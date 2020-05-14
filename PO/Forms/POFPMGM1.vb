Imports Microsoft.Office.Interop
Imports System.Drawing

Public Class POFPMGM1

    Dim rowPOTPMGM1 As DataRow
    Dim sqlPOTPMGMX As String = ""
    Dim PROGRAM_NO As String
    Dim images_folder As String = "C:\dmp\Images"
    Dim images As New Dictionary(Of String, List(Of System.Drawing.Bitmap))
 
    Dim APPR_STATUS_CODE_BackColors As New Dictionary(Of String, System.Drawing.Color)
    Dim APPR_STATUS_CODE_ForeColors As New Dictionary(Of String, System.Drawing.Color)
     
    Dim CONV_TOPIC_NOs As New List(Of String)

    Dim colors() As System.Drawing.Color = {Color.Green, Color.Purple, Color.Blue, Color.Red}

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")

        With dst

            Dim sqlBrands As String = ASCMAIN1.Flattened_List("PROGRAM_NO", "BRAND_CODE", "(SELECT DISTINCT PROGRAM_NO, BRAND_CODE FROM POTPMGM2)", ",", "")

            sqlPOTPMGMX = "Select POTPMGM1.*, X.STYLE_CODE_1, X.STYLES, X.BRAND_CODE_1, B.BRAND_CODES BRANDS" & vbCrLf _
            & " from POTPMGM1" & vbCrLf _
            & ", (Select PROGRAM_NO, MIN (STYLE_CODE) STYLE_CODE_1, Count (*) STYLES, MIN (BRAND_CODE) BRAND_CODE_1 from POTPMGM2 group by PROGRAM_NO) X" & vbCrLf _
            & ", (" & sqlBrands & ") B" & vbCrLf _
            & " where X.PROGRAM_NO = POTPMGM1.PROGRAM_NO" & vbCrLf _
            & "   and B.PROGRAM_NO (+) = POTPMGM1.PROGRAM_NO"
            ASCMAIN1.sql = sqlPOTPMGMX
            Create_TDA(.Tables.Add, "POTPMGMX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "POTPMGM1", "*")
            With .Tables("POTPMGM1")
                .Columns.Add("LOGO", GetType(System.Byte()))
                '.PrimaryKey = New DataColumn() {.Columns("PROGRAM_NO")}
            End With

            ASCMAIN1.sql = "Select POTPMGM2.*" & vbCrLf _
                & " from POTPMGM2 where POTPMGM2.PROGRAM_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPMGM2", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select POTPMGM3.*, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & " from POTPMGM3, ICTCOLR1 where ICTCOLR1.COLOR_CODE = POTPMGM3.COLOR_CODE" & vbCrLf _
                & " and POTPMGM3.PROGRAM_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPMGM3", "**", 0, True, "V", 2)


            ASCMAIN1.sql = "Select * from POTPMGM9 where PROGRAM_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPMGM9", "**", 0, True, "V", 0)

            ASCMAIN1.sql = "Select POTCTOP1.*" & vbCrLf _
                & " from POTCTOP1" & vbCrLf _
                & " where POTCTOP1.PROGRAM_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTCTOP1", "**", 0, True, "V", 1)

            ' NEED TO PULL IN ONLY THOSE CONVS THAT MATCH MY CATEGORICAL SUBSCRIPTIONS

            ASCMAIN1.sql = "Select TATCONV1.*, TATCONPU.CONV_ACK_IND, DECODE(TATCONPU.CONV_NO,NULL,'0','1') CONV_ACK_REQD" & vbCrLf _
                & " from POTCTOP1, TATCONV1, TATCONPU" & vbCrLf _
                & " where POTCTOP1.PROGRAM_NO = :PARM1" & vbCrLf _
                & "   and TATCONPU.USER_ID (+) = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
                & "   and TATCONPU.CONV_NO (+) = TATCONV1.CONV_NO" & vbCrLf _
                & "   and TATCONV1.TABLE_NAME = 'POTCTOP1' and TATCONV1.TABLE_KEY = POTCTOP1.CONV_TOPIC_NO" & vbCrLf _
                & "   and TATCONV1.CONV_NO_PREV is Null"
            Create_TDA(.Tables.Add, "TATCONV1", "**", 0, True, "V", 1)


            ASCMAIN1.sql = "Select TATCONV1.*, TATCONPU.CONV_ACK_IND, DECODE(TATCONPU.CONV_NO,NULL,'0','1') CONV_ACK_REQD" & vbCrLf _
                & " from POTCTOP1, TATCONV1, TATCONPU" & vbCrLf _
                & " where POTCTOP1.PROGRAM_NO = :PARM1" & vbCrLf _
                & "   and TATCONPU.USER_ID (+) = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
                & "   and TATCONPU.CONV_NO (+) = TATCONV1.CONV_NO" & vbCrLf _
                & "   and TATCONV1.TABLE_NAME = 'POTCTOP1' and TATCONV1.TABLE_KEY = POTCTOP1.CONV_TOPIC_NO" & vbCrLf _
                & "   and TATCONV1.CONV_NO_PREV is NOT Null"
            Create_TDA(.Tables.Add, "TATCONV1_R", "**", 0, False, "V", 1)

            Create_Relation("TATCONV1", "TATCONV1_R", "CONV_NO", "CONV_NO_PREV")

            Create_TDA(.Tables.Add, "ASTATTA2", "*", 3)

            Create_TDA(.Tables.Add, "POTPGMC1", "*", 0, False)
            Fill_Records("POTPGMC1")


            Create_TDA(.Tables.Add, "ICTBRAN1", "*", 0, False)
            With .Tables("ICTBRAN1").Columns
                .Add("SEL")
            End With
            .Tables("ICTBRAN1").Columns("SEL").DefaultValue = "0"

            Fill_Records("ICTBRAN1")

            Create_TDA(.Tables.Add, "POTPMGMO", "*", 1)

            Create_TDA(.Tables.Add, "POTPMGMD", "*", 1)

            ASCMAIN1.sql = "Select POTPMGMC.*, POTPGMC1.PROGRAM_CATGY_DESC" & vbCrLf _
                & ", POTPGMC1.PROGRAM_CATGY_PHASE, POTPGMC1.PROGRAM_CATGY_SEQ, POTPGMC1.PROGRAM_CATGY_USAGE, POTPGMC1.PROGRAM_CATGY_DEFAULT" & vbCrLf _
                & " from POTPGMC1,POTPMGMC" & vbCrLf _
                & " where POTPGMC1.PROGRAM_CATGY_CODE = POTPMGMC.PROGRAM_CATGY_CODE" & vbCrLf _
                & "   and POTPMGMC.PROGRAM_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPMGMC", "**", 0, True, "V", 2)
            With .Tables("POTPMGMC").Columns
                .Add("SEL")
                .Add("TASKS", GetType(System.Int32))
                .Add("DATE_PLANNED_1", GetType(System.DateTime))
                .Add("DATE_PLANNED_2", GetType(System.DateTime))
                .Add("ASSIGNED_TO")
                .Add("SCHEDULE")
            End With
            .Tables("POTPMGMC").Columns("SEL").DefaultValue = "0"

            Create_Relation("POTPMGMC", "POTCTOP1", "PROGRAM_CATGY_CODE")

            With .Tables("POTPMGMC")
                .Columns("TASKS").Expression = "COUNT(CHILD.CONV_TOPIC_NO)"
                .Columns("DATE_PLANNED_1").Expression = "MIN(CHILD.DATE_PLANNED)"
                .Columns("DATE_PLANNED_2").Expression = "MAX(CHILD.DATE_PLANNED)"
            End With

            With .Tables("POTCTOP1")
                .Columns.Add("PROGRAM_CATGY_STATUS", GetType(System.String), "PARENT.PROGRAM_CATGY_STATUS")
                 
            End With


            Create_TDA(.Tables.Add, "POTPMGM8", "*", 1, True)

            Create_Relation("ICTBRAN1", "POTPMGM8", "BRAND_CODE")
            .Tables("ICTBRAN1").Columns.Add("COLLS", GetType(System.Int32), "COUNT(CHILD.COLLECTION_NO)")
        End With

        grdPOTPMGM2.DataSource = dst.Tables("POTPMGM2")
        grdPOTPMGM3.DataSource = dst.Tables("POTPMGM3")
        grdPOTPMGMC.DataSource = dst.Tables("POTPMGMC")
        grdPOTPMGMO.DataSource = dst.Tables("POTPMGMO")

        grdPOTPMGM9.DataSource = dst.Tables("POTPMGM9")

        grdPOTPMGMX.DataSource = dst.Tables("POTPMGMX")
        grdICTBRAN1.DataSource = dst.Tables("ICTBRAN1")

        grdPOTCTOP1.DataSource = dst.Tables("POTCTOP1")
        grdTATCONV1.DataSource = dst.Tables("TATCONV1")

        grdTATCONV1.DisplayLayout.Bands(1).ColHeadersVisible = False

        For Each gcol As UltraWinGrid.UltraGridColumn In grdTATCONV1.DisplayLayout.Bands(0).Columns
            If gcol.Key = "CONV_ACK_IND" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        Create_Summary(grdPOTPMGMX, "PROGRAM_NO", "Count")


        For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTPMGMC.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Color.White
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            If gcol.Key = "SEL" Then
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.Header.Appearance.BackColor2 = Color.LightGray
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        Create_Summary(grdPOTPMGMC, "PROGRAM_CATGY_CODE", "Count")
        Create_Summary(grdPOTPMGMC, "SEL")


        For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTPMGMC.DisplayLayout.Bands(1).Columns
            gcol.Header.Appearance.BackColor = Color.White
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            gcol.Header.Appearance.BackColor2 = Color.LightGray
        Next

        ASCMAIN1.Add_Value_List(grdPOTPMGMC, "PROGRAM_CATGY_PHASE", Nothing, New String() {":", "D:Development (Pre-Booking)", "B:Booking (PO)"})

        ASCMAIN1.Add_Value_List(grdPOTPMGMC, "PROGRAM_CATGY_STATUS", Nothing, New String() {":", "A:Active", "C:Completed", "F:Future"})
        ASCMAIN1.Add_Value_List(grdPOTCTOP1, "PROGRAM_CATGY_STATUS", Nothing, New String() {":", "A:Active", "C:Completed", "F:Future"})

        SplitContainer1.Panel1Collapsed = True

        MakeTransparent(chkShowAllStyles)
        MakeTransparent(chkShowAllBrands)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                '   Validate_Code("STEP_TEMPLATE")
                If Absx1.txtFor("PROGRAM_NAME").Text = "" Then
                    EMsg &= vbCr & "You must specify a Program Name"
                End If
                If Absx1.txtFor("SEASON_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify a Season"
                End If
                If Absx1.dteFor("PROGRAM_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "You must specify a Program Inception Date"
                End If

            Case "View", "Edit"
                If Absx1.txtFor("PROGRAM_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a valid Program No"
                Else
                    PROGRAM_NO = Absx1.txtFor("PROGRAM_NO").Text
                    rowPOTPMGM1 = LookUp("POTPMGM1", PROGRAM_NO)
                    If rowPOTPMGM1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Program No " & PROGRAM_NO
                    End If
                End If

            Case "Update"
                If Absx1.txtFor("PROGRAM_NAME").Text = "" Then
                    EMsg &= vbCr & "You must specify a Program Name"
                End If
                If Absx1.txtFor("PROGRAM_TAG").Text = "" Then
                    EMsg &= vbCr & "You must specify a Program Tag (to appear in the Calendar)"
                End If

                If dst.Tables("POTPMGM2").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Styles Entered"
                Else
                    For Each rowPOTPMGM2 As DataRow In dst.Tables("POTPMGM2").Select("", "", DataViewRowState.CurrentRows)
                    Next
                End If

                If dst.Tables("POTPMGMO").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Users Specified"
                Else
                    For Each rowPOTPMGM2 As DataRow In dst.Tables("POTPMGM2").Select("", "", DataViewRowState.CurrentRows)
                    Next
                End If

                If EMsg = "" Then

                End If

            Case "Delete"

                If ASCMAIN1.USER_ID <> rowPOTPMGM1.Item("INIT_OPER") & "" Then
                    EMsg &= vbCr & "Only " & rowPOTPMGM1.Item("INIT_OPER") & " may Delete this Program"
                End If

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete this Program", _
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
                If dst.Tables("POTPMGM2").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Styles on the Program Sheet"
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

                '  Update_Record_TDA("POTPMGM2", "1=1")
                Synch_TABLE_NAME("POTPMGM1")

                Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    rowPOTPMGM1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
                End If


                Print_Report_Begin()


                Dim RPT As String = "ICRQUOT1"


                If eItemKey = "email" Then
                    Dim tempFileName As String = rowPOTPMGM1.Item("PROGRAM_NO")
                    Dim REPORT_NO As String = Generate_Report(RPT, "Program Sheet", "", "", "PDF", tempFileName, False)
                    ' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
                    Print_Report_End(, True)
                    email_Quote(tempFileName)
                Else
                    Generate_Report(RPT, "Program Sheet")
                    Print_Report_End()
                End If

 
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

                Setup_tabPOTPMGMX()
                .Groups("Show Programs").Visible = Not ScreenMode
                .Groups("Task Owners").Visible = ScreenMode
                .Groups("Program Image").Visible = ScreenMode
            End With
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grpAttributes.Visible = ScreenMode

        tabPOTPMGMX.Visible = Not ScreenMode

        If ScreenMode Then

            cmdChangeProgramImage.Visible = ((EntryMode = "N") Or (EntryMode = "E"))
            btnCreateTopics.Visible = ((EntryMode = "N") Or (EntryMode = "E"))

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(grpAttributes, (EntryMode = "V"))
            Set_Read_Only_for_ctl(Absx1.txtFor("PROGRAM_NAME"), (EntryMode = "V"))
            Set_Read_Only_for_ctl(Absx1.txtFor("PROGRAM_TAG"), (EntryMode = "V"))
            With tabDetails
                .SelectedTab = .Tabs("Styles")
            End With

            If EntryMode = "V" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdPOTPMGM2, grdPOTPMGM3, grdPOTPMGMC, grdPOTPMGMO, grdPOTCTOP1}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                grdICTBRAN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                With grdICTBRAN1.DisplayLayout.Bands(1).Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With

                '  grdPOTCTOP1.DisplayLayout.Bands(0).Columns("PROGRAM_CATGY_CODE").Hidden = True
                grdTATCONV1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdPOTPMGM2, grdPOTPMGM3, grdPOTPMGMC, grdPOTPMGMO, grdPOTCTOP1}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Next
                grdICTBRAN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                With grdICTBRAN1.DisplayLayout.Bands(1).Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With

                '  grdPOTCTOP1.DisplayLayout.Bands(0).Columns("PROGRAM_CATGY_CODE").Hidden = False
                grdTATCONV1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If


            Setup_grdPOTPMGM8_brand()
        Else
            Clear_Record()

            splCalendar.Parent = tabPOTPMGMX.Tabs("Calendar").TabPage
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTPMGMX", "POTPMGM1", "POTPMGM2", "POTPMGM2", "POTPMGM3", "POTPMGM8", _
             "POTPMGMC", "POTPMGMO", _
             "POTCTOP1", "TATCONV1", "TATCONV1_R"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        grdPOTPMGM2.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Absx1.txtFor("PROGRAM_NAME").Text = ""
        'Absx1.dteFor("SHIP_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("PROGRAM_NO").Text = ""

        PROGRAM_NO = ""
        images.Clear()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowPOTPMGM1 = dst.Tables("POTPMGM1").NewRow
            PROGRAM_NO = ASCMAIN1.Next_Control_No("POTPMGM1.PROGRAM_NO")
            rowPOTPMGM1.Item("PROGRAM_NO") = PROGRAM_NO
            rowPOTPMGM1.Item("PROGRAM_NAME") = HFs("PROGRAM_NAME")
            rowPOTPMGM1.Item("SEASON_CODE") = HFs("SEASON_CODE")
            'rowPOTPMGM1.Item("SHIP_DATE") = HFs("SHIP_DATE")
            ' rowPOTPMGM1.Item("TOTAL_QTY") = Val(HFs("TOTAL_QTY"))
            ' rowPOTPMGM1.Item("CUST_CODE") = HFs("CUST_CODE")
            ' rowPOTPMGM1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowPOTPMGM1.Item("PROGRAM_TAG") = HFs("PROGRAM_TAG")
            rowPOTPMGM1.Item("PROGRAM_DATE") = HFs("PROGRAM_DATE")
            rowPOTPMGM1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTPMGM1.Item("INIT_DATE") = DATETIME_STAMP
            rowPOTPMGM1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTPMGM1.Item("LAST_DATE") = DATETIME_STAMP
            dst.Tables("POTPMGM1").Rows.Add(rowPOTPMGM1)
        Else
            rowPOTPMGM1 = Fill_Record("POTPMGM1", PROGRAM_NO)
            dst.AcceptChanges()
        End If

        images.Clear()

        Fill_Records("POTPMGM2", PROGRAM_NO)
        Sort_grdColumns(grdPOTPMGM2, "STYLE_CODE")

        Fill_Records("POTPMGMO", PROGRAM_NO)
        Sort_grdColumns(grdPOTPMGMO, "USER_ID")

        Fill_Records("POTPMGMC", PROGRAM_NO)
        For Each row As DataRow In dst.Tables("POTPMGMC").Select("")
            row.Item("SEL") = "1"
        Next
        For Each row As DataRow In dst.Tables("POTPGMC1").Select("")
            Dim PROGRAM_CATGY_CODE As String = row.Item("PROGRAM_CATGY_CODE")
            Dim rowPOTPMGMC As DataRow = dst.Tables("POTPMGMC").Rows.Find(New String() {PROGRAM_NO, PROGRAM_CATGY_CODE})
            If rowPOTPMGMC Is Nothing Then
                rowPOTPMGMC = dst.Tables("POTPMGMC").Rows.Add(New String() {PROGRAM_NO, PROGRAM_CATGY_CODE})
                rowPOTPMGMC("PROGRAM_CATGY_DESC") = row.Item("PROGRAM_CATGY_DESC")
                rowPOTPMGMC("PROGRAM_CATGY_PHASE") = row.Item("PROGRAM_CATGY_PHASE")
                rowPOTPMGMC("PROGRAM_CATGY_SEQ") = row.Item("PROGRAM_CATGY_SEQ")
                rowPOTPMGMC("PROGRAM_CATGY_USAGE") = row.Item("PROGRAM_CATGY_USAGE")
                rowPOTPMGMC("PROGRAM_CATGY_DEFAULT") = row.Item("PROGRAM_CATGY_DEFAULT")
                rowPOTPMGMC("PROGRAM_CATGY_STATUS") = "A"
                If EntryMode = "N" And row.Item("PROGRAM_CATGY_DEFAULT") = "1" Then
                    rowPOTPMGMC("SEL") = "1"
                End If
            End If
        Next

        With grdPOTPMGMC.DisplayLayout.Bands(0).SortedColumns
            .Clear()
            .Add("PROGRAM_CATGY_PHASE", True, True) ' DESC SO THAT DEV PRECEDES BOOKING
            .Add("PROGRAM_CATGY_SEQ", False, False)
            .Add("PROGRAM_CATGY_CODE", False, False)
        End With
        grdPOTPMGMC.Rows.ExpandAll(False)

        Fill_Records("POTPMGM8", PROGRAM_NO)
        Fill_Records("POTPMGM9", PROGRAM_NO)

        Fill_Records("POTCTOP1", PROGRAM_NO)
        Fill_Records("TATCONV1", PROGRAM_NO)
        Fill_Records("TATCONV1_R", PROGRAM_NO)

        Setup_TATCONV1()

        For Each rowPOTPMGMC As DataRow In dst.Tables("POTPMGMC").Select("")
            Dim PROGRAM_CATGY_CODE As String = rowPOTPMGMC.Item("PROGRAM_CATGY_CODE")
            Dim ASSIGNED_TO As String = ""
            For Each rowPOTPMGMD As DataRow In dst.Tables("POTPMGMD").Select("PROGRAM_CATGY_CODE = '" & PROGRAM_CATGY_CODE & "'", "USER_ID")
                Dim USER_ID As String = rowPOTPMGMD.Item("USER_ID")
                ASSIGNED_TO &= "," & USER_ID
            Next
            rowPOTPMGMC.Item("ASSIGNED_TO") = Mid(ASSIGNED_TO, 2)
        Next

        Load_Program_Graphic()

        EnforceConstraints(True)

        Dim TASKS As Int32 = dst.Tables("POTCTOP1").Compute("COUNT(CONV_TOPIC_NO)", "DATE_PLANNED IS NOT NULL")
        If TASKS > 0 Then
            Dim DATE_PLANNED_MIN As Date = dst.Tables("POTCTOP1").Compute("MIN(DATE_PLANNED)", "")
            Dim DATE_PLANNED_MAX As Date = dst.Tables("POTCTOP1").Compute("MAX(DATE_PLANNED)", "")
            Dim DAYS As Int32 = DATE_PLANNED_MAX.Subtract(DATE_PLANNED_MIN).TotalDays + 1
            Dim WEEKS As Int32 = DAYS / 7 + 1

            For Each row As DataRow In dst.Tables("POTPMGMC").Select("TASKS <> 0")
                Dim DATE_PLANNED_1 As Date = row.Item("DATE_PLANNED_1")
                Dim DATE_PLANNED_2 As Date = row.Item("DATE_PLANNED_1")
                ' Dim SCHEDULE As String = "".PadLeft(DAYS)
                Dim SCHEDULE As String = "".PadLeft(WEEKS)
                Dim DAY_1 As Int32 = DATE_PLANNED_1.Subtract(DATE_PLANNED_MIN).TotalDays
                Dim DAY_2 As Int32 = DATE_PLANNED_2.Subtract(DATE_PLANNED_MIN).TotalDays
                Dim WEEK_1 As Int32 = DAY_1 / 7
                Dim WEEK_2 As Int32 = DAY_2 / 7
                ' Mid(SCHEDULE, DAY_1 + 1, DAY_2 - DAY_1 + 1) = "-"
                Mid(SCHEDULE, WEEK_1 + 1, WEEK_2 - WEEK_1 + 1) = "-"
                For Each rowPOTCTOP1 As DataRow In row.GetChildRows("POTPMGMC_POTCTOP1")
                    If rowPOTCTOP1.Item("DATE_PLANNED") & "" <> "" Then
                        Dim DATE_PLANNED As Date = rowPOTCTOP1.Item("DATE_PLANNED")
                        Dim DAY As Int32 = DATE_PLANNED.Subtract(DATE_PLANNED_MIN).TotalDays
                        Dim WEEK As Int32 = DAY / 7
                        ' Mid(SCHEDULE, DAY + 1, 1) = "T"
                        Mid(SCHEDULE, WEEK + 1, 1) = "T"
                    End If
                Next
                row.Item("SCHEDULE") = SCHEDULE
            Next
            grdPOTPMGMC.DisplayLayout.Bands(0).Columns("SCHEDULE").Hidden = False
        Else
            grdPOTPMGMC.DisplayLayout.Bands(0).Columns("SCHEDULE").Hidden = True
        End If
 
        Setup_grdPOTPMGM2()

        grdPOTPMGMO.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("ASTATTA2")
        Update_Record_TDA("POTPMGM1")
        Update_Record_TDA("POTPMGM2") 
        Update_Record_TDA("POTPMGM8")
        Update_Record_TDA("POTPMGM9")

        grdPOTPMGMO.ActiveRow = Nothing ' NEED TO DO THIS OTHERWISE WE GET AN ERROR ABOUT PROGRAM_NO DOES NOT PERMIT NULLS - MAYBE BECAUSE THERE IS ONLY 1 EDITABLE FIELD IN THIS GRID
        Update_Record_TDA("POTPMGMO")

        ASCDATA1.DeleteRows(dst.Tables("POTPMGMC"), "ISNULL(SEL,'0') <> '1' AND ISNULL(TASKS,0) = 0 AND ISNULL(ASSIGNED_TO,'') = ''")
        Update_Record_TDA("POTPMGMC", "PROGRAM_NO = '" & PROGRAM_NO & "'")

        Update_Record_TDA("POTCTOP1")

        dst.Tables("POTPMGMD").Rows.Clear()
        For Each POTPMGMC As DataRow In dst.Tables("POTPMGMC").Select("ISNULL(ASSIGNED_TO,'') <> ''")
            Dim PROGRAM_CATGY_CODE As String = POTPMGMC.Item("PROGRAM_CATGY_CODE")
            Dim ASSIGNED_TO As String = POTPMGMC.Item("ASSIGNED_TO") & ""
            For Each USER_ID As String In Split(ASSIGNED_TO, ",")
                dst.Tables("POTPMGMD").Rows.Add(New String() {PROGRAM_NO, PROGRAM_CATGY_CODE, USER_ID})
            Next
        Next
        Update_Record_TDA("POTPMGMD", "PROGRAM_NO = '" & PROGRAM_NO & "'")

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()

        For Each TABLE_NAME As String In New String() {"POTPMGM1", "POTPMGM2", "POTPMGM8", "POTPMGM9", "POTPMGMC", "POTPMGMD", "POTPMGMO"}
            Delete_Records(TABLE_NAME)
        Next

        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where PROGRAM_NO = '" & Absx1.txtFor("PROGRAM_NO").Text & "'")
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

                Absx1.txtFor("PROGRAM_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTPMGM1"
            E.COLUMN_NAME = "PROGRAM_NO"
            E.CODE_VALUE = Absx1.txtFor("PROGRAM_NO").Text
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
        Load_Popup_Menu(grdPOTPMGMX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdPOTPMGM2, "SBBBB", "Show Filter", "Get Styles", "Product Line Maintenance", "Product Line Inquiry", "Sequence as Shown")
        Load_Popup_Menu(grdPOTPMGM3, "B", "Get Colors")
        Load_Popup_Menu(grdTATCONV1, "BB", "Respond", "New Message")
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

        Select Case e.SourceControl.Name
            Case "grdTATCONV1"
                tlb_btn = DirectCast(tlb_pop.Tools("Respond"), UltraWinToolbars.ButtonTool)
                ' tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And grdTATCONV1.ActiveRow IsNot Nothing
                tlb_btn.SharedProps.Visible = grdTATCONV1.ActiveRow IsNot Nothing
                If grdPOTPMGMO.ActiveRow IsNot Nothing Then
                    ' tlb_btn.SharedProps.Caption = "Respond to " & grdPOTPMGM9.ActiveRow.Cells("CONV_NO").Value
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("New Message"), UltraWinToolbars.ButtonTool)
                ' tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And grdPOTCTOP1.ActiveRow IsNot Nothing
                tlb_btn.SharedProps.Visible = grdPOTCTOP1.ActiveRow IsNot Nothing
                If grdPOTPMGMO.ActiveRow IsNot Nothing Then
                    ' tlb_btn.SharedProps.Caption = "Respond to " & grdPOTPMGM9.ActiveRow.Cells("CONV_NO").Value
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdPOTPMGM2"
                    tlb_btn = DirectCast(tlb_pop.Tools("Sequence as Shown"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Styles"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Case "grdPOTPMGM3"
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Colors"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            'Case "Get Styles"
            '    Get_Styles()

            'Case "Get Colors"
            '    Get_Colors()
                 
            Case "New Message"

                Using F As New TAC.TAFPMGMM("", Nothing, Me, True)

                    F.EntryMode = "N"
                    F.CONV_SUBJECT = grdPOTCTOP1.ActiveRow.Cells("CONV_TOPIC_DESC").Value
                    '  F.CONV_NOTES = "{Enter Notes Here}"
                    F.CONV_TOPIC_NO = grdPOTCTOP1.ActiveRow.Cells("CONV_TOPIC_NO").Value
                    Dim PROGRAM_CATGY_CODE As String = grdPOTCTOP1.ActiveRow.Cells("PROGRAM_CATGY_CODE").Value
                    F.PROGRAM_CATGY_CODE = PROGRAM_CATGY_CODE

                    If grdPOTPMGMO.Selected.Rows.Count = 1 Then
                        F.MESSAGE_BY = grdPOTPMGMO.Selected.Rows(0).Cells("USER_ID").Value
                        grdPOTPMGMO.Selected.Rows.Clear()
                    End If

                    F.ShowDialog()
                    If F.result = "U" Then
                        ASCMAIN1.sql = "Select TATCONV1.*, TATCONPU.CONV_ACK_IND, DECODE(TATCONPU.CONV_NO,NULL,'0','1') CONV_ACK_REQD" & vbCrLf _
                            & " from TATCONV1, TATCONPU" & vbCrLf _
                            & " where TATCONV1.CONV_NO = '" & F.CONV_NO & "'" & vbCrLf _
                            & "   and TATCONPU.CONV_NO (+) = TATCONV1.CONV_NO"
                        Fill_Records("TATCONV1", , False, ASCMAIN1.sql)
                    End If
                    F.Dispose()
                End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            Case "Respond"
                Using F As New TAC.TAFPMGMM("", Nothing, Me, True)

                    F.CONV_NO_PREV = grd.ActiveRow.Cells("CONV_NO").Value
                    F.CONV_NO_PREV_NOTES = grd.ActiveRow.Cells("CONV_NOTES").Value & ""
                    If grd.ActiveRow.Cells("CONV_NO_PREV").Value & "" <> "" Then
                        F.CONV_NO_PREV = grd.ActiveRow.Cells("CONV_NO_PREV").Value
                        F.CONV_NO_PREV_NOTES = grd.ActiveRow.Cells("CONV_NOTES").Value & ""
                    End If

                    F.EntryMode = "N"
                    F.CONV_SUBJECT = grdPOTCTOP1.ActiveRow.Cells("CONV_TOPIC_DESC").Value
                    F.CONV_TOPIC_NO = grdPOTCTOP1.ActiveRow.Cells("CONV_TOPIC_NO").Value

                    Dim PROGRAM_CATGY_CODE As String = grdPOTCTOP1.ActiveRow.Cells("PROGRAM_CATGY_CODE").Value
                    F.PROGRAM_CATGY_CODE = PROGRAM_CATGY_CODE

                    If grdPOTPMGMO.Selected.Rows.Count = 1 Then
                        F.MESSAGE_BY = grdPOTPMGMO.Selected.Rows(0).Cells("USER_ID").Value
                        grdPOTPMGMO.Selected.Rows.Clear()
                    End If

                    F.ShowDialog()

                    If F.result = "U" Then
                        Dim rowTATCONV1 As DataRow = LookUp("TATCONV1", F.CONV_NO_PREV)
                        Dim rowTATCONV1_local As DataRow = dst.Tables("TATCONV1").Rows.Find(F.CONV_NO_PREV)
                        rowTATCONV1_local.Item("CONV_NOTES") = rowTATCONV1.Item("CONV_NOTES")

                        ASCMAIN1.sql = "Select TATCONV1.*, TATCONPU.CONV_ACK_IND, DECODE(TATCONPU.CONV_NO,NULL,'0','1') CONV_ACK_REQD" & vbCrLf _
                            & " from TATCONV1, TATCONPU" & vbCrLf _
                            & " where TATCONV1.CONV_NO = '" & F.CONV_NO & "'" & vbCrLf _
                            & "   and TATCONPU.CONV_NO (+) = TATCONV1.CONV_NO"
                        Fill_Records("TATCONV1_R", , False, ASCMAIN1.sql)

                    End If
                    F.Dispose()
                End Using


            Case "Sequence as Shown"
                Dim SEQ As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTPMGM2.Rows
                    SEQ += 10
                    grow.Cells("SEQ").Value = SEQ
                    grow.Update()
                Next

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
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "PROGRAM_NO"
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

#Region "grdPOTPMGM2"

    Private Sub grdPOTPMGM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTPMGM2.AfterCellUpdate
        If Not e.Cell.Row.IsDataRow Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "STYLE_NO"
            Case "STYLE_CODE"

                'grdCodeDesc(grdPOTPMGM2, "ICTPLIN2", "STYLE_CODE_PLM", "STYLE_DESC")
                '' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                'If cdr IsNot Nothing Then
                '    Dim STYLE_CODE_PLM As String = e.Cell.Value
                '    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                '    'e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = cdr.Item("SALES_DIVISION_CODE")
                '    'e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = cdr.Item("STYLE_CLASS_CODE") & ""
                '    'e.Cell.Row.Cells("STYLE_PRICE").Value = cdr.Item("STYLE_PRICE")

                'Else
                '    grdPOTPMGM2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                'End If
        End Select
    End Sub

    Private Sub grdPOTPMGM2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTPMGM2.AfterRowActivate

        If Not grdPOTPMGM2.ActiveRow.IsDataRow Then Exit Sub

        With grdPOTPMGM2.DisplayLayout.Bands(0)
            If grdPOTPMGM2.ActiveRow.IsAddRow Then
                grdPOTPMGM2.ActiveCell = grdPOTPMGM2.ActiveRow.Cells("STYLE_CODE")
                grdPOTPMGM2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else

            End If
        End With

        Setup_grdPOTPMGM2()
    End Sub

    Sub Setup_grdPOTPMGM2()


        If grdPOTPMGM2.ActiveRow Is Nothing OrElse (grdPOTPMGM2.ActiveRow.IsAddRow Or Not grdPOTPMGM2.ActiveRow.IsDataRow) Then


            tabStyle.Visible = False
        Else
            Dim STYLE_NO As String = grdPOTPMGM2.ActiveRow.Cells("STYLE_NO").Value & ""
            Dim STYLE_CODE As String = grdPOTPMGM2.ActiveRow.Cells("STYLE_CODE").Value & ""
            tabStyle.Visible = True
            Dim dvw As DataView = DirectCast(grdPOTPMGM9.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PROGRAM_NO = '" & PROGRAM_NO & "' and STYLE_NO = '" & STYLE_NO & "'"
            Sort_grdColumns(grdPOTPMGM9, "INIT_DATE")
            grdPOTPMGM9.Text = "Conversation Log for Style " & STYLE_CODE
        End If
    End Sub

    Private Sub grdPOTPMGM2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTPMGM2.AfterRowsDeleted

        If images.Count > 0 Then
            Dim STYLE_NOs As New List(Of String)
            For Each rowPOTPMGM2 As DataRow In dst.Tables("POTPMGM2").Select("")
                Dim STYLE_NO As String = rowPOTPMGM2.Item("STYLE_NO")
                If Not STYLE_NOs.Contains(STYLE_NO) Then
                    STYLE_NOs.Add(STYLE_NO)
                End If
            Next
            Dim STYLE_NOs_to_delete As New List(Of String)
            For Each k As String In images.Keys
                If Not STYLE_NOs.Contains(k) Then
                    STYLE_NOs_to_delete.Add(k)
                End If
            Next
            For Each k As String In STYLE_NOs_to_delete
                images.Remove(k)
            Next
        End If

        Setup_grdPOTPMGM2()
    End Sub
    Private Sub grdPOTPMGM2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTPMGM2.AfterRowUpdate
        If e.Row IsNot Nothing AndAlso Not e.Row.IsAddRow Then

        End If
        Dim PROGRAM_NO As String = e.Row.Cells("PROGRAM_NO").Value
        Dim STYLE_NO As String = e.Row.Cells("STYLE_NO").Value
        Dim rowPOTPMGM2 As DataRow = dst.Tables("POTPMGM2").Rows.Find(New Object() {PROGRAM_NO, STYLE_NO})
        If Not images.ContainsKey(STYLE_NO) Then
            Get_Images(STYLE_NO)
        End If
    End Sub

    Private Sub grdPOTPMGM2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTPMGM2.BeforeExitEditMode
        If grdPOTPMGM2.ActiveCell Is Nothing Then Exit Sub
        If Not grdPOTPMGM2.ActiveRow.IsDataRow Then Exit Sub
        With grdPOTPMGM2.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If
                    End If
                    'If .Text <> "" Then
                    '    cdr = LookUp("ICTPLIN2", .Text)
                    '    If cdr Is Nothing Then
                    '        ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                    '        If .Value IsNot Nothing Then
                    '            .Value = ""
                    '        End If
                    '        e.Cancel = True
                    '    End If
                    'End If
            End Select
        End With
    End Sub

    Private Sub grdPOTPMGM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTPMGM2.BeforeRowUpdate
        With grdPOTPMGM2
            If e.Row.Cells("STYLE_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else

            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PROGRAM_NO").Text = "" Then
                    .ActiveRow.Cells("PROGRAM_NO").Value = Absx1.CtlFor("PROGRAM_NO").Text
                    .ActiveRow.Cells("STYLE_NO").Value = ASCMAIN1.Next_Control_No("POTPMGM2.STYLE_NO")
                    .ActiveRow.Cells("BRAND_CODE").Value = grdICTBRAN1.ActiveRow.ParentRow.Cells("BRAND_CODE").Value
                    .ActiveRow.Cells("COLLECTION_NO").Value = grdICTBRAN1.ActiveRow.Cells("COLLECTION_NO").Value
                    .ActiveRow.Cells("SEQ").Value = Val(dst.Tables("POTPMGM2").Compute("Max(SEQ)", "") & "") + 10
                End If
            End If
        End With

    End Sub

    Private Sub grdPOTPMGM2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTPMGM2.ClickCellButton

        If grdPOTPMGM2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
        End Select
        grdClickCellButton(grdPOTPMGM2, sql_where, True)

    End Sub

#End Region
     
 
     
#Region "grdICTBRAN1"
    Private Sub grdICTBRAN1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTBRAN1.AfterCellUpdate
        Select Case e.Cell.Column.Key

        End Select
    End Sub

    Private Sub grdICTBRAN1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTBRAN1.AfterRowActivate
        If grdICTBRAN1.ActiveRow.Band.Index = 1 Then
            Setup_grdPOTPMGM8_brand()
        End If
    End Sub

    Private Sub grdICTBRAN1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTBRAN1.BeforeRowUpdate
        If grdICTBRAN1.ActiveRow.Band.Index = 1 Then
            With grdICTBRAN1
                If Not e.Cancel Then
                    If e.Row.Cells("PROGRAM_NO").Text = "" Then
                        .ActiveRow.Cells("PROGRAM_NO").Value = Absx1.CtlFor("PROGRAM_NO").Text
                        .ActiveRow.Cells("BRAND_CODE").Value = .ActiveRow.ParentRow.Cells("BRAND_CODE").Value
                        Dim sqlx As String = "PROGRAM_NO = '" & .ActiveRow.Cells("PROGRAM_NO").Value & "'"
                        .ActiveRow.Cells("COLLECTION_NO").Value = Val(dst.Tables("POTPMGM8").Compute("MAX(COLLECTION_NO)", sqlx) & "") + 1
                        ' .ActiveRow.Cells("SEQ").Value = Val(dst.Tables("ICTBRAN1").Compute("Max(SEQ)", "") & "") + 10

                    End If
                End If
            End With
        End If
    End Sub

    Private Sub grdICTBRAN1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTBRAN1.ClickCellButton
        'Dim sql_where As String = ""
        'grdClickCellButton(grdICTBRAN1, sql_where, sql_where <> "")

        'If e.Cell.Row.Band.Key = "ICTBRAN1_POTPMGM2" Then
        '    If e.Cell.Row.Cells("TASK_STATUS").Value & "" = "U" Then
        '        MsgBox("Task has not yet been assigned")
        '    Else
        '        Dim EVENT_KEY As String = e.Cell.Row.Cells("PROGRAM_NO").Value & ":" & e.Cell.Row.Cells("COLLECTION_NO").Value & ":" & e.Cell.Row.Cells("TASK_LNO").Value
        '        Edit_Task(EVENT_KEY)
        '    End If
        'End If

    End Sub

    Sub Setup_grdPOTPMGM8_brand()

        If grdICTBRAN1.ActiveRow Is Nothing OrElse grdICTBRAN1.ActiveRow.IsAddRow OrElse grdICTBRAN1.ActiveRow.Band.Index <> 1 Then
            splStyle.Visible = False
        Else
            splStyle.Visible = True
            Dim COLLECTION_NO As Int32 = Val(grdICTBRAN1.ActiveRow.Cells("COLLECTION_NO").Value & "")
            Dim COLLECTION_NAME As String = grdICTBRAN1.ActiveRow.Cells("COLLECTION_NAME").Value & ""
            Dim dvw As DataView = DirectCast(grdPOTPMGM2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "COLLECTION_NO = " & CStr(COLLECTION_NO)
            grdPOTPMGM2.Text = "Styles defined to Collection " & CStr(COLLECTION_NO) & ":" & COLLECTION_NAME
        End If
    End Sub
#End Region

#Region "grdPOTPMGM9"

    Private Sub grdPOTPMGM9_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTPMGM9.AfterRowActivate
        If Not grdPOTPMGM9.ActiveRow.IsDataRow Then Exit Sub

        With grdPOTPMGM9.DisplayLayout.Bands(0)
            If grdPOTPMGM9.ActiveRow.IsAddRow Then
                .Columns("PROGRAM_COMMENT").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdPOTPMGM9.ActiveCell = grdPOTPMGM9.ActiveRow.Cells("PROGRAM_COMMENT")
                grdPOTPMGM9.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("PROGRAM_COMMENT").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdPOTPMGM9_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTPMGM9.BeforeRowUpdate
        With grdPOTPMGM9
            If Not e.Cancel Then
                If e.Row.Cells("PROGRAM_NO").Text = "" Then
                    .ActiveRow.Cells("PROGRAM_NO").Value = Absx1.CtlFor("PROGRAM_NO").Text
                    .ActiveRow.Cells("STYLE_NO").Value = grdPOTPMGM2.ActiveRow.Cells("STYLE_NO").Value
                    .ActiveRow.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                    .ActiveRow.Cells("INIT_DATE").Value = DATETIME_STAMP
                End If
            End If
        End With
    End Sub

#End Region

#Region "grdPOTPMGMO"

    Private Sub grdPOTPMGMO_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTPMGMO.AfterCellUpdate
        If Not e.Cell.Row.IsDataRow Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "USER_ID"
 
                LookUp("ASTUSER1", e.Cell.Row.Cells("USER_ID").Value)
                If cdr IsNot Nothing Then
 
                Else
                    grdPOTPMGMO.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If
        End Select
    End Sub

    Private Sub grdPOTPMGMO_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTPMGMO.AfterRowActivate

        If Not grdPOTPMGMO.ActiveRow.IsDataRow Then Exit Sub

        With grdPOTPMGMO.DisplayLayout.Bands(0)
            If grdPOTPMGMO.ActiveRow.IsAddRow Then
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdPOTPMGMO.ActiveCell = grdPOTPMGMO.ActiveRow.Cells("USER_ID")
                grdPOTPMGMO.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdPOTPMGMO_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTPMGMO.AfterRowUpdate
        grdPOTPMGMO.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Private Sub grdPOTPMGMO_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTPMGMO.BeforeExitEditMode
        If grdPOTPMGMO.ActiveCell Is Nothing Then Exit Sub
        If Not grdPOTPMGMO.ActiveRow.IsDataRow Then Exit Sub
        With grdPOTPMGMO.ActiveCell
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

    Private Sub grdPOTPMGMO_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTPMGMO.BeforeRowsDeleted

    End Sub

     
    Private Sub grdPOTPMGMO_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTPMGMO.BeforeRowUpdate
        With grdPOTPMGMO
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
                '?     .ActiveRow.Cells("PROGRAM_NO").Value = Absx1.CtlFor("PROGRAM_NO").Text

            End If
        End With
    End Sub

    Private Sub grdPOTPMGMO_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTPMGMO.ClickCellButton

        If grdPOTPMGMO.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "USER_ID"
        End Select
        grdClickCellButton(grdPOTPMGMO, sql_where, True)

    End Sub

#End Region

    Private Sub grdPOTPMGMX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTPMGMX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("PROGRAM_NO").Text = e.Row.Cells("PROGRAM_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        ASCMAIN1.sql = sqlPOTPMGMX
        Dim PROGRAM_NO As String = Absx1.txtFor("PROGRAM_NO").Text
        If optShow.Value = "A" And PROGRAM_NO = "" Then
            grdPOTPMGMX.Text = "All Programs"
        ElseIf optShow.Value = "M" Then
            ASCMAIN1.sql &= " and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            grdPOTPMGMX.Text = "Programs entered or modified by Me"
        ElseIf optShow.Value = "C" Or PROGRAM_NO <> "" Then
            ASCMAIN1.sql &= " and PROGRAM_NO = '" & PROGRAM_NO & "'"
            grdPOTPMGMX.Text = "Programs associated with " & PROGRAM_NO
        End If
        Fill_Records("POTPMGMX")
        Sort_grdColumns(grdPOTPMGMX, "PROGRAM_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub email_Quote(tempFileName As String)
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        ATTACHMENTs.Add(tempFileName & ".pdf", ASCMAIN1.Folders("Temp") & tempFileName & ".pdf")

        Dim SUBJECT As String = "Program Sheet"
        Dim PFX As String = ""

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        If CUST_CODE <> "" Then
            EMAIL_ADDRESSs.Add(rowARTCUST1.Item("CUST_EMAIL") & "", rowARTCUST1.Item("CUST_CONTACT") & "")
        End If

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                SUBJECT, "POTPMGM1", False, True, CUST_CODE, CUST_NAME, "Customer")
        If SEND_NO <> "" Then
            TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "PGMEML", "Program Sheet emailed", SEND_NO)
        End If
    End Sub

    Private Sub optShow_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShow.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub

    Sub Get_Images(STYLE_NO As String)
        Dim I As New List(Of System.Drawing.Bitmap)

        Dim IMAGE_FOLDER As String = images_folder & "\COLUMN_NAME\PROGRAM_NO\STYLE_NO\" & STYLE_NO
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

        If images.ContainsKey(STYLE_NO) Then
            images(STYLE_NO) = I
        Else
            images.Add(STYLE_NO, I)
        End If
    End Sub

  
    Private Sub tabPOTPMGMX_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabPOTPMGMX.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabPOTPMGMX()
    End Sub

    Sub Setup_tabPOTPMGMX()
        With UltraExplorerBar1
            .Groups("Screen Control").Visible = (tabPOTPMGMX.SelectedTab.Key = "Programs")
            .Groups("Show Programs").Visible = (tabPOTPMGMX.SelectedTab.Key = "Programs")
            .Groups("Filters").Visible = (tabPOTPMGMX.SelectedTab.Key = "Calendar")
        End With

        spl.Panel1Collapsed = (tabPOTPMGMX.SelectedTab.Key = "Calendar")
    End Sub
  
    Private Sub grdPOTPMGMO_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTPMGMO.InitializeRow

        If e.Row.IsAddRow Then
            e.Row.Cells("USER_ID").Appearance.ForeColor = Color.Empty
            Exit Sub
        End If
        Dim USER_ID As String = e.Row.Cells("USER_ID").Value & ""
        If APPR_STATUS_CODE_ForeColors.ContainsKey(USER_ID) Then
            e.Row.Cells("USER_ID").Appearance.ForeColor = APPR_STATUS_CODE_ForeColors(USER_ID)
        End If
    End Sub
     
     
    Private Sub chkShowAllBrands_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAllBrands.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim dvw As DataView = DirectCast(grdICTBRAN1.DataSource, DataTable).DefaultView
        If chkShowAllBrands.Checked Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "SEL = '1' OR COLLS <> 0"
        End If
    End Sub

    Private Sub btnMessage_Click(sender As Object, e As EventArgs) Handles btnMessage.Click
        Using F As New TAC.TAFPMGMM("", Nothing, Me, True)
            F.EntryMode = "N"
            F.CONV_SUBJECT = "Kohl's - cost reduction suggestions"
            Dim CONV_TOPIC_NO As String = grdPOTCTOP1.ActiveRow.Cells("CONV_TOPIC_NO").Value
            F.CONV_TOPIC_NO = CONV_TOPIC_NO
            F.CONV_NOTES = ""




            F.ShowDialog()
            If F.result = "U" Then
                ' TAC.DECMAIN1.Fill_DETJOBMP(JOB_NO, Me, grdDETJOBMP, , , , chkPauseAutoPlan.Checked)
            End If
            F.Dispose()
        End Using
    End Sub

    Private Sub grdPOTCTOP1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTCTOP1.AfterRowActivate
        Setup_TATCONV1()
    End Sub

    Sub Setup_TATCONV1()
        If grdPOTCTOP1.ActiveRow IsNot Nothing AndAlso grdPOTCTOP1.ActiveRow.IsDataRow AndAlso Not grdPOTCTOP1.ActiveRow.IsAddRow Then
            Dim dvw As DataView = DirectCast(grdTATCONV1.DataSource, DataTable).DefaultView
            Dim CONV_TOPIC_NO As String = grdPOTCTOP1.ActiveRow.Cells("CONV_TOPIC_NO").Value
            Dim CONV_TOPIC_DESC As String = grdPOTCTOP1.ActiveRow.Cells("CONV_TOPIC_DESC").Value
            dvw.RowFilter = "TABLE_KEY = '" & CONV_TOPIC_NO & "'"
            grdTATCONV1.Visible = True
            grdTATCONV1.Rows.ExpandAll(True)
            grdTATCONV1.Text = "Conversation Log for " & CONV_TOPIC_DESC
            Sort_grdColumns(grdTATCONV1, "CONV_DATE")
        Else
            grdTATCONV1.Visible = False
        End If
    End Sub
    Private Sub grdPOTCTOP1_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdPOTCTOP1.BeforeRowUpdate
        With grdPOTCTOP1
            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PROGRAM_NO").Text = "" Then
                    .ActiveRow.Cells("PROGRAM_NO").Value = Absx1.CtlFor("PROGRAM_NO").Text
                    .ActiveRow.Cells("CONV_TOPIC_NO").Value = ASCMAIN1.Next_Control_No("POTCTOP1.CONV_TOPIC_NO")
                    .ActiveRow.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                    .ActiveRow.Cells("INIT_DATE").Value = Now + ASCMAIN1.NowTSD
                    .ActiveRow.Cells("LAST_OPER").Value = .ActiveRow.Cells("INIT_OPER").Value
                    .ActiveRow.Cells("LAST_DATE").Value = .ActiveRow.Cells("INIT_DATE").Value
                Else
                    .ActiveRow.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                    .ActiveRow.Cells("LAST_DATE").Value = Now + ASCMAIN1.NowTSD
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTCTOP1_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdPOTCTOP1.ClickCellButton

        If EntryMode = "N" Or EntryMode = "E" Then
        Else
            Exit Sub
        End If

        'ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PROGRAM_CATGY_CODE")

        'If ASCMAIN1.CodeSelector.SQL <> "" Then
        '    ASCMAIN1.CodeSelector.MultipleSelections = True
        '    '  ASCMAIN1.CodeSelector.UseDataFromTable = tbl
        '    Dim F As New ASFCODE1
        '    F.ShowDialog()
        '    F.Dispose()
        '    If ASCMAIN1.CodeSelector.Selections <> 0 Then

        '        Dim CONV_TOPIC_CATGYS As String = ""
        '        For Each PROGRAM_CATGY_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
        '            CONV_TOPIC_CATGYS &= "," & PROGRAM_CATGY_CODE
        '        Next
        '        e.Cell.Row.Cells("CONV_TOPIC_CATGYS").Value = Mid(CONV_TOPIC_CATGYS, 2)
        '        e.Cell.Row.Update()
        '    End If
        'End If
    End Sub

    Private Sub grdTATCONV1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdTATCONV1.AfterRowActivate
        If EntryMode = "N" Or EntryMode = "E" Then
            If grdTATCONV1.ActiveRow.Cells("CONV_ACK_REQD").Value & "" = "1" And grdTATCONV1.ActiveRow.Cells("CONV_ACK_IND").Value & "" <> "1" Then
                grdTATCONV1.ActiveRow.Cells("CONV_ACK_IND").Column.CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End If
    End Sub

    Private Sub grdTATCONV1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdTATCONV1.InitializeRow
        e.Row.Cells("WHO_AND_WHEN").Value = e.Row.Cells("INIT_OPER").Value & vbCrLf & e.Row.Cells("INIT_DATE").Text
        If e.Row.Cells("CONV_ACK_REQD").Value & "" = "1" Then
            e.Row.Cells("CONV_ACK_IND").Appearance.BackColor = Color.Empty
        Else
            e.Row.Cells("CONV_ACK_IND").Appearance.BackColor = Color.LightGray
        End If
    End Sub

    Private Sub grdTATCONV1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATCONV1.InitializeLayout

    End Sub

    Private Sub grdTATCONV1_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdTATCONV1.BeforeRowActivate
        e.Row.Band.Columns("CONV_ACK_IND").CellActivation = UltraWinGrid.Activation.NoEdit
    End Sub

    Private Sub btnCreateTopics_Click(sender As Object, e As EventArgs) Handles btnCreateTopics.Click

        If MsgBox("Topics will be created for any Selected Category that does not already have a dedicated topic." _
                  & vbCrLf & vbCrLf & "OK to Continue?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then

            Dim HITS As Integer = 0
            Dim MISSES As Integer = 0

            For Each row As DataRow In dst.Tables("POTPMGMC").Select("SEL='1'")
                Dim PROGRAM_CATGY_CODE As String = row.Item("PROGRAM_CATGY_CODE")
                Dim rows() As DataRow = dst.Tables("POTCTOP1").Select("PROGRAM_CATGY_CODE = '" & PROGRAM_CATGY_CODE & "'")
                If rows.Length = 0 Then
                    Dim rowPOTCTOP1 As DataRow = dst.Tables("POTCTOP1").NewRow
                    With rowPOTCTOP1
                        .Item("PROGRAM_NO") = PROGRAM_NO
                        .Item("CONV_TOPIC_NO") = ASCMAIN1.Next_Control_No("POTCTOP1.CONV_TOPIC_NO")
                        .Item("CONV_TOPIC_DESC") = row.Item("PROGRAM_CATGY_DESC")
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("LAST_OPER") = .Item("INIT_OPER")
                        .Item("LAST_DATE") = .Item("INIT_DATE")
                        .Item("PROGRAM_CATGY_CODE") = PROGRAM_CATGY_CODE
                    End With
                    dst.Tables("POTCTOP1").Rows.Add(rowPOTCTOP1)
                    HITS += 1
                Else
                    MISSES += 1
                End If
            Next
            MsgBox("Topics Created = " & CStr(HITS) & vbCrLf & "Topics Already Created = " & CStr(MISSES), MsgBoxStyle.OkOnly, "Verification")
        End If
    End Sub

    Private Sub grdPOTPMGMC_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdPOTPMGMC.ClickCellButton

        If EntryMode = "N" Or EntryMode = "E" Then
        Else
            Exit Sub
        End If

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("USER_ID")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim USER_IDs As String = ""
            For Each row As DataRow In dst.Tables("POTPMGMO").Select("")
                Dim USER_ID As String = row.Item("USER_ID")
                USER_IDs &= ",'" & USER_ID & "'"
            Next
            If USER_IDs = "" Then
                MsgBox("No Users associated with this Program", MsgBoxStyle.OkOnly, "Cannot Assign Users")
                Exit Sub
            Else
                ASCMAIN1.CodeSelector.Custom_sql_where = "USER_ID in (" & Mid(USER_IDs, 2) & ")"
                ASCMAIN1.CodeSelector.SQL &= " where " & ASCMAIN1.CodeSelector.Custom_sql_where
            End If

            '  ASCMAIN1.CodeSelector.UseDataFromTable = tbl
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then

                Dim ASSIGNED_TO As String = ""
                For Each USER_ID As String In ASCMAIN1.CodeSelector.SelectedCodes
                    ASSIGNED_TO &= "," & USER_ID
                Next
                e.Cell.Row.Cells("ASSIGNED_TO").Value = Mid(ASSIGNED_TO, 2)
                e.Cell.Row.Update()
            End If
        End If
    End Sub
     
    Private Sub cmdChangeProgramImage_Click(sender As Object, e As EventArgs) Handles cmdChangeProgramImage.Click
        Dim FILENAME As String = ""
        Dim FOLDER As String = ASCMAIN1.Folders("Images") & "COLUMN_NAME\PROGRAM_NO\"

        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Image to use as the Program Graphic"
            ' openFileDialog1.Filter = "png files (*.png)|*.png"
            openFileDialog1.Filter = "jpg files (*.jpg)|*.jpg|png files (*.png)|*.png"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            Dim FOLDER_NAME As String = ASCMAIN1.Folders("Images") & "COLUMN_NAME\PROGRAM_NO\"

            My.Computer.FileSystem.CopyFile(FILENAME, FOLDER_NAME & PROGRAM_NO & ".png", True)
            Load_Program_Graphic()
        End If
    End Sub

    Sub Load_Program_Graphic()
        Dim FOLDER As String = ASCMAIN1.Folders("Images") & "COLUMN_NAME\PROGRAM_NO\"
        Dim FILENAME As String = PROGRAM_NO & ".png"
        If Not My.Computer.FileSystem.FileExists(FOLDER & FILENAME) Then
            FILENAME = Replace(FILENAME, ".png", ".jpg")
        End If
        If Not My.Computer.FileSystem.FileExists(FOLDER & FILENAME) Then
            FILENAME = ""
        End If
        If FILENAME = "" Then
            ProgramImage.Visible = False
        Else
            Dim b As System.Drawing.Bitmap = ASCMAIN1.Get_Image(FOLDER, FILENAME) ' , False, , , Nothing)
            ProgramImage.Image = b
            ProgramImage.ScaleImage = ScaleImage.Always
            ProgramImage.MaintainAspectRatio = True

            ProgramImage.Visible = True
        End If

    End Sub

    Private Sub grdPOTCTOP1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTCTOP1.InitializeLayout

    End Sub
 
    Private Sub btnImportFrom_Click(sender As Object, e As EventArgs) Handles btnImportFrom.Click

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PROGRAM_NO")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            '  ASCMAIN1.CodeSelector.UseDataFromTable = tbl
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections = 1 Then
                Dim PROGRAM_NO_to_import_from As String = ASCMAIN1.CodeSelector.SelectedCode

                ASCMAIN1.sql = "Select * from POTPMGMO where PROGRAM_NO = '" & PROGRAM_NO_to_import_from & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    row.Item("PROGRAM_NO") = PROGRAM_NO
                    dst.Tables("POTPMGMO").Rows.Add(row.ItemArray)
                Next

                'ASCMAIN1.sql = "Select * from POTPMGMC where PROGRAM_NO = '" & PROGRAM_NO_to_import_from & "'"
                'For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                '    row.Item("PROGRAM_NO") = PROGRAM_NO
                '    dst.Tables("POTPMGMO").Rows.Add(row.ItemArray)
                'Next

            End If
        End If
    End Sub
End Class