Imports System
Public Class WBFIMGC1
    Dim InquiryOnly As Boolean = False
    Dim MasterImages As List(Of String)
    Dim WebImages As List(Of String)
    Dim WebUpdated As Boolean = False
#Region "ABS Standard Routines"

    ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst

            ASCMAIN1.sql = ""

            Create_TDA(.Tables.Add, "ICTWPICS", "*")
            .Tables("ICTWPICS").Columns.Add("FACTORY", GetType(System.String))
            .Tables("ICTWPICS").Columns.Add("STYLE_STATUS", GetType(System.String))
            .Tables("ICTWPICS").Columns.Add("WEB", GetType(System.String))
            Create_TDA(.Tables.Add, "ICTWPICM", "*")
            Create_TDA(.Tables.Add, "ICTWPICW", "*")

            Create_TDA(.Tables.Add, "WBTSTYL1", "*")
            Create_TDA(.Tables.Add, "WBTSTYL2", "*")

        End With

        grdICTWPICS.DataSource = dst.Tables("ICTWPICS")
        grdICTWPICW.DataSource = dst.Tables("ICTWPICW")
        grdICTWPICM.DataSource = dst.Tables("ICTWPICM")
        grdNotes.DataSource = dst.Tables("ICTWPICS")

        Create_Summary(grdICTWPICS, "STYLE_CODE", "Count", "", "###,##0")

        Sort_grdColumns(grdICTWPICS, "STYLE_CODE, COLOR_CODE")
        Sort_grdColumns(grdICTWPICW, "STYLE_CODE, COLOR_CODE")
        Sort_grdColumns(grdICTWPICM, "STYLE_CODE, COLOR_CODE")

        'grdSOTRUSSE.DisplayLayout.UseFixedHeaders = True
        'With grdSOTRUSSE.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "ORDR_QTY", "ORDR_UNIT_PRICE"}
        '        .Columns(COLUMN_NAME).Header.Fixed = True
        '    Next
        'End With

        tab.Visible = False
        'grdSOTORDRX.Parent = tab.Parent

        'Get Folder location from parameters.
        Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")
        If IsNothing(rowWBTPARM1) Then
            MsgBox("Images Not Set-Up In Web Parameters File!", MsgBoxStyle.Critical, "Parameters")
        Else
            If rowWBTPARM1.Item("WB_PARM_MASTER_IMAGES").ToString.Length = 0 Then
                MsgBox("Master Images Folder Not Set-Up In Web Parameters File!", MsgBoxStyle.Critical, "Parameters")
            Else
                txtWB_PARM_MASTER_IMAGES.Text = rowWBTPARM1.Item("WB_PARM_MASTER_IMAGES").ToString
            End If
            If rowWBTPARM1.Item("WB_PARM_WEB_IMAGES").ToString.Length = 0 Then
                MsgBox("Web Images Folder Not Set-Up In Web Parameters File!", MsgBoxStyle.Critical, "Parameters")
            Else
                txtWB_PARM_WEB_IMAGES.Text = rowWBTPARM1.Item("WB_PARM_WEB_IMAGES").ToString
            End If
            If rowWBTPARM1.Item("WB_PARM_FINAL_IMAGES").ToString.Length = 0 Then
                MsgBox("Final Images Folder Not Set-Up In Web Parameters File!", MsgBoxStyle.Critical, "Parameters")
            Else
                txtWB_PARM_FINAL_IMAGES.Text = rowWBTPARM1.Item("WB_PARM_FINAL_IMAGES").ToString
                txtWB_PARM_DISC_IMAGES.Text = rowWBTPARM1.Item("WB_PARM_FINAL_IMAGES").ToString
            End If
        End If

        Dim lbli As New Text.StringBuilder() With {.Length = 0}
        lbli.AppendLine("1) Make Sure Images Are In Web Folder.")
        lbli.AppendLine("2) Change configuration.")
        lbli.AppendLine("3) Load Records.")
        lbli.AppendLine("4) Remove Prompt.")
        lbli.AppendLine("5) Hit Auto-complete.")
        lbli.AppendLine("6) Right-click all or select items for web.")
        lbli.AppendLine("7) Remove Dupes and Deal with it.")
        lbli.AppendLine("8) Send to web.")
        lbli.AppendLine("9) Update.")
        lblInstructions.Text = lbli.ToString
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit"
                Dim x As Boolean = False
                If x = False Then
                    EMsg &= vbCr & "Some Kind Of Error."
                End If
                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text) Then
                        Exit Sub
                    End If
                End If
            Case "Update"
                If txtWB_PARM_MASTER_IMAGES.Text.Length = 0 Or
                    txtWB_PARM_WEB_IMAGES.Text.Length = 0 Or
                    txtWB_PARM_FINAL_IMAGES.Text.Length = 0 Or
                    txtWB_PARM_DISC_IMAGES.Text.Length = 0 Then
                    MsgBox("Folder Locations In Config Can Not Be Blank", MsgBoxStyle.Critical, "Config Issue")
                End If
                Dim i As Integer = TryRemoveDups()
                If i > 0 Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Removals Happened"
                    Dim iMSG As New System.Text.StringBuilder
                    iMSG.AppendLine("Are You Sure You Want to Update?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg &= vbCr & "Update Canceled"
                    End If

                End If
                'Mode_Settings(False)
            Case "Cancel"

            Case "Load Records"
                'Check to see if FinalImageLocation is empty and if it's not ask if they want to clean it out.

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Private Function TryRemoveDups() As Integer
        Dim SelCnt As Integer = Remove_Dups()
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Un-Selections"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine(SelCnt & " Rows Un-Selected Because they Were Already On The Web")
        iMSG.AppendLine("You Will Need To Re-Run Send To Web Before Updating.")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        Return SelCnt
    End Function

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)
            Case "Cancel"
                Call Mode_Settings(False)
            Case "Load Records"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Potentials"
                'Call BuildPotentials()
            Case "Send To Web"
                Call SendToWeb()
            Case "Remove Dups"
                Dim i As Integer = TryRemoveDups()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Edit").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Done").Visible = False

                .Groups("Screen Control").Items("Update").Visible = ScreenMode
                .Groups("Screen Control").Items("Cancel").Visible = ScreenMode
                .Groups("Screen Control").Items("Send To Web").Visible = ScreenMode
                .Groups("Screen Control").Items("Remove Dups").Visible = ScreenMode

                .Groups("Screen Control").Items("Potentials").Visible = False
                'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne" Then
                '    .Groups("Screen Control").Items("Potentials").Visible = True
                'Else
                '    .Groups("Screen Control").Items("Potentials").Visible = False
                'End If

            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdSOTORDRX.Visible = Not tf
        'With grdSOTRUSSE.DisplayLayout.Override
        '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .AllowDelete = DefaultableBoolean.False
        '    .AllowUpdate = DefaultableBoolean.True
        'End With
        'For i As Integer = 0 To grdSOTRUSSE.DisplayLayout.Bands(0).Columns.Count - 1
        '    grdSOTRUSSE.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        'Next i
        'For Each COLNAME As String In New String() {"NEW_QTY", "NEW_UPC", "NEW_SKU", "NEW_COLOR_CODE", "NEW_ORDR_UNIT_PRICE"}
        '    grdSOTRUSSE.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        'Next
        'For Each COLNAME As String In New String() {"NEW_QTY", "NEW_UPC", "NEW_SKU", "NEW_COLOR_CODE", "NEW_ORDR_UNIT_PRICE"}
        '    grdSOTRUSSE.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        'Next

        'If Not ScreenMode Then
        '    RefreshSOTORDRX()
        'End If

        'grdICTWPICS.DisplayLayout.Bands(0).Columns("No Img").Hidden = False
        'grdNotes.DisplayLayout.Bands(0).Columns("IMG_NOT_FOUND").Hidden = False

    End Sub

    Sub Clear_Record()
        dst.Tables("ICTSTYL1").Rows.Clear()
    End Sub

    Private Sub ReBuildPictures()
        BeginTrans()

        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("Truncate table ICTWPICS")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        SQLS.Length = 0
        SQLS.AppendLine("Truncate table ICTWPICM")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        SQLS.Length = 0
        SQLS.AppendLine("Truncate table ICTWPICW")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        SQLS.Length = 0
        SQLS.AppendLine("INSERT INTO ICTWPICS")
        SQLS.AppendLine(" (STYLE_CODE, COLOR_CODE, COLOR_CODE_LONG, STYLE_COLOR_IMAGE_NAME)")
        SQLS.AppendLine(" SELECT ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_CODE_LONG, NULL AS STYLE_COLOR_IMAGE_NAME ")
        SQLS.AppendLine(" FROM ICTSTYC1, ICTCOLR1, ICTSTAT2")
        SQLS.AppendLine(" WHERE ICTSTYC1.COLOR_CODE = ICTCOLR1.COLOR_CODE")
        SQLS.AppendLine(" AND ICTSTYC1.STYLE_CODE = ICTSTAT2.STYLE_CODE (+)")
        SQLS.AppendLine(" AND ICTSTYC1.COLOR_CODE = ICTSTAT2.COLOR_CODE (+)")
        'SQLS.AppendLine(" AND ICTSTYC1.STYLE_CODE = 'MTF19430'")
        SQLS.AppendLine(" GROUP BY ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_CODE_LONG")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        If Not chkINCLDISC.Checked Then
            SQLS.Length = 0
            SQLS.AppendLine("DELETE FROM ICTWPICS")
            SQLS.AppendLine(" WHERE (STYLE_CODE, COLOR_CODE)")
            SQLS.AppendLine(" IN")
            SQLS.AppendLine(" (")
            SQLS.AppendLine("SELECT S1.STYLE_CODE, C1.COLOR_CODE")
            SQLS.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1, ICTSTAT2 S2")
            SQLS.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE")
            SQLS.AppendLine("AND C1.STYLE_CODE = S2.STYLE_CODE (+)")
            SQLS.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
            SQLS.AppendLine("AND (S1.STYLE_STATUS = 'D' OR S1.STYLE_STATUS = 'N')")
            SQLS.AppendLine("HAVING SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) <= 0 AND SUM(NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0)) <=0")
            SQLS.AppendLine("GROUP BY S1.STYLE_CODE, C1.COLOR_CODE")
            SQLS.AppendLine(" )")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
        End If


        CommitTrans()

        Fill_Records("ICTWPICS", , True, "SELECT * FROM ICTWPICS")
        For Each rowICTWPICS As DataRow In dst.Tables("ICTWPICS").Select()
            Dim SQLS1 As New System.Text.StringBuilder() With {.Length = 0}
            SQLS1.AppendLine(String.Format("Select NVL(STYLE_STATUS,'A') AS STYLE_STATUS from ICTSTYL1 where STYLE_CODE = '{0}'", rowICTWPICS.Item("STYLE_CODE")))
            ASCMAIN1.sql = SQLS1.ToString()
            Dim STYLE_STATUS As String = ASCDATA1.GetDataValue
            rowICTWPICS.Item("STYLE_STATUS") = STYLE_STATUS
        Next

        SQLS.Length = 0
        SQLS.AppendLine("Select Count(*) as REC_CNT from ICTWPICS")
        ASCMAIN1.sql = SQLS.ToString()
        Dim TotalRecs As Int64 = Val(ASCDATA1.GetDataValue)
        Dim CurrentRec As Int64 = 0
        Dim RecFound As Int64 = 0
        MasterImages = (From chkFile In IO.Directory.EnumerateFiles(txtWB_PARM_MASTER_IMAGES.Text, "*.jpg", IO.SearchOption.TopDirectoryOnly)).ToList()
        For i As Int64 = 0 To MasterImages.Count - 1
            MasterImages(i) = Replace(MasterImages(i), txtWB_PARM_MASTER_IMAGES.Text, "").ToUpper
        Next
        WebImages = (From chkFile In IO.Directory.EnumerateFiles(txtWB_PARM_WEB_IMAGES.Text, "*.jpg", IO.SearchOption.TopDirectoryOnly)).ToList()
        For i As Int64 = 0 To WebImages.Count - 1
            WebImages(i) = Replace(WebImages(i), txtWB_PARM_WEB_IMAGES.Text, "").ToUpper
        Next
        For Each rowICTWPICS As DataRow In dst.Tables("ICTWPICS").Select()
            CurrentRec += 1
            If FillImageInfo(rowICTWPICS) Then
                RecFound += 1
            End If
            ASCMAIN1.Progress(String.Format("{0} of {1} Searched. {2} Found.", CurrentRec, TotalRecs, RecFound))
            'If RecFound > 1000 Then Exit For
        Next
    End Sub

    Sub Load_Record()

        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Rebuild Picture Table"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("Do You Want To Rebuild The Picture Table?")
        iMSG.AppendLine("This Will Remove All Prior Pictures And")
        iMSG.AppendLine("Re-Build Them From The File System.")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            ReBuildPictures()
        Else
            Fill_Records("ICTWPICS", , True, "SELECT * FROM ICTWPICS")
            Fill_Records("ICTWPICM", , True, "SELECT * FROM ICTWPICM")
            Fill_Records("ICTWPICW", , True, "SELECT * FROM ICTWPICW")
        End If


        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        EnforceConstraints(True)

        ASCMAIN1.Progress("Fetching Factory Codes")
        For Each rowICTWPICS As DataRow In dst.Tables("ICTWPICS").Select()
            rowICTWPICS.Item("FACTORY") = GetFactoryCode(rowICTWPICS.Item("STYLE_CODE").ToString)
        Next
        ASCMAIN1.Progress("")

        'For Each grow As UltraWinGrid.UltraGridRow In grdICTWPICS.Rows
        '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", grow.Cells.Item("STYLE_CODE").Text)
        '    If Not IsNothing(rowICTSTYL1) Then
        '        grow.Cells.Item("STYLE_CODE").ToolTipText = rowICTSTYL1.Item("STYLE_DESC").ToString
        '        grow.Cells.Item("COLOR_CODE").ToolTipText = rowICTSTYL1.Item("STYLE_DESC").ToString
        '        grow.Cells.Item("COLOR_CODE").ToolTipText = rowICTSTYL1.Item("STYLE_DESC").ToString
        '    Else
        '        grow.Cells.Item("STYLE_CODE").ToolTipText = ""
        '    End If
        'Next

        'If EntryMode = "N" Then
        'Else
        '    'dst.AcceptChanges()
        'End If
        Dim dvw As DataView = DirectCast(grdICTWPICS.DataSource, DataTable).DefaultView
        dvw.RowFilter = "WEB_IMAGE_NAME <> '' OR MASTER_IMAGE_NAME <> ''"

        Setup_IMAGES()
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()
        'For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        'Next
        'Call CommitTrans("Order / Quote Deleted")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Update_Record_TDA("ICTWPICS")
        Update_Record_TDA("ICTWPICM")
        Update_Record_TDA("ICTWPICW")
        'Update_Record_TDA("WBTSTYL1", "DELETE FROM WBTSTYL1")
        'Update_Record_TDA("WBTSTYL2", "DELETE FROM WBTSTYL2")
        Update_Record_TDA("WBTSTYL1")
        Update_Record_TDA("WBTSTYL2")

        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("UPDATE WBTPARM1")
        SQLS.AppendLine(String.Format(" SET WB_PARM_MASTER_IMAGES = '{0}',", txtWB_PARM_MASTER_IMAGES.Text))
        SQLS.AppendLine(String.Format(" WB_PARM_WEB_IMAGES = '{0}',", txtWB_PARM_WEB_IMAGES.Text))
        SQLS.AppendLine(String.Format(" WB_PARM_FINAL_IMAGES = '{0}'", txtWB_PARM_FINAL_IMAGES.Text))
        SQLS.AppendLine(" WHERE WB_PARM_KEY = 'Z'")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()



        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        'Print_Report_Begin()
        'Generate_Report("SORORDRO")
        'Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWPICS, "SSBB", "Show Filter", "Show GroupBox", "Select All For Web", "Select None For Web")
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
            Case "grdICTWPICS"
                e.Tool.ToolbarsManager.Tools("Select All For Web").SharedProps.Visible = True
                e.Tool.ToolbarsManager.Tools("Select None For Web").SharedProps.Visible = True
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

        Select Case e.Tool.Key
            Case "Select All For Web"
                For Each grow As UltraWinGrid.UltraGridRow In grdICTWPICS.Rows
                    If grow.VisibleIndex <> -1 Then
                        grow.Cells.Item("WEB").Value = "1"
                    End If
                Next
                grdICTWPICS.UpdateData()
            Case "Select None For Web"
                For Each grow As UltraWinGrid.UltraGridRow In grdICTWPICS.Rows
                    If grow.VisibleIndex <> -1 Then
                        grow.Cells.Item("WEB").Value = ""
                    End If

                Next
                grdICTWPICS.UpdateData()
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                'FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
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
    Private Function FillImageInfo(ByRef rowICTWPICS As DataRow) As Boolean
        Dim RetVal As Boolean = False
        Dim FileMatch As String = ""
        Dim STYLE_CODE As String = rowICTWPICS.Item("STYLE_CODE").ToString
        Dim COLOR_CODE As String = rowICTWPICS.Item("COLOR_CODE").ToString
        Dim COLOR_CODE_LONG As String = rowICTWPICS.Item("COLOR_CODE_LONG").ToString
        'If STYLE_CODE = "MT22660" Then Stop

        For imgpass As Integer = 1 To 2
            If imgpass = 1 Then
                FileMatch = String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE).ToUpper
                If Not MasterImages.Contains(FileMatch) Then
                    FileMatch = String.Format("{0}{1}.jpg", STYLE_CODE, COLOR_CODE).ToUpper
                    If Not MasterImages.Contains(FileMatch) Then
                        FileMatch = ""
                    End If
                End If
                If FileMatch.Length > 0 Then
                    rowICTWPICS.Item("MASTER_IMAGE_NAME") = FileMatch
                    RetVal = True
                End If
            Else
                FileMatch = String.Format("{0}{1}.jpg", STYLE_CODE, COLOR_CODE).ToUpper
                If Not WebImages.Contains(FileMatch) Then
                    FileMatch = String.Format("{0}{1}.jpg", STYLE_CODE, COLOR_CODE_LONG).ToUpper
                    If Not WebImages.Contains(FileMatch) Then
                        FileMatch = String.Format("{0}.jpg", STYLE_CODE).ToUpper
                        If Not WebImages.Contains(FileMatch) Then
                            Dim results As List(Of String) = WebImages.FindAll(Function(value As String) value.StartsWith(STYLE_CODE & COLOR_CODE))
                            If results.Count > 0 Then
                                If results.Count = 1 Then
                                    FileMatch = results.Item(0)
                                Else
                                    FileMatch = "MULTI"
                                    For i As Integer = 1 To results.Count
                                        Dim rowICTWPICW As DataRow
                                        rowICTWPICW = dst.Tables("ICTWPICW").NewRow
                                        rowICTWPICW.Item("STYLE_CODE") = STYLE_CODE
                                        rowICTWPICW.Item("COLOR_CODE") = COLOR_CODE
                                        rowICTWPICW.Item("WEB_IMAGE_NAME") = results(i - 1)
                                        'dst.Tables("SOTCORDR").Rows.Add(rowSOTCORDR)
                                        dst.Tables("ICTWPICW").Rows.Add(rowICTWPICW)
                                    Next
                                End If
                            Else
                                If results.Count = 0 Then
                                    Dim results2 As List(Of String) = WebImages.FindAll(Function(value As String) value.StartsWith(STYLE_CODE))
                                    If results2.Count > 0 Then
                                        If results2.Count = 1 Then
                                            FileMatch = results2.Item(0)
                                        Else
                                            FileMatch = "MULTI"
                                            For i As Integer = 1 To results2.Count
                                                Dim rowICTWPICW As DataRow
                                                rowICTWPICW = dst.Tables("ICTWPICW").NewRow
                                                rowICTWPICW.Item("STYLE_CODE") = STYLE_CODE
                                                rowICTWPICW.Item("COLOR_CODE") = COLOR_CODE
                                                rowICTWPICW.Item("WEB_IMAGE_NAME") = results2(i - 1)
                                                'dst.Tables("SOTCORDR").Rows.Add(rowSOTCORDR)
                                                dst.Tables("ICTWPICW").Rows.Add(rowICTWPICW)
                                            Next
                                        End If
                                    Else
                                        FileMatch = ""
                                    End If
                                Else
                                    FileMatch = ""
                                End If
                            End If
                        End If
                    End If
                End If
                If FileMatch.Length > 0 Then
                    rowICTWPICS.Item("WEB_IMAGE_NAME") = FileMatch
                    RetVal = True
                End If
            End If

        Next
        Return RetVal
    End Function

    Sub Setup_IMAGES()
        'MT16782
        If grdICTWPICS.ActiveRow Is Nothing OrElse (Not grdICTWPICS.ActiveRow.IsDataRow Or grdICTWPICS.ActiveRow.IsAddRow) Then
            'grpSOTORDR3.Visible = False
        Else
            Dim STYLE_CODE As String = grdICTWPICS.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTWPICS.ActiveRow.Cells("COLOR_CODE").Value
            Dim dvw As DataView = DirectCast(grdICTWPICM.DataSource, DataTable).DefaultView
            dvw.RowFilter = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)

            Dim dvw2 As DataView = DirectCast(grdICTWPICW.DataSource, DataTable).DefaultView
            dvw2.RowFilter = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)

            Dim MASTER_IMAGE_NAME As String = grdICTWPICS.ActiveRow.Cells("MASTER_IMAGE_NAME").Value & ""
            If MASTER_IMAGE_NAME.Length > 0 Then
                picMaster.ImageLocation = txtWB_PARM_MASTER_IMAGES.Text & MASTER_IMAGE_NAME
            Else
                picMaster.ImageLocation = ""
            End If

            Dim WEB_IMAGE_NAME As String = grdICTWPICS.ActiveRow.Cells("WEB_IMAGE_NAME").Value & ""
            If WEB_IMAGE_NAME.Length > 0 Then
                picWeb.ImageLocation = txtWB_PARM_WEB_IMAGES.Text & WEB_IMAGE_NAME
            Else
                picWeb.ImageLocation = ""
            End If
            If grdICTWPICS.ActiveRow.Cells("PROCESSED").Value & "" = "1" Then
                If grdICTWPICS.ActiveRow.Cells.Item("IMG_SRC").Value & "" = "M" Then
                    chkImageMaster.Checked = True
                Else
                    chkImageMaster.Checked = False
                End If
                If grdICTWPICS.ActiveRow.Cells.Item("IMG_SRC").Value & "" = "W" Then
                    chkImageWeb.Checked = True
                Else
                    chkImageWeb.Checked = False
                End If
            Else
                chkImageMaster.Checked = False
                chkImageWeb.Checked = False
            End If

            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", grdICTWPICS.ActiveRow.Cells.Item("STYLE_CODE").Text)
            If Not IsNothing(rowICTSTYL1) Then
                grdICTWPICS.ActiveRow.Cells.Item("STYLE_CODE").ToolTipText = rowICTSTYL1.Item("STYLE_DESC").ToString
                grdICTWPICS.ActiveRow.Cells.Item("COLOR_CODE").ToolTipText = rowICTSTYL1.Item("STYLE_DESC").ToString
                grdICTWPICS.ActiveRow.Cells.Item("COLOR_CODE").ToolTipText = rowICTSTYL1.Item("STYLE_DESC").ToString
            Else
                grdICTWPICS.ActiveRow.Cells.Item("STYLE_CODE").ToolTipText = ""
            End If

        End If
    End Sub

    Private Function GetSelImage() As String
        Dim RetVal As String = ""
        If chkImageMaster.Checked Then
            RetVal = Replace(picMaster.ImageLocation, txtWB_PARM_MASTER_IMAGES.Text, "")
        End If
        If chkImageWeb.Checked Then
            RetVal = Replace(picWeb.ImageLocation, txtWB_PARM_WEB_IMAGES.Text, "")
        End If
        Return RetVal
    End Function

    Private Function GetSelPath() As String
        Dim RetVal As String = ""
        If chkImageMaster.Checked Then
            RetVal = txtWB_PARM_MASTER_IMAGES.Text
        End If
        If chkImageWeb.Checked Then
            RetVal = txtWB_PARM_WEB_IMAGES.Text
        End If
        Return RetVal
    End Function

    Private Function GetIMG_SRC() As String
        Dim RetVal As String = "O"
        If chkImageMaster.Checked Then
            RetVal = "M"
        End If
        If chkImageWeb.Checked Then
            RetVal = "W"
        End If
        Return RetVal
    End Function

    Private Function MakeFinal(ByRef curRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As Boolean
        Dim RetVal As Boolean = False
        Dim ProceedCopy As Boolean = True
        Dim STYLE_COLOR_IMAGE_NAME As String = GetSelImage()
        Dim FullSource As String = GetSelPath() & STYLE_COLOR_IMAGE_NAME
        Dim IMG_SRC As String = GetIMG_SRC()

        Dim FullDest As String = ""
        If isStyleColorActive(curRow.Cells.Item("STYLE_CODE").Text, curRow.Cells.Item("COLOR_CODE").Text) Then
            FullDest = String.Format("{0}{1}-{2}.jpg", txtWB_PARM_FINAL_IMAGES.Text, curRow.Cells.Item("STYLE_CODE").Text, curRow.Cells.Item("COLOR_CODE").Text)
        Else
            FullDest = String.Format("{0}{1}-{2}.jpg", txtWB_PARM_DISC_IMAGES.Text, curRow.Cells.Item("STYLE_CODE").Text, curRow.Cells.Item("COLOR_CODE").Text)
        End If

        If STYLE_COLOR_IMAGE_NAME.Length = 0 Then
            Return RetVal
            Exit Function
        End If
        If curRow.Cells.Item("COLOR_CODE").Text.Contains("/") Then
            Return RetVal
            Exit Function
        End If
        If IO.File.Exists(FullDest) Then
            If chkPrompOver.Checked Then
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Overwrite File"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine(String.Format("File {0} Already Exists In Destination.", STYLE_COLOR_IMAGE_NAME))
                iMSG.AppendLine("Do You Want To Over-Write It?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    IO.File.Delete(FullDest)
                Else
                    ProceedCopy = False
                End If
            Else
                IO.File.Delete(FullDest)
            End If
        End If
        If ProceedCopy Then
            System.IO.File.Copy(FullSource, FullDest)
            curRow.Cells.Item("PROCESSED").Value = "1"
            curRow.Cells.Item("STYLE_COLOR_IMAGE_NAME").Value = STYLE_COLOR_IMAGE_NAME
            curRow.Cells.Item("USER_NOTES").Value = txtNotes.Text
            curRow.Cells.Item("IMG_SRC").Value = IMG_SRC
            curRow.Cells.Item("IMG_NOT_FOUND").Value = "0"
            RetVal = True
        End If
        Return RetVal
    End Function

    Private Function isStyleColorActive(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As Boolean
        Dim RetVal As Boolean = False
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT NVL(STYLE_COLOR_STATUS,'A')")
        SQLS.AppendLine("FROM ICTSTYC1")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim STYLE_COLOR_STATUS As String = ASCDATA1.GetDataValue & String.Empty
        If STYLE_COLOR_STATUS = "A" Then
            RetVal = True
        End If
        Return RetVal
    End Function

    Private Function GetFactoryCode(ByVal STYLE_CODE As String) As String
        Dim RetVal As String = ""
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("SELECT VEND_SUPPLIER_ID")
        SQLS.AppendLine("FROM ICTSTYL1, APTVEND1")
        SQLS.AppendLine("WHERE ICTSTYL1.VEND_CODE = APTVEND1.VEND_CODE")
        SQLS.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        RetVal = ASCDATA1.GetDataValue
        Return RetVal
    End Function
#End Region

#Region "FormControls"
    Private Sub FilterMaster()
        If chkShowAllColors.Checked Then
            chkHideFinished.Checked = False
            chkNOIMG.Checked = False
            Dim dvw As DataView = DirectCast(grdICTWPICS.DataSource, DataTable).DefaultView
            dvw.RowFilter = ""
        Else
            Dim dvw As DataView = DirectCast(grdICTWPICS.DataSource, DataTable).DefaultView
            Dim RowFilter As String = "(WEB_IMAGE_NAME <> '' OR MASTER_IMAGE_NAME <> '')"
            If chkHideFinished.Checked Then
                RowFilter += " AND ISNULL(PROCESSED,'0') <> '1'"
            End If
            If chkNOIMG.Checked Then
                RowFilter += " AND ISNULL(IMG_NOT_FOUND,'0') <> '1'"
            End If
            dvw.RowFilter = RowFilter
        End If
    End Sub

    Private Sub chkShowAllColors_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowAllColors.CheckedChanged
        FilterMaster()
    End Sub

    Private Sub btnSelectImage_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectImage.Click
        MakeFinal(grdICTWPICS.ActiveRow)
    End Sub

    Private Sub btnWB_PARM_MASTER_IMAGES_Click(sender As System.Object, e As System.EventArgs) Handles btnWB_PARM_MASTER_IMAGES.Click
        Dim newFolder As String
        FolderBrowserDialog1.ShowDialog()
        newFolder = FolderBrowserDialog1.SelectedPath
        If newFolder.Length > 0 Then
            If Not newFolder.EndsWith("\") Then
                newFolder = newFolder & "\"
                txtWB_PARM_MASTER_IMAGES.Text = newFolder
            End If
        End If
    End Sub

    Private Sub btnWB_PARM_WEB_IMAGES_Click(sender As System.Object, e As System.EventArgs) Handles btnWB_PARM_WEB_IMAGES.Click
        Dim newFolder As String
        FolderBrowserDialog1.ShowDialog()
        newFolder = FolderBrowserDialog1.SelectedPath
        If newFolder.Length > 0 Then
            If Not newFolder.EndsWith("\") Then
                newFolder = newFolder & "\"
                txtWB_PARM_WEB_IMAGES.Text = newFolder
            End If
        End If
    End Sub
    Private Sub btnWB_PARM_DISC_IMAGES_Click(sender As System.Object, e As System.EventArgs) Handles btnWB_PARM_DISC_IMAGES.Click
        Dim newFolder As String
        FolderBrowserDialog1.ShowDialog()
        newFolder = FolderBrowserDialog1.SelectedPath
        If newFolder.Length > 0 Then
            If Not newFolder.EndsWith("\") Then
                newFolder = newFolder & "\"
                txtWB_PARM_DISC_IMAGES.Text = newFolder
            End If
        End If
    End Sub


    Private Sub btnWB_PARM_FINAL_IMAGES_Click(sender As System.Object, e As System.EventArgs) Handles btnWB_PARM_FINAL_IMAGES.Click
        Dim newFolder As String
        FolderBrowserDialog1.ShowDialog()
        newFolder = FolderBrowserDialog1.SelectedPath
        If newFolder.Length > 0 Then
            If Not newFolder.EndsWith("\") Then
                newFolder = newFolder & "\"
                txtWB_PARM_FINAL_IMAGES.Text = newFolder
            End If
        End If
    End Sub

    Private Sub txtWB_PARM_MASTER_IMAGES_LostFocus(sender As Object, e As System.EventArgs) Handles txtWB_PARM_MASTER_IMAGES.LostFocus
        If Not txtWB_PARM_MASTER_IMAGES.Text.EndsWith("\") Then
            txtWB_PARM_MASTER_IMAGES.Text = txtWB_PARM_MASTER_IMAGES.Text & "\"
        End If
    End Sub

    Private Sub txtWB_PARM_WEB_IMAGES_LostFocus(sender As Object, e As System.EventArgs) Handles txtWB_PARM_WEB_IMAGES.LostFocus
        If Not txtWB_PARM_WEB_IMAGES.Text.EndsWith("\") Then
            txtWB_PARM_WEB_IMAGES.Text = txtWB_PARM_WEB_IMAGES.Text & "\"
        End If
    End Sub

    Private Sub txtWB_PARM_DISC_IMAGES_LostFocus(sender As Object, e As System.EventArgs) Handles txtWB_PARM_DISC_IMAGES.LostFocus
        If Not txtWB_PARM_DISC_IMAGES.Text.EndsWith("\") Then
            txtWB_PARM_DISC_IMAGES.Text = txtWB_PARM_DISC_IMAGES.Text & "\"
        End If
    End Sub

    Private Sub txtWB_PARM_FINAL_IMAGES_LostFocus(sender As Object, e As System.EventArgs) Handles txtWB_PARM_FINAL_IMAGES.LostFocus
        If Not txtWB_PARM_FINAL_IMAGES.Text.EndsWith("\") Then
            txtWB_PARM_FINAL_IMAGES.Text = txtWB_PARM_FINAL_IMAGES.Text & "\"
        End If
    End Sub

    Private Sub chkHideFinished_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkHideFinished.CheckedChanged
        FilterMaster()
    End Sub

    Private Sub btnAUTOCOMPLETE_Click(sender As System.Object, e As System.EventArgs) Handles btnAUTOCOMPLETE.Click
        For Each grow As UltraWinGrid.UltraGridRow In grdICTWPICS.Rows
            grdICTWPICS.ActiveRow = grow
            Dim AutoCheck As Boolean = False
            txtNotes.Text = ""
            If Not picMaster.ImageLocation.EndsWith("MULTI") And Not picWeb.ImageLocation.EndsWith("MULTI") Then
                If picMaster.ImageLocation = "" And picWeb.ImageLocation <> "" Then
                    AutoCheck = True
                    chkImageMaster.Checked = False
                    chkImageWeb.Checked = True
                    txtNotes.Text = "Auto-Set From Web Image."
                End If
                If picWeb.ImageLocation = "" And picMaster.ImageLocation <> "" Then
                    AutoCheck = True
                    chkImageMaster.Checked = True
                    chkImageWeb.Checked = False
                    txtNotes.Text = "Auto-Set From Master Image."
                End If
                If AutoCheck Then
                    MakeFinal(grdICTWPICS.ActiveRow)
                End If
                If picWeb.ImageLocation = "" And picMaster.ImageLocation = "" Then
                    grow.Cells.Item("USER_NOTES").Value = "Auto-Set As No Image Found."
                    grow.Cells.Item("IMG_NOT_FOUND").Value = "1"
                    grow.Cells.Item("PROCESSED").Value = "1"
                Else
                    grow.Cells.Item("IMG_NOT_FOUND").Value = "0"
                End If
            End If
        Next
        'Update_Record()
    End Sub
#End Region

#Region "grdICTWPICS"

    Private Sub grdICTWPICS_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTWPICS.AfterRowActivate
        Setup_IMAGES()
    End Sub

#End Region

#Region "grdICTWPICM"
    Private Sub grdICTWPICM_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTWPICM.AfterRowActivate
        Dim MASTER_IMAGE_NAME As String = grdICTWPICM.ActiveRow.Cells("MASTER_IMAGE_NAME").Value & ""
        If MASTER_IMAGE_NAME.Length > 0 Then
            picMaster.ImageLocation = txtWB_PARM_MASTER_IMAGES.Text & MASTER_IMAGE_NAME
        Else
            MASTER_IMAGE_NAME = grdICTWPICS.ActiveRow.Cells("MASTER_IMAGE_NAME").Value & ""
            If MASTER_IMAGE_NAME.Length > 0 Then
                picMaster.ImageLocation = txtWB_PARM_MASTER_IMAGES.Text & MASTER_IMAGE_NAME
            Else
                picMaster.ImageLocation = ""
            End If
        End If
    End Sub
#End Region

#Region "grdICTWPICW"
    Private Sub grdICTWPICW_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTWPICW.AfterRowActivate
        Dim WEB_IMAGE_NAME As String = grdICTWPICW.ActiveRow.Cells("WEB_IMAGE_NAME").Value & ""
        If WEB_IMAGE_NAME.Length > 0 Then
            picWeb.ImageLocation = txtWB_PARM_WEB_IMAGES.Text & WEB_IMAGE_NAME
        Else
            WEB_IMAGE_NAME = grdICTWPICS.ActiveRow.Cells("WEB_IMAGE_NAME").Value & ""
            If WEB_IMAGE_NAME.Length > 0 Then
                picWeb.ImageLocation = txtWB_PARM_WEB_IMAGES.Text & WEB_IMAGE_NAME
            Else
                picWeb.ImageLocation = ""
            End If
        End If
    End Sub
#End Region

#Region "PictureBoxes"
    Private Sub picMaster_DoubleClick(sender As Object, e As System.EventArgs) Handles picMaster.DoubleClick
        Dim frmSOFIMGV1 As New WBFIMGV1(Me, picMaster.ImageLocation)
        frmSOFIMGV1.Show()
    End Sub

    Private Sub picWeb_DoubleClick(sender As Object, e As System.EventArgs) Handles picWeb.DoubleClick
        Dim frmSOFIMGV1 As New WBFIMGV1(Me, picWeb.ImageLocation)
        frmSOFIMGV1.Show()
    End Sub

    Private Sub PictureBox3_DoubleClick(sender As Object, e As System.EventArgs) Handles PictureBox3.DoubleClick
        Dim frmSOFIMGV1 As New WBFIMGV1(Me, PictureBox3.ImageLocation)
        frmSOFIMGV1.Show()
    End Sub
#End Region

    Private Sub chkNOIMG_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkNOIMG.CheckedChanged
        FilterMaster()
    End Sub

    Private Sub txtFactory_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtFactory.TextChanged
        Dim dvw As DataView = DirectCast(grdICTWPICS.DataSource, DataTable).DefaultView
        If txtFactory.Text.Length = 0 Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = String.Format("FACTORY = '{0}'", txtFactory.Text)
        End If
    End Sub

    'Private Sub BuildPotentials()
    '    MsgBox("I Don't think this routine is applicable anymore", MsgBoxStyle.Critical, "No Mass")
    '    Exit Sub
    '    Dim RecCount As Integer = 0
    '    If Not dst.Tables.Contains("ICTSTYLP") Then
    '        Dim sql As New Text.StringBuilder
    '        sql.Length = 0
    '        sql.AppendLine("SELECT")
    '        sql.AppendLine("S1.STYLE_CODE,")
    '        sql.AppendLine("C1.COLOR_CODE,")
    '        sql.AppendLine("S1.STYLE_DESC,")
    '        sql.AppendLine("S1.style_status,")
    '        sql.AppendLine("C1.STYLE_COLOR_STATUS,")
    '        sql.AppendLine("0 AS MSOH,")
    '        sql.AppendLine("0 AS MSFT,")
    '        sql.AppendLine("0 AS SWOH,")
    '        sql.AppendLine("0 AS SWFT,")
    '        sql.AppendLine("(S1.STYLE_CODE || '-' || C1.COLOR_CODE || '.JPG') AS IMG_NAME,")
    '        sql.AppendLine("'0' AS IMG_FOUND,")
    '        sql.AppendLine("'0' AS USE_ON_WEB")
    '        sql.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1")
    '        sql.AppendLine("WHERE S1.STYLE_CODE = c1.style_code")
    '        ASCMAIN1.sql = sql.ToString()
    '        Create_TDA(dst.Tables.Add, "ICTSTYX1", "**", 0, False, "", 2)
    '    Else
    '        dst.Tables("ICTSTYX1").Clear()
    '    End If

    '    If Not dst.Tables.Contains("ICTSTYC1") Then
    '        Dim sqls As New Text.StringBuilder
    '        sqls.Length = 0
    '        sqls.AppendLine("SELECT * FROM")
    '        sqls.AppendLine("  (")
    '        sqls.AppendLine("   SELECT C1.STYLE_CODE, C1.COLOR_CODE,")
    '        sqls.AppendLine("   9999 AS ORDR_QTY,")
    '        sqls.AppendLine("   C2.COLOR_DESC AS COLOR_CODE_LONG,")
    '        sqls.AppendLine("   C1.STYLE_COLOR_STATUS,")
    '        sqls.AppendLine("   CASE WHEN")
    '        sqls.AppendLine("   SUM(")
    '        sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
    '        sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
    '        sqls.AppendLine("     ELSE 0")
    '        sqls.AppendLine("     END) < 0")
    '        sqls.AppendLine("   THEN")
    '        sqls.AppendLine("     0")
    '        sqls.AppendLine("   ELSE")
    '        sqls.AppendLine("   SUM(")
    '        sqls.AppendLine("     CASE S2.WHSE_CODE")
    '        sqls.AppendLine("     WHEN 'MS'")
    '        sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
    '        sqls.AppendLine("     ELSE 0")
    '        sqls.AppendLine("     END)")
    '        sqls.AppendLine("   END AS MSOH,")
    '        sqls.AppendLine("   CASE WHEN")
    '        sqls.AppendLine("   SUM(")
    '        sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
    '        sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
    '        sqls.AppendLine("     ELSE 0")
    '        sqls.AppendLine("     END) <= 0")
    '        sqls.AppendLine("   THEN")
    '        sqls.AppendLine("     0")
    '        sqls.AppendLine("   ELSE")
    '        sqls.AppendLine("     CASE WHEN")
    '        sqls.AppendLine("       SUM(")
    '        sqls.AppendLine("       CASE S2.WHSE_CODE")
    '        sqls.AppendLine("       WHEN 'MS'")
    '        sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
    '        sqls.AppendLine("       ELSE 0")
    '        sqls.AppendLine("       END) < 0")
    '        sqls.AppendLine("     THEN")
    '        sqls.AppendLine("       0")
    '        sqls.AppendLine("     ELSE")
    '        sqls.AppendLine("     SUM(")
    '        sqls.AppendLine("       CASE S2.WHSE_CODE")
    '        sqls.AppendLine("       WHEN 'MS'")
    '        sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
    '        sqls.AppendLine("       ELSE 0")
    '        sqls.AppendLine("       END) END")
    '        sqls.AppendLine("   END AS MSFT,")
    '        sqls.AppendLine(" CASE WHEN")
    '        sqls.AppendLine("   SUM(")
    '        sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
    '        sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
    '        sqls.AppendLine("     ELSE 0")
    '        sqls.AppendLine("     END) < 0")
    '        sqls.AppendLine("   THEN")
    '        sqls.AppendLine("     0")
    '        sqls.AppendLine("   ELSE")
    '        sqls.AppendLine("   SUM(")
    '        sqls.AppendLine("     CASE S2.WHSE_CODE")
    '        sqls.AppendLine("     WHEN 'SW'")
    '        sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
    '        sqls.AppendLine("     ELSE 0")
    '        sqls.AppendLine("     END)")
    '        sqls.AppendLine("   END AS SWOH,")
    '        sqls.AppendLine("   CASE WHEN")
    '        sqls.AppendLine("   SUM(")
    '        sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
    '        sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
    '        sqls.AppendLine("     ELSE 0")
    '        sqls.AppendLine("     END) <= 0")
    '        sqls.AppendLine("   THEN")
    '        sqls.AppendLine("     0")
    '        sqls.AppendLine("   ELSE")
    '        sqls.AppendLine("     CASE WHEN")
    '        sqls.AppendLine("       SUM(")
    '        sqls.AppendLine("       CASE S2.WHSE_CODE")
    '        sqls.AppendLine("       WHEN 'SW'")
    '        sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
    '        sqls.AppendLine("       ELSE 0")
    '        sqls.AppendLine("       END) < 0")
    '        sqls.AppendLine("     THEN")
    '        sqls.AppendLine("       0")
    '        sqls.AppendLine("     ELSE")
    '        sqls.AppendLine("     SUM(")
    '        sqls.AppendLine("       CASE S2.WHSE_CODE")
    '        sqls.AppendLine("       WHEN 'SW'")
    '        sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
    '        sqls.AppendLine("       ELSE 0")
    '        sqls.AppendLine("       END) END")
    '        sqls.AppendLine("   END AS SWFT")
    '        sqls.AppendLine("   FROM ICTSTYC1 C1")
    '        sqls.AppendLine("   LEFT JOIN ICTSTAT2 S2")
    '        sqls.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
    '        sqls.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
    '        sqls.AppendLine("   INNER JOIN ICTCOLR1 C2")
    '        sqls.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
    '        sqls.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
    '        sqls.AppendLine("  )")
    '        sqls.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') or (MSOH <> 0) or (MSFT <> 0) or (SWOH <> 0)  or (SWFT <> 0))")
    '        sqls.AppendLine("  AND STYLE_CODE = :PARM1")
    '        sqls.AppendLine("  AND COLOR_CODE = :PARM2")
    '        ASCMAIN1.sql = sqls.ToString
    '        Create_TDA(dst.Tables.Add, "ICTSTYC1", "**", 0, False, "VV", 2)
    '    Else
    '        dst.Tables("ICTSTYC1").Clear()
    '    End If
    '    Fill_Records("ICTSTYX1")

    '    For Each rowICTSTYX1 As DataRow In dst.Tables("ICTSTYX1").Select()
    '        RecCount = RecCount + 1
    '        ASCMAIN1.Progress("Record: " & RecCount)

    '        Dim TOTALOHFT As Integer = 0
    '        Dim STYLE_CODE As String = rowICTSTYX1.Item("STYLE_CODE").ToString & ""
    '        Dim COLOR_CODE As String = rowICTSTYX1.Item("COLOR_CODE").ToString & ""

    '        Fill_Records("ICTSTYC1", New Object() {STYLE_CODE, COLOR_CODE}, True)
    '        If dst.Tables("ICTSTYC1").Rows.Count > 0 Then
    '            If dst.Tables("ICTSTYC1").Rows.Count > 1 Then
    '                Stop
    '            Else
    '                rowICTSTYX1.Item("MSOH") = Val(dst.Tables("ICTSTYC1").Rows(0).Item("MSOH").ToString & "")
    '                rowICTSTYX1.Item("MSFT") = Val(dst.Tables("ICTSTYC1").Rows(0).Item("MSFT").ToString & "")
    '                rowICTSTYX1.Item("SWOH") = Val(dst.Tables("ICTSTYC1").Rows(0).Item("SWOH").ToString & "")
    '                rowICTSTYX1.Item("SWFT") = Val(dst.Tables("ICTSTYC1").Rows(0).Item("SWFT").ToString & "")

    '                TOTALOHFT = TOTALOHFT + Val(dst.Tables("ICTSTYC1").Rows(0).Item("MSOH").ToString & "")
    '                TOTALOHFT = TOTALOHFT + Val(dst.Tables("ICTSTYC1").Rows(0).Item("MSFT").ToString & "")
    '                TOTALOHFT = TOTALOHFT + Val(dst.Tables("ICTSTYC1").Rows(0).Item("SWOH").ToString & "")
    '                TOTALOHFT = TOTALOHFT + Val(dst.Tables("ICTSTYC1").Rows(0).Item("SWFT").ToString & "")
    '            End If
    '        End If

    '        Dim File_Location As String = "\\192.168.110.224\c$\RGI_PORTAL\images\product\" & rowICTSTYX1.Item("IMG_NAME").ToString & ""
    '        If IO.File.Exists(File_Location) Then
    '            rowICTSTYX1.Item("IMG_FOUND") = "1"
    '            If rowICTSTYX1.Item("IMG_FOUND").ToString & "" = "1" Then
    '                If TOTALOHFT > 0 Or rowICTSTYX1.Item("STYLE_COLOR_STATUS").ToString & "" = "A" Then
    '                    rowICTSTYX1.Item("USE_ON_WEB") = "1"
    '                End If
    '            End If
    '        Else
    '            rowICTSTYX1.Item("IMG_FOUND") = "0"
    '        End If
    '    Next
    '    MsgBox("Table Created")
    'End Sub

    Private Sub SendToWeb()
        If Not dst.Tables.Contains("ICTSTYC1") Then
            Dim sqls As New Text.StringBuilder
            sqls.Length = 0
            sqls.AppendLine("SELECT * FROM")
            sqls.AppendLine("  (")
            sqls.AppendLine("   SELECT C1.STYLE_CODE, C1.COLOR_CODE,")
            sqls.AppendLine("   9999 AS ORDR_QTY,")
            sqls.AppendLine("   C2.COLOR_DESC AS COLOR_CODE_LONG,")
            sqls.AppendLine("   C1.STYLE_COLOR_STATUS,")
            sqls.AppendLine("   CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) < 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE")
            sqls.AppendLine("     WHEN 'MS'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END)")
            sqls.AppendLine("   END AS MSOH,")
            sqls.AppendLine("   CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) <= 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("     CASE WHEN")
            sqls.AppendLine("       SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'MS'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) < 0")
            sqls.AppendLine("     THEN")
            sqls.AppendLine("       0")
            sqls.AppendLine("     ELSE")
            sqls.AppendLine("     SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'MS'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) END")
            sqls.AppendLine("   END AS MSFT,")
            sqls.AppendLine(" CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) < 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE")
            sqls.AppendLine("     WHEN 'SW'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END)")
            sqls.AppendLine("   END AS SWOH,")
            sqls.AppendLine("   CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) <= 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("     CASE WHEN")
            sqls.AppendLine("       SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'SW'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) < 0")
            sqls.AppendLine("     THEN")
            sqls.AppendLine("       0")
            sqls.AppendLine("     ELSE")
            sqls.AppendLine("     SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'SW'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) END")
            sqls.AppendLine("   END AS SWFT")
            sqls.AppendLine("   FROM ICTSTYC1 C1")
            sqls.AppendLine("   LEFT JOIN ICTSTAT2 S2")
            sqls.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
            sqls.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
            sqls.AppendLine("   INNER JOIN ICTCOLR1 C2")
            sqls.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
            sqls.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
            sqls.AppendLine("  )")
            'sqls.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') or (MSOH <> 0) or (MSFT <> 0) or (SWOH <> 0)  or (SWFT <> 0))")
            sqls.AppendLine("  WHERE STYLE_CODE = :PARM1")
            sqls.AppendLine("  AND COLOR_CODE = :PARM2")
            ASCMAIN1.sql = sqls.ToString
            Create_TDA(dst.Tables.Add, "ICTSTYC1", "**", 0, False, "VV", 2)
        Else
            dst.Tables("ICTSTYC1").Clear()
        End If
        Fill_Records("WBTSTYL1")
        Fill_Records("WBTSTYL2")
        Dim lastSTYLE_CODE As String = ""
        For Each rowICTWPICS As DataRow In dst.Tables("ICTWPICS").Select("WEB = '1'", "STYLE_CODE, COLOR_CODE")
            If rowICTWPICS.Item("WEB").ToString & "" = "1" Then
                Dim Filter1 As String = String.Format("STYLE_CODE = '{0}'", rowICTWPICS.Item("STYLE_CODE"))
                Dim Filter2 As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", rowICTWPICS.Item("STYLE_CODE"), rowICTWPICS.Item("COLOR_CODE"))
                If dst.Tables.Item("WBTSTYL2").Select(Filter2).Count = 0 Then
                    Fill_Records("ICTSTYC1", New Object() {rowICTWPICS.Item("STYLE_CODE"), rowICTWPICS.Item("COLOR_CODE")}, True)
                    If dst.Tables("ICTSTYC1").Rows.Count <> 1 Then
                        MsgBox(String.Format("Style: {0}-{1}", rowICTWPICS.Item("STYLE_CODE"), rowICTWPICS.Item("COLOR_CODE")), MsgBoxStyle.Critical, "Inventory Problem With")
                    End If
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowICTWPICS.Item("STYLE_CODE"))
                    If rowICTWPICS.Item("STYLE_CODE") <> lastSTYLE_CODE Then
                        lastSTYLE_CODE = rowICTWPICS.Item("STYLE_CODE")
                        If dst.Tables.Item("WBTSTYL1").Select(Filter1).Count = 0 Then
                            Dim newWBTSTYL1 As DataRow = dst.Tables.Item("WBTSTYL1").NewRow
                            newWBTSTYL1.Item("STYLE_CODE") = rowICTWPICS.Item("STYLE_CODE").ToString
                            newWBTSTYL1.Item("STYLE_STATUS") = rowICTSTYL1.Item("STYLE_STATUS").ToString
                            newWBTSTYL1.Item("STYLE_FULL_DESC") = rowICTSTYL1.Item("STYLE_DESC").ToString
                            newWBTSTYL1.Item("WEB_IND") = "1"
                            newWBTSTYL1.Item("UPLOAD_BATCH") = ""
                            newWBTSTYL1.Item("DEFAULT_IMAGE") = rowICTWPICS.Item("STYLE_COLOR_IMAGE_NAME").ToString
                            newWBTSTYL1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            newWBTSTYL1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            newWBTSTYL1.Item("INIT_DATE") = DATETIME_STAMP
                            newWBTSTYL1.Item("LAST_DATE") = DATETIME_STAMP
                            dst.Tables.Item("WBTSTYL1").Rows.Add(newWBTSTYL1)
                        End If
                    End If
                    Dim newWBTSTYL2 As DataRow = dst.Tables.Item("WBTSTYL2").NewRow
                    newWBTSTYL2.Item("STYLE_CODE") = rowICTWPICS.Item("STYLE_CODE").ToString
                    newWBTSTYL2.Item("COLOR_CODE") = rowICTWPICS.Item("COLOR_CODE").ToString
                    newWBTSTYL2.Item("COLOR_CODE_LONG") = dst.Tables("ICTSTYC1").Rows(0).Item("COLOR_CODE_LONG").ToString
                    newWBTSTYL2.Item("COLOR_STATUS") = dst.Tables("ICTSTYC1").Rows(0).Item("STYLE_COLOR_STATUS").ToString
                    newWBTSTYL2.Item("MSOH") = dst.Tables("ICTSTYC1").Rows(0).Item("MSOH").ToString
                    newWBTSTYL2.Item("MSFT") = dst.Tables("ICTSTYC1").Rows(0).Item("MSFT").ToString
                    newWBTSTYL2.Item("SWOH") = dst.Tables("ICTSTYC1").Rows(0).Item("SWOH").ToString
                    newWBTSTYL2.Item("SWFT") = dst.Tables("ICTSTYC1").Rows(0).Item("SWFT").ToString
                    newWBTSTYL2.Item("IMG_NAME") = rowICTWPICS.Item("STYLE_COLOR_IMAGE_NAME").ToString
                    newWBTSTYL2.Item("IMG_FOUND") = "1"
                    dst.Tables.Item("WBTSTYL2").Rows.Add(newWBTSTYL2)

                End If

            End If
        Next
        Dim msg As String = "Web Update Complete"
        msg = msg & vbCrLf & "Make Sure You Update And"
        msg = msg & vbCrLf & "Move All The Finished Images"
        msg = msg & vbCrLf & "To The Live Image Folder."
        MsgBox(msg, MsgBoxStyle.Exclamation, "Finished")
    End Sub

    Private Sub WBFIMGC1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F2 Then
            MakeFinal(grdICTWPICS.ActiveRow)
        End If
    End Sub

    Private Function Remove_Dups() As Integer
        Dim RetVal As Integer = 0
        For Each grow As UltraWinGrid.UltraGridRow In grdICTWPICS.Rows
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("Select Count(*) as RecCnt from WBTSTYL1 where STYLE_CODE = '{0}'", grow.Cells.Item("STYLE_CODE").Text))
            ASCMAIN1.sql = SQLS.ToString()
            Dim RecCnt As Int16 = Val(ASCDATA1.GetDataValue)
            If RecCnt > 0 Then
                grow.Cells.Item("WEB").Value = "0"
                RetVal += 1
            End If
        Next
        Return RetVal
    End Function

End Class