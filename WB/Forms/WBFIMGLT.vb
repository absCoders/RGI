
Imports System.Text
Imports System.IO

Public Class WBFIMGLT
    Dim S As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
    Dim TTM As New UltraWinToolTip.UltraToolTipManager
    Dim IMAGES_FOLDER_HIGH As String = ""
    Dim IMAGES_FOLDER_LOW As String = ""
    Dim IMAGE_DEFAULT As String = ""


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Get_PARM("ICTPARMI")
        IMAGES_FOLDER_HIGH = ROWs("ICTPARMI").Item("IMAGES_FOLDER_HIGH") & String.Empty
        IMAGES_FOLDER_LOW = ROWs("ICTPARMI").Item("IMAGES_FOLDER_LOW") & String.Empty
        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            'IMAGES_FOLDER_HIGH = ""
            'IMAGES_FOLDER_LOW = ""
        End If

        With dst

            S.Length = 0
            S.AppendLine("SELECT * FROM ICTIMAGT")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTIMAGT", "**", 0, False)
            Fill_Records("ICTIMAGT")

            S.Length = 0
            S.AppendLine("SELECT * FROM WBTSTYLD")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "WBTSTYLD", "**", 0, False)
            Fill_Records("WBTSTYLD")

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("DISTINCT STYLE_CODE")
            S.AppendLine("FROM ECTESTY1")
            S.AppendLine("WHERE (NVL(SHIP_ECOM,'0') = '1' OR NVL(SHIP_DROP,'0') = '1')")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ECTESTY1", "**", 0, False)
            Fill_Records("ECTESTY1")

            S.Length = 0
            S.AppendLine("SELECT * FROM")
            S.AppendLine("(")
            S.AppendLine("SELECT")
            S.AppendLine("S1.STYLE_CODE,")
            S.AppendLine("C1.COLOR_CODE,")
            S.AppendLine("S1.STYLE_STATUS,")
            S.AppendLine("C1.STYLE_COLOR_STATUS,")
            S.AppendLine("S1.STYLE_DESC,")
            S.AppendLine("SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))) AS AVAIL")
            S.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1, ICTSTAT2 S2")
            S.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE")
            S.AppendLine("AND C1.STYLE_CODE = S2.STYLE_CODE (+)")
            S.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
            S.AppendLine("AND S2.WHSE_CODE = 'MS'")
            S.AppendLine("GROUP BY")
            S.AppendLine("S1.STYLE_CODE,")
            S.AppendLine("C1.COLOR_CODE,")
            S.AppendLine("S1.STYLE_STATUS,")
            S.AppendLine("C1.STYLE_COLOR_STATUS,")
            S.AppendLine("S1.STYLE_DESC")
            S.AppendLine(")")
            S.AppendLine("WHERE (STYLE_COLOR_STATUS = 'A' OR AVAIL > 0)")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "WBTIMGLT", "**", 0, False)
            With .Tables("WBTIMGLT")
                .Columns.Add("WEB").DefaultValue = 0
                .Columns.Add("ECOM").DefaultValue = 0
                .Columns.Add("LOWREZ").DefaultValue = 0

                IMAGE_DEFAULT = dst.Tables("ICTIMAGT").Select("IMAGE_DEFAULT = '1'").FirstOrDefault.Item("IMAGE_CODE").ToString & String.Empty
                If IMAGE_DEFAULT.Length > 0 Then
                    .Columns.Add(IMAGE_DEFAULT).DefaultValue = 0
                End If

                For Each rowICTIMAGT As DataRow In dst.Tables("ICTIMAGT").Select("IMAGE_DEFAULT = '0'", "IMAGE_CODE")
                    .Columns.Add(rowICTIMAGT.Item("IMAGE_CODE").ToString & String.Empty).DefaultValue = 0
                Next

                .Columns.Add("NOMATCH").DefaultValue = 0
            End With
        End With

        'Fill_Records("ECTSZIO1")

        grdWBTIMGLT.DataSource = dst.Tables("WBTIMGLT")

        Create_Summary(grdWBTIMGLT, "STYLE_CODE", "Count", "", "###,##0")

        ASCMAIN1.Add_Value_List(grdWBTIMGLT, "STYLE_STATUS", , New String() {":", "A:Active", "N:No Re-Order", "D:Discontinued"})
        ASCMAIN1.Add_Value_List(grdWBTIMGLT, "STYLE_COLOR_STATUS", , New String() {":", "A:Active", "N:No Re-Order", "D:Discontinued"})

        Sort_grdColumns(grdWBTIMGLT, "STYLE_CODE, COLOR_CODE", False)

        For Each rowICTIMAGT As DataRow In dst.Tables("ICTIMAGT").Select("IMAGE_DEFAULT = '0'")
            Dim IMAGE_CODE As String = rowICTIMAGT.Item("IMAGE_CODE").ToString & String.Empty
            Dim IMAGE_DESC As String = rowICTIMAGT.Item("IMAGE_DESC").ToString & String.Empty
            'grdWBTIMGLT.DisplayLayout.Bands(0).Columns.Add(IMAGE_CODE, IMAGE_CODE)
            grdWBTIMGLT.DisplayLayout.Bands(0).Columns.Item(IMAGE_CODE).Style = UltraWinGrid.ColumnStyle.CheckBox
            grdWBTIMGLT.DisplayLayout.Bands(0).Columns.Item(IMAGE_CODE).Header.ToolTipText = IMAGE_DESC
            grdWBTIMGLT.DisplayLayout.Bands(0).Columns.Item(IMAGE_CODE).Hidden = False
        Next

        grdWBTIMGLT.DisplayLayout.Bands(0).Columns("AVAIL").Format = "###,##0"

        With grdWBTIMGLT.DisplayLayout.Bands(0).Columns
            .Item(IMAGE_DEFAULT).Style = UltraWinGrid.ColumnStyle.CheckBox
            .Item(IMAGE_DEFAULT).Header.ToolTipText = "Default High Rez Image"
            .Item(IMAGE_DEFAULT).Hidden = False

            .Item("NOMATCH").Style = UltraWinGrid.ColumnStyle.CheckBox
            .Item("NOMATCH").Header.ToolTipText = "No Matches Found"
            .Item("NOMATCH").Hidden = False

            .Item("WEB").Style = UltraWinGrid.ColumnStyle.CheckBox
            .Item("WEB").Header.ToolTipText = "Style / Color On Shopsite"
            .Item("WEB").Header.Appearance.TextHAlign = HAlign.Center
            .Item("WEB").Hidden = False
            .Item("WEB").Editor.DataFilter = New CheckEditorDataFilter
            .Item("WEB").CellClickAction = UltraWinGrid.CellClickAction.CellSelect
            .Item("WEB").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

            .Item("LOWREZ").Style = UltraWinGrid.ColumnStyle.CheckBox
            .Item("LOWREZ").Header.ToolTipText = "Low Resolution Version Available"
            .Item("LOWREZ").Hidden = False

            .Item("ECOM").Style = UltraWinGrid.ColumnStyle.CheckBox
            .Item("ECOM").Header.ToolTipText = "Style / Color On E-Commerce"
            .Item("ECOM").Hidden = False
        End With

        With grdWBTIMGLT.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        For i As Integer = 0 To grdWBTIMGLT.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBTIMGLT.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

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

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

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
        Load_Popup_Menu(grdWBTIMGLT, "SSB", "Show Filter", "Show GroupBox", "View Image")
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

        If grd.Selected.Rows.Count = 0 Then
            MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "View Image"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells.Item("STYLE_CODE").Value
                Dim COLOR_CODE As String = grd.ActiveRow.Cells.Item("COLOR_CODE").Value
                Dim frmIMAGE As New TAC.TAFIMGV1(Me, STYLE_CODE, COLOR_CODE, "M")
                With frmIMAGE
                    .ShowDialog()
                End With
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
        ASCMAIN1.Progress("Refreshing Styles", "")
        Fill_Records("WBTIMGLT")
        For Each rowWBTIMGLT As DataRow In dst.Tables("WBTIMGLT").Select()
            Dim STYLE_CODE As String = rowWBTIMGLT.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowWBTIMGLT.Item("COLOR_CODE").ToString & String.Empty

            For Each COL As String In New String() {"ECOM", "NOMATCH", "LOWREZ"}
                rowWBTIMGLT.Item(COL) = "0"
            Next
            For Each rowICTIMAGT As DataRow In dst.Tables("ICTIMAGT").Select()
                rowWBTIMGLT.Item(rowICTIMAGT.Item("IMAGE_CODE").ToString & String.Empty) = "0"
            Next

            Dim WFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
            If dst.Tables.Item("WBTSTYLD").Select(WFilter).Count > 0 Then
                rowWBTIMGLT.Item("WEB") = "1"
            Else
                rowWBTIMGLT.Item("WEB") = "0"
            End If

            Dim EFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            If dst.Tables.Item("ECTESTY1").Select(EFilter).Count > 0 Then
                rowWBTIMGLT.Item("ECOM") = "1"
            Else
                rowWBTIMGLT.Item("ECOM") = "0"
            End If

        Next
        MakeFileNameData()

        MsgBox("Analysis Complete", vbOKOnly, "Done")
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub MakeFileNameData()
        ASCMAIN1.Progress("Fetching Info From File System", "")
        Dim FILES_HIGH As String() = IO.Directory.GetFiles(IMAGES_FOLDER_HIGH)
        Dim FILES_LOW As String() = IO.Directory.GetFiles(IMAGES_FOLDER_LOW)
        Dim EXT As String = ".JPG"

        ASCMAIN1.Progress("Processing Low Rez Files", "")
        For Each FL As String In FILES_LOW
            FL = FL.ToUpper
            Dim STYLE_CODE As String = ""
            Dim COLOR_CODE As String = ""
            Dim FLAST As String = ""
            If FL.Length > 4 Then
                If FL.EndsWith(EXT) Then
                    Dim endP As Int64 = FL.ToUpper.IndexOf(EXT)
                    Dim begP As Int64 = FL.LastIndexOf("\") + 1
                    Dim FULL_STYLE As String = FL.Substring(begP, endP - begP)
                    If FULL_STYLE.Length > 1 Then
                        If FULL_STYLE.IndexOf("-") > 0 Then
                            STYLE_CODE = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-"))
                        End If
                        FULL_STYLE = FULL_STYLE.Replace(STYLE_CODE + "-", "")
                    End If
                    If FULL_STYLE.Length > 1 Then
                        If FULL_STYLE.IndexOf("-") > 0 Then
                            COLOR_CODE = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-") + 1)
                            FULL_STYLE = FULL_STYLE.Replace(COLOR_CODE + "-", "")
                        Else
                            COLOR_CODE = FULL_STYLE
                            FULL_STYLE = ""
                        End If
                    End If
                    If FULL_STYLE.Length > 1 Then
                        If FULL_STYLE.IndexOf("-") > 0 Then
                            FLAST = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-") + 1)
                        End If
                        FULL_STYLE = FULL_STYLE.Replace(FLAST + "-", "")
                    End If
                    If STYLE_CODE.Length > 0 And COLOR_CODE.Length > 0 Then
                        Dim LFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                        Try
                            Dim rowWBTIMGLT As DataRow = dst.Tables.Item("WBTIMGLT").Select(LFilter).FirstOrDefault
                            If Not IsNothing(rowWBTIMGLT) Then
                                rowWBTIMGLT.Item("LOWREZ") = "1"
                            End If
                        Catch ex As Exception
                            'Skip this shit
                        End Try

                    End If
                End If
            End If
        Next

        ASCMAIN1.Progress("Processing High Rez Files", "")
        For Each FH As String In FILES_HIGH
            FH = FH.ToUpper
            Dim STYLE_CODE As String = ""
            Dim COLOR_CODE As String = ""
            Dim FLAST As String = ""
            If FH.Length > 4 Then
                If FH.EndsWith(EXT) Then
                    Dim endP As Int64 = FH.ToUpper.IndexOf(EXT)
                    Dim begP As Int64 = FH.LastIndexOf("\") + 1
                    Dim FULL_STYLE As String = FH.Substring(begP, endP - begP)
                    If FULL_STYLE.Length > 1 Then
                        If FULL_STYLE.IndexOf("-") > 0 Then
                            STYLE_CODE = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-"))
                        End If
                        FULL_STYLE = FULL_STYLE.Replace(STYLE_CODE + "-", "")
                    End If
                    If FULL_STYLE.Length > 1 Then
                        If FULL_STYLE.IndexOf("-") > 0 Then
                            COLOR_CODE = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-") + 1)
                            FULL_STYLE = FULL_STYLE.Replace(COLOR_CODE + "-", "")
                        Else
                            COLOR_CODE = FULL_STYLE
                            FULL_STYLE = ""
                        End If
                    End If
                    If FULL_STYLE.Length > 1 Then
                        If FULL_STYLE.IndexOf("-") > 0 Then
                            FLAST = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-") + 1)
                        End If
                        FULL_STYLE = FULL_STYLE.Replace(FLAST + "-", "")
                    End If
                    If STYLE_CODE.Length > 0 And COLOR_CODE.Length > 0 Then
                        Dim LFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                        Try
                            Dim rowWBTIMGLT As DataRow = dst.Tables.Item("WBTIMGLT").Select(LFilter).FirstOrDefault
                            If Not IsNothing(rowWBTIMGLT) Then
                                If FLAST.Length = 0 Then
                                    rowWBTIMGLT.Item(IMAGE_DEFAULT) = "1"
                                Else
                                    If rowWBTIMGLT.Table.Columns.Contains(FLAST) Then
                                        rowWBTIMGLT.Item(FLAST) = "1"
                                    Else
                                        rowWBTIMGLT.Item("NOMATCH") = "1"
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            'Skip this shit
                        End Try

                    End If
                End If
            End If
        Next

    End Sub

    Private Sub grdWBTIMGLT_MouseHover(sender As Object, e As EventArgs) Handles grdWBTIMGLT.MouseHover
        If 1 = 1 Then

        End If
    End Sub
#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class