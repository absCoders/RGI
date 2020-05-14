Imports System.Drawing
Imports System.Math

Public Class WHFTRAN1


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select WHTMOVE2.*, WHSE_TRAN_TYPE from WHTMOVE1, WHTMOVE2"
            Create_TDA(.Tables.Add, "WHTMOVE2", "**", 0, False, "", 2)
            .Tables("WHTMOVE2").Columns.Add("TRAN_TYPE_COL")

            Create_TDA(.Tables.Add, "ASTUSER1", "*", , False, "", 1)

        End With
        grdASTUSER1.DataSource = dst.Tables("ASTUSER1")
        grdWHTMOVEX.DataSource = dst.Tables("WHTMOVE2")

        ASCMAIN1.Add_Value_List(grdWHTMOVEX, "TRAN_TYPE_COL", , New String() {":", "D:Deposit", "F:Finalized", "M:Move", "X:Cycle", "U:Wave Unit", "L:Wave Load", "C:Wave Case"})


        'With grdWHTWRTN2.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME In New String() {"CASES", "UNITS", "WH_RTN_NO", "WH_RTN_LNO", "STYLE_DESC", "COLOR_DESC"}
        '        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
        '    Next
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        If New String() {"STYLE_CODE", "COLOR_CODE", "CTN_PACK_QTY"}.Contains(gcol.Key) Then
        '            gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
        '            gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        '        ElseIf New String() {"CASES", "UNITS"}.Contains(gcol.Key) Then
        '            gcol.CellAppearance.BackColor = Drawing.Color.Beige
        '        End If
        '        gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
        '    Next
        'End With

        '        Create_Summary(grdASTUSER1, New String() {"CASES", "UNITS"})
        '        ASCMAIN1.Add_Value_List(grdWHTRTRNX, "WH_RTN_STATUS", , New String() {":", "S:SAVED", "C:COMPLETED", "F:FINALIZED"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "New"
               
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Select Users"

                ASCMAIN1.CodeSelector.Get_SQL("USER_ID")
                ASCMAIN1.CodeSelector.SQL = "Select USER_ID, USER_NAME, USER_COMPANY, USER_EMAIL from ASTUSER1 Where USER_ID in (" _
                & " Select Distinct USER_ID from ASTUSER2 where security_code = 'NJE') and USER_STATUS = 'A'"

                ASCMAIN1.CodeSelector.MultipleSelections = True
                Using F As New ASFCODE1
                    F.ShowDialog()
                End Using
                If ASCMAIN1.CodeSelector.Selections <> 0 Then

                    For Each AC As String In ASCMAIN1.CodeSelector.SelectedCodes
                        Dim row As DataRow = clsASCBASE1.LookUp("ASTUSER1", AC)

                        Dim rowASTUSER1 As DataRow = dst.Tables("ASTUSER1").NewRow
                        With rowASTUSER1
                            .Item("USER_ID") = row.Item("USER_ID")
                            .Item("USER_NAME") = row.Item("USER_NAME")
                        End With
                        dst.Tables("ASTUSER1").Rows.Add(rowASTUSER1)
                    Next
                    Load_Record()
                    Mode_Settings(True)
                End If


            Case "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Select Users").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                End With
            End With
        End If
        'Set_Read_Only(UltraGroupBox1, ScreenMode)
        tab0.Visible = Not tf
        chkINIT_DATE.Checked = False
        grpDates.Enabled = Not ScreenMode
        If ScreenMode Then


        Else
            Clear_Record()
        End If


    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ASTUSER1", "WHTMOVE2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        fpDaysBack.Value = 30
        optHistory.Value = "D"
        dteFROM.Value = DateAdd("d", -30, Now)
        dteTO.Value = Now
        'dteFROM.MaxDate = Nothing
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data ...")
        Me.Cursor = Cursors.WaitCursor
        'Application.DoEvents()

        Dim Users As String = ""
        For Each rowASTUSER1 As DataRow In dst.Tables("ASTUSER1").Select
            Users &= ",'" & rowASTUSER1.Item("USER_ID") & "'"
        Next
        If Users <> "" Then
            Users = Mid(Users, 2)

            ASCMAIN1.sql = "Select M2.WHSE_TRAN_NO, M2.WHSE_TRAN_LNO, M2.LOCATION_CODE_FROM, " _
            & " M2.LOCATION_CODE_TO,M2.BAR_CODE, M2.WHSE_TRAN_QTY, " _
            & " M2.STYLE_CODE, M2.COLOR_CODE, M2.INIT_OPER," _
            & " M2.INIT_DATE, M2.LAST_OPER, M2.LAST_DATE, M2.STATUS, " _
            & " M2.LOAD_NO_FROM, M2.LOAD_NO_TO, M2.BAR_CODE_OTHER, WHSE_TRAN_TYPE, WHSE_TRAN_TYPE TRAN_TYPE_COL  from WHTMOVE1, WHTMOVE2 M2" _
            & " Where WHTMOVE1.WHSE_TRAN_NO = M2.WHSE_TRAN_NO" _
            & " And WHTMOVE1.INIT_OPER in (" & Users & ")"
            If optHistory.Value = "R" Then
                ASCMAIN1.sql &= " And WHTMOVE1.INIT_DATE >= '" & Format(dteFROM.Value, "dd-MMM-yy") & "' and WHTMOVE1.INIT_DATE <= '" & Format(dteTO.Value, "dd-MMM-yy") & "' "
            Else
                ASCMAIN1.sql &= " And WHTMOVE1.INIT_DATE >= '" & Format(DateAdd("d", fpDaysBack.Value * -1, Now), "dd-MMM-yy") & "' "
            End If


            Dim sql_WHTCYCLEX As String = " Select L1.CYCLE_NO, Rownum, LOCATION_CODE_ORIG, LOCATION_CODE, " _
            & " L2.BAR_CODE, To_Number(CYCLE_SCAN), '' STYLE_CODE, '' COLOR_CODE," _
            & " INIT_OPER, INIT_DATE, LAST_OPER, LAST_DATE, CYCLE_STATUS , " _
            & " '' as LOAD_NO_FROM,'' as LOAD_NO_TO, '' as BAR_CODE_OTHER , 'X' as WHSE_TRAN_TYPE,  'X' TRAN_TYPE_COL" _
            & " from WHTCYCL1 L1, WHTCYCL2 L2" _
            & " Where L1.CYCLE_NO = L2.CYCLE_NO" _
            & " And INIT_OPER in (" & Users & ")"
            If optHistory.Value = "R" Then
                sql_WHTCYCLEX &= " And INIT_DATE >= '" & Format(dteFROM.Value, "dd-MMM-yy") & "' and INIT_DATE <= '" & Format(dteTO.Value, "dd-MMM-yy") & "' "
            Else
                sql_WHTCYCLEX &= " And INIT_DATE >= '" & Format(DateAdd("d", fpDaysBack.Value * -1, Now), "dd-MMM-yy") & "' "
            End If

            Dim WHTCYCLEX As String = ASCMAIN1.Temp_Table("Select distinct CYCLE_NO from (" & sql_WHTCYCLEX & ")")
            ASCMAIN1.sql &= " Union " & sql_WHTCYCLEX

            ASCMAIN1.sql &= " Union " _
            & " Select CYCLE_NO, Rownum, LOCATION_CODE LOCATION_CODE_ORIG, LOCATION_CODE, " _
            & " '' BAR_CODE, To_Number(CASES_BOOK), '' STYLE_CODE, '' COLOR_CODE," _
            & " INIT_OPER, INIT_DATE, LAST_OPER, LAST_DATE, CYCLE_STATUS , " _
            & " '' as LOAD_NO_FROM,'' as LOAD_NO_TO, '' as BAR_CODE_OTHER , 'X' as WHSE_TRAN_TYPE,  'X' TRAN_TYPE_COL" _
            & " from WHTCYCL1" _
            & " Where INIT_OPER in (" & Users & ")" _
            & " And CYCLE_NO not in (Select CYCLE_NO from " & WHTCYCLEX & ")"
            If optHistory.Value = "R" Then
                ASCMAIN1.sql &= " And INIT_DATE >= '" & Format(dteFROM.Value, "dd-MMM-yy") & "' and INIT_DATE <= '" & Format(dteTO.Value, "dd-MMM-yy") & "' "
            Else
                ASCMAIN1.sql &= " And INIT_DATE >= '" & Format(DateAdd("d", fpDaysBack.Value * -1, Now), "dd-MMM-yy") & "' "
            End If

            ASCMAIN1.sql &= " Union " & vbCrLf _
            & " Select L1.WAVE_INST_NO, " & vbCrLf _
            & " Rownum, LOCATION_CODE, LOCATION_CODE_OTHER,  L2.BAR_CODE, To_Number(LOCATION_QTY_PICK), " & vbCrLf _
            & " STYLE_CODE, COLOR_CODE, INIT_OPER, INIT_DATE, " & vbCrLf _
            & " LAST_OPER, LAST_DATE, WAVE_PICK_TYPE ,  LOAD_NO, LOAD_NO_OTHER, '' as BAR_CODE_OTHER , WAVE_PICK_TYPE as WHSE_TRAN_TYPE, WAVE_PICK_TYPE TRAN_TYPE_COL " & vbCrLf _
            & " from WHTINST1 L1, WHTINST2 L2 Where L1.WAVE_INST_NO = L2.WAVE_INST_NO " & vbCrLf _
            & " And INIT_OPER in (" & Users & ")"
            If optHistory.Value = "R" Then
                ASCMAIN1.sql &= " And INIT_DATE >= '" & Format(dteFROM.Value, "dd-MMM-yy") & "' and INIT_DATE <= '" & Format(dteTO.Value, "dd-MMM-yy") & "' "
            Else
                ASCMAIN1.sql &= " And INIT_DATE >= '" & Format(DateAdd("d", fpDaysBack.Value * -1, Now), "dd-MMM-yy") & "' "
            End If
            Fill_Records("WHTMOVE2", , , ASCMAIN1.sql)
        End If
        ' Sort_grdColumns(grdASTUSER1, "USER_NAME")

        Sort_grdColumns(grdWHTMOVEX, "INIT_DATE".ToLower)


        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub


    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where &= " AND WHSE_CTN_CTL = 'C'"
        End Select
    End Sub

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        'If ScreenMode Then
        '    E.TABLE_NAME = "SOTRTRN1"
        '    E.COLUMN_NAME = "RTRN_NO"
        '    E.CODE_VALUE = Absx1.txtFor("RTRN_NO").Text
        '    E.DESC_VALUE = "Return"
        '    E.ATTACHMENT_NOTES = ""
        'End If

        Return E
    End Function


#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTMOVEX, "SSBBBB", "Show Filter", "Show GroupBox", "Location Inquiry using Location From", "Location Inquiry using Location To", "Style Status Inquiry", "Location Inquiry using Style")
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

        Select Case grd.Name
            Case "grdSOTRTRN2"


        End Select

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
            Case ""

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Location Inquiry using Location From"
                Context_Launch("Select", "L:" & grd.ActiveRow.Cells("LOCATION_CODE_FROM").Value, e.Tool.Key, "WHFLOCS1")
            Case "Location Inquiry using Location To"
                Context_Launch("Select", "L:" & grd.ActiveRow.Cells("LOCATION_CODE_TO").Value, e.Tool.Key, "WHFLOCS1")
            Case "Location Inquiry using Style"
                Context_Launch("Select", "S:" & grd.ActiveRow.Cells("STYLE_CODE").Value, e.Tool.Key, "WHFLOCS1")
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



#End Region

    Private Sub optHistory_ValueChanged(sender As Object, e As EventArgs) Handles optHistory.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        splDays.Panel1Collapsed = IIf(optHistory.Value = "R", False, True)
        splDays.Panel2Collapsed = IIf(optHistory.Value = "R", True, False)
    End Sub

    Private Sub grdASTUSER1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdASTUSER1.AfterRowActivate
        If SELECTION_NO = 0 Then Exit Sub
        Dim sqlUser As String = ""
        If grdASTUSER1.ActiveRow.IsAddRow Then
            Exit Sub
        Else
            grdWHTMOVEX.Text = "Transaction Details"
            If grdASTUSER1.ActiveRow IsNot Nothing Then
                sqlUser = " INIT_OPER = '" & grdASTUSER1.ActiveRow.Cells("USER_ID").Value & "'"
                grdWHTMOVEX.Text = "Transaction Details for " & grdASTUSER1.ActiveRow.Cells("USER_NAME").Value
            End If
        End If
        ASCMAIN1.Progress("Now Loading Data ...")
        Me.Cursor = Cursors.WaitCursor
        Dim dvw As DataView = DirectCast(grdWHTMOVEX.DataSource, DataTable).DefaultView
        dvw.RowFilter = sqlUser
        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdASTUSER1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTUSER1.InitializeLayout

    End Sub


    Private Sub dteTO_ValueChanged(sender As Object, e As EventArgs) Handles dteTO.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If dteTO.Value <> Nothing Then
            ' dteFROM.MaxDate = DateValue(dteTO.Value)
        Else
            dteFROM.MaxDate = Nothing
        End If

    End Sub

    Private Sub chkINIT_DATE_CheckedChanged(sender As Object, e As EventArgs) Handles chkINIT_DATE.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        grdWHTMOVEX.DisplayLayout.Bands(0).Columns("INIT_DATE").Hidden = IIf(chkINIT_DATE.Checked, True, False)
    End Sub
End Class