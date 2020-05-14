Imports ABSolution

Public Class SOFCARRA


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Create_TDA(.Tables.Add, "SOTCARRA", "*")
            .Tables("SOTCARRA").Columns.Add("STATE_NAME", GetType(System.String))

            grdSOTCARRA.DataSource = dst.Tables("SOTCARRA")
            ASCMAIN1.Add_Value_List(grdSOTCARRA, "STATE_NAME", "SELECT STATE_CODE, STATE_NAME FROM TATSTATE")

            Create_TDA(.Tables.Add, "TATSTATE", "*")
            Fill_Records("TATSTATE", String.Empty, True, "SELECT STATE_CODE, STATE_NAME FROM TATSTATE")

            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Fill_Records("SOTCARR1", String.Empty, True, "SELECT * FROM SOTCARR1")
            Dim row As DataRow = dst.Tables("SOTCARR1").NewRow
            row.Item("CARRIER_CODE") = "*"
            row.Item("CARRIER_DESC") = "All Carriers"
            dst.Tables("SOTCARR1").Rows.Add(row)

        End With

        Create_Summary(grdSOTCARRA, "CARRIER_CODE", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CARRIER_CODE", False, True)

            Case "Cancel"

                If MessageBox.Show("Do you want to Cancel changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Update"

                Dim warnings As String = String.Empty
                Dim carrierStateList As New List(Of String)

                ' Validate the percentages do not cross.
                Dim CARRIER_CODE As String = String.Empty
                Dim STATE_CODE As String = String.Empty
                Dim START_PERC As Decimal = 0
                Dim END_PERC As Decimal = 0

                For Each rowSOTCARRA As DataRow In dst.Tables("SOTCARRA").Select("", "CARRIER_CODE, STATE_CODE, START_PERC")

                    ' Start a new Carrier / State  
                    If CARRIER_CODE <> rowSOTCARRA.Item("CARRIER_CODE") & String.Empty OrElse STATE_CODE <> rowSOTCARRA.Item("STATE_CODE") & String.Empty Then
                        CARRIER_CODE = rowSOTCARRA.Item("CARRIER_CODE")
                        STATE_CODE = rowSOTCARRA.Item("STATE_CODE")
                        START_PERC = Val(rowSOTCARRA.Item("START_PERC") & String.Empty)
                        END_PERC = Val(rowSOTCARRA.Item("END_PERC") & String.Empty)
                    Else
                        If Val(rowSOTCARRA.Item("START_PERC") & String.Empty) < END_PERC Then
                            EMsg &= vbCr & "Carrier (" & CARRIER_CODE & "), State (" & STATE_CODE & ") has an invalid sequence."
                            Exit Select
                        ElseIf Val(rowSOTCARRA.Item("START_PERC") & String.Empty) <> END_PERC Then
                            If Not carrierStateList.Contains(CARRIER_CODE & "/" & STATE_CODE) Then
                                carrierStateList.Add(CARRIER_CODE & "/" & STATE_CODE)
                                warnings &= vbCr & "There are gaps in Carrier (" & CARRIER_CODE & "), State (" & STATE_CODE & ")"
                            End If
                        End If

                        START_PERC = Val(rowSOTCARRA.Item("START_PERC") & String.Empty)
                        END_PERC = Val(rowSOTCARRA.Item("END_PERC") & String.Empty)

                    End If

                    If START_PERC >= END_PERC Then
                        EMsg &= vbCr & "The Start Percentage (% >) must be less then End Percengtage (% <=). Carrier (" & CARRIER_CODE & "), State (" & STATE_CODE & ")"
                        Exit Select
                    End If
                Next

                If warnings.Length > 0 Then
                    If MessageBox.Show(warnings & Environment.NewLine & Environment.NewLine & "Do you want to Update anyway?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                Else
                    If MessageBox.Show("Do you want to Update changes?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If

                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        grdSOTCARRA.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SOTCARRA").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("CARRIER_CODE").Clear()

        Clear_All_Filters(grdSOTCARRA)

    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading 1099 information")

        If Absx1.txtFor("CARRIER_CODE").TextLength = 0 Then
            Absx1.txtFor("CARRIER_CODE").Text = "*"
        End If

        EnforceConstraints(False)
        Fill_Records("SOTCARRA", String.Empty, True, "SELECT SOTCARRA.*, TATSTATE.STATE_NAME " _
                     & " FROM SOTCARRA, TATSTATE " _
                     & " WHERE SOTCARRA.CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'" _
                     & " AND SOTCARRA.STATE_CODE = TATSTATE.STATE_CODE (+)")
        EnforceConstraints(True)

        Sort_grdColumns(grdSOTCARRA, "CARRIER_CODE,STATE_CODE,START_PERC")

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()
            Update_Record_TDA("SOTCARRA", "DELETE FROM SOTCARRA where CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTCARRA, "SS", "Show Filter", "Show GroupBox")
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

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "Control Procedures"

    Private Sub grdSOTCARRA_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTCARRA.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "STATE_CODE"
                Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").Rows.Find(e.Cell.Value & String.Empty)
                If rowTATSTATE Is Nothing Then
                    e.Cell.Row.Cells("STATE_NAME").Value = String.Empty
                Else
                    e.Cell.Row.Cells("STATE_NAME").Value = rowTATSTATE.Item("STATE_NAME") & String.Empty
                End If
        End Select
    End Sub

    Private Sub grdSOTCARRA_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSOTCARRA.AfterRowsDeleted
        Sort_grdColumns(grdSOTCARRA, "CARRIER_CODE,STATE_CODE,START_PERC")
    End Sub

    Private Sub grdSOTCARRA_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSOTCARRA.AfterRowUpdate
        Sort_grdColumns(grdSOTCARRA, "CARRIER_CODE,STATE_CODE,START_PERC")
    End Sub

    Private Sub grdSOTCARRA_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCARRA.BeforeRowUpdate

        Dim CARRIER_CODE As String = Absx1.txtFor("CARRIER_CODE").Text

        Dim STATE_CODE As String = e.Row.Cells("STATE_CODE").Value
        Dim START_PERC As Decimal = Val(e.Row.Cells("START_PERC").Value & String.Empty)
        Dim END_PERC As Decimal = Val(e.Row.Cells("END_PERC").Value & String.Empty)

        CARRIER_CODE = CARRIER_CODE.Trim.ToUpper
        STATE_CODE = STATE_CODE.Trim.ToUpper

        If dst.Tables("SOTCARR1").Rows.Find(CARRIER_CODE) Is Nothing Then
            MessageBox.Show("Invalid Carrier Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If dst.Tables("TATSTATE").Rows.Find(STATE_CODE) Is Nothing Then
            MessageBox.Show("Invalid State Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        e.Row.Cells("CARRIER_CODE").Value = CARRIER_CODE
        e.Row.Cells("STATE_CODE").Value = STATE_CODE

        If START_PERC >= END_PERC Then
            MessageBox.Show("The Start Percentage (% >) must be less then End Percengtage (% <=).", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

    End Sub

    Private Sub grdSOTCARRA_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTCARRA.ClickCellButton

        Dim sql_where As String = String.Empty

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "CARRIER_CODE"
                    grdClickCellButton(grdSOTCARRA, sql_where)

                Case "STATE_CODE"
                    grdClickCellButton(grdSOTCARRA, sql_where)
            End Select
        End With

    End Sub

#End Region

End Class