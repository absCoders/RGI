Public Class POFWPDM2

    'Public APT As Infragistics.Win.UltraWinSchedule.Appointment
    Public rowPOTWPDM6 As DataRow
    Public UPDATED As Boolean = False
    Public CUST_CODE As String
    Public VEND_CODE As String
    Public TASK_ASSIGNED_TO As String
    Public STYLE_GROUP_NO As String

    Public CUST_NAME As String
    Public VEND_NAME As String
    Public STYLE_GROUP_NAME As String

    Public frmASFBASE0 As ASFBASE0
    Dim WORK_IDs As New Dictionary(Of String, String)

    Dim STEP_LNO As Int32
    Dim TASK_LNO As Int32

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "POTWPDM5", "*")
            Create_TDA(.Tables.Add, "POTWPDM6", "*")
            With .Tables("POTWPDM6").Columns
                .Add("STEP_DESC")
                .Add("STEP_STAGE")
                .Add("STEP_ACTION_DATE_NAME")
                .Add("STEP_ACTION_DATE", GetType(System.DateTime))
                .Add("WORK_PERFORMED")
                .Add("WORK_COMPLETED")
            End With

            'Dim DT As New DataTable("POTWPDMA")
            'DT = frmASFBASE0.dst.Tables("POTWPDM6").Clone

            Create_TDA(.Tables.Add, "POTWPDM7", "*")
            ' Create_TDA(.Tables.Add, "POTWPDM7", "*")
            '.Tables.Add(frmASFBASE0.dst.Tables("POTWPDM5"))
            '.Tables.Add(frmASFBASE0.dst.Tables("POTWPDM6"))
            '.Tables.Add(frmASFBASE0.dst.Tables("POTWPDM7"))
        End With

        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("VEND_CODE").Text = VEND_CODE
        Absx1.txtFor("STYLE_GROUP_NO").Text = STYLE_GROUP_NO

        Absx1.txtFor("CUST_NAME").Text = CUST_NAME
        Absx1.txtFor("VEND_NAME").Text = VEND_NAME
        Absx1.txtFor("STYLE_GROUP_NAME").Text = STYLE_GROUP_NAME

        TABLE_NAME = "POTWPDM6"

        grdASTATTA2.DataSource = frmASFBASE0.dst.Tables("ASTATTA2")

        Dim rowPOTWPDM6 As DataRow = dst.Tables("POTWPDM6").NewRow
        rowPOTWPDM6.ItemArray = Me.rowPOTWPDM6.ItemArray

        STEP_LNO = Val(rowPOTWPDM6.Item("STEP_LNO") & "")
        TASK_LNO = Val(rowPOTWPDM6.Item("TASK_LNO") & "")

        Dim rowPOTWPDM5 As DataRow = frmASFBASE0.dst.Tables("POTWPDM5").Rows.Find(New Object() {STYLE_GROUP_NO, STEP_LNO})
        rowPOTWPDM6.Item("STEP_ACTION_DATE") = rowPOTWPDM5.Item("STEP_ACTION_DATE")

        dst.Tables("POTWPDM6").Rows.Add(rowPOTWPDM6)

        Dim sql As String = "STEP_LNO = " & CStr(STEP_LNO) & " and TASK_LNO = " & CStr(TASK_LNO)

        Dim WORK_ID_previous As String = ""
        For Each row As DataRow In frmASFBASE0.dst.Tables("POTWPDM7").Select(sql, "WORK_ID")
            Dim WORK_ID As String = row.Item("WORK_ID")
            WORK_IDs.Add(WORK_ID, WORK_ID_previous)
            WORK_ID_previous = WORK_ID
        Next

        grdPOTWPDM7.DataSource = frmASFBASE0.dst.Tables("POTWPDM7")

        Dim dvw As DataView

        dvw = DirectCast(grdPOTWPDM7.DataSource, DataTable).DefaultView
        dvw.RowFilter = sql
        Sort_grdColumns(grdPOTWPDM7, "WORK_ID")

        dvw = New DataView(frmASFBASE0.dst.Tables("POTWPDM6"))
        dvw.RowFilter = "STEP_LNO = " & rowPOTWPDM6.Item("STEP_LNO")
        grdPOTWPDM6.DataSource = dvw
        grdPOTWPDM6.Text = "All Tasks in Step " & CStr(STEP_LNO) & ":" & rowPOTWPDM6.Item("STEP_DESC")
        Sort_grdColumns(grdPOTWPDM6, "TASK_ID")

        lblSTEP.Text = "Step " & rowPOTWPDM6.Item("STEP_LNO")
        lblTASK.Text = "Task " & rowPOTWPDM6.Item("TASK_LNO")

        ASCMAIN1.sql = "" ' NECESSARY BECAUSE ASCMAIN1.SQL HAS STUFF IN IT 
        ASCMAIN1.Add_Value_List(cbeTASK_DIR, "TASK_DIR")

        ASCMAIN1.Add_Value_List(grdPOTWPDM6, "TASK_DIR")
        ASCMAIN1.Add_Value_List(grdPOTWPDM6, "TASK_STATUS")

        ASCMAIN1.Add_Value_List(grdPOTWPDM7, "TASK_STATUS")

        Absx1.txtFor("WORK_PERFORMED").Appearance.BackColor = Drawing.Color.LightGreen
        Absx1.chkFor("WORK_COMPLETED").Appearance.BackColor = Drawing.Color.LightGreen
        dteSTEP_ACTION_DATE.Appearance.BackColor = Drawing.Color.LightGreen
        dteTASK_COMPLETED.Appearance.BackColor = Drawing.Color.LightGreen

        dteTASK_COMPLETED.Visible = False
        lblTASK_COMPLETED.Visible = False

        Me.Text = "Work Performed on Step " & CStr(STEP_LNO) & ", Task " & CStr(TASK_LNO) & ": " & rowPOTWPDM6.Item("STEP_DESC") & " - " & rowPOTWPDM6.Item("TASK_DESC") & "; Task ID " & rowPOTWPDM6.Item("TASK_ID")

        Absx1.txtFor("WORK_PERFORMED").Focus()

        If rowPOTWPDM6.Item("STEP_ACTION_DATE_NAME") & "" <> "" Then
            lblSTEP_ACTION_DATE.Visible = True
            dteSTEP_ACTION_DATE.Visible = True
            lblSTEP_ACTION_DATE.Text = rowPOTWPDM6.Item("STEP_ACTION_DATE_NAME") & ""
        End If

        lblNote.Visible = ASCMAIN1.USER_ID <> Absx1.txtFor("TASK_ASSIGNED_TO").Text

        If rowPOTWPDM6.Item("TASK_STATUS") = "C" Then 'If optTASK_STATUS.Value = "C" Then
            cmdUpdate.Visible = False
            cmdCancel.Text = "Done"
            Set_Read_Only_for_ctl(optTASK_STATUS, True)
            Absx1.txtFor("WORK_PERFORMED").Visible = False
            chkComplete.Visible = False
            Set_Read_Only_for_ctl(Absx1.txtFor("TASK_NOTE"), True)
            Set_Read_Only_for_ctl(Absx1.txtFor("TASK_ASSIGNED_TO"), True)
            Set_Read_Only_for_ctl(Absx1.CtlFor("TASK_DUE"), True)
            Set_Read_Only_for_ctl(Absx1.CtlFor("STEP_ACTION_DATE"), True)
            lblWORK_PERFORMED1.Visible = False
            lblWORK_PERFORMED2.Visible = False
        End If

        If frmASFBASE0.EntryMode = "E" Or frmASFBASE0.EntryMode = "N" Then
        Else
            cmdUpdate.Visible = False
            cmdCancel.Text = "Done"
            Set_Read_Only(SplitContainer1.Panel1, True)
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        Dim EMsg As String = ""
        If Absx1.txtFor("TASK_ASSIGNED_TO").Text = "" Then
            EMsg &= vbCr & "Task Owner is Required"
        Else
            If LookUp("ASTUSER1", Absx1.txtFor("TASK_ASSIGNED_TO").Text) Is Nothing Then
                EMsg &= vbCr & "Invalid Value specified for Task Owner"
            End If
        End If

        If Not IsDate(dteTASK_ASSIGNED.Value) OrElse Not IsDate(dteTASK_DUE.Value) Then
            EMsg &= vbCr & "Values for Date Assigned and Date Due are required"
        ElseIf Format(dteTASK_ASSIGNED.Value, "yyyyMMdd") > Format(dteTASK_DUE.Value, "yyyyMMdd") Then
            EMsg &= vbCr & "Date Due may not be prior to Date Assigned"
        End If

        Dim sql_prior_incomplete As String = "STEP_LNO = " & CStr(STEP_LNO) & " and TASK_LNO < " & CStr(TASK_LNO) & " and ISNULL(TASK_STATUS,'?') <> 'C'"
        If chkComplete.Checked Then
            If frmASFBASE0.dst.Tables("POTWPDM6").Select(sql_prior_incomplete).Length <> 0 Then
                If MsgBox("There are other Tasks scheduled before this Task that are not Complete Yet." _
                          & vbCrLf & "Do you want to Complete all of those prior Tasks in this single Update?", _
                          MsgBoxStyle.YesNo, "You are Completing a Task ahead of Prior Tasks") = MsgBoxResult.No Then
                    MsgBox("You Must Complete Tasks in Order", MsgBoxStyle.OkOnly, "Cannot Skip Ahead to Complete Tasks")
                    Exit Sub
                End If
            End If

            If rowPOTWPDM6.Item("STEP_ACTION_DATE_NAME") & "" <> "" Then
                If dteSTEP_ACTION_DATE.Value & "" = "" Then
                    EMsg &= vbCr & "You Must Specify a Value for " & lblSTEP_ACTION_DATE.Text
                End If
            End If
            If optTASK_STATUS.Value = "U" Then
                EMsg &= vbCr & "Cannot set Task Status to Unassigned if you are Completing the Task"
            End If
        Else
            If optTASK_STATUS.Value = "C" Then
                EMsg &= vbCr & "Cannot set Task Status to Complete using Option Set - use the Checkbox to indicate that the Task is Complete"
            End If
        End If


        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Perform the Requested Action")
            Exit Sub
        End If

        DATETIME_STAMP = DateTime.Now + ASCMAIN1.NowTSD
        Synch_TABLE_NAME("POTWPDM6")
        INIT_LAST("POTWPDM6", True, , True)

        Dim rowPOTWPDM6_work As DataRow = dst.Tables("POTWPDM6").Rows(0)

        If rowPOTWPDM6_work.Item("TASK_STATUS") = "U" Then
            rowPOTWPDM6_work.Item("TASK_ASSIGNED_TO") = DBNull.Value
        End If

        Dim rowPOTWPDM6_base As DataRow = frmASFBASE0.dst.Tables("POTWPDM6").Rows.Find _
                    (New Object() {rowPOTWPDM6_work.Item("STYLE_GROUP_NO"), _
                                   rowPOTWPDM6_work.Item("STEP_LNO"), _
                                   rowPOTWPDM6_work.Item("TASK_LNO")})

        rowPOTWPDM6_base.Item("LAST_DATE") = DATETIME_STAMP
        rowPOTWPDM6_base.Item("LAST_OPER") = ASCMAIN1.USER_ID

        rowPOTWPDM6_base.Item("TASK_NOTE") = rowPOTWPDM6_work.Item("TASK_NOTE")
        rowPOTWPDM6_base.Item("TASK_STATUS") = rowPOTWPDM6_work.Item("TASK_STATUS")
        rowPOTWPDM6_base.Item("TASK_DUE") = rowPOTWPDM6_work.Item("TASK_DUE")
        rowPOTWPDM6_base.Item("TASK_ASSIGNED_TO") = rowPOTWPDM6_work.Item("TASK_ASSIGNED_TO")

        If chkComplete.Checked Then
            rowPOTWPDM6_base.Item("TASK_STATUS") = "C"
            rowPOTWPDM6_base.Item("TASK_COMPLETED_BY") = ASCMAIN1.USER_ID
            rowPOTWPDM6_base.Item("TASK_COMPLETED") = rowPOTWPDM6_work.Item("TASK_COMPLETED")

            If frmASFBASE0.dst.Tables("POTWPDM6").Select(sql_prior_incomplete).Length <> 0 Then
                For Each rowPOTWPDM6_prior As DataRow In frmASFBASE0.dst.Tables("POTWPDM6").Select(sql_prior_incomplete)
                    rowPOTWPDM6_prior.Item("TASK_STATUS") = "C"
                    rowPOTWPDM6_prior.Item("TASK_COMPLETED_BY") = ASCMAIN1.USER_ID
                    rowPOTWPDM6_prior.Item("TASK_COMPLETED") = rowPOTWPDM6_work.Item("TASK_COMPLETED")

                    Dim rowPOTWPDM7_prior As DataRow = frmASFBASE0.dst.Tables("POTWPDM7").NewRow
                    For Each dcol As DataColumn In frmASFBASE0.dst.Tables("POTWPDM7").Columns
                        If frmASFBASE0.dst.Tables("POTWPDM6").Columns.Contains(dcol.ColumnName) Then
                            rowPOTWPDM7_prior.Item(dcol.ColumnName) = rowPOTWPDM6_prior.Item(dcol.ColumnName)
                        End If
                    Next
                    ' rowPOTWPDM7_prior.Item("TASK_LNO") = rowPOTWPDM6_prior.Item("TASK_LNO")
                    rowPOTWPDM7_prior.Item("WORK_ID") = ASCMAIN1.Next_Control_No("POTWPDM7.WORK_ID")
                    rowPOTWPDM7_prior.Item("WORK_PERFORMED") = "Completed by Task Id " & rowPOTWPDM6_base.Item("TASK_ID")
                    frmASFBASE0.dst.Tables("POTWPDM7").Rows.Add(rowPOTWPDM7_prior)
                Next
            End If
        End If


        If rowPOTWPDM6_work.Item("STEP_ACTION_DATE_NAME") & "" <> "" AndAlso dteSTEP_ACTION_DATE.Value & "" <> "" Then
            Dim rowPOTWPDM5_base As DataRow = frmASFBASE0.dst.Tables("POTWPDM5").Rows.Find _
                        (New Object() {rowPOTWPDM6_work.Item("STYLE_GROUP_NO"), _
                                       rowPOTWPDM6_work.Item("STEP_LNO")})
            rowPOTWPDM5_base.Item("STEP_ACTION_DATE") = dteSTEP_ACTION_DATE.Value
        End If

        Dim rowPOTWPDM7_base As DataRow = frmASFBASE0.dst.Tables("POTWPDM7").NewRow
        For Each dcol As DataColumn In frmASFBASE0.dst.Tables("POTWPDM7").Columns
            If dst.Tables("POTWPDM6").Columns.Contains(dcol.ColumnName) Then
                rowPOTWPDM7_base.Item(dcol.ColumnName) = rowPOTWPDM6_work.Item(dcol.ColumnName)
            End If
        Next
        rowPOTWPDM7_base.Item("WORK_ID") = ASCMAIN1.Next_Control_No("POTWPDM7.WORK_ID")
        frmASFBASE0.dst.Tables("POTWPDM7").Rows.Add(rowPOTWPDM7_base)

        'Update_Record_TDA("POTWPDM6")
        'Update_Record_TDA("POTWPDM7")
        'Update_Record_TDA("POTWPDM5")

        UPDATED = True

        Me.Close()
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As System.Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)
        Select Case COLUMN_NAME
            'Case "WH_OPER_ID"
            '    If VEHICLE_CODE <> "WH" Then
            '        sql_where = "TATUSER1.SALES_DIVISION_CODE = '" & CUST_CODE & "'"
            '    Else
            '        sql_where = "SALES_DIVISION_CODE = '" & CUST_CODE & "'"
            '    End If
        End Select
    End Sub

    Private Sub grdPOTWPDM6_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTWPDM6.InitializeRow
        If rowPOTWPDM6 Is Nothing Then Exit Sub
        If e.Row.Cells("TASK_LNO").Value = rowPOTWPDM6.Item("TASK_LNO") Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        End If
    End Sub

    Private Sub chkComplete_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkComplete.CheckedChanged
        dteTASK_COMPLETED.Visible = chkComplete.Checked
        lblTASK_COMPLETED.Visible = chkComplete.Checked

        If chkComplete.Checked Then
            dteTASK_COMPLETED.Value = Now.Date
        Else
            dteTASK_COMPLETED.Value = DBNull.Value
        End If
    End Sub

    Private Sub grdPOTWPDM7_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTWPDM7.InitializeRow
        If WORK_IDs.Count = 0 Then Exit Sub
        Dim WORK_ID As String = e.Row.Cells("WORK_ID").Value & ""
        If WORK_IDs.ContainsKey(WORK_ID) Then
            Dim WORK_ID_previous As String = WORK_IDs(WORK_ID)
            If WORK_ID_previous <> "" Then
                Dim row As DataRow = frmASFBASE0.dst.Tables("POTWPDM7").Select("WORK_ID = '" & WORK_ID_previous & "'")(0)
                For Each COLUMN_NAME As String In New String() {"TASK_DESC", "TASK_DIR", "TASK_NOTE", "TASK_STATUS", "TASK_ASSIGNED_TO", "TASK_ASSIGNED", "TASK_DUE", "TASK_COMPLETED_BY", "TASK_COMPLETED", "STEP_ACTION_DATE"}
                    If row.Item(COLUMN_NAME) & "" <> e.Row.Cells(COLUMN_NAME).Value & "" Then
                        e.Row.Cells(COLUMN_NAME).Appearance.BackColor = Drawing.Color.LightPink
                    End If
                Next
            End If
        End If
    End Sub
End Class