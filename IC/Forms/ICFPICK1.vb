Imports ABSolution

Public Class ICFPICK1

    Private CUST_CODE As String = String.Empty
    Private PICK_NO As String = String.Empty
    Private ORDR_CUST_PO As String = String.Empty
    Private wkTable As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            InquiryMode = MENU_ITEM_OBJECT = "ICFPICKI"


            Create_TDA(.Tables.Add, "SOTSHIP1", "*", , , , , "SHIP_DATE_RECEIVED,SHIPPED_ACTUAL,SHIP_NOTES,SHIP_DATE_PACKED")
            Create_Lookup("ARTCUST1")

            ASCMAIN1.sql = " Select  0 LINE_NO, SHIP_DATE_RECEIVED, min(sotpick1.ordr_no) min_ordr_no, max(sotpick1.ordr_no) max_ordr_no, min(sotpick1.pick_no) min_pick_no, "
            ASCMAIN1.sql &= " max(sotpick1.pick_no) max_pick_no, count(sotpick1.pick_no) tkts, sotordr0.cust_code, sotordr0.ordr_cust_po, sotordr0.ordr_ship_date, sotordr0.ordr_cancel_date,"
            ASCMAIN1.sql &= " SHIP_DATE_SHIPPED, sotship1.shipped_actual, sotsvia1.ship_via_desc, sotship1.ship_notes, sotship1.ship_bol_no,sotship1.SHIP_PICK_PRINTED, sotpick1.pick_status, sotship1.SHIP_DATE_PACKED,"
            ASCMAIN1.sql &= " SOTORDR0.ordr_group_no"
            ASCMAIN1.sql &= " From sotship1, sotordr0, sotpick1, sotsvia1, SOTORDR1 "
            ASCMAIN1.sql &= " where sotship1.ship_bol_no = sotpick1.ship_bol_no and sotship1.ordr_group_no = sotordr0.ordr_group_no and"
            ASCMAIN1.sql &= " sotship1.ship_via_code = sotsvia1.ship_via_code (+) "
            ASCMAIN1.sql &= " and sotordr0.ORDR_GROUP_NO = sotordr1.ORDR_GROUP_NO and rownum < 1"
            ASCMAIN1.sql &= " group by  sotordr0.cust_code, sotordr0.ordr_cust_po, sotordr0.ordr_ship_date, sotship1.ship_bol_no, SHIP_DATE_RECEIVED, SHIP_DATE_SHIPPED, "
            ASCMAIN1.sql &= " sotordr0.ordr_cancel_date, sotsvia1.ship_via_desc, sotship1.shipped_actual, sotship1.ship_notes, sotship1.SHIP_PICK_PRINTED, sotpick1.pick_status, sotship1.SHIP_DATE_PACKED, SOTORDR0.ordr_group_no"
            Create_TDA(.Tables.Add, "SOTSHIPX", ASCMAIN1.sql, 0, False)

            'ASCMAIN1.sql = " select 0 LINE_NO, SHIP_DATE_RECEIVED, min(sotpick1.ordr_no) min_ordr_no, max(sotpick1.ordr_no) max_ordr_no, min(sotpick1.pick_no) min_pick_no, "
            'ASCMAIN1.sql &= " max(sotpick1.pick_no) max_pick_no, count(sotpick1.pick_no) tkts, sotordr0.cust_code, sotordr0.ordr_cust_po, sotordr0.ordr_ship_date, sotordr0.ordr_cancel_date,"
            'ASCMAIN1.sql &= " SHIP_DATE_SHIPPED, sotship1.shipped_actual, sotsvia1.ship_via_desc, sotship1.ship_notes, sotship1.ship_bol_no,sotship1.SHIP_PICK_PRINTED, sotpick1.pick_status, sotship1.SHIP_DATE_PACKED"
            'ASCMAIN1.sql &= " From sotship1, sotordr0, sotpick1, sotsvia1, SOTORDR1 "
            'ASCMAIN1.sql &= " where sotship1.ship_bol_no = sotpick1.ship_bol_no and sotship1.ordr_group_no = sotordr0.ordr_group_no and"
            'ASCMAIN1.sql &= " sotship1.ship_via_code = sotsvia1.ship_via_code (+) "
            'ASCMAIN1.sql &= " and sotordr0.ORDR_GROUP_NO = sotordr1.ORDR_GROUP_NO and rownum < 1"
            'ASCMAIN1.sql &= " group by sotordr0.cust_code, sotordr0.ordr_cust_po, sotordr0.ordr_ship_date, sotship1.ship_bol_no, SHIP_DATE_RECEIVED, SHIP_DATE_SHIPPED, "
            'ASCMAIN1.sql &= " sotordr0.ordr_cancel_date, sotsvia1.ship_via_desc, sotship1.shipped_actual, sotship1.ship_notes, sotship1.SHIP_PICK_PRINTED, sotpick1.pick_status, sotship1.SHIP_DATE_PACKED"
            'Create_TDA(.Tables.Add, "SOTSHIPC", ASCMAIN1.sql, 0, False)
        End With

        grdICWPICK1.DataSource = dst.Tables("SOTSHIPX")
        Create_Summary(grdICWPICK1, "LINE_NO", "Count")

        Absx1.dteFor("DTE0").MinDate = DateAdd(DateInterval.Year, -2, DateTime.Now)
        Absx1.dteFor("DTE1").MinDate = DateAdd(DateInterval.Year, -2, DateTime.Now)
        Absx1.dteFor("DTE2").MinDate = DateAdd(DateInterval.Year, -1, DateTime.Now)

        Absx1.dteFor("DTE0").DateTime = CDate(DateAdd(DateInterval.Day, -90, DateTime.Now).ToString("MM/dd/yyyy"))
        Absx1.dteFor("DTE1").DateTime = CDate(DateTime.Now.ToString("MM/dd/yyyy"))
        Absx1.dteFor("DTE2").DateTime = CDate(DateTime.Now.ToString("MM/dd/yyyy"))

        txtNotes.MaxLength = dst.Tables("SOTSHIP1").Columns("SHIP_NOTES").MaxLength

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Customer()

            Case "Cancel"

            Case "Update"
                'If EMsg.Length = 0 Then
                '    Excel_Export(grdICWPICK1)
                'End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

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

                '.Groups("Select Using").Visible = Not tf
                .Groups("Mass Update").Visible = tf
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpSelectUsing, ScreenMode)
        grdICWPICK1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()

        EnforceConstraints(False)

        For Each tableName As String In New String() {"SOTSHIP1", "SOTSHIPX"}
            dst.Tables(tableName).Rows.Clear()
        Next

        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Clear()
        Absx1.txtFor("PICK_NO").Clear()
        Absx1.txtFor("ORDR_CUST_PO").Clear()

        txtNotes.Clear()
        chkMassUpdate.Checked = False

        CUST_CODE = String.Empty
        PICK_NO = String.Empty
        ORDR_CUST_PO = String.Empty

        grdICWPICK1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        With grdICWPICK1.DisplayLayout.Override
            If InquiryMode Then
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            Else
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.True
            End If
        End With

    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        EnforceConstraints(False)

        CUST_CODE = Absx1.txtFor("CUST_CODE").Text.Trim
        PICK_NO = Absx1.txtFor("PICK_NO").Text.Trim
        ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text.Trim

        Dim sql As String = String.Empty

        '   Load Data into Work File(s)
        sql = " select 0 LINE_NO, SHIP_DATE_RECEIVED, min(sotpick1.ordr_no) min_ordr_no, max(sotpick1.ordr_no) max_ordr_no, min(sotpick1.pick_no) min_pick_no, "
        sql &= " max(sotpick1.pick_no) max_pick_no, count(sotpick1.pick_no) tkts, sotordr0.cust_code, sotordr0.ordr_cust_po, sotordr0.ordr_ship_date, sotordr0.ordr_cancel_date,"
        sql &= " SHIP_DATE_SHIPPED, sotship1.shipped_actual, sotsvia1.ship_via_desc, sotship1.ship_notes, sotship1.ship_bol_no,sotship1.SHIP_PICK_PRINTED, sotpick1.pick_status, sotship1.SHIP_DATE_PACKED, SOTORDR0.ordr_group_no "
        sql &= " From sotship1, sotordr0, sotpick1, sotsvia1 "
        sql &= " where sotship1.ship_bol_no = sotpick1.ship_bol_no and sotship1.ordr_group_no = sotordr0.ordr_group_no and"
        sql &= " sotship1.ship_via_code = sotsvia1.ship_via_code (+)"

        If CUST_CODE <> "" Then
            sql &= " and sotordr0.CUST_CODE = '" & CUST_CODE & "'"
        End If

        If PICK_NO <> "" Then
            sql &= " and sotpick1.PICK_NO = " & PICK_NO & ""
        End If

        If ORDR_CUST_PO <> "" Then
            sql &= " and sotordr0.ORDR_CUST_PO = '" & ORDR_CUST_PO & "'"
        End If

        If chkUD.Checked Then

            Select Case optUD.Value

                Case "SD"
                    sql &= " and sotordr0.ordr_ship_date >= '" & Absx1.dteFor("DTE0").DateTime.ToString("dd-MMM-yy") & "' "
                    sql &= " and sotordr0.ordr_ship_date <= '" & Absx1.dteFor("DTE1").DateTime.ToString("dd-MMM-yy") & "' "

                Case "DS"
                    sql &= " and sotship1.shipped_actual >= '" & Absx1.dteFor("DTE0").DateTime.ToString("dd-MMM-yy") & "' "
                    sql &= " and sotship1.shipped_actual <= '" & Absx1.dteFor("DTE1").DateTime.ToString("dd-MMM-yy") & "' "

                Case "RD"
                    sql &= " and sotship1.SHIP_PICK_PRINTED >= '" & Absx1.dteFor("DTE0").DateTime.ToString("dd-MMM-yy") & "' "
                    sql &= " and sotpick1.pick_status <> 'F'"
                    sql &= " and sotship1.SHIP_PICK_PRINTED < '" & DateAdd(DateInterval.Day, 1, Absx1.dteFor("DTE1").DateTime).ToString("dd-MMM-yy") & "' "

                Case "DR"
                    sql &= " and sotship1.ship_date_received >= '" & Absx1.dteFor("DTE0").DateTime.ToString("dd-MMM-yy") & "' "
                    sql &= " and sotship1.ship_date_received <= '" & Absx1.dteFor("DTE1").DateTime.ToString("dd-MMM-yy") & "' "
            End Select

        End If

        sql &= " group by sotordr0.cust_code, sotordr0.ordr_cust_po, sotordr0.ordr_ship_date, sotship1.ship_bol_no"
        sql &= " ,sotordr0.ordr_cancel_date, sotsvia1.ship_via_desc, SHIP_DATE_RECEIVED, ship_date_shipped, sotship1.shipped_actual, sotship1.ship_notes, sotship1.SHIP_PICK_PRINTED, sotpick1.pick_status, sotship1.SHIP_DATE_PACKED, SOTORDR0.ordr_group_no"

        Try
            If wkTable.Length = 0 Then
                wkTable = ASCMAIN1.Temp_Table(sql)
            Else
                ASCDATA1.ExecuteSQL("Truncate Table " & wkTable)
                ASCDATA1.ExecuteSQL("Insert Into " & wkTable & " " & sql)
            End If

            ASCDATA1.ExecuteSQL("Update " & wkTable & " set LINE_NO = Rownum")

            Fill_Records("SOTSHIPX", String.Empty, True, "SELECT * FROM " & wkTable)
            Fill_Records("SOTSHIP1", String.Empty, True, "SELECT * FROM SOTSHIP1 WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & wkTable & ")")

        Catch ex As Exception
            dst.Tables("SOTSHIPX").Clear()
            dst.Tables("SOTSHIP1").Clear()
            MessageBox.Show(ex.Message, "Load Data", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            EnforceConstraints(True)
        End Try

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Private Sub Update_Record()

        Try
            ASCMAIN1.Progress("Processing", String.Empty)

            For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("", "", DataViewRowState.CurrentRows)

                ASCMAIN1.Progress("-", rowSOTSHIPX.Item("SHIP_BOL_NO"))

                If rowSOTSHIPX.RowState = DataRowState.Unchanged Then
                    Continue For
                End If

                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", rowSOTSHIPX.Item("SHIP_BOL_NO"), , , False) Then
                    Continue For
                End If

                If Not ASCMAIN1.Logical_Lock("SOTORDR0", rowSOTSHIPX.Item("ORDR_GROUP_NO"), , , False) Then
                    Continue For
                End If

                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(rowSOTSHIPX.Item("SHIP_BOL_NO"))
                If rowSOTSHIP1 Is Nothing Then
                    Continue For
                End If

                rowSOTSHIP1.Item("SHIP_DATE_RECEIVED") = rowSOTSHIPX.Item("SHIP_DATE_RECEIVED")
                rowSOTSHIP1.Item("SHIPPED_ACTUAL") = rowSOTSHIPX.Item("SHIPPED_ACTUAL")
                rowSOTSHIP1.Item("SHIP_NOTES") = rowSOTSHIPX.Item("SHIP_NOTES")
                rowSOTSHIP1.Item("SHIP_DATE_PACKED") = rowSOTSHIPX.Item("SHIP_DATE_PACKED")
            Next

            Try
                BeginTrans()
                Update_Record_TDA("SOTSHIP1")
                CommitTrans("Update Successful")
            Catch ex As Exception
                Rollback(ex.Message)
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

#End Region

#Region "Form Procedures"

    Sub Validate_Customer()

        CUST_CODE = String.Empty

        Absx1.txtFor("CUST_CODE").Text = Absx1.txtFor("CUST_CODE").Text.ToUpper.Trim
        If Absx1.txtFor("CUST_CODE").TextLength = 0 Then
            Exit Sub
        End If

        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)

        If rowARTCUST1 Is Nothing Then
            EMsg &= vbCr & CUST_CODE & " is not a valid cust code"
            Exit Sub
        End If

        If rowARTCUST1.Item("CUST_STATUS") & String.Empty = String.Empty Then
            EMsg &= vbCr = "Customer Does Not have a Valid Status Code"
            Exit Sub
        ElseIf rowARTCUST1.Item("CUST_STATUS") & String.Empty <> "A" Then
            EMsg &= vbCr & "Customer is not Active"
            Exit Sub
        Else
            CUST_CODE = Absx1.txtFor("CUST_CODE").Text
        End If

    End Sub

#End Region

#Region "Form Controls"

    Private Sub chkMassUpdate_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMassUpdate.CheckedChanged
        optMass.Enabled = chkMassUpdate.Checked
        txtNotes.Enabled = chkMassUpdate.Checked AndAlso optMass.Value = "N"
        cmdMS.Enabled = chkMassUpdate.Checked
    End Sub

    Private Sub chkUD_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkUD.CheckedChanged
        optUD.Enabled = chkUD.Checked
    End Sub

    Private Sub cmdMS_Click(sender As System.Object, e As System.EventArgs) Handles cmdMS.Click

        Dim UPDATE_COL As String = String.Empty
        Dim UPDATE_FRM As String = String.Empty
        Dim sql As String = String.Empty
        Dim row As DataRow = Nothing

        If grdICWPICK1.Selected.Rows.Count <> 0 Then
            For i As Integer = 0 To grdICWPICK1.Selected.Rows.Count - 1
                UPDATE_FRM = Format(Absx1.dteFor("DTE2").DateTime, "MM/dd/yyyy")
                Select Case optMass.Value
                    Case Is = "R"
                        UPDATE_COL = "SHIP_DATE_RECEIVED"
                    Case Is = "S"
                        UPDATE_COL = "SHIPPED_ACTUAL"
                    Case Is = "P"
                        UPDATE_COL = "SHIP_DATE_PACKED"
                    Case Is = "N"
                        UPDATE_COL = "SHIP_NOTES"
                        txtNotes.Text = txtNotes.Text.Replace("*", "").Replace("/", "").Replace("'", "").Replace("#", "")
                        txtNotes.Text = txtNotes.Text.Replace(Space(2), Space(1))
                        UPDATE_FRM = txtNotes.Text
                End Select

                row = dst.Tables("SOTSHIPX").Select("SHIP_BOL_NO = '" & grdICWPICK1.Selected.Rows(i).Cells("SHIP_BOL_NO").Value & "'")(0)
                row.Item(UPDATE_COL) = UPDATE_FRM
            Next i
        End If

        If optMass.Value = "N" Then
            txtNotes.Clear()
        End If

    End Sub

    Private Sub grdICWPICK1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICWPICK1.InitializeRow
        If e.Row.Cells("PICK_STATUS").Value & String.Empty = "D" Then
            e.Row.Appearance.BackColor = Color.Red
        End If
    End Sub

    Private Sub optMass_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optMass.ValueChanged
        txtNotes.Enabled = optMass.Value = "N"
    End Sub

#End Region

End Class