Public Class ICFCOSTM
    Public STYLE_CODE As String
    Public STYLE_DESC As String
    Public COLOR_CODE As String
    Public COLOR_DESC As String
    Public updated As Boolean = False
    Public isBatchMarkdown As Boolean = False
    Public batchTranNo As String
    Public markdownBatch As DataTable = Nothing
    Public batchMarkdownColumns As List(Of String)
    Public frmBase As ASFBASE1

    Dim SQL_RECORD_NOS As String = ""
    Dim COLOR_CODEs As New List(Of String)

    Dim OPS_YYYYPP_last_closed As String

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.Text = "Style Cost Maintenance" & IIf(isBatchMarkdown, " - Batch Markdowns: " & batchTranNo, "")

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTCOST1.*, '' IMPORT_CODE" _
                & " from ICTCOST1" _
                & " where ICTCOST1.STYLE_CODE = :PARM1" _
                & "   and ICTCOST1.COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTCOST1", "**", 0, True, "VV", 0)

            If isBatchMarkdown Then
                For Each TABLE_NAME As String In New String() {"ICTCOST1", "ICTCOSTE", "ICTCOSTD", "ICTCOSTN"}
                    If TABLE_NAME <> "ICTCOST1" Then
                        ASCMAIN1.sql = "Select ICTCOST1.*, '' IMPORT_CODE" _
                            & " from ICTCOST1"
                        Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 3)
                    End If
                    For Each COLUMN_NAME As String In batchMarkdownColumns
                        With .Tables(TABLE_NAME)
                            .Columns.Add(COLUMN_NAME, GetType(System.Double))
                        End With
                    Next
                Next
            End If

        End With

        grdICTCOST1.DataSource = dst.Tables("ICTCOST1")

        If isBatchMarkdown Then

            grdICTCOSTE.DataSource = dst.Tables("ICTCOSTE")
            grdICTCOSTD.DataSource = dst.Tables("ICTCOSTD")
            grdICTCOSTN.DataSource = dst.Tables("ICTCOSTN")
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTCOST1, grdICTCOSTE, grdICTCOSTD, grdICTCOSTN}
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.True
                End With
                ASCMAIN1.Add_Value_List(grd, "TRAN_TYPE", Nothing, New String() {"M:Markdown"})
            Next

        Else

            With grdICTCOST1.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.True
            End With

        End If



        SplitContainer2.Panel1Collapsed = isBatchMarkdown
        cmdAdd.Visible = Not isBatchMarkdown

        If Not isBatchMarkdown Then
            grdICTCOST1.Parent = SplitContainer3.Panel1
            SplitContainer3.Panel2Collapsed = True
            Absx1.txtFor("STYLE_CODE").Text = Me.STYLE_CODE
            Absx1.txtFor("STYLE_DESC").Text = Me.STYLE_DESC
            Absx1.txtFor("COLOR_CODE").Text = Me.COLOR_CODE
            Absx1.txtFor("COLOR_DESC").Text = Me.COLOR_DESC
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)

            Absx1.numFor("SUB_UNIT_PACK_QTY").Value = rowICTSTYL1.Item("SUB_UNIT_PACK_QTY")
            Absx1.numFor("CARTON_PACK_QTY").Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
            ASCMAIN1.Add_Value_List(grdICTCOST1, "TRAN_TYPE", Nothing, New String() {":", "M:Markdown", "B:Baseline", "J:Adjustment", "Z:Zero Cost"})
            Fill_Records("ICTCOST1", New String() {STYLE_CODE, COLOR_CODE})

            Create_Summary(grdICTCOST1, "TRAN_NO", "Count")

            COLOR_CODEs.Clear()
            ASCMAIN1.sql = "Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("COLOR_CODE <> '" & COLOR_CODE & "'", "COLOR_CODE")
                COLOR_CODEs.Add(row.Item("COLOR_CODE"))
            Next

            chkAllColors.Checked = (COLOR_CODEs.Count > 0)
            chkAllColors.Visible = (COLOR_CODEs.Count > 0)
            Absx1.dteFor("TRAN_DATE").Value = Now.Date

        Else
            Absx1.dteFor("TRAN_DATE").Value = New DateTime(Now.Year, Now.Month, 1)
            Absx1.dteFor("TRAN_DATE").Enabled = False
            chkAllowZeroDollar.Visible = False
            numTRAN_COST.Visible = False
            lblTRAN_COST.Visible = False

            grdICTCOST1.Parent = tabImportResults.Tabs("Valid Records").TabPage
            SplitContainer3.Panel1Collapsed = True
            Setup_Batch_Markdowns()
        End If




        ASCMAIN1.sql = "Select MAX(OPS_YYYYPP) from ICTCOSTP where UPDATED = '1'"
        OPS_YYYYPP_last_closed = ASCDATA1.GetDataValue




    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"

        End Select
    End Sub
#End Region
    Private Sub Setup_Batch_Markdowns()
        Dim badPeriod As String = ""
        Dim recordsFailed As Boolean = False
        If markdownBatch IsNot Nothing Then
            For Each row As DataRow In markdownBatch.Select()
                Dim TBL As String = "ICTCOST1"
                Dim TRAN_NO As String = row.Item("TRAN_NO") & ""
                Dim TRAN_REF As String = "Batch MARK"
                Dim STYLE_CODE As String = row.Item("STYLE_CODE") & ""
                Dim COLOR_CODE As String = row.Item("COLOR_CODE") & ""
                Dim OPS_YYYYPP As String = row.Item("OPS_YYYYPP") & ""
                Dim recordExists As Boolean = False

                If InStr(SQL_RECORD_NOS, TRAN_NO) = 0 Then
                    SQL_RECORD_NOS &= ",'" & TRAN_NO & "'"
                End If

                Dim sqlRecordExists As String = "Select * from ICTCOST1 where STYLE_CODE = :PARM1" & vbCrLf _
                                                & " AND COLOR_CODE = :PARM2 AND OPS_YYYYPP = :PARM3"

                Dim rowCheck As DataRow = ASCDATA1.GetDataRow(sqlRecordExists, "VVV", New Object() {STYLE_CODE, COLOR_CODE, OPS_YYYYPP})
                If rowCheck IsNot Nothing Then
                    TBL = "ICTCOST1"
                    recordExists = True
                    'TRAN_REF = "Dup Record"
                    'recordsFailed = False
                Else
                    Dim currPeriod As String = IIf(ASCMAIN1.Running_in_VS, "201804", ASCMAIN1.CYP) 'rdw - old db
                    If OPS_YYYYPP <> currPeriod Then

                        'recordsFailed = True
                    End If
                    Select Case row.Item("IMPORT_CODE")
                        Case "0" ' VALID

                        Case "E" 'INVALIE PERIOD
                            TBL = "ICTCOSTE"
                            TRAN_REF = "Bad Period"
                        Case "N" 'NO MARKDOWN
                            TBL = "ICTCOSTN"
                            TRAN_REF = "Null New $"

                    End Select
                End If

                Dim markdownRecord As DataRow = dst.Tables(TBL).NewRow
                markdownRecord.ItemArray = row.ItemArray
                markdownRecord.Item("TRAN_REF") = TRAN_REF
                If recordExists Then
                    Dim duplicateMarkdown As DataRow = dst.Tables("ICTCOSTD").NewRow
                    duplicateMarkdown.ItemArray = rowCheck.ItemArray
                    dst.Tables("ICTCOSTD").Rows.Add(duplicateMarkdown)
                End If


                dst.Tables(TBL).Rows.Add(markdownRecord)
            Next
        End If

        grdICTCOST1.Text = "Valid Markdowns"
        grdICTCOSTD.Text = "Existing Markdowns - Will be replaced"
        grdICTCOSTE.Text = "Invalid Records"
        grdICTCOSTN.Text = "No Changes"

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
    {grdICTCOST1, grdICTCOSTD, grdICTCOSTE, grdICTCOSTN}
            With grd.DisplayLayout.Bands(0)
                .Columns("STYLE_CODE").Hidden = False
                .Columns("COLOR_CODE").Hidden = False
                .Columns("STYLE_CODE").Header.Caption = "Style"
                .Columns("COST_NOW").Header.Caption = "Cost Now"
                .Columns("ON_HAND").Header.Caption = "On Hand"
                .Columns("VALUE_NOW").Header.Caption = "Value Now"
                .Columns("COST_NEW").Header.Caption = "Cost New"
                .Columns("VALUE_NEW").Header.Caption = "Value New"
                .Columns("MARKDOWN").Header.Caption = "MD Amt"
                .Columns("TRAN_TYPE").Hidden = True
                .Columns("TRAN_NO").Hidden = True
                For Each COLUMN_NAME As String In batchMarkdownColumns
                    With .Columns(COLUMN_NAME)
                        .Hidden = False
                        .Header.Appearance.TextHAlign = HAlign.Right
                        .CellAppearance.TextHAlign = HAlign.Right
                        If COLUMN_NAME.StartsWith("COST") Then
                            .Width = 80
                            .Format = "#,##0.0000"
                            If COLUMN_NAME = "COST_NEW" Then
                                .CellAppearance.BackColor = Drawing.Color.Yellow
                            End If
                        Else
                            .Width = 100
                            .Format = "#,##0.00"
                            Create_Summary(grd, COLUMN_NAME, "Sum")
                        End If
                    End With
                Next
            End With
            Create_Summary(grd, "STYLE_CODE", "Count")
        Next

        optTRAN_TYPE.Enabled = Not isBatchMarkdown

        'If recordsFailed Then
        '    Using fr As New ASFMSGBF
        '        fr.Show_grd(dst.Tables("ICTCOSTE"), frmBase, "Some Rows Failed to Import - Please Check Tran Ref Column for Messages")
        '    End Using
        'End If

    End Sub
    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        BeginTrans()

        If isBatchMarkdown Then
            For Each rowD As DataRow In dst.Tables("ICTCOSTD").Select()
                Dim STYLE_CODE As String = rowD.Item("STYLE_CODE") & ""
                Dim COLOR_CODE As String = rowD.Item("COLOR_CODE") & ""
                Dim TRAN_NO As String = rowD.Item("TRAN_NO") & ""
                ASCMAIN1.sql = "Delete from ICTCOST1 WHERE STYLE_CODE = :PARM1 AND COLOR_CODE = :PARM2 AND TRAN_NO = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {STYLE_CODE, COLOR_CODE, TRAN_NO})
            Next

            Update_Record_TDA("ICTCOST1")
        Else
            Update_Record_TDA("ICTCOST1", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
        End If

        If SQL_RECORD_NOS <> "" Then
            ASCMAIN1.sql = "INSERT INTO ICTCOSTL " & vbCrLf _
                & " (OPS_YYYYPP_FIFO,STYLE_CODE,COLOR_CODE,OPS_YYYYPP,TRAN_DATE,RECORD_NO" & vbCrLf _
                & ",TRAN_TYPE,TRAN_REF,TRAN_QTY,TRAN_COST,TRAN_NO,TRAN_LNO)" & vbCrLf _
                & " Select OPS_YYYYPP, STYLE_CODE, COLOR_CODE, OPS_YYYYPP, TRAN_DATE, 'L' || SUBSTR(TRAN_NO,2) RECORD_NO" & vbCrLf _
                & ", TRAN_TYPE, TRAN_REF, TRAN_QTY, TRAN_COST, TRAN_NO, 0 TRAN_LNO" & vbCrLf _
                & " from ICTCOST1 where ICTCOST1.TRAN_NO in (" & Mid(SQL_RECORD_NOS, 2) & ")"
            ASCDATA1.ExecuteSQL()
        End If

        CommitTrans("Update Successful")

        updated = True
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdAdd_Click(sender As System.Object, e As System.EventArgs) Handles cmdAdd.Click
        ' note that we are not supporting Baseline, because I could not understand the difference between M and B

        ASCMAIN1.sql = "Select MIN(OPS_YYYYPP) from GLTPARM2 where PRD_END_DATE >= :PARM1"
        Dim OPS_YYYYPP As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "D", New Object() {Absx1.dteFor("TRAN_DATE").Value})

        If OPS_YYYYPP_last_closed <> "" And OPS_YYYYPP <= OPS_YYYYPP_last_closed Then
            If OPS_YYYYPP > "201300" And (ASCMAIN1.USER_ID = "wjz" Or ASCMAIN1.USER_ID = "anna" Or ASCMAIN1.USER_ID = "lshalom") Then
                MsgBox("You are using a temporary feature to allow you to update costs in a previously closed period." & vbCrLf & vbCrLf & "Proceed with caution.", MsgBoxStyle.OkOnly, "Cannot Add Cost Record into Closed Costing Period")
            Else
                MsgBox("Period " & OPS_YYYYPP & " has been Closed", MsgBoxStyle.OkOnly, "Cannot Add Cost Record into Closed Costing Period")
                Exit Sub
            End If
        End If

        Dim TRAN_TYPE As String = Absx1.optFor("TRAN_TYPE").Value

        If TRAN_TYPE = "M" Then
            If dst.Tables("ICTCOST1").Select("TRAN_TYPE = 'M' and OPS_YYYYPP = '" & OPS_YYYYPP & "'").Length > 0 Then
                MsgBox("Period " & OPS_YYYYPP & " already has a Markdown Record", MsgBoxStyle.OkOnly, "Cannot have 2 Markdowns in Same Period")
                Exit Sub
            End If
        ElseIf TRAN_TYPE = "J" Then
            If Val(Absx1.numFor("TRAN_QTY").Value & "") = 0 Then
                MsgBox("Cost Adjustment Records must have a Non-Zero Qty", MsgBoxStyle.OkOnly, "Cannot have 0 Qty for Cost Adjustment")
                Exit Sub
            End If
        End If

        If Val(Absx1.numFor("TRAN_COST").Value & "") = 0 Then
            If MsgBox("The record you are establishing has a 0 value for cost." _
                       & vbCrLf & vbCrLf & "OK to continue with Update?", _
                       MsgBoxStyle.YesNo, "Verificatin") = MsgBoxResult.No Then Exit Sub
        End If

        Dim rowICTCOST1 As DataRow = dst.Tables("ICTCOST1").NewRow
        With rowICTCOST1
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            Dim TRAN_NO As String = ASCMAIN1.Next_Control_No("ICTCOST1.TRAN_NO", 1)
            If OPS_YYYYPP_last_closed <> "" And OPS_YYYYPP <= OPS_YYYYPP_last_closed Then
                SQL_RECORD_NOS &= ",'" & TRAN_NO & "'"
            End If
            .Item("TRAN_NO") = TRAN_NO
            .Item("TRAN_TYPE") = Absx1.optFor("TRAN_TYPE").Value
            .Item("TRAN_REF") = ""
            .Item("TRAN_DATE") = Absx1.dteFor("TRAN_DATE").Value
            .Item("OPS_YYYYPP") = OPS_YYYYPP
            If TRAN_TYPE = "M" Then
            Else
                .Item("TRAN_QTY") = Absx1.numFor("TRAN_QTY").Value
            End If
            .Item("TRAN_COST") = Absx1.numFor("TRAN_COST").Value
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
        End With
        dst.Tables("ICTCOST1").Rows.Add(rowICTCOST1)

        If chkAllColors.Checked Then
            For Each COLOR_CODE As String In COLOR_CODEs
                Dim rowICTCOST1X As DataRow = dst.Tables("ICTCOST1").NewRow
                rowICTCOST1X.ItemArray = rowICTCOST1.ItemArray
                rowICTCOST1X.Item("COLOR_CODE") = COLOR_CODE
                dst.Tables("ICTCOST1").Rows.Add(rowICTCOST1X)
            Next
        End If

    End Sub

    Private Sub grdICTCOST1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTCOST1.AfterRowActivate
        If isBatchMarkdown Then
            Absx1.numFor("TRAN_COST").Value = Val(grdICTCOST1.ActiveRow.Cells("TRAN_COST").Text & "")
        End If
    End Sub

    Private Sub grdICTCOST1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdICTCOST1.AfterRowsDeleted

    End Sub

    Private Sub grdICTCOST1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTCOST1.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim OPS_YYYYPP As String = grow.Cells("OPS_YYYYPP").Value
            If OPS_YYYYPP <= OPS_YYYYPP_last_closed Then
                If OPS_YYYYPP > "201300" And (ASCMAIN1.USER_ID = "wjz" Or ASCMAIN1.USER_ID = "anna" Or ASCMAIN1.USER_ID = "lshalom") Then
                    MsgBox("You are using a temporary feature to allow you to update costs in a previously closed period." & vbCrLf & vbCrLf & "Proceed with caution.", MsgBoxStyle.OkOnly, "Cannot Delete Cost Record when Costing Period is Closed")
                Else
                    MsgBox("Period " & OPS_YYYYPP & " has been Closed", MsgBoxStyle.OkOnly, "Cannot Add Cost Record into Closed Costing Period")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Next
        e.DisplayPromptMsg = False
    End Sub

    Private Sub optTRAN_TYPE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optTRAN_TYPE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub

        lblTRAN_QTY.Visible = Not (optTRAN_TYPE.Value = "M")
        numTRAN_QTY.Visible = Not (optTRAN_TYPE.Value = "M")

    End Sub

    Private Sub chkAllColors_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAllColors.CheckedChanged
        grdICTCOST1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = Not chkAllColors.Checked
    End Sub

    Private Sub tabImportResults_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabImportResults.SelectedTabChanged

    End Sub
End Class