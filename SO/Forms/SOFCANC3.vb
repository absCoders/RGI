Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

Public Class SOFCANC3

    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow
    Dim SOTCANCY As String = ""
    Dim sqlORDR_GROUP_NOs As String
    Dim O_ORDR_GROUP_NOs As New List(Of String)
    Dim sqlPICK_NOs = "','"


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
        Create_Temp_Tables(True)

        With dst

            ASCMAIN1.sql = "Select SOTCANCY.*" & vbCrLf _
                & " from " & SOTCANCY & " SOTCANCY"
            MyBase.Create_TDA(.Tables.Add, "SOTCANCY", "**", 0, False, "", 7)

            ASCMAIN1.sql = "select SOTPICK1.PICK_NO, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, SOTPICK1.SHIP_BOL_NO, SOTORDR1.ORDR_NO, " & vbCrLf _
                & " SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_DATE, SOTPICK1.PICK_RELEASED, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & " from SOTPICK1, SOTORDR1 " & vbCrLf _
                & " where SOTORDR1.CUST_CODE = :PARM1 AND SOTORDR1.WHSE_CODE = :PARM2 " & vbCrLf _
                & "and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO " & vbCrLf _
                & "and SOTORDR1.ORDR_STATUS in ('P','O') " & vbCrLf _
                & "and SOTORDR1.ORDR_TYPE_CODE = 'B2C' " & vbCrLf _
                & "and SOTORDR1.ORDR_SOURCE = 'E' " & vbCrLf _
                & "and SOTPICK1.PICK_STATUS = 'P' " & vbCrLf
            MyBase.Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "VV", 1)
            With .Tables("SOTPICKX")
                .Columns.Add("SELECTED", GetType(System.String)) ', "IIF(ordr_qty - (ordr_qty_open + ordr_qty_canc) = 0 ,'1','0')")
                .Columns("SELECTED").DefaultValue = "1"
            End With

            ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV2,SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.CUST_CODE = :PARM1 " & vbCrLf _
                & "   and SOTRSRV2.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTRSRV2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_QTY_OPEN > 0" & vbCrLf
            Create_TDA(.Tables.Add, "SOTRSRVX", "**", 0, False, "VVV", 0)

            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & " where ORDR_NO = :PARM1"
            MyBase.Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, True, "V", 1, "ORDR_STATUS")

            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & " from SOTORDR2" & vbCrLf _
                & " where ORDR_NO = :PARM1"
            MyBase.Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2, "ORDR_QTY_OPEN,ORDR_QTY_CANC,ORDR_STATUS")

            Create_TDA(.Tables.Add, "SOTRSRV1", "*")
            Create_TDA(.Tables.Add, "SOTRSRV2", "*")


            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_OPEN from ICTSTAT2"
            Create_TDA(.Tables.Add, "ICTSTATO", "**", 0, False, String.Empty, 3)
        End With

        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        grdSOTCANCY.DataSource = dst.Tables("SOTCANCY")

        Create_Summary(grdSOTPICKX, "ORDR_NO", "Count")
        Create_Summary(grdSOTPICKX, "SELECTED", "Sum")
        Show_Filter(grdSOTPICKX, True)

        grdSOTPICKX.DisplayLayout.Bands(0).Columns("SELECTED").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdSOTPICKX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "ORDR_NO", "ORDR_CUST_PO", "ORDR_GROUP_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key.StartsWith("ORDR_AMT") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    GCOL.Width = 80
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTPICKX, GCOL.Key)
                ElseIf GCOL.Key.StartsWith("ORDR_QTY") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    GCOL.Width = 70
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTPICKX, GCOL.Key)
                ElseIf GCOL.Key.StartsWith("ORDR_CNT") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                    GCOL.Width = 50
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTPICKX, GCOL.Key)
                ElseIf New String() {"ORDR_NO", "ORDR_GROUP_NO", "ORDR_CUST_PO", "SHIP_BOL_NO", "PICK_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightCoral
                    GCOL.Width = 110
                ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "PICK_RELEASED", "ORDR_CANCEL_DATE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    GCOL.Width = 90
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdSOTCANCY, "ORDR_NO", "Count")
        With grdSOTCANCY.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"PICK_NO", "ORDR_NO", "ORDR_STATUS"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key.StartsWith("ORDR_QTY") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    GCOL.Width = 70
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTCANCY, GCOL.Key)
                ElseIf GCOL.Key.StartsWith("COUNT") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    GCOL.Width = 70
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTCANCY, GCOL.Key)
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View", "Edit"

                CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

                If CUST_CODE = "" Then
                    EMsg &= vbCr & "No Customer Defined"
                Else
                    If grdSOTPICKX.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Orders in Selection Grid"
                    Else
                        Dim rows() As DataRow = dst.Tables("SOTPICKX").Select("SELECTED='1'")
                        If rows.Length = 0 Then
                            EMsg &= vbCr & "No Orders Selected"
                        Else
                            If rows(0).Item("CUST_CODE") <> CUST_CODE Then
                                EMsg &= vbCr & "Orders in Selection grid do not appear to belong to Customer Defined"
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Code"
                    End If
                    rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Warehouse Code"
                    End If
                End If

                'RER - removed below I am handling reservations, and selection criteria is only for B2C EDI, although it could be open to others as well.
                'If eItemKey = "Edit" Then
                '    'We should only see e-comm orders here, check for that
                '    For Each row As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                '        Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                '        ASCMAIN1.sql = "Select Count (*) HITS from SOTORDR2,SOTORDR1" & vbCrLf _
                '            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                '            & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                '            & "   and SOTORDR2.RSRV_NO is Not Null"
                '        Dim HITS As Int64 = Val(ASCDATA1.GetDataValue)
                '        If HITS <> 0 Then
                '            EMsg &= vbCr & "Reservation Restoration is not supported by this Mode"
                '            ' note to future ABS developer - if VAN does not care that reservations will be restored, 
                '            ' there really is no harm in allowing them to continue
                '            ' in truth - it would not be so difficult to restore the reservation either
                '            Exit For
                '        End If
                '    Next
                'End If


                If eItemKey = "Edit" Then
                    'Need to check that we're de-releasing all opne picks for an order else skip order
                    If EMsg = "" Then
                        If optStatus.Value = "REL" Then
                            ASCMAIN1.sql = "SELECT ORDR_NO, SUM (PICKS) PICKS, SUM (PICKS_DEREL) PICKS_DEREL FROM (" & vbCrLf _
                                & " select ordr_no, count (*) picks, 0 PICKS_DEREL from sotpick1 where ordr_no in" & vbCrLf _
                                & " (select distinct ordr_no from " & SOTCANCY & " ) and pick_status = 'P' GROUP BY ORDR_NO" & vbCrLf _
                                & " UNION " & vbCrLf _
                                & " select ordr_no, 0 PICKS, count (*) picks_derel from " & SOTCANCY & "  GROUP BY ORDR_NO)" & vbCrLf _
                                & " GROUP BY ORDR_NO" & vbCrLf _
                                & " HAVING SUM (PICKS) <> SUM (PICKS_DEREL)"
                            For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select()
                                EMsg += "," & row.Item("ORDR_NO")
                            Next
                            If EMsg <> "" Then
                                EMsg = "The following Orders have multiple open Picks:" & vbCrLf & EMsg.Substring(1)
                            End If
                        End If
                    End If
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        For Each row As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        Next
                        If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub
                    End If
                End If

            Case "Cancel"
                    If EntryMode = "V" Then
                    Else
                        If MsgBox("Are you sure you want to Cancel your changes?",
                                MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

            Case "Update"
                    'We're cancelling orders, not style - everything goes.
                    'If dst.Tables("SOTCANCY").Select("ORDR_QTY_OPEN <> ORIG_QTY_OPEN").Length = 0 Then
                    '    EMsg &= vbCr & "No records have been updated"
                    'End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View", "Edit"
                MyBase.EntryMode = Mid(eItemKey, 1, 1)
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Cancel"
                Me.Mode_Settings(False)

            Case "Print"
                Me.Print_Record()

            Case "Update"
                Me.Update_Record()
                Me.Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("View").Visible = False
                    If Not ScreenMode Or (EntryMode = "V") Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode

                    '.Items("Cancel").Visible = (EntryMode = "N" Or EntryMode = "E")
                    .Items("Update").Visible = (EntryMode = "N" Or EntryMode = "E")
                    .Items("Update").Text = "De-release and Cancel Orders"
                    If optStatus.Value = "OPEN" Then
                        .Items("Update").Text = "Cancel Open Orders"
                    End If
                    .Items("Cancel").Text = "Exit Update"

                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Print").Visible = False
                End With

            End With
        End If

        If ASCMAIN1.Running_in_VS = True Then
            btnTest.Visible = True
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)


        grdSOTPICKX.Visible = Not ScreenMode
 
        grdSOTPICKX.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = ScreenMode

        If ScreenMode Then
            With grdSOTCANCY.DisplayLayout.Override
                If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        End If

        With grdSOTPICKX.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTPICKX", "SOTCANCY", "SOTORDR1", "SOTORDR2", "SOTRSRVX", "SOTRSRV1", "SOTRSRV2", "ICTSTATO"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        MyBase.EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("WHSE_CODE").Text = ""

        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE


        Clear_All_Filters(grdSOTPICKX)
    End Sub

    Private Sub Load_Record()

        MyBase.EnforceConstraints(False)

        ASCDATA1.DeleteRows(dst.Tables("SOTPICKX"), "SELECTED <> '1'")

        Create_Temp_Tables(False)

        Fill_Records("SOTCANCY")

        MyBase.EnforceConstraints(True)

        Update_Totals()

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Sub Print_Record()
        Create_Report()
    End Sub

    Function Create_Report() As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Dim REPORT_NAME As String = "SORALLO1"
        Dim RPT As String = REPORT_NAME

        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        dst.Tables("SOTALLOZ").Rows.Clear()

        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLOZ").Rows.Clear()

        Dim STYLE_CODEs As New List(Of String)
        Dim CUST_CODEs As New List(Of String)

        For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO1")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() _
                    {"ALLO_CTL_NO", "STYLE_CODE", "DATE_START", "DATE_END", "INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE", "ALLOW_OVER",
                     "ITEM_DESC", "COLLECTION_CODE", "BRAND_CODE", "ITEM_BASIC_PROMO", "ITEM_SNU_CODE", "QTY_ALLO_PLAN", "QTY_ALLO_TOTAL", "ITEM_DATE_TO_SHIP"}
                    If COLUMN_NAME = "BRAND_CODE" Then
                    Else
                        rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                    End If
                Next
                .Rows.Add(rowR)
            End With

            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")

            Fill_Records("SOTALLOZ", ALLO_CTL_NO, False)


            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            If Not STYLE_CODEs.Contains(STYLE_CODE) Then
                Fill_Records("ICTSTAT2", STYLE_CODE, False)

                Dim rowR As DataRow = REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").NewRow
                For Each DC As DataColumn In dst.Tables("ICTITEM1").Columns
                    Dim COLUMN_NAME As String = DC.ColumnName
                    If REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").Columns.Contains(COLUMN_NAME) Then
                        '      rowR.Item(COLUMN_NAME) = rowICTITEM1.Item(COLUMN_NAME)
                    End If
                Next
                REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").Rows.Add(rowR)

                Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
                Dim imgba() As Byte = Nothing
                Dim IMAGE_FILENAME As String = FOLDER_NAME & "\" & STYLE_CODE & ".JPG"
                If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                    rowR.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                Else
                    IMAGE_FILENAME = FOLDER_NAME & "\" & STYLE_CODE & ".PNG"
                    If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                        rowR.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                    End If
                End If
            End If
        Next

        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()

            Dim SUBT As String = "Allocations by Item/Customer (Screen Report)"
            .CR_params.Add("SUBT", SUBT) ' "")
            .CR_params.Add("PAGE_EJECT", "0")
            .CR_params.Add("EXC_ONLY", "0")
            .CR_params.Add("SUMMARY", "0")
            .Generate_Report(RPT, Me.Text, SUBT)
            .Print_Report_End()

        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return ""
    End Function

    Private Sub Update_Record()

        Dim ORDR_L As String = ""
        Dim ORDR_NO As String = ""
        Dim ORDR_STATUS As String = ""
        Dim OPEN As Int64 = 0

        Dim rowSOTORDR1 As DataRow
        Dim rowSOTORDR2 As DataRow
        Dim ORDR_GROUP_NOs As New List(Of String)
        Dim ORDR_GROUP_NO As String = ""

        Dim ORDR_NOs As New List(Of String)

        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR2").Rows.Clear()

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO in (select distinct ORDR_NO from " & SOTCANCY & ")"
        Fill_Records("SOTORDR1", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO in (select distinct ORDR_NO from " & SOTCANCY & ")"
        Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "select * from SOTPICK1 where PICK_NO in (select PICK_NO from " & SOTCANCY & ")"
        Dim SOTPICK1 As String = ASCMAIN1.Temp_Table
        
        Try
            MyBase.BeginTrans()

            If optStatus.Value = "REL" Then
                TAC.SOCMAIN1.DeRelease(SOTPICK1)
            End If

            Dim Ordr2Select As String = ""
            ASCMAIN1.Progress("Now Retracting Commitments")
            For Each row As DataRow In dst.Tables("SOTCANCY").Select("", "ORDR_NO, CUST_STORE_NO")
                ORDR_NO = row.Item("ORDR_NO")
                ASCMAIN1.Progress("-", ORDR_NO)

                If ORDR_NO <> ORDR_L Then
                    ORDR_NOs.Add(ORDR_NO)
                    rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                    ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO")
                    If ORDR_GROUP_NOs.IndexOf(ORDR_GROUP_NO) = -1 Then
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                    End If
                    Dependent_Updates(-1, ORDR_NO, ORDR_GROUP_NO)
                End If
                Ordr2Select = String.Format("ORDR_NO = '{0}'", ORDR_NO)
                Dim ROWS() As DataRow = dst.Tables("SOTORDR2").Select(Ordr2Select)
                Dim Status As String = ""
                For Each rowSOTORDR2 In ROWS
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = IIf(IsDBNull(rowSOTORDR2.Item("ORDR_QTY_CANC")), 0, rowSOTORDR2.Item("ORDR_QTY_CANC")) + rowSOTORDR2.Item("ORDR_QTY_OPEN") + IIf(IsDBNull(rowSOTORDR2.Item("ORDR_QTY_PICK")), 0, rowSOTORDR2.Item("ORDR_QTY_PICK"))
                    rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                    If rowSOTORDR2.Item("ORDR_QTY_OPEN") <> 0 Then
                        Status = "O"
                    ElseIf rowSOTORDR2.Item("ORDR_QTY_PICK") <> 0 Then
                        Status = "P"
                    ElseIf Not IsDBNull(rowSOTORDR2.Item("ORDR_QTY_SHIP")) AndAlso rowSOTORDR2.Item("ORDR_QTY_SHIP") <> 0 Then
                        Status = "F"
                    Else
                        Status = "C"
                    End If
                    rowSOTORDR2.Item("ORDR_STATUS") = Status
                    ORDR_L = ORDR_NO
                Next
            Next

            ASCMAIN1.sql = "Select SOTRSRV1.* from SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.RSRV_STATUS = 'O' and SOTRSRV1.CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SOTRSRV1", "", True, ASCMAIN1.sql)
            ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV1,SOTRSRV2" & vbCrLf _
                & " where SOTRSRV1.RSRV_STATUS = 'O' and SOTRSRV1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO"
            Fill_Records("SOTRSRV2", "", True, ASCMAIN1.sql)

            ASCMAIN1.Progress("Now Updating Sales Order Commitments")
            For Each ORDR_NO In ORDR_NOs
                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

                'ORDR_NO = rowSOTORDR1.Item("ORDR_NO")
                ASCMAIN1.Progress("-", ORDR_NO)
                ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO")
                ORDR_STATUS = ""
                OPEN = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", String.Format("ORDR_NO = '{0}'", ORDR_NO) & ""))

                If OPEN <> 0 Then
                    ORDR_STATUS = "O"
                Else
                    ORDR_STATUS = "C"
                End If
                rowSOTORDR1.Item("ORDR_STATUS") = ORDR_STATUS
                Dependent_Updates(1, ORDR_NO, ORDR_GROUP_NO)

                'Note RGI has one order per group, this will need to change for other clients if you chose to use
                If rowSOTORDR1.Item("ORDR_SOURCE") = "E" Then
                    ASCMAIN1.sql = "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '855' AND CUST_CODE = '" & CUST_CODE & "'"
                    Dim row As DataRow = ASCDATA1.GetDataRow
                    If row IsNot Nothing Then
                        TAC.EDC855O1.Generate_855(clsASCBASE1, ORDR_GROUP_NO)
                    End If
                End If
                TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "ORDCXL", "E-comm Order De-released and Cancelled")
            Next

            ASCMAIN1.Progress("Now Saving Orders & Reservations")

            Update_Record_TDA("SOTORDR1")

            Update_Record_TDA("SOTORDR2")

            'Update_BAs("SOTORDR2")

            Update_Record_TDA("SOTRSRV1")
            Update_Record_TDA("SOTRSRV2")

            ASCMAIN1.Progress("Now Updating Style/Color Commitments")

            For Each rowICTSTATO As DataRow In dst.Tables("ICTSTATO").Select("")
                Dim STYLE_CODE As String = rowICTSTATO.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowICTSTATO.Item("COLOR_CODE")
                Dim WHSE_CODE As String = rowICTSTATO.Item("WHSE_CODE")
                Dim WHSE_QTY_OPEN As String = Val(rowICTSTATO.Item("WHSE_QTY_OPEN") & "")
                If WHSE_QTY_OPEN <> 0 Then
                    TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", WHSE_QTY_OPEN)
                End If
            Next

            For Each ORDR_GROUP_NO In ORDR_GROUP_NOs
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
            Next

            ASCMAIN1.Progress("")
            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub
    Sub Dependent_Updates(S As Integer, ORDR_NO As String, ORDR_GROUP_NO As String)

        Dim QTY_TO_COMMIT As Int64

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim rowSOTORDR2s() As DataRow

        If S = -1 Then
            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
            rowSOTORDR1 = ASCDATA1.GetDataRow

            ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
            rowSOTORDR2s = ASCDATA1.GetDataTable.Select("")
        Else
            rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            rowSOTORDR2s = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
        End If

        'ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        'For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
        For Each rowSOTORDR2 As DataRow In rowSOTORDR2s
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")
            Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")

            If S = -1 Then
                If rowSOTORDR2.Item("RSRV_NO") & "" <> "" Then
                    'Only restore this reservation line if it hasn't been substitutioned.  Per Gabe 07/30/02 - WR.
                    Dim row As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, rowSOTORDR2.Item("ORDR_LNO")})
                    If row IsNot Nothing Then  'Added for Angela. 1/24/05.  She was adding styles to range that had pulled from reservation already.
                        If row.Item("STYLE_CODE_SUB") & "" = "" Then
                            Update_SOTRSRVx(rowSOTORDR2, S, ORDR_GROUP_NO)
                        End If
                    End If
                End If
            Else

                Dim rowSOTRSRVX As DataRow = Nothing ' Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                '& " order by SOTRSRV1.ORDR_CANCEL_DATE"

                If S = -1 Then
                    rowSOTRSRVX = Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                Else
                    Dim sqlw As String = String.Format("STYLE_CODE = '{0}' and COLOR_CODE = '{1}' and RSRV_QTY_OPEN > 0", STYLE_CODE, COLOR_CODE)
                    Dim rows() As DataRow = dst.Tables("SOTRSRV2").Select(sqlw)
                    If rows.Length <> 0 Then
                        rowSOTRSRVX = rows(0)
                    End If
                End If

                Dim Ps() As Object

                If rowSOTRSRVX IsNot Nothing Then
                    rowSOTORDR2.Item("RSRV_NO") = rowSOTRSRVX.Item("RSRV_NO")
                    rowSOTORDR2.Item("RSRV_LNO") = rowSOTRSRVX.Item("RSRV_LNO")
                    Ps = {rowSOTRSRVX.Item("RSRV_NO"), rowSOTRSRVX.Item("RSRV_LNO")}
                    Update_SOTRSRVx(rowSOTORDR2, S, ORDR_GROUP_NO)
                Else
                    rowSOTORDR2.Item("RSRV_NO") = DBNull.Value
                    rowSOTORDR2.Item("RSRV_LNO") = DBNull.Value
                    Ps = {DBNull.Value, DBNull.Value}
                End If

                'Update_Record_TDA("SOTORDR2")

                'ASCMAIN1.sql = "Update SOTORDR2 Set RSRV_NO = :PARM1, RSRV_LNO = :PARM2" _
                '    & " where ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
                'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VN", Ps)
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE")
                If S = -1 Then
                    TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", S * QTY_TO_COMMIT)
                Else
                    Dim rowICTSTATO As DataRow = dst.Tables("ICTSTATO").Rows.Find(New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE})
                    If rowICTSTATO Is Nothing Then
                        rowICTSTATO = dst.Tables("ICTSTATO").Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, 0})
                    End If
                    rowICTSTATO.Item("WHSE_QTY_OPEN") = Val(rowICTSTATO.Item("WHSE_QTY_OPEN") & "") + QTY_TO_COMMIT
                End If
            End If
        Next

    End Sub

    Sub Update_SOTRSRVx(rowSOTORDR2 As DataRow, S As Integer, ORDR_GROUP_NO As String)
        Dim RSRV_NO As String = rowSOTORDR2.Item("RSRV_NO") & ""
        Dim RSRV_LNO As Int64 = Val(rowSOTORDR2.Item("RSRV_LNO") & "")

        Dim rowSOTRSRV1 As DataRow = Nothing
        Dim rowSOTRSRV2 As DataRow = Nothing
        If S = -1 Then
            rowSOTRSRV1 = Fill_Record("SOTRSRV1", RSRV_NO)
            rowSOTRSRV2 = Fill_Record("SOTRSRV2", New String() {RSRV_NO, RSRV_LNO})
        Else
            rowSOTRSRV1 = dst.Tables("SOTRSRV1").Rows.Find(RSRV_NO)
            rowSOTRSRV2 = dst.Tables("SOTRSRV2").Rows.Find(New Object() {RSRV_NO, RSRV_LNO})
        End If
        Dim WHSE_CODE As String = rowSOTRSRV1.Item("WHSE_CODE")

        With rowSOTRSRV2
            Dim RSRV_QTY As Int64 = .Item("RSRV_QTY")
            Dim RSRV_QTY_OPEN As Int64 = Val(.Item("RSRV_QTY_OPEN") & "")
            Dim RSRV_QTY_CANC As Int64 = Val(.Item("RSRV_QTY_CANC") & "")
            Dim RSRV_QTY_USED As Int64 = Val(.Item("RSRV_QTY_USED") & "") _
                          + S * Val(rowSOTORDR2.Item("ORDR_QTY") & "")

            '  + S * Val(rowSOTORDR2.Item("ORDR_QTY_ORIG") & "") - USING ORDR_QTY_ORIG WILL ALWAYS HAVE 0 IMPACT WHEN CHANGING THE ORDER
            Dim RSRV_QTY_OPEN_OLD As Int64 = RSRV_QTY_OPEN
            RSRV_QTY_OPEN = RSRV_QTY - RSRV_QTY_CANC - RSRV_QTY_USED
            If RSRV_QTY_OPEN < 0 Then
                RSRV_QTY_OPEN = 0
            End If
            Dim RSRV_QTY_OPEN_NEW As Int64 = RSRV_QTY_OPEN
            .Item("RSRV_QTY_USED") = RSRV_QTY_USED
            .Item("RSRV_QTY_OPEN") = RSRV_QTY_OPEN

            Dim QTY_TO_COMMIT As Int64 = RSRV_QTY_OPEN_NEW - RSRV_QTY_OPEN_OLD
            If QTY_TO_COMMIT <> 0 Then
                Dim STYLE_CODE As String = .Item("STYLE_CODE")
                Dim COLOR_CODE As String = .Item("COLOR_CODE")
                If S = -1 Then
                    TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY_TO_COMMIT)
                Else
                    Dim rowICTSTATO As DataRow = dst.Tables("ICTSTATO").Rows.Find(New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE})
                    If rowICTSTATO Is Nothing Then
                        rowICTSTATO = dst.Tables("ICTSTATO").Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, 0})
                    End If
                    rowICTSTATO.Item("WHSE_QTY_OPEN") = Val(rowICTSTATO.Item("WHSE_QTY_OPEN") & "") + QTY_TO_COMMIT
                End If
            End If

        End With

        Dim RSRV_QTY_OPEN_total As Int64 = 0
        If S = -1 Then
            Update_Record_TDA("SOTRSRV2")

            ASCMAIN1.sql = "Select Sum (RSRV_QTY_OPEN) from SOTRSRV2 where RSRV_NO = :PARM1"
            RSRV_QTY_OPEN_total = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {RSRV_NO}))
        Else
            RSRV_QTY_OPEN_total = Val(dst.Tables("SOTRSRV2").Compute("sum(RSRV_QTY_OPEN)", "RSRV_NO = '" & RSRV_NO & "'") & "")
        End If

        If RSRV_QTY_OPEN_total = 0 Then
            rowSOTRSRV1.Item("RSRV_STATUS") = "F"
        Else
            rowSOTRSRV1.Item("RSRV_STATUS") = "O"
        End If

        If S = -1 Then
            Update_Record_TDA("SOTRSRV1")
        End If

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICKX, "SBBBB", "Show Filter", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdSOTCANCY, "SSBBBB", "Show Filter", "Show GroupBox", "Clear Qtys", "Restore Qtys", "Clear Qtys for", "Restore Qtys for", "Sales Order Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
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

            Case "grdSOTPICKX"

            Case "grdSOTCANCY"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else

            Select Case e.SourceControl.Name
                'Case "grdSOTALLOX", "grdICTITEM1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Me.Cursor = Cursors.WaitCursor
        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.Key <> "Show All Levels" Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next
            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                Next
            
                ASCMAIN1.Progress("")

        End Select
        Me.Cursor = Cursors.Default

        If grd Is Nothing OrElse (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow) Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If



        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode Then
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        Set_SOTPICKX()
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)

            Case "CUST_CODE"
                Set_SOTPICKX()

        End Select
    End Sub
#End Region


#Region "grdSOTPICKX"
    Private Sub grdSOTPICKX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICKX.InitializeRow
        With e.Row.Cells("SELECTED")
            If .Value & "" = "1" Then
                .Appearance.BackColor = System.Drawing.Color.LightGreen
            Else
                .Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub
#End Region


    Sub Set_SOTPICKX()

        If SOTCANCY = "" Then Exit Sub
        If ScreenMode Then Exit Sub

        '    Create_Temp_Tables(False)

        If optStatus.Value = "OPEN" Then
            ASCMAIN1.sql = "select SOTORDR1.ORDR_NO PICK_NO, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, '0' SHIP_BOL_NO, SOTORDR1.ORDR_NO, " & vbCrLf _
                & " SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_DATE, '' PICK_RELEASED, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & " from SOTORDR1 " & vbCrLf _
                & " where SOTORDR1.CUST_CODE = :PARM1 AND SOTORDR1.WHSE_CODE = :PARM2 " & vbCrLf _
                & "and SOTORDR1.ORDR_STATUS in ('O') " & vbCrLf _
                & "and SOTORDR1.ORDR_TYPE_CODE = 'B2C' " & vbCrLf _
                & "and SOTORDR1.ORDR_SOURCE = 'E' " & vbCrLf
        Else
            ASCMAIN1.sql = "select SOTPICK1.PICK_NO, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, SOTPICK1.SHIP_BOL_NO, SOTORDR1.ORDR_NO, " & vbCrLf _
                            & " SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_DATE, SOTPICK1.PICK_RELEASED, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                            & " from SOTPICK1, SOTORDR1 " & vbCrLf _
                            & " where SOTORDR1.CUST_CODE = :PARM1 AND SOTORDR1.WHSE_CODE = :PARM2 " & vbCrLf _
                            & "and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO " & vbCrLf _
                            & "and SOTORDR1.ORDR_STATUS in ('P','O') " & vbCrLf _
                            & "and SOTORDR1.ORDR_TYPE_CODE = 'B2C' " & vbCrLf _
                            & "and SOTORDR1.ORDR_SOURCE = 'E' " & vbCrLf _
                            & "and SOTPICK1.PICK_STATUS = 'P' " & vbCrLf
        End If

        Fill_Records("SOTPICKX", New Object() {Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("WHSE_CODE").Text},True,ASCMAIN1.sql)
        For Each row As DataRow In dst.Tables("SOTPICKX").Select("") '("ORDR_QTY <> (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_CANC)")
            row("SELECTED") = "0"
        Next

        grdSOTPICKX.Text = "Pick Tickets in Pick for " & Absx1.txtFor("CUST_CODE").Text & " in Whse " & Absx1.txtFor("WHSE_CODE").Text
    End Sub

    Sub Update_Totals()
        'For ictr As Integer = 1 To iColumn
        '    If ALLO_CTL_NOi(ictr) <> "" Then
        '        Dim QTY_ALLO As Int64 = Val(dst.Tables("SOTALLOC").Compute("SUM(ALLO_" & Format(ictr, "00") & ")", "") & "")
        '        Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NOi(ictr))
        '        rowSOTALLO1.Item("QTY_ALLO_TOTAL") = QTY_ALLO
        '    End If
        'Next

    End Sub

    Sub Create_Temp_Tables(initialize As Boolean)

        If initialize Then
            '  dteORDR_DATE.Value = Now.Date.AddDays(-1)
        End If

        Dim CUST_CODE As String = ""
        Dim WHSE_CODE As String = ""


        sqlORDR_GROUP_NOs = ",''"
        sqlPICK_NOs = ",''"
        O_ORDR_GROUP_NOs.Clear()

        If Not initialize Then
            'CUST_CODE = Absx1.txtFor("CUST_CODE").Text
            'WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
            sqlORDR_GROUP_NOs = ""
            For Each row As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                sqlORDR_GROUP_NOs &= ",'" & row("ORDR_GROUP_NO") & "'"
                sqlPICK_NOs &= ",'" & row("PICK_NO") & "'"
                O_ORDR_GROUP_NOs.Add(row("ORDR_GROUP_NO"))
            Next
        End If

        If optStatus.Value = "OPEN" Then
            ASCMAIN1.sql = "" _
            & "Select SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, '0' PICK_NO, '0' SHIP_BOL_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_STATUS" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_OPEN) ORIG_QTY_OPEN" & vbCrLf _
            & " from SOTORDR1,SOTORDR2" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   And SOTORDR1.ORDR_GROUP_NO in (" & vbCrLf _
            & Mid(sqlORDR_GROUP_NOs, 2) & ")" & vbCrLf _
            & " Group By SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_STATUS"
        Else
            ASCMAIN1.sql = "" _
            & "Select SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTPICK1.PICK_NO, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_STATUS" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM(SOTORDR2.ORDR_QTY_OPEN) ORIG_QTY_OPEN" & vbCrLf _
            & " from SOTORDR1,SOTORDR2,SOTPICK1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   And SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "   And SOTPICK1.PICK_NO in (" & vbCrLf _
            & Mid(sqlPICK_NOs, 2) & ")" & vbCrLf _
            & " Group By SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTPICK1.PICK_NO, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_STATUS"
        End If
        

        If initialize Then
            SOTCANCY = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Delete from " & SOTCANCY)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCANCY & " " & ASCMAIN1.sql)
        End If

    End Sub

  

    Private Sub grdSOTCANCY_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdSOTCANCY.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ORDR_QTY_OPEN"
                Dim ORDR_QTY As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY").Value)
                Dim ORDR_QTY_OPEN As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_OPEN").Value)
                Dim ORDR_QTY_PICK As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_PICK").Value)
                Dim ORDR_QTY_SHIP As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_SHIP").Value)
                Dim ORDR_QTY_CANC As Int64 = ORDR_QTY - (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP)
                e.Cell.Row.Cells("ORDR_QTY_CANC").Value = ORDR_QTY_CANC

            Case "CANCEL"
                If e.Cell.Row.Cells("CANCEL").Value = "1" Then
                    e.Cell.Row.Cells("ORDR_QTY_OPEN").Value = 0
                Else
                    e.Cell.Row.Cells("ORDR_QTY_OPEN").Value = e.Cell.Row.Cells("ORDR_QTY_CANC").Value
                End If

        End Select
    End Sub

    Private Sub grdSOTCANCY_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTCANCY.AfterRowActivate
        If Not grdSOTCANCY.ActiveRow.IsDataRow Then
            Exit Sub
        End If
        If grdSOTCANCY.ActiveRow.Cells("ORDR_STATUS").Value = "O" Then
            grdSOTCANCY.ActiveRow.Cells("ORDR_QTY_OPEN").Column.CellActivation = Activation.AllowEdit
        Else
            grdSOTCANCY.ActiveRow.Cells("ORDR_QTY_OPEN").Column.CellActivation = Activation.NoEdit
        End If
    End Sub

    Private Sub grdSOTCANCY_BeforeExitEditMode(sender As Object, e As BeforeExitEditModeEventArgs) Handles grdSOTCANCY.BeforeExitEditMode
        'If grdSOTCANCY.ActiveCell.Column.Key = "ORDR_QTY_OPEN" Then
        '    Dim ORDR_QTY As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY").Value & "")
        '    Dim ORDR_QTY_OPEN As Int64 = Val(grdSOTCANCY.ActiveCell.Text & "")

        '    Dim ORDR_QTY_PICK As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_PICK").Value & "")
        '    Dim ORDR_QTY_SHIP As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_SHIP").Value & "")
        '    'Dim ORDR_QTY_CANC As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_CANC").Value & "")

        '    If ORDR_QTY < (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP) Then
        '        e.Cancel = True
        '    End If
        'End If
    End Sub

    Private Sub grdSOTCANCY_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdSOTCANCY.BeforeCellUpdate
        If e.Cell.Column.Key = "ORDR_QTY_OPEN" And Not grdSOTCANCY.ActiveCell Is Nothing Then
            Dim ORDR_QTY As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY").Value & "")
            Dim ORDR_QTY_OPEN As Int64 = e.NewValue 'Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_OPEN").Value & "") ' Val(grdSOTCANCY.ActiveCell.Text & "")
            'Dim STYLE_CODE As String = grdSOTCANCY.ActiveCell.Row.Cells("STYLE_CODE").Value
            'Dim COLOR_CODE As String = grdSOTCANCY.ActiveCell.Row.Cells("COLOR_CODE").Value
            Dim ORDR_QTY_PICK As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_PICK").Value & "")
            Dim ORDR_QTY_SHIP As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_SHIP").Value & "")
            'Dim ORDR_QTY_CANC As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_CANC").Value & "")


            If ORDR_QTY < (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP) Or ORDR_QTY_OPEN < 0 Then
                MsgBox("Invalid Qty")
                e.Cancel = True
            End If

        End If
    End Sub

    Private Sub grdSOTPICKX_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdSOTPICKX.BeforeCellUpdate
        If e.Cell.Column.Key = "SELECTED" And Not grdSOTPICKX.ActiveCell Is Nothing Then
            If ASCMAIN1.Running_in_VS Then Exit Sub
            'If Val(grdSOTPICKX.ActiveCell.Row.Cells("ORDR_QTY").Value & "") <>
            '    Math.Abs(Val(grdSOTPICKX.ActiveCell.Row.Cells("ORDR_QTY_OPEN").Value & "") _
            '             + Val(grdSOTPICKX.ActiveCell.Row.Cells("ORDR_QTY_PICK").Value & "") _
            '             + Val(grdSOTPICKX.ActiveCell.Row.Cells("ORDR_QTY_CANC").Value & "")) Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Private Sub txtCUST_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtCUST_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If SOTCANCY = "" Then Exit Sub

        Set_SOTPICKX()
    End Sub

    Private Sub txtWHSE_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtWHSE_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If SOTCANCY = "" Then Exit Sub

        Set_SOTPICKX()
    End Sub

    Private Sub optStatus_ValueChanged(sender As Object, e As EventArgs) Handles optStatus.ValueChanged
        Set_SOTPICKX()

    End Sub

    Private Sub Test_Cancel() Handles btnTest.Click

        Dim PO_list As String() = {"30886774", "30873770", "30885723", "30873829", "30876203" _
, "30877345", "30878180", "30878300", "30879879", "30879656", "30875759", "30876292" _
, "30868703", "30879389", "30873562", "30877130", "30879870", "30878338", "30880997" _
, "30868708", "30879552", "30878219", "30878289", "30875686", "30880045"}

        For Each PO As String In PO_list
            Dim row As DataRow = ASCDATA1.GetDataRow("SELECT * From SOTORDR1 where ORDR_CUST_PO = '" & PO & "'")
            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
            If row IsNot Nothing Then
                TAC.EDC855O1.Generate_855(clsASCBASE1, ORDR_GROUP_NO)
            Else
                Stop

            End If

        Next


    End Sub
End Class