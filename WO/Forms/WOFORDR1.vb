Imports System.Xml
Imports System.Drawing
Imports System.Text

Public Class WOFORDR1
    Dim WKORDER_NO As String
    Dim WHSE_CODE As String
    Dim rowWOTORDR1 As DataRow
    Dim rowWOTORDR5 As DataRow
    Dim STYLE_CODE As String
    Dim COLORS As New List(Of String)
    Dim TransmitToWhse As Boolean = False
    Dim TotalUnits As Integer
    'Dim ToolTips As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "SELECT WOTORDR1.* from WOTORDR1"
            Create_TDA(.Tables.Add, "WOTORDRX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "WOTORDR1", "*")

            ASCMAIN1.sql = "Select WOTORDR3.*" & vbCrLf _
            & " from WOTORDR3" & vbCrLf _
            & " where WOTORDR3.WKORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "WOTORDR3", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select WOTORDR4.*" & vbCrLf _
            & " from WOTORDR4" & vbCrLf _
            & " where WOTORDR4.WKORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "WOTORDR4", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select WOTORDR5.*" & vbCrLf _
            & " from WOTORDR5" & vbCrLf _
            & " where WOTORDR5.WKORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "WOTORDR5", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "SELECT WOTSIZE1.* from WOTSIZE1"
            Create_TDA(.Tables.Add, "WOTSIZE1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "SELECT WHTPPKM1.* from WHTPPKM1"
            Create_TDA(.Tables.Add, "WHTPPKM1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "SELECT WHTPPKM2.* from WHTPPKM2"
            Create_TDA(.Tables.Add, "WHTPPKM2", "**", 0, False, "", 3)

            ASCMAIN1.sql = "SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_QTY) ORDR_QTY" & vbCrLf _
                & " FROM SOTORDR2" & vbCrLf _
                & " WHERE ROWNUM < 0" & vbCrLf _
                & " GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
                & " ORDER BY STYLE_CODE, COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)

        End With

        grdWOTORDRX.DataSource = dst.Tables("WOTORDRX")
        grdWOTORDR3.DataSource = dst.Tables("WOTORDR3")
        grdWOTORDR4.DataSource = dst.Tables("WOTORDR4")

        Bind_Controls(grpWOTORDR1, "WOTORDR1")

        ASCMAIN1.Add_Value_List(grdWOTORDRX, "WKORDER_STATUS", Nothing, New String() {":", "P:Pending", "0:Transmitted", "1:Transmitted", "2:Transmitted", "3:Transmitted", "C:Completed", "X:Deleted", "4:Recall Requested", "5:Recall Accepted", "6:Recall Denied", "7:Recalled"})

        Check_InquiryMode()

    End Sub

    Sub Check_InquiryMode()
        If Not IsNothing(optStatus.Value) And optStatus.Value <> "P" Then
            InquiryMode = True
        Else
            InquiryMode = (MENU_ITEM_OBJECT = "WOFORDRI")
        End If

        With UltraExplorerBar1.Groups("Screen Control")
            If EntryMode = "V" Then
                InquiryMode = True
                If optStatus.Value = "P" Then
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    .Items("Edit").Visible = True
                Else
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    .Items("Edit").Visible = False
                End If
                If "012T".Contains(optStatus.Value.ToString) Then
                    .Items("Recall").Settings.Enabled = DefaultableBoolean.True
                    .Items("Recall").Visible = True
                Else
                    .Items("Recall").Settings.Enabled = DefaultableBoolean.False
                    .Items("Recall").Visible = False
                End If
            End If
            .Items("New").Visible = Not InquiryMode
            '.Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
            .Items("Delete").Visible = Not InquiryMode
            .Items("Transmit").Visible = Not InquiryMode
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdWOTORDR3, grdWOTORDR4}
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With
        Next

        Set_Read_Only(grpWOTORDR1, InquiryMode)
        Set_Read_Only(Panel1, InquiryMode)
        Set_Read_Only(Panel2, InquiryMode)
        Set_Read_Only(Panel3, InquiryMode)
        Set_Read_Only(Panel4, InquiryMode)
        Set_Read_Only(Panel6, InquiryMode)
        Set_Read_Only(Panel7, InquiryMode)
        Set_Read_Only(splPOTSHIP1, InquiryMode)
        btnAddStyles.Enabled = Not InquiryMode
        btnClear.Enabled = Not InquiryMode
        btnStyleMatch.Enabled = Not InquiryMode
        cboAPPLY_TKT.Enabled = Not InquiryMode
        cboAPPLY_HGR.Enabled = Not InquiryMode
        cboAPPLY_STK.Enabled = Not InquiryMode
        cboAPPLY_TAG.Enabled = Not InquiryMode
        cboREMOVE_TKT.Enabled = Not InquiryMode
        cboREMOVE_HGR.Enabled = Not InquiryMode
        cboREMOVE_STK.Enabled = Not InquiryMode
        cboREMOVE_TAG.Enabled = Not InquiryMode

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdWOTORDR3, grdWOTORDR4}
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With
        Next

        btnStyleMatch.Enabled = Not InquiryMode
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("WKORDER_DESC").Text = "" Then
                    EMsg &= vbCr & "You Must Enter A Valid Description."
                End If
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Select A Warehouse."
                End If
            Case "View", "Edit"

            Case "Update"
                If Not AreAllStylesMatched() Then
                    EMsg &= vbCr & "Styles Created Records Found Without References To Styles Used."
                End If
                If Val(cboAPPLY_HGR.Text) > 0 Or Val(cboAPPLY_STK.Text) > 0 Or Val(cboAPPLY_TKT.Text) > 0 Then
                    If IsDBNull(Absx1.numFor("APPLY_UNITS").Value) Then
                        EMsg &= vbCr & "You Must Supply Units To Apply When Applying Tickets, Hangers or Stickers."
                    Else
                        If Val(Absx1.numFor("APPLY_UNITS").Value) = 0 Then
                            EMsg &= vbCr & "You Must Supply Units To Apply When Applying Tickets, Hangers or Stickers."
                        End If
                    End If
                End If
                If Val(cboREMOVE_TKT.Text) > 0 Or Val(cboREMOVE_HGR.Text) > 0 Then
                    If IsDBNull(Absx1.numFor("REMOVE_UNITS").Value) Then
                        EMsg &= vbCr & "You Must Supply Units To Remove When Removing Tickets, Hangers or Stickers."
                    Else
                        If Val(Absx1.numFor("REMOVE_UNITS").Value) = 0 Then
                            EMsg &= vbCr & "You Must Supply Units To Remove When Removing Tickets, Hangers or Stickers."
                        End If
                    End If
                End If
                If TBLs.Item("WOTORDR3").Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must Supply Styles To Be Worked On In The Styles Used Grid."
                End If
            Case "Cancel", "Done"
                'EMsg &= vbCr & "This Feature Not Enabled Yet!"
            Case "Delete"
                If optWKORDER_STATUS.Value <> "P" Then
                    EMsg &= vbCr & "Only Orders In A Pending Status May Be Deleted!"
                Else
                    Dim iResult As MsgBoxResult = MessageBox.Show("Are You Sure You Want To Delete This Work Order?", "This Can Not Be Reversed", MessageBoxButtons.YesNo)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg &= vbCr & "Transmission Cancelled At Your Request"
                    End If
                End If
            Case "Print"
                'EMsg &= vbCr & "This Feature Not Enabled Yet!"
            Case "Transmit"
                If ScreenMode Then
                    If optWKORDER_STATUS.Value <> "P" Then
                        EMsg &= vbCr & "Only Orders In A Pending Status May Be Transmitted!"
                    Else
                        Dim iResult As MsgBoxResult = MessageBox.Show("Are You Sure You Want To Transmit This Work Order?", "This Can Not Be Reversed", MessageBoxButtons.YesNo)
                        If iResult <> MsgBoxResult.Yes Then
                            EMsg &= vbCr & "Transmission Cancelled At Your Request"
                        End If
                    End If
                Else
                    EMsg &= vbCr & "Please Select and Enter The Order You Wish To Transmit First"
                End If
            Case "Recall"
                EMsg &= vbCr & "This Feature Is Under Testing And Is Not Yet Available."
            Case "Update From ADS"
                EMsg &= vbCr & "This Feature Has To Be Tested By Wayne Before Use. Please Let Him Know There Is Data Ready."
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
                Check_InquiryMode()
                Mode_Settings(True)
            Case "View"
                EntryMode = "V"
                Load_Record()
                Check_InquiryMode()
                Mode_Settings(True)
            Case "Update"
                Update_Record(True)
                Mode_Settings(False)
            Case "Cancel"
                Mode_Settings(False)
            Case "Done"
                Mode_Settings(False)
            Case "Delete"
                MarkAsDeleted()
                Update_Record(False)
                Mode_Settings(False)
            Case "Print"
                Print_Record()
            Case "Transmit"
                TransmitToWhse = True
                Update_Record(False)
                TransmitToWhse = False
                Mode_Settings(False)
            Case "Update From ADS"
                Stop
                SyncRecordsWithADS("2")
                SyncRecordsWithADS("3")
                Mode_Settings(False)
            Case "Recall"
                Stop
                SyncRecordsWithADS("2")
                SyncRecordsWithADS("3")
                RecallWO()
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
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode
                    If (EntryMode = "E" And ScreenMode) Then
                        .Items("Delete").Settings.Enabled = DefaultableBoolean.True
                        .Items("Transmit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                        .Items("Transmit").Settings.Enabled = DefaultableBoolean.False
                    End If
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("View").Visible = InquiryMode Or (EntryMode = "V" Or Not ScreenMode)
                    .Items("Done").Visible = InquiryMode Or (EntryMode = "V" And ScreenMode)
                    '.Items("Transmit").Visible = InquiryMode Or (EntryMode = "V" And ScreenMode)

                    '.Items("New").Visible = Not InquiryMode And Not ScreenMode
                    '.Items("Update").Visible = ((EntryMode = "N" Or EntryMode = "E") And ScreenMode)
                    '.Items("Cancel").Visible = ((EntryMode = "N" Or EntryMode = "E") And ScreenMode)
                    '.Items("Delete").Visible = ((EntryMode = "E") And ScreenMode)
                    '.Items("Cancel PO").Visible = ((EntryMode = "E") And ScreenMode)

                    .Items("Print").Visible = InquiryMode Or (EntryMode = "V" And ScreenMode)
                End With

                '.Groups("Totals").Visible = ScreenMode
                .Groups("Status Filter").Visible = Not ScreenMode 'And InquiryMode

                If Not ScreenMode Then
                    optStatus.Value = "P"
                    '.Groups("Line Item Commands").Visible = False
                    '.Groups("Cost Calculation").Visible = False
                End If

            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'setup_tabPOTORDR1()

        ' splPOTORDR1.Visible = tf
        grdWOTORDRX.Visible = Not tf

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdWOTORDR3, grdWOTORDR4}
                With grd.DisplayLayout.Override
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowUpdate = DefaultableBoolean.True
                        .AllowDelete = DefaultableBoolean.True
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    End If
                End With
            Next
        Else
            Clear_Record()
        End If

        optWKORDER_STATUS.Enabled = False
        dteDATE_PULLED.Enabled = False
        dteDATE_COMPLETED.Enabled = False
        txtUNITS_WORKED.Enabled = False
        txtEXT_JOB_NO.Enabled = False
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"WOTORDR1", "WOTORDR3", "WOTORDR4", "WOTORDR5", "WOTSIZE1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_WOTORDRX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            WKORDER_NO = ASCMAIN1.Next_Control_No("WKORDER_NO")
        Else
            WKORDER_NO = Absx1.txtFor("WKORDER_NO").Text
        End If
        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Value

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowWOTORDR1 = dst.Tables("WOTORDR1").NewRow
            rowWOTORDR1.Item("WKORDER_NO") = WKORDER_NO
            rowWOTORDR1.Item("WHSE_CODE") = WHSE_CODE
            rowWOTORDR1.Item("WKORDER_DESC") = Absx1.txtFor("WKORDER_DESC").Text
            rowWOTORDR1.Item("WKORDER_STATUS") = "P"
            rowWOTORDR1.Item("APPLY_TKT") = 0
            rowWOTORDR1.Item("APPLY_HGR") = 0
            rowWOTORDR1.Item("APPLY_STK") = 0
            rowWOTORDR1.Item("APPLY_TAG") = 0
            rowWOTORDR1.Item("REMOVE_TKT") = 0
            rowWOTORDR1.Item("REMOVE_HGR") = 0
            rowWOTORDR1.Item("REMOVE_STK") = 0
            rowWOTORDR1.Item("REMOVE_TAG") = 0
            dst.Tables("WOTORDR1").Rows.Add(rowWOTORDR1)

            rowWOTORDR5 = dst.Tables("WOTORDR5").NewRow
            rowWOTORDR5.Item("WKORDER_NO") = WKORDER_NO
            dst.Tables("WOTORDR5").Rows.Add(rowWOTORDR5)
        Else
            rowWOTORDR1 = Fill_Record("WOTORDR1", WKORDER_NO)
            rowWOTORDR5 = Fill_Record("WOTORDR5", WKORDER_NO)
        End If
        Fill_Records("WOTORDR3", WKORDER_NO)
        Fill_Records("WOTORDR4", WKORDER_NO)
        Fill_Records("WOTSIZE1")
        'Fill_Records("WOTORDR5", WKORDER_NO)
        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        If Not InquiryMode Then
            With grdWOTORDR3.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                Next
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("FROM_QTY_REQ").CellActivation = UltraWinGrid.Activation.AllowEdit
            End With
        End If
        SetWKORDER_STATUS()
        'Sort_grdColumns(grdPOTORDR3, "COLOR_NO")
    End Sub

    Sub Cancel_Order()
        'This need to be vetted with ADS before we know what to do.
        'Me.Cursor = Cursors.WaitCursor
        'BeginTrans()
        'Dependent_Updates(-1, WKORDER_NO, True)
        'CommitTrans("Order " & WKORDER_NO & " has been Cancelled")
        'Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()
        Dependent_Updates(-1, WKORDER_NO)
        For Each TABLE_NAME In New String() {"WOTORDR1", "WOTORDR3", "WOTORDR4", "WOTORDR5"}
            ASCDATA1.ExecuteSQL(String.Format("Delete from {0} where WKORDER_NO = '{1}'", TABLE_NAME, WKORDER_NO))
        Next
        CommitTrans(String.Format("PO {0} has been Deleted", WKORDER_NO))
        Me.Cursor = Cursors.Default
    End Sub

    Sub Dependent_Updates(S As Integer, WKORDER_NO As String, Optional cancel_po As Boolean = False)

        ' there is the usual fuzziness around Close PO vs Cancel PO 
        ' since we support Delete (which is like Cancel PO for VANs purposes), we might weant to rename Cancel PO to Close PO

        'ASCMAIN1.sql = ""
        'ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record(ShowMsg As Boolean)
        BeginTrans()

        If EntryMode <> "N" Then
            Dependent_Updates(-1, WKORDER_NO)
        End If

        Dim sqlx As String = String.Format("WKORDER_NO = '{0}'", WKORDER_NO)
        INIT_LAST("WOTORDR1", False, sqlx, True)
        INIT_LAST("WOTORDR5", True, sqlx, True)

        If TransmitToWhse Then
            For Each rowWOTORDR1 As DataRow In dst.Tables("WOTORDR1").Select()
                rowWOTORDR1.Item("DATE_SENT") = Now()
                rowWOTORDR1.Item("WKORDER_STATUS") = "0"
            Next
            rowWOTORDR1.Item("WKORDER_STATUS") = "0"
        End If

        Update_Record_TDA("WOTORDR1", sqlx)
        Update_Record_TDA("WOTORDR3", sqlx)
        Update_Record_TDA("WOTORDR4", sqlx)
        Update_Record_TDA("WOTORDR5", sqlx)

        Dependent_Updates(1, WKORDER_NO)

        'Turn this back on when you are ready to send records to ADS.
        If TransmitToWhse Then
            'Need to do this in case we are re-transmiting.   We can only do this if we previously set the records to 0 or 1.
            ASCMAIN1.sql = String.Format("DELETE FROM ADS.WOTORDR1@ADSIIS WHERE WKORDER_NO = '{0}' AND WKORDER_STATUS IN ('0','1')", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = String.Format("DELETE FROM ADS.WOTORDR3@ADSIIS WHERE WKORDER_NO = '{0}'", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = String.Format("DELETE FROM ADS.WOTORDR4@ADSIIS WHERE WKORDER_NO = '{0}'", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = String.Format("DELETE FROM ADS.WOTORDR5@ADSIIS WHERE WKORDER_NO = '{0}'", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = String.Format("INSERT INTO ADS.WOTORDR1@ADSIIS SELECT * FROM WOTORDR1 WHERE WKORDER_NO = '{0}'", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = String.Format("INSERT INTO ADS.WOTORDR3@ADSIIS SELECT * FROM WOTORDR3 WHERE WKORDER_NO = '{0}'", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = String.Format("INSERT INTO ADS.WOTORDR4@ADSIIS SELECT * FROM WOTORDR4 WHERE WKORDER_NO = '{0}'", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = String.Format("INSERT INTO ADS.WOTORDR5@ADSIIS SELECT * FROM WOTORDR5 WHERE WKORDER_NO = '{0}'", WKORDER_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = "DELETE FROM ADS.WOTSIZE1@ADSIIS"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = "INSERT INTO ADS.WOTSIZE1@ADSIIS SELECT * FROM WOTSIZE1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If

        If ShowMsg Then
            CommitTrans("Update Complete")
        Else
            CommitTrans("")
        End If
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("WKORDER_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "WOTORDR1"
            E.COLUMN_NAME = "WKORDER_NO"
            E.CODE_VALUE = "" ' HFs("CUST_CODE")
            E.DESC_VALUE = "" ' HFs("CUST_NAME")
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "WOTORDR1"
        E.TABLE_KEY_CAPTION = "WO"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("WKORDER_NO").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("WKORDER_DESC").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "N")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Dim Where_Prep As New System.Text.StringBuilder
        Select Case COLUMN_NAME
            Case "WKORDER_NO"
                If InquiryMode Then
                    If optStatus.Value = "O" Then
                        'Change the selection some day once we want to filter the selector
                        'sql_where = " AND WKORDER_NO in (Select DISTINCT WKORDER_NO from WOTORDR1 where WKORDER_STATUS in ('P')) "
                    End If
                Else
                    'Change the selection some day once we want to filter the selector
                    'sql_where = " AND WKORDER_NO in (Select DISTINCT WKORDER_NO from WOTORDR1 where WKORDER_STATUS in ('P')) "
                End If
            Case "ORDR_NO"
                If txtCUST_CODE.Text.Length = 0 Then
                    MsgBox("You Must First Select A Customer Before Selecting It's Orders", MsgBoxStyle.OkOnly, "Customer Not Slected")
                    Cancel = True
                Else
                    Where_Prep.AppendLine(String.Format(" AND CUST_CODE = '{0}' AND ORDR_STATUS IN ('O', 'P')", txtCUST_CODE.Text))
                    If txtORDR_GROUP_NO.TextLength > 0 Then
                        Where_Prep.AppendLine(String.Format(" AND ORDR_GROUP_NO = '{0}'", txtORDR_GROUP_NO.Text))
                    End If
                    If txtSHIP_BOL_NO.TextLength > 0 Then
                        Where_Prep.AppendLine(String.Format(" AND ORDR_NO IN (SELECT DISTINCT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO  = '{0}')", txtSHIP_BOL_NO.Text))
                    End If
                End If
                sql_where = Where_Prep.ToString()
            Case "ORDR_GROUP_NO"
                If txtCUST_CODE.Text.Length = 0 Then
                    MsgBox("You Must First Select A Customer Before Selecting It's Groups", MsgBoxStyle.OkOnly, "Customer Not Slected")
                    Cancel = True
                Else
                    'txtORDR_NO
                    Where_Prep.AppendLine(String.Format(" AND ORDR_GROUP_NO IN (SELECT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_STATUS IN ('O', 'P') AND CUST_CODE = '{0}' )", txtCUST_CODE.Text))
                    If txtSHIP_BOL_NO.TextLength > 0 Then
                        Where_Prep.AppendLine(String.Format(" AND ORDR_GROUP_NO IN (SELECT DISTINCT ORDR_GROUP_NO FROM SOTSHIP1 WHERE SHIP_BOL_NO  = '{0}')", txtSHIP_BOL_NO.Text))
                    End If
                    If txtORDR_NO.TextLength > 0 Then
                        Where_Prep.AppendLine(String.Format(" AND ORDR_GROUP_NO IN (SELECT DISTINCT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_NO = '{0}')", txtORDR_NO.Text))
                    End If
                    sql_where = Where_Prep.ToString()
                End If
            Case "SHIP_BOL_NO"
                If txtCUST_CODE.Text.Length = 0 Then
                    MsgBox("You Must First Select A Customer Before Selecting It's Shipments", MsgBoxStyle.OkOnly, "Customer Not Slected")
                    Cancel = True
                Else
                    Where_Prep.AppendLine(String.Format(" AND CUST_CODE = '{0}' AND SHIP_STATUS = 'P'", txtCUST_CODE.Text))
                    If txtORDR_NO.TextLength > 0 Then
                        Where_Prep.AppendLine(String.Format(" AND SOTSHIP1.SHIP_BOL_NO IN (SELECT DISTINCT SHIP_BOL_NO FROM SOTPICK1 WHERE ORDR_NO = '{0}')", txtORDR_NO.Text))
                    End If
                    If txtORDR_GROUP_NO.TextLength > 0 Then
                        Where_Prep.AppendLine(String.Format(" AND SOTSHIP1.ORDR_GROUP_NO = '{0}'", txtORDR_GROUP_NO.Text))
                    End If
                    sql_where = Where_Prep.ToString()
                End If
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWOTORDRX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "WO Inquiry")
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

                Case "grdWOTORDR2"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Update All ETA Dates"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E") _
                    '    And grd.ActiveCell IsNot Nothing AndAlso grd.ActiveCell.Column.Key = "PO_DATE_ETA"

                    'tlb_btn = DirectCast(tlb_pop.Tools("Update All Ship Dates"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E") _
                    '    And grd.ActiveCell IsNot Nothing AndAlso grd.ActiveCell.Column.Key = "PO_DATE_SHIP_BY"

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "WO Inquiry"
                Dim WKORDER_NO As String = grd.ActiveRow.Cells("WKORDER_NO").Text
                Context_Launch("Load", WKORDER_NO, e.Tool.Key, "WOFORDRI", "F", "WOE")
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "WKORDER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select
    End Sub

    Public Overrides Sub num_Leave(sender As Object, e As System.EventArgs)
        'Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        'Dependent_Calculations(COLUMN_NAME)
    End Sub

    Private Sub SetGroupNoData()
        Dim Rec_Cnt As Int16 = 0
        If txtORDR_GROUP_NO.TextLength > 0 Then
            ASCMAIN1.sql = String.Format("SELECT COUNT(DISTINCT ORDR_NO) FROM SOTORDR1 WHERE ORDR_GROUP_NO = '{0}'", txtORDR_GROUP_NO.Text)
            Rec_Cnt = Val(ASCDATA1.GetDataValue)
            If Rec_Cnt = 1 Then
                ASCMAIN1.sql = String.Format("SELECT DISTINCT ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO = '{0}'", txtORDR_GROUP_NO.Text)
                txtORDR_NO.Text = ASCDATA1.GetDataValue
            Else
                txtORDR_NO.Text = ""
            End If
            ASCMAIN1.sql = String.Format("SELECT COUNT(DISTINCT SHIP_BOL_NO) FROM SOTSHIP1 WHERE ORDR_GROUP_NO = '{0}'", txtORDR_GROUP_NO.Text)
            Rec_Cnt = Val(ASCDATA1.GetDataValue)
            If Rec_Cnt = 1 Then
                ASCMAIN1.sql = String.Format("SELECT DISTINCT SHIP_BOL_NO FROM SOTSHIP1 WHERE ORDR_GROUP_NO = '{0}'", txtORDR_GROUP_NO.Text)
                txtSHIP_BOL_NO.Text = ASCDATA1.GetDataValue
            Else
                txtSHIP_BOL_NO.Text = ""
            End If
            ASCMAIN1.sql = String.Format("SELECT ORDR_CUST_PO FROM SOTORDR0 WHERE ORDR_GROUP_NO = '{0}'", txtORDR_GROUP_NO.Text)
            txtORDR_GROUP_PO.Text = ASCDATA1.GetDataValue

            Dim ORDR_SHIP_DATE As DateTime
            Dim ORDR_CANCEL_DATE As DateTime
            Dim RecFound As Boolean = False
            If txtORDR_GROUP_NO.Text.Length > 0 Then
                Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", txtORDR_GROUP_NO.Text)
                If Not IsNothing(rowSOTORDR0) Then
                    RecFound = True
                    ORDR_SHIP_DATE = Format(rowSOTORDR0.Item("ORDR_SHIP_DATE"), "MM/dd/yy")
                    ORDR_CANCEL_DATE = Format(rowSOTORDR0.Item("ORDR_CANCEL_DATE"), "MM/dd/yy")
                End If
            End If
            If RecFound Then
                txtSHIP_DATE.DateTime = ORDR_SHIP_DATE
                txtCANCEL_DATE.DateTime = ORDR_CANCEL_DATE
                txtDEADLINE_DATE.DateTime = ORDR_CANCEL_DATE
            Else
                txtSHIP_DATE.Value = DBNull.Value
                txtCANCEL_DATE.Value = DBNull.Value
                txtDEADLINE_DATE.Value = DBNull.Value
            End If
        Else
            txtORDR_NO.Text = ""
            txtSHIP_BOL_NO.Text = ""
            txtORDR_GROUP_PO.Text = ""
        End If
    End Sub

    Private Sub SetOrderNoData()
        Dim Rec_Cnt As Int16 = 0
        If txtORDR_NO.TextLength > 0 Then
            If txtORDR_GROUP_NO.Text.Length = 0 Then
                ASCMAIN1.sql = String.Format("SELECT COUNT(DISTINCT ORDR_GROUP_NO) FROM SOTORDR1 WHERE ORDR_NO = '{0}'", txtORDR_NO.Text)
                Rec_Cnt = Val(ASCDATA1.GetDataValue)
                If Rec_Cnt = 1 Then
                    ASCMAIN1.sql = String.Format("SELECT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_NO = '{0}'", txtORDR_NO.Text)
                    txtORDR_GROUP_NO.Text = ASCDATA1.GetDataValue
                Else
                    txtORDR_GROUP_NO.Text = ""
                    txtORDR_GROUP_PO.Text = ""
                End If
            End If

            ASCMAIN1.sql = String.Format("SELECT ORDR_CUST_PO FROM SOTORDR1 WHERE ORDR_NO = '{0}'", txtORDR_NO.Text)
            txtORDR_PO.Text = ASCDATA1.GetDataValue

            If txtSHIP_BOL_NO.Text.Length = 0 Then
                ASCMAIN1.sql = String.Format("SELECT COUNT(DISTINCT SHIP_BOL_NO) FROM SOTPICK1 WHERE ORDR_NO = '{0}'", txtORDR_NO.Text)
                Rec_Cnt = Val(ASCDATA1.GetDataValue)
                If Rec_Cnt = 1 Then
                    ASCMAIN1.sql = String.Format("SELECT DISTINCT SHIP_BOL_NO FROM SOTPICK1 WHERE ORDR_NO = '{0}'", txtORDR_NO.Text)
                    txtSHIP_BOL_NO.Text = ASCDATA1.GetDataValue
                Else
                    txtSHIP_BOL_NO.Text = ""
                End If
            End If

        Else
            txtORDR_PO.Text = ""
        End If
    End Sub

    Private Sub SetBOLNoData()
        Dim Rec_Cnt As Int16 = 0
        If txtSHIP_BOL_NO.TextLength > 0 Then
            If txtORDR_NO.Text.Length = 0 Then
                ASCMAIN1.sql = String.Format("SELECT COUNT(DISTINCT ORDR_NO) FROM SOTPICK1 WHERE SHIP_BOL_NO = '{0}'", txtSHIP_BOL_NO.Text)
                Rec_Cnt = Val(ASCDATA1.GetDataValue)
                If Rec_Cnt = 1 Then
                    ASCMAIN1.sql = String.Format("SELECT DISTINCT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '{0}'", txtSHIP_BOL_NO.Text)
                    txtORDR_NO.Text = ASCDATA1.GetDataValue
                End If
            End If

            If txtORDR_GROUP_NO.Text.Length = 0 Then
                ASCMAIN1.sql = String.Format("SELECT COUNT(DISTINCT ORDR_GROUP_NO) FROM SOTSHIP1 WHERE SHIP_BOL_NO = '{0}'", txtSHIP_BOL_NO.Text)
                Rec_Cnt = Val(ASCDATA1.GetDataValue)
                If Rec_Cnt = 1 Then
                    ASCMAIN1.sql = String.Format("SELECT DISTINCT ORDR_GROUP_NO FROM SOTSHIP1 WHERE SHIP_BOL_NO = '{0}'", txtSHIP_BOL_NO.Text)
                    txtORDR_GROUP_NO.Text = ASCDATA1.GetDataValue
                End If
            End If
        Else
            txtSHIP_DATE.Value = DBNull.Value
            txtCANCEL_DATE.Value = DBNull.Value
            txtDEADLINE_DATE.Value = DBNull.Value
        End If
    End Sub
    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        With Absx1.txtFor(COLUMN_NAME)
            Select Case COLUMN_NAME

                Case "ORDR_GROUP_NO"
                    SetGroupNoData()
                Case "ORDR_NO"
                    SetOrderNoData()
                Case "SHIP_BOL_NO"
                    SetBOLNoData()
                Case Else

            End Select

        End With
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WKORDER_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub CheckedChanged_Special(COLUMN_NAME As String, chk As Infragistics.Win.UltraWinEditors.UltraCheckEditor)
        MyBase.CheckedChanged_Special(COLUMN_NAME, chk)
        Select Case COLUMN_NAME
            Case "LINE_CLOSED"

            Case "LINE_FINISHED"
                'Don't forget there is a similar area in the grid for CMT with Inv and FOB types.
                Dim EMsg As String = ""
                If EMsg <> "" Then
                    MsgBox(EMsg, vbOKOnly, "You Can Not Perform This Action For The Following Reasons")
                End If
        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "PO_DATE_SHIP_BY"
        End Select

    End Sub

    Public Overrides Sub num_ValueChanged(sender As Object, e As System.EventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If COLUMN_NAME = "PO_COST_COMM" Or COLUMN_NAME = "PO_COST_BUFFER" Then
            'Dependent_Calculations(COLUMN_NAME)
        End If
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "FREIGHT_ENTERED_BY"
            '    Dim blnFREIGHT_ENTERED_BY_Container As Boolean = (Absx1.optFor(COLUMN_NAME).Value = "C")
            '    With grdPOTORDR3.DisplayLayout.Bands(0)
            '        .Columns("CBM_RATE").Hidden = blnFREIGHT_ENTERED_BY_Container
            '        .Columns("CBM").Hidden = blnFREIGHT_ENTERED_BY_Container

            '        .Columns("BOL_FEE").Hidden = blnFREIGHT_ENTERED_BY_Container
            '        .Columns("FREIGHT_AMT").Hidden = Not blnFREIGHT_ENTERED_BY_Container
            '        .Columns("CBM").Hidden = Not blnFREIGHT_ENTERED_BY_Container
            '        .Columns("TRUCKING").Hidden = Not blnFREIGHT_ENTERED_BY_Container
            '    End With
        End Select
    End Sub
#End Region
    Private Function Get_Code_SQL(p1 As String) As String
        Throw New NotImplementedException
    End Function

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_WOTORDRX()
    End Sub

    Private Sub SetComboBoxes()
        Dim cboList As New List(Of ComboBox)
        cboList.Add(cboAPPLY_TKT)
        cboList.Add(cboAPPLY_HGR)
        cboList.Add(cboAPPLY_STK)
        cboList.Add(cboAPPLY_TAG)
        cboList.Add(cboREMOVE_TKT)
        cboList.Add(cboREMOVE_HGR)
        cboList.Add(cboREMOVE_STK)
        cboList.Add(cboREMOVE_TAG)
        For Each cbo As System.Windows.Forms.ComboBox In cboList
            cbo.Items.Clear()
            For topI As Integer = 0 To 2
                cbo.Items.Add(topI.ToString())
            Next
        Next
    End Sub

    Sub Load_WOTORDRX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ASCMAIN1.sql = "Select * from WOTORDR1"
        Select Case optStatus.Value
            Case "P"
                ASCMAIN1.sql &= " where WKORDER_STATUS = 'P'"
                grdWOTORDRX.Text = "Showing All Pending Work Orders"
            Case "T"
                ASCMAIN1.sql &= " where WKORDER_STATUS IN ('0','1','2','3')"
                grdWOTORDRX.Text = "Showing All Transmitted Work Orders"
            Case "C"
                ASCMAIN1.sql &= " where WKORDER_STATUS IN ('C')"
                grdWOTORDRX.Text = "Showing All Completed Work Orders"
            Case "X"
                ASCMAIN1.sql &= " where WKORDER_STATUS = 'X'"
                grdWOTORDRX.Text = "Showing All Cancelled Work Orders"
            Case "R"
                ASCMAIN1.sql &= " where WKORDER_STATUS IN ('4','5')"
                grdWOTORDRX.Text = "Showing All Pending Recall Work Orders"
            Case "D"
                ASCMAIN1.sql &= " where WKORDER_STATUS = ('6')"
                grdWOTORDRX.Text = "Showing All Denied Recall Work Orders"
            Case Else
                grdWOTORDRX.Text = "Showing All Work Orders"
        End Select
        Fill_Records("WOTORDRX", "", True, ASCMAIN1.sql)

        SetComboBoxes()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdWOTORDRX_DoubleClick(sender As Object, e As System.EventArgs) Handles grdWOTORDRX.DoubleClick
        If grdWOTORDRX.ActiveRow IsNot Nothing AndAlso grdWOTORDRX.ActiveRow.IsDataRow Then
            Absx1.txtFor("WKORDER_NO").Text = grdWOTORDRX.ActiveRow.Cells("WKORDER_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdWOTORDRX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWOTORDRX.InitializeLayout

    End Sub

    Private Sub grdWOTORDR3_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWOTORDR3.AfterRowActivate
        If grdWOTORDR3.ActiveRow.IsAddRow Then
            If grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                grdWOTORDR3.ActiveCell = grdWOTORDR3.ActiveRow.Cells("STYLE_CODE")
            End If
        Else
            If grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value <> "" And grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value <> STYLE_CODE Then
                Validate_Style(grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value, False)
            End If
        End If
    End Sub

    Function Validate_Style(style_z As String, m As Boolean) As DataRow

        ' IF THE STYLE CODE IS "", THEN DO NOT PERMIT THE ENTRY OF ANYTHING ELSE

        STYLE_CODE = ""
        Dim e As String = ""

        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", style_z)

        If rowICTSTYL1 Is Nothing Then
            e = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                e = "Item Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                e = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                e = "Item does not have a valid Division Code" & vbCrLf
            End If
            If Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") = 0 Then
                e = "Styles With Empty Unit Packs Are Not Allowed. Please Change the Masterfile" & vbCrLf
            End If

        End If

        COLORS.Clear()

        If e = "" Then
            ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & style_z & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "COLOR_CODE")
                COLORS.Add(row.Item("COLOR_CODE"))
            Next
            ' WE NEED TO CODE THE EQUIVALENT OF THIS IN .NET - A SUGGESTED LIST OF COLORS FOR THE STYLE FOR GRDPOTORDR2
            'z = ""
            'ReDim COLORS(100)
            'Do While Not dyn.EOF
            '    i = i + 1
            '    z = z & ",'" & dyn.Fields("COLOR_CODE").Value & "'"
            '    COLORS(i) = dyn.Fields("COLOR_CODE").Value
            '    dyn.MoveNext()
            'Loop
            'COLORS(0) = Mid$(z, 2)
            'If COLORS(0) <> "" Then
            '    Sql = "Select * from ICWCOLR1 where COLOR_CODE in (" & COLORS(0) & ")"
            '    Sql = Sql & " order by COLOR_CODE"
            '    datICWCOLR1.RecordSource = Sql
            '    datICWCOLR1.Refresh()
            '    ssdICWCOLR1.Refresh()
            'End If
        End If

        'If e <> "" And grdWOTORDR3.ActiveRow.IsAddRow Then
        '    If m Then
        '        MsgBox(e, vbOKOnly, "Style Code Entered is Invalid because ...")
        '    End If
        'Else
        '    If e = "" Then
        '        STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE")
        '    End If
        'End If
        Return rowICTSTYL1
    End Function

    Private Sub grdWOTORDR3_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWOTORDR3.BeforeRowUpdate
        With grdWOTORDR3
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", New String() {STYLE_CODE})
            If rowICTSTYL1 Is Nothing Then
                MsgBox("Invalid Style: " & STYLE_CODE)
                e.Cancel = True
            End If

            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            If rowICTSTYC1 Is Nothing Then
                MsgBox("Invalid Color: " & COLOR_CODE)
                e.Cancel = True
            End If

            If Not e.Cancel Then
                If Val(e.Row.Cells("WKORDER_LNO").Value & "") = 0 Then
                    e.Row.Cells("WKORDER_NO").Value = WKORDER_NO
                    e.Row.Cells("WKORDER_LNO").Value = Val(dst.Tables("WOTORDR3").Compute("Max(WKORDER_LNO)", "") & "") + 1
                    e.Row.Cells("ITEM_CODE").Value = e.Row.Cells("STYLE_CODE").Value & e.Row.Cells("COLOR_CODE").Value
                End If
            End If
        End With
    End Sub

    Private Sub grdWOTORDR3_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWOTORDR3.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdWOTORDR3, sql_where)
                If IsDBNull(grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value) Then
                    Exit Sub
                End If
                Validate_Style(grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value, False)
                If COLORS.Count = 1 Then
                    grdWOTORDR3.ActiveRow.Cells("COLOR_CODE").Value = COLORS(0)
                Else
                    grdWOTORDR3.ActiveRow.Cells("COLOR_CODE").Value = ""
                End If
            Case "COLOR_CODE"
                If IsDBNull(grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value) Then
                    Exit Sub
                End If
                If grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value <> "" Then
                    Dim sql_where As String = String.Format(" COLOR_CODE IN (SELECT DISTINCT COLOR_CODE FROM ICTSTYC1 WHERE STYLE_CODE = '{0}')", grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value)
                    grdClickCellButton(grdWOTORDR3, sql_where)
                    If IsDBNull(grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value) Then
                        Exit Sub
                    End If
                    Validate_Style(grdWOTORDR3.ActiveRow.Cells("STYLE_CODE").Value, False)
                Else
                    MsgBox("You Must First Select A Style")
                End If
            Case ""
                Dim sql_where As String = ""
                grdClickCellButton(grdWOTORDR3, sql_where)
        End Select
    End Sub

    Private Sub grdWOTORDR4_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWOTORDR4.BeforeRowUpdate
        With grdWOTORDR4
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", New String() {STYLE_CODE})
            If rowICTSTYL1 Is Nothing Then
                MsgBox("Invalid Style: " & STYLE_CODE)
                e.Cancel = True
            End If

            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            If rowICTSTYC1 Is Nothing Then
                MsgBox("Invalid Color: " & COLOR_CODE)
                e.Cancel = True
            End If

            If Not e.Cancel Then
                If Val(e.Row.Cells("WKORDER_LNO").Value & "") = 0 Then
                    e.Row.Cells("WKORDER_NO").Value = WKORDER_NO
                    e.Row.Cells("WKORDER_LNO").Value = Val(dst.Tables("WOTORDR4").Compute("Max(WKORDER_LNO)", "") & "") + 1
                    e.Row.Cells("ITEM_CODE").Value = e.Row.Cells("STYLE_CODE").Value & e.Row.Cells("COLOR_CODE").Value
                End If
            End If
        End With
    End Sub

    Private Sub grdWOTORDR4_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWOTORDR4.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdWOTORDR4, sql_where)
                If Not IsDBNull(grdWOTORDR4.ActiveRow.Cells("STYLE_CODE").Value) Then
                    Validate_Style(grdWOTORDR4.ActiveRow.Cells("STYLE_CODE").Value, False)
                End If
                If COLORS.Count = 1 Then
                    grdWOTORDR4.ActiveRow.Cells("COLOR_CODE").Value = COLORS(0)
                Else
                    grdWOTORDR4.ActiveRow.Cells("COLOR_CODE").Value = ""
                End If
            Case "COLOR_CODE"
                If grdWOTORDR4.ActiveRow.Cells("STYLE_CODE").Value <> "" Then
                    Dim sql_where As String = String.Format(" COLOR_CODE IN (SELECT DISTINCT COLOR_CODE FROM ICTSTYC1 WHERE STYLE_CODE = '{0}')", grdWOTORDR4.ActiveRow.Cells("STYLE_CODE").Value)
                    grdClickCellButton(grdWOTORDR4, sql_where)
                    If Not IsDBNull(grdWOTORDR4.ActiveRow.Cells("STYLE_CODE").Value) Then
                        Validate_Style(grdWOTORDR4.ActiveRow.Cells("STYLE_CODE").Value, False)
                    End If
                Else
                    MsgBox("You Must First Select A Style")
                End If
            Case ""
                Dim sql_where As String = ""
                grdClickCellButton(grdWOTORDR4, sql_where)
        End Select
    End Sub

    Private Sub btnStyleMatch_Click(sender As Object, e As System.EventArgs) Handles btnStyleMatch.Click
        Dim FromCount As Integer = grdWOTORDR3.Selected.Rows.Count
        Dim ToCount As Integer = grdWOTORDR4.Selected.Rows.Count
        Dim EMsg As New System.Text.StringBuilder
        Dim Row3Selected As Boolean = False
        Dim Row4Selected As Boolean = False
        'Dim LCOLOR As Color
        If FromCount = 0 And ToCount = 0 Then
            EMsg.AppendLine("You Must Select At Style Lines From Each List Then Press To Join Them.")
        Else
            If FromCount = 0 Or ToCount = 0 Then
                If FromCount = 0 Then
                    EMsg.AppendLine("You Must Select At Least One Style Line To Use")
                End If
                If ToCount = 0 Then
                    EMsg.AppendLine("You Must Select At Least One Line To Create")
                End If
            Else
                For Each Row3 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR3.Selected.Rows
                    If Not IsDBNull(Row3.Cells("REF_LNO").Value) Then
                        Row3Selected = True
                    End If
                Next
                For Each Row4 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR4.Selected.Rows
                    If Not IsDBNull(Row4.Cells("REF_LNO").Value) Then
                        Row4Selected = True
                    End If
                Next
                If Row3Selected Or Row4Selected Then
                    Dim iResult As MsgBoxResult = MessageBox.Show("Some Of The Selected Lines Were Previously Used.  OK To Reset them?", "Reset Lines?", MessageBoxButtons.YesNo)
                    If iResult = MsgBoxResult.Yes Then
                        For Each Row3 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR3.Selected.Rows
                            If Not IsDBNull(Row3.Cells("REF_LNO").Value) Then
                                For Each rowWOTORDR4 As DataRow In dst.Tables("WOTORDR4").Select(String.Format("REF_LNO = {0}", Row3.Cells("REF_LNO").Value))
                                    rowWOTORDR4.Item("REF_LNO") = Null
                                Next
                                For Each rowWOTORDR3 As DataRow In dst.Tables("WOTORDR3").Select(String.Format("REF_LNO = {0}", Row3.Cells("REF_LNO").Value))
                                    rowWOTORDR3.Item("REF_LNO") = Null
                                Next
                            End If
                            Row3.Cells("REF_LNO").Value = Null
                        Next
                        For Each Row4 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR4.Selected.Rows
                            If Not IsDBNull(Row4.Cells("REF_LNO").Value) Then
                                For Each rowWOTORDR4 As DataRow In dst.Tables("WOTORDR4").Select(String.Format("REF_LNO = {0}", Row4.Cells("REF_LNO").Value))
                                    rowWOTORDR4.Item("REF_LNO") = Null
                                Next
                                For Each rowWOTORDR3 As DataRow In dst.Tables("WOTORDR3").Select(String.Format("REF_LNO = {0}", Row4.Cells("REF_LNO").Value))
                                    rowWOTORDR3.Item("REF_LNO") = Null
                                Next
                            End If
                            Row4.Cells("REF_LNO").Value = Null
                        Next
                    Else
                        EMsg.AppendLine("Reset Cancellation Selected")
                    End If
                End If
            End If
        End If

        If EMsg.Length > 0 Then
            MessageBox.Show(EMsg.ToString, "Problem Joining Records")
        Else
            Dim REF_LNO As Integer = Val(dst.Tables("WOTORDR3").Compute("Max(REF_LNO)", "") & "") + 1
            For Each Row3 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR3.Selected.Rows
                Row3.Cells("REF_LNO").Value = REF_LNO
                'Row3.Cells("REF_LNO").Appearance.BackColor = LCOLOR
            Next
            For Each Row4 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR4.Selected.Rows
                Row4.Cells("REF_LNO").Value = REF_LNO
                'Row4.Cells("REF_LNO").Appearance.BackColor = LCOLOR
            Next
            'For Each Row3 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR3.Rows
            '    If IsDBNull(Row3.Cells("REF_LNO").Value) Then
            '        Row3.Cells("REF_LNO").Appearance.BackColor = Color.Empty
            '    Else
            '        Select Case Row3.Cells("REF_LNO").Value
            '            Case 1
            '                LCOLOR = Color.Cyan
            '            Case 2
            '                LCOLOR = Color.DarkGreen
            '            Case 3
            '                LCOLOR = Color.DarkRed
            '            Case 4
            '                LCOLOR = Color.DeepPink
            '            Case 5
            '                LCOLOR = Color.Blue
            '            Case Else
            '                LCOLOR = Color.Goldenrod
            '        End Select
            '        Row3.Cells("REF_LNO").Appearance.BackColor = LCOLOR
            '    End If
            'Next
            'For Each Row4 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR4.Rows
            '    If IsDBNull(Row4.Cells("REF_LNO").Value) Then
            '        Row4.Cells("REF_LNO").Appearance.BackColor = Color.Empty
            '    Else
            '        Select Case Row4.Cells("REF_LNO").Value
            '            Case 1
            '                LCOLOR = Color.Cyan
            '            Case 2
            '                LCOLOR = Color.DarkGreen
            '            Case 3
            '                LCOLOR = Color.DarkRed
            '            Case 4
            '                LCOLOR = Color.DeepPink
            '            Case 5
            '                LCOLOR = Color.Blue
            '            Case Else
            '                LCOLOR = Color.Goldenrod
            '        End Select
            '        Row4.Cells("REF_LNO").Appearance.BackColor = LCOLOR
            '    End If
            'Next
            grdWOTORDR4.UpdateData()
            grdWOTORDR3.UpdateData()
            grdWOTORDR3.Selected.Rows.Clear()
            grdWOTORDR4.Selected.Rows.Clear()
        End If
    End Sub

    Private Sub btnStyleMatch_MouseHover(sender As Object, e As System.EventArgs) Handles btnStyleMatch.MouseHover
        ShowToolTip(sender, e)
    End Sub

    Private Sub grdWOTORDR3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWOTORDR3.InitializeRow
        SetMatchColors(e)
        CalcTotalUnits()
    End Sub

    Private Sub grdWOTORDR4_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWOTORDR4.InitializeRow
        SetMatchColors(e)
    End Sub

    Private Sub SetMatchColors(e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        Dim LCOLOR As Color
        If IsDBNull(e.Row.Cells("REF_LNO").Value) Then
            e.Row.Cells("REF_LNO").Appearance.BackColor = Color.Empty
        Else
            Select Case e.Row.Cells("REF_LNO").Value
                Case 1
                    LCOLOR = Color.Cyan
                Case 2
                    LCOLOR = Color.DarkGreen
                Case 3
                    LCOLOR = Color.DarkRed
                Case 4
                    LCOLOR = Color.DeepPink
                Case 5
                    LCOLOR = Color.Blue
                Case Else
                    LCOLOR = Color.Goldenrod
            End Select
            e.Row.Cells("REF_LNO").Appearance.BackColor = LCOLOR
        End If
    End Sub

    Private Function AreAllStylesMatched() As Boolean
        Dim RetVal As Boolean = True
        For Each Row4 As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWOTORDR4.Rows
            If IsDBNull(Row4.Cells("REF_LNO").Value) Then
                RetVal = False
            Else
                If RetVal Then
                    If Not IsNumeric(Row4.Cells("REF_LNO").Value) Then
                        RetVal = False
                    Else
                        If dst.Tables("WOTORDR3").Select(String.Format("REF_LNO = {0}", Val(Row4.Cells("REF_LNO").Value))).Count() = 0 Then
                            RetVal = False
                        End If
                    End If
                End If
            End If
        Next
        Return RetVal
    End Function

    Private Sub txtSHIP_DATE_MouseHover(sender As Object, e As System.EventArgs) Handles txtSHIP_DATE.MouseHover
        ShowToolTip(sender, e)
    End Sub

    Private Sub ShowToolTip(sender As Object, e As System.EventArgs)
        tip.AutoPopDelay = 3000
        tip.InitialDelay = 3000
        tip.DisplayStyle = ToolTipDisplayStyle.BalloonTip
        tip.ShowToolTip(sender)
    End Sub

    Private Sub txtCANCEL_DATE_MouseHover(sender As Object, e As System.EventArgs) Handles txtCANCEL_DATE.MouseHover
        ShowToolTip(sender, e)
    End Sub

    Private Sub txtORDR_GROUP_PO_MouseHover(sender As Object, e As System.EventArgs) Handles txtORDR_GROUP_PO.MouseHover
        ShowToolTip(sender, e)
    End Sub

    Private Sub txtORDR_PO_MouseHover(sender As Object, e As System.EventArgs) Handles txtORDR_PO.MouseHover
        ShowToolTip(sender, e)
    End Sub

    Private Sub btnAddStyles_MouseHover(sender As Object, e As System.EventArgs) Handles btnAddStyles.MouseHover
        ShowToolTip(sender, e)
    End Sub

    Private Sub btnClear_MouseHover(sender As Object, e As System.EventArgs) Handles btnClear.MouseHover
        ShowToolTip(sender, e)
    End Sub

    Private Sub btnAddStyles_Click(sender As System.Object, e As System.EventArgs) Handles btnAddStyles.Click
        Dim ORDR_NO As String = txtORDR_NO.Text
        Dim ORDR_GROUP_NO As String = txtORDR_GROUP_NO.Text
        Dim SQL As String = ""
        Dim WKORDER_LNO As Integer = 0
        Dim Msg As String = "This Will Clear Both Left And Right Grids And"
        Msg += vbCr
        Msg += "Load The Styles And Colors From The Orders"
        Msg += vbCr
        Msg += "Selected Above."
        If ORDR_NO.Length = 0 And ORDR_GROUP_NO = 0 Then
            MsgBox("No Order or Group Selected Above.", MsgBoxStyle.OkOnly, "Select Order Or Group")
            Exit Sub
        End If
        Dim iResponse As MsgBoxResult = MsgBox(Msg, MsgBoxStyle.OkCancel, "Load Styles From Orders")
        If iResponse = MsgBoxResult.Ok Then
            If ORDR_NO.Length > 0 Then
                SQL = "SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_QTY) ORDR_QTY" & vbCrLf _
                    & " FROM SOTORDR2" & vbCrLf _
                    & " WHERE ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                    & " GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
                    & " ORDER BY STYLE_CODE, COLOR_CODE"
            Else
                SQL = "SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_QTY) ORDR_QTY" & vbCrLf _
                    & " FROM SOTORDR2" & vbCrLf _
                    & " WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')" & vbCrLf _
                    & " GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
                    & " ORDER BY STYLE_CODE, COLOR_CODE"
            End If
            dst.Tables("WOTORDR3").Rows.Clear()
            dst.Tables("WOTORDR4").Rows.Clear()
            dst.Tables("SOTORDR2").Rows.Clear()
            ASCMAIN1.sql = SQL
            Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "STYLE_CODE")
                WKORDER_LNO = WKORDER_LNO + 1
                Dim rowWOTORDR3 As DataRow = dst.Tables("WOTORDR3").NewRow
                rowWOTORDR3.Item("WKORDER_NO") = WKORDER_NO
                rowWOTORDR3.Item("WKORDER_LNO") = WKORDER_LNO
                rowWOTORDR3.Item("REF_LNO") = Null
                rowWOTORDR3.Item("STYLE_CODE") = rowSOTORDR2.Item("STYLE_CODE").ToString
                rowWOTORDR3.Item("COLOR_CODE") = rowSOTORDR2.Item("COLOR_CODE").ToString
                rowWOTORDR3.Item("PPK_CODE") = Null
                rowWOTORDR3.Item("ITEM_CODE") = String.Format("{0}{1}", rowSOTORDR2.Item("STYLE_CODE"), rowSOTORDR2.Item("COLOR_CODE"))
                rowWOTORDR3.Item("FROM_QTY_REQ") = rowSOTORDR2.Item("ORDR_QTY").ToString
                rowWOTORDR3.Item("FROM_QTY_ACT") = Null
                dst.Tables("WOTORDR3").Rows.Add(rowWOTORDR3)
            Next

        End If
    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        dst.Tables("WOTORDR3").Rows.Clear()
        dst.Tables("WOTORDR4").Rows.Clear()
        CalcTotalUnits()
    End Sub

    Private Sub CalcTotalUnits()
        TotalUnits = 0
        For Each rowWOTORDR3 As DataRow In dst.Tables("WOTORDR3").Select()
            TotalUnits += Val(rowWOTORDR3.Item("FROM_QTY_REQ") & "")
        Next
    End Sub

    Private Sub grdWOTORDR3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWOTORDR3.AfterRowUpdate
        CalcTotalUnits()
    End Sub

    Private Sub txtVarious_DoubleClick(sender As Object, e As System.EventArgs) Handles txtHANGER_ONE_QTY.DoubleClick, _
        txtHANGER_TWO_QTY.DoubleClick, _
        txtHANGER_REMOVE_QTY.DoubleClick, _
        txtHANGER_CAP_QTY.DoubleClick, _
        txtAPPLY_UNITS.DoubleClick, _
        txtREMOVE_UNITS.DoubleClick, _
        txtTKTS_CONT_FORM.DoubleClick, _
        txtLBL_CUT_QTY.DoubleClick, _
        txtREPACK_NEW_CARTS.DoubleClick, _
        txtSETS_MADE_QTY.DoubleClick, _
        txtSETS_SEPERATED_QTY.DoubleClick, _
        txtTKTS_UNIQUE_PER_CARTON.DoubleClick
        Dim txtBox As Infragistics.Win.UltraWinEditors.UltraNumericEditor = sender
        If TotalUnits = 0 Then
            txtBox.Value = Null
        Else
            txtBox.Value = TotalUnits
        End If
    End Sub

    Private Sub grdWOTORDR3_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWOTORDR3.AfterRowsDeleted
        CalcTotalUnits()
    End Sub

    Private Sub txtORDR_GROUP_NO_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtORDR_GROUP_NO.ValueChanged
        SetGroupNoData()
    End Sub

    Private Sub txtORDR_NO_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtORDR_NO.ValueChanged
        SetOrderNoData()
    End Sub

    Private Sub txtSHIP_BOL_NO_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtSHIP_BOL_NO.ValueChanged
        SetBOLNoData()
    End Sub

    Private Sub SyncRecordsWithADS(ByVal WKORDER_STATUS_FROM As String)
        Dim WKORDER_STATUS_TO As String = ""
        Select Case WKORDER_STATUS_FROM
            Case "2"
                WKORDER_STATUS_TO = "1"
            Case "3"
                WKORDER_STATUS_TO = "C"
            Case Else
                Exit Sub
        End Select

        Dim SQLS As New StringBuilder() With {.Length = 0}

        'This needs to be tested!!!!
        If 1 = 1 Then
            Exit Sub
        End If

        BeginTrans()
        ASCMAIN1.sql = String.Format("UPDATE ADS.WOTORDR1@ADSIIS SET WKORDER_STATUS = 'V' WHERE WKORDER_STATUS = '{0}'", WKORDER_STATUS_FROM)
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        SQLS.AppendLine("SELECT *")
        SQLS.AppendLine(" FROM ADS.WOTORDR1@ADSIIS")
        SQLS.AppendLine(" WHERE WKORDER_STATUS = 'V'")
        Using tblWOTORDR1_LP As DataTable = ASCDATA1.GetDataTable(SQLS.ToString())
            For Each rowWOTORDR1_LP As DataRow In tblWOTORDR1_LP.Rows
                'Update Oracle Tables
                Dim WKORDER_NO As String = rowWOTORDR1_LP.Item("WKORDER_NO").ToString

                SQLS.AppendLine("UPDATE WOTORDR1")
                SQLS.AppendLine(String.Format(" SET WKORDER_STATUS = '{0}'", WKORDER_STATUS_TO))
                SQLS.AppendLine(String.Format(" ,WKORDER_NOTES = '{0}'", rowWOTORDR1_LP.Item("WKORDER_NOTES").ToString))
                SQLS.AppendLine(String.Format(" ,LAST_OPER = '{0}'", ASCMAIN1.USER_ID))
                SQLS.AppendLine(String.Format(" ,LAST_DATE = '{0}'", DateValue(Format$(Now, "MM/dd/yyyy"))))
                SQLS.AppendLine(String.Format(" WHERE WKORDER_NO = '{0}'", WKORDER_NO))
                ASCDATA1.ExecuteSQL(SQLS.ToString())

                SQLS.AppendLine("SELECT *")
                SQLS.AppendLine(" FROM ADS.WOTORDR3@ADSIIS")
                SQLS.AppendLine(String.Format(" WHERE WKORDER_NO = '{0}'", WKORDER_NO))
                Using tblWOTORDR3_LP As DataTable = ASCDATA1.GetDataTable(SQLS.ToString())
                    For Each rowWOTORDR3_LP As DataRow In tblWOTORDR3_LP.Rows
                        Dim WKORDER_LNO As Integer = rowWOTORDR3_LP.Item("WKORDER_LNO").ToString
                        SQLS.AppendLine("UPDATE WOTORDR3")
                        SQLS.AppendLine(String.Format(" SET FROM_QTY_ACT = {0}", rowWOTORDR3_LP.Item("FROM_QTY_ACT").ToString))
                        SQLS.AppendLine(String.Format(" WHERE WKORDER_NO = '{0}' AND WKORDER_LNO = {1}", WKORDER_NO, WKORDER_LNO))
                        ASCDATA1.ExecuteSQL(SQLS.ToString())
                    Next
                End Using

                SQLS.AppendLine("SELECT *")
                SQLS.AppendLine(" FROM ADS.WOTORDR4@ADSIIS")
                SQLS.AppendLine(String.Format(" WHERE WKORDER_NO = '{0}'", WKORDER_NO))
                Using tblWOTORDR4_LP As DataTable = ASCDATA1.GetDataTable(SQLS.ToString())
                    For Each rowWOTORDR4_LP As DataRow In tblWOTORDR4_LP.Rows
                        Dim WKORDER_LNO As Integer = rowWOTORDR4_LP.Item("WKORDER_LNO").ToString
                        SQLS.AppendLine("UPDATE WOTORDR4")
                        SQLS.AppendLine(String.Format(" SET FROM_QTY_ACT = {0}", rowWOTORDR4_LP.Item("FROM_QTY_ACT").ToString))
                        SQLS.AppendLine(String.Format(" WHERE WKORDER_NO = '{0}' AND WKORDER_LNO = {1}", WKORDER_NO, WKORDER_LNO))
                        ASCDATA1.ExecuteSQL(SQLS.ToString())
                    Next
                End Using

                SQLS.AppendLine("SELECT *")
                SQLS.AppendLine(" FROM ADS.WOTORDR5@ADSIIS")
                SQLS.AppendLine(String.Format(" WHERE WKORDER_NO = '{0}'", WKORDER_NO))
                Using tblWOTORDR5_LP As DataTable = ASCDATA1.GetDataTable(SQLS.ToString())
                    For Each rowWOTORDR5_LP As DataRow In tblWOTORDR5_LP.Rows
                        SQLS.AppendLine("UPDATE WOTORDR5")
                        SQLS.AppendLine(String.Format(" SET DATE_PULLED = '{0}'", rowWOTORDR5_LP.Item("DATE_PULLED").ToString))
                        SQLS.AppendLine(String.Format(" ,DATE_COMPLETED = '{0}'", rowWOTORDR5_LP.Item("DATE_COMPLETED").ToString))
                        SQLS.AppendLine(String.Format(" ,EXT_JOB_NO = '{0}'", rowWOTORDR5_LP.Item("EXT_JOB_NO").ToString))
                        SQLS.AppendLine(String.Format(" ,VAS_JOB_QTY = {0}", rowWOTORDR5_LP.Item("VAS_JOB_QTY").ToString))
                        SQLS.AppendLine(String.Format(" ,EXT_JOB_NO = '{0}'", rowWOTORDR5_LP.Item("EXT_JOB_NO").ToString))
                        SQLS.AppendLine(String.Format(" ,UNITS_WORKED = '{0}'", rowWOTORDR5_LP.Item("UNITS_WORKED").ToString))
                        SQLS.AppendLine(String.Format(" ,LAST_OPER = '{0}'", ASCMAIN1.USER_ID))
                        SQLS.AppendLine(String.Format(" ,LAST_DATE = '{0}'", DateValue(Format$(Now, "MM/dd/yyyy"))))
                        SQLS.AppendLine(String.Format(" WHERE WKORDER_NO = '{0}'", WKORDER_NO))
                        ASCDATA1.ExecuteSQL(SQLS.ToString())
                    Next
                End Using
            Next
        End Using

        ASCMAIN1.sql = String.Format("UPDATE ADS.WOTORDR1@ADSIIS SET WKORDER_STATUS = '{0}' WHERE WKORDER_STATUS = 'V'", WKORDER_STATUS_TO)
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        CommitTrans()
    End Sub

    Private Sub grdWOTORDR3_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWOTORDR3.BeforeExitEditMode
        With grdWOTORDR3.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE", "COLOR_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdWOTORDR4_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWOTORDR4.BeforeExitEditMode
        With grdWOTORDR4.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE", "COLOR_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub txtWHSE_CODE_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtWHSE_CODE.Validating
        txtWHSE_CODE.Text = txtWHSE_CODE.Text.ToUpper()
        ASCMAIN1.sql = String.Format("SELECT COUNT(*) FROM ICTWHSE1 WHERE WHSE_CODE = '{0}'", txtWHSE_CODE.Text)
        Dim RECCNT As Int16 = Val(ASCDATA1.GetDataValue)
        If RECCNT = 0 Then
            MessageBox.Show("Invalid Warehouse")
            txtWHSE_CODE.Text = ""
        End If
    End Sub

    Private Sub MarkAsDeleted()
        For Each rowWOTORDR1 As DataRow In dst.Tables("WOTORDR1").Select()
            rowWOTORDR1.Item("WKORDER_STATUS") = "X"
        Next
    End Sub

    Private Sub RecallWO()
        Throw New NotImplementedException
    End Sub

    Private Sub SetWKORDER_STATUS()
        If Not IsNothing(rowWOTORDR1) Then
            If Not IsDBNull(rowWOTORDR1.Item("WKORDER_STATUS")) Then
                Select Case rowWOTORDR1.Item("WKORDER_STATUS")
                    Case "P"
                        optWKORDER_STATUS.Value = "P"
                    Case "0", "1", "2", "3"
                        optWKORDER_STATUS.Value = "T"
                    Case "C"
                        optWKORDER_STATUS.Value = "C"
                    Case "X"
                        optWKORDER_STATUS.Value = "X"
                    Case "4", "5", "7"
                        optWKORDER_STATUS.Value = "R"
                    Case "6"
                        optWKORDER_STATUS.Value = "D"
                    Case Else
                        optWKORDER_STATUS.Value = "P"
                End Select

            End If
        End If

    End Sub
    Sub Print_Record()
        Synch_TABLE_NAME("WOTORDR1")
        Print_Report_Begin()
        'Dim MODE As String = "S"
        'If receipt_mode Then MODE = "R"
        'If cost_calc Then MODE = "C"
        'CR_params.Add("MODE", MODE)
        Generate_Report("WORORDR1", Me.Text, String.Format("Work Order {0}{1}", WKORDER_NO, IIf((EntryMode <> "E"), "", " - Edit in Process")))
        Print_Report_End()
    End Sub
End Class