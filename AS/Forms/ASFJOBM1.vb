Imports ABSolution

Public Class ASFJOBM1

    '***********************************************************************************************
    ' Open Issues
    '   Some reports require selections to me made. How does user know to create a Saved Setting.
    '   If report is supposed to be updated and for some reason the report cannot be updated
    '       example: out of balance, then the variable 'updateSuccessful' needs to be set to false.
    '       Need to allow for all circumstances where an Updateable report cannot be updated.
    '       What about if an Updateable Report does not have any rows to Update. This should be fine.
    '       example: Open Payables Report has no records to update.
    '***********************************************************************************************

    ' Tables Needed

    '  CREATE TABLE ASTJOBM1 ( 
    '  JOB_STREAM_CODE  VARCHAR2 (10)  NOT NULL, 
    '  JOB_STREAM_DESC  VARCHAR2 (40), 
    '  JOB_STREAM_TYPE  VARCHAR2 (1), 
    '  INIT_OPER VARCHAR2(8),
    '  INIT_DATE DATE,
    '  LAST_OPER VARCHAR2(8),
    '  LAST_DATE DATE,
    '  ACTIVE VARCHAR2(1), 
    '  USER_ID VARCHAR2(8)
    '  PRIMARY KEY ( JOB_STREAM_CODE ) ) ;  

    '  CREATE TABLE ASTJOBM2 ( 
    '  JOB_STREAM_CODE  VARCHAR2 (10)  NOT NULL, 
    '  JOB_STREAM_LNO   NUMBER (3)    NOT NULL, 
    '  REPORT_ID        VARCHAR2 (8), 
    '  SET_ID           VARCHAR2 (10), 
    '  MENU_ID          VARCHAR2 (8), 
    '  UPDATE_REPORT    VARCHAR2 (1), 
    '  PRIMARY KEY ( JOB_STREAM_CODE, JOB_STREAM_LNO ) ) ; 

    '  CREATE TABLE ASTJOBM3 ( 
    '  JOB_STREAM_XNO   VARCHAR2 (10)  NOT NULL, 
    '  JOB_STREAM_CODE  VARCHAR2 (10)  NOT NULL, 
    '  JOB_STREAM_LNO   NUMBER (3)    NOT NULL, 
    '  REPORT_ID        VARCHAR2 (8), 
    '  SET_ID           VARCHAR2 (10), 
    '  MENU_ID          VARCHAR2 (8), 
    '  INIT_DATE        DATE, 
    '  LAST_DATE        DATE, 
    '  SUCCESS          VARCHAR2 (1), 
    '  PRIMARY KEY ( JOB_STREAM_XNO, JOB_STREAM_CODE, JOB_STREAM_LNO ) ) ; 

    '  CREATE TABLE ASTJOBM4 ( 
    '  JOB_STREAM_XNO   VARCHAR2 (10)  NOT NULL, 
    '  JOB_STREAM_CODE  VARCHAR2 (10), 
    '  JOB_STREAM_DESC  VARCHAR2 (40), 
    '  JOB_STREAM_TYPE  VARCHAR2 (1), 
    '  INIT_OPER        VARCHAR2 (8), 
    '  LAST_OPER        VARCHAR2 (8), 
    '  INIT_DATE        DATE, 
    '  LAST_DATE        DATE, 
    '  PRIMARY KEY ( JOB_STREAM_XNO ) ) ; 

    '  CREATE TABLE ASTJOBM5 ( 
    '  JOB_STREAM_CODE          VARCHAR2 (10)  NOT NULL, 
    '  JOB_START_TIME            DATE, 
    '  JOB_RECURR_PATTERN        VARCHAR2 (1), 
    '  JOB_RANGE_LIMIT           VARCHAR2 (1), 
    '  JOB_RANGE_MAX_OCCUR  NUMBER (3), 
    '  JOB_RANGE_END_DATE         DATE, 
    '  JOB_PATTERN_DAY_OF_MONTH    NUMBER (2), 
    '  JOB_PATTERN_DAYS_OF_WEEK    VARCHAR2 (11), 
    '  JOB_PATTERN_INTERVAL      NUMBER (6), 
    '  JOB_PATTERN_MONTH_OF_YEAR   NUMBER (2), 
    '  JOB_OCCUR_OF_DAY_IN_MONTH    NUMBER (2), 
    '  JOB_PATTERN_TYPE          NUMBER (1), 
    '  PRIMARY KEY ( JOB_STREAM_CODE ) ) ;

    ' CREATE TABLE  ASTJOBM6 (
    ' JOB_STREAM_CODE VARCHAR2(10),
    ' JOB_START_TIME DATE,
    ' PROCESSED VARCHAR2(0),
    ' PRIMARY KEY (JOB_STREAM_CODE, JOB_START_TIME));

    ' Create Views
    '    ASTJOBM1.ASTJOBM1
    '      select columns JOB_STREAM_CODE (return value), JOB_STREAM_DESC

    '    ASTMENU1.MENU_ITEM_OBJECT
    '      where clause: MENU_ITEM_TYPE = 'R'
    '      select columns MENU_ID (hide), MENU_ITEM_DESC, MENU_ITEM_OBJECT (return value)

    ' ********************** NOTES *************************
    ' Need View for JOB_STREAM_TYPE S=Standard, W=Web Reports in TACMAIN1.Code_Values
    ' Need View for ASTJOBM1.JOB_STREAM_CODE
    ' Need View for ASTMENU1.MENU_ITEM_OBJECT - Where MENU_ITEM_TYPE - 'R' 
    ' Need View for ASTROPT1.SET_ID

    ' Needed in ASCMAIN1
    '   Public Shared JOB_STREAM_CODE As String = String.Empty
    '   Public Shared JOB_STREAM_FORM_NAME As String = String.Empty
    '   Public Shared JOB_STREAM_XNO As String = String.Empty
    '   Public Shared JOB_STREAM_LNO As Int16 = 0

    ' New Class TAC.ASCAPPT1
    ' Code changes to ASFLOGON, ASFMAIN1, ASFSRPTM - These changes exist in VDI

    Private sqlASTJOBM4_C As String = String.Empty
    Private JOB_STREAM_CODE As String = String.Empty
    Private REPORT_ID As String = String.Empty
    Private SET_ID As String = String.Empty
    Private rowASTJOBM1 As DataRow
    Private threadCompleted As Boolean = False

    Private JOB_MENU_ITEM_OBJECT As String
    Private JOB_MENU_ITEM_TYPE As String
    Private JOB_MENU_ID As String
    Private JOB_STREAM_XNO As String

    Private ExecutingScript As Boolean = False
    Private JOB_STREAM_LNO As Int16 = 0
    Private tblASTJOBM2 As DataTable = Nothing
    Private recurrenceChanged As Boolean = False
    Private recurringEventDates As New List(Of Date)

#Region "ABS Standard Routines"

    Private Sub ASFJOBM1_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown

        If Environment.GetCommandLineArgs.Count >= 6 Then
            If Environment.GetCommandLineArgs.ElementAt(4) = "JS" Then
                Dim JOB_STREAM_CODE As String = Environment.GetCommandLineArgs.ElementAt(5)

                If ASCMAIN1.JOB_STREAM_CODE = JOB_STREAM_CODE Then
                    Try
                        txtJOB_STREAM_CODE.Text = JOB_STREAM_CODE
                        System.Threading.Thread.Sleep(2000)
                        MyBase.Click_Command("View")
                        System.Threading.Thread.Sleep(2000)
                        MyBase.Click_Command("Execute")
                        System.Threading.Thread.Sleep(2000)
                    Catch ex As Exception
                    Finally
                        Me.BeginInvoke(New MethodInvoker(AddressOf Me.Close))
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub ASFJOBM1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If ExecutingScript Then
            timExecute.Start()
        End If
    End Sub

    ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        With dst

            Create_TDA(.Tables.Add, "ASTJOBM1", "*")
            Create_TDA(.Tables.Add, "ASTJOBM2", "*", 1)
            Create_TDA(.Tables.Add, "ASTJOBM5", "*", 1)
            Create_TDA(.Tables.Add, "ASTJOBM6", "*", 1)

            .Tables("ASTJOBM2").Columns.Add("MENU_ITEM_DESC", GetType(System.String))
            .Tables("ASTJOBM2").Columns.Add("SET_DESC", GetType(System.String))

            Create_TDA(.Tables.Add, "ASTJOBM3", "*")
            .Tables("ASTJOBM3").Columns.Add("MENU_ITEM_DESC", GetType(System.String))
            .Tables("ASTJOBM3").Columns.Add("SET_DESC", GetType(System.String))
            .Tables("ASTJOBM3").Columns.Add("ELAPSED", GetType(System.String))

            Create_TDA(.Tables.Add, "ASTJOBM4", "*")
            .Tables("ASTJOBM4").Columns.Add("JOBS", GetType(System.Int64))
            .Tables("ASTJOBM4").Columns.Add("DONE", GetType(System.Int64))
            .Tables("ASTJOBM4").Columns.Add("ELAPSED", GetType(System.String))

            sqlASTJOBM4_C = "Select JOB_STREAM_XNO, Count (*) JOBS, Count (LAST_DATE) DONE from ASTJOBM3 GROUP BY JOB_STREAM_XNO"
            Create_TDA(.Tables.Add, "ASTJOBM4_C", sqlASTJOBM4_C, 0, False, "", 1)

            sql = "Select * from ASTMENU1 where MENU_ITEM_TYPE = 'R'"
            Create_TDA(.Tables.Add, "ASTMENU1", sql)
            Fill_Records("ASTMENU1", String.Empty, True, sql)

            sql = "Select * from ASTSPRF1 where JOB_STREAM_XNO = :PARM1"
            Create_TDA(.Tables.Add, "ASTSPRF1", sql, 0, False, "V", 1)

            sql = "Select * from ASTROPT1"
            Create_TDA(.Tables.Add, "ASTROPT1", sql)
            Fill_Records("ASTROPT1", String.Empty, True, sql)

            ' JOB_STREAM_CODE
            .Relations.Add("ASTJOBM4_ASTJOBM4_C", _
                    New DataColumn() {.Tables("ASTJOBM4").Columns("JOB_STREAM_XNO")}, _
                    New DataColumn() {.Tables("ASTJOBM4_C").Columns("JOB_STREAM_XNO")})

            .Tables("ASTJOBM4").Columns("JOBS").Expression = "SUM(CHILD(ASTJOBM4_ASTJOBM4_C).JOBS)"
            .Tables("ASTJOBM4").Columns("DONE").Expression = "SUM(CHILD(ASTJOBM4_ASTJOBM4_C).DONE)"

            grdASTJOBM2.DataSource = dst.Tables("ASTJOBM2")
            grdASTJOBM3.DataSource = dst.Tables("ASTJOBM3")
            grdASTJOBM4.DataSource = dst.Tables("ASTJOBM4")

            .Tables.Add("ASTJOBTYPE")
            .Tables("ASTJOBTYPE").Columns.Add("JOB_STREAM_TYPE", GetType(System.String))
            .Tables("ASTJOBTYPE").Columns.Add("JOB_STREAM_TYPE_DESC", GetType(System.String))
            Dim tacTACMAIN1 As New TAC.TACMAIN1
            Dim jstDict As New Dictionary(Of String, String)
            jstDict = tacTACMAIN1.CodeValues("JOB_STREAM_TYPE")
            For Each kvp As KeyValuePair(Of String, String) In jstDict
                .Tables("ASTJOBTYPE").Rows.Add(New Object() {kvp.Key, kvp.Value})
            Next
            cmbJOB_STREAM_TYPE.DataSource = dst.Tables("ASTJOBTYPE")
            cmbJOB_STREAM_TYPE.ValueMember = "JOB_STREAM_TYPE"
            cmbJOB_STREAM_TYPE.DisplayMember = "JOB_STREAM_TYPE_DESC"

        End With

        Bind_Controls(grpHeader, "ASTJOBM1")

        splJobs.Parent = tabMain.Parent
        grdASTJOBM2.Parent = tabMain.Parent
        tabMain.Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = String.Empty

        Dim sql As String = String.Empty
        Dim zMsg As String = String.Empty

        Select Case eItemKey

            Case "New"
                MyBase.Absx1.txtFor("JOB_STREAM_CODE").Text = MyBase.Absx1.txtFor("JOB_STREAM_CODE").Text.ToUpper.Trim
                If MyBase.Absx1.txtFor("JOB_STREAM_CODE").TextLength = 0 Then
                    EMsg &= vbCrLf & "Job Stream Code is required."
                Else
                    Validate_Code("JOB_STREAM_CODE", True)
                End If

                MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text = MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text.ToUpper.Trim
                If MyBase.Absx1.txtFor("JOB_STREAM_DESC").TextLength = 0 Then
                    EMsg &= vbCrLf & "Job Stream Description is required."
                Else
                    MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text = StrConv(MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text, VbStrConv.ProperCase)
                End If

                If cmbJOB_STREAM_TYPE.Value = String.Empty Then
                    EMsg &= vbCrLf & "Job Stream Type is required."
                End If

            Case "Edit", "View"
                MyBase.Absx1.txtFor("JOB_STREAM_CODE").Text = MyBase.Absx1.txtFor("JOB_STREAM_CODE").Text.ToUpper.Trim
                Validate_Code("JOB_STREAM_CODE", False)

            Case "Update"
                EMsg = ValidateData()

                If EntryMode = "N" AndAlso recurringEventDates.Count = 0 Then
                    EMsg &= vbCr & "You must Set and Save the Recurrence for this Job Stream Entry."
                End If

                If EMsg.Length = 0 Then
                    If MessageBox.Show("Do you want to Update changes?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If MessageBox.Show("Do you want to Cancel changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Delete"
                If MessageBox.Show("Do you want to Delete this Job Stream?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Execute"
                If Me.MdiParent.MdiChildren.Length <> 1 Then
                    EMsg &= vbCr & "To execute a Job Stream, " & MENU_ITEM_DESC & " form must be the only open form."
                End If
        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                MyBase.EntryMode = "N"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "View"
                MyBase.EntryMode = "V"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Edit"
                MyBase.EntryMode = "E"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Update"
                Update_Record()
                Me.Mode_Settings(False)

            Case "Cancel", "Done"
                Me.Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Me.Mode_Settings(False)

            Case "Execute"
                ExecuteJobStream(txtJOB_STREAM_CODE.Text)

            Case "Re-Number"
                Dim eMessage As String = ReNumberDetails()
                If eMessage.Length > 0 Then
                    MessageBox.Show(eMessage, "Re Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            Case "Recurrence"
                DisplayRecurrence()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1.Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If
                .Items("Done").Settings.Enabled = iScreenMode

                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                If EntryMode = "V" Then
                    .Items("Delete").Settings.Enabled = iScreenMode
                Else
                    .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                End If

                .Items("Re-Number").Settings.Enabled = iScreenMode
                .Items("Recurrence").Settings.Enabled = iScreenMode
                .Items("Execute").Settings.Enabled = iScreenMode

                '.Items("New").Visible = EntryMode <> "V"
                '.Items("View").Visible = EntryMode <> "V"
                '.Items("Edit").Visible = EntryMode <> "V"
                .Items("Done").Visible = EntryMode = "V"
                .Items("Update").Visible = EntryMode <> "V"
                .Items("Cancel").Visible = EntryMode <> "V"
                .Items("Delete").Visible = EntryMode = "V"
                .Items("Re-Number").Visible = EntryMode <> "V"
                .Items("Recurrence").Visible = EntryMode <> "V"
                .Items("Execute").Visible = EntryMode = "V"

            End With
        End If

        If EntryMode = "V" Then
            MyBase.Set_Read_Only(grpHeader, True)
            With grdASTJOBM2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        Else
            MyBase.Set_Read_Only(grpHeader, False)
            MyBase.Set_Read_Only(txtJOB_STREAM_CODE, ScreenMode)
            MyBase.Set_Read_Only(txtUSER_ID, Not ScreenMode)
            MyBase.Set_Read_Only(chkActive, Not ScreenMode)
            With grdASTJOBM2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            End With
        End If

        If ScreenMode Then
            grdASTJOBM2.Visible = True
            splJobs.Visible = False
        Else
            Me.Clear_Record()
            splJobs.Visible = True
            grdASTJOBM2.Visible = False
        End If

    End Sub

    Private Sub Clear_Record()
        MyBase.EnforceConstraints(False)

        For Each tableName As String In New String() {"ASTJOBM1", "ASTJOBM2", "ASTJOBM3", "ASTJOBM4", "ASTJOBM5", "ASTJOBM6", "ASTJOBM4_C", "ASTSPRF1"}
            dst.Tables(tableName).Rows.Clear()
        Next

        Dim sql As String = String.Empty

        sql = "Select ASTJOBM3.*, ASTMENU1.MENU_ITEM_DESC, ASTROPT1.SET_DESC,"
        sql &= " lpad(nvl(floor(((ASTJOBM3.last_date - ASTJOBM3.init_date)*24*60*60)/3600), 0), 2, '0')"
        sql &= " || ':' ||"
        sql &= " lpad(nvl(floor((((ASTJOBM3.last_date - ASTJOBM3.init_date)*24*60*60) -"
        sql &= " floor(((ASTJOBM3.last_date - ASTJOBM3.init_date)*24*60*60)/3600)*3600)/60), 0), 2, '0')"
        sql &= " || ':' ||"
        sql &= " lpad(nvl(round((((ASTJOBM3.last_date - ASTJOBM3.init_date)*24*60*60) -"
        sql &= " floor(((ASTJOBM3.last_date - ASTJOBM3.init_date)*24*60*60)/3600)*3600 -"
        sql &= " (floor((((ASTJOBM3.last_date - ASTJOBM3.init_date)*24*60*60) -"
        sql &= " floor(((ASTJOBM3.last_date - ASTJOBM3.init_date)*24*60*60)/3600)*3600)/60)*60) )), 0), 2, '0')"
        sql &= " ELAPSED"
        sql &= " from ASTJOBM3, ASTMENU1, ASTROPT1 "
        sql &= "   where ASTMENU1.MENU_ID (+) = ASTJOBM3.MENU_ID"
        sql &= "   and ASTMENU1.MENU_ITEM_OBJECT (+) = ASTJOBM3.REPORT_ID"
        sql &= "   and ASTROPT1.FORM_NAME (+) = ASTJOBM3.REPORT_ID"
        sql &= "   and ASTROPT1.SET_ID (+) = ASTJOBM3.SET_ID"
        Fill_Records("ASTJOBM3", "", True, sql)

        sql = "Select ASTJOBM4.*, "
        sql &= " lpad(nvl(floor(((ASTJOBM4.last_date - ASTJOBM4.init_date)*24*60*60)/3600), 0), 2, '0')"
        sql &= " || ':' ||"
        sql &= " lpad(nvl(floor((((ASTJOBM4.last_date - ASTJOBM4.init_date)*24*60*60) -"
        sql &= " floor(((ASTJOBM4.last_date - ASTJOBM4.init_date)*24*60*60)/3600)*3600)/60), 0), 2, '0')"
        sql &= " || ':' ||"
        sql &= " lpad(nvl(round((((ASTJOBM4.last_date - ASTJOBM4.init_date)*24*60*60) -"
        sql &= " floor(((ASTJOBM4.last_date - ASTJOBM4.init_date)*24*60*60)/3600)*3600 -"
        sql &= " (floor((((ASTJOBM4.last_date - ASTJOBM4.init_date)*24*60*60) -"
        sql &= " floor(((ASTJOBM4.last_date - ASTJOBM4.init_date)*24*60*60)/3600)*3600)/60)*60) )), 0), 2, '0')"
        sql &= " ELAPSED"
        sql &= " FROM ASTJOBM4"
        Fill_Records("ASTJOBM4", "", True, sql)
        Fill_Records("ASTJOBM4_C", "", True, sqlASTJOBM4_C)

        Clear_All_Filters(grdASTJOBM3)
        Clear_All_Filters(grdASTJOBM4)
        Sort_grdColumns(grdASTJOBM4, "JOB_STREAM_XNO,JOB_STREAM_CODE")

        JOB_STREAM_CODE = String.Empty
        REPORT_ID = String.Empty
        SET_ID = String.Empty
        JOB_STREAM_LNO = 0
        ExecutingScript = False

        txtJOB_STREAM_CODE.Clear()
        txtJOB_STREAM_DESC.Clear()
        rowASTJOBM1 = Nothing
        recurrenceChanged = False
        recurringEventDates.Clear()

        MyBase.EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Loading Job Stream")
        Call Save_Header_Fields(grpHeader)

        Dim sql As String = String.Empty

        MyBase.EnforceConstraints(False)

        JOB_STREAM_CODE = HFs("JOB_STREAM_CODE")

        If EntryMode = "N" Then
            rowASTJOBM1 = dst.Tables("ASTJOBM1").NewRow
            rowASTJOBM1.Item("JOB_STREAM_CODE") = HFs("JOB_STREAM_CODE")
            rowASTJOBM1.Item("JOB_STREAM_DESC") = HFs("JOB_STREAM_DESC")
            rowASTJOBM1.Item("JOB_STREAM_TYPE") = cmbJOB_STREAM_TYPE.Value
            rowASTJOBM1.Item("ACTIVE") = "1"
            rowASTJOBM1.Item("USER_ID") = ASCMAIN1.USER_ID
            dst.Tables("ASTJOBM1").Rows.Add(rowASTJOBM1)
            recurrenceChanged = False
        Else
            Fill_Records("ASTJOBM1", JOB_STREAM_CODE)

            sql = "Select ASTJOBM2.*, ASTMENU1.MENU_ITEM_DESC, ASTROPT1.SET_DESC "
            sql &= " from ASTJOBM2,ASTMENU1,ASTROPT1 "
            sql &= " where ASTJOBM2.JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'"
            sql &= " and ASTMENU1.MENU_ID (+) = ASTJOBM2.MENU_ID"
            sql &= " and ASTMENU1.MENU_ITEM_TYPE (+) = 'R'"
            sql &= " and ASTMENU1.MENU_ITEM_OBJECT (+) = ASTJOBM2.REPORT_ID"
            sql &= " and ASTROPT1.FORM_NAME (+) = ASTJOBM2.REPORT_ID"
            sql &= " and ASTROPT1.SET_ID (+) = ASTJOBM2.SET_ID"

            Fill_Records("ASTJOBM2", "", True, sql)
            Fill_Records("ASTJOBM5", JOB_STREAM_CODE)

            If EntryMode = "V" Then
                grdASTJOBM2.DisplayLayout.Bands(0).Columns("REPORT_ID").Style = UltraWinGrid.ColumnStyle.Default
                grdASTJOBM2.DisplayLayout.Bands(0).Columns("SET_ID").Style = UltraWinGrid.ColumnStyle.Default
                'grdASTJOBM2.DisplayLayout.Bands(0).Columns("REPORT_ID").ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.OnCellActivate
                'grdASTJOBM2.DisplayLayout.Bands(0).Columns("SET_ID").ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.OnCellActivate
            Else
                grdASTJOBM2.DisplayLayout.Bands(0).Columns("REPORT_ID").Style = UltraWinGrid.ColumnStyle.EditButton
                grdASTJOBM2.DisplayLayout.Bands(0).Columns("SET_ID").Style = UltraWinGrid.ColumnStyle.EditButton
            End If
        End If

        rowASTJOBM1 = dst.Tables("ASTJOBM1").Rows(0)


        If dst.Tables("ASTJOBM5").Rows.Count = 0 Then
            Dim rowASTJOBM5 As DataRow = dst.Tables("ASTJOBM5").NewRow
            rowASTJOBM5.Item("JOB_STREAM_CODE") = HFs("JOB_STREAM_CODE")
            rowASTJOBM5.Item("JOB_START_TIME") = CDate(DateTime.Now.ToShortDateString & " 8:00 AM")
            rowASTJOBM5.Item("JOB_RECURR_PATTERN") = DirectCast(UltraWinSchedule.RecurrencePatternFrequency.Daily, Integer)
            rowASTJOBM5.Item("JOB_RANGE_LIMIT") = DirectCast(UltraWinSchedule.RecurrenceRangeLimit.NoLimit, Integer)
            rowASTJOBM5.Item("JOB_RANGE_MAX_OCCUR") = 10
            rowASTJOBM5.Item("JOB_PATTERN_INTERVAL") = 1
            rowASTJOBM5.Item("JOB_RANGE_END_DATE") = DateTime.Now
            rowASTJOBM5.Item("JOB_PATTERN_DAY_OF_MONTH") = 1
            rowASTJOBM5.Item("JOB_PATTERN_DAYS_OF_WEEK") = DirectCast(UltraWinSchedule.RecurrencePatternDaysOfWeek.AllWeekdays, Integer)
            rowASTJOBM5.Item("JOB_PATTERN_INTERVAL") = 1
            rowASTJOBM5.Item("JOB_PATTERN_MONTH_OF_YEAR") = DateTime.Now.Month
            rowASTJOBM5.Item("JOB_OCCUR_OF_DAY_IN_MONTH") = UltraWinSchedule.RecurrencePatternOccurrenceOfDayInMonth.None
            rowASTJOBM5.Item("JOB_PATTERN_TYPE") = DirectCast(UltraWinSchedule.RecurrencePatternType.Calculated, Integer)
            dst.Tables("ASTJOBM5").Rows.Add(rowASTJOBM5)
        End If


        txtJOB_STREAM_DESC.Text = rowASTJOBM1.Item("JOB_STREAM_DESC")
        cmbJOB_STREAM_TYPE.Value = rowASTJOBM1.Item("JOB_STREAM_TYPE")

        MyBase.EnforceConstraints(True)

        Clear_All_Filters(grdASTJOBM2)
        Sort_grdColumns(grdASTJOBM2, "JOB_STREAM_LNO")

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()

            INIT_LAST("ASTJOBM1")
            Update_Record_TDA("ASTJOBM1", "JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'")
            Update_Record_TDA("ASTJOBM2", "JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'")
            Update_Record_TDA("ASTJOBM5", "JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'")

            ' Create the times the jobs will fire
            If recurringEventDates.Count > 0 Then
                For Each apptDate As Date In recurringEventDates
                    If apptDate > DateTime.Now Then
                        dst.Tables("ASTJOBM6").Rows.Add(New Object() {JOB_STREAM_CODE, apptDate, "0"})
                    End If
                Next
                Update_Record_TDA("ASTJOBM6", "JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "' AND TRUNC(JOB_START_TIME) >= TRUNC(SYSDATE) and NVL(PROCESSED, '0') = '0'")
            End If

            MyBase.CommitTrans("Update Complete")
        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub

    Private Sub Delete_Record()
        Try
            Dim sql As String = String.Empty

            MyBase.BeginTrans()

            For Each table As String In New String() {"ASTJOBM1", "ASTJOBM2", "ASTJOBM3", "ASTJOBM4", "ASTJOBM5", "ASTJOBM6"}
                sql = "Delete from " & table & " where JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'"
                ASCDATA1.ExecuteSQL(sql)
            Next
            MyBase.CommitTrans("Job Stream (" & JOB_STREAM_CODE & ") successfully deleted.")
        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdASTJOBM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTJOBM2.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "REPORT_ID"
                If e.Cell.Value & String.Empty <> REPORT_ID Then
                    REPORT_ID = e.Cell.Value
                    e.Cell.Row.Cells("SET_ID").Value = String.Empty
                    e.Cell.Row.Cells("SET_DESC").Value = String.Empty
                    e.Cell.Row.Cells("MENU_ITEM_DESC").Value = String.Empty
                    e.Cell.Row.Cells("MENU_ID").Value = String.Empty

                    If dst.Tables("ASTMENU1").Select("MENU_ITEM_OBJECT = '" & REPORT_ID & "'").Length > 0 Then
                        Dim row As DataRow = dst.Tables("ASTMENU1").Select("MENU_ITEM_OBJECT = '" & REPORT_ID & "'")(0)
                        e.Cell.Row.Cells("MENU_ITEM_DESC").Value = row.Item("MENU_ITEM_DESC")
                        e.Cell.Row.Cells("MENU_ID").Value = row.Item("MENU_ID")
                    End If
                End If

            Case "SET_ID"
                If e.Cell.Value & String.Empty <> SET_ID Then
                    SET_ID = e.Cell.Value
                    e.Cell.Row.Cells("SET_DESC").Value = String.Empty

                    If dst.Tables("ASTROPT1").Select("FORM_NAME = '" & REPORT_ID & "' AND SET_ID = '" & SET_ID & "'").Length > 0 Then
                        Dim row As DataRow = dst.Tables("ASTROPT1").Select("FORM_NAME = '" & REPORT_ID & "' AND SET_ID = '" & SET_ID & "'")(0)
                        e.Cell.Row.Cells("SET_DESC").Value = row.Item("SET_DESC")
                    End If
                End If
        End Select

    End Sub

    Private Sub grdASTJOBM2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTJOBM2.AfterExitEditMode

        If grdASTJOBM2.ActiveCell Is Nothing Then
            Exit Sub
        End If

        Select Case grdASTJOBM2.ActiveCell.Column.Key
            Case "REPORT_ID"
                grdASTJOBM2.ActiveCell.Value = grdASTJOBM2.ActiveCell.Text.ToUpper
            Case "SET_ID"
                grdASTJOBM2.ActiveCell.Value = grdASTJOBM2.ActiveCell.Text.ToUpper
        End Select
    End Sub

    Private Sub grdASTJOBM2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTJOBM2.AfterRowActivate
        REPORT_ID = grdASTJOBM2.ActiveRow.Cells("REPORT_ID").Text
        SET_ID = grdASTJOBM2.ActiveRow.Cells("SET_ID").Text
    End Sub

    Private Sub grdASTJOBM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTJOBM2.BeforeRowUpdate

        EMsg = String.Empty

        Dim REPORT_ID As String = e.Row.Cells("REPORT_ID").Value & String.Empty
        REPORT_ID = REPORT_ID.Trim.ToUpper

        Dim SET_ID As String = e.Row.Cells("SET_ID").Value & String.Empty
        SET_ID = SET_ID.Trim.ToUpper

        If REPORT_ID.Length = 0 Then
            EMsg &= vbCr & "Report ID is required."
        Else
            e.Row.Cells("REPORT_ID").Value = REPORT_ID
            If dst.Tables("ASTMENU1").Select("MENU_ITEM_OBJECT = '" & REPORT_ID & "'").Length = 0 Then
                EMsg &= vbCr & "Report ID is invalid."
            End If
        End If

        If SET_ID.Length = 0 Then
            EMsg &= vbCr & "Set ID is required."
        Else
            If dst.Tables("ASTROPT1").Select("SET_ID = '" & SET_ID & "' AND FORM_NAME = '" & REPORT_ID & "'").Length = 0 Then
                Fill_Records("ASTROPT1", String.Empty, True, "Select * from ASTROPT1")
                If dst.Tables("ASTROPT1").Select("SET_ID = '" & SET_ID & "' AND FORM_NAME = '" & REPORT_ID & "'").Length = 0 Then
                    EMsg &= vbCr & "Set ID is invalid for the selected Report ID."
                End If
            End If
        End If

        If EMsg.Length > 0 Then
            MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If Val(e.Row.Cells("JOB_STREAM_LNO").Value & String.Empty) = 0 Then
            e.Row.Cells("JOB_STREAM_LNO").Value = Val(dst.Tables("ASTJOBM2").Compute("MAX(JOB_STREAM_LNO)", "") & String.Empty) + 1
        End If

        e.Row.Cells("JOB_STREAM_CODE").Value = JOB_STREAM_CODE
    End Sub

    Private Sub grdASTJOBM2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTJOBM2.ClickCellButton

        Dim sql_where As String = String.Empty

        Select Case grdASTJOBM2.ActiveCell.Column.Key
            Case "REPORT_ID"
                sql_where = "MENU_ITEM_TYPE = 'R'"
                MyBase.grdClickCellButton(grdASTJOBM2, sql_where, False, "", "MENU_ITEM_OBJECT")

            Case "SET_ID"
                sql_where = "FORM_NAME = '" & e.Cell.Row.Cells("REPORT_ID").Text & "'"
                MyBase.grdClickCellButton(grdASTJOBM2, sql_where, False)
        End Select
    End Sub

    Private Sub grdASTJOBM4_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTJOBM4.AfterRowActivate

        Dim JOB_STREAM_XNO As String = String.Empty

        If grdASTJOBM4.ActiveRow IsNot Nothing Then
            JOB_STREAM_XNO = grdASTJOBM4.ActiveRow.Cells("JOB_STREAM_XNO").Text
        End If

        Dim view As DataView = New DataView(dst.Tables("ASTJOBM3"))
        view.RowFilter = "JOB_STREAM_XNO = '" & JOB_STREAM_XNO & "'"
        view.Sort = "JOB_STREAM_LNO"
        grdASTJOBM3.DataSource = view
    End Sub

    Private Sub timExecute_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timExecute.Tick
        timExecute.Stop()

        Dim sql As String = String.Empty
        Dim rowASTJOBM3 As DataRow = Nothing

        Try
            If JOB_STREAM_LNO > 0 Then

                ' Update Menu ID value
                If tblASTJOBM2.Select("JOB_STREAM_LNO = " & JOB_STREAM_LNO).Length > 0 Then
                    Dim rowASTJOBM2 As DataRow = tblASTJOBM2.Select("JOB_STREAM_LNO = " & JOB_STREAM_LNO)(0)
                    ' See if the report was successful.
                    sql = "Update ASTJOBM3 set MENU_ID = '" & rowASTJOBM2.Item("MENU_ID") & "'"
                    sql &= " Where JOB_STREAM_XNO = '" & JOB_STREAM_XNO & "'"
                    sql &= " And JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'"
                    sql &= " And JOB_STREAM_LNO = " & JOB_STREAM_LNO
                    ASCDATA1.ExecuteSQL(sql)
                End If

                ' See if the report was successful.
                sql = "Select * from ASTJOBM3"
                sql &= " Where JOB_STREAM_XNO = '" & JOB_STREAM_XNO & "'"
                sql &= " And JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'"
                sql &= " And JOB_STREAM_LNO = " & JOB_STREAM_LNO

                rowASTJOBM3 = ASCDATA1.GetDataRow(sql)
                Dim menuItemDesc As String = String.Empty
                menuItemDesc = rowASTJOBM3.Item("REPORT_ID")

                Dim errorMessage As String = String.Empty
                If rowASTJOBM3 Is Nothing OrElse rowASTJOBM3.Item("SUCCESS") & String.Empty <> "1" Then
                    errorMessage = "It appears the report (" & menuItemDesc & ") did not execute properly. The Job will terminate."
                End If

                If errorMessage.Length > 0 Then
                    MessageBox.Show(errorMessage, "Execute Job", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ' Done processing the job.
                    Me.Mode_Settings(False)
                    Exit Sub
                End If

            End If

            JOB_STREAM_LNO += 1
            For Each rowASTJOBM2 As DataRow In tblASTJOBM2.Select("JOB_STREAM_LNO >= " & JOB_STREAM_LNO, "JOB_STREAM_LNO")
                ASCMAIN1.JOB_STREAM_CODE = rowASTJOBM2.Item("JOB_STREAM_CODE") & String.Empty
                ASCMAIN1.JOB_STREAM_LNO = rowASTJOBM2.Item("JOB_STREAM_LNO") & String.Empty
                ASCMAIN1.JOB_STREAM_FORM_NAME = rowASTJOBM2.Item("REPORT_ID") & String.Empty
                ASCMAIN1.JOB_STREAM_XNO = JOB_STREAM_XNO

                JOB_STREAM_LNO = Val(rowASTJOBM2.Item("JOB_STREAM_LNO") & String.Empty)
                JOB_MENU_ITEM_OBJECT = rowASTJOBM2.Item("REPORT_ID") & String.Empty
                JOB_MENU_ID = rowASTJOBM2.Item("MENU_ID") & String.Empty
                JOB_MENU_ITEM_TYPE = "R"

                ASCMAIN1.Launch_Form(JOB_MENU_ITEM_OBJECT, JOB_MENU_ITEM_TYPE, JOB_MENU_ID)
                Exit Sub

            Next
            ' Done processing the job.
            Me.Mode_Settings(False)

        Catch ex As Exception
            ExecutingScript = False
            MessageBox.Show("Execution terminated for the folloiwng reason: " & ex.Message, _
                             "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Try
                BeginTrans()
                sql = "Update ASTJOBM4 set LAST_OPER = '" & ASCMAIN1.USER_ID & "', LAST_DATE = SYSDATE WHERE JOB_STREAM_XNO = '" & JOB_STREAM_XNO & "'"
                ASCDATA1.ExecuteSQL(sql)
                CommitTrans()
            Catch ex As Exception
                Rollback()
            End Try

        End Try
    End Sub

#End Region

#Region "Form Procedures"

    Private Function ReNumberDetails() As String

        Dim eMessage As String = String.Empty
        Try
            dst.Tables("ASTJOBM2").AcceptChanges()

            Dim JOB_STREAM_LNO As Int16 = 1
            For Each rowASTJOBM2 As DataRow In dst.Tables("ASTJOBM2").Select("", "JOB_STREAM_LNO", DataViewRowState.CurrentRows)
                rowASTJOBM2.Item("JOB_STREAM_LNO") = JOB_STREAM_LNO
                JOB_STREAM_LNO += 1
            Next

            Sort_grdColumns(grdASTJOBM2, "JOB_STREAM_LNO")

        Catch ex As Exception
            eMessage = ex.Message
        End Try

        Return eMessage

    End Function

    Private Function ValidateData() As String
        Dim eMessage As String = String.Empty

        Try
            MyBase.Absx1.txtFor("JOB_STREAM_CODE").Text = MyBase.Absx1.txtFor("JOB_STREAM_CODE").Text.ToUpper.Trim
            If MyBase.Absx1.txtFor("JOB_STREAM_CODE").TextLength = 0 Then
                eMessage &= vbCrLf & "Job Stream Code is required."
            End If

            MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text = MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text.ToUpper.Trim
            If MyBase.Absx1.txtFor("JOB_STREAM_DESC").TextLength = 0 Then
                eMessage &= vbCrLf & "Job Stream Description is required."
            Else
                MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text = StrConv(MyBase.Absx1.txtFor("JOB_STREAM_DESC").Text, VbStrConv.ProperCase)
            End If

            If cmbJOB_STREAM_TYPE.Value = String.Empty Then
                eMessage &= vbCrLf & "Job Stream Type is required."
            End If

            Dim numMessage As String = ReNumberDetails()
            If numMessage.Length > 0 Then
                eMessage &= vbCrLf & numMessage
            End If

            For Each rowASTJOBM2 As DataRow In dst.Tables("ASTJOBM2").Select

                rowASTJOBM2.Item("REPORT_ID") = (rowASTJOBM2.Item("REPORT_ID") & String.Empty).ToString.Trim.ToUpper
                Dim REPORT_ID As String = rowASTJOBM2.Item("REPORT_ID") & String.Empty

                rowASTJOBM2.Item("SET_ID") = (rowASTJOBM2.Item("SET_ID") & String.Empty).ToString.Trim.ToUpper
                Dim SET_ID As String = rowASTJOBM2.Item("SET_ID") & String.Empty

                If rowASTJOBM2.Item("REPORT_ID") & String.Empty = String.Empty Then
                    eMessage &= vbCr & "Report ID is required on Line " & rowASTJOBM2.Item("JOB_STREAM_LNO")
                Else
                    If dst.Tables("ASTMENU1").Select("MENU_ITEM_OBJECT = '" & rowASTJOBM2.Item("REPORT_ID") & "'").Length = 0 Then
                        eMessage &= vbCr & "Report ID is invalid on Line " & rowASTJOBM2.Item("JOB_STREAM_LNO")
                    End If
                End If

                If SET_ID.Length > 0 Then
                    If dst.Tables("ASTROPT1").Select("SET_ID = '" & SET_ID & "' AND FORM_NAME = '" & REPORT_ID & "'").Length = 0 Then
                        eMessage &= vbCr & "Set ID is invalid for the selected Report ID on Line " & rowASTJOBM2.Item("JOB_STREAM_LNO")
                    End If
                End If
            Next

            If eMessage.Length = 0 AndAlso dst.Tables("ASTJOBM2").Rows.Count = 0 Then
                eMessage &= vbCrLf & "Job Stream details are required."
            End If

        Catch ex As Exception
            eMessage = ex.Message
        End Try

        If eMessage.Length = 0 Then
            rowASTJOBM1.Item("JOB_STREAM_CODE") = txtJOB_STREAM_CODE.Text
            rowASTJOBM1.Item("JOB_STREAM_DESC") = txtJOB_STREAM_DESC.Text
            rowASTJOBM1.Item("JOB_STREAM_TYPE") = cmbJOB_STREAM_TYPE.Value
        End If

        Return eMessage

    End Function

    Private Sub ExecuteJobStream(ByVal JOB_STREAM_CODE As String)

        Try

            Dim sql As String = String.Empty
            Dim rowASTJOBM3 As DataRow = Nothing
            Dim REPORT_ID As String = String.Empty

            Dim rowASTJOBM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ASTJOBM1 WHERE JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'")
            tblASTJOBM2 = ASCDATA1.GetDataTable("SELECT * FROM ASTJOBM2 WHERE JOB_STREAM_CODE = '" & JOB_STREAM_CODE & "'", "ASTJOBM2")

            If rowASTJOBM1 Is Nothing OrElse tblASTJOBM2.Rows.Count = 0 Then
                MessageBox.Show("No reports to process.", "Execute Job Stream", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            dst.Tables("ASTJOBM3").Rows.Clear()
            dst.Tables("ASTJOBM4_C").Rows.Clear()
            dst.Tables("ASTJOBM4").Rows.Clear()
            JOB_STREAM_XNO = ASCMAIN1.Next_Control_No("ASTJOBM3.JOB_STREAM_XNO")

            Dim rowASTJOBM4 As DataRow = dst.Tables("ASTJOBM4").NewRow

            rowASTJOBM4.Item("JOB_STREAM_XNO") = JOB_STREAM_XNO
            rowASTJOBM4.Item("JOB_STREAM_CODE") = rowASTJOBM1.Item("JOB_STREAM_CODE")
            rowASTJOBM4.Item("JOB_STREAM_DESC") = rowASTJOBM1.Item("JOB_STREAM_DESC")
            rowASTJOBM4.Item("JOB_STREAM_TYPE") = rowASTJOBM1.Item("JOB_STREAM_TYPE")
            rowASTJOBM4.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowASTJOBM4.Item("INIT_DATE") = DateTime.Now
            dst.Tables("ASTJOBM4").Rows.Add(rowASTJOBM4)

            Try
                BeginTrans()
                Update_Record_TDA("ASTJOBM4")
                ' NEED FOR CONSISTENCY OF DATES
                sql = "UPDATE ASTJOBM4 SET INIT_DATE = SYSDATE WHERE JOB_STREAM_XNO = '" & JOB_STREAM_XNO & "'"
                ASCDATA1.ExecuteSQL(sql)
                CommitTrans()
                ExecutingScript = True
                timExecute.Start()

            Catch ex As Exception
                Rollback()
                MessageBox.Show("The following error occurred: " & ex.Message, "Execute Job", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Execute Job", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub DisplayRecurrence()

        Try
            Dim objASTAPPT1 As New TAC.ASCAPPT1
            Dim rowASTJOBM5 As DataRow = dst.Tables("ASTJOBM5").Rows(0)

            With objASTAPPT1
                .StartTime = rowASTJOBM5.Item("JOB_START_TIME")
                .EndTime = rowASTJOBM5.Item("JOB_START_TIME")
                .PatternFrequency = Val(rowASTJOBM5.Item("JOB_RECURR_PATTERN") & String.Empty)
                .RangeLimit = Val(rowASTJOBM5.Item("JOB_RANGE_LIMIT") & String.Empty)
                .RangeMaxOccurrences = Val(rowASTJOBM5.Item("JOB_RANGE_MAX_OCCUR") & String.Empty)
                .RangeEndDate = CDate(rowASTJOBM5.Item("JOB_RANGE_END_DATE") & String.Empty)
                .PatternOccurrenceOfDayInMonth = Val(rowASTJOBM5.Item("JOB_OCCUR_OF_DAY_IN_MONTH") & String.Empty)
                .PatternInterval = Val(rowASTJOBM5.Item("JOB_PATTERN_INTERVAL") & String.Empty)
                .PatternDayOfMonth = Val(rowASTJOBM5.Item("JOB_PATTERN_DAY_OF_MONTH") & String.Empty)
                .PatternType = Val(rowASTJOBM5.Item("JOB_PATTERN_TYPE") & String.Empty)
                .PatternDaysOfWeek = Val(rowASTJOBM5.Item("JOB_PATTERN_DAYS_OF_WEEK") & String.Empty)
                .PatternMonthOfYear = Val(rowASTJOBM5.Item("JOB_PATTERN_MONTH_OF_YEAR") & String.Empty)

                ' This is true if the User clicks the OK Button.
                If .DisplayAppointment(recurringEventDates) Then

                    If EntryMode = "E" Then
                        recurrenceChanged = False
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_START_TIME", DataRowVersion.Original) <> .StartTime
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_RECURR_PATTERN", DataRowVersion.Original) <> DirectCast(.PatternFrequency, Integer)
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_RANGE_LIMIT", DataRowVersion.Original) <> DirectCast(.RangeLimit, Integer)
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_RANGE_MAX_OCCUR", DataRowVersion.Original) <> .RangeMaxOccurrences
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_PATTERN_INTERVAL", DataRowVersion.Original) <> .PatternInterval
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_RANGE_END_DATE", DataRowVersion.Original) <> .RangeEndDate
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_PATTERN_DAY_OF_MONTH", DataRowVersion.Original) <> .PatternDayOfMonth
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_PATTERN_DAYS_OF_WEEK", DataRowVersion.Original) <> DirectCast(.PatternDaysOfWeek, Integer)
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_PATTERN_MONTH_OF_YEAR", DataRowVersion.Original) <> .PatternMonthOfYear
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_OCCUR_OF_DAY_IN_MONTH", DataRowVersion.Original) <> DirectCast(.PatternOccurrenceOfDayInMonth, Integer)
                        recurrenceChanged = recurrenceChanged OrElse rowASTJOBM5.Item("JOB_PATTERN_TYPE", DataRowVersion.Original) <> DirectCast(.PatternType, Integer)
                    ElseIf EntryMode = "N" Then
                        recurrenceChanged = True
                    End If

                    If Not recurrenceChanged Then
                        recurringEventDates.Clear()
                    End If

                    rowASTJOBM5.Item("JOB_START_TIME") = .StartTime
                    rowASTJOBM5.Item("JOB_RECURR_PATTERN") = DirectCast(.PatternFrequency, Integer)
                    rowASTJOBM5.Item("JOB_RANGE_LIMIT") = DirectCast(.RangeLimit, Integer)
                    rowASTJOBM5.Item("JOB_RANGE_MAX_OCCUR") = .RangeMaxOccurrences
                    rowASTJOBM5.Item("JOB_PATTERN_INTERVAL") = .PatternInterval
                    rowASTJOBM5.Item("JOB_RANGE_END_DATE") = .RangeEndDate
                    rowASTJOBM5.Item("JOB_PATTERN_DAY_OF_MONTH") = .PatternDayOfMonth
                    rowASTJOBM5.Item("JOB_PATTERN_DAYS_OF_WEEK") = DirectCast(.PatternDaysOfWeek, Integer)
                    rowASTJOBM5.Item("JOB_PATTERN_MONTH_OF_YEAR") = .PatternMonthOfYear
                    rowASTJOBM5.Item("JOB_OCCUR_OF_DAY_IN_MONTH") = DirectCast(.PatternOccurrenceOfDayInMonth, Integer)
                    rowASTJOBM5.Item("JOB_PATTERN_TYPE") = DirectCast(.PatternType, Integer)
                End If
            End With

        Catch ex As Exception
            MessageBox.Show("The following Error Occurred: " & ex.Message, "Appointment Recurrence", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#End Region

End Class