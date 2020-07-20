Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Ipc
Imports System.Text.RegularExpressions
Public Class ASFLOGON
    Declare Function ProcessIdToSessionId Lib "kernel32.dll" (ByVal dwProcessId As Int32, ByRef pSessionId As Int32) As Int32

    Private LoginError As String = String.Empty

    Public Function GetSessionId() As Int32
        Dim _currentProcess As Process = Process.GetCurrentProcess()
        Dim _processID As Int32 = _currentProcess.Id
        Dim _sessionID As Int32
        Dim _result As Boolean = ProcessIdToSessionId(_processID, _sessionID)
        Return _sessionID
    End Function

    Public Interface ICommunicationService
        Sub Command(ByVal commandtext As String)
    End Interface

    Private Sub ASFLOGON_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If txtUSER_ID.Text <> "" Then
            txtUSER_PASSWORD.Focus()
        End If
    End Sub

    Private Sub ASFLOGON_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim z As String
        'UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.CLIENT_CODE & ".bmp")
        UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & "ABS" & ".bmp")

        If Len(My.Application.Info.AssemblyName) = 8 And My.Application.Info.AssemblyName <> "ASEMENU1" Then
            Me.Left = -10000
        Else
            Call ASCMAIN1.Center(Me)

            z = GetSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFLOGON.USER_ID")
            If z <> "" Then
                ASCMAIN1.USER_ID = z
                ASCMAIN1.DBS_COMPANY = GetSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFLOGON.DBS_COMPANY")
                ASCMAIN1.DBS_SERVER = GetSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFLOGON.DBS_SERVER")
            End If
        End If
        'MsgBox("2c")

        txtDBS_COMPANY.Text = ASCMAIN1.DBS_COMPANY
        txtDBS_SERVER.Text = ASCMAIN1.DBS_SERVER
        txtUSER_ID.Text = ASCMAIN1.USER_ID

        If txtUSER_ID.Text = "" Then
            txtUSER_ID.Focus()
        Else
            txtUSER_PASSWORD.Focus()
        End If

        z = My.Application.CommandLineArgs.ToString
        z = Mid$(z, InStr(z & " ", " ") + 1)

        If My.User.Name = "wjz" Then
            Try
                Send_Message_to_Splash_Screen()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        'MsgBox("2d")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        ASCMAIN1.USER_ID = ""
        Me.Close()
    End Sub

    Private Sub cmdLogOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLogOn.Click
        Call Logon_Attempt()
    End Sub

    Sub Kill_Ghost_Sessions()

        If Environment.GetCommandLineArgs.Count >= 6 Then
            If Environment.GetCommandLineArgs.ElementAt(4) = "JS" Then
                Exit Sub
            End If
        End If

        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            Exit Sub
        End If

        Dim Sql As String = "Select SID,SERIAL# FROM V$SESSION WHERE TERMINAL = " _
        & " (Select UserEnv('TERMINAL') from DUAL)" _
        & " and AUDSID <> " _
        & " (Select UserEnv('SESSIONID') from DUAL)" _
        & " and UPPER(USERNAME) = '" & UCase$(ASCMAIN1.DBS_COMPANY) & "'"
        If ASCMAIN1.Running_in_VS Then
            Sql = Sql & " and UPPER(PROGRAM) = '" & UCase(My.Application.Info.AssemblyName & ".vshost.EXE'")
        Else
            Sql = Sql & " and UPPER(PROGRAM) = '" & UCase(My.Application.Info.AssemblyName & ".EXE'")
        End If
        Sql = Sql & " and UPPER(OSUSER) = '" & UCase(My.User.Name) & "'"
        Sql = Sql & " and STATUS = 'INACTIVE'"

        Dim tbl As DataTable = ASCDATA1.GetDataTable(Sql)
        If tbl.Rows.Count > 0 Then
            If MsgBoxResult.Yes = MsgBox("Kill these other Sessions? ", vbYesNo + vbQuestion, "There are Other Sessions Logged into ABSolution from This Station") Then

                For Each row As DataRow In tbl.Rows
                    Sql = "ALTER SYSTEM KILL SESSION '" & row.Item(0) & "," & row.Item(1) & "'"
                    ASCDATA1.ExecuteSQL(Sql)
                Next
            End If
        End If

    End Sub

    Public Sub Setup_Automated_Login(USER_ID As String, DBS_COMPANY As String)
        txtUSER_ID.Text = USER_ID
        txtDBS_COMPANY.Text = DBS_COMPANY

        ASCMAIN1.oraCon.ConnectionString = "Data Source=" & "" & ";User ID=" & DBS_COMPANY & ";Password=" & DBS_COMPANY

        ' txtDBS_PASSWORD.Text = "ANE"
    End Sub

    Sub Logon_Attempt()

        Dim Environment_UserDomainName As String = Environment.UserDomainName
        Dim Environment_UserName As String = Environment.UserName

        LoginError = String.Empty

        If txtUSER_ID.Text = "" Then
            Exit Sub
        End If

        Try

            Me.Cursor = Cursors.WaitCursor

            txtUSER_ID.Text = txtUSER_ID.Text.ToLower.Trim
            txtDBS_COMPANY.Text = txtDBS_COMPANY.Text.ToUpper.Trim
            txtDBS_SERVER.Text = txtDBS_SERVER.Text.ToUpper.Trim

            Dim Logon As String
            Logon = "Y"

            ASCMAIN1.DBS_COMPANY = txtDBS_COMPANY.Text
            ASCMAIN1.DBS_PASSWORD = ASCMAIN1.DBS_COMPANY
            ASCMAIN1.DBS_SERVER = txtDBS_SERVER.Text

            lblStatus.ForeColor = Color.Blue
            lblStatus.Text = "Now Attempting to Log-On"
            lblStatus.Visible = True
            If Not ASCMAIN1.ABSWEB Then Application.DoEvents()

            ASCMAIN1.DBS_PASSWORD = ASCMAIN1.DBS_COMPANY
            LoginError = String.Empty
            If Not Logon_Attempt_Succeeded() Then
                LoginError = String.Empty
                ASCMAIN1.DBS_PASSWORD = Get_DBS_PASSWORD_from_Password_Service(txtDBS_SERVER.Text, txtDBS_COMPANY.Text)
                If ASCMAIN1.DBS_PASSWORD = "" OrElse Not Logon_Attempt_Succeeded() Then
                    ASCMAIN1.DBS_PASSWORD = txtDBS_PASSWORD.Text
                    If ASCMAIN1.DBS_PASSWORD = "" OrElse Not Logon_Attempt_Succeeded() Then
                        ASCMAIN1.DBS_PASSWORD = ASCMAIN1.DBS_COMPANY
                        If Not Logon_Attempt_Succeeded() Then
                            lblStatus.ForeColor = Color.Red
                            lblStatus.Text = "Invalid Connection Credentials"
                            Me.Cursor = Cursors.Default
                            Application.DoEvents()
                            If LoginError.Length > 0 Then
                                MessageBox.Show(LoginError, "Login", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If

            ASCMAIN1.sql = "Select * from ASTPARM1 where AS_PARM_KEY = 'Z'"
            Dim tblASTPARM1 As DataTable = ASCDATA1.GetDataTable
            ASCMAIN1.rowASTPARM1 = tblASTPARM1.Rows(0)
			Dim AS_PARM_USERDOMAIN AS String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_USERDOMAIN") & ""
			dim domainUserMatches as Boolean = (Environment_UserName.toLower() = txtUSER_ID.Text And Environment_UserDomainName.toLower() = AS_PARM_USERDOMAIN.toLower())
            Dim Use_Encryption As Boolean = False
            Dim MD5 As New ASCSCMD5

            'ASCMAIN1.sql = "Select * from ASTPARMP where AS_PARM_KEY = 'Z'"
            Dim tblASTPARMP As DataTable = ASCDATA1.GetDataTable("Select * from ASTPARMP where AS_PARM_KEY = 'Z'", "ASTPARMP")
            If tblASTPARMP.Rows.Count = 0 Then
                Logon = "P"
            Else
                If tblASTPARMP.Rows(0).Item("AS_PARM_PWD_ENCRYPTED").ToString = "1" Then
                    Use_Encryption = True
                End If

                ASCMAIN1.sql = "Select * from ASTUSER1 where USER_ID = :PARM1"
                Dim rowASTUSER1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {txtUSER_ID.Text})

                If rowASTUSER1 Is Nothing Then
                    Logon = "N"
                Else

                    If rowASTUSER1.Item("USER_STATUS").ToString <> "A" Then
                        Logon = "I"
                    Else

                        If Not domainUserMatches Then
                            If Use_Encryption Then
                                Dim ENCRYPTED As String = MD5.DigestStrToHexStr(txtUSER_PASSWORD.Text)
                                If rowASTUSER1.Item("USER_PASSWORD") & "" <> "" Then
                                    If rowASTUSER1.Item("USER_PASSWORD").ToString <> MD5.DigestStrToHexStr(txtUSER_PASSWORD.Text) Then
                                        If ASCMAIN1.Running_in_VS Then
                                        Else
                                            Logon = "N"
                                        End If
                                    End If
                                End If
                            Else
                                If rowASTUSER1.Item("USER_PASSWORD").ToString <> (txtUSER_PASSWORD.Text) Then
                                    If ASCMAIN1.Running_in_VS Then
                                    Else
                                        Logon = "N"
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If ASCMAIN1.SOLUTION = "SEA" Then
                        If rowASTUSER1.Item("USER_CODES").ToString <> "" And _
                           rowASTUSER1.Item("USER_CODES").ToString <> ASCMAIN1.DBS_COMPANY Then
                            Logon = "N"
                        End If
                    End If

                End If

                If Logon <> "N" Then
                    Dim USER_SUSPEND_DATE As String
                    USER_SUSPEND_DATE = rowASTUSER1.Item("USER_SUSPEND_DATE").ToString

                    ' PROBABLY SHOULD HAVE JUST CACHED ROWASTUSER1
                    ASCMAIN1.USER_CODES = rowASTUSER1.Item("USER_CODES").ToString
                    ASCMAIN1.USER_EMAIL = rowASTUSER1.Item("USER_EMAIL").ToString
                    ASCMAIN1.USER_NAME = rowASTUSER1.Item("USER_NAME").ToString

                    ASCMAIN1.USER_MENU_ITEM_OBJECT = rowASTUSER1.Item("USER_MENU_ITEM_OBJECT").ToString
                    Dim CLA As String = Command()

                    Dim ii As Integer = InStr(UCase(CLA), "/FORM=")
                    If ii <> 0 Then
                        If ASCMAIN1.USER_MENU_ITEM_OBJECT <> "" And ASCMAIN1.USER_MENU_ITEM_OBJECT <> Mid(CLA, ii + 6, 8) Then
                            MsgBox("You are not Set Up to run the Application Requested", MsgBoxStyle.OkOnly, "Security / Access Violation")
                            End
                        Else
                            ASCMAIN1.USER_MENU_ITEM_OBJECT = Mid(CLA, ii + 6, 8)
                        End If
                    End If

                    If IsDate(USER_SUSPEND_DATE) Then
                        ' Look for Suspension of Login
                        If DateDiff("D", USER_SUSPEND_DATE, Now()) >= 0 Then
                            Logon = "S"
                        End If
                    End If
                End If

                If Logon = "Y" And Not domainUserMatches Then
                    Dim AS_PARM_PWD_DAYS_EXPIRE As Integer
                    AS_PARM_PWD_DAYS_EXPIRE = 0
                    Dim USER_PASSWORD_LAST_DATE As String

                    AS_PARM_PWD_DAYS_EXPIRE = Val(tblASTPARMP.Rows(0)("AS_PARM_PWD_DAYS_EXPIRE").ToString)

                    USER_PASSWORD_LAST_DATE = rowASTUSER1.Item("USER_PASSWORD_LAST_DATE").ToString
                    If Not IsDate(USER_PASSWORD_LAST_DATE) Then
                        USER_PASSWORD_LAST_DATE = DateValue(Format$(Now, "MM/dd/yyyy"))
                    End If

                    If AS_PARM_PWD_DAYS_EXPIRE > 0 Then
                        If DateDiff("D", USER_PASSWORD_LAST_DATE, Now) > AS_PARM_PWD_DAYS_EXPIRE Then

                            MsgBox("Your password has expired and must be changed.", MsgBoxStyle.OkOnly, "Change Password")

                            ASCMAIN1.USER_ID = txtUSER_ID.Text
                            ASCMAIN1.USER_PASSWORD = txtUSER_PASSWORD.Text
                            ASCMAIN1.Message = ""
                            Dim PF As New ASFPWDC1
                            PF.ShowDialog()
                            PF.Dispose()

                            If ASCMAIN1.Message = "" Then
                                Logon = "F"
                            End If
                        End If
                    End If

                    If Logon = "Y" Then
                        If rowASTUSER1.Item("USER_ID").ToString <> "wjz" Then
                            If tblASTPARMP.Rows(0).Item("AS_PARM_DENY_LOGON_MSG").ToString <> "" Then
                                Logon = "D"
                            End If
                        End If
                    End If
                End If
            End If

            If Logon = "Y" Then
                Dim COMPANY_CODEs As New List(Of String)
                ASCMAIN1.sql = "Select COMPANY_CODE from ASTUSER4 where USER_ID = :PARM1"
                For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {ASCMAIN1.USER_ID}).Rows
                    COMPANY_CODEs.Add(row.Item("COMPANY_CODE"))
                Next
                If COMPANY_CODEs.Count <> 0 Then
                    If Not COMPANY_CODEs.Contains(ASCMAIN1.DBS_COMPANY) Then
                        Logon = "C"
                    End If
                End If
            End If

            If Logon <> "Y" Then
                lblStatus.ForeColor = Color.Red
                Select Case Logon
                    Case "P"
                        lblStatus.Text = "Invalid/Missing Password Settings"
                    Case "I"
                        lblStatus.Text = "Invalid User Status"
                    Case "N"
                        lblStatus.Text = "Invalid User ID and/or Password"
                    Case "S"
                        lblStatus.Text = "User ID Suspended - See Administrator"
                    Case "F"
                        lblStatus.Text = "Invalid Password Change"
                    Case "D"
                        lblStatus.Text = "Log-On Temporarily Denied: " & tblASTPARMP.Rows(0).Item("AS_PARM_DENY_LOGON_MSG").ToString
                    Case "C"
                        lblStatus.Text = "No Access to Company " & ASCMAIN1.DBS_COMPANY
                End Select
                Me.Cursor = Cursors.Default
                Application.DoEvents()
                Exit Sub
            End If

            ' ************************* LOG ON SUCCESSFUL ************************

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "Select GETDATE()"
            Else
                ASCMAIN1.sql = "Select SYSDATE from DUAL"
            End If
            ASCMAIN1.oraCmd.CommandText = ASCMAIN1.sql
            Dim db_datetime As Date = ASCMAIN1.oraCmd.ExecuteScalar
            ASCMAIN1.NowTSD = db_datetime.Subtract(Now)

            ASCMAIN1.USER_ID = txtUSER_ID.Text
            ASCMAIN1.USER_PASSWORD = txtUSER_PASSWORD.Text

            ASCMAIN1.sql = "Select * from ASTSECK1"
            ASCMAIN1.tblASTSECK1 = ASCDATA1.GetDataTable

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.USER_SECURITY_CODEs = ""
                ASCMAIN1.sql = "Select SECURITY_CODE from ASTUSER2 where USER_ID = '" & ASCMAIN1.USER_ID & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                    ASCMAIN1.USER_SECURITY_CODEs &= "," & row.Item(0)
                Next
                If ASCMAIN1.USER_SECURITY_CODEs <> "" Then
                    ASCMAIN1.USER_SECURITY_CODEs = ASCMAIN1.USER_SECURITY_CODEs.Substring(1)
                End If
            Else
                ASCMAIN1.sql = ASCMAIN1.Flattened_List("USER_ID", "SECURITY_CODE", "ASTUSER2", ",", "USER_ID = '" & ASCMAIN1.USER_ID & "'")
                Dim tblASTUSER2s As DataTable
                tblASTUSER2s = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                If tblASTUSER2s.Rows.Count = 0 Then
                    ASCMAIN1.USER_SECURITY_CODEs = ""
                Else
                    ASCMAIN1.USER_SECURITY_CODEs = tblASTUSER2s.Rows(0).Item(1)
                End If
            End If

            Dim USER_GROUP_ID As String

            ASCMAIN1.sql = "Select USER_GROUP_ID from ASTUSER3 where USER_ID = :PARM1"
            For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {ASCMAIN1.USER_ID}).Rows
                USER_GROUP_ID = row.Item("USER_GROUP_ID")

                ASCMAIN1.sql = ASCMAIN1.Flattened_List("USER_ID", "SECURITY_CODE", "ASTUSER2", ",", "USER_ID = '" & USER_GROUP_ID & "'")
                Dim tblASTUSER2g As DataTable
                tblASTUSER2g = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                If tblASTUSER2g.Rows.Count > 0 Then

                    For Each SECURITY_CODE As String In tblASTUSER2g.Rows(0).Item(1).Split(",")
                        If Not ASCMAIN1.USER_SECURITY_CODEs.Split(",").Contains(SECURITY_CODE) Then
                            ASCMAIN1.USER_SECURITY_CODEs &= "," & SECURITY_CODE
                        End If
                    Next

                End If
            Next

            ASCMAIN1.SESSION_NO = ASCMAIN1.Next_Control_No("ASTLOGS1.SESSION_NO")
            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.DBS_SESSION_ID = 1
            Else
                Dim rowSession As DataRow = ASCDATA1.GetDataRow("Select UserEnv('SESSIONID'), UserEnv('TERMINAL') from DUAL")
                ASCMAIN1.DBS_SESSION_ID = rowSession.Item(0)
            End If
            ASCMAIN1.COMPUTER_NAME = My.Computer.Name

            SaveSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFLOGON.USER_ID", ASCMAIN1.USER_ID)
            SaveSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFLOGON.DBS_COMPANY", ASCMAIN1.DBS_COMPANY)
            SaveSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFLOGON.DBS_SERVER", ASCMAIN1.DBS_SERVER)

            Kill_Ghost_Sessions()
            ASCMAIN1.Get_Current_YP()

            '         ASCMAIN1.sql = "Select * from ASTPARM1 where AS_PARM_KEY = 'Z'"
            '        Dim tblASTPARM1 As DataTable = ASCDATA1.GetDataTable
            '        ASCMAIN1.rowASTPARM1 = tblASTPARM1.Rows(0)

            ASCMAIN1.CLIENT = ASCMAIN1.SOLUTION
            ASCMAIN1.CLIENT = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SCHEMA_PWD")

            If Not ASCMAIN1.Running_in_VS Then
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then ' necessary because V1 uses G: drive in ASTPARM1
                    ASCMAIN1.Folders("Archive") = "R:\VDI\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\"
                    ASCMAIN1.Folders("Attach") = "R:\VDI\ATTACH\" & ASCMAIN1.DBS_COMPANY & "\"
                    If Not ASCMAIN1.Running_in_VS Then
                        ASCMAIN1.Folders("Reports") = "R:\VDI\REPORTS\"
                    End If
                    ASCMAIN1.Folders("SharedRoot") = "R:\VDI\"
                Else
                    ASCMAIN1.Folders("Archive") = ASCMAIN1.rowASTPARM1("AS_PARM_ARCHIVE_FOLDER") & "\" & ASCMAIN1.DBS_COMPANY & "\"
                    ASCMAIN1.Folders("Attach") = ASCMAIN1.rowASTPARM1("AS_PARM_ATTACHMENT_FOLDER") & "\" & ASCMAIN1.DBS_COMPANY & "\"
                    ASCMAIN1.Folders("SharedRoot") = ASCMAIN1.rowASTPARM1("AS_PARM_SHARED_ROOT_FOLDER") & "\"
                End If

                If ASCMAIN1.DBS_SERVER = "ANE" Or ASCMAIN1.DBS_COMPANY = "ANE" Then ' necessary because V1 uses G: drive in ASTPARM1
                    ASCMAIN1.Folders("Archive") = "G:\EXP\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\"
                    ASCMAIN1.Folders("Attach") = "G:\EXP\ATTACH\" & ASCMAIN1.DBS_COMPANY & "\"
                    If Not ASCMAIN1.Running_in_VS Then
                        ASCMAIN1.Folders("Reports") = "G:\EXP\REPORTS\"
                    End If
                    ASCMAIN1.Folders("SharedRoot") = "G:\EXP\"

                End If
            Else
                ASCMAIN1.Folders("Archive") = ASCMAIN1.Folders("Archive") & ASCMAIN1.DBS_COMPANY & "\"
                ASCMAIN1.Folders("Attach") = ASCMAIN1.Folders("Attach") & ASCMAIN1.DBS_COMPANY & "\"
                ASCMAIN1.Folders("SharedRoot") = ASCMAIN1.rowASTPARM1("AS_PARM_SHARED_ROOT_FOLDER") & "\"
            End If
            If ASCMAIN1.Folders("SharedRoot") = "\" Then
                ASCMAIN1.Folders("SharedRoot") = "S:\" & ASCMAIN1.CLIENT & "\"
            End If

            ASCMAIN1.tblASTFFMT1 = ASCDATA1.GetDataTable("*", "ASTFFMT1")

            If Not ASCMAIN1.ABSWEB Then
                lblStatus.Text = "Log-On Successful"
                Me.Cursor = Cursors.Default
                Application.DoEvents()
            End If

            Call ASCMAIN1.Temp_Table_Cleanup()

            Dim INIT_DATE As Date = Now + ASCMAIN1.NowTSD

            Dim tblASTOPST1 As New DataTable
            With ASCDATA1.GetDataAdapter(tblASTOPST1, "ASTOPST1", "*", True, -1, False)
                Dim rowASTOPST1 As DataRow = tblASTOPST1.NewRow
                rowASTOPST1.Item("USER_ID") = ASCMAIN1.USER_ID
                rowASTOPST1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowASTOPST1.Item("INIT_DATE") = INIT_DATE
                rowASTOPST1.Item("YYYYPP") = ASCMAIN1.CYP
                rowASTOPST1.Item("SELECTION_NO") = 0
                rowASTOPST1.Item("RE_XNO") = 0
                rowASTOPST1.Item("PRD_CLOSE_IND") = ASCMAIN1.EOM
                rowASTOPST1.Item("FORM_INSTANCE_NO") = ASCMAIN1.Next_Control_No("ASFLOGON.FORM_INSTANCE_NO")
                tblASTOPST1.Rows.Add(rowASTOPST1)
                .Update(tblASTOPST1)
                .Dispose()
            End With

            Dim tblASTLOGS1 As New DataTable
            With ASCDATA1.GetDataAdapter(tblASTLOGS1, "ASTLOGS1", "*", True, -1, False)
                Dim rowASTLOGS1 As DataRow = tblASTLOGS1.NewRow
                rowASTLOGS1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowASTLOGS1.Item("USER_ID") = ASCMAIN1.USER_ID
                rowASTLOGS1.Item("SESSION_ID") = ASCMAIN1.DBS_SESSION_ID
                rowASTLOGS1.Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
                rowASTLOGS1.Item("DATE_LOGGED_ON") = INIT_DATE
                rowASTLOGS1.Item("SESSION_STATUS") = "A"
                tblASTLOGS1.Rows.Add(rowASTLOGS1)
                .Update(tblASTLOGS1)
                .Dispose()
            End With


            ' WTS Session ID

            ASCMAIN1.WTS_SESSION_ID = GetSessionId()
            ASCMAIN1.EncryptionKey = "0ff1c3" & ASCMAIN1.DBS_COMPANY

            If ASCMAIN1.DBS_COMPANY = "TST" Or ASCMAIN1.USER_ID = "TST" Then
                MsgBox("You are in the Test Company", vbOKOnly, "Please Note")
            End If

            If ASCMAIN1.ABSWEB Then
                ' NORMALLY CLALED FROM ASFMAIN1.FORM_SHOWN

                ASCMAIN1.Load_Views()
                ASCMAIN1.Load_MRUs()
            End If

            Me.Cursor = Cursors.Default
            Me.Close()

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


    End Sub

    Function Logon_Attempt_Succeeded() As Boolean
        Logon_Attempt_Succeeded = False

        If ASCMAIN1.DBS_PASSWORD <> "" Then
            Try
                If ASCMAIN1.oraCon.State = ConnectionState.Open Then
                    ASCMAIN1.oraCon.Close()
                End If

                Dim DEVELOPMENT_MACHINE_TNS As String = "(DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = orcl)))"
                'DEVELOPMENT_MACHINE_TNS = ""

                If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    ASCMAIN1.oraCon.ConnectionString = "Data Source=" & IIf(ASCMAIN1.DBS_SERVER = "", ".", ASCMAIN1.DBS_SERVER) & ";Initial Catalog=" & ASCMAIN1.DBS_COMPANY & "; " & IIf(ASCMAIN1.DBS_SERVER = "", "User ID='ODG'", "User ID='sa';Password='0ff1c3';") & ";Integrated Security=" & IIf(ASCMAIN1.DBS_SERVER = "", "True", "False") & ";MultipleActiveResultSets=True"
                Else
                    ASCMAIN1.oraCon.ConnectionString = "Data Source=" & IIf(ASCMAIN1.DBS_SERVER = "", DEVELOPMENT_MACHINE_TNS, ASCMAIN1.DBS_SERVER) & ";User ID=" & ASCMAIN1.DBS_COMPANY & ";Password=" & ASCMAIN1.DBS_PASSWORD & ";pooling=false"
                End If

                ASCMAIN1.oraCon.Open()
                ASCMAIN1.oraCmd = ASCMAIN1.oraCon.CreateCommand
                If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    'ASCMAIN1.oraCmd.CommandText = "Set Transaction Isolation Level Snapshot"
                    'ASCMAIN1.oraCmd.ExecuteNonQuery()
                End If

                ASCMAIN1.oraSP.CommandType = CommandType.StoredProcedure
                ASCMAIN1.oraSP.Connection = ASCMAIN1.oraCon

                Logon_Attempt_Succeeded = True

                'Dim myIpaddress As System.Net.IPAddress
                'Dim strhost As String
                Dim myWorkstation As String = System.Net.Dns.GetHostName()
                Dim IPAddress As String = _
                System.Net.Dns.GetHostEntry(myWorkstation).AddressList(0).ToString()
                ASCMAIN1.DBS_IP_ADDRESS = IPAddress
                ASCMAIN1.DBS_SERVER_NAME = myWorkstation

                'ASCMAIN1.DBS_IP_ADDRESS = ASCDATA1.GetDataValue("Select UTL_INADDR.GET_HOST_ADDRESS FROM DUAL")
                'ASCMAIN1.DBS_SERVER_NAME = ASCDATA1.GetDataValue("Select UTL_INADDR.GET_HOST_NAME FROM DUAL")
            Catch ex As Exception
                loginError = $"ABSolution Login Error: {ex.Message}"
                ' message below reveals the password
                'MsgBox(ex.Message & vbCr & ASCMAIN1.oraCon.ConnectionString)
            End Try
        End If

    End Function

    ''' <summary>
    ''' Get Oracle Server IP from TNSNAMES.ORA to issue a request for the password from a service running on the Oracle Server.
    ''' </summary>
    ''' <returns>Oracle Server Password for Company (Schema) and Server (SID) specified</returns>
    ''' <remarks>This process keeps the password encrypted on the server.  The password may be chnaged using the service.</remarks>
    Public Function Get_DBS_PASSWORD_from_Password_Service(ByVal DBS_SERVER As String, ByVal DBS_COMPANY As String) As String

        Get_DBS_PASSWORD_from_Password_Service = ""


        Try
            Dim PWD_HOST As String = String.Empty
            Dim PWD_PORT As Long = 0

            GetHostAndPortFromConfig(DBS_SERVER, PWD_HOST, PWD_PORT)

            Dim c As New System.Net.Sockets.Socket(
                       Net.Sockets.AddressFamily.InterNetwork,
                       Net.Sockets.SocketType.Stream,
                       Net.Sockets.ProtocolType.Tcp)

            c.SendTimeout = 3000
            c.ReceiveTimeout = 3000
            c.Connect(PWD_HOST, PWD_PORT)

            Dim request As String = "PROCURE " & DBS_COMPANY

            Dim BytesToSend() As Byte =
                        System.Text.ASCIIEncoding.ASCII.GetBytes(request)

            c.Send(BytesToSend, BytesToSend.Length, Net.Sockets.SocketFlags.None)

            Dim BytesToReceive(100) As Byte
            c.ReceiveTimeout = 3000
            Dim length As Int16 = 0

            Try
                length = c.Receive(BytesToReceive)
            Catch ex As Exception
            End Try

            Dim password As String = String.Empty
            password = System.Text.ASCIIEncoding.ASCII.GetString(BytesToReceive).ToString

            If length > 0 Then
                password = password.Substring(0, length)
            Else
                password = String.Empty
            End If

            Try
                c.Shutdown(Net.Sockets.SocketShutdown.Both)
                c.Close()
            Catch ex As Exception

            End Try

            Get_DBS_PASSWORD_from_Password_Service = password

        Catch ex As Exception

        End Try

    End Function

    Sub Send_Message_to_Splash_Screen()
        Dim ipcCh As New IpcChannel("myClient")
        ChannelServices.RegisterChannel(ipcCh, False)

        Dim obj As ICommunicationService =
          DirectCast(Activator.GetObject(GetType(ICommunicationService),
          "ipc://IPChannelName/SreeniRemoteObj"), ICommunicationService)
        obj.Command("Close")

        ChannelServices.UnregisterChannel(ipcCh)
    End Sub

    Private Sub ASFLOGON_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        If Environment.GetCommandLineArgs.Count >= 4 Then

            Dim company As String = Environment.GetCommandLineArgs.ElementAt(1)

            If txtDBS_COMPANY.Text.Length > 0 AndAlso txtDBS_COMPANY.Text.ToUpper = company.ToUpper Then
                txtUSER_ID.Text = Environment.GetCommandLineArgs.ElementAt(2)
                txtUSER_PASSWORD.Text = Environment.GetCommandLineArgs.ElementAt(3)
                Me.cmdLogOn_Click(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub GetHostAndPortFromConfig(ByVal DBS_SERVER As String, ByRef PWD_HOST As String, ByRef PWD_PORT As Int16)

        PWD_PORT = 4444
        PWD_HOST = String.Empty

        Try
            Dim appPath As String = Application.StartupPath
            appPath &= "\" & "" & "tnsnames.ora"

            If My.Computer.FileSystem.FileExists(appPath) Then
                Dim tnsNamesText = IO.File.ReadAllText(appPath)
                Dim tnsPattern = $"\n\s*{DBS_SERVER}\s*=.*?HOST\s*=\s*([^()]+)\)"
                Dim tnsMatch As Match = Regex.Match(tnsNamesText, tnsPattern, RegexOptions.Singleline)
                PWD_HOST = tnsMatch.Groups(1).Value

            End If
        Catch ex As Exception

        End Try


        Try
            Dim cp As System.Configuration.SettingsPropertyCollection = My.Settings.Properties
            For Each cc As System.Configuration.SettingsProperty In cp
                Select Case cc.Name
                    Case "PasswordServiceHost"
                        PWD_HOST = My.Settings.PropertyValues("PasswordServiceHost").Property.DefaultValue ' Hostname or IP Address of Oracle Password Server
                    Case "PasswordServicePort"
                        PWD_PORT = Val(My.Settings.PropertyValues("PasswordServicePort").Property.DefaultValue & String.Empty) ' Port for Database Password Request
                End Select
            Next
        Catch ex As Exception

        End Try

    End Sub

End Class