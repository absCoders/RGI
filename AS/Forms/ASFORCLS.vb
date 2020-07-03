Imports System.ComponentModel

Public Class ASFORCLS

#Region "Form Variables"

    Private Const RequestAllOracleUsertEntries As String = "*****"
    Private ReadOnly dt As New DataTable

    Private DBS_SERVER As String = String.Empty
    Private DBS_COMPANY As String = String.Empty
    Private clsPropertyPage As New PropertyPageClass

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dt
            .Columns.Add("USER", GetType(System.String))
            .Columns.Add("PASSWORD", GetType(System.String))
        End With
        grdUsers.DataSource = dt

        grdUsers.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdUsers.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdUsers.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

        DBS_SERVER = ASCMAIN1.DBS_SERVER
        DBS_COMPANY = ASCMAIN1.DBS_COMPANY

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Set Password"

            Case "Refresh"

            Case "Done"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Load_Record()
                Mode_Settings(True)

            Case "Set Password"
                SetPassword()

            Case "Refresh"
                LoadAllOracleUserEntries()

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Set Password").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        splMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
    End Sub

    Sub Load_Record()
        GetHostAndPortFromConfig()
        LoadAllOracleUserEntries()

        propertySheet.SelectedObject = clsPropertyPage
    End Sub

    Sub Delete_Record()

    End Sub

    Sub Update_Record()

        Try
            BeginTrans()

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Form Procedures"

    Private Sub SetPassword()
        Try
            txtUID.Text = txtUID.Text.Trim.ToUpper
            txtPWD.Text = txtPWD.Text.Trim
            txtPWDconf.Text = txtPWDconf.Text.Trim

            If txtUID.TextLength = 0 OrElse txtPWD.TextLength = 0 Then
                MessageBox.Show("The following are required: User, and Password.", "Set Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If txtPWD.Text <> txtPWDconf.Text Then
                MessageBox.Show("Password and Confirmation Password do Not match.", "Set Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Set_Password(txtUID.Text, txtPWD.Text)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Set Password", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            txtUID.Clear()
            txtPWD.Clear()
            txtPWDconf.Clear()
            LoadAllOracleUserEntries()
        End Try
    End Sub

    Public Function Get_DBS_PASSWORD_from_Password_Service(ByVal RequestAll As Boolean) As String

        Get_DBS_PASSWORD_from_Password_Service = String.Empty

        Try
            Using c As New System.Net.Sockets.Socket(
                       Net.Sockets.AddressFamily.InterNetwork,
                       Net.Sockets.SocketType.Stream,
                       Net.Sockets.ProtocolType.Tcp)

                c.SendTimeout = 5
                c.Connect(clsPropertyPage.PasswordServiceHostAddress, clsPropertyPage.PasswordServicePort)

                Dim request As String = String.Empty

                If Not RequestAll Then
                    request = "PROCURE " & DBS_SERVER & vbTab & DBS_COMPANY
                Else

                    request = "PROCURE " & RequestAllOracleUsertEntries
                End If

                Dim BytesToSend() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(request)
                c.Send(BytesToSend, BytesToSend.Length, Net.Sockets.SocketFlags.None)

                Dim BytesToReceive(2000) As Byte
                c.ReceiveTimeout = 5000
                Dim length As Int16 = c.Receive(BytesToReceive)
                Dim password As String = System.Text.ASCIIEncoding.ASCII.GetString(BytesToReceive).ToString
                password = password.Substring(0, length)

                Try
                    c.Shutdown(Net.Sockets.SocketShutdown.Both)
                    c.Close()
                Catch ex As Exception

                End Try

                Get_DBS_PASSWORD_from_Password_Service = password
            End Using

        Catch ex As Exception
            MessageBox.Show("Get DBS Password Service Error: " _
                            & ex.Message & Environment.NewLine & Environment.NewLine _
                            & "IP Address: " & clsPropertyPage.PasswordServiceHostAddress & Environment.NewLine _
                            & "Port: " & clsPropertyPage.PasswordServicePort, "Get DBS Password Service", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    Private Sub GetHostAndPortFromConfig()

        clsPropertyPage.PasswordServicePort = 4444
        clsPropertyPage.PasswordServiceHostAddress = String.Empty

        Try
            Dim appPath As String = Application.StartupPath
            appPath &= "\" & "" & "ABS.exe.Config"
            If Not My.Computer.FileSystem.FileExists(appPath) Then
                appPath = Application.StartupPath
                appPath &= "\" & "" & "ABS.Config"
            End If

            If My.Computer.FileSystem.FileExists(appPath) Then
                Dim dsConn As New DataSet
                ' Need to extract the status code from the 
                Using sr As New IO.StreamReader(appPath)
                    dsConn.ReadXml(sr)
                    sr.Close()
                    sr.Dispose()
                End Using

                If dsConn.Tables.Contains("dataSource") Then
                    If dsConn.Tables("dataSource").Select("Alias = '" & DBS_SERVER & "'").Length > 0 Then
                        Dim CONN As String = dsConn.Tables("dataSource").Select("Alias = '" & DBS_SERVER & "'")(0).Item("descriptor") & String.Empty
                        CONN = CONN.Replace(" ", "").ToUpper
                        CONN = CONN.Split("HOST=")(1).Split("=")(1).Split(")")(0)
                        clsPropertyPage.PasswordServiceHostAddress = CONN
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

        Try
            Dim cp As System.Configuration.SettingsPropertyCollection = My.Settings.Properties
            For Each cc As System.Configuration.SettingsProperty In cp
                Select Case cc.Name
                    Case "PasswordServiceHost"
                        clsPropertyPage.PasswordServiceHostAddress = My.Settings.PropertyValues("PasswordServiceHost").Property.DefaultValue ' Hostname or IP Address of Oracle Password Server
                    Case "PasswordServicePort"
                        clsPropertyPage.PasswordServicePort = Val(My.Settings.PropertyValues("PasswordServicePort").Property.DefaultValue & String.Empty) ' Port for Database Password Request
                End Select
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub LoadAllOracleUserEntries()

        Try
            dt.Rows.Clear()

            Dim entries As String = Get_DBS_PASSWORD_from_Password_Service(True)
            If entries.Length > 0 Then
                Dim oracleUsers() As String = entries.Split(",")
                For Each entry As String In oracleUsers
                    Dim pwds() As String = entry.Split(vbTab)
                    If pwds(0) & String.Empty <> String.Empty Then
                        Select Case pwds.Length
                            Case 0
                                ' Nothing 
                            Case 1
                                dt.Rows.Add(New Object() {pwds(0)})
                            Case 2
                                dt.Rows.Add(New Object() {pwds(0), pwds(1)})
                            Case Else
                                dt.Rows.Add(New Object() {pwds(0), pwds(1)})
                        End Select
                    End If
                Next
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Load All Oracle Users Entries", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Sort_grdColumns(grdUsers, "USER")
        End Try

    End Sub

    Private Sub Set_Password(ByVal UID As String,
                             ByVal Password As String)
        Try
            Using c As New System.Net.Sockets.Socket(
                       Net.Sockets.AddressFamily.InterNetwork,
                       Net.Sockets.SocketType.Stream,
                       Net.Sockets.ProtocolType.Tcp)


                c.SendTimeout = 5
                c.Connect(clsPropertyPage.PasswordServiceHostAddress, clsPropertyPage.PasswordServicePort)

                Dim request As String = "SET " & UID & vbTab & Password

                Dim BytesToSend() As Byte =
                            System.Text.ASCIIEncoding.ASCII.GetBytes(request)

                c.Send(BytesToSend, BytesToSend.Length, Net.Sockets.SocketFlags.None)

                Dim BytesToReceive(500) As Byte
                c.ReceiveTimeout = 5000

                Dim length As Int16 = c.Receive(BytesToReceive)
                Dim response As String = System.Text.ASCIIEncoding.ASCII.GetString(BytesToReceive).ToString
                response = response.Substring(0, length)

                MessageBox.Show("Response sent back from server: " & response, "Set Password", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Try
                    c.Shutdown(Net.Sockets.SocketShutdown.Both)
                    c.Close()
                Catch ex As Exception
                End Try

            End Using

        Catch ex As Exception
            MessageBox.Show("Set Password Error: " _
                            & ex.Message & Environment.NewLine & Environment.NewLine _
                            & "IP Address: " & clsPropertyPage.PasswordServiceHostAddress & Environment.NewLine _
                            & "Port: " & clsPropertyPage.PasswordServicePort, "Set Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try

    End Sub

#End Region

#Region "Form Classes"

    Private Class PropertyPageClass

        Dim m_PasswordServiceHostAddress As String = String.Empty
        Dim m_PasswordServicePort As Int32 = 4444

        <CategoryAttribute("Password Service"),
           Browsable(True),
           [ReadOnly](False),
           BindableAttribute(True),
           DefaultValueAttribute(""),
           DesignOnly(False),
           DescriptionAttribute("Host Address")>
        Public Property PasswordServiceHostAddress As String
            Get
                Return m_PasswordServiceHostAddress
            End Get

            Set(ByVal Value As String)
                m_PasswordServiceHostAddress = Value
            End Set
        End Property

        <CategoryAttribute("Password Service"),
           Browsable(True),
           [ReadOnly](False),
           BindableAttribute(True),
           DefaultValueAttribute(4444),
           DesignOnly(False),
           DescriptionAttribute("Host Port")>
        Public Property PasswordServicePort As Int32
            Get
                Return m_PasswordServicePort
            End Get

            Set(ByVal Value As Int32)
                m_PasswordServicePort = Value
            End Set
        End Property

    End Class

#End Region

End Class