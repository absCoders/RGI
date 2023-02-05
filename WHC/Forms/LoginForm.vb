Public Class LoginForm

    Private _serverConnection As ServerConnection
    Private _gunEnvironmentVariables As Dictionary(Of String, String)

    Public _loginResults As New Dictionary(Of String, String)

    Public Shared Function Login(ByVal serverConnection As ServerConnection, ByRef gunEnvironmentVariables As Dictionary(Of String, String)) As Dictionary(Of String, String)
        Dim x As New LoginForm()
        x._serverConnection = serverConnection
        x._gunEnvironmentVariables = gunEnvironmentVariables
        x.ShowDialog()
        Return x._loginResults
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try

            Dim serverOptions As New Dictionary(Of String, String)
            serverOptions.Add("USER", txtUserID.Text)
            serverOptions.Add("PASS", txtPassword.Text)
            serverOptions.Add("GUNID", Me._gunEnvironmentVariables("GUNID"))
            serverOptions.Add("GUNIP", Me._gunEnvironmentVariables("GUNIP"))
            If _gunEnvironmentVariables.ContainsKey("SCHEMA") Then
                serverOptions.Add("SCHEMA", Me._gunEnvironmentVariables("SCHEMA"))
            End If


            _serverConnection.connectToServer(False, Nothing)
            _serverConnection.sendMessageToServer(serverOptions)
            Dim _loginResultMenu = _serverConnection.readFromServer()


            Dim loginResults = Util.parseOptions(_loginResultMenu)

            If loginResults.ContainsKey("ERROR") Then
                MsgBox(loginResults("ERROR"))
                Exit Sub
            Else
                _gunEnvironmentVariables.Add("USER", serverOptions("USER"))
                _gunEnvironmentVariables.Add("PASS", serverOptions("PASS"))
                _gunEnvironmentVariables.Add("SESSION", loginResults("SESSION"))

                _loginResults = loginResults
            End If
        Catch ex As Exception
            MsgBox("Connection error... wait a few seconds and try again")
            Exit Sub
        End Try
        Me.Close()
    End Sub

    Private Sub PressOK(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Chr(13) Then
            btnOK_Click(sender, e)
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        _loginResults.Add("EXIT", "EXIT")
        Application.Exit()
    End Sub

    'Private Sub txtInput_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUserID.KeyPress
    '    If e.KeyChar = Chr(13) Then 'Chr(13) is the Enter Key
    '        Me.Close()
    '    End If
    'End Sub
End Class