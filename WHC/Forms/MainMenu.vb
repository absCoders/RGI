Imports System.Net.Sockets
Imports System.IO
Imports System.Net
Imports System.Threading

Public Class MainMenu

    Private gunID As String
    Private userID As String
    Private sessionNo As String
    Private gunIPaddress As String

    Public serverIP As String
    Public serverPort As String
    Public dbSchema As String = ""
    Private serverConnection As ServerConnection

    Private menuItems As New Dictionary(Of String, String)
    Private gunEnvironmentVariables As New Dictionary(Of String, String)()

    Public Shared Sub Main()
        SingleInstance.Run(New MainMenu())
    End Sub

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Try
            Dim ipTextPath = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\server.txt"
            Dim serverFile = New StreamReader(New FileStream(ipTextPath, FileMode.Open, FileAccess.Read))
            serverIP = serverFile.ReadLine()
            serverPort = serverFile.ReadLine()
            If Not serverFile.EndOfStream Then
                dbSchema = serverFile.ReadLine()
            End If

            If serverIP & "" = "" Or serverPort & "" = "" Then
                MsgBox("Server IP or port not found" & vbCrLf & "Check that server.txt is in the directory with this app and contains IP on line 1 and port on line 2")
            End If
        Catch ex As Exception
            MsgBox(String.Format("Error loading server IP and port: {0}" & vbCrLf & "Check that server.txt is in the directory with this app and contains IP on line 1 and port on line 2", ex.Message))
            AddHandler Load, Function(s, e) CloseMe()
            'server.txt not found, use default serverIP
        End Try

        gunID = Dns.GetHostName()
        Dim ipEntry As IPHostEntry = Dns.GetHostEntry(gunID)
        gunIPaddress = ipEntry.AddressList(0).ToString()

        gunEnvironmentVariables.Add("GUNID", gunID) 'get gun ID
        gunEnvironmentVariables.Add("GUNIP", gunIPaddress) 'get gun IP address
        If Not String.IsNullOrEmpty(dbSchema) Then
            gunEnvironmentVariables.Add("SCHEMA", dbSchema)
        End If

        Try
            serverConnection = New ServerConnection(serverIP, serverPort)
        Catch
            AddHandler Load, Function(s, e) CloseMe()
            Exit Sub
        End Try

        Dim menuDict As Dictionary(Of String, String)
        Do
            menuDict = LoginForm.Login(serverConnection, gunEnvironmentVariables)

            If (menuDict.ContainsKey("ERROR")) Then
                MsgBox("Invalid Login -- try again")
            ElseIf menuDict.ContainsKey("EXIT") Then
                Application.Exit()
            End If
        Loop While (menuDict.ContainsKey("ERROR"))

        If Not menuDict.ContainsKey("EXIT") Then
            generateMenuGUI(menuDict)
        End If



    End Sub

    Private Function CloseMe()
        Application.Exit()
        Return 0
    End Function

    Private Sub generateMenuGUI(ByVal menuDict As Dictionary(Of String, String))

        Dim count As Integer = 0

        For Each menuItem In menuDict.Keys
            If menuItem.StartsWith("MENU") Then
                Dim appID As String = menuItem.Substring(4)
                menuItems.Add(menuDict(menuItem), appID)

                Dim menuLabel = New LinkLabel()
                menuLabel.Text = menuDict(menuItem)
                menuLabel.Top = (count * 30) + 10
                menuLabel.Left = 10
                menuLabel.Width = 220
                AddHandler menuLabel.Click, AddressOf menuItemClicked
                Me.Controls.Add(menuLabel)
                count += 1
            End If
        Next

    End Sub

    Private Sub menuItemClicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim clickedLabel = CType(sender, LinkLabel)
        Dim appSeqNo = 0
        Dim appID = menuItems(clickedLabel.Text)
        Dim appName = clickedLabel.Text

        If Not serverConnection.connected Then
            Try
                serverConnection.connectToServer(True, gunEnvironmentVariables)
                serverConnection.readFromServer() 'resends the menu
            Catch ex As Exception
                MsgBox("Error connecting... try again")
                serverConnection.connected = False
                Exit Sub
            End Try
        End If

        Dim dialogString
        Try
            'tell server which app we are starting up
            Dim serverOptions As New Dictionary(Of String, String)
            serverOptions.Add("APPID", appID)
            serverOptions.Add("STATE", 0)

            serverConnection.sendMessageToServer(serverOptions)

            'server will create state and tell us what to display
            dialogString = serverConnection.readFromServer()
        Catch
            MsgBox("Connection error... try again...")
            serverConnection.connected = False
            Exit Sub
        End Try


        While (dialogString <> "")

            Dim options = Util.parseOptions(dialogString)

            If options.ContainsKey("EXIT") Or (options.ContainsKey("STATE") AndAlso options("STATE") = "EXIT") Then
                Exit While
            End If

            Dim inputForm As InputForm = New InputForm(appName, options)

            If options.ContainsKey("ERROR") Then
                'MessageBox.Show(options("ERROR"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
                MessageBoxForm.Show(Me, options("ERROR"), "Error")
                'errorForm.ShowDialog()
            End If

            inputForm.ShowDialog()

            If inputForm.scannedData = "CANCEL" Then
                Exit While
            End If

            Dim sendOptions As New Dictionary(Of String, String)
            sendOptions.Add("APPID", appID)
            sendOptions.Add("STATE", options("STATE"))
            sendOptions.Add("SCAN", inputForm.scannedData)
            If inputForm.buttonPressed & "" <> "" Then
                sendOptions.Add("BTN", inputForm.buttonPressed)
            End If
            'we will need to also send the button that was pressed (for the class)

            Try
                serverConnection.sendMessageToServer(sendOptions)
                dialogString = serverConnection.readFromServer() 'final response is simply an ACK? what about invalid?
                'MsgBox(dialogString)
            Catch ex As Exception
                MsgBox("Connection error -- try again") 'could also be a database error... need to differentiate
                serverConnection.connected = False
                Exit Sub
            End Try

            inputForm.Close()
        End While
    End Sub


    Private Sub MainMenu_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress

        Select Case e.KeyChar
            Case "1"
                'linkLocations_Click(Me, Nothing)
        End Select
    End Sub

    Private Sub MainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'SingleInstance.HideTaskBar(Me.Handle)

    End Sub

    Private Sub MainMenu_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        'Me.BringToFront()
        'SingleInstance.ShowTaskBar(Me.Handle)
    End Sub

    Private Sub MainMenu_Closing(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'Ask for PASSWORD
        'If InputForm.GetInput("Enter password:", "Password required") <> "ane" Then
        '    e.Cancel = True
        'End If
    End Sub

End Class


