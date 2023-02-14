Imports System.Text.RegularExpressions

Public Class InputForm

    Public scannedData As String
    Public buttonPressed As String = ""
    Public serverResult As String
    Public procToCall As String
    Private formType As String

    Private MyAudioController As Symbol.Audio.Controller = Nothing


    Public Sub New(ByVal appName As String, ByVal options As Dictionary(Of String, String))
        InitializeComponent()
        'parse guiString to create graphics
        'we will always be displaying a text prompt

        CreateGUI(appName, options)
        'options("TEXT") = prompt text for label
        'options("TYPE") = dialog type -- either display an input textbox with an OK, or an OK/Cancel, or a simple OK
        'we will then expect either a scan, or an OK/Cancel (to verify/submit scanned data)
        'we may need a LIST option for waves
    End Sub

    Private Sub CreateGUI(ByVal appName As String, ByVal options As Dictionary(Of String, String))
        Dim appLabel As Label = New Label()
        appLabel.Text = appName
        appLabel.Width = 220
        Me.Controls.Add(appLabel)

        Dim messageLabel = New AutoSizeLabel()
        If options.ContainsKey("MESSAGE") Then
            'informational message at the top
            messageLabel.Height = 70
            messageLabel.Width = 220
            messageLabel.Top = 20
            messageLabel.Text = options("MESSAGE")
            Me.Controls.Add(messageLabel)
        End If

        Dim promptLabel = New AutoSizeLabel()
        promptLabel.Height = 30
        promptLabel.Top = If(options.ContainsKey("MESSAGE"), messageLabel.Height + 35, 25)
        promptLabel.Width = 220
        promptLabel.Text = options("TEXT")

        Me.Controls.Add(promptLabel)
        Dim promptHeight = promptLabel.Height

        Me.formType = options("TYPE")
        Select Case options("TYPE")
            Case "S"
                Dim txtScan = New TextBox()
                txtScan.Top = promptLabel.Top + promptLabel.Height + 10
                txtScan.Height = 15
                txtScan.Name = "txtScan"
                Me.Controls.Add(txtScan)

                Dim createdButtons As Boolean = False
                For i = 1 To 3
                    Dim btnX = "BTN" & i.ToString()
                    If options.ContainsKey(btnX) Then
                        'Create the button with the given text
                        Dim frmBtn = New Button()
                        frmBtn.Text = options(btnX)
                        frmBtn.Top = txtScan.Top + 26
                        frmBtn.Height = 25
                        frmBtn.Width = 60
                        frmBtn.Left = 10 + (i - 1) * 70
                        AddHandler frmBtn.Click, AddressOf btnOK_Click
                        Me.Controls.Add(frmBtn)
                        createdButtons = True
                    End If
                Next

                If Not createdButtons Then
                    Dim btnOK = New Button()
                    btnOK.Text = "OK"
                    btnOK.Top = txtScan.Top + 26
                    btnOK.Height = 25
                    AddHandler btnOK.Click, AddressOf btnOK_Click
                    Me.Controls.Add(btnOK)
                    Dim btnExit = New Button()
                    btnExit.Text = "Exit"
                    btnExit.Top = btnOK.Top
                    btnExit.Height = 25
                    btnExit.Left = 80
                    AddHandler btnExit.Click, AddressOf btnExit_Click
                    Me.Controls.Add(btnExit)
                    txtScan.Focus()
                End If
            Case "Y"

                Dim createdButtons As Boolean = False
                For i = 1 To 3
                    Dim btnX = "BTN" & i.ToString()
                    If options.ContainsKey(btnX) Then
                        'Create the button with the given text
                        Dim frmBtn = New Button()
                        frmBtn.Text = options(btnX)
                        frmBtn.Top = promptLabel.Top + promptLabel.Height + 25
                        frmBtn.Height = 25
                        frmBtn.Width = 60
                        frmBtn.Left = 10 + (i - 1) * 70
                        AddHandler frmBtn.Click, AddressOf btnOK_Click
                        Me.Controls.Add(frmBtn)
                        createdButtons = True
                    End If
                Next

                If Not createdButtons Then
                    Dim btnOK = New Button()
                    btnOK.Text = "OK"
                    btnOK.Height = 25
                    btnOK.Top = promptLabel.Top + promptLabel.Height + 25
                    AddHandler btnOK.Click, AddressOf btnOK_Click
                    Dim btnCancel = New Button()
                    btnCancel.Top = btnOK.Top
                    btnCancel.Height = 25
                    btnCancel.Left = 80
                    btnCancel.Text = "Exit"
                    AddHandler btnCancel.Click, AddressOf btnExit_Click
                    Me.Controls.Add(btnOK)
                    Me.Controls.Add(btnCancel)
                End If
        End Select

        Dim MyDevice As Symbol.Audio.Device = _
        CType(Symbol.StandardForms.SelectDevice.Select( _
        Symbol.Audio.Controller.Title, _
        Symbol.Audio.Device.AvailableDevices), Symbol.Audio.Device)
        MyAudioController = New Symbol.Audio.StandardAudio(MyDevice)

        If (options.ContainsKey("BEEP") AndAlso options("BEEP") = "ERROR") Or options.ContainsKey("ERROR") Then
            ErrorBeep()
        End If

    End Sub

    Private Function GetControlByName(ByVal ctrlName As String) As Control
        For Each c As Control In Me.Controls
            If (c.Name = ctrlName) Then
                Return c
            End If
        Next
        Return Nothing
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If GetControlByName("txtScan") IsNot Nothing Then
            Me.scannedData = GetControlByName("txtScan").Text & ""
        End If
        Me.buttonPressed = CType(sender, Button).Text

        Me.Close()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.scannedData = "CANCEL"
        Me.Close()
    End Sub

    Private Sub PressEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Chr(13) Then
            If GetControlByName("txtScan") IsNot Nothing Then
                Me.scannedData = GetControlByName("txtScan").Text & ""
            End If
            Me.Close()
        End If
    End Sub

    Private Sub Barcode1_OnRead(ByVal sender As System.Object, ByVal readerData As Symbol.Barcode.ReaderData) Handles Barcode1.OnRead
        GetControlByName("txtScan").Text = readerData.Text

        If Me.formType = "S" Then
            Me.scannedData = GetControlByName("txtScan").Text

            If Me.scannedData & "" = "" Then
                Exit Sub
            End If
        End If
        Me.buttonPressed = ""
        Me.Close()
    End Sub

    Private Sub InputForm_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Barcode1.EnableScanner = False
        Me.MyAudioController.Dispose()
    End Sub

    Private Sub ErrorBeep()
        'duration in ms, frequency in hz
        'Me.MyAudioController.PlayAudio(250, 200)

        Try 'prevent crash if Error in audio
            For i = 1 To 5
                Me.MyAudioController.PlayAudio(80, 800)
            Next
        Catch
            'eat error
        End Try
    End Sub

    Private Sub InputForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Barcode1.EnableScanner = True
    End Sub
End Class