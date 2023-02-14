Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.WindowsCE.Forms

Public Class NonFullscreenForm
    Inherits Form
    Private centered As Boolean = True

#Region "Native Platform Invoke"

    <DllImport("coredll.dll")> _
    Private Shared Function SetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As UInt32) As UInt32
    End Function

    <DllImport("aygshell.dll")> _
    Private Shared Function SHDoneButton(ByVal hwndRequester As IntPtr, ByVal dwState As UInt32) As Integer
    End Function

    Private ReadOnly GWL_STYLE As Integer = (-16)

    Private ReadOnly WS_CAPTION As UInt32 = &HC00000     ' WS_BORDER | WS_DLGFRAME  
    Private ReadOnly WS_BORDER As UInt32 = &H800000
    Private ReadOnly WS_POPUP As UInt32 = &H8000000

    Private ReadOnly SHDB_SHOW As UInt32 = &H1
    Private ReadOnly SHDB_HIDE As UInt32 = &H2

#End Region

    Public Sub New()
    End Sub

    Public Property CenterFormOnScreen() As Boolean
        Get
            Return centered
        End Get
        Set(ByVal value As Boolean)

            centered = value

            If (centered) Then
                CenterWithinScreen()
            End If
        End Set
    End Property

    Protected Overrides Sub OnLoad(ByVal e As EventArgs)

        ' By default if you set a form's size within
        ' the Visual Studio form designer it won't
        ' take into account the additional height of
        ' the caption, so we'll add that height here...
        Me.Height += SystemInformation.MenuHeight

        MyBase.OnLoad(e)

        ' Add the border and caption we removed from the form
        ' when we set the Form's FormBorderStyle property to None.
        ' We do this at the Win32 API level, which causes the .NET
        ' Compact Framework wrapper to get out of sync.
        Dim style As UInteger = WS_BORDER Or WS_CAPTION Or WS_POPUP

        SetWindowLong(Handle, GWL_STYLE, style)

        ' Add/Remove an [OK] button from the dialog's
        ' caption bar as required
        SHDoneButton(Handle, If(ControlBox, SHDB_SHOW, SHDB_HIDE))

        ' Center the form if requested
        If (centered) Then
            CenterWithinScreen()
        End If
    End Sub

    Protected Overrides Sub OnResize(ByVal e As EventArgs)

        MyBase.OnResize(e)

        ' If the dialog changes size and we want to be
        ' centered we may need to move the dialog to
        ' keep it centered.
        If (centered) Then
            CenterWithinScreen()
        End If
    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
        MyBase.OnKeyDown(e)

        ' If we have an [OK] button in the caption pressing
        ' Return or Escape should close the dialog
        If (Me.ControlBox) Then
            If (e.KeyCode = Keys.Return Or e.KeyCode = Keys.Escape) Then

                Me.DialogResult = DialogResult.OK
            End If
        End If
    End Sub

    Private Sub CenterWithinScreen()

        ' Move the position of this form to center it within the
        ' working area of the desktop
        Dim x As Integer = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Dim y As Integer = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2

        Me.Location = New Point(x, y)
    End Sub
End Class