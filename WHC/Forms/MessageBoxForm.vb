Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text

    Partial Public Class MessageBoxForm
        Inherits NonFullscreenForm
        Private MyAutoScaleFactor As SizeF
        Private beepType As MessageBeepType = MessageBeepType.[Default]

#Region "Native Platform Invoke"

        <DllImport("coredll.dll")> _
        Private Shared Function MessageBeep(ByVal uType As MessageBeepType) As Integer
        End Function

        <DllImport("coredll.dll")> _
        Private Shared Function DrawText(ByVal hDC As IntPtr, ByVal lpString As String, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal uFormat As UInteger) As Integer
        End Function

        <DllImport("coredll.dll")> _
        Private Shared Function SelectObject(ByVal hDC As IntPtr, ByVal hobj As IntPtr) As IntPtr
        End Function

        <DllImport("coredll.dll")> _
        Private Shared Function DeleteObject(ByVal HGDIOBJ As IntPtr) As Integer
        End Function

        Public Enum MessageBeepType As UInteger
            None = &HFFFFFFFEUI
            Simple = &HFFFFFFFFUI
            ' Simple beep
            [Default] = &H0
            ' MB_OK
            SystemHand = &H10
            ' MB_ICONHAND
            SystemQuestion = &H20
            ' MB_ICONQUESTION
            SystemExclamation = &H30
            ' MB_ICONEXCLAMATION
            SystemAsterisk = &H40
            ' MB_ICONASTERISK
        End Enum

        Private Structure RECT
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer

            Public Sub New(ByVal rc As Rectangle)
                Me.Left = rc.Left
                Me.Right = rc.Right
                Me.Top = rc.Top
                Me.Bottom = rc.Bottom
            End Sub
        End Structure

        Private Shared ReadOnly DT_LEFT As UInteger = &H0
        Private Shared ReadOnly DT_TOP As UInteger = &H0
        Private Shared ReadOnly DT_WORDBREAK As UInteger = &H10
        Private Shared ReadOnly DT_CALCRECT As UInteger = &H400
        Private Shared ReadOnly DT_NOPREFIX As UInteger = &H800

#End Region

        Private Sub New()
            InitializeComponent()

            ' We have a couple of coordinates specified within this
            ' class. The coordinates are calculated assuming a 96DPI
            ' screen. Calculate a scaling factor to translate these
            ' coordinates into the required values for the current
            ' screen's DPI.
            MyAutoScaleFactor = New SizeF(Me.AutoScaleDimensions.Width / 96.0F, Me.AutoScaleDimensions.Height / 96.0F)
        End Sub

        ''' <summary>
        ''' Displays a message box to the user with a blank caption and a single
        ''' [OK] button on the caption.
        ''' </summary>
        ''' <param name="owner">The form which is wanting to display this message box.</param>
        ''' <param name="message">The message to display in the client area of the message box.</param>
        ''' <param name="caption">The title to display in the caption of the message box.</param>
        ''' <returns>DialogResult.OK</returns>
        Public Shared Function Show(ByVal owner As Form, ByVal message As String) As DialogResult
            Return Show(owner, message, Nothing, Nothing)
        End Function

        ''' <summary>
        ''' Displays a message box to the user with a single [OK] button on the caption.
        ''' </summary>
        ''' <param name="owner">The form which is wanting to display this message box.</param>
        ''' <param name="message">The message to display in the client area of the message box.</param>
        ''' <param name="caption">The title to display in the caption of the message box.</param>
        ''' <returns>DialogResult.OK</returns>
        Public Shared Function Show(ByVal owner As Form, ByVal message As String, ByVal caption As String) As DialogResult
            Return Show(owner, message, caption, Nothing)
        End Function

        ''' <summary>
        ''' Displays a message box to the user with a blank caption.
        ''' </summary>
        ''' <param name="owner">The form which is wanting to display this message box.</param>
        ''' <param name="message">The message to display in the client area of the message box.</param>
        ''' <param name="buttons">A dictionary containing button labels (and their respective DialogResult values).</param>
        ''' <returns>The DialogResult value of the button which caused the message box to dismiss.</returns>
        Public Shared Function Show(ByVal owner As Form, ByVal message As String, ByVal buttons As Dictionary(Of String, DialogResult)) As DialogResult
            Return Show(owner, message, Nothing, buttons)
        End Function

        ''' <summary>
        ''' Displays a message box to the user.
        ''' </summary>
        ''' <param name="owner">The form which is wanting to display this message box.</param>
        ''' <param name="message">The message to display in the client area of the message box.</param>
        ''' <param name="caption">The title to display in the caption of the message box.</param>
        ''' <param name="buttons">A dictionary containing button labels (and their respective DialogResult values).</param>
        ''' <returns>The DialogResult value of the button which caused the message box to dismiss.</returns>
        Public Shared Function Show(ByVal owner As Form, ByVal message As String, ByVal caption As String, ByVal buttons As Dictionary(Of String, DialogResult)) As DialogResult
            ' Create a new MessageBoxForm and configure the
            ' caption and message text strings.
            Dim form As New MessageBoxForm()
            form.Owner = owner

            form.Text = caption
            form.lblMessage.Text = message

            ' We only want the ControlBox (i.e. the [OK] button) if
            ' there are no custom buttons for this message box
            form.ControlBox = (buttons Is Nothing OrElse buttons.Count = 0)

            ' Resize the form to make it just large enough to fit
            ' the contents and no larger.
            SizeFormToContent(form, Not form.ControlBox)

            ' Add the desired buttons to the form. They are all
            ' docked left into a container panel, so that we
            ' don't need to bother with manual co-ordinate
            ' calculations within this class.
            If Not form.ControlBox Then
                form.SuspendLayout()

                Dim buttonWidth As Integer = (form.pnlButtons.Width - CInt(Math.Floor(form.MyAutoScaleFactor.Width * 2)) - (buttons.Count - 1) * form.pnlButtons.Left) / buttons.Count
                For Each kp As KeyValuePair(Of String, DialogResult) In buttons
                    ' Add a button to the form
                    Dim button As New Button()
                    button.Text = kp.Key
                    button.DialogResult = kp.Value
                    button.Dock = DockStyle.Left
                    button.Width = buttonWidth
                    form.pnlButtons.Controls.Add(button)
                    form.pnlButtons.Controls.SetChildIndex(button, 0)

                    ' Add a seperator panel to the form to
                    ' seperate this button from the next one
                    Dim seperator As New Panel()
                    seperator.Width = form.pnlButtons.Left
                    seperator.Dock = DockStyle.Left
                    form.pnlButtons.Controls.Add(seperator)
                    form.pnlButtons.Controls.SetChildIndex(seperator, 0)
                Next

                form.ResumeLayout()
            End If

            ' Finally we are ready to
            ' display the dialog box...
            Return form.ShowDialog()
        End Function

        Protected Overrides Sub OnLoad(ByVal e As EventArgs)
            MyBase.OnLoad(e)

            ' When the message box is displayed play
            ' the desired sound effect
            If beepType <> MessageBeepType.None Then
                MessageBeep(beepType)
            End If
        End Sub

        ' Calculate how big we want the message box text label
        ' to be to fit the user's specified text
        Private Shared Function CalcTextSize(ByVal form As MessageBoxForm) As Size
            Dim bounds As New RECT(form.lblMessage.ClientRectangle)
            Dim titleWidth As Integer

            ' Measure how tall/wide the text will be assuming we
            ' can't go wider than the maximum width of the label.
            Using g As Graphics = form.CreateGraphics()
                Dim hdc As IntPtr = IntPtr.Zero
                Dim hFont As IntPtr = IntPtr.Zero
                Try
                    hdc = g.GetHdc()
                    hFont = form.lblMessage.Font.ToHfont()

                    ' Get the Drawtext Win32 API to measure the width and height of the
                    ' text. Using the width specified in bounds.Right - bounds.Left to
                    ' dictate the word wrapping behaviour.
                    Dim oldFont As IntPtr = SelectObject(hdc, hFont)
                    DrawText(hdc, form.lblMessage.Text, form.lblMessage.Text.Length, bounds, DT_LEFT Or DT_TOP Or DT_CALCRECT Or DT_NOPREFIX Or DT_WORDBREAK)
                    SelectObject(hdc, oldFont)

                    ' Calculate how wide the caption will be (roughly)
                    titleWidth = CInt(Math.Floor(g.MeasureString(form.Text, form.Font).Width * 1.5)) + SystemInformation.MenuHeight
                Finally
                    ' Tidy up our native resources
                    If hFont <> IntPtr.Zero Then
                        DeleteObject(hFont)
                    End If

                    If hdc <> IntPtr.Zero Then
                        g.ReleaseHdc(hdc)
                    End If
                End Try
            End Using

            ' Return a rectangle just large enough to hold the
            ' message we want to display. Very narrow dialogs look
            ' ugly and probably don't leave enough space for the
            ' caption, so if the width would be narrower than 
            ' what we roughly calculate is required to make the
            ' caption fully visible, we use this width instead.
            ' This seems to give a behaviour similiar to the
            ' Win32 MessageBox API.
            If (bounds.Right - bounds.Left) < titleWidth Then
                Return New Size(titleWidth, bounds.Bottom - bounds.Top)
            Else
                Return New Size(bounds.Right - bounds.Left + CInt(Math.Floor(form.MyAutoScaleFactor.Width * 4)), bounds.Bottom - bounds.Top)
            End If
        End Function

        Private Shared Sub SizeFormToContent(ByVal form As MessageBoxForm, ByVal hasButtons As Boolean)
            ' Only show the panel of buttons if we
            ' have atleast one button to display...
            form.pnlButtons.Visible = hasButtons

            ' We'll limit the width of our message box form to
            ' 3/4ths of the working area of the screen.
            form.Width = (Screen.PrimaryScreen.WorkingArea.Width \ 4) * 3

            ' Figure out how big we need to make our text label
            ' to display the user's message
            Dim textSize As Size = CalcTextSize(form)

            ' And based upon this, calculate and then eventually
            ' set the form's size so that the label is of the
            ' correct size (due to it's anchoring, as we resize
            ' the form we'll also resize the label automatically).
            Dim width As Integer = 2 * form.lblMessage.Left + textSize.Width

            Dim height As Integer = 2 * form.lblMessage.Top + textSize.Height
            If hasButtons Then
                height += form.lblMessage.Top + form.pnlButtons.Height
            End If

            form.Size = New Size(width, height)
            form.lblMessage.Height = textSize.Height
        End Sub
    End Class
