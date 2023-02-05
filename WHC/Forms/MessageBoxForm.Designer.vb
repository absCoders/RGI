Partial Class MessageBoxForm
    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.IContainer = Nothing
    Private mainMenu1 As System.Windows.Forms.MainMenu

    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "Windows Form Designer generated code"

    ''' <summary>
    ''' Required method for Designer support - do not modify
    ''' the contents of this method with the code editor.
    ''' </summary>
    Private Sub InitializeComponent()
        Me.mainMenu1 = New System.Windows.Forms.MainMenu()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.pnlButtons = New System.Windows.Forms.Panel()
        Me.SuspendLayout()
        ' 
        ' lblMessage
        ' 
        Me.lblMessage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblMessage.Location = New System.Drawing.Point(9, 9)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(140, 62)
        Me.lblMessage.Text = "This is my example message!"
        ' 
        ' pnlButtons
        ' 
        Me.pnlButtons.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlButtons.Location = New System.Drawing.Point(9, 80)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Size = New System.Drawing.Size(140, 26)
        ' 
        ' MessageBoxForm
        ' 
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.CenterFormOnScreen = True
        Me.ClientSize = New System.Drawing.Size(158, 115)
        Me.Controls.Add(Me.pnlButtons)
        Me.Controls.Add(Me.lblMessage)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Location = New System.Drawing.Point(50, 50)
        Me.Menu = Me.mainMenu1
        Me.MinimizeBox = False
        Me.Name = "MessageBoxForm"
        Me.Text = "MessageBoxForm"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private lblMessage As System.Windows.Forms.Label
    Private pnlButtons As System.Windows.Forms.Panel
End Class