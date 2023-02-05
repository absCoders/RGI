<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class LoginForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtUserID = New System.Windows.Forms.TextBox
        Me.lblText = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.lblPassword = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(82, 6)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(132, 23)
        Me.txtUserID.TabIndex = 0
        '
        'lblText
        '
        Me.lblText.Location = New System.Drawing.Point(3, 9)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(73, 20)
        Me.lblText.Text = "User ID:"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(96, 75)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(50, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        '
        'lblPassword
        '
        Me.lblPassword.Location = New System.Drawing.Point(3, 41)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(73, 20)
        Me.lblPassword.Text = "Password:"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(164, 75)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(50, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(82, 35)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(132, 23)
        Me.txtPassword.TabIndex = 1
        '
        'LoginForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(238, 117)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lblText)
        Me.Controls.Add(Me.txtUserID)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(50, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LoginForm"
        Me.Text = "Input"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents lblText As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
End Class
