<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TAFZPLT1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.picPNG = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.picPNG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'picPNG
        '
        Me.picPNG.Location = New System.Drawing.Point(0, 0)
        Me.picPNG.Name = "picPNG"
        Me.picPNG.Size = New System.Drawing.Size(259, 225)
        Me.picPNG.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picPNG.TabIndex = 0
        Me.picPNG.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.Controls.Add(Me.picPNG)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(606, 424)
        Me.Panel1.TabIndex = 1
        '
        'TAFZPLT1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(606, 424)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "TAFZPLT1"
        Me.Text = "TAFZPLT1"
        CType(Me.picPNG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents picPNG As PictureBox
    Friend WithEvents Panel1 As Panel
End Class
