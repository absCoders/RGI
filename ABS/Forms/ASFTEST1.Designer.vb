<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFTEST1
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
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraButton1
        '
        Me.UltraButton1.Location = New System.Drawing.Point(9, 3)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(98, 42)
        Me.UltraButton1.TabIndex = 0
        Me.UltraButton1.Text = "UltraButton1"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraButton2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraButton1)
        Me.SplitContainer1.Size = New System.Drawing.Size(292, 266)
        Me.SplitContainer1.SplitterDistance = 47
        Me.SplitContainer1.TabIndex = 1
        '
        'UltraButton2
        '
        Me.UltraButton2.Location = New System.Drawing.Point(191, 2)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(98, 42)
        Me.UltraButton2.TabIndex = 1
        Me.UltraButton2.Text = "UltraButton2"
        '
        'ASFTEST1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "ASFTEST1"
        Me.Text = "ASFTEST1"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
End Class
