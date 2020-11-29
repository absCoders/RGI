<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TAFIMGV1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE2
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.panIMAGE = New System.Windows.Forms.Panel()
        Me.imgSTYLE = New System.Windows.Forms.PictureBox()
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton()
        Me.rdoRezH = New System.Windows.Forms.RadioButton()
        Me.rdoRezL = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboICTIMAGT = New System.Windows.Forms.ComboBox()
        Me.cmdDone = New Infragistics.Win.Misc.UltraButton()
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.panIMAGE.SuspendLayout()
        CType(Me.imgSTYLE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(553, 516)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 516)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(553, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 516)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(553, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 516)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(553, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.panIMAGE)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.rdoRezH)
        Me.SplitContainer1.Panel2.Controls.Add(Me.rdoRezL)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboICTIMAGT)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdDone)
        Me.SplitContainer1.Size = New System.Drawing.Size(553, 516)
        Me.SplitContainer1.SplitterDistance = 437
        Me.SplitContainer1.TabIndex = 2
        '
        'panIMAGE
        '
        Me.panIMAGE.AutoScroll = True
        Me.panIMAGE.Controls.Add(Me.imgSTYLE)
        Me.panIMAGE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.panIMAGE.Location = New System.Drawing.Point(0, 0)
        Me.panIMAGE.Name = "panIMAGE"
        Me.panIMAGE.Size = New System.Drawing.Size(553, 437)
        Me.panIMAGE.TabIndex = 0
        '
        'imgSTYLE
        '
        Me.imgSTYLE.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.imgSTYLE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.imgSTYLE.Location = New System.Drawing.Point(0, 0)
        Me.imgSTYLE.Name = "imgSTYLE"
        Me.imgSTYLE.Size = New System.Drawing.Size(553, 437)
        Me.imgSTYLE.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.imgSTYLE.TabIndex = 1
        Me.imgSTYLE.TabStop = False
        '
        'cmdSave
        '
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdSave.Location = New System.Drawing.Point(68, 38)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(50, 25)
        Me.cmdSave.TabIndex = 16
        Me.cmdSave.Text = "Save"
        '
        'rdoRezH
        '
        Me.rdoRezH.AutoSize = True
        Me.rdoRezH.Checked = True
        Me.rdoRezH.Location = New System.Drawing.Point(418, 43)
        Me.rdoRezH.Name = "rdoRezH"
        Me.rdoRezH.Size = New System.Drawing.Size(82, 20)
        Me.rdoRezH.TabIndex = 13
        Me.rdoRezH.TabStop = True
        Me.rdoRezH.Text = "High Rez"
        Me.rdoRezH.UseVisualStyleBackColor = True
        '
        'rdoRezL
        '
        Me.rdoRezL.AutoSize = True
        Me.rdoRezL.Location = New System.Drawing.Point(418, 19)
        Me.rdoRezL.Name = "rdoRezL"
        Me.rdoRezL.Size = New System.Drawing.Size(80, 20)
        Me.rdoRezL.TabIndex = 12
        Me.rdoRezL.Text = "Low Rez"
        Me.rdoRezL.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(135, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 16)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Image Type"
        '
        'cboICTIMAGT
        '
        Me.cboICTIMAGT.FormattingEnabled = True
        Me.cboICTIMAGT.Location = New System.Drawing.Point(138, 38)
        Me.cboICTIMAGT.Name = "cboICTIMAGT"
        Me.cboICTIMAGT.Size = New System.Drawing.Size(260, 24)
        Me.cboICTIMAGT.TabIndex = 10
        '
        'cmdDone
        '
        Me.cmdDone.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdDone.Location = New System.Drawing.Point(12, 38)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(50, 25)
        Me.cmdDone.TabIndex = 2
        Me.cmdDone.Text = "Done"
        '
        'TAFIMGV1
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(553, 516)
        Me.ControlBox = False
        Me.Name = "TAFIMGV1"
        Me.Text = "Image Viewer"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.panIMAGE.ResumeLayout(False)
        CType(Me.imgSTYLE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cmdDone As Infragistics.Win.Misc.UltraButton
    Friend WithEvents rdoRezH As RadioButton
    Friend WithEvents rdoRezL As RadioButton
    Friend WithEvents Label1 As Label
    Friend WithEvents cboICTIMAGT As ComboBox
    Friend WithEvents panIMAGE As Panel
    Friend WithEvents imgSTYLE As PictureBox
    Friend WithEvents cmdSave As Misc.UltraButton
End Class
