<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TACCUBE1
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
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem5 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.optEntryType = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraNumericEditor4 = New Infragistics.Win.UltraWinEditors.UltraNumericEditor()
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraNumericEditor2 = New Infragistics.Win.UltraWinEditors.UltraNumericEditor()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraNumericEditor1 = New Infragistics.Win.UltraWinEditors.UltraNumericEditor()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraNumericEditor3 = New Infragistics.Win.UltraWinEditors.UltraNumericEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.btnCalculate = New Infragistics.Win.Misc.UltraButton()
        Me.btnCancel = New Infragistics.Win.Misc.UltraButton()
        Me.btnOk = New Infragistics.Win.Misc.UltraButton()
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optEntryType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.UltraNumericEditor4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraNumericEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraNumericEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraNumericEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.btnCancel)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.btnOk)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.btnCalculate)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraNumericEditor3)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraLabel1)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraNumericEditor4)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraLabel5)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraNumericEditor2)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraLabel4)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraNumericEditor1)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraLabel3)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraGroupBox1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(284, 281)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 281)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(284, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 281)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(284, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 281)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(284, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'optEntryType
        '
        Me.optEntryType.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.optEntryType.CheckedIndex = 1
        ValueListItem2.DataValue = "E"
        ValueListItem2.DisplayText = "English"
        ValueListItem5.DataValue = "M"
        ValueListItem5.DisplayText = "Metric"
        Me.optEntryType.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem2, ValueListItem5})
        Me.optEntryType.Location = New System.Drawing.Point(16, 23)
        Me.optEntryType.Name = "optEntryType"
        Me.optEntryType.Size = New System.Drawing.Size(155, 19)
        Me.optEntryType.TabIndex = 1
        Me.optEntryType.Text = "Metric"
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.UltraGroupBox1.Controls.Add(Me.optEntryType)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(233, 61)
        Me.UltraGroupBox1.TabIndex = 0
        Me.UltraGroupBox1.Text = "Entry Type"
        '
        'UltraNumericEditor4
        '
        Me.Absx1.SetABSColumnName(Me.UltraNumericEditor4, "PKG_H")
        Me.UltraNumericEditor4.Location = New System.Drawing.Point(88, 148)
        Me.UltraNumericEditor4.MaxValue = 100
        Me.UltraNumericEditor4.MinValue = 0
        Me.UltraNumericEditor4.Name = "UltraNumericEditor4"
        Me.UltraNumericEditor4.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.UltraNumericEditor4.Size = New System.Drawing.Size(71, 25)
        Me.UltraNumericEditor4.TabIndex = 4
        '
        'UltraLabel5
        '
        Me.UltraLabel5.AutoSize = True
        Me.UltraLabel5.Location = New System.Drawing.Point(17, 155)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(49, 18)
        Me.UltraLabel5.TabIndex = 190
        Me.UltraLabel5.Text = "Height"
        '
        'UltraNumericEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraNumericEditor2, "PKG_W")
        Me.UltraNumericEditor2.Location = New System.Drawing.Point(88, 117)
        Me.UltraNumericEditor2.MaxValue = 100
        Me.UltraNumericEditor2.MinValue = 0
        Me.UltraNumericEditor2.Name = "UltraNumericEditor2"
        Me.UltraNumericEditor2.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.UltraNumericEditor2.Size = New System.Drawing.Size(71, 25)
        Me.UltraNumericEditor2.TabIndex = 3
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(17, 124)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(44, 18)
        Me.UltraLabel4.TabIndex = 189
        Me.UltraLabel4.Text = "Width"
        '
        'UltraNumericEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraNumericEditor1, "PKG_L")
        Me.UltraNumericEditor1.Location = New System.Drawing.Point(88, 86)
        Me.UltraNumericEditor1.MaxValue = 100
        Me.UltraNumericEditor1.MinValue = 0
        Me.UltraNumericEditor1.Name = "UltraNumericEditor1"
        Me.UltraNumericEditor1.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.UltraNumericEditor1.Size = New System.Drawing.Size(71, 25)
        Me.UltraNumericEditor1.TabIndex = 2
        '
        'UltraLabel3
        '
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(17, 93)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(51, 18)
        Me.UltraLabel3.TabIndex = 188
        Me.UltraLabel3.Text = "Length"
        '
        'UltraNumericEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraNumericEditor3, "PKG_C")
        Appearance1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Appearance1.BackColorDisabled = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.UltraNumericEditor3.Appearance = Appearance1
        Me.UltraNumericEditor3.Enabled = False
        Me.UltraNumericEditor3.FormatString = "0.000000"
        Me.UltraNumericEditor3.Location = New System.Drawing.Point(88, 196)
        Me.UltraNumericEditor3.MaxValue = 10000.0R
        Me.UltraNumericEditor3.MinValue = 0
        Me.UltraNumericEditor3.Name = "UltraNumericEditor3"
        Me.UltraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.[Double]
        Me.UltraNumericEditor3.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.UltraNumericEditor3.Size = New System.Drawing.Size(71, 25)
        Me.UltraNumericEditor3.TabIndex = 5
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(17, 203)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(39, 18)
        Me.UltraLabel1.TabIndex = 192
        Me.UltraLabel1.Text = "Cube"
        '
        'btnCalculate
        '
        Me.btnCalculate.Location = New System.Drawing.Point(169, 196)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(112, 25)
        Me.btnCalculate.TabIndex = 6
        Me.btnCalculate.Text = "Calculate"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(204, 242)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(77, 32)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(129, 242)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(77, 32)
        Me.btnOk.TabIndex = 7
        Me.btnOk.Text = "Ok"
        '
        'TACCUBE1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 281)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "TACCUBE1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Calculate Cubic Feet"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        Me.ASFBASE2_Fill_Panel.PerformLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optEntryType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        CType(Me.UltraNumericEditor4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraNumericEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraNumericEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraNumericEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents optEntryType As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents UltraNumericEditor3 As Infragistics.Win.UltraWinEditors.UltraNumericEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraNumericEditor4 As Infragistics.Win.UltraWinEditors.UltraNumericEditor
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraNumericEditor2 As Infragistics.Win.UltraWinEditors.UltraNumericEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraNumericEditor1 As Infragistics.Win.UltraWinEditors.UltraNumericEditor
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents btnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents btnOk As Infragistics.Win.Misc.UltraButton
    Friend WithEvents btnCalculate As Infragistics.Win.Misc.UltraButton
End Class
