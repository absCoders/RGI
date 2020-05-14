<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SOFORDRQ
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE2
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim ValueListItem3 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem6 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem24 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem7 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem9 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem23 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem10 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem14 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem22 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem17 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem18 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem19 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem20 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem21 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.pnlMultiPrice = New System.Windows.Forms.Panel()
        Me.grpPVC = New System.Windows.Forms.GroupBox()
        Me.panPVCALWAYS = New System.Windows.Forms.Panel()
        Me.optPRICE_TIER_PVC = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.UltraLabel61 = New Infragistics.Win.Misc.UltraLabel()
        Me.grpNONPVC = New System.Windows.Forms.GroupBox()
        Me.panEXTRA = New System.Windows.Forms.Panel()
        Me.lblCUST_DISC_PCT_EXTRA = New Infragistics.Win.Misc.UltraLabel()
        Me.optDISC_PCT_EXTRA = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.numDISC_PCT = New Infragistics.Win.UltraWinEditors.UltraNumericEditor()
        Me.panNONALWAYS = New System.Windows.Forms.Panel()
        Me.UltraLabel59 = New Infragistics.Win.Misc.UltraLabel()
        Me.optPRICE_TIER = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.grpFEPricing = New System.Windows.Forms.GroupBox()
        Me.panFEFD = New System.Windows.Forms.Panel()
        Me.optFEFD = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.panFEExtra = New System.Windows.Forms.Panel()
        Me.numFEFDFACTOR = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optQuoteType = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.cmdFinished = New Infragistics.Win.Misc.UltraButton()
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
        Me.pnlMultiPrice.SuspendLayout()
        Me.grpPVC.SuspendLayout()
        Me.panPVCALWAYS.SuspendLayout()
        CType(Me.optPRICE_TIER_PVC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpNONPVC.SuspendLayout()
        Me.panEXTRA.SuspendLayout()
        CType(Me.optDISC_PCT_EXTRA, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numDISC_PCT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panNONALWAYS.SuspendLayout()
        CType(Me.optPRICE_TIER, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFEPricing.SuspendLayout()
        Me.panFEFD.SuspendLayout()
        CType(Me.optFEFD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panFEExtra.SuspendLayout()
        CType(Me.numFEFDFACTOR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.optQuoteType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Margin = New System.Windows.Forms.Padding(6)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(657, 225)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 225)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(657, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 225)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(657, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 225)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(657, 0)
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.pnlMultiPrice)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(657, 225)
        Me.SplitContainer1.SplitterDistance = 185
        Me.SplitContainer1.TabIndex = 2
        '
        'pnlMultiPrice
        '
        Me.pnlMultiPrice.Controls.Add(Me.grpPVC)
        Me.pnlMultiPrice.Controls.Add(Me.grpNONPVC)
        Me.pnlMultiPrice.Controls.Add(Me.grpFEPricing)
        Me.pnlMultiPrice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMultiPrice.Location = New System.Drawing.Point(0, 0)
        Me.pnlMultiPrice.Name = "pnlMultiPrice"
        Me.pnlMultiPrice.Size = New System.Drawing.Size(657, 185)
        Me.pnlMultiPrice.TabIndex = 0
        Me.pnlMultiPrice.Visible = False
        '
        'grpPVC
        '
        Me.grpPVC.Controls.Add(Me.panPVCALWAYS)
        Me.grpPVC.Location = New System.Drawing.Point(231, 5)
        Me.grpPVC.Name = "grpPVC"
        Me.grpPVC.Size = New System.Drawing.Size(163, 171)
        Me.grpPVC.TabIndex = 5
        Me.grpPVC.TabStop = False
        Me.grpPVC.Text = "Customer PVC Pricing"
        '
        'panPVCALWAYS
        '
        Me.panPVCALWAYS.Controls.Add(Me.optPRICE_TIER_PVC)
        Me.panPVCALWAYS.Controls.Add(Me.UltraLabel61)
        Me.panPVCALWAYS.Location = New System.Drawing.Point(6, 22)
        Me.panPVCALWAYS.Name = "panPVCALWAYS"
        Me.panPVCALWAYS.Size = New System.Drawing.Size(93, 103)
        Me.panPVCALWAYS.TabIndex = 218
        '
        'optPRICE_TIER_PVC
        '
        Me.Absx1.SetABSBindToTable(Me.optPRICE_TIER_PVC, False)
        Me.optPRICE_TIER_PVC.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem3.DataValue = "5C"
        ValueListItem3.DisplayText = "5-9 Case"
        ValueListItem6.DataValue = "FC"
        ValueListItem6.DisplayText = "Full Case"
        ValueListItem24.DataValue = "PC"
        ValueListItem24.DisplayText = "Order Qty"
        Me.optPRICE_TIER_PVC.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem3, ValueListItem6, ValueListItem24})
        Me.optPRICE_TIER_PVC.Location = New System.Drawing.Point(3, 24)
        Me.optPRICE_TIER_PVC.Name = "optPRICE_TIER_PVC"
        Me.optPRICE_TIER_PVC.Size = New System.Drawing.Size(87, 71)
        Me.optPRICE_TIER_PVC.TabIndex = 207
        '
        'UltraLabel61
        '
        Me.UltraLabel61.AutoSize = True
        Me.UltraLabel61.Location = New System.Drawing.Point(3, 3)
        Me.UltraLabel61.Name = "UltraLabel61"
        Me.UltraLabel61.Size = New System.Drawing.Size(81, 18)
        Me.UltraLabel61.TabIndex = 206
        Me.UltraLabel61.Text = "Always Use"
        '
        'grpNONPVC
        '
        Me.grpNONPVC.Controls.Add(Me.panEXTRA)
        Me.grpNONPVC.Controls.Add(Me.panNONALWAYS)
        Me.grpNONPVC.Location = New System.Drawing.Point(6, 4)
        Me.grpNONPVC.Name = "grpNONPVC"
        Me.grpNONPVC.Size = New System.Drawing.Size(219, 172)
        Me.grpNONPVC.TabIndex = 4
        Me.grpNONPVC.TabStop = False
        Me.grpNONPVC.Text = "Customer Non-PVC Pricing"
        '
        'panEXTRA
        '
        Me.panEXTRA.Controls.Add(Me.lblCUST_DISC_PCT_EXTRA)
        Me.panEXTRA.Controls.Add(Me.optDISC_PCT_EXTRA)
        Me.panEXTRA.Controls.Add(Me.numDISC_PCT)
        Me.panEXTRA.Location = New System.Drawing.Point(110, 23)
        Me.panEXTRA.Name = "panEXTRA"
        Me.panEXTRA.Size = New System.Drawing.Size(101, 115)
        Me.panEXTRA.TabIndex = 220
        '
        'lblCUST_DISC_PCT_EXTRA
        '
        Me.lblCUST_DISC_PCT_EXTRA.AutoSize = True
        Me.lblCUST_DISC_PCT_EXTRA.Location = New System.Drawing.Point(3, 3)
        Me.lblCUST_DISC_PCT_EXTRA.Name = "lblCUST_DISC_PCT_EXTRA"
        Me.lblCUST_DISC_PCT_EXTRA.Size = New System.Drawing.Size(87, 18)
        Me.lblCUST_DISC_PCT_EXTRA.TabIndex = 221
        Me.lblCUST_DISC_PCT_EXTRA.Text = "Extra Disc%"
        '
        'optDISC_PCT_EXTRA
        '
        Me.Absx1.SetABSBindToTable(Me.optDISC_PCT_EXTRA, False)
        Me.optDISC_PCT_EXTRA.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem7.DataValue = "2"
        ValueListItem7.DisplayText = "10%"
        ValueListItem9.DataValue = "1"
        ValueListItem9.DisplayText = "5%"
        ValueListItem23.DataValue = "0"
        ValueListItem23.DisplayText = "N/A"
        Me.optDISC_PCT_EXTRA.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem7, ValueListItem9, ValueListItem23})
        Me.optDISC_PCT_EXTRA.Location = New System.Drawing.Point(5, 19)
        Me.optDISC_PCT_EXTRA.Name = "optDISC_PCT_EXTRA"
        Me.optDISC_PCT_EXTRA.Size = New System.Drawing.Size(60, 63)
        Me.optDISC_PCT_EXTRA.TabIndex = 220
        '
        'numDISC_PCT
        '
        Me.Absx1.SetABSBindToTable(Me.numDISC_PCT, False)
        Me.numDISC_PCT.AlwaysInEditMode = True
        Me.numDISC_PCT.Location = New System.Drawing.Point(5, 21)
        Me.numDISC_PCT.Name = "numDISC_PCT"
        Me.numDISC_PCT.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.numDISC_PCT.Size = New System.Drawing.Size(85, 25)
        Me.numDISC_PCT.TabIndex = 219
        '
        'panNONALWAYS
        '
        Me.panNONALWAYS.Controls.Add(Me.UltraLabel59)
        Me.panNONALWAYS.Controls.Add(Me.optPRICE_TIER)
        Me.panNONALWAYS.Location = New System.Drawing.Point(6, 23)
        Me.panNONALWAYS.Name = "panNONALWAYS"
        Me.panNONALWAYS.Size = New System.Drawing.Size(101, 115)
        Me.panNONALWAYS.TabIndex = 212
        '
        'UltraLabel59
        '
        Me.UltraLabel59.AutoSize = True
        Me.UltraLabel59.Location = New System.Drawing.Point(3, 3)
        Me.UltraLabel59.Name = "UltraLabel59"
        Me.UltraLabel59.Size = New System.Drawing.Size(81, 18)
        Me.UltraLabel59.TabIndex = 205
        Me.UltraLabel59.Text = "Always Use"
        '
        'optPRICE_TIER
        '
        Me.Absx1.SetABSBindToTable(Me.optPRICE_TIER, False)
        Me.optPRICE_TIER.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem10.DataValue = "FC"
        ValueListItem10.DisplayText = "Full Case"
        ValueListItem14.DataValue = "HC"
        ValueListItem14.DisplayText = "Half Case"
        ValueListItem22.DataValue = "PC"
        ValueListItem22.DisplayText = "Order Qty"
        ValueListItem17.DataValue = "SP"
        ValueListItem17.DisplayText = "Disc%"
        Me.optPRICE_TIER.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem10, ValueListItem14, ValueListItem22, ValueListItem17})
        Me.optPRICE_TIER.Location = New System.Drawing.Point(3, 21)
        Me.optPRICE_TIER.Name = "optPRICE_TIER"
        Me.optPRICE_TIER.Size = New System.Drawing.Size(95, 74)
        Me.optPRICE_TIER.TabIndex = 204
        '
        'grpFEPricing
        '
        Me.grpFEPricing.Controls.Add(Me.panFEFD)
        Me.grpFEPricing.Controls.Add(Me.panFEExtra)
        Me.grpFEPricing.Location = New System.Drawing.Point(424, 5)
        Me.grpFEPricing.Name = "grpFEPricing"
        Me.grpFEPricing.Size = New System.Drawing.Size(158, 177)
        Me.grpFEPricing.TabIndex = 3
        Me.grpFEPricing.TabStop = False
        Me.grpFEPricing.Text = "Customer FE Pricing"
        '
        'panFEFD
        '
        Me.panFEFD.Controls.Add(Me.optFEFD)
        Me.panFEFD.Location = New System.Drawing.Point(6, 21)
        Me.panFEFD.Name = "panFEFD"
        Me.panFEFD.Size = New System.Drawing.Size(101, 88)
        Me.panFEFD.TabIndex = 222
        '
        'optFEFD
        '
        Me.Absx1.SetABSBindToTable(Me.optFEFD, False)
        Me.optFEFD.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem18.CheckState = System.Windows.Forms.CheckState.Checked
        ValueListItem18.DataValue = "FE"
        ValueListItem18.DisplayText = "FE"
        ValueListItem19.DataValue = "FD"
        ValueListItem19.DisplayText = "FD"
        ValueListItem20.DataValue = "FEM"
        ValueListItem20.DisplayText = "FE Mix"
        ValueListItem21.DataValue = "FDM"
        ValueListItem21.DisplayText = "FD Mix"
        Me.optFEFD.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem18, ValueListItem19, ValueListItem20, ValueListItem21})
        Me.optFEFD.Location = New System.Drawing.Point(3, 6)
        Me.optFEFD.Name = "optFEFD"
        Me.optFEFD.Size = New System.Drawing.Size(87, 74)
        Me.optFEFD.TabIndex = 204
        '
        'panFEExtra
        '
        Me.panFEExtra.Controls.Add(Me.numFEFDFACTOR)
        Me.panFEExtra.Controls.Add(Me.Label2)
        Me.panFEExtra.Location = New System.Drawing.Point(6, 113)
        Me.panFEExtra.Name = "panFEExtra"
        Me.panFEExtra.Size = New System.Drawing.Size(101, 57)
        Me.panFEExtra.TabIndex = 221
        '
        'numFEFDFACTOR
        '
        Me.numFEFDFACTOR.Location = New System.Drawing.Point(5, 26)
        Me.numFEFDFACTOR.Maximum = New Decimal(New Integer() {0, 0, 0, 0})
        Me.numFEFDFACTOR.Minimum = New Decimal(New Integer() {14, 0, 0, -2147483648})
        Me.numFEFDFACTOR.Name = "numFEFDFACTOR"
        Me.numFEFDFACTOR.Size = New System.Drawing.Size(86, 23)
        Me.numFEFDFACTOR.TabIndex = 4
        Me.numFEFDFACTOR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "FEFD Factor"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optQuoteType)
        Me.GroupBox1.Controls.Add(Me.cmdFinished)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(657, 36)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'optQuoteType
        '
        Me.optQuoteType.CheckedIndex = 0
        ValueListItem1.CheckState = System.Windows.Forms.CheckState.Checked
        ValueListItem1.DataValue = "S"
        ValueListItem1.DisplayText = "Standard Quote"
        ValueListItem2.DataValue = "M"
        ValueListItem2.DisplayText = "Multi-Priced Quote"
        Me.optQuoteType.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem1, ValueListItem2})
        Me.optQuoteType.Location = New System.Drawing.Point(92, 6)
        Me.optQuoteType.Name = "optQuoteType"
        Me.optQuoteType.Size = New System.Drawing.Size(281, 20)
        Me.optQuoteType.TabIndex = 201
        Me.optQuoteType.Tag = ""
        Me.optQuoteType.Text = "Standard Quote"
        '
        'cmdFinished
        '
        Me.cmdFinished.Location = New System.Drawing.Point(6, 1)
        Me.cmdFinished.Name = "cmdFinished"
        Me.cmdFinished.Size = New System.Drawing.Size(70, 33)
        Me.cmdFinished.TabIndex = 8
        Me.cmdFinished.Text = "Finished"
        '
        'SOFORDRQ
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(657, 225)
        Me.ControlBox = False
        Me.Margin = New System.Windows.Forms.Padding(6, 8, 6, 8)
        Me.Name = "SOFORDRQ"
        Me.Text = "Quote Options"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.pnlMultiPrice.ResumeLayout(False)
        Me.grpPVC.ResumeLayout(False)
        Me.panPVCALWAYS.ResumeLayout(False)
        Me.panPVCALWAYS.PerformLayout()
        CType(Me.optPRICE_TIER_PVC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpNONPVC.ResumeLayout(False)
        Me.panEXTRA.ResumeLayout(False)
        Me.panEXTRA.PerformLayout()
        CType(Me.optDISC_PCT_EXTRA, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numDISC_PCT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panNONALWAYS.ResumeLayout(False)
        Me.panNONALWAYS.PerformLayout()
        CType(Me.optPRICE_TIER, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFEPricing.ResumeLayout(False)
        Me.panFEFD.ResumeLayout(False)
        CType(Me.optFEFD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panFEExtra.ResumeLayout(False)
        Me.panFEExtra.PerformLayout()
        CType(Me.numFEFDFACTOR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.optQuoteType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdFinished As Infragistics.Win.Misc.UltraButton
    Friend WithEvents pnlMultiPrice As Panel
    Friend WithEvents optQuoteType As UltraWinEditors.UltraOptionSet
    Friend WithEvents grpPVC As GroupBox
    Friend WithEvents panPVCALWAYS As Panel
    Friend WithEvents optPRICE_TIER_PVC As UltraWinEditors.UltraOptionSet
    Friend WithEvents UltraLabel61 As Misc.UltraLabel
    Friend WithEvents grpNONPVC As GroupBox
    Friend WithEvents panEXTRA As Panel
    Friend WithEvents lblCUST_DISC_PCT_EXTRA As Misc.UltraLabel
    Friend WithEvents optDISC_PCT_EXTRA As UltraWinEditors.UltraOptionSet
    Friend WithEvents numDISC_PCT As UltraWinEditors.UltraNumericEditor
    Friend WithEvents panNONALWAYS As Panel
    Friend WithEvents UltraLabel59 As Misc.UltraLabel
    Friend WithEvents optPRICE_TIER As UltraWinEditors.UltraOptionSet
    Friend WithEvents grpFEPricing As GroupBox
    Friend WithEvents panFEFD As Panel
    Friend WithEvents optFEFD As UltraWinEditors.UltraOptionSet
    Friend WithEvents panFEExtra As Panel
    Friend WithEvents numFEFDFACTOR As NumericUpDown
    Friend WithEvents Label2 As Label
End Class
