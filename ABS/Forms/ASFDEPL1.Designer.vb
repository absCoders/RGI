<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ASFDEPL1
    Inherits System.Windows.Forms.Form

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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SELECTED")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DLL_NAME")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DLL_DESC")
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CLIENT")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IPADDRESS_PROD")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IPADDRESS_TEST")
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Me.grdDLLS = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cmdDeploy = New Infragistics.Win.Misc.UltraButton()
        Me.cmbClient = New Infragistics.Win.UltraWinGrid.UltraCombo()
        Me.lblUSER_ID = New Infragistics.Win.Misc.UltraLabel()
        Me.optRegion = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.btnSelect = New Infragistics.Win.Misc.UltraButton()
        Me.btnDeSelect = New Infragistics.Win.Misc.UltraButton()
        CType(Me.grdDLLS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.cmbClient, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optRegion, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdDLLS
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdDLLS.DisplayLayout.Appearance = Appearance1
        Me.grdDLLS.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ExtendLastColumn
        UltraGridColumn2.Header.Caption = "Sel"
        UltraGridColumn2.Header.VisiblePosition = 0
        UltraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn2.Width = 52
        UltraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn3.Header.Caption = "Assembly"
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Width = 140
        UltraGridColumn4.Header.Caption = "Description"
        UltraGridColumn4.Header.VisiblePosition = 2
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn2, UltraGridColumn3, UltraGridColumn4})
        Me.grdDLLS.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdDLLS.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdDLLS.DisplayLayout.CaptionAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdDLLS.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdDLLS.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdDLLS.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdDLLS.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdDLLS.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdDLLS.DisplayLayout.MaxColScrollRegions = 1
        Me.grdDLLS.DisplayLayout.MaxRowScrollRegions = 1
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdDLLS.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdDLLS.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdDLLS.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdDLLS.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdDLLS.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.grdDLLS.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdDLLS.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdDLLS.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdDLLS.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdDLLS.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdDLLS.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdDLLS.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdDLLS.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdDLLS.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdDLLS.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.grdDLLS.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[True]
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdDLLS.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdDLLS.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdDLLS.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdDLLS.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdDLLS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdDLLS.Location = New System.Drawing.Point(0, 0)
        Me.grdDLLS.Margin = New System.Windows.Forms.Padding(4)
        Me.grdDLLS.Name = "grdDLLS"
        Me.grdDLLS.Size = New System.Drawing.Size(910, 462)
        Me.grdDLLS.TabIndex = 1
        Me.grdDLLS.Text = "DLLs to Deploy from Development to Live Server"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdDLLS)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDeSelect)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSelect)
        Me.SplitContainer1.Panel2.Controls.Add(Me.optRegion)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdDeploy)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmbClient)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblUSER_ID)
        Me.SplitContainer1.Size = New System.Drawing.Size(910, 532)
        Me.SplitContainer1.SplitterDistance = 462
        Me.SplitContainer1.TabIndex = 2
        '
        'cmdDeploy
        '
        Me.cmdDeploy.Location = New System.Drawing.Point(338, 30)
        Me.cmdDeploy.Name = "cmdDeploy"
        Me.cmdDeploy.Size = New System.Drawing.Size(139, 26)
        Me.cmdDeploy.TabIndex = 2
        Me.cmdDeploy.TabStop = False
        Me.cmdDeploy.Text = "Deploy"
        '
        'cmbClient
        '
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Appearance13.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cmbClient.DisplayLayout.Appearance = Appearance13
        UltraGridColumn1.Header.Caption = "Client"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn5.Header.Caption = "Production"
        UltraGridColumn5.Header.VisiblePosition = 1
        UltraGridColumn6.Header.Caption = "Test"
        UltraGridColumn6.Header.VisiblePosition = 2
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn5, UltraGridColumn6})
        Me.cmbClient.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.cmbClient.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cmbClient.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance14.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance14.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance14.BorderColor = System.Drawing.SystemColors.Window
        Me.cmbClient.DisplayLayout.GroupByBox.Appearance = Appearance14
        Appearance15.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cmbClient.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance15
        Me.cmbClient.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance16.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance16.BackColor2 = System.Drawing.SystemColors.Control
        Appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance16.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cmbClient.DisplayLayout.GroupByBox.PromptAppearance = Appearance16
        Me.cmbClient.DisplayLayout.MaxColScrollRegions = 1
        Me.cmbClient.DisplayLayout.MaxRowScrollRegions = 1
        Appearance17.BackColor = System.Drawing.SystemColors.Window
        Appearance17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmbClient.DisplayLayout.Override.ActiveCellAppearance = Appearance17
        Appearance18.BackColor = System.Drawing.SystemColors.Highlight
        Appearance18.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cmbClient.DisplayLayout.Override.ActiveRowAppearance = Appearance18
        Me.cmbClient.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cmbClient.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance19.BackColor = System.Drawing.SystemColors.Window
        Me.cmbClient.DisplayLayout.Override.CardAreaAppearance = Appearance19
        Appearance20.BorderColor = System.Drawing.Color.Silver
        Appearance20.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cmbClient.DisplayLayout.Override.CellAppearance = Appearance20
        Me.cmbClient.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cmbClient.DisplayLayout.Override.CellPadding = 0
        Appearance21.BackColor = System.Drawing.SystemColors.Control
        Appearance21.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance21.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance21.BorderColor = System.Drawing.SystemColors.Window
        Me.cmbClient.DisplayLayout.Override.GroupByRowAppearance = Appearance21
        Appearance22.TextHAlignAsString = "Left"
        Me.cmbClient.DisplayLayout.Override.HeaderAppearance = Appearance22
        Me.cmbClient.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cmbClient.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance23.BackColor = System.Drawing.SystemColors.Window
        Appearance23.BorderColor = System.Drawing.Color.Silver
        Me.cmbClient.DisplayLayout.Override.RowAppearance = Appearance23
        Me.cmbClient.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance24.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cmbClient.DisplayLayout.Override.TemplateAddRowAppearance = Appearance24
        Me.cmbClient.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cmbClient.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cmbClient.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cmbClient.Location = New System.Drawing.Point(13, 28)
        Me.cmbClient.Name = "cmbClient"
        Me.cmbClient.Size = New System.Drawing.Size(137, 26)
        Me.cmbClient.TabIndex = 0
        '
        'lblUSER_ID
        '
        Me.lblUSER_ID.AutoSize = True
        Me.lblUSER_ID.Location = New System.Drawing.Point(13, 6)
        Me.lblUSER_ID.Margin = New System.Windows.Forms.Padding(4)
        Me.lblUSER_ID.Name = "lblUSER_ID"
        Me.lblUSER_ID.Size = New System.Drawing.Size(43, 18)
        Me.lblUSER_ID.TabIndex = 14
        Me.lblUSER_ID.Text = "Client"
        '
        'optRegion
        '
        Me.optRegion.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.optRegion.CheckedIndex = 0
        ValueListItem1.CheckState = System.Windows.Forms.CheckState.Checked
        ValueListItem1.DataValue = "P"
        ValueListItem1.DisplayText = "Production"
        ValueListItem2.DataValue = "T"
        ValueListItem2.DisplayText = "Test"
        Me.optRegion.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem1, ValueListItem2})
        Me.optRegion.Location = New System.Drawing.Point(156, 35)
        Me.optRegion.Name = "optRegion"
        Me.optRegion.Size = New System.Drawing.Size(146, 16)
        Me.optRegion.TabIndex = 1
        Me.optRegion.Text = "Production"
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(483, 30)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(139, 26)
        Me.btnSelect.TabIndex = 3
        Me.btnSelect.TabStop = False
        Me.btnSelect.Text = "Select All"
        '
        'btnDeSelect
        '
        Me.btnDeSelect.Location = New System.Drawing.Point(628, 30)
        Me.btnDeSelect.Name = "btnDeSelect"
        Me.btnDeSelect.Size = New System.Drawing.Size(139, 26)
        Me.btnDeSelect.TabIndex = 4
        Me.btnDeSelect.TabStop = False
        Me.btnDeSelect.Text = "Deselect All"
        '
        'ASFDEPL1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(910, 532)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ASFDEPL1"
        Me.Text = "ABSolution Deployment Utility"
        CType(Me.grdDLLS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.cmbClient, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optRegion, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdDLLS As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblUSER_ID As Misc.UltraLabel
    Friend WithEvents cmbClient As UltraWinGrid.UltraCombo
    Friend WithEvents cmdDeploy As Misc.UltraButton
    Friend WithEvents optRegion As UltraWinEditors.UltraOptionSet
    Friend WithEvents btnDeSelect As Misc.UltraButton
    Friend WithEvents btnSelect As Misc.UltraButton
End Class
