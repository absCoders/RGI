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
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Assemblies", -1)
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
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CLIENT")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IPADDRESS_PROD")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IPADDRESS_TEST")
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
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.grdDLLS = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lblClientIP = New Infragistics.Win.Misc.UltraLabel()
        Me.btnDeSelect = New Infragistics.Win.Misc.UltraButton()
        Me.btnSelect = New Infragistics.Win.Misc.UltraButton()
        Me.optRegion = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.cmdDeploy = New Infragistics.Win.Misc.UltraButton()
        Me.cmbClient = New Infragistics.Win.UltraWinGrid.UltraCombo()
        Me.lblUSER_ID = New Infragistics.Win.Misc.UltraLabel()
        CType(Me.grdDLLS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.optRegion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbClient, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblClientIP)
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
        'lblClientIP
        '
        Appearance13.ForeColor = System.Drawing.Color.Red
        Me.lblClientIP.Appearance = Appearance13
        Me.lblClientIP.AutoSize = True
        Me.lblClientIP.Location = New System.Drawing.Point(157, 35)
        Me.lblClientIP.Margin = New System.Windows.Forms.Padding(4)
        Me.lblClientIP.Name = "lblClientIP"
        Me.lblClientIP.Size = New System.Drawing.Size(120, 18)
        Me.lblClientIP.TabIndex = 15
        Me.lblClientIP.Text = "Client IP Address"
        '
        'btnDeSelect
        '
        Me.btnDeSelect.Location = New System.Drawing.Point(759, 30)
        Me.btnDeSelect.Name = "btnDeSelect"
        Me.btnDeSelect.Size = New System.Drawing.Size(139, 26)
        Me.btnDeSelect.TabIndex = 4
        Me.btnDeSelect.TabStop = False
        Me.btnDeSelect.Text = "Deselect All"
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(614, 30)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(139, 26)
        Me.btnSelect.TabIndex = 3
        Me.btnSelect.TabStop = False
        Me.btnSelect.Text = "Select All"
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
        Me.optRegion.Location = New System.Drawing.Point(156, 12)
        Me.optRegion.Name = "optRegion"
        Me.optRegion.Size = New System.Drawing.Size(146, 16)
        Me.optRegion.TabIndex = 1
        Me.optRegion.Text = "Production"
        '
        'cmdDeploy
        '
        Me.cmdDeploy.Location = New System.Drawing.Point(469, 30)
        Me.cmdDeploy.Name = "cmdDeploy"
        Me.cmdDeploy.Size = New System.Drawing.Size(139, 26)
        Me.cmdDeploy.TabIndex = 2
        Me.cmdDeploy.TabStop = False
        Me.cmdDeploy.Text = "Deploy"
        '
        'cmbClient
        '
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cmbClient.DisplayLayout.Appearance = Appearance14
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
        Appearance15.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance15.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance15.BorderColor = System.Drawing.SystemColors.Window
        Me.cmbClient.DisplayLayout.GroupByBox.Appearance = Appearance15
        Appearance16.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cmbClient.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance16
        Me.cmbClient.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance17.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance17.BackColor2 = System.Drawing.SystemColors.Control
        Appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance17.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cmbClient.DisplayLayout.GroupByBox.PromptAppearance = Appearance17
        Me.cmbClient.DisplayLayout.MaxColScrollRegions = 1
        Me.cmbClient.DisplayLayout.MaxRowScrollRegions = 1
        Appearance18.BackColor = System.Drawing.SystemColors.Window
        Appearance18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmbClient.DisplayLayout.Override.ActiveCellAppearance = Appearance18
        Appearance19.BackColor = System.Drawing.SystemColors.Highlight
        Appearance19.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cmbClient.DisplayLayout.Override.ActiveRowAppearance = Appearance19
        Me.cmbClient.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cmbClient.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Me.cmbClient.DisplayLayout.Override.CardAreaAppearance = Appearance20
        Appearance21.BorderColor = System.Drawing.Color.Silver
        Appearance21.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cmbClient.DisplayLayout.Override.CellAppearance = Appearance21
        Me.cmbClient.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cmbClient.DisplayLayout.Override.CellPadding = 0
        Appearance22.BackColor = System.Drawing.SystemColors.Control
        Appearance22.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance22.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance22.BorderColor = System.Drawing.SystemColors.Window
        Me.cmbClient.DisplayLayout.Override.GroupByRowAppearance = Appearance22
        Appearance23.TextHAlignAsString = "Left"
        Me.cmbClient.DisplayLayout.Override.HeaderAppearance = Appearance23
        Me.cmbClient.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cmbClient.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance24.BackColor = System.Drawing.SystemColors.Window
        Appearance24.BorderColor = System.Drawing.Color.Silver
        Me.cmbClient.DisplayLayout.Override.RowAppearance = Appearance24
        Me.cmbClient.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance25.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cmbClient.DisplayLayout.Override.TemplateAddRowAppearance = Appearance25
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
        Me.lblUSER_ID.Location = New System.Drawing.Point(13, 10)
        Me.lblUSER_ID.Margin = New System.Windows.Forms.Padding(4)
        Me.lblUSER_ID.Name = "lblUSER_ID"
        Me.lblUSER_ID.Size = New System.Drawing.Size(43, 18)
        Me.lblUSER_ID.TabIndex = 14
        Me.lblUSER_ID.Text = "Client"
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
        CType(Me.optRegion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbClient, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents lblClientIP As Misc.UltraLabel
End Class
