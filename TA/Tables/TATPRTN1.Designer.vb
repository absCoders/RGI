<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class TATPRTN1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFCODEM
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
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor2 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.grpExtracts = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton()
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton()
        Me.btnOpenOrders = New Infragistics.Win.Misc.UltraButton()
        Me.UltraLabel7 = New Infragistics.Win.Misc.UltraLabel()
        Me.dteEnd = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.dteStart = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton()
        Me.Panel1.SuspendLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpExtracts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpExtracts.SuspendLayout()
        CType(Me.dteEnd, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteStart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.grpExtracts)
        Me.Panel1.Controls.Add(Me.UltraLabel3)
        Me.Panel1.Controls.Add(Me.UltraTextEditor2)
        Me.Panel1.Controls.Add(Me.UltraLabel2)
        Me.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.Panel1.Size = New System.Drawing.Size(772, 507)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Margins.Bottom = 0
        Me.UltraExplorerBar1.Margins.Left = 0
        Me.UltraExplorerBar1.Margins.Right = 0
        Me.UltraExplorerBar1.Margins.Top = 0
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 554)
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(776, 574)
        '
        'grdASFBASEX
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance1
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.Highlight
        Appearance3.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance3
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance4.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance4
        Appearance5.BorderColor = System.Drawing.Color.Silver
        Appearance5.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance5
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance6.BackColor = System.Drawing.SystemColors.Control
        Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance6.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance6.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance6
        Appearance7.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance7
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance8.BackColor = System.Drawing.SystemColors.Window
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance8
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance9.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance9
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 574)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(989, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 574)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(989, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 574)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(989, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraLabel3
        '
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(16, 44)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(93, 18)
        Me.UltraLabel3.TabIndex = 13
        Me.UltraLabel3.Text = "Partner Code"
        '
        'UltraTextEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor2, "PARTNER_NAME")
        Me.UltraTextEditor2.Location = New System.Drawing.Point(122, 70)
        Me.UltraTextEditor2.Name = "UltraTextEditor2"
        Me.UltraTextEditor2.Size = New System.Drawing.Size(316, 25)
        Me.UltraTextEditor2.TabIndex = 12
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(16, 73)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(98, 18)
        Me.UltraLabel2.TabIndex = 11
        Me.UltraLabel2.Text = "Partner Name"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "PARTNER_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(122, 39)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(100, 25)
        Me.UltraTextEditor1.TabIndex = 10
        '
        'grpExtracts
        '
        Me.grpExtracts.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.grpExtracts.Controls.Add(Me.UltraButton3)
        Me.grpExtracts.Controls.Add(Me.UltraButton2)
        Me.grpExtracts.Controls.Add(Me.UltraButton1)
        Me.grpExtracts.Controls.Add(Me.btnOpenOrders)
        Me.grpExtracts.Controls.Add(Me.UltraLabel7)
        Me.grpExtracts.Controls.Add(Me.dteEnd)
        Me.grpExtracts.Controls.Add(Me.UltraLabel1)
        Me.grpExtracts.Controls.Add(Me.dteStart)
        Me.grpExtracts.Location = New System.Drawing.Point(444, 3)
        Me.grpExtracts.Name = "grpExtracts"
        Me.grpExtracts.Size = New System.Drawing.Size(311, 287)
        Me.grpExtracts.TabIndex = 16
        Me.grpExtracts.Text = "Extracts"
        Me.grpExtracts.Visible = False
        '
        'UltraButton2
        '
        Me.UltraButton2.Location = New System.Drawing.Point(21, 238)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(284, 31)
        Me.UltraButton2.TabIndex = 47
        Me.UltraButton2.TabStop = False
        Me.UltraButton2.Text = "Style Master Upload"
        Me.UltraButton2.TextRenderingMode = Infragistics.Win.TextRenderingMode.GDI
        '
        'UltraButton1
        '
        Me.UltraButton1.Location = New System.Drawing.Point(21, 141)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(284, 31)
        Me.UltraButton1.TabIndex = 46
        Me.UltraButton1.TabStop = False
        Me.UltraButton1.Text = "Finished Goods Purchase Orders Open"
        '
        'btnOpenOrders
        '
        Me.btnOpenOrders.Location = New System.Drawing.Point(21, 92)
        Me.btnOpenOrders.Name = "btnOpenOrders"
        Me.btnOpenOrders.Size = New System.Drawing.Size(284, 31)
        Me.btnOpenOrders.TabIndex = 45
        Me.btnOpenOrders.TabStop = False
        Me.btnOpenOrders.Text = "Open Sales Orders"
        '
        'UltraLabel7
        '
        Me.UltraLabel7.AutoSize = True
        Me.UltraLabel7.Location = New System.Drawing.Point(7, 57)
        Me.UltraLabel7.Name = "UltraLabel7"
        Me.UltraLabel7.Size = New System.Drawing.Size(58, 18)
        Me.UltraLabel7.TabIndex = 44
        Me.UltraLabel7.Text = "To Date"
        '
        'dteEnd
        '
        Me.dteEnd.Location = New System.Drawing.Point(130, 52)
        Me.dteEnd.Name = "dteEnd"
        Me.dteEnd.ReadOnly = True
        Me.dteEnd.Size = New System.Drawing.Size(122, 25)
        Me.dteEnd.TabIndex = 17
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(7, 28)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(76, 18)
        Me.UltraLabel1.TabIndex = 42
        Me.UltraLabel1.Text = "From Date"
        '
        'dteStart
        '
        Me.dteStart.Location = New System.Drawing.Point(130, 23)
        Me.dteStart.Name = "dteStart"
        Me.dteStart.Size = New System.Drawing.Size(122, 25)
        Me.dteStart.TabIndex = 16
        '
        'UltraButton3
        '
        Me.UltraButton3.Location = New System.Drawing.Point(21, 191)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(284, 31)
        Me.UltraButton3.TabIndex = 48
        Me.UltraButton3.TabStop = False
        Me.UltraButton3.Text = "Invoices"
        Me.UltraButton3.TextRenderingMode = Infragistics.Win.TextRenderingMode.GDI
        '
        'TATPRTN1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "TATPRTN1"
        Me.Text = "TATPRTN1"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpExtracts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpExtracts.ResumeLayout(False)
        Me.grpExtracts.PerformLayout()
        CType(Me.dteEnd, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteStart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor2 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents grpExtracts As Misc.UltraGroupBox
    Friend WithEvents btnOpenOrders As Misc.UltraButton
    Friend WithEvents UltraLabel7 As Misc.UltraLabel
    Friend WithEvents dteEnd As UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents UltraLabel1 As Misc.UltraLabel
    Friend WithEvents dteStart As UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents UltraButton1 As Misc.UltraButton
    Friend WithEvents UltraButton2 As Misc.UltraButton
    Friend WithEvents UltraButton3 As Misc.UltraButton
End Class
