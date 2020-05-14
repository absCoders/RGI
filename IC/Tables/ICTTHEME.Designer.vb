<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ICTTHEME
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFCODEM
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
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ICTTHEMX", -1)
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("THEME_DESC", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("USAGE")
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SEASON_CODE_MIN")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SEASON_CODE_MAX")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("THEME_CODE_MIN")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("THEME_CODE_MAX")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("THEME_NO")
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor2 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor3 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.chkThemeGenerator = New ABSCS.ABSCheckBox()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraNumericEditor2 = New Infragistics.Win.UltraWinEditors.UltraNumericEditor()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.grdICTTHEMX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.cmdGetXLS = New Infragistics.Win.Misc.UltraButton()
        Me.chkDeleteSeason = New ABSCS.ABSCheckBox()
        Me.cmdGenerateThemeCodes = New Infragistics.Win.Misc.UltraButton()
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtSeason = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
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
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkThemeGenerator, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraNumericEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.grdICTTHEMX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkDeleteSeason, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSeason, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.SplitContainer1)
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
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Appearance20.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance20
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance21.BackColor = System.Drawing.SystemColors.Window
        Appearance21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance21
        Appearance22.BackColor = System.Drawing.SystemColors.Highlight
        Appearance22.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance22
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance23.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance23
        Appearance24.BorderColor = System.Drawing.Color.Silver
        Appearance24.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance24
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance25.BackColor = System.Drawing.SystemColors.Control
        Appearance25.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance25.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance25.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance25
        Appearance26.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance26
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance27.BackColor = System.Drawing.SystemColors.Window
        Appearance27.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance27
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance28.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance28
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
        'UltraLabel2
        '
        Appearance4.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel2.Appearance = Appearance4
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(12, 38)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(80, 18)
        Me.UltraLabel2.TabIndex = 23
        Me.UltraLabel2.Text = "Description"
        '
        'UltraLabel1
        '
        Appearance3.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel1.Appearance = Appearance3
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(11, 8)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(89, 18)
        Me.UltraLabel1.TabIndex = 22
        Me.UltraLabel1.Text = "Theme Code"
        '
        'UltraTextEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor2, "THEME_DESC")
        Me.UltraTextEditor2.Location = New System.Drawing.Point(134, 34)
        Me.UltraTextEditor2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor2.Name = "UltraTextEditor2"
        Me.UltraTextEditor2.Size = New System.Drawing.Size(391, 25)
        Me.UltraTextEditor2.TabIndex = 21
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "THEME_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(134, 4)
        Me.UltraTextEditor1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(108, 25)
        Me.UltraTextEditor1.TabIndex = 20
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraLabel4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraTextEditor3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkThemeGenerator)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraLabel3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraNumericEditor2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraLabel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraTextEditor2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraLabel2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(772, 507)
        Me.SplitContainer1.SplitterDistance = 123
        Me.SplitContainer1.TabIndex = 115
        '
        'UltraLabel4
        '
        Appearance1.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel4.Appearance = Appearance1
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(12, 98)
        Me.UltraLabel4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(54, 18)
        Me.UltraLabel4.TabIndex = 259
        Me.UltraLabel4.Text = "Season"
        '
        'UltraTextEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor3, "SEASON_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor3, True)
        Me.Absx1.SetABSViewName(Me.UltraTextEditor3, "SEASON_CODE")
        Me.UltraTextEditor3.Location = New System.Drawing.Point(134, 94)
        Me.UltraTextEditor3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor3.Name = "UltraTextEditor3"
        Me.UltraTextEditor3.Size = New System.Drawing.Size(152, 25)
        Me.UltraTextEditor3.TabIndex = 258
        '
        'chkThemeGenerator
        '
        Me.chkThemeGenerator.Location = New System.Drawing.Point(248, 7)
        Me.chkThemeGenerator.Name = "chkThemeGenerator"
        Me.chkThemeGenerator.Size = New System.Drawing.Size(190, 20)
        Me.chkThemeGenerator.TabIndex = 257
        Me.chkThemeGenerator.Text = "Show Theme Generator"
        '
        'UltraLabel3
        '
        Appearance2.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel3.Appearance = Appearance2
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(12, 69)
        Me.UltraLabel3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(109, 18)
        Me.UltraLabel3.TabIndex = 198
        Me.UltraLabel3.Text = "Theme Number"
        '
        'UltraNumericEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraNumericEditor2, "THEME_NO")
        Me.UltraNumericEditor2.AlwaysInEditMode = True
        Me.UltraNumericEditor2.Location = New System.Drawing.Point(134, 64)
        Me.UltraNumericEditor2.MaxValue = 100.0R
        Me.UltraNumericEditor2.MinValue = 0
        Me.UltraNumericEditor2.Name = "UltraNumericEditor2"
        Me.UltraNumericEditor2.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.UltraNumericEditor2.Size = New System.Drawing.Size(58, 25)
        Me.UltraNumericEditor2.TabIndex = 197
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.grdICTTHEMX)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.cmdGetXLS)
        Me.SplitContainer2.Panel2.Controls.Add(Me.chkDeleteSeason)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cmdGenerateThemeCodes)
        Me.SplitContainer2.Panel2.Controls.Add(Me.UltraLabel5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.txtSeason)
        Me.SplitContainer2.Size = New System.Drawing.Size(772, 380)
        Me.SplitContainer2.SplitterDistance = 527
        Me.SplitContainer2.TabIndex = 303
        '
        'grdICTTHEMX
        '
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Appearance5.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdICTTHEMX.DisplayLayout.Appearance = Appearance5
        UltraGridColumn5.Format = "#,000"
        UltraGridColumn5.Header.Caption = "Theme Description"
        UltraGridColumn5.Header.VisiblePosition = 0
        UltraGridColumn5.Width = 248
        Appearance6.TextHAlignAsString = "Right"
        UltraGridColumn1.CellAppearance = Appearance6
        Appearance7.TextHAlignAsString = "Right"
        UltraGridColumn1.Header.Appearance = Appearance7
        UltraGridColumn1.Header.Caption = "# SC"
        UltraGridColumn1.Header.VisiblePosition = 5
        UltraGridColumn1.Width = 72
        UltraGridColumn3.Header.Caption = "Season Min"
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Width = 90
        UltraGridColumn4.Header.Caption = "Season Max"
        UltraGridColumn4.Header.VisiblePosition = 2
        UltraGridColumn4.Width = 95
        UltraGridColumn7.Header.Caption = "Theme Code Min"
        UltraGridColumn7.Header.VisiblePosition = 3
        UltraGridColumn7.Width = 127
        UltraGridColumn8.Header.Caption = "Theme Code Min"
        UltraGridColumn8.Header.VisiblePosition = 4
        UltraGridColumn6.Header.Caption = "Theme No"
        UltraGridColumn6.Header.VisiblePosition = 6
        UltraGridColumn6.Width = 92
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn5, UltraGridColumn1, UltraGridColumn3, UltraGridColumn4, UltraGridColumn7, UltraGridColumn8, UltraGridColumn6})
        Me.grdICTTHEMX.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdICTTHEMX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance8.TextHAlignAsString = "Left"
        Me.grdICTTHEMX.DisplayLayout.CaptionAppearance = Appearance8
        Appearance9.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTTHEMX.DisplayLayout.GroupByBox.Appearance = Appearance9
        Appearance10.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTTHEMX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance10
        Me.grdICTTHEMX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdICTTHEMX.DisplayLayout.GroupByBox.Hidden = True
        Appearance11.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance11.BackColor2 = System.Drawing.SystemColors.Control
        Appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance11.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTTHEMX.DisplayLayout.GroupByBox.PromptAppearance = Appearance11
        Me.grdICTTHEMX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdICTTHEMX.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdICTTHEMX.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdICTTHEMX.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance12.BackColor = System.Drawing.SystemColors.Window
        Appearance12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdICTTHEMX.DisplayLayout.Override.ActiveCellAppearance = Appearance12
        Me.grdICTTHEMX.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdICTTHEMX.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTTHEMX.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTTHEMX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdICTTHEMX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Me.grdICTTHEMX.DisplayLayout.Override.CardAreaAppearance = Appearance13
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Appearance14.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdICTTHEMX.DisplayLayout.Override.CellAppearance = Appearance14
        Me.grdICTTHEMX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdICTTHEMX.DisplayLayout.Override.CellPadding = 0
        Appearance15.BackColor = System.Drawing.SystemColors.Control
        Appearance15.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance15.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance15.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTTHEMX.DisplayLayout.Override.GroupByRowAppearance = Appearance15
        Appearance16.TextHAlignAsString = "Left"
        Me.grdICTTHEMX.DisplayLayout.Override.HeaderAppearance = Appearance16
        Me.grdICTTHEMX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdICTTHEMX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance17.BackColor = System.Drawing.SystemColors.Window
        Appearance17.BorderColor = System.Drawing.Color.Silver
        Me.grdICTTHEMX.DisplayLayout.Override.RowAppearance = Appearance17
        Me.grdICTTHEMX.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.grdICTTHEMX.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
        Appearance18.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdICTTHEMX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance18
        Me.grdICTTHEMX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdICTTHEMX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdICTTHEMX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdICTTHEMX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTTHEMX.Location = New System.Drawing.Point(0, 0)
        Me.grdICTTHEMX.Name = "grdICTTHEMX"
        Me.grdICTTHEMX.Size = New System.Drawing.Size(527, 380)
        Me.grdICTTHEMX.TabIndex = 302
        Me.grdICTTHEMX.Text = "All Theme Codes"
        '
        'cmdGetXLS
        '
        Me.cmdGetXLS.Location = New System.Drawing.Point(13, 177)
        Me.cmdGetXLS.Name = "cmdGetXLS"
        Me.cmdGetXLS.Size = New System.Drawing.Size(214, 27)
        Me.cmdGetXLS.TabIndex = 259
        Me.cmdGetXLS.Text = "Get XLS"
        '
        'chkDeleteSeason
        '
        Me.chkDeleteSeason.Location = New System.Drawing.Point(13, 69)
        Me.chkDeleteSeason.Name = "chkDeleteSeason"
        Me.chkDeleteSeason.Size = New System.Drawing.Size(190, 20)
        Me.chkDeleteSeason.TabIndex = 258
        Me.chkDeleteSeason.Text = "Delete Season"
        '
        'cmdGenerateThemeCodes
        '
        Me.cmdGenerateThemeCodes.Location = New System.Drawing.Point(13, 36)
        Me.cmdGenerateThemeCodes.Name = "cmdGenerateThemeCodes"
        Me.cmdGenerateThemeCodes.Size = New System.Drawing.Size(214, 27)
        Me.cmdGenerateThemeCodes.TabIndex = 252
        Me.cmdGenerateThemeCodes.Text = "Generate Theme Codes"
        '
        'UltraLabel5
        '
        Appearance19.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel5.Appearance = Appearance19
        Me.UltraLabel5.AutoSize = True
        Me.UltraLabel5.Location = New System.Drawing.Point(13, 8)
        Me.UltraLabel5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(54, 18)
        Me.UltraLabel5.TabIndex = 202
        Me.UltraLabel5.Text = "Season"
        '
        'txtSeason
        '
        Me.Absx1.SetABSBindToTable(Me.txtSeason, False)
        Me.Absx1.SetABSColumnName(Me.txtSeason, "SEASON_CODE2")
        Me.Absx1.SetABSHasButton(Me.txtSeason, True)
        Me.Absx1.SetABSViewName(Me.txtSeason, "SEASON_CODE")
        Me.txtSeason.Location = New System.Drawing.Point(75, 4)
        Me.txtSeason.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtSeason.MaxLength = 5
        Me.txtSeason.Name = "txtSeason"
        Me.txtSeason.Size = New System.Drawing.Size(152, 25)
        Me.txtSeason.TabIndex = 201
        '
        'ICTTHEME
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "ICTTHEME"
        Me.Text = "ICTTHEME"
        Me.Panel1.ResumeLayout(False)
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
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkThemeGenerator, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraNumericEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.grdICTTHEMX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkDeleteSeason, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSeason, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor2 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraLabel3 As Misc.UltraLabel
    Friend WithEvents UltraNumericEditor2 As UltraWinEditors.UltraNumericEditor
    Friend WithEvents grdICTTHEMX As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtSeason As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents cmdGenerateThemeCodes As Infragistics.Win.Misc.UltraButton
    Friend WithEvents chkThemeGenerator As ABSCS.ABSCheckBox
    Friend WithEvents chkDeleteSeason As ABSCS.ABSCheckBox
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor3 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents cmdGetXLS As Infragistics.Win.Misc.UltraButton
End Class
