<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SOTTRCK1
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
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTTOTE1", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTE_NO")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SLOT_NO")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTE_LABEL_PRINT_IND")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTE_CLASS_CODE")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SEL")
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim ValueListItem6 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem7 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.btnPrintTruckPlacard = New Infragistics.Win.Misc.UltraButton()
        Me.grdSOTTOTE1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTextEditor10 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor11 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblTotes = New Infragistics.Win.Misc.UltraLabel()
        Me.numTotes = New Infragistics.Win.UltraWinEditors.UltraNumericEditor()
        Me.btnPrintSelected = New Infragistics.Win.Misc.UltraButton()
        Me.btnPrintToteLabels = New Infragistics.Win.Misc.UltraButton()
        Me.optTRUCK_TYPE = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtMini = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.splTruck = New System.Windows.Forms.SplitContainer()
        Me.splTote1 = New System.Windows.Forms.SplitContainer()
        Me.grpPrintToteLabels = New Infragistics.Win.Misc.UltraGroupBox()
        Me.chkOneTotePerLabel = New ABSCS.ABSCheckBox()
        Me.btnTest = New Infragistics.Win.Misc.UltraButton()
        Me.cmd27 = New Infragistics.Win.Misc.UltraButton()
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
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTTOTE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numTotes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTRUCK_TYPE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtMini, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.splTruck, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splTruck.Panel1.SuspendLayout()
        Me.splTruck.Panel2.SuspendLayout()
        Me.splTruck.SuspendLayout()
        CType(Me.splTote1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splTote1.Panel1.SuspendLayout()
        Me.splTote1.Panel2.SuspendLayout()
        Me.splTote1.SuspendLayout()
        CType(Me.grpPrintToteLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpPrintToteLabels.SuspendLayout()
        CType(Me.chkOneTotePerLabel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.splTruck)
        Me.Panel1.Size = New System.Drawing.Size(849, 593)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Location = New System.Drawing.Point(0, 20)
        Me.UltraExplorerBar1.Margins.Bottom = 0
        Me.UltraExplorerBar1.Margins.Left = 0
        Me.UltraExplorerBar1.Margins.Right = 0
        Me.UltraExplorerBar1.Margins.Top = 0
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(165, 640)
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(853, 660)
        '
        'grdASFBASEX
        '
        Appearance19.BackColor = System.Drawing.SystemColors.Window
        Appearance19.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance19
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Appearance20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance20
        Appearance21.BackColor = System.Drawing.SystemColors.Highlight
        Appearance21.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance21
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance22.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance22
        Appearance23.BorderColor = System.Drawing.Color.Silver
        Appearance23.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance23
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance24.BackColor = System.Drawing.SystemColors.Control
        Appearance24.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance24.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance24.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance24
        Appearance25.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance25
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance26.BackColor = System.Drawing.SystemColors.Window
        Appearance26.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance26
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance27.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance27
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 660)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(989, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 660)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(989, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 660)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(989, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraLabel1
        '
        Appearance4.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel1.Appearance = Appearance4
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(11, 15)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(65, 18)
        Me.UltraLabel1.TabIndex = 45
        Me.UltraLabel1.Text = "Truck No"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "TRUCK_NO")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(125, 10)
        Me.UltraTextEditor1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(96, 25)
        Me.UltraTextEditor1.TabIndex = 0
        '
        'btnPrintTruckPlacard
        '
        Appearance2.TextHAlignAsString = "Left"
        Me.btnPrintTruckPlacard.Appearance = Appearance2
        Me.btnPrintTruckPlacard.Location = New System.Drawing.Point(363, 10)
        Me.btnPrintTruckPlacard.Name = "btnPrintTruckPlacard"
        Me.btnPrintTruckPlacard.Size = New System.Drawing.Size(153, 25)
        Me.btnPrintTruckPlacard.TabIndex = 52
        Me.btnPrintTruckPlacard.TabStop = False
        Me.btnPrintTruckPlacard.Text = "Print Truck ID Label"
        '
        'grdSOTTOTE1
        '
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Appearance5.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTTOTE1.DisplayLayout.Appearance = Appearance5
        UltraGridColumn1.Header.Caption = "Tote No"
        UltraGridColumn1.Header.VisiblePosition = 1
        UltraGridColumn3.Header.Caption = "Slot No"
        UltraGridColumn3.Header.VisiblePosition = 3
        UltraGridColumn3.Width = 78
        UltraGridColumn2.Header.Caption = "Printed"
        UltraGridColumn2.Header.VisiblePosition = 4
        UltraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn2.Width = 77
        UltraGridColumn5.Header.Caption = "Class"
        UltraGridColumn5.Header.VisiblePosition = 2
        UltraGridColumn5.Width = 52
        UltraGridColumn4.Header.Caption = "Sel"
        UltraGridColumn4.Header.VisiblePosition = 0
        UltraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn4.Width = 47
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn3, UltraGridColumn2, UltraGridColumn5, UltraGridColumn4})
        Me.grdSOTTOTE1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTTOTE1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance6.TextHAlignAsString = "Left"
        Me.grdSOTTOTE1.DisplayLayout.CaptionAppearance = Appearance6
        Appearance7.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance7.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance7.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTTOTE1.DisplayLayout.GroupByBox.Appearance = Appearance7
        Appearance8.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTTOTE1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance8
        Me.grdSOTTOTE1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTTOTE1.DisplayLayout.GroupByBox.Hidden = True
        Appearance9.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance9.BackColor2 = System.Drawing.SystemColors.Control
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTTOTE1.DisplayLayout.GroupByBox.PromptAppearance = Appearance9
        Me.grdSOTTOTE1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTTOTE1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTTOTE1.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdSOTTOTE1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Appearance10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTTOTE1.DisplayLayout.Override.ActiveCellAppearance = Appearance10
        Me.grdSOTTOTE1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTTOTE1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTTOTE1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTTOTE1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Me.grdSOTTOTE1.DisplayLayout.Override.CardAreaAppearance = Appearance11
        Appearance12.BorderColor = System.Drawing.Color.Silver
        Appearance12.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTTOTE1.DisplayLayout.Override.CellAppearance = Appearance12
        Me.grdSOTTOTE1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTTOTE1.DisplayLayout.Override.CellPadding = 0
        Appearance13.BackColor = System.Drawing.SystemColors.Control
        Appearance13.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance13.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance13.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTTOTE1.DisplayLayout.Override.GroupByRowAppearance = Appearance13
        Appearance14.TextHAlignAsString = "Left"
        Me.grdSOTTOTE1.DisplayLayout.Override.HeaderAppearance = Appearance14
        Me.grdSOTTOTE1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTTOTE1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance15.BackColor = System.Drawing.SystemColors.Window
        Appearance15.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTTOTE1.DisplayLayout.Override.RowAppearance = Appearance15
        Me.grdSOTTOTE1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance16.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTTOTE1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance16
        Me.grdSOTTOTE1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTTOTE1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTTOTE1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTTOTE1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTTOTE1.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTTOTE1.Name = "grdSOTTOTE1"
        Me.grdSOTTOTE1.Size = New System.Drawing.Size(445, 466)
        Me.grdSOTTOTE1.TabIndex = 5
        Me.grdSOTTOTE1.Text = "Totes in Truck"
        '
        'UltraTextEditor10
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor10, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor10, "WHSE_DESC")
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor10, "WHSE_CODE")
        Me.UltraTextEditor10.Location = New System.Drawing.Point(219, 39)
        Me.UltraTextEditor10.Name = "UltraTextEditor10"
        Me.UltraTextEditor10.ReadOnly = True
        Me.UltraTextEditor10.Size = New System.Drawing.Size(297, 25)
        Me.UltraTextEditor10.TabIndex = 2
        Me.UltraTextEditor10.TabStop = False
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(11, 46)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(42, 18)
        Me.UltraLabel2.TabIndex = 173
        Me.UltraLabel2.Text = "Whse"
        '
        'UltraTextEditor11
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor11, "WHSE_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor11, True)
        Me.UltraTextEditor11.Location = New System.Drawing.Point(125, 39)
        Me.UltraTextEditor11.Name = "UltraTextEditor11"
        Me.UltraTextEditor11.Size = New System.Drawing.Size(96, 25)
        Me.UltraTextEditor11.TabIndex = 1
        '
        'lblTotes
        '
        Me.lblTotes.AutoSize = True
        Me.lblTotes.Location = New System.Drawing.Point(11, 101)
        Me.lblTotes.Name = "lblTotes"
        Me.lblTotes.Size = New System.Drawing.Size(41, 18)
        Me.lblTotes.TabIndex = 186
        Me.lblTotes.Text = "Totes"
        '
        'numTotes
        '
        Me.numTotes.FormatString = "#0"
        Me.numTotes.Location = New System.Drawing.Point(125, 94)
        Me.numTotes.MaxValue = 99
        Me.numTotes.MinValue = 1.0R
        Me.numTotes.Name = "numTotes"
        Me.numTotes.NullText = "0"
        Me.numTotes.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.numTotes.Size = New System.Drawing.Size(74, 25)
        Me.numTotes.TabIndex = 4
        Me.numTotes.TabStop = False
        Me.numTotes.Value = 1
        '
        'btnPrintSelected
        '
        Appearance18.TextHAlignAsString = "Left"
        Me.btnPrintSelected.Appearance = Appearance18
        Me.btnPrintSelected.Location = New System.Drawing.Point(15, 109)
        Me.btnPrintSelected.Name = "btnPrintSelected"
        Me.btnPrintSelected.Size = New System.Drawing.Size(232, 25)
        Me.btnPrintSelected.TabIndex = 8
        Me.btnPrintSelected.TabStop = False
        Me.btnPrintSelected.Text = "Print Labels for Selected Totes"
        '
        'btnPrintToteLabels
        '
        Appearance17.TextHAlignAsString = "Left"
        Me.btnPrintToteLabels.Appearance = Appearance17
        Me.btnPrintToteLabels.Location = New System.Drawing.Point(15, 78)
        Me.btnPrintToteLabels.Name = "btnPrintToteLabels"
        Me.btnPrintToteLabels.Size = New System.Drawing.Size(232, 25)
        Me.btnPrintToteLabels.TabIndex = 7
        Me.btnPrintToteLabels.TabStop = False
        Me.btnPrintToteLabels.Text = "Print Tote Labels not yet Printed"
        '
        'optTRUCK_TYPE
        '
        Me.Absx1.SetABSColumnName(Me.optTRUCK_TYPE, "TRUCK_TYPE")
        Me.optTRUCK_TYPE.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem6.DataValue = "P"
        ValueListItem6.DisplayText = "Pre-Configured"
        ValueListItem7.DataValue = "X"
        ValueListItem7.DisplayText = "Custom"
        ValueListItem1.DataValue = "R"
        ValueListItem1.DisplayText = "Regular"
        Me.optTRUCK_TYPE.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem6, ValueListItem7, ValueListItem1})
        Me.optTRUCK_TYPE.Location = New System.Drawing.Point(125, 70)
        Me.optTRUCK_TYPE.Name = "optTRUCK_TYPE"
        Me.optTRUCK_TYPE.Size = New System.Drawing.Size(280, 18)
        Me.optTRUCK_TYPE.TabIndex = 3
        '
        'UltraLabel3
        '
        Appearance3.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel3.Appearance = Appearance3
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(11, 70)
        Me.UltraLabel3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(79, 18)
        Me.UltraLabel3.TabIndex = 189
        Me.UltraLabel3.Text = "Truck Type"
        '
        'txtMini
        '
        Me.txtMini.Location = New System.Drawing.Point(15, 46)
        Me.txtMini.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtMini.Name = "txtMini"
        Me.txtMini.ReadOnly = True
        Me.txtMini.Size = New System.Drawing.Size(179, 25)
        Me.txtMini.TabIndex = 6
        Me.txtMini.TabStop = False
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(15, 23)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(149, 18)
        Me.UltraLabel4.TabIndex = 174
        Me.UltraLabel4.Text = "Mini Label IP Address"
        '
        'splTruck
        '
        Me.splTruck.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splTruck.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.splTruck.IsSplitterFixed = True
        Me.splTruck.Location = New System.Drawing.Point(0, 0)
        Me.splTruck.Name = "splTruck"
        Me.splTruck.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splTruck.Panel1
        '
        Me.splTruck.Panel1.Controls.Add(Me.cmd27)
        Me.splTruck.Panel1.Controls.Add(Me.btnPrintTruckPlacard)
        Me.splTruck.Panel1.Controls.Add(Me.UltraLabel3)
        Me.splTruck.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.splTruck.Panel1.Controls.Add(Me.optTRUCK_TYPE)
        Me.splTruck.Panel1.Controls.Add(Me.UltraLabel1)
        Me.splTruck.Panel1.Controls.Add(Me.UltraTextEditor11)
        Me.splTruck.Panel1.Controls.Add(Me.lblTotes)
        Me.splTruck.Panel1.Controls.Add(Me.UltraLabel2)
        Me.splTruck.Panel1.Controls.Add(Me.numTotes)
        Me.splTruck.Panel1.Controls.Add(Me.UltraTextEditor10)
        '
        'splTruck.Panel2
        '
        Me.splTruck.Panel2.Controls.Add(Me.splTote1)
        Me.splTruck.Size = New System.Drawing.Size(849, 593)
        Me.splTruck.SplitterDistance = 123
        Me.splTruck.TabIndex = 190
        '
        'splTote1
        '
        Me.splTote1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splTote1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.splTote1.IsSplitterFixed = True
        Me.splTote1.Location = New System.Drawing.Point(0, 0)
        Me.splTote1.Name = "splTote1"
        '
        'splTote1.Panel1
        '
        Me.splTote1.Panel1.Controls.Add(Me.grdSOTTOTE1)
        '
        'splTote1.Panel2
        '
        Me.splTote1.Panel2.Controls.Add(Me.grpPrintToteLabels)
        Me.splTote1.Size = New System.Drawing.Size(849, 466)
        Me.splTote1.SplitterDistance = 445
        Me.splTote1.TabIndex = 0
        '
        'grpPrintToteLabels
        '
        Me.grpPrintToteLabels.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.grpPrintToteLabels.Controls.Add(Me.chkOneTotePerLabel)
        Me.grpPrintToteLabels.Controls.Add(Me.btnTest)
        Me.grpPrintToteLabels.Controls.Add(Me.UltraLabel4)
        Me.grpPrintToteLabels.Controls.Add(Me.btnPrintToteLabels)
        Me.grpPrintToteLabels.Controls.Add(Me.txtMini)
        Me.grpPrintToteLabels.Controls.Add(Me.btnPrintSelected)
        Me.grpPrintToteLabels.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpPrintToteLabels.Location = New System.Drawing.Point(0, 0)
        Me.grpPrintToteLabels.Name = "grpPrintToteLabels"
        Me.grpPrintToteLabels.Size = New System.Drawing.Size(400, 466)
        Me.grpPrintToteLabels.TabIndex = 187
        Me.grpPrintToteLabels.Text = "Print Tote Labels"
        '
        'chkOneTotePerLabel
        '
        Me.Absx1.SetABSBindToTable(Me.chkOneTotePerLabel, False)
        Me.chkOneTotePerLabel.Location = New System.Drawing.Point(15, 140)
        Me.chkOneTotePerLabel.Name = "chkOneTotePerLabel"
        Me.chkOneTotePerLabel.Size = New System.Drawing.Size(191, 18)
        Me.chkOneTotePerLabel.TabIndex = 176
        Me.chkOneTotePerLabel.Text = "Print One Tote Per Label"
        '
        'btnTest
        '
        Me.btnTest.Location = New System.Drawing.Point(200, 46)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(47, 25)
        Me.btnTest.TabIndex = 175
        Me.btnTest.TabStop = False
        Me.btnTest.Text = "Test"
        '
        'cmd27
        '
        Appearance1.TextHAlignAsString = "Center"
        Me.cmd27.Appearance = Appearance1
        Me.cmd27.Location = New System.Drawing.Point(238, 94)
        Me.cmd27.Name = "cmd27"
        Me.cmd27.Size = New System.Drawing.Size(167, 25)
        Me.cmd27.TabIndex = 190
        Me.cmd27.TabStop = False
        Me.cmd27.Text = "Create 27 Totes"
        '
        'SOTTRCK1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1236, 825)
        Me.Name = "SOTTRCK1"
        Me.Text = "SOTTRCK1"
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
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTTOTE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numTotes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTRUCK_TYPE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtMini, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splTruck.Panel1.ResumeLayout(False)
        Me.splTruck.Panel1.PerformLayout()
        Me.splTruck.Panel2.ResumeLayout(False)
        CType(Me.splTruck, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splTruck.ResumeLayout(False)
        Me.splTote1.Panel1.ResumeLayout(False)
        Me.splTote1.Panel2.ResumeLayout(False)
        CType(Me.splTote1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splTote1.ResumeLayout(False)
        CType(Me.grpPrintToteLabels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpPrintToteLabels.ResumeLayout(False)
        Me.grpPrintToteLabels.PerformLayout()
        CType(Me.chkOneTotePerLabel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents btnPrintTruckPlacard As Misc.UltraButton
    Friend WithEvents grdSOTTOTE1 As UltraWinGrid.UltraGrid
    Friend WithEvents UltraTextEditor10 As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Misc.UltraLabel
    Friend WithEvents UltraTextEditor11 As UltraWinEditors.UltraTextEditor
    Friend WithEvents lblTotes As Misc.UltraLabel
    Friend WithEvents numTotes As UltraWinEditors.UltraNumericEditor
    Friend WithEvents btnPrintSelected As Misc.UltraButton
    Friend WithEvents btnPrintToteLabels As Misc.UltraButton
    Friend WithEvents UltraLabel3 As Misc.UltraLabel
    Friend WithEvents optTRUCK_TYPE As UltraWinEditors.UltraOptionSet
    Friend WithEvents txtMini As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Misc.UltraLabel
    Friend WithEvents splTruck As SplitContainer
    Friend WithEvents splTote1 As SplitContainer
    Friend WithEvents grpPrintToteLabels As Misc.UltraGroupBox
    Friend WithEvents btnTest As Misc.UltraButton
    Friend WithEvents chkOneTotePerLabel As ABSCS.ABSCheckBox
    Friend WithEvents cmd27 As Misc.UltraButton
End Class
