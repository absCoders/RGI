<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SOFORDR2
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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDR5", -1)
        Dim UltraGridColumn23 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn24 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR_TYPE")
        Dim UltraGridColumn25 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR_CODE")
        Dim UltraGridColumn26 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn27 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR1")
        Dim UltraGridColumn85 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR2")
        Dim UltraGridColumn86 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CITY")
        Dim UltraGridColumn87 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STATE")
        Dim UltraGridColumn88 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ZIP_CODE")
        Dim UltraGridColumn89 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_COUNTRY")
        Dim UltraGridColumn90 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CONTACT")
        Dim UltraGridColumn91 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_PHONE")
        Dim UltraGridColumn92 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_EXT")
        Dim UltraGridColumn93 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_FAX")
        Dim UltraGridColumn94 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_EMAIL")
        Dim UltraGridColumn95 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SYNC_BATCH")
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.grdARTCUST2 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.grpBuyerOnfo = New Infragistics.Win.Misc.UltraGroupBox()
        Me.btnBuyer = New System.Windows.Forms.Button()
        Me.txtORDR_BUYER_EMAIL = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel11 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtORDR_BUYER_NAME = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel10 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtORDR_CATEGORY = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel9 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtORDR_MESSAGE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel()
        Me.grpType = New Infragistics.Win.Misc.UltraGroupBox()
        Me.rdoQUOTE = New System.Windows.Forms.RadioButton()
        Me.rdoORDER = New System.Windows.Forms.RadioButton()
        Me.txtFRT_TERMS = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel47 = New Infragistics.Win.Misc.UltraLabel()
        Me.chkPinDisc = New System.Windows.Forms.CheckBox()
        Me.optPinDisc2 = New System.Windows.Forms.RadioButton()
        Me.optPinDisc1 = New System.Windows.Forms.RadioButton()
        Me.lblPinDisc = New System.Windows.Forms.Label()
        Me.txtTERM_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtORDR_SHIP_INSTR = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel7 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel6 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtWHSE_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.datCANCELDATE = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.datSTARTDATE = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.txtSHIP_VIA_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtORDR_CUST_PO = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        Me.imgSTYLE = New System.Windows.Forms.PictureBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnCancel = New Infragistics.Win.Misc.UltraButton()
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
        Me.GroupBox2.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.grdARTCUST2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.grpBuyerOnfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpBuyerOnfo.SuspendLayout()
        CType(Me.txtORDR_BUYER_EMAIL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtORDR_BUYER_NAME, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtORDR_CATEGORY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtORDR_MESSAGE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpType.SuspendLayout()
        CType(Me.txtFRT_TERMS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTERM_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtORDR_SHIP_INSTR, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWHSE_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.datCANCELDATE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.datSTARTDATE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSHIP_VIA_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtORDR_CUST_PO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgSTYLE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(959, 485)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 485)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(959, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 485)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(959, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 485)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(959, 0)
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.imgSTYLE)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(959, 485)
        Me.SplitContainer1.SplitterDistance = 432
        Me.SplitContainer1.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.SplitContainer2)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(959, 432)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 19)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.grdARTCUST2)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Size = New System.Drawing.Size(953, 410)
        Me.SplitContainer2.SplitterDistance = 240
        Me.SplitContainer2.TabIndex = 0
        '
        'grdARTCUST2
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdARTCUST2.DisplayLayout.Appearance = Appearance1
        UltraGridColumn23.Header.VisiblePosition = 0
        UltraGridColumn23.Hidden = True
        UltraGridColumn24.Header.VisiblePosition = 1
        UltraGridColumn24.Hidden = True
        UltraGridColumn25.Header.Caption = "Address Code"
        UltraGridColumn25.Header.VisiblePosition = 2
        UltraGridColumn25.Width = 104
        UltraGridColumn26.Header.Caption = "Name"
        UltraGridColumn26.Header.VisiblePosition = 3
        UltraGridColumn26.Width = 185
        UltraGridColumn27.Header.Caption = "Address 1"
        UltraGridColumn27.Header.VisiblePosition = 4
        UltraGridColumn27.Width = 192
        UltraGridColumn85.Header.Caption = "Address 2"
        UltraGridColumn85.Header.VisiblePosition = 5
        UltraGridColumn86.Header.Caption = "City"
        UltraGridColumn86.Header.VisiblePosition = 6
        UltraGridColumn87.Header.Caption = "ST"
        UltraGridColumn87.Header.VisiblePosition = 7
        UltraGridColumn87.Width = 45
        UltraGridColumn88.Header.Caption = "Zip Code"
        UltraGridColumn88.Header.VisiblePosition = 8
        UltraGridColumn88.Width = 94
        UltraGridColumn89.Header.Caption = "Country"
        UltraGridColumn89.Header.VisiblePosition = 9
        UltraGridColumn89.Width = 92
        UltraGridColumn90.Header.VisiblePosition = 10
        UltraGridColumn90.Hidden = True
        UltraGridColumn91.Header.VisiblePosition = 11
        UltraGridColumn91.Hidden = True
        UltraGridColumn92.Header.VisiblePosition = 12
        UltraGridColumn92.Hidden = True
        UltraGridColumn93.Header.VisiblePosition = 13
        UltraGridColumn93.Hidden = True
        UltraGridColumn94.Header.VisiblePosition = 14
        UltraGridColumn94.Hidden = True
        UltraGridColumn95.Header.VisiblePosition = 15
        UltraGridColumn95.Hidden = True
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn23, UltraGridColumn24, UltraGridColumn25, UltraGridColumn26, UltraGridColumn27, UltraGridColumn85, UltraGridColumn86, UltraGridColumn87, UltraGridColumn88, UltraGridColumn89, UltraGridColumn90, UltraGridColumn91, UltraGridColumn92, UltraGridColumn93, UltraGridColumn94, UltraGridColumn95})
        Me.grdARTCUST2.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdARTCUST2.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdARTCUST2.DisplayLayout.CaptionAppearance = Appearance2
        Me.grdARTCUST2.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdARTCUST2.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdARTCUST2.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdARTCUST2.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdARTCUST2.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdARTCUST2.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdARTCUST2.DisplayLayout.MaxColScrollRegions = 1
        Me.grdARTCUST2.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdARTCUST2.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdARTCUST2.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdARTCUST2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdARTCUST2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdARTCUST2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdARTCUST2.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdARTCUST2.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.Color.Transparent
        Me.grdARTCUST2.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdARTCUST2.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdARTCUST2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdARTCUST2.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdARTCUST2.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdARTCUST2.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdARTCUST2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdARTCUST2.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdARTCUST2.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdARTCUST2.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdARTCUST2.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdARTCUST2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdARTCUST2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdARTCUST2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdARTCUST2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdARTCUST2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdARTCUST2.Location = New System.Drawing.Point(0, 0)
        Me.grdARTCUST2.Name = "grdARTCUST2"
        Me.grdARTCUST2.Size = New System.Drawing.Size(953, 240)
        Me.grdARTCUST2.TabIndex = 19
        Me.grdARTCUST2.Text = "Ship-To Selector"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.grpBuyerOnfo)
        Me.Panel1.Controls.Add(Me.txtORDR_CATEGORY)
        Me.Panel1.Controls.Add(Me.UltraLabel9)
        Me.Panel1.Controls.Add(Me.txtORDR_MESSAGE)
        Me.Panel1.Controls.Add(Me.UltraLabel8)
        Me.Panel1.Controls.Add(Me.grpType)
        Me.Panel1.Controls.Add(Me.txtFRT_TERMS)
        Me.Panel1.Controls.Add(Me.UltraLabel47)
        Me.Panel1.Controls.Add(Me.chkPinDisc)
        Me.Panel1.Controls.Add(Me.optPinDisc2)
        Me.Panel1.Controls.Add(Me.optPinDisc1)
        Me.Panel1.Controls.Add(Me.lblPinDisc)
        Me.Panel1.Controls.Add(Me.txtTERM_CODE)
        Me.Panel1.Controls.Add(Me.txtORDR_SHIP_INSTR)
        Me.Panel1.Controls.Add(Me.UltraLabel7)
        Me.Panel1.Controls.Add(Me.UltraLabel6)
        Me.Panel1.Controls.Add(Me.txtWHSE_CODE)
        Me.Panel1.Controls.Add(Me.UltraLabel5)
        Me.Panel1.Controls.Add(Me.UltraLabel4)
        Me.Panel1.Controls.Add(Me.datCANCELDATE)
        Me.Panel1.Controls.Add(Me.datSTARTDATE)
        Me.Panel1.Controls.Add(Me.txtSHIP_VIA_CODE)
        Me.Panel1.Controls.Add(Me.UltraLabel2)
        Me.Panel1.Controls.Add(Me.txtORDR_CUST_PO)
        Me.Panel1.Controls.Add(Me.UltraLabel1)
        Me.Panel1.Controls.Add(Me.UltraLabel3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(953, 166)
        Me.Panel1.TabIndex = 0
        '
        'grpBuyerOnfo
        '
        Me.grpBuyerOnfo.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.grpBuyerOnfo.Controls.Add(Me.btnBuyer)
        Me.grpBuyerOnfo.Controls.Add(Me.txtORDR_BUYER_EMAIL)
        Me.grpBuyerOnfo.Controls.Add(Me.UltraLabel11)
        Me.grpBuyerOnfo.Controls.Add(Me.txtORDR_BUYER_NAME)
        Me.grpBuyerOnfo.Controls.Add(Me.UltraLabel10)
        Me.grpBuyerOnfo.Location = New System.Drawing.Point(715, 58)
        Me.grpBuyerOnfo.Name = "grpBuyerOnfo"
        Me.grpBuyerOnfo.Size = New System.Drawing.Size(230, 103)
        Me.grpBuyerOnfo.TabIndex = 220
        Me.grpBuyerOnfo.Text = "Buyer Information"
        '
        'btnBuyer
        '
        Me.btnBuyer.Location = New System.Drawing.Point(56, 70)
        Me.btnBuyer.Name = "btnBuyer"
        Me.btnBuyer.Size = New System.Drawing.Size(168, 23)
        Me.btnBuyer.TabIndex = 197
        Me.btnBuyer.Text = "Select/Add Buyer"
        Me.btnBuyer.UseVisualStyleBackColor = True
        '
        'txtORDR_BUYER_EMAIL
        '
        Me.Absx1.SetABSColumnName(Me.txtORDR_BUYER_EMAIL, "ORDR_BUYER_EMAIL")
        Me.Absx1.SetABSTableName(Me.txtORDR_BUYER_EMAIL, "ICTSTYL1")
        Me.txtORDR_BUYER_EMAIL.Enabled = False
        Me.txtORDR_BUYER_EMAIL.Location = New System.Drawing.Point(56, 46)
        Me.txtORDR_BUYER_EMAIL.Name = "txtORDR_BUYER_EMAIL"
        Me.txtORDR_BUYER_EMAIL.Size = New System.Drawing.Size(168, 25)
        Me.txtORDR_BUYER_EMAIL.TabIndex = 196
        '
        'UltraLabel11
        '
        Me.UltraLabel11.AutoSize = True
        Me.UltraLabel11.Location = New System.Drawing.Point(6, 52)
        Me.UltraLabel11.Name = "UltraLabel11"
        Me.UltraLabel11.Size = New System.Drawing.Size(46, 18)
        Me.UltraLabel11.TabIndex = 195
        Me.UltraLabel11.Text = "E-Mail"
        '
        'txtORDR_BUYER_NAME
        '
        Me.Absx1.SetABSColumnName(Me.txtORDR_BUYER_NAME, "ORDR_BUYER_NAME")
        Me.Absx1.SetABSTableName(Me.txtORDR_BUYER_NAME, "ICTSTYL1")
        Me.txtORDR_BUYER_NAME.Enabled = False
        Me.txtORDR_BUYER_NAME.Location = New System.Drawing.Point(56, 22)
        Me.txtORDR_BUYER_NAME.Name = "txtORDR_BUYER_NAME"
        Me.txtORDR_BUYER_NAME.Size = New System.Drawing.Size(168, 25)
        Me.txtORDR_BUYER_NAME.TabIndex = 194
        '
        'UltraLabel10
        '
        Me.UltraLabel10.AutoSize = True
        Me.UltraLabel10.Location = New System.Drawing.Point(6, 26)
        Me.UltraLabel10.Name = "UltraLabel10"
        Me.UltraLabel10.Size = New System.Drawing.Size(44, 18)
        Me.UltraLabel10.TabIndex = 108
        Me.UltraLabel10.Text = "Name"
        '
        'txtORDR_CATEGORY
        '
        Me.Absx1.SetABSColumnName(Me.txtORDR_CATEGORY, "ORDR_CATEGORY")
        Me.Absx1.SetABSTableName(Me.txtORDR_CATEGORY, "ICTSTYL1")
        Me.txtORDR_CATEGORY.Location = New System.Drawing.Point(586, 31)
        Me.txtORDR_CATEGORY.Name = "txtORDR_CATEGORY"
        Me.txtORDR_CATEGORY.Size = New System.Drawing.Size(125, 25)
        Me.txtORDR_CATEGORY.TabIndex = 219
        '
        'UltraLabel9
        '
        Me.UltraLabel9.AutoSize = True
        Me.UltraLabel9.Location = New System.Drawing.Point(586, 13)
        Me.UltraLabel9.Name = "UltraLabel9"
        Me.UltraLabel9.Size = New System.Drawing.Size(94, 18)
        Me.UltraLabel9.TabIndex = 218
        Me.UltraLabel9.Text = "Order Theme"
        '
        'txtORDR_MESSAGE
        '
        Me.Absx1.SetABSColumnName(Me.txtORDR_MESSAGE, "ORDR_MESSAGE")
        Me.txtORDR_MESSAGE.Location = New System.Drawing.Point(7, 114)
        Me.txtORDR_MESSAGE.MaxLength = 255
        Me.txtORDR_MESSAGE.Name = "txtORDR_MESSAGE"
        Me.txtORDR_MESSAGE.Size = New System.Drawing.Size(576, 25)
        Me.txtORDR_MESSAGE.TabIndex = 216
        '
        'UltraLabel8
        '
        Me.UltraLabel8.AutoSize = True
        Me.UltraLabel8.Location = New System.Drawing.Point(7, 99)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(135, 18)
        Me.UltraLabel8.TabIndex = 217
        Me.UltraLabel8.Text = "Internal Comments"
        '
        'grpType
        '
        Me.grpType.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.grpType.Controls.Add(Me.rdoQUOTE)
        Me.grpType.Controls.Add(Me.rdoORDER)
        Me.grpType.Location = New System.Drawing.Point(589, 58)
        Me.grpType.Name = "grpType"
        Me.grpType.Size = New System.Drawing.Size(122, 77)
        Me.grpType.TabIndex = 215
        Me.grpType.Text = "Order Type"
        '
        'rdoQUOTE
        '
        Me.rdoQUOTE.AutoSize = True
        Me.rdoQUOTE.Location = New System.Drawing.Point(10, 46)
        Me.rdoQUOTE.Name = "rdoQUOTE"
        Me.rdoQUOTE.Size = New System.Drawing.Size(66, 20)
        Me.rdoQUOTE.TabIndex = 1
        Me.rdoQUOTE.Text = "Quote"
        Me.rdoQUOTE.UseVisualStyleBackColor = True
        '
        'rdoORDER
        '
        Me.rdoORDER.AutoSize = True
        Me.rdoORDER.Checked = True
        Me.rdoORDER.Location = New System.Drawing.Point(10, 22)
        Me.rdoORDER.Name = "rdoORDER"
        Me.rdoORDER.Size = New System.Drawing.Size(62, 20)
        Me.rdoORDER.TabIndex = 0
        Me.rdoORDER.TabStop = True
        Me.rdoORDER.Text = "Order"
        Me.rdoORDER.UseVisualStyleBackColor = True
        '
        'txtFRT_TERMS
        '
        Me.Absx1.SetABSColumnName(Me.txtFRT_TERMS, "FRT_TERMS")
        Me.Absx1.SetABSHasButton(Me.txtFRT_TERMS, True)
        Me.txtFRT_TERMS.Location = New System.Drawing.Point(842, 31)
        Me.txtFRT_TERMS.Name = "txtFRT_TERMS"
        Me.txtFRT_TERMS.Size = New System.Drawing.Size(103, 25)
        Me.txtFRT_TERMS.TabIndex = 213
        '
        'UltraLabel47
        '
        Me.UltraLabel47.AutoSize = True
        Me.UltraLabel47.Location = New System.Drawing.Point(842, 13)
        Me.UltraLabel47.Name = "UltraLabel47"
        Me.UltraLabel47.Size = New System.Drawing.Size(70, 18)
        Me.UltraLabel47.TabIndex = 214
        Me.UltraLabel47.Text = "Frt Terms"
        '
        'chkPinDisc
        '
        Me.chkPinDisc.AutoSize = True
        Me.chkPinDisc.Location = New System.Drawing.Point(8, 142)
        Me.chkPinDisc.Name = "chkPinDisc"
        Me.chkPinDisc.Size = New System.Drawing.Size(221, 20)
        Me.chkPinDisc.TabIndex = 212
        Me.chkPinDisc.Text = "Pin This Customers Discounts"
        Me.chkPinDisc.UseVisualStyleBackColor = True
        '
        'optPinDisc2
        '
        Me.optPinDisc2.AutoSize = True
        Me.optPinDisc2.Location = New System.Drawing.Point(634, 141)
        Me.optPinDisc2.Name = "optPinDisc2"
        Me.optPinDisc2.Size = New System.Drawing.Size(86, 20)
        Me.optPinDisc2.TabIndex = 211
        Me.optPinDisc2.TabStop = True
        Me.optPinDisc2.Text = "5-9 Case"
        Me.optPinDisc2.UseVisualStyleBackColor = True
        Me.optPinDisc2.Visible = False
        '
        'optPinDisc1
        '
        Me.optPinDisc1.AutoSize = True
        Me.optPinDisc1.Location = New System.Drawing.Point(553, 141)
        Me.optPinDisc1.Name = "optPinDisc1"
        Me.optPinDisc1.Size = New System.Drawing.Size(85, 20)
        Me.optPinDisc1.TabIndex = 210
        Me.optPinDisc1.TabStop = True
        Me.optPinDisc1.Text = "Full Case"
        Me.optPinDisc1.UseVisualStyleBackColor = True
        Me.optPinDisc1.Visible = False
        '
        'lblPinDisc
        '
        Me.lblPinDisc.AutoSize = True
        Me.lblPinDisc.Location = New System.Drawing.Point(224, 143)
        Me.lblPinDisc.Name = "lblPinDisc"
        Me.lblPinDisc.Size = New System.Drawing.Size(331, 16)
        Me.lblPinDisc.TabIndex = 209
        Me.lblPinDisc.Text = "To The Following Discounts For PVC Items ONLY:"
        Me.lblPinDisc.Visible = False
        '
        'txtTERM_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtTERM_CODE, "TERM_CODE")
        Me.Absx1.SetABSHasButton(Me.txtTERM_CODE, True)
        Me.txtTERM_CODE.Location = New System.Drawing.Point(715, 31)
        Me.txtTERM_CODE.Name = "txtTERM_CODE"
        Me.txtTERM_CODE.Size = New System.Drawing.Size(125, 25)
        Me.txtTERM_CODE.TabIndex = 208
        '
        'txtORDR_SHIP_INSTR
        '
        Me.Absx1.SetABSColumnName(Me.txtORDR_SHIP_INSTR, "ORDR_SHIP_INSTR")
        Me.txtORDR_SHIP_INSTR.Location = New System.Drawing.Point(7, 73)
        Me.txtORDR_SHIP_INSTR.MaxLength = 255
        Me.txtORDR_SHIP_INSTR.Name = "txtORDR_SHIP_INSTR"
        Me.txtORDR_SHIP_INSTR.Size = New System.Drawing.Size(576, 25)
        Me.txtORDR_SHIP_INSTR.TabIndex = 206
        '
        'UltraLabel7
        '
        Me.UltraLabel7.AutoSize = True
        Me.UltraLabel7.Location = New System.Drawing.Point(7, 58)
        Me.UltraLabel7.Name = "UltraLabel7"
        Me.UltraLabel7.Size = New System.Drawing.Size(148, 18)
        Me.UltraLabel7.TabIndex = 207
        Me.UltraLabel7.Text = "Shipping Instructions"
        '
        'UltraLabel6
        '
        Me.UltraLabel6.AutoSize = True
        Me.UltraLabel6.Location = New System.Drawing.Point(715, 12)
        Me.UltraLabel6.Name = "UltraLabel6"
        Me.UltraLabel6.Size = New System.Drawing.Size(85, 18)
        Me.UltraLabel6.TabIndex = 205
        Me.UltraLabel6.Text = "Terms Code"
        '
        'txtWHSE_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtWHSE_CODE, "WHSE_CODE")
        Me.Absx1.SetABSHasButton(Me.txtWHSE_CODE, True)
        Me.txtWHSE_CODE.Location = New System.Drawing.Point(366, 31)
        Me.txtWHSE_CODE.Name = "txtWHSE_CODE"
        Me.txtWHSE_CODE.Size = New System.Drawing.Size(88, 25)
        Me.txtWHSE_CODE.TabIndex = 203
        '
        'UltraLabel5
        '
        Me.UltraLabel5.AutoSize = True
        Me.UltraLabel5.Location = New System.Drawing.Point(366, 13)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(80, 18)
        Me.UltraLabel5.TabIndex = 204
        Me.UltraLabel5.Text = "Warehouse"
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(237, 13)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(86, 18)
        Me.UltraLabel4.TabIndex = 202
        Me.UltraLabel4.Text = "Cancel Date"
        '
        'datCANCELDATE
        '
        Me.datCANCELDATE.Location = New System.Drawing.Point(237, 30)
        Me.datCANCELDATE.Name = "datCANCELDATE"
        Me.datCANCELDATE.Size = New System.Drawing.Size(126, 25)
        Me.datCANCELDATE.TabIndex = 200
        '
        'datSTARTDATE
        '
        Me.datSTARTDATE.Location = New System.Drawing.Point(109, 30)
        Me.datSTARTDATE.Name = "datSTARTDATE"
        Me.datSTARTDATE.Size = New System.Drawing.Size(126, 25)
        Me.datSTARTDATE.TabIndex = 199
        '
        'txtSHIP_VIA_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtSHIP_VIA_CODE, "SHIP_VIA_CODE")
        Me.Absx1.SetABSHasButton(Me.txtSHIP_VIA_CODE, True)
        Me.txtSHIP_VIA_CODE.Location = New System.Drawing.Point(7, 30)
        Me.txtSHIP_VIA_CODE.Name = "txtSHIP_VIA_CODE"
        Me.txtSHIP_VIA_CODE.Size = New System.Drawing.Size(99, 25)
        Me.txtSHIP_VIA_CODE.TabIndex = 197
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(10, 13)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(60, 18)
        Me.UltraLabel2.TabIndex = 198
        Me.UltraLabel2.Text = "Ship Via"
        '
        'txtORDR_CUST_PO
        '
        Me.Absx1.SetABSColumnName(Me.txtORDR_CUST_PO, "ORDR_CUST_PO")
        Me.Absx1.SetABSTableName(Me.txtORDR_CUST_PO, "ICTSTYL1")
        Me.txtORDR_CUST_PO.Location = New System.Drawing.Point(458, 31)
        Me.txtORDR_CUST_PO.MaxLength = 20
        Me.txtORDR_CUST_PO.Name = "txtORDR_CUST_PO"
        Me.txtORDR_CUST_PO.Size = New System.Drawing.Size(125, 25)
        Me.txtORDR_CUST_PO.TabIndex = 195
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(458, 12)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(93, 18)
        Me.UltraLabel1.TabIndex = 196
        Me.UltraLabel1.Text = "Customer PO"
        '
        'UltraLabel3
        '
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(109, 12)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(75, 18)
        Me.UltraLabel3.TabIndex = 201
        Me.UltraLabel3.Text = "Start Date"
        '
        'imgSTYLE
        '
        Me.imgSTYLE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.imgSTYLE.Location = New System.Drawing.Point(0, 0)
        Me.imgSTYLE.Name = "imgSTYLE"
        Me.imgSTYLE.Size = New System.Drawing.Size(959, 432)
        Me.imgSTYLE.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.imgSTYLE.TabIndex = 0
        Me.imgSTYLE.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.cmdDone)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(959, 49)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(101, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(83, 33)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'cmdDone
        '
        Me.cmdDone.Location = New System.Drawing.Point(13, 8)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(83, 33)
        Me.cmdDone.TabIndex = 3
        Me.cmdDone.Text = "Select"
        '
        'SOFORDR2
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(959, 485)
        Me.ControlBox = False
        Me.Name = "SOFORDR2"
        Me.Text = "Select Ship To"
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
        Me.GroupBox2.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.grdARTCUST2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.grpBuyerOnfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpBuyerOnfo.ResumeLayout(False)
        Me.grpBuyerOnfo.PerformLayout()
        CType(Me.txtORDR_BUYER_EMAIL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtORDR_BUYER_NAME, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtORDR_CATEGORY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtORDR_MESSAGE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpType.ResumeLayout(False)
        Me.grpType.PerformLayout()
        CType(Me.txtFRT_TERMS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTERM_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtORDR_SHIP_INSTR, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWHSE_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.datCANCELDATE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.datSTARTDATE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSHIP_VIA_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtORDR_CUST_PO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgSTYLE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents imgSTYLE As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdDone As Infragistics.Win.Misc.UltraButton
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grdARTCUST2 As UltraWinGrid.UltraGrid
    Friend WithEvents Panel1 As Panel
    Friend WithEvents grpBuyerOnfo As Misc.UltraGroupBox
    Friend WithEvents btnBuyer As Button
    Friend WithEvents txtORDR_BUYER_EMAIL As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel11 As Misc.UltraLabel
    Friend WithEvents txtORDR_BUYER_NAME As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel10 As Misc.UltraLabel
    Friend WithEvents txtORDR_CATEGORY As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel9 As Misc.UltraLabel
    Friend WithEvents txtORDR_MESSAGE As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel8 As Misc.UltraLabel
    Friend WithEvents grpType As Misc.UltraGroupBox
    Friend WithEvents rdoQUOTE As RadioButton
    Friend WithEvents rdoORDER As RadioButton
    Friend WithEvents txtFRT_TERMS As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel47 As Misc.UltraLabel
    Friend WithEvents chkPinDisc As CheckBox
    Friend WithEvents optPinDisc2 As RadioButton
    Friend WithEvents optPinDisc1 As RadioButton
    Friend WithEvents lblPinDisc As Label
    Friend WithEvents txtTERM_CODE As UltraWinEditors.UltraTextEditor
    Friend WithEvents txtORDR_SHIP_INSTR As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel7 As Misc.UltraLabel
    Friend WithEvents UltraLabel6 As Misc.UltraLabel
    Friend WithEvents txtWHSE_CODE As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel5 As Misc.UltraLabel
    Friend WithEvents UltraLabel4 As Misc.UltraLabel
    Friend WithEvents datCANCELDATE As UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents datSTARTDATE As UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents txtSHIP_VIA_CODE As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Misc.UltraLabel
    Friend WithEvents txtORDR_CUST_PO As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Misc.UltraLabel
    Friend WithEvents UltraLabel3 As Misc.UltraLabel
End Class
