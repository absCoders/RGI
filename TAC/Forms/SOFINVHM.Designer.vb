<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SOFINVHM
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
        Dim Appearance86 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTINVH1_I", -1)
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_TYPE")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_NO")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NO")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_SALES")
        Dim UltraGridColumn18 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_COGS")
        Dim UltraGridColumn19 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_FREIGHT")
        Dim UltraGridColumn20 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_MISC_CHG")
        Dim UltraGridColumn21 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_TOTAL_AMOUNT")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REASON_CODE")
        Dim UltraGridColumn22 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_DATE")
        Dim UltraGridColumn116 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_UPDATED")
        Dim UltraGridColumn23 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_YYYYPP_UPDATED")
        Dim UltraGridColumn118 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_BILL_TO_CUST")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("POST_CODE")
        Dim UltraGridColumn37 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_BOL_NO")
        Dim UltraGridColumn99 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SALES_DIVISION_CODE")
        Dim UltraGridColumn30 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_PRINTED")
        Dim UltraGridColumn119 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_810_BATCH_NO")
        Dim UltraGridColumn36 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_NO_CONS")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TERM_CODE")
        Dim UltraGridColumn24 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn25 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn38 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PICK_NO")
        Dim UltraGridColumn39 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_NO_REV")
        Dim UltraGridColumn91 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_FACTOR_IND")
        Dim UltraGridColumn124 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_SURCHARGE_IND")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SREP_CODE")
        Dim UltraGridColumn42 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_COMMENT")
        Dim UltraGridColumn26 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REGISTER_XNO")
        Dim UltraGridColumn41 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_NO_REV_BY")
        Dim UltraGridColumn90 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SREP2_CODE")
        Dim UltraGridColumn125 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REVISED_CUST_STORE_NO")
        Dim UltraGridColumn126 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_REVISED_DATE")
        Dim UltraGridColumn127 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_REVISED_OPER")
        Dim UltraGridColumn128 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_RETRANSMIT_IND")
        Dim UltraGridColumn129 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORIG_CUST_STORE_NO")
        Dim UltraGridColumn130 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_TOTAL_AMOUNT_CURR")
        Dim UltraGridColumn27 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CURR_CODE")
        Dim UltraGridColumn28 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CURR_EXCH_RATE")
        Dim UltraGridColumn133 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_SALES_CURR")
        Dim UltraGridColumn137 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_FREIGHT_CURR")
        Dim UltraGridColumn140 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_MISC_CHG_CURR")
        Dim UltraGridColumn141 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_TOTAL_AMT_CURR")
        Dim UltraGridColumn142 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GST_TAX")
        Dim UltraGridColumn143 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GST_TAX_CURR")
        Dim UltraGridColumn144 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GEN_IND")
        Dim UltraGridColumn145 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GEN_XNO")
        Dim UltraGridColumn252 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GEN_DATE")
        Dim UltraGridColumn253 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DOCUMENTKEY")
        Dim UltraGridColumn254 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_DURING_PHY")
        Dim UltraGridColumn87 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DEPT")
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_TYPE_CODE")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_BILL_TO_CUST")
        Dim UltraGridColumn255 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE_TO")
        Dim UltraGridColumn256 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CC_TRANS_ID")
        Dim UltraGridColumn107 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_NO")
        Dim UltraGridColumn102 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REGISTER_IND")
        Dim UltraGridColumn88 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REGISTER_DATE")
        Dim Appearance87 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance88 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance89 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance90 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance91 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance92 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance93 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance94 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance95 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance96 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance97 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance139 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.splItems = New System.Windows.Forms.SplitContainer()
        Me.grpHeader = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grdSOTINVH1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.btnUpdate = New Infragistics.Win.Misc.UltraButton()
        Me.btnCancel = New Infragistics.Win.Misc.UltraButton()
        Me.txtTERM_DESC = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblTERM_CODE = New Infragistics.Win.Misc.UltraLabel()
        Me.txtTERM_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtSREP2_NAME = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblSREP2_CODE = New Infragistics.Win.Misc.UltraLabel()
        Me.txtSREP2_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtSREP_NAME = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblSREP_CODE = New Infragistics.Win.Misc.UltraLabel()
        Me.txtSREP_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.splItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splItems.Panel1.SuspendLayout()
        Me.splItems.Panel2.SuspendLayout()
        Me.splItems.SuspendLayout()
        CType(Me.grpHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpHeader.SuspendLayout()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox2.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grdSOTINVH1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.txtTERM_DESC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTERM_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSREP2_NAME, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSREP2_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSREP_NAME, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSREP_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.splItems)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(735, 431)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 431)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(735, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 431)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(735, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 431)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(735, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'splItems
        '
        Me.splItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splItems.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.splItems.IsSplitterFixed = True
        Me.splItems.Location = New System.Drawing.Point(0, 0)
        Me.splItems.Name = "splItems"
        Me.splItems.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splItems.Panel1
        '
        Me.splItems.Panel1.Controls.Add(Me.grpHeader)
        '
        'splItems.Panel2
        '
        Me.splItems.Panel2.Controls.Add(Me.SplitContainer1)
        Me.splItems.Size = New System.Drawing.Size(735, 431)
        Me.splItems.SplitterDistance = 100
        Me.splItems.TabIndex = 0
        '
        'grpHeader
        '
        Me.grpHeader.Controls.Add(Me.UltraGroupBox2)
        Me.grpHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpHeader.Location = New System.Drawing.Point(0, 0)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(735, 100)
        Me.grpHeader.TabIndex = 0
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.UltraGroupBox2.Controls.Add(Me.txtTERM_DESC)
        Me.UltraGroupBox2.Controls.Add(Me.lblTERM_CODE)
        Me.UltraGroupBox2.Controls.Add(Me.txtTERM_CODE)
        Me.UltraGroupBox2.Controls.Add(Me.txtSREP2_NAME)
        Me.UltraGroupBox2.Controls.Add(Me.lblSREP2_CODE)
        Me.UltraGroupBox2.Controls.Add(Me.txtSREP2_CODE)
        Me.UltraGroupBox2.Controls.Add(Me.txtSREP_NAME)
        Me.UltraGroupBox2.Controls.Add(Me.lblSREP_CODE)
        Me.UltraGroupBox2.Controls.Add(Me.txtSREP_CODE)
        Me.UltraGroupBox2.Location = New System.Drawing.Point(6, 3)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(496, 95)
        Me.UltraGroupBox2.TabIndex = 122
        Me.UltraGroupBox2.Text = "New Values"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdSOTINVH1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.UltraGroupBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(735, 327)
        Me.SplitContainer1.SplitterDistance = 288
        Me.SplitContainer1.TabIndex = 168
        '
        'grdSOTINVH1
        '
        Appearance86.BackColor = System.Drawing.SystemColors.Window
        Appearance86.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTINVH1.DisplayLayout.Appearance = Appearance86
        UltraGridColumn3.Header.Caption = "T"
        UltraGridColumn3.Header.VisiblePosition = 0
        UltraGridColumn3.Width = 23
        UltraGridColumn4.Header.Caption = "Invoice No"
        UltraGridColumn4.Header.VisiblePosition = 1
        UltraGridColumn4.Width = 111
        UltraGridColumn5.Header.Caption = "Customer"
        UltraGridColumn5.Header.VisiblePosition = 2
        UltraGridColumn6.Header.Caption = "Store"
        UltraGridColumn6.Header.VisiblePosition = 3
        UltraGridColumn6.Width = 71
        UltraGridColumn7.Header.Caption = "PO/Ref"
        UltraGridColumn7.Header.VisiblePosition = 4
        UltraGridColumn7.Width = 113
        UltraGridColumn8.Header.Caption = "Order No"
        UltraGridColumn8.Header.VisiblePosition = 5
        UltraGridColumn8.Width = 105
        UltraGridColumn9.Header.Caption = "Whse"
        UltraGridColumn9.Header.VisiblePosition = 6
        UltraGridColumn9.Width = 57
        UltraGridColumn17.Header.Caption = "Sales"
        UltraGridColumn17.Header.VisiblePosition = 9
        UltraGridColumn18.Header.VisiblePosition = 10
        UltraGridColumn18.Hidden = True
        UltraGridColumn19.Header.Caption = "Freight"
        UltraGridColumn19.Header.VisiblePosition = 11
        UltraGridColumn19.Width = 69
        UltraGridColumn20.Header.Caption = "Misc"
        UltraGridColumn20.Header.VisiblePosition = 12
        UltraGridColumn20.Width = 93
        UltraGridColumn21.Header.Caption = "Total"
        UltraGridColumn21.Header.VisiblePosition = 15
        UltraGridColumn13.Header.Caption = "RC"
        UltraGridColumn13.Header.VisiblePosition = 8
        UltraGridColumn13.Hidden = True
        UltraGridColumn13.Width = 33
        UltraGridColumn22.Header.Caption = "Inv Date"
        UltraGridColumn22.Header.VisiblePosition = 18
        UltraGridColumn22.Width = 112
        UltraGridColumn116.Header.VisiblePosition = 13
        UltraGridColumn116.Hidden = True
        UltraGridColumn23.Header.VisiblePosition = 19
        UltraGridColumn23.Hidden = True
        UltraGridColumn118.Header.VisiblePosition = 14
        UltraGridColumn118.Hidden = True
        UltraGridColumn10.Header.Caption = "Post"
        UltraGridColumn10.Header.VisiblePosition = 58
        UltraGridColumn10.Width = 59
        UltraGridColumn37.Header.Caption = "Ship BOL"
        UltraGridColumn37.Header.VisiblePosition = 25
        UltraGridColumn37.Width = 106
        UltraGridColumn99.Header.Caption = "Div"
        UltraGridColumn99.Header.VisiblePosition = 54
        UltraGridColumn99.Width = 40
        UltraGridColumn30.Header.VisiblePosition = 22
        UltraGridColumn30.Hidden = True
        UltraGridColumn119.Header.VisiblePosition = 17
        UltraGridColumn119.Hidden = True
        UltraGridColumn36.Header.VisiblePosition = 24
        UltraGridColumn36.Hidden = True
        UltraGridColumn11.Header.Caption = "Term"
        UltraGridColumn11.Header.VisiblePosition = 59
        UltraGridColumn11.Width = 59
        UltraGridColumn24.Header.Caption = "Entered"
        UltraGridColumn24.Header.VisiblePosition = 49
        UltraGridColumn25.Header.Caption = "By"
        UltraGridColumn25.Header.VisiblePosition = 51
        UltraGridColumn25.Width = 75
        UltraGridColumn38.Header.Caption = "Pick No"
        UltraGridColumn38.Header.VisiblePosition = 26
        UltraGridColumn38.Width = 108
        UltraGridColumn39.Header.Caption = "Reversing"
        UltraGridColumn39.Header.VisiblePosition = 27
        UltraGridColumn39.Width = 106
        UltraGridColumn91.Header.VisiblePosition = 38
        UltraGridColumn91.Hidden = True
        UltraGridColumn124.Header.VisiblePosition = 23
        UltraGridColumn124.Hidden = True
        UltraGridColumn12.Header.Caption = "SRep"
        UltraGridColumn12.Header.VisiblePosition = 7
        UltraGridColumn12.Width = 52
        UltraGridColumn42.Header.Caption = "Comment"
        UltraGridColumn42.Header.VisiblePosition = 31
        UltraGridColumn26.Header.Caption = "Register XNo"
        UltraGridColumn26.Header.VisiblePosition = 53
        UltraGridColumn26.Width = 111
        UltraGridColumn41.Header.Caption = "Reversed By"
        UltraGridColumn41.Header.VisiblePosition = 29
        UltraGridColumn41.Width = 107
        UltraGridColumn90.Header.Caption = "SRep2"
        UltraGridColumn90.Header.VisiblePosition = 37
        UltraGridColumn125.Header.VisiblePosition = 28
        UltraGridColumn125.Hidden = True
        UltraGridColumn126.Header.VisiblePosition = 30
        UltraGridColumn126.Hidden = True
        UltraGridColumn127.Header.VisiblePosition = 32
        UltraGridColumn127.Hidden = True
        UltraGridColumn128.Header.VisiblePosition = 33
        UltraGridColumn128.Hidden = True
        UltraGridColumn129.Header.VisiblePosition = 34
        UltraGridColumn129.Hidden = True
        UltraGridColumn130.Header.VisiblePosition = 35
        UltraGridColumn130.Hidden = True
        UltraGridColumn27.Header.VisiblePosition = 20
        UltraGridColumn27.Hidden = True
        UltraGridColumn28.Header.VisiblePosition = 21
        UltraGridColumn28.Hidden = True
        UltraGridColumn133.Header.VisiblePosition = 39
        UltraGridColumn133.Hidden = True
        UltraGridColumn137.Header.VisiblePosition = 40
        UltraGridColumn137.Hidden = True
        UltraGridColumn140.Header.VisiblePosition = 41
        UltraGridColumn140.Hidden = True
        UltraGridColumn141.Header.VisiblePosition = 42
        UltraGridColumn141.Hidden = True
        UltraGridColumn142.Header.VisiblePosition = 43
        UltraGridColumn142.Hidden = True
        UltraGridColumn143.Header.VisiblePosition = 44
        UltraGridColumn143.Hidden = True
        UltraGridColumn144.Header.VisiblePosition = 45
        UltraGridColumn144.Hidden = True
        UltraGridColumn145.Header.VisiblePosition = 46
        UltraGridColumn145.Hidden = True
        UltraGridColumn252.Header.VisiblePosition = 47
        UltraGridColumn252.Hidden = True
        UltraGridColumn253.Header.VisiblePosition = 50
        UltraGridColumn253.Hidden = True
        UltraGridColumn254.Header.VisiblePosition = 52
        UltraGridColumn254.Hidden = True
        UltraGridColumn87.Header.VisiblePosition = 36
        UltraGridColumn87.Hidden = True
        UltraGridColumn46.Header.Caption = "Type"
        UltraGridColumn46.Header.VisiblePosition = 16
        UltraGridColumn46.Width = 52
        UltraGridColumn14.Header.Caption = "Bill-To"
        UltraGridColumn14.Header.VisiblePosition = 48
        UltraGridColumn255.Header.VisiblePosition = 55
        UltraGridColumn255.Hidden = True
        UltraGridColumn256.Header.VisiblePosition = 57
        UltraGridColumn256.Hidden = True
        UltraGridColumn107.Header.Caption = "Auth No"
        UltraGridColumn107.Header.VisiblePosition = 61
        UltraGridColumn102.Header.VisiblePosition = 56
        UltraGridColumn102.Hidden = True
        UltraGridColumn88.Header.Caption = "Register Date"
        UltraGridColumn88.Header.VisiblePosition = 60
        UltraGridColumn88.Width = 114
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9, UltraGridColumn17, UltraGridColumn18, UltraGridColumn19, UltraGridColumn20, UltraGridColumn21, UltraGridColumn13, UltraGridColumn22, UltraGridColumn116, UltraGridColumn23, UltraGridColumn118, UltraGridColumn10, UltraGridColumn37, UltraGridColumn99, UltraGridColumn30, UltraGridColumn119, UltraGridColumn36, UltraGridColumn11, UltraGridColumn24, UltraGridColumn25, UltraGridColumn38, UltraGridColumn39, UltraGridColumn91, UltraGridColumn124, UltraGridColumn12, UltraGridColumn42, UltraGridColumn26, UltraGridColumn41, UltraGridColumn90, UltraGridColumn125, UltraGridColumn126, UltraGridColumn127, UltraGridColumn128, UltraGridColumn129, UltraGridColumn130, UltraGridColumn27, UltraGridColumn28, UltraGridColumn133, UltraGridColumn137, UltraGridColumn140, UltraGridColumn141, UltraGridColumn142, UltraGridColumn143, UltraGridColumn144, UltraGridColumn145, UltraGridColumn252, UltraGridColumn253, UltraGridColumn254, UltraGridColumn87, UltraGridColumn46, UltraGridColumn14, UltraGridColumn255, UltraGridColumn256, UltraGridColumn107, UltraGridColumn102, UltraGridColumn88})
        Me.grdSOTINVH1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTINVH1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance87.TextHAlignAsString = "Left"
        Me.grdSOTINVH1.DisplayLayout.CaptionAppearance = Appearance87
        Appearance88.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance88.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance88.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance88.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTINVH1.DisplayLayout.GroupByBox.Appearance = Appearance88
        Appearance89.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTINVH1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance89
        Me.grdSOTINVH1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTINVH1.DisplayLayout.GroupByBox.Hidden = True
        Appearance90.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance90.BackColor2 = System.Drawing.SystemColors.Control
        Appearance90.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance90.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTINVH1.DisplayLayout.GroupByBox.PromptAppearance = Appearance90
        Me.grdSOTINVH1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTINVH1.DisplayLayout.MaxRowScrollRegions = 1
        Appearance91.BackColor = System.Drawing.SystemColors.Window
        Appearance91.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTINVH1.DisplayLayout.Override.ActiveCellAppearance = Appearance91
        Appearance92.BackColor = System.Drawing.SystemColors.Highlight
        Appearance92.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdSOTINVH1.DisplayLayout.Override.ActiveRowAppearance = Appearance92
        Me.grdSOTINVH1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTINVH1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTINVH1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTINVH1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTINVH1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance93.BackColor = System.Drawing.SystemColors.Window
        Me.grdSOTINVH1.DisplayLayout.Override.CardAreaAppearance = Appearance93
        Appearance94.BorderColor = System.Drawing.Color.Silver
        Appearance94.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTINVH1.DisplayLayout.Override.CellAppearance = Appearance94
        Me.grdSOTINVH1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTINVH1.DisplayLayout.Override.CellPadding = 0
        Appearance95.BackColor = System.Drawing.SystemColors.Control
        Appearance95.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance95.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance95.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance95.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTINVH1.DisplayLayout.Override.GroupByRowAppearance = Appearance95
        Appearance96.TextHAlignAsString = "Left"
        Me.grdSOTINVH1.DisplayLayout.Override.HeaderAppearance = Appearance96
        Me.grdSOTINVH1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTINVH1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance97.BackColor = System.Drawing.SystemColors.Window
        Appearance97.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTINVH1.DisplayLayout.Override.RowAppearance = Appearance97
        Me.grdSOTINVH1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance139.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTINVH1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance139
        Me.grdSOTINVH1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTINVH1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTINVH1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTINVH1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTINVH1.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTINVH1.Name = "grdSOTINVH1"
        Me.grdSOTINVH1.Size = New System.Drawing.Size(735, 288)
        Me.grdSOTINVH1.TabIndex = 3
        Me.grdSOTINVH1.Text = "Invoices"
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.btnUpdate)
        Me.UltraGroupBox1.Controls.Add(Me.btnCancel)
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(735, 35)
        Me.UltraGroupBox1.TabIndex = 0
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(554, 2)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(81, 31)
        Me.btnUpdate.TabIndex = 117
        Me.btnUpdate.Text = "Update"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(641, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(81, 31)
        Me.btnCancel.TabIndex = 118
        Me.btnCancel.TabStop = False
        Me.btnCancel.Text = "Cancel"
        '
        'txtTERM_DESC
        '
        Me.Absx1.SetABSColumnName(Me.txtTERM_DESC, "TERM_DESC")
        Me.Absx1.SetABSParentColumnName(Me.txtTERM_DESC, "TERM_CODE")
        Me.txtTERM_DESC.Location = New System.Drawing.Point(193, 65)
        Me.txtTERM_DESC.Name = "txtTERM_DESC"
        Me.txtTERM_DESC.ReadOnly = True
        Me.txtTERM_DESC.Size = New System.Drawing.Size(228, 25)
        Me.txtTERM_DESC.TabIndex = 186
        '
        'lblTERM_CODE
        '
        Me.lblTERM_CODE.AutoSize = True
        Me.lblTERM_CODE.Location = New System.Drawing.Point(6, 69)
        Me.lblTERM_CODE.Name = "lblTERM_CODE"
        Me.lblTERM_CODE.Size = New System.Drawing.Size(47, 18)
        Me.lblTERM_CODE.TabIndex = 185
        Me.lblTERM_CODE.Text = "Terms"
        '
        'txtTERM_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtTERM_CODE, "TERM_CODE")
        Me.Absx1.SetABSHasButton(Me.txtTERM_CODE, True)
        Me.txtTERM_CODE.Location = New System.Drawing.Point(91, 65)
        Me.txtTERM_CODE.Name = "txtTERM_CODE"
        Me.txtTERM_CODE.Size = New System.Drawing.Size(103, 25)
        Me.txtTERM_CODE.TabIndex = 184
        '
        'txtSREP2_NAME
        '
        Me.Absx1.SetABSColumnName(Me.txtSREP2_NAME, "SREP_NAME")
        Me.Absx1.SetABSParentColumnName(Me.txtSREP2_NAME, "SREP2_CODE")
        Me.Absx1.SetABSViewName(Me.txtSREP2_NAME, "SOTSREP1")
        Me.txtSREP2_NAME.Location = New System.Drawing.Point(193, 41)
        Me.txtSREP2_NAME.Name = "txtSREP2_NAME"
        Me.txtSREP2_NAME.ReadOnly = True
        Me.txtSREP2_NAME.Size = New System.Drawing.Size(228, 25)
        Me.txtSREP2_NAME.TabIndex = 183
        '
        'lblSREP2_CODE
        '
        Me.lblSREP2_CODE.AutoSize = True
        Me.lblSREP2_CODE.Location = New System.Drawing.Point(6, 45)
        Me.lblSREP2_CODE.Name = "lblSREP2_CODE"
        Me.lblSREP2_CODE.Size = New System.Drawing.Size(84, 18)
        Me.lblSREP2_CODE.TabIndex = 182
        Me.lblSREP2_CODE.Text = "Sales Rep 2"
        '
        'txtSREP2_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtSREP2_CODE, "SREP2_CODE")
        Me.Absx1.SetABSHasButton(Me.txtSREP2_CODE, True)
        Me.txtSREP2_CODE.Location = New System.Drawing.Point(91, 41)
        Me.txtSREP2_CODE.Name = "txtSREP2_CODE"
        Me.txtSREP2_CODE.Size = New System.Drawing.Size(103, 25)
        Me.txtSREP2_CODE.TabIndex = 181
        '
        'txtSREP_NAME
        '
        Me.Absx1.SetABSColumnName(Me.txtSREP_NAME, "SREP_NAME")
        Me.Absx1.SetABSParentColumnName(Me.txtSREP_NAME, "SREP_CODE")
        Me.txtSREP_NAME.Location = New System.Drawing.Point(193, 17)
        Me.txtSREP_NAME.Name = "txtSREP_NAME"
        Me.txtSREP_NAME.ReadOnly = True
        Me.txtSREP_NAME.Size = New System.Drawing.Size(228, 25)
        Me.txtSREP_NAME.TabIndex = 180
        '
        'lblSREP_CODE
        '
        Me.lblSREP_CODE.AutoSize = True
        Me.lblSREP_CODE.Location = New System.Drawing.Point(6, 21)
        Me.lblSREP_CODE.Name = "lblSREP_CODE"
        Me.lblSREP_CODE.Size = New System.Drawing.Size(84, 18)
        Me.lblSREP_CODE.TabIndex = 179
        Me.lblSREP_CODE.Text = "Sales Rep 1"
        '
        'txtSREP_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtSREP_CODE, "SREP_CODE")
        Me.Absx1.SetABSHasButton(Me.txtSREP_CODE, True)
        Me.txtSREP_CODE.Location = New System.Drawing.Point(91, 17)
        Me.txtSREP_CODE.Name = "txtSREP_CODE"
        Me.txtSREP_CODE.Size = New System.Drawing.Size(103, 25)
        Me.txtSREP_CODE.TabIndex = 178
        '
        'SOFINVHM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(735, 431)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "SOFINVHM"
        Me.Text = "Invoice Maintenance"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splItems.Panel1.ResumeLayout(False)
        Me.splItems.Panel2.ResumeLayout(False)
        CType(Me.splItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splItems.ResumeLayout(False)
        CType(Me.grpHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpHeader.ResumeLayout(False)
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox2.ResumeLayout(False)
        Me.UltraGroupBox2.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grdSOTINVH1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        CType(Me.txtTERM_DESC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTERM_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSREP2_NAME, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSREP2_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSREP_NAME, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSREP_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents splItems As System.Windows.Forms.SplitContainer
    Friend WithEvents grpHeader As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents btnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents btnUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents grdSOTINVH1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents txtTERM_DESC As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblTERM_CODE As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtTERM_CODE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtSREP2_NAME As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblSREP2_CODE As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtSREP2_CODE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtSREP_NAME As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblSREP_CODE As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtSREP_CODE As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
