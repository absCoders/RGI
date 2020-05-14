<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SOFORDRC
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
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDR1", -1)
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NO")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NAME")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_FOB")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn15 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_DATE")
        Dim UltraGridColumn16 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CANCEL_DATE")
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("POST_CODE")
        Dim UltraGridColumn18 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_VIA_CODE")
        Dim UltraGridColumn19 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_INSTR")
        Dim UltraGridColumn20 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TERM_CODE")
        Dim UltraGridColumn21 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SREP_CODE")
        Dim UltraGridColumn22 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn91 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_PICK_SEQ")
        Dim UltraGridColumn92 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REASON_CODE")
        Dim UltraGridColumn93 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SALES_DIVISION_CODE")
        Dim UltraGridColumn94 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn95 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn146 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn147 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn148 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_RECD")
        Dim UltraGridColumn149 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SOURCE")
        Dim UltraGridColumn150 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DEPT")
        Dim UltraGridColumn151 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FRT_TERMS")
        Dim UltraGridColumn152 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ADDR_TYPE_ST")
        Dim UltraGridColumn153 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_BOOKED")
        Dim UltraGridColumn154 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_YYYYPP_BOOKED")
        Dim UltraGridColumn155 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_CLOSED")
        Dim UltraGridColumn156 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_YYYYPP_CLOSED")
        Dim UltraGridColumn157 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_PRIORITY")
        Dim UltraGridColumn158 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_STATUS")
        Dim UltraGridColumn159 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_JRNL_NO")
        Dim UltraGridColumn160 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_GROUP_NO")
        Dim UltraGridColumn161 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_HOLD")
        Dim UltraGridColumn162 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_HOLD_CODES")
        Dim UltraGridColumn163 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_BILL_TO_CUST")
        Dim UltraGridColumn164 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_DC_NO")
        Dim UltraGridColumn165 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_BATCH_NO")
        Dim UltraGridColumn166 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_DOC_SEQ_NO")
        Dim UltraGridColumn167 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_APPOINTMENT")
        Dim UltraGridColumn168 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_INV_COMMENT")
        Dim UltraGridColumn169 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_FACTOR_IND")
        Dim UltraGridColumn170 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_PRE_ALLOC")
        Dim UltraGridColumn171 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_MERCH_TYPE")
        Dim UltraGridColumn172 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SREP2_CODE")
        Dim UltraGridColumn173 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CURR_CODE")
        Dim UltraGridColumn174 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CURR_EXCH_RATE")
        Dim UltraGridColumn175 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_VALUE_CHANGE_REMARK")
        Dim UltraGridColumn176 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_VALUE_CHANGE_OPER")
        Dim UltraGridColumn177 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_VALUE_CHANGE_DATE")
        Dim UltraGridColumn178 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_DURING_PHY")
        Dim UltraGridColumn179 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ORIG_SHIP_DATE")
        Dim UltraGridColumn180 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ORIG_CANCEL_DATE")
        Dim UltraGridColumn181 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_HOLD_REASON")
        Dim UltraGridColumn182 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SYNC_BATCH")
        Dim UltraGridColumn85 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TORDR")
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn251 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("YEAR")
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
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
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem3 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem4 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblAuth = New Infragistics.Win.Misc.UltraLabel()
        Me.grpSOTORDRS = New System.Windows.Forms.GroupBox()
        Me.grdSOTORDRS = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.cboCUST_CREDIT_CARD_EXP_DATE = New System.Windows.Forms.ComboBox()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtCUST_CREDIT_CARD_VER_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtShowNumbers = New System.Windows.Forms.CheckBox()
        Me.lblCCNo = New Infragistics.Win.Misc.UltraLabel()
        Me.txtCUST_CREDIT_CARD_NO = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtLAST4 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtCUST_CREDIT_CARD_ADDR1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel7 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtCUST_CREDIT_CARD_ZIP_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel10 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtCUST_CREDIT_CARD_STATE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel9 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtCUST_CREDIT_CARD_CITY = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtCUST_CREDIT_CARD_NAME = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel6 = New Infragistics.Win.Misc.UltraLabel()
        Me.optCC_TYPE = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.imgSTYLE = New System.Windows.Forms.PictureBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnCustAdd = New Infragistics.Win.Misc.UltraButton()
        Me.txtCCPA_AMT = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.btnCancel = New Infragistics.Win.Misc.UltraButton()
        Me.cmdFinished = New Infragistics.Win.Misc.UltraButton()
        Me.lblAMEX = New System.Windows.Forms.Label()
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
        Me.grpSOTORDRS.SuspendLayout()
        CType(Me.grdSOTORDRS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.txtCUST_CREDIT_CARD_VER_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_CREDIT_CARD_NO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLAST4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.txtCUST_CREDIT_CARD_ADDR1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_CREDIT_CARD_ZIP_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_CREDIT_CARD_STATE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_CREDIT_CARD_CITY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_CREDIT_CARD_NAME, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optCC_TYPE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgSTYLE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.txtCCPA_AMT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Margin = New System.Windows.Forms.Padding(6)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(710, 432)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 432)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(710, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 432)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(710, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 432)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(710, 0)
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
        Me.SplitContainer1.Size = New System.Drawing.Size(710, 432)
        Me.SplitContainer1.SplitterDistance = 332
        Me.SplitContainer1.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblAMEX)
        Me.GroupBox2.Controls.Add(Me.lblAuth)
        Me.GroupBox2.Controls.Add(Me.grpSOTORDRS)
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.optCC_TYPE)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(710, 332)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'lblAuth
        '
        Appearance17.ForeColor = System.Drawing.Color.Red
        Me.lblAuth.Appearance = Appearance17
        Me.lblAuth.AutoSize = True
        Me.lblAuth.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAuth.Location = New System.Drawing.Point(206, 3)
        Me.lblAuth.Name = "lblAuth"
        Me.lblAuth.Size = New System.Drawing.Size(0, 0)
        Me.lblAuth.TabIndex = 212
        '
        'grpSOTORDRS
        '
        Me.grpSOTORDRS.Controls.Add(Me.grdSOTORDRS)
        Me.grpSOTORDRS.Location = New System.Drawing.Point(18, 229)
        Me.grpSOTORDRS.Name = "grpSOTORDRS"
        Me.grpSOTORDRS.Size = New System.Drawing.Size(649, 146)
        Me.grpSOTORDRS.TabIndex = 214
        Me.grpSOTORDRS.TabStop = False
        Me.grpSOTORDRS.Text = "Orders On Group XXXX"
        '
        'grdSOTORDRS
        '
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTORDRS.DisplayLayout.Appearance = Appearance2
        UltraGridColumn7.Header.Caption = "Order No"
        UltraGridColumn7.Header.VisiblePosition = 1
        UltraGridColumn7.Width = 107
        UltraGridColumn8.Header.Caption = "Order Date"
        UltraGridColumn8.Header.VisiblePosition = 3
        UltraGridColumn8.Hidden = True
        UltraGridColumn8.Width = 90
        UltraGridColumn9.Header.Caption = "Cust Code"
        UltraGridColumn9.Header.VisiblePosition = 5
        UltraGridColumn9.Hidden = True
        UltraGridColumn9.Width = 88
        UltraGridColumn10.Header.Caption = "Cust Name"
        UltraGridColumn10.Header.VisiblePosition = 6
        UltraGridColumn10.Width = 172
        UltraGridColumn11.Header.Caption = "Store no"
        UltraGridColumn11.Header.VisiblePosition = 7
        UltraGridColumn11.Hidden = True
        UltraGridColumn11.Width = 85
        UltraGridColumn12.Header.Caption = "Store Name"
        UltraGridColumn12.Header.VisiblePosition = 8
        UltraGridColumn12.Hidden = True
        UltraGridColumn12.Width = 178
        UltraGridColumn13.Header.VisiblePosition = 9
        UltraGridColumn13.Hidden = True
        UltraGridColumn14.Header.Caption = "Cust PO"
        UltraGridColumn14.Header.VisiblePosition = 10
        UltraGridColumn14.Width = 121
        UltraGridColumn15.Header.Caption = "Ship Date"
        UltraGridColumn15.Header.VisiblePosition = 11
        UltraGridColumn15.Width = 90
        UltraGridColumn16.Header.Caption = "Cancel Date"
        UltraGridColumn16.Header.VisiblePosition = 13
        UltraGridColumn16.Hidden = True
        UltraGridColumn16.Width = 105
        UltraGridColumn17.Header.VisiblePosition = 14
        UltraGridColumn17.Hidden = True
        UltraGridColumn18.Header.Caption = "Ship Via"
        UltraGridColumn18.Header.VisiblePosition = 18
        UltraGridColumn18.Hidden = True
        UltraGridColumn18.Width = 68
        UltraGridColumn19.Header.VisiblePosition = 15
        UltraGridColumn19.Hidden = True
        UltraGridColumn20.Header.VisiblePosition = 16
        UltraGridColumn20.Hidden = True
        UltraGridColumn21.Header.VisiblePosition = 17
        UltraGridColumn21.Hidden = True
        UltraGridColumn22.Header.Caption = "Whse"
        UltraGridColumn22.Header.VisiblePosition = 19
        UltraGridColumn22.Hidden = True
        UltraGridColumn22.Width = 46
        UltraGridColumn91.Header.VisiblePosition = 20
        UltraGridColumn91.Hidden = True
        UltraGridColumn92.Header.VisiblePosition = 21
        UltraGridColumn92.Hidden = True
        UltraGridColumn93.Header.VisiblePosition = 22
        UltraGridColumn93.Hidden = True
        UltraGridColumn94.Header.VisiblePosition = 23
        UltraGridColumn94.Hidden = True
        UltraGridColumn95.Header.VisiblePosition = 24
        UltraGridColumn95.Hidden = True
        UltraGridColumn146.Header.VisiblePosition = 25
        UltraGridColumn146.Hidden = True
        UltraGridColumn147.Header.VisiblePosition = 26
        UltraGridColumn147.Hidden = True
        UltraGridColumn148.Header.VisiblePosition = 27
        UltraGridColumn148.Hidden = True
        UltraGridColumn149.Header.VisiblePosition = 28
        UltraGridColumn149.Hidden = True
        UltraGridColumn150.Header.VisiblePosition = 29
        UltraGridColumn150.Hidden = True
        UltraGridColumn151.Header.VisiblePosition = 30
        UltraGridColumn151.Hidden = True
        UltraGridColumn152.Header.VisiblePosition = 31
        UltraGridColumn152.Hidden = True
        UltraGridColumn153.Header.VisiblePosition = 32
        UltraGridColumn153.Hidden = True
        UltraGridColumn154.Header.VisiblePosition = 33
        UltraGridColumn154.Hidden = True
        UltraGridColumn155.Header.VisiblePosition = 34
        UltraGridColumn155.Hidden = True
        UltraGridColumn156.Header.VisiblePosition = 35
        UltraGridColumn156.Hidden = True
        UltraGridColumn157.Header.VisiblePosition = 36
        UltraGridColumn157.Hidden = True
        UltraGridColumn158.Header.Caption = "Status"
        UltraGridColumn158.Header.VisiblePosition = 2
        UltraGridColumn158.Hidden = True
        UltraGridColumn158.Width = 82
        UltraGridColumn159.Header.VisiblePosition = 37
        UltraGridColumn159.Hidden = True
        UltraGridColumn160.Header.Caption = "Group"
        UltraGridColumn160.Header.VisiblePosition = 0
        UltraGridColumn160.Hidden = True
        UltraGridColumn160.Width = 98
        UltraGridColumn161.Header.VisiblePosition = 38
        UltraGridColumn161.Hidden = True
        UltraGridColumn162.Header.VisiblePosition = 39
        UltraGridColumn162.Hidden = True
        UltraGridColumn163.Header.VisiblePosition = 40
        UltraGridColumn163.Hidden = True
        UltraGridColumn164.Header.VisiblePosition = 41
        UltraGridColumn164.Hidden = True
        UltraGridColumn165.Header.VisiblePosition = 42
        UltraGridColumn165.Hidden = True
        UltraGridColumn166.Header.VisiblePosition = 43
        UltraGridColumn166.Hidden = True
        UltraGridColumn167.Header.VisiblePosition = 44
        UltraGridColumn167.Hidden = True
        UltraGridColumn168.Header.VisiblePosition = 45
        UltraGridColumn168.Hidden = True
        UltraGridColumn169.Header.VisiblePosition = 46
        UltraGridColumn169.Hidden = True
        UltraGridColumn170.Header.VisiblePosition = 47
        UltraGridColumn170.Hidden = True
        UltraGridColumn171.Header.VisiblePosition = 48
        UltraGridColumn171.Hidden = True
        UltraGridColumn172.Header.VisiblePosition = 49
        UltraGridColumn172.Hidden = True
        UltraGridColumn173.Header.VisiblePosition = 50
        UltraGridColumn173.Hidden = True
        UltraGridColumn174.Header.VisiblePosition = 51
        UltraGridColumn174.Hidden = True
        UltraGridColumn175.Header.VisiblePosition = 52
        UltraGridColumn175.Hidden = True
        UltraGridColumn176.Header.VisiblePosition = 53
        UltraGridColumn176.Hidden = True
        UltraGridColumn177.Header.VisiblePosition = 54
        UltraGridColumn177.Hidden = True
        UltraGridColumn178.Header.VisiblePosition = 55
        UltraGridColumn178.Hidden = True
        UltraGridColumn179.Header.VisiblePosition = 56
        UltraGridColumn179.Hidden = True
        UltraGridColumn180.Header.VisiblePosition = 57
        UltraGridColumn180.Hidden = True
        UltraGridColumn181.Header.VisiblePosition = 58
        UltraGridColumn181.Hidden = True
        UltraGridColumn182.Header.VisiblePosition = 59
        UltraGridColumn182.Hidden = True
        Appearance3.TextHAlignAsString = "Right"
        UltraGridColumn85.CellAppearance = Appearance3
        Appearance4.TextHAlignAsString = "Right"
        UltraGridColumn85.Header.Appearance = Appearance4
        UltraGridColumn85.Header.Caption = "Total"
        UltraGridColumn85.Header.VisiblePosition = 12
        UltraGridColumn85.Width = 117
        UltraGridColumn251.Header.Caption = "Year"
        UltraGridColumn251.Header.VisiblePosition = 4
        UltraGridColumn251.Hidden = True
        UltraGridColumn251.Width = 68
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn7, UltraGridColumn8, UltraGridColumn9, UltraGridColumn10, UltraGridColumn11, UltraGridColumn12, UltraGridColumn13, UltraGridColumn14, UltraGridColumn15, UltraGridColumn16, UltraGridColumn17, UltraGridColumn18, UltraGridColumn19, UltraGridColumn20, UltraGridColumn21, UltraGridColumn22, UltraGridColumn91, UltraGridColumn92, UltraGridColumn93, UltraGridColumn94, UltraGridColumn95, UltraGridColumn146, UltraGridColumn147, UltraGridColumn148, UltraGridColumn149, UltraGridColumn150, UltraGridColumn151, UltraGridColumn152, UltraGridColumn153, UltraGridColumn154, UltraGridColumn155, UltraGridColumn156, UltraGridColumn157, UltraGridColumn158, UltraGridColumn159, UltraGridColumn160, UltraGridColumn161, UltraGridColumn162, UltraGridColumn163, UltraGridColumn164, UltraGridColumn165, UltraGridColumn166, UltraGridColumn167, UltraGridColumn168, UltraGridColumn169, UltraGridColumn170, UltraGridColumn171, UltraGridColumn172, UltraGridColumn173, UltraGridColumn174, UltraGridColumn175, UltraGridColumn176, UltraGridColumn177, UltraGridColumn178, UltraGridColumn179, UltraGridColumn180, UltraGridColumn181, UltraGridColumn182, UltraGridColumn85, UltraGridColumn251})
        Me.grdSOTORDRS.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTORDRS.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance5.TextHAlignAsString = "Left"
        Me.grdSOTORDRS.DisplayLayout.CaptionAppearance = Appearance5
        Me.grdSOTORDRS.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance6.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance6.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDRS.DisplayLayout.GroupByBox.Appearance = Appearance6
        Appearance7.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDRS.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance7
        Me.grdSOTORDRS.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTORDRS.DisplayLayout.GroupByBox.Hidden = True
        Appearance8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance8.BackColor2 = System.Drawing.SystemColors.Control
        Appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance8.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDRS.DisplayLayout.GroupByBox.PromptAppearance = Appearance8
        Me.grdSOTORDRS.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTORDRS.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTORDRS.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance9.BackColor = System.Drawing.SystemColors.Window
        Appearance9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTORDRS.DisplayLayout.Override.ActiveCellAppearance = Appearance9
        Me.grdSOTORDRS.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTORDRS.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDRS.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDRS.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTORDRS.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance10.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTORDRS.DisplayLayout.Override.CardAreaAppearance = Appearance10
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Appearance11.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTORDRS.DisplayLayout.Override.CellAppearance = Appearance11
        Me.grdSOTORDRS.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdSOTORDRS.DisplayLayout.Override.CellPadding = 0
        Appearance12.BackColor = System.Drawing.SystemColors.Control
        Appearance12.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance12.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance12.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDRS.DisplayLayout.Override.GroupByRowAppearance = Appearance12
        Appearance13.TextHAlignAsString = "Left"
        Me.grdSOTORDRS.DisplayLayout.Override.HeaderAppearance = Appearance13
        Me.grdSOTORDRS.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTORDRS.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTORDRS.DisplayLayout.Override.RowAppearance = Appearance14
        Me.grdSOTORDRS.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance15.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTORDRS.DisplayLayout.Override.TemplateAddRowAppearance = Appearance15
        Me.grdSOTORDRS.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTORDRS.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTORDRS.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTORDRS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTORDRS.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTORDRS.Location = New System.Drawing.Point(3, 19)
        Me.grdSOTORDRS.Name = "grdSOTORDRS"
        Me.grdSOTORDRS.Size = New System.Drawing.Size(643, 124)
        Me.grdSOTORDRS.TabIndex = 14
        Me.grdSOTORDRS.Text = "Orders"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cboCUST_CREDIT_CARD_EXP_DATE)
        Me.GroupBox4.Controls.Add(Me.UltraLabel2)
        Me.GroupBox4.Controls.Add(Me.txtCUST_CREDIT_CARD_VER_CODE)
        Me.GroupBox4.Controls.Add(Me.UltraLabel5)
        Me.GroupBox4.Controls.Add(Me.txtShowNumbers)
        Me.GroupBox4.Controls.Add(Me.lblCCNo)
        Me.GroupBox4.Controls.Add(Me.txtCUST_CREDIT_CARD_NO)
        Me.GroupBox4.Controls.Add(Me.UltraLabel4)
        Me.GroupBox4.Controls.Add(Me.txtLAST4)
        Me.GroupBox4.Location = New System.Drawing.Point(126, 22)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(547, 81)
        Me.GroupBox4.TabIndex = 213
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Credit Card Information"
        '
        'cboCUST_CREDIT_CARD_EXP_DATE
        '
        Me.cboCUST_CREDIT_CARD_EXP_DATE.FormattingEnabled = True
        Me.cboCUST_CREDIT_CARD_EXP_DATE.Location = New System.Drawing.Point(275, 37)
        Me.cboCUST_CREDIT_CARD_EXP_DATE.Name = "cboCUST_CREDIT_CARD_EXP_DATE"
        Me.cboCUST_CREDIT_CARD_EXP_DATE.Size = New System.Drawing.Size(102, 24)
        Me.cboCUST_CREDIT_CARD_EXP_DATE.TabIndex = 219
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(275, 18)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(73, 18)
        Me.UltraLabel2.TabIndex = 218
        Me.UltraLabel2.Text = "Expiration"
        '
        'txtCUST_CREDIT_CARD_VER_CODE
        '
        Me.Absx1.SetABSBindToTable(Me.txtCUST_CREDIT_CARD_VER_CODE, False)
        Me.Absx1.SetABSColumnName(Me.txtCUST_CREDIT_CARD_VER_CODE, "CUST_CREDIT_CARD_VER_CODE")
        Me.txtCUST_CREDIT_CARD_VER_CODE.Location = New System.Drawing.Point(481, 37)
        Me.txtCUST_CREDIT_CARD_VER_CODE.MaxLength = 10
        Me.txtCUST_CREDIT_CARD_VER_CODE.Name = "txtCUST_CREDIT_CARD_VER_CODE"
        Me.txtCUST_CREDIT_CARD_VER_CODE.Size = New System.Drawing.Size(52, 25)
        Me.txtCUST_CREDIT_CARD_VER_CODE.TabIndex = 216
        Me.txtCUST_CREDIT_CARD_VER_CODE.Tag = "Verification Code"
        '
        'UltraLabel5
        '
        Me.UltraLabel5.AutoSize = True
        Me.UltraLabel5.Location = New System.Drawing.Point(481, 18)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(41, 18)
        Me.UltraLabel5.TabIndex = 217
        Me.UltraLabel5.Text = "CVV2"
        '
        'txtShowNumbers
        '
        Me.txtShowNumbers.AutoSize = True
        Me.txtShowNumbers.Location = New System.Drawing.Point(8, 63)
        Me.txtShowNumbers.Name = "txtShowNumbers"
        Me.txtShowNumbers.Size = New System.Drawing.Size(124, 20)
        Me.txtShowNumbers.TabIndex = 4
        Me.txtShowNumbers.Text = "Show Numbers"
        Me.txtShowNumbers.UseVisualStyleBackColor = True
        '
        'lblCCNo
        '
        Me.lblCCNo.AutoSize = True
        Me.lblCCNo.Location = New System.Drawing.Point(8, 18)
        Me.lblCCNo.Name = "lblCCNo"
        Me.lblCCNo.Size = New System.Drawing.Size(104, 18)
        Me.lblCCNo.TabIndex = 111
        Me.lblCCNo.Text = "Credit Card No"
        '
        'txtCUST_CREDIT_CARD_NO
        '
        Me.Absx1.SetABSBindToTable(Me.txtCUST_CREDIT_CARD_NO, False)
        Me.Absx1.SetABSColumnName(Me.txtCUST_CREDIT_CARD_NO, "CUST_CREDIT_CARD_NO")
        Me.txtCUST_CREDIT_CARD_NO.Location = New System.Drawing.Point(8, 36)
        Me.txtCUST_CREDIT_CARD_NO.MaxLength = 25
        Me.txtCUST_CREDIT_CARD_NO.Name = "txtCUST_CREDIT_CARD_NO"
        Me.txtCUST_CREDIT_CARD_NO.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCUST_CREDIT_CARD_NO.Size = New System.Drawing.Size(246, 25)
        Me.txtCUST_CREDIT_CARD_NO.TabIndex = 1
        Me.txtCUST_CREDIT_CARD_NO.Tag = "Credit Card Number"
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(383, 18)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(89, 18)
        Me.UltraLabel4.TabIndex = 198
        Me.UltraLabel4.Text = "Last 4 Digits"
        '
        'txtLAST4
        '
        Me.Absx1.SetABSBindToTable(Me.txtLAST4, False)
        Me.txtLAST4.Location = New System.Drawing.Point(383, 37)
        Me.txtLAST4.Name = "txtLAST4"
        Me.txtLAST4.ReadOnly = True
        Me.txtLAST4.Size = New System.Drawing.Size(92, 25)
        Me.txtLAST4.TabIndex = 197
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtCUST_CREDIT_CARD_ADDR1)
        Me.GroupBox3.Controls.Add(Me.UltraLabel7)
        Me.GroupBox3.Controls.Add(Me.txtCUST_CREDIT_CARD_ZIP_CODE)
        Me.GroupBox3.Controls.Add(Me.UltraLabel10)
        Me.GroupBox3.Controls.Add(Me.txtCUST_CREDIT_CARD_STATE)
        Me.GroupBox3.Controls.Add(Me.UltraLabel9)
        Me.GroupBox3.Controls.Add(Me.txtCUST_CREDIT_CARD_CITY)
        Me.GroupBox3.Controls.Add(Me.UltraLabel8)
        Me.GroupBox3.Controls.Add(Me.txtCUST_CREDIT_CARD_NAME)
        Me.GroupBox3.Controls.Add(Me.UltraLabel6)
        Me.GroupBox3.Location = New System.Drawing.Point(18, 109)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(641, 128)
        Me.GroupBox3.TabIndex = 212
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Cardholder Name And Address"
        '
        'txtCUST_CREDIT_CARD_ADDR1
        '
        Me.Absx1.SetABSBindToTable(Me.txtCUST_CREDIT_CARD_ADDR1, False)
        Me.Absx1.SetABSColumnName(Me.txtCUST_CREDIT_CARD_ADDR1, "CUST_CREDIT_CARD_ADDR1")
        Me.txtCUST_CREDIT_CARD_ADDR1.Location = New System.Drawing.Point(125, 52)
        Me.txtCUST_CREDIT_CARD_ADDR1.MaxLength = 100
        Me.txtCUST_CREDIT_CARD_ADDR1.Name = "txtCUST_CREDIT_CARD_ADDR1"
        Me.txtCUST_CREDIT_CARD_ADDR1.Size = New System.Drawing.Size(416, 25)
        Me.txtCUST_CREDIT_CARD_ADDR1.TabIndex = 4
        Me.txtCUST_CREDIT_CARD_ADDR1.Tag = "Street Address"
        '
        'UltraLabel7
        '
        Me.UltraLabel7.AutoSize = True
        Me.UltraLabel7.Location = New System.Drawing.Point(15, 59)
        Me.UltraLabel7.Name = "UltraLabel7"
        Me.UltraLabel7.Size = New System.Drawing.Size(104, 18)
        Me.UltraLabel7.TabIndex = 217
        Me.UltraLabel7.Text = "Street Address"
        '
        'txtCUST_CREDIT_CARD_ZIP_CODE
        '
        Me.Absx1.SetABSBindToTable(Me.txtCUST_CREDIT_CARD_ZIP_CODE, False)
        Me.Absx1.SetABSColumnName(Me.txtCUST_CREDIT_CARD_ZIP_CODE, "CUST_CREDIT_CARD_ZIP_CODE")
        Me.txtCUST_CREDIT_CARD_ZIP_CODE.Location = New System.Drawing.Point(497, 87)
        Me.txtCUST_CREDIT_CARD_ZIP_CODE.MaxLength = 15
        Me.txtCUST_CREDIT_CARD_ZIP_CODE.Name = "txtCUST_CREDIT_CARD_ZIP_CODE"
        Me.txtCUST_CREDIT_CARD_ZIP_CODE.Size = New System.Drawing.Size(65, 25)
        Me.txtCUST_CREDIT_CARD_ZIP_CODE.TabIndex = 7
        Me.txtCUST_CREDIT_CARD_ZIP_CODE.Tag = "Zip Code"
        '
        'UltraLabel10
        '
        Me.UltraLabel10.AutoSize = True
        Me.UltraLabel10.Location = New System.Drawing.Point(469, 91)
        Me.UltraLabel10.Name = "UltraLabel10"
        Me.UltraLabel10.Size = New System.Drawing.Size(26, 18)
        Me.UltraLabel10.TabIndex = 215
        Me.UltraLabel10.Text = "Zip"
        '
        'txtCUST_CREDIT_CARD_STATE
        '
        Me.Absx1.SetABSBindToTable(Me.txtCUST_CREDIT_CARD_STATE, False)
        Me.Absx1.SetABSColumnName(Me.txtCUST_CREDIT_CARD_STATE, "CUST_CREDIT_CARD_STATE")
        Me.txtCUST_CREDIT_CARD_STATE.Location = New System.Drawing.Point(408, 87)
        Me.txtCUST_CREDIT_CARD_STATE.MaxLength = 2
        Me.txtCUST_CREDIT_CARD_STATE.Name = "txtCUST_CREDIT_CARD_STATE"
        Me.txtCUST_CREDIT_CARD_STATE.Size = New System.Drawing.Size(43, 25)
        Me.txtCUST_CREDIT_CARD_STATE.TabIndex = 6
        Me.txtCUST_CREDIT_CARD_STATE.Tag = "State"
        '
        'UltraLabel9
        '
        Me.UltraLabel9.AutoSize = True
        Me.UltraLabel9.Location = New System.Drawing.Point(368, 94)
        Me.UltraLabel9.Name = "UltraLabel9"
        Me.UltraLabel9.Size = New System.Drawing.Size(41, 18)
        Me.UltraLabel9.TabIndex = 213
        Me.UltraLabel9.Text = "State"
        '
        'txtCUST_CREDIT_CARD_CITY
        '
        Me.Absx1.SetABSBindToTable(Me.txtCUST_CREDIT_CARD_CITY, False)
        Me.Absx1.SetABSColumnName(Me.txtCUST_CREDIT_CARD_CITY, "CUST_CREDIT_CARD_CITY")
        Me.txtCUST_CREDIT_CARD_CITY.Location = New System.Drawing.Point(125, 87)
        Me.txtCUST_CREDIT_CARD_CITY.MaxLength = 100
        Me.txtCUST_CREDIT_CARD_CITY.Name = "txtCUST_CREDIT_CARD_CITY"
        Me.txtCUST_CREDIT_CARD_CITY.Size = New System.Drawing.Size(237, 25)
        Me.txtCUST_CREDIT_CARD_CITY.TabIndex = 5
        Me.txtCUST_CREDIT_CARD_CITY.Tag = "City"
        '
        'UltraLabel8
        '
        Me.UltraLabel8.AutoSize = True
        Me.UltraLabel8.Location = New System.Drawing.Point(88, 94)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(31, 18)
        Me.UltraLabel8.TabIndex = 209
        Me.UltraLabel8.Text = "City"
        '
        'txtCUST_CREDIT_CARD_NAME
        '
        Me.Absx1.SetABSBindToTable(Me.txtCUST_CREDIT_CARD_NAME, False)
        Me.Absx1.SetABSColumnName(Me.txtCUST_CREDIT_CARD_NAME, "CUST_CREDIT_CARD_NAME")
        Me.txtCUST_CREDIT_CARD_NAME.Location = New System.Drawing.Point(125, 21)
        Me.txtCUST_CREDIT_CARD_NAME.MaxLength = 100
        Me.txtCUST_CREDIT_CARD_NAME.Name = "txtCUST_CREDIT_CARD_NAME"
        Me.txtCUST_CREDIT_CARD_NAME.Size = New System.Drawing.Size(416, 25)
        Me.txtCUST_CREDIT_CARD_NAME.TabIndex = 3
        Me.txtCUST_CREDIT_CARD_NAME.Tag = "Name"
        '
        'UltraLabel6
        '
        Me.UltraLabel6.AutoSize = True
        Me.UltraLabel6.Location = New System.Drawing.Point(75, 28)
        Me.UltraLabel6.Name = "UltraLabel6"
        Me.UltraLabel6.Size = New System.Drawing.Size(44, 18)
        Me.UltraLabel6.TabIndex = 205
        Me.UltraLabel6.Text = "Name"
        '
        'optCC_TYPE
        '
        ValueListItem1.DataValue = "V"
        ValueListItem1.DisplayText = "Visa"
        ValueListItem2.DataValue = "M"
        ValueListItem2.DisplayText = "Mastercard"
        ValueListItem3.DataValue = "D"
        ValueListItem3.DisplayText = "Discover"
        ValueListItem4.DataValue = "A"
        ValueListItem4.DisplayText = "Amex"
        Me.optCC_TYPE.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem1, ValueListItem2, ValueListItem3, ValueListItem4})
        Me.optCC_TYPE.Location = New System.Drawing.Point(18, 25)
        Me.optCC_TYPE.Name = "optCC_TYPE"
        Me.optCC_TYPE.Size = New System.Drawing.Size(102, 78)
        Me.optCC_TYPE.TabIndex = 199
        Me.optCC_TYPE.Tag = "Credit Card Type"
        '
        'imgSTYLE
        '
        Me.imgSTYLE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.imgSTYLE.Location = New System.Drawing.Point(0, 0)
        Me.imgSTYLE.Name = "imgSTYLE"
        Me.imgSTYLE.Size = New System.Drawing.Size(710, 332)
        Me.imgSTYLE.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.imgSTYLE.TabIndex = 0
        Me.imgSTYLE.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCustAdd)
        Me.GroupBox1.Controls.Add(Me.txtCCPA_AMT)
        Me.GroupBox1.Controls.Add(Me.UltraLabel1)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.cmdFinished)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(710, 96)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'btnCustAdd
        '
        Me.btnCustAdd.Location = New System.Drawing.Point(164, 13)
        Me.btnCustAdd.Name = "btnCustAdd"
        Me.btnCustAdd.Size = New System.Drawing.Size(171, 33)
        Me.btnCustAdd.TabIndex = 218
        Me.btnCustAdd.Text = "Use Customer Address"
        '
        'txtCCPA_AMT
        '
        Me.Absx1.SetABSBindToTable(Me.txtCCPA_AMT, False)
        Appearance16.TextHAlignAsString = "Right"
        Me.txtCCPA_AMT.Appearance = Appearance16
        Me.txtCCPA_AMT.Location = New System.Drawing.Point(526, 13)
        Me.txtCCPA_AMT.Name = "txtCCPA_AMT"
        Me.txtCCPA_AMT.ReadOnly = True
        Me.txtCCPA_AMT.Size = New System.Drawing.Size(112, 25)
        Me.txtCCPA_AMT.TabIndex = 210
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(401, 15)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(119, 18)
        Me.UltraLabel1.TabIndex = 211
        Me.UltraLabel1.Text = "Credit Card Total"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(88, 13)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(70, 33)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel"
        '
        'cmdFinished
        '
        Me.cmdFinished.Location = New System.Drawing.Point(12, 13)
        Me.cmdFinished.Name = "cmdFinished"
        Me.cmdFinished.Size = New System.Drawing.Size(70, 33)
        Me.cmdFinished.TabIndex = 8
        Me.cmdFinished.Text = "Finished"
        '
        'lblAMEX
        '
        Me.lblAMEX.AutoSize = True
        Me.lblAMEX.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAMEX.ForeColor = System.Drawing.Color.Red
        Me.lblAMEX.Location = New System.Drawing.Point(17, 3)
        Me.lblAMEX.Name = "lblAMEX"
        Me.lblAMEX.Size = New System.Drawing.Size(618, 16)
        Me.lblAMEX.TabIndex = 215
        Me.lblAMEX.Text = "!!! Please Inform Your Customer That Amex Orders Will Incure A 3.5% Surcharge !!!" &
    ""
        Me.lblAMEX.Visible = False
        '
        'SOFORDRC
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(710, 432)
        Me.ControlBox = False
        Me.Margin = New System.Windows.Forms.Padding(6, 8, 6, 8)
        Me.Name = "SOFORDRC"
        Me.Text = "Credit Card Information"
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
        Me.GroupBox2.PerformLayout()
        Me.grpSOTORDRS.ResumeLayout(False)
        CType(Me.grdSOTORDRS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.txtCUST_CREDIT_CARD_VER_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_CREDIT_CARD_NO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLAST4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.txtCUST_CREDIT_CARD_ADDR1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_CREDIT_CARD_ZIP_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_CREDIT_CARD_STATE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_CREDIT_CARD_CITY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_CREDIT_CARD_NAME, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optCC_TYPE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgSTYLE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.txtCCPA_AMT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents imgSTYLE As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdFinished As Infragistics.Win.Misc.UltraButton
    Friend WithEvents btnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents txtCUST_CREDIT_CARD_NO As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblCCNo As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtLAST4 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents grpSOTORDRS As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtCUST_CREDIT_CARD_ADDR1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel7 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCUST_CREDIT_CARD_ZIP_CODE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel10 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCUST_CREDIT_CARD_STATE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel9 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCUST_CREDIT_CARD_CITY As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel8 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCUST_CREDIT_CARD_NAME As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel6 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents optCC_TYPE As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents grdSOTORDRS As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents txtCCPA_AMT As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents lblAuth As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtShowNumbers As System.Windows.Forms.CheckBox
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCUST_CREDIT_CARD_VER_CODE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboCUST_CREDIT_CARD_EXP_DATE As System.Windows.Forms.ComboBox
    Friend WithEvents btnCustAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents lblAMEX As Label
End Class
