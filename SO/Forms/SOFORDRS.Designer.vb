<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SOFORDRS
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
        Me.components = New System.ComponentModel.Container()
        Dim UltraToolTipInfo1 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Does the customer require a call ahead from the carrier to schedule the delivery?" &
        "", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo2 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Shipping out of the US-  Customer needs to provide us with a customs broker eithe" &
        "r to a port or Canada border", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo3 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Shipping out of the US-  Customer needs to provide us with a customs broker eithe" &
        "r to a port or Canada border.", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo4 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Closed for vacation, Special delivery hours and days, closed certain days etc…", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo5 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Closed certain hours/days/vacation, etc. Please list all.", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo6 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Churches, Schools, Hospitals, Municipal Buildings etc…", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo7 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Limited Access.", Infragistics.Win.ToolTipImage.[Default], "LA Title", Infragistics.Win.DefaultableBoolean.[True])
        Dim UltraToolTipInfo8 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Churches, Schools, Hospitals, Municipal Buildings etc.", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo9 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Do you have a loading dock or forklift?", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo10 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Please provide contact name and number", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo11 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Will the driver need to move a pallet beyond adjacent loading area?", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim UltraToolTipInfo12 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Will order be shipped to a home-based business or residential street?", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDR5", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VERIFIED")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR_CODE")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR1")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR2")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ADDR3")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CITY")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STATE")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_ZIP_CODE")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_COUNTRY")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_ORDR_NO")
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
        Dim UltraToolTipInfo13 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Click On Each Line To Verify More Than One Ship-To If Available.", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtAPPOINTMENT_REQUIRED_NOTE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtBROKER_NOTE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.chkBROKER = New System.Windows.Forms.CheckBox()
        Me.txtIRREGULAR_HOURS_NOTE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.chkIRREGULAR_HOURS = New System.Windows.Forms.CheckBox()
        Me.txtLIMITED_ACCESS_NOTE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.chkLIMITED_ACCESS = New System.Windows.Forms.CheckBox()
        Me.chkGATE_LIFT_REQ = New System.Windows.Forms.CheckBox()
        Me.chkAPPOINTMENT_REQUIRED_NOTE = New System.Windows.Forms.CheckBox()
        Me.chkINSIDE_REQ = New System.Windows.Forms.CheckBox()
        Me.txtLAST_OPER = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel13 = New Infragistics.Win.Misc.UltraLabel()
        Me.dteLAST_DATE = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.UltraLabel12 = New Infragistics.Win.Misc.UltraLabel()
        Me.lblSHIPINFO = New System.Windows.Forms.Label()
        Me.chkRESIDENTIAL_ORDR = New System.Windows.Forms.CheckBox()
        Me.btnDone = New System.Windows.Forms.Button()
        Me.lblAuth = New Infragistics.Win.Misc.UltraLabel()
        Me.grpSOTORDRS = New System.Windows.Forms.GroupBox()
        Me.grdARTCUSX2 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraToolTipManager1 = New Infragistics.Win.UltraWinToolTip.UltraToolTipManager(Me.components)
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.txtAPPOINTMENT_REQUIRED_NOTE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtBROKER_NOTE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtIRREGULAR_HOURS_NOTE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLIMITED_ACCESS_NOTE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLAST_OPER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteLAST_DATE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpSOTORDRS.SuspendLayout()
        CType(Me.grdARTCUSX2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.Panel1)
        Me.ASFBASE2_Fill_Panel.Margin = New System.Windows.Forms.Padding(6)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(903, 406)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 406)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(903, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 406)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(903, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 406)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(6)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(903, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(903, 406)
        Me.Panel1.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Panel2)
        Me.GroupBox2.Controls.Add(Me.btnDone)
        Me.GroupBox2.Controls.Add(Me.lblAuth)
        Me.GroupBox2.Controls.Add(Me.grpSOTORDRS)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(903, 406)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(122, 380)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(535, 13)
        Me.Label1.TabIndex = 218
        Me.Label1.Text = "Note: Hover Your Mouse Over Any Item Above To See More Information About It's Mea" &
    "ning."
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtAPPOINTMENT_REQUIRED_NOTE)
        Me.Panel2.Controls.Add(Me.txtBROKER_NOTE)
        Me.Panel2.Controls.Add(Me.chkBROKER)
        Me.Panel2.Controls.Add(Me.txtIRREGULAR_HOURS_NOTE)
        Me.Panel2.Controls.Add(Me.chkIRREGULAR_HOURS)
        Me.Panel2.Controls.Add(Me.txtLIMITED_ACCESS_NOTE)
        Me.Panel2.Controls.Add(Me.chkLIMITED_ACCESS)
        Me.Panel2.Controls.Add(Me.chkGATE_LIFT_REQ)
        Me.Panel2.Controls.Add(Me.chkAPPOINTMENT_REQUIRED_NOTE)
        Me.Panel2.Controls.Add(Me.chkINSIDE_REQ)
        Me.Panel2.Controls.Add(Me.txtLAST_OPER)
        Me.Panel2.Controls.Add(Me.UltraLabel13)
        Me.Panel2.Controls.Add(Me.dteLAST_DATE)
        Me.Panel2.Controls.Add(Me.UltraLabel12)
        Me.Panel2.Controls.Add(Me.lblSHIPINFO)
        Me.Panel2.Controls.Add(Me.chkRESIDENTIAL_ORDR)
        Me.Panel2.Location = New System.Drawing.Point(12, 195)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(882, 169)
        Me.Panel2.TabIndex = 217
        '
        'txtAPPOINTMENT_REQUIRED_NOTE
        '
        Me.Absx1.SetABSBindToTable(Me.txtAPPOINTMENT_REQUIRED_NOTE, False)
        Me.Absx1.SetABSHasButton(Me.txtAPPOINTMENT_REQUIRED_NOTE, True)
        Me.txtAPPOINTMENT_REQUIRED_NOTE.Location = New System.Drawing.Point(442, 139)
        Me.txtAPPOINTMENT_REQUIRED_NOTE.Name = "txtAPPOINTMENT_REQUIRED_NOTE"
        Me.txtAPPOINTMENT_REQUIRED_NOTE.Size = New System.Drawing.Size(416, 25)
        Me.txtAPPOINTMENT_REQUIRED_NOTE.TabIndex = 140
        UltraToolTipInfo1.ToolTipText = "Does the customer require a call ahead from the carrier to schedule the delivery?" &
    ""
        Me.UltraToolTipManager1.SetUltraToolTip(Me.txtAPPOINTMENT_REQUIRED_NOTE, UltraToolTipInfo1)
        Me.txtAPPOINTMENT_REQUIRED_NOTE.Visible = False
        '
        'txtBROKER_NOTE
        '
        Me.Absx1.SetABSBindToTable(Me.txtBROKER_NOTE, False)
        Me.Absx1.SetABSHasButton(Me.txtBROKER_NOTE, True)
        Me.txtBROKER_NOTE.Location = New System.Drawing.Point(442, 113)
        Me.txtBROKER_NOTE.Name = "txtBROKER_NOTE"
        Me.txtBROKER_NOTE.Size = New System.Drawing.Size(416, 25)
        Me.txtBROKER_NOTE.TabIndex = 139
        UltraToolTipInfo2.ToolTipText = "Shipping out of the US-  Customer needs to provide us with a customs broker eithe" &
    "r to a port or Canada border"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.txtBROKER_NOTE, UltraToolTipInfo2)
        Me.txtBROKER_NOTE.Visible = False
        '
        'chkBROKER
        '
        Me.chkBROKER.AutoSize = True
        Me.chkBROKER.Location = New System.Drawing.Point(310, 116)
        Me.chkBROKER.Name = "chkBROKER"
        Me.chkBROKER.Size = New System.Drawing.Size(68, 20)
        Me.chkBROKER.TabIndex = 138
        Me.chkBROKER.Text = "Broker"
        UltraToolTipInfo3.ToolTipText = "Shipping out of the US-  Customer needs to provide us with a customs broker eithe" &
    "r to a port or Canada border."
        Me.UltraToolTipManager1.SetUltraToolTip(Me.chkBROKER, UltraToolTipInfo3)
        Me.chkBROKER.UseVisualStyleBackColor = True
        '
        'txtIRREGULAR_HOURS_NOTE
        '
        Me.Absx1.SetABSBindToTable(Me.txtIRREGULAR_HOURS_NOTE, False)
        Me.Absx1.SetABSHasButton(Me.txtIRREGULAR_HOURS_NOTE, True)
        Me.txtIRREGULAR_HOURS_NOTE.Location = New System.Drawing.Point(442, 86)
        Me.txtIRREGULAR_HOURS_NOTE.Name = "txtIRREGULAR_HOURS_NOTE"
        Me.txtIRREGULAR_HOURS_NOTE.Size = New System.Drawing.Size(416, 25)
        Me.txtIRREGULAR_HOURS_NOTE.TabIndex = 137
        UltraToolTipInfo4.ToolTipText = "Closed for vacation, Special delivery hours and days, closed certain days etc…"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.txtIRREGULAR_HOURS_NOTE, UltraToolTipInfo4)
        Me.txtIRREGULAR_HOURS_NOTE.Visible = False
        '
        'chkIRREGULAR_HOURS
        '
        Me.chkIRREGULAR_HOURS.AutoSize = True
        Me.chkIRREGULAR_HOURS.Location = New System.Drawing.Point(310, 89)
        Me.chkIRREGULAR_HOURS.Name = "chkIRREGULAR_HOURS"
        Me.chkIRREGULAR_HOURS.Size = New System.Drawing.Size(124, 20)
        Me.chkIRREGULAR_HOURS.TabIndex = 136
        Me.chkIRREGULAR_HOURS.Text = "Irregular Hours"
        UltraToolTipInfo5.ToolTipText = "Closed certain hours/days/vacation, etc. Please list all."
        Me.UltraToolTipManager1.SetUltraToolTip(Me.chkIRREGULAR_HOURS, UltraToolTipInfo5)
        Me.chkIRREGULAR_HOURS.UseVisualStyleBackColor = True
        '
        'txtLIMITED_ACCESS_NOTE
        '
        Me.Absx1.SetABSBindToTable(Me.txtLIMITED_ACCESS_NOTE, False)
        Me.Absx1.SetABSHasButton(Me.txtLIMITED_ACCESS_NOTE, True)
        Me.txtLIMITED_ACCESS_NOTE.Location = New System.Drawing.Point(442, 60)
        Me.txtLIMITED_ACCESS_NOTE.Name = "txtLIMITED_ACCESS_NOTE"
        Me.txtLIMITED_ACCESS_NOTE.Size = New System.Drawing.Size(416, 25)
        Me.txtLIMITED_ACCESS_NOTE.TabIndex = 135
        UltraToolTipInfo6.ToolTipText = "Churches, Schools, Hospitals, Municipal Buildings etc…"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.txtLIMITED_ACCESS_NOTE, UltraToolTipInfo6)
        Me.txtLIMITED_ACCESS_NOTE.Visible = False
        '
        'chkLIMITED_ACCESS
        '
        Me.chkLIMITED_ACCESS.AutoSize = True
        Me.chkLIMITED_ACCESS.Location = New System.Drawing.Point(310, 63)
        Me.chkLIMITED_ACCESS.Name = "chkLIMITED_ACCESS"
        Me.chkLIMITED_ACCESS.Size = New System.Drawing.Size(125, 20)
        Me.chkLIMITED_ACCESS.TabIndex = 134
        Me.chkLIMITED_ACCESS.Text = "Limited Access"
        UltraToolTipInfo7.Enabled = Infragistics.Win.DefaultableBoolean.[True]
        UltraToolTipInfo7.ToolTipText = "Limited Access."
        UltraToolTipInfo7.ToolTipTextFormatted = "Limited Access Fomatted"
        UltraToolTipInfo7.ToolTipTitle = "LA Title"
        Me.tip.SetUltraToolTip(Me.chkLIMITED_ACCESS, UltraToolTipInfo7)
        UltraToolTipInfo8.ToolTipText = "Churches, Schools, Hospitals, Municipal Buildings etc."
        Me.UltraToolTipManager1.SetUltraToolTip(Me.chkLIMITED_ACCESS, UltraToolTipInfo8)
        Me.chkLIMITED_ACCESS.UseVisualStyleBackColor = True
        '
        'chkGATE_LIFT_REQ
        '
        Me.chkGATE_LIFT_REQ.AutoSize = True
        Me.chkGATE_LIFT_REQ.Location = New System.Drawing.Point(12, 116)
        Me.chkGATE_LIFT_REQ.Name = "chkGATE_LIFT_REQ"
        Me.chkGATE_LIFT_REQ.Size = New System.Drawing.Size(145, 20)
        Me.chkGATE_LIFT_REQ.TabIndex = 133
        Me.chkGATE_LIFT_REQ.Text = "Lift Gate Required"
        UltraToolTipInfo9.ToolTipText = "Do you have a loading dock or forklift?"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.chkGATE_LIFT_REQ, UltraToolTipInfo9)
        Me.chkGATE_LIFT_REQ.UseVisualStyleBackColor = True
        '
        'chkAPPOINTMENT_REQUIRED_NOTE
        '
        Me.chkAPPOINTMENT_REQUIRED_NOTE.AutoSize = True
        Me.chkAPPOINTMENT_REQUIRED_NOTE.Location = New System.Drawing.Point(310, 142)
        Me.chkAPPOINTMENT_REQUIRED_NOTE.Name = "chkAPPOINTMENT_REQUIRED_NOTE"
        Me.chkAPPOINTMENT_REQUIRED_NOTE.Size = New System.Drawing.Size(119, 20)
        Me.chkAPPOINTMENT_REQUIRED_NOTE.TabIndex = 132
        Me.chkAPPOINTMENT_REQUIRED_NOTE.Text = "Appt Required"
        UltraToolTipInfo10.ToolTipText = "Please provide contact name and number"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.chkAPPOINTMENT_REQUIRED_NOTE, UltraToolTipInfo10)
        Me.chkAPPOINTMENT_REQUIRED_NOTE.UseVisualStyleBackColor = True
        '
        'chkINSIDE_REQ
        '
        Me.chkINSIDE_REQ.AutoSize = True
        Me.chkINSIDE_REQ.Location = New System.Drawing.Point(12, 89)
        Me.chkINSIDE_REQ.Name = "chkINSIDE_REQ"
        Me.chkINSIDE_REQ.Size = New System.Drawing.Size(184, 20)
        Me.chkINSIDE_REQ.TabIndex = 131
        Me.chkINSIDE_REQ.Text = "Inside Delivery Required"
        UltraToolTipInfo11.ToolTipText = "Will the driver need to move a pallet beyond adjacent loading area?"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.chkINSIDE_REQ, UltraToolTipInfo11)
        Me.chkINSIDE_REQ.UseVisualStyleBackColor = True
        '
        'txtLAST_OPER
        '
        Me.Absx1.SetABSBindToTable(Me.txtLAST_OPER, False)
        Me.Absx1.SetABSHasButton(Me.txtLAST_OPER, True)
        Me.txtLAST_OPER.Location = New System.Drawing.Point(289, 20)
        Me.txtLAST_OPER.Name = "txtLAST_OPER"
        Me.txtLAST_OPER.ReadOnly = True
        Me.txtLAST_OPER.Size = New System.Drawing.Size(99, 25)
        Me.txtLAST_OPER.TabIndex = 129
        '
        'UltraLabel13
        '
        Me.UltraLabel13.AutoSize = True
        Me.UltraLabel13.Location = New System.Drawing.Point(292, 3)
        Me.UltraLabel13.Name = "UltraLabel13"
        Me.UltraLabel13.Size = New System.Drawing.Size(78, 18)
        Me.UltraLabel13.TabIndex = 130
        Me.UltraLabel13.Text = "Verified By"
        '
        'dteLAST_DATE
        '
        Me.dteLAST_DATE.DropDownButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Never
        Me.dteLAST_DATE.Location = New System.Drawing.Point(173, 20)
        Me.dteLAST_DATE.Name = "dteLAST_DATE"
        Me.dteLAST_DATE.ReadOnly = True
        Me.dteLAST_DATE.Size = New System.Drawing.Size(109, 25)
        Me.dteLAST_DATE.TabIndex = 127
        '
        'UltraLabel12
        '
        Me.UltraLabel12.AutoSize = True
        Me.UltraLabel12.Location = New System.Drawing.Point(173, 2)
        Me.UltraLabel12.Name = "UltraLabel12"
        Me.UltraLabel12.Size = New System.Drawing.Size(89, 18)
        Me.UltraLabel12.TabIndex = 128
        Me.UltraLabel12.Text = "Last Verified"
        '
        'lblSHIPINFO
        '
        Me.lblSHIPINFO.AutoSize = True
        Me.lblSHIPINFO.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSHIPINFO.Location = New System.Drawing.Point(9, 3)
        Me.lblSHIPINFO.Name = "lblSHIPINFO"
        Me.lblSHIPINFO.Size = New System.Drawing.Size(153, 16)
        Me.lblSHIPINFO.TabIndex = 126
        Me.lblSHIPINFO.Text = "Ship-To Information"
        '
        'chkRESIDENTIAL_ORDR
        '
        Me.chkRESIDENTIAL_ORDR.AutoSize = True
        Me.chkRESIDENTIAL_ORDR.Location = New System.Drawing.Point(12, 63)
        Me.chkRESIDENTIAL_ORDR.Name = "chkRESIDENTIAL_ORDR"
        Me.chkRESIDENTIAL_ORDR.Size = New System.Drawing.Size(154, 20)
        Me.chkRESIDENTIAL_ORDR.TabIndex = 125
        Me.chkRESIDENTIAL_ORDR.Text = "Residential Delivery"
        UltraToolTipInfo12.ToolTipText = "Will order be shipped to a home-based business or residential street?"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.chkRESIDENTIAL_ORDR, UltraToolTipInfo12)
        Me.chkRESIDENTIAL_ORDR.UseVisualStyleBackColor = True
        '
        'btnDone
        '
        Me.btnDone.Location = New System.Drawing.Point(6, 370)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.Size = New System.Drawing.Size(75, 23)
        Me.btnDone.TabIndex = 215
        Me.btnDone.Text = "Done"
        Me.btnDone.UseVisualStyleBackColor = True
        '
        'lblAuth
        '
        Appearance1.ForeColor = System.Drawing.Color.Red
        Me.lblAuth.Appearance = Appearance1
        Me.lblAuth.AutoSize = True
        Me.lblAuth.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAuth.Location = New System.Drawing.Point(206, 3)
        Me.lblAuth.Name = "lblAuth"
        Me.lblAuth.Size = New System.Drawing.Size(0, 0)
        Me.lblAuth.TabIndex = 212
        '
        'grpSOTORDRS
        '
        Me.grpSOTORDRS.Controls.Add(Me.grdARTCUSX2)
        Me.grpSOTORDRS.Location = New System.Drawing.Point(12, 9)
        Me.grpSOTORDRS.Name = "grpSOTORDRS"
        Me.grpSOTORDRS.Size = New System.Drawing.Size(882, 183)
        Me.grpSOTORDRS.TabIndex = 214
        Me.grpSOTORDRS.TabStop = False
        Me.grpSOTORDRS.Text = "Ship-Tos For Order"
        '
        'grdARTCUSX2
        '
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdARTCUSX2.DisplayLayout.Appearance = Appearance2
        UltraGridColumn1.Header.Caption = "Verified"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Width = 83
        UltraGridColumn3.Header.Caption = "Code"
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Width = 80
        UltraGridColumn4.Header.Caption = "Name"
        UltraGridColumn4.Header.VisiblePosition = 2
        UltraGridColumn4.Width = 184
        UltraGridColumn7.Header.Caption = "Addr1"
        UltraGridColumn7.Header.VisiblePosition = 3
        UltraGridColumn7.Width = 228
        UltraGridColumn8.Header.Caption = "Addr2"
        UltraGridColumn8.Header.VisiblePosition = 4
        UltraGridColumn9.Header.Caption = "Addr3"
        UltraGridColumn9.Header.VisiblePosition = 5
        UltraGridColumn10.Header.Caption = "City"
        UltraGridColumn10.Header.VisiblePosition = 6
        UltraGridColumn11.Header.Caption = "State"
        UltraGridColumn11.Header.VisiblePosition = 7
        UltraGridColumn11.Width = 65
        UltraGridColumn12.Header.Caption = "Zip"
        UltraGridColumn12.Header.VisiblePosition = 8
        UltraGridColumn12.Width = 71
        UltraGridColumn13.Header.Caption = "Country"
        UltraGridColumn13.Header.VisiblePosition = 9
        UltraGridColumn13.Width = 70
        UltraGridColumn2.Header.VisiblePosition = 10
        UltraGridColumn2.Hidden = True
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn3, UltraGridColumn4, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9, UltraGridColumn10, UltraGridColumn11, UltraGridColumn12, UltraGridColumn13, UltraGridColumn2})
        Me.grdARTCUSX2.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdARTCUSX2.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance3.TextHAlignAsString = "Left"
        Me.grdARTCUSX2.DisplayLayout.CaptionAppearance = Appearance3
        Appearance4.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance4.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance4.BorderColor = System.Drawing.SystemColors.Window
        Me.grdARTCUSX2.DisplayLayout.GroupByBox.Appearance = Appearance4
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdARTCUSX2.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance5
        Me.grdARTCUSX2.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdARTCUSX2.DisplayLayout.GroupByBox.Hidden = True
        Appearance6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance6.BackColor2 = System.Drawing.SystemColors.Control
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance6.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdARTCUSX2.DisplayLayout.GroupByBox.PromptAppearance = Appearance6
        Me.grdARTCUSX2.DisplayLayout.MaxColScrollRegions = 1
        Me.grdARTCUSX2.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdARTCUSX2.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdARTCUSX2.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Appearance7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdARTCUSX2.DisplayLayout.Override.ActiveCellAppearance = Appearance7
        Me.grdARTCUSX2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        Me.grdARTCUSX2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdARTCUSX2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdARTCUSX2.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdARTCUSX2.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance8.BackColor = System.Drawing.SystemColors.Window
        Me.grdARTCUSX2.DisplayLayout.Override.CardAreaAppearance = Appearance8
        Appearance9.BorderColor = System.Drawing.Color.Silver
        Appearance9.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdARTCUSX2.DisplayLayout.Override.CellAppearance = Appearance9
        Me.grdARTCUSX2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdARTCUSX2.DisplayLayout.Override.CellPadding = 0
        Appearance10.BackColor = System.Drawing.SystemColors.Control
        Appearance10.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance10.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance10.BorderColor = System.Drawing.SystemColors.Window
        Me.grdARTCUSX2.DisplayLayout.Override.GroupByRowAppearance = Appearance10
        Appearance11.TextHAlignAsString = "Left"
        Me.grdARTCUSX2.DisplayLayout.Override.HeaderAppearance = Appearance11
        Me.grdARTCUSX2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdARTCUSX2.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance12.BackColor = System.Drawing.SystemColors.Window
        Appearance12.BorderColor = System.Drawing.Color.Silver
        Me.grdARTCUSX2.DisplayLayout.Override.RowAppearance = Appearance12
        Me.grdARTCUSX2.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.grdARTCUSX2.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
        Appearance13.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdARTCUSX2.DisplayLayout.Override.TemplateAddRowAppearance = Appearance13
        Me.grdARTCUSX2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdARTCUSX2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdARTCUSX2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdARTCUSX2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdARTCUSX2.Location = New System.Drawing.Point(3, 19)
        Me.grdARTCUSX2.Name = "grdARTCUSX2"
        Me.grdARTCUSX2.Size = New System.Drawing.Size(876, 161)
        Me.grdARTCUSX2.TabIndex = 25
        UltraToolTipInfo13.ToolTipText = "Click On Each Line To Verify More Than One Ship-To If Available."
        Me.UltraToolTipManager1.SetUltraToolTip(Me.grdARTCUSX2, UltraToolTipInfo13)
        '
        'UltraToolTipManager1
        '
        Me.UltraToolTipManager1.ContainingControl = Me
        '
        'SOFORDRS
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(903, 406)
        Me.ControlBox = False
        Me.Margin = New System.Windows.Forms.Padding(6, 8, 6, 8)
        Me.Name = "SOFORDRS"
        Me.Text = "Ship-To Verification"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.txtAPPOINTMENT_REQUIRED_NOTE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtBROKER_NOTE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtIRREGULAR_HOURS_NOTE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLIMITED_ACCESS_NOTE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLAST_OPER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteLAST_DATE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpSOTORDRS.ResumeLayout(False)
        CType(Me.grdARTCUSX2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btnDone As Button
    Friend WithEvents lblAuth As Misc.UltraLabel
    Friend WithEvents grpSOTORDRS As GroupBox
    Friend WithEvents grdARTCUSX2 As UltraWinGrid.UltraGrid
    Friend WithEvents Panel2 As Panel
    Friend WithEvents txtIRREGULAR_HOURS_NOTE As UltraWinEditors.UltraTextEditor
    Friend WithEvents chkIRREGULAR_HOURS As CheckBox
    Friend WithEvents txtLIMITED_ACCESS_NOTE As UltraWinEditors.UltraTextEditor
    Friend WithEvents chkLIMITED_ACCESS As CheckBox
    Friend WithEvents chkGATE_LIFT_REQ As CheckBox
    Friend WithEvents chkAPPOINTMENT_REQUIRED_NOTE As CheckBox
    Friend WithEvents chkINSIDE_REQ As CheckBox
    Friend WithEvents txtLAST_OPER As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel13 As Misc.UltraLabel
    Friend WithEvents dteLAST_DATE As UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents UltraLabel12 As Misc.UltraLabel
    Friend WithEvents lblSHIPINFO As Label
    Friend WithEvents chkRESIDENTIAL_ORDR As CheckBox
    Friend WithEvents txtBROKER_NOTE As UltraWinEditors.UltraTextEditor
    Friend WithEvents chkBROKER As CheckBox
    Friend WithEvents txtAPPOINTMENT_REQUIRED_NOTE As UltraWinEditors.UltraTextEditor
    Friend WithEvents Label1 As Label
    Friend WithEvents UltraToolTipManager1 As UltraWinToolTip.UltraToolTipManager
End Class
