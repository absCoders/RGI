<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SOFPICKU
    Inherits ASFBASE2

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
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDQ1", -1)
        Dim UltraGridColumn30 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GROUP_KEY")
        Dim UltraGridColumn31 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NPIX_NO")
        Dim UltraGridColumn32 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn33 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CNT")
        Dim UltraGridColumn34 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_OPEN")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_BACK")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_ALLO")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VOL_INDEX_TOT")
        Dim UltraGridColumn35 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_AVA")
        Dim UltraGridColumn36 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_TYPE")
        Dim UltraGridColumn37 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_TYPE_CODE")
        Dim UltraGridColumn38 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_DELIVERY_DATE")
        Dim UltraGridColumn39 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PARTNER_CODE")
        Dim UltraGridColumn40 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAB_CODE")
        Dim UltraGridColumn41 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SO_NOTES")
        Dim UltraGridColumn42 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn43 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn44 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_SHIP_TO_NO")
        Dim UltraGridColumn45 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NO")
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_SHIP_TO_NAME")
        Dim UltraGridColumn47 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO")
        Dim UltraGridColumn48 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SOURCE")
        Dim UltraGridColumn49 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DESTINATION")
        Dim UltraGridColumn50 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("KITTING")
        Dim UltraGridColumn51 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("JOB_NO")
        Dim UltraGridColumn52 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_CODE")
        Dim UltraGridColumn53 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_UPC_CODE")
        Dim UltraGridColumn54 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FRAME_STYLE_DESC")
        Dim UltraGridColumn55 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FRAME_SOURCE_CTL_NO")
        Dim UltraGridColumn56 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_XFR_BATCH_NO")
        Dim UltraGridColumn57 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SEL")
        Dim UltraGridColumn58 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTE_NO")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SLOT_NO")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTE_CLASS_CODE")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTE_CLASS_MIN_QTY")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTE_CLASS_MAX_QTY")
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
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTSCAN1", -1)
        Dim UltraGridColumn59 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SCAN_NO")
        Dim UltraGridColumn60 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SCAN")
        Dim UltraGridColumn61 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RESULT", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ERR")
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
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.grdSOTORDQ1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdSOTSCAN1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        Me.lblInstruction1 = New Infragistics.Win.Misc.UltraLabel()
        Me.chkCustomTruck = New ABSCS.ABSCheckBox()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton()
        Me.lblInstruction2 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtSCAN = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblPreConfigured = New Infragistics.Win.Misc.UltraLabel()
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton()
        Me.lblTruckIsReady = New Infragistics.Win.Misc.UltraLabel()
        Me.btnSimulateScan = New Infragistics.Win.Misc.UltraButton()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.lblInstruction = New Infragistics.Win.Misc.UltraLabel()
        Me.txtTRUCK = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
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
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.grdSOTORDQ1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTSCAN1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkCustomTruck, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSCAN, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTRUCK, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(1435, 456)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 456)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(1435, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 456)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(1435, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 456)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(1435, 0)
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
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.UltraLabel3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblInstruction1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkCustomTruck)
        Me.SplitContainer1.Panel2.Controls.Add(Me.UltraLabel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblInstruction2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtSCAN)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblPreConfigured)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdUpdate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblTruckIsReady)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSimulateScan)
        Me.SplitContainer1.Panel2.Controls.Add(Me.UltraLabel2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblInstruction)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtTRUCK)
        Me.SplitContainer1.Size = New System.Drawing.Size(1435, 456)
        Me.SplitContainer1.SplitterDistance = 319
        Me.SplitContainer1.TabIndex = 0
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.grdSOTORDQ1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.grdSOTSCAN1)
        Me.SplitContainer2.Size = New System.Drawing.Size(1435, 319)
        Me.SplitContainer2.SplitterDistance = 818
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 0
        '
        'grdSOTORDQ1
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTORDQ1.DisplayLayout.Appearance = Appearance1
        UltraGridColumn30.Header.VisiblePosition = 0
        UltraGridColumn30.Hidden = True
        UltraGridColumn31.Header.Caption = "NPI No"
        UltraGridColumn31.Header.VisiblePosition = 2
        UltraGridColumn31.Hidden = True
        UltraGridColumn31.Width = 69
        UltraGridColumn32.Header.Caption = "Whse"
        UltraGridColumn32.Header.VisiblePosition = 3
        UltraGridColumn32.Hidden = True
        UltraGridColumn32.Width = 64
        UltraGridColumn33.Header.Caption = "#Orders"
        UltraGridColumn33.Header.VisiblePosition = 4
        UltraGridColumn33.Hidden = True
        UltraGridColumn33.Width = 56
        UltraGridColumn34.Header.Caption = "Open"
        UltraGridColumn34.Header.VisiblePosition = 6
        UltraGridColumn34.Width = 50
        UltraGridColumn6.Header.Caption = "Back"
        UltraGridColumn6.Header.VisiblePosition = 7
        UltraGridColumn6.Width = 50
        UltraGridColumn7.Header.Caption = "Allo"
        UltraGridColumn7.Header.VisiblePosition = 8
        UltraGridColumn7.Width = 50
        UltraGridColumn8.Format = "#,##0"
        UltraGridColumn8.Header.Caption = "VolTot"
        UltraGridColumn8.Header.VisiblePosition = 35
        UltraGridColumn8.Width = 70
        UltraGridColumn35.Header.Caption = "Qty Ava"
        UltraGridColumn35.Header.VisiblePosition = 9
        UltraGridColumn35.Hidden = True
        UltraGridColumn35.Width = 69
        UltraGridColumn36.Header.Caption = "PO Type"
        UltraGridColumn36.Header.VisiblePosition = 10
        UltraGridColumn36.Hidden = True
        UltraGridColumn36.Width = 132
        UltraGridColumn37.Header.Caption = "Type"
        UltraGridColumn37.Header.VisiblePosition = 11
        UltraGridColumn37.Hidden = True
        UltraGridColumn37.Width = 51
        UltraGridColumn38.Header.Caption = "Deliver By"
        UltraGridColumn38.Header.VisiblePosition = 12
        UltraGridColumn38.Width = 98
        UltraGridColumn39.Header.Caption = "Partner"
        UltraGridColumn39.Header.VisiblePosition = 13
        UltraGridColumn39.Hidden = True
        UltraGridColumn39.Width = 67
        UltraGridColumn40.Header.Caption = "Lab"
        UltraGridColumn40.Header.VisiblePosition = 14
        UltraGridColumn40.Hidden = True
        UltraGridColumn40.Width = 52
        UltraGridColumn41.Header.Caption = "Order Notes"
        UltraGridColumn41.Header.VisiblePosition = 18
        UltraGridColumn41.Width = 108
        UltraGridColumn42.Header.Caption = "Order No"
        UltraGridColumn42.Header.VisiblePosition = 5
        UltraGridColumn42.Width = 100
        UltraGridColumn43.Header.Caption = "Customer"
        UltraGridColumn43.Header.VisiblePosition = 19
        UltraGridColumn43.Width = 87
        UltraGridColumn44.Header.Caption = "ShipTo"
        UltraGridColumn44.Header.VisiblePosition = 20
        UltraGridColumn44.Width = 52
        UltraGridColumn45.Header.Caption = "Store"
        UltraGridColumn45.Header.VisiblePosition = 22
        UltraGridColumn45.Width = 67
        UltraGridColumn46.Header.Caption = "Ship-To Name"
        UltraGridColumn46.Header.VisiblePosition = 21
        UltraGridColumn46.Hidden = True
        UltraGridColumn47.Header.Caption = "Customer PO"
        UltraGridColumn47.Header.VisiblePosition = 23
        UltraGridColumn47.Hidden = True
        UltraGridColumn47.Width = 115
        UltraGridColumn48.Header.Caption = "Source"
        UltraGridColumn48.Header.VisiblePosition = 16
        UltraGridColumn48.Hidden = True
        UltraGridColumn48.Width = 86
        UltraGridColumn49.Header.Caption = "Destination"
        UltraGridColumn49.Header.VisiblePosition = 15
        UltraGridColumn49.Hidden = True
        UltraGridColumn49.Width = 85
        UltraGridColumn50.Header.Caption = "Kitting"
        UltraGridColumn50.Header.VisiblePosition = 17
        UltraGridColumn50.Hidden = True
        UltraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn50.Width = 57
        UltraGridColumn51.Header.Caption = "Job No"
        UltraGridColumn51.Header.VisiblePosition = 24
        UltraGridColumn51.Hidden = True
        UltraGridColumn51.Width = 72
        UltraGridColumn52.Header.Caption = "Item Code"
        UltraGridColumn52.Header.VisiblePosition = 25
        UltraGridColumn52.Hidden = True
        UltraGridColumn53.Header.Caption = "UPC Code"
        UltraGridColumn53.Header.VisiblePosition = 26
        UltraGridColumn53.Hidden = True
        UltraGridColumn53.Width = 97
        UltraGridColumn54.Header.Caption = "Description"
        UltraGridColumn54.Header.VisiblePosition = 27
        UltraGridColumn54.Hidden = True
        UltraGridColumn55.Header.Caption = "Source Ctl"
        UltraGridColumn55.Header.VisiblePosition = 28
        UltraGridColumn55.Hidden = True
        UltraGridColumn55.Width = 110
        UltraGridColumn56.Header.Caption = "Xfr Batch No"
        UltraGridColumn56.Header.VisiblePosition = 29
        UltraGridColumn56.Hidden = True
        UltraGridColumn56.Width = 104
        UltraGridColumn57.Header.Caption = "Sel"
        UltraGridColumn57.Header.VisiblePosition = 1
        UltraGridColumn57.Hidden = True
        UltraGridColumn57.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn57.Width = 40
        UltraGridColumn58.Header.Caption = "Tote No"
        UltraGridColumn58.Header.VisiblePosition = 30
        UltraGridColumn58.Width = 65
        UltraGridColumn2.Header.Caption = "Slot#"
        UltraGridColumn2.Header.VisiblePosition = 31
        UltraGridColumn2.Width = 56
        UltraGridColumn3.Header.Caption = "Cls"
        UltraGridColumn3.Header.VisiblePosition = 32
        UltraGridColumn3.Width = 50
        UltraGridColumn4.Header.Caption = "MinVol"
        UltraGridColumn4.Header.VisiblePosition = 33
        UltraGridColumn4.Width = 60
        UltraGridColumn5.Header.Caption = "MaxVol"
        UltraGridColumn5.Header.VisiblePosition = 34
        UltraGridColumn5.Width = 70
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn30, UltraGridColumn31, UltraGridColumn32, UltraGridColumn33, UltraGridColumn34, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn35, UltraGridColumn36, UltraGridColumn37, UltraGridColumn38, UltraGridColumn39, UltraGridColumn40, UltraGridColumn41, UltraGridColumn42, UltraGridColumn43, UltraGridColumn44, UltraGridColumn45, UltraGridColumn46, UltraGridColumn47, UltraGridColumn48, UltraGridColumn49, UltraGridColumn50, UltraGridColumn51, UltraGridColumn52, UltraGridColumn53, UltraGridColumn54, UltraGridColumn55, UltraGridColumn56, UltraGridColumn57, UltraGridColumn58, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5})
        Me.grdSOTORDQ1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTORDQ1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdSOTORDQ1.DisplayLayout.CaptionAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDQ1.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDQ1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdSOTORDQ1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTORDQ1.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDQ1.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdSOTORDQ1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTORDQ1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTORDQ1.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdSOTORDQ1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTORDQ1.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdSOTORDQ1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTORDQ1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDQ1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTORDQ1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDQ1.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTORDQ1.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdSOTORDQ1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTORDQ1.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDQ1.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdSOTORDQ1.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdSOTORDQ1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTORDQ1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTORDQ1.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdSOTORDQ1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTORDQ1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdSOTORDQ1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTORDQ1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTORDQ1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTORDQ1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTORDQ1.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTORDQ1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdSOTORDQ1.Name = "grdSOTORDQ1"
        Me.grdSOTORDQ1.Size = New System.Drawing.Size(818, 319)
        Me.grdSOTORDQ1.TabIndex = 167
        Me.grdSOTORDQ1.Text = "Open Orders"
        '
        'grdSOTSCAN1
        '
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Appearance13.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTSCAN1.DisplayLayout.Appearance = Appearance13
        UltraGridColumn59.Header.Caption = "#"
        UltraGridColumn59.Header.VisiblePosition = 0
        UltraGridColumn59.Width = 36
        UltraGridColumn60.Header.Caption = "Scan"
        UltraGridColumn60.Header.VisiblePosition = 1
        UltraGridColumn60.Width = 73
        UltraGridColumn61.Header.Caption = "Result"
        UltraGridColumn61.Header.VisiblePosition = 2
        UltraGridColumn61.Width = 358
        UltraGridColumn1.Header.VisiblePosition = 3
        UltraGridColumn1.Hidden = True
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn59, UltraGridColumn60, UltraGridColumn61, UltraGridColumn1})
        Me.grdSOTSCAN1.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.grdSOTSCAN1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance14.TextHAlignAsString = "Left"
        Me.grdSOTSCAN1.DisplayLayout.CaptionAppearance = Appearance14
        Appearance15.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance15.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance15.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTSCAN1.DisplayLayout.GroupByBox.Appearance = Appearance15
        Appearance16.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTSCAN1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance16
        Me.grdSOTSCAN1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTSCAN1.DisplayLayout.GroupByBox.Hidden = True
        Appearance17.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance17.BackColor2 = System.Drawing.SystemColors.Control
        Appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance17.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTSCAN1.DisplayLayout.GroupByBox.PromptAppearance = Appearance17
        Me.grdSOTSCAN1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTSCAN1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTSCAN1.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdSOTSCAN1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance18.BackColor = System.Drawing.SystemColors.Window
        Appearance18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTSCAN1.DisplayLayout.Override.ActiveCellAppearance = Appearance18
        Me.grdSOTSCAN1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTSCAN1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTSCAN1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTSCAN1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance19.BackColor = System.Drawing.SystemColors.Window
        Me.grdSOTSCAN1.DisplayLayout.Override.CardAreaAppearance = Appearance19
        Appearance20.BorderColor = System.Drawing.Color.Silver
        Appearance20.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTSCAN1.DisplayLayout.Override.CellAppearance = Appearance20
        Me.grdSOTSCAN1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTSCAN1.DisplayLayout.Override.CellPadding = 0
        Appearance21.BackColor = System.Drawing.SystemColors.Control
        Appearance21.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance21.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance21.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTSCAN1.DisplayLayout.Override.GroupByRowAppearance = Appearance21
        Appearance22.TextHAlignAsString = "Left"
        Me.grdSOTSCAN1.DisplayLayout.Override.HeaderAppearance = Appearance22
        Me.grdSOTSCAN1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTSCAN1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance23.BackColor = System.Drawing.SystemColors.Window
        Appearance23.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTSCAN1.DisplayLayout.Override.RowAppearance = Appearance23
        Me.grdSOTSCAN1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance24.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTSCAN1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance24
        Me.grdSOTSCAN1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTSCAN1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTSCAN1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTSCAN1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTSCAN1.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTSCAN1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdSOTSCAN1.Name = "grdSOTSCAN1"
        Me.grdSOTSCAN1.Size = New System.Drawing.Size(612, 319)
        Me.grdSOTSCAN1.TabIndex = 173
        Me.grdSOTSCAN1.Text = "Scan History"
        '
        'UltraLabel3
        '
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(10, 98)
        Me.UltraLabel3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(150, 22)
        Me.UltraLabel3.TabIndex = 190
        Me.UltraLabel3.Text = "Available Trucks:"
        '
        'lblInstruction1
        '
        Appearance25.ForeColor = System.Drawing.Color.Blue
        Me.lblInstruction1.Appearance = Appearance25
        Me.lblInstruction1.AutoSize = True
        Me.lblInstruction1.Location = New System.Drawing.Point(168, 98)
        Me.lblInstruction1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.lblInstruction1.Name = "lblInstruction1"
        Me.lblInstruction1.Size = New System.Drawing.Size(101, 22)
        Me.lblInstruction1.TabIndex = 191
        Me.lblInstruction1.Text = "N999,N999"
        '
        'chkCustomTruck
        '
        Me.chkCustomTruck.Location = New System.Drawing.Point(820, 44)
        Me.chkCustomTruck.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.chkCustomTruck.Name = "chkCustomTruck"
        Me.chkCustomTruck.Size = New System.Drawing.Size(155, 21)
        Me.chkCustomTruck.TabIndex = 188
        Me.chkCustomTruck.Text = "Custom Truck"
        Me.chkCustomTruck.Visible = False
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(10, 16)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(105, 22)
        Me.UltraLabel1.TabIndex = 182
        Me.UltraLabel1.Text = "Instruction:"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(1266, 44)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(86, 37)
        Me.cmdCancel.TabIndex = 179
        Me.cmdCancel.Text = "Cancel"
        '
        'lblInstruction2
        '
        Appearance26.ForeColor = System.Drawing.Color.Blue
        Me.lblInstruction2.Appearance = Appearance26
        Me.lblInstruction2.AutoSize = True
        Me.lblInstruction2.Location = New System.Drawing.Point(820, 16)
        Me.lblInstruction2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.lblInstruction2.Name = "lblInstruction2"
        Me.lblInstruction2.Size = New System.Drawing.Size(166, 22)
        Me.lblInstruction2.TabIndex = 187
        Me.lblInstruction2.Text = "Scan Tote in Slot 1"
        Me.lblInstruction2.Visible = False
        '
        'txtSCAN
        '
        Me.txtSCAN.Location = New System.Drawing.Point(1005, 10)
        Me.txtSCAN.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.txtSCAN.Name = "txtSCAN"
        Me.txtSCAN.Size = New System.Drawing.Size(160, 29)
        Me.txtSCAN.TabIndex = 177
        '
        'lblPreConfigured
        '
        Appearance27.ForeColor = System.Drawing.Color.Red
        Me.lblPreConfigured.Appearance = Appearance27
        Me.lblPreConfigured.AutoSize = True
        Me.lblPreConfigured.Location = New System.Drawing.Point(208, 52)
        Me.lblPreConfigured.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.lblPreConfigured.Name = "lblPreConfigured"
        Me.lblPreConfigured.Size = New System.Drawing.Size(174, 22)
        Me.lblPreConfigured.TabIndex = 186
        Me.lblPreConfigured.Text = "Pre-Configured with"
        Me.lblPreConfigured.Visible = False
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Location = New System.Drawing.Point(1174, 44)
        Me.cmdUpdate.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(86, 37)
        Me.cmdUpdate.TabIndex = 178
        Me.cmdUpdate.Text = "Update"
        '
        'lblTruckIsReady
        '
        Appearance28.FontData.SizeInPoints = 24.0!
        Appearance28.ForeColor = System.Drawing.Color.Fuchsia
        Me.lblTruckIsReady.Appearance = Appearance28
        Me.lblTruckIsReady.AutoSize = True
        Me.lblTruckIsReady.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.Solid
        Me.lblTruckIsReady.Location = New System.Drawing.Point(474, 16)
        Me.lblTruckIsReady.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.lblTruckIsReady.Name = "lblTruckIsReady"
        Me.lblTruckIsReady.Size = New System.Drawing.Size(326, 56)
        Me.lblTruckIsReady.TabIndex = 185
        Me.lblTruckIsReady.Text = "Truck Is Ready"
        Me.lblTruckIsReady.Visible = False
        '
        'btnSimulateScan
        '
        Me.btnSimulateScan.Location = New System.Drawing.Point(1174, 7)
        Me.btnSimulateScan.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnSimulateScan.Name = "btnSimulateScan"
        Me.btnSimulateScan.Size = New System.Drawing.Size(178, 37)
        Me.btnSimulateScan.TabIndex = 184
        Me.btnSimulateScan.Text = "Simulate Scan"
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(10, 52)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(53, 22)
        Me.UltraLabel2.TabIndex = 181
        Me.UltraLabel2.Text = "Truck"
        '
        'lblInstruction
        '
        Appearance29.ForeColor = System.Drawing.Color.Blue
        Me.lblInstruction.Appearance = Appearance29
        Me.lblInstruction.AutoSize = True
        Me.lblInstruction.Location = New System.Drawing.Point(122, 16)
        Me.lblInstruction.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.lblInstruction.Name = "lblInstruction"
        Me.lblInstruction.Size = New System.Drawing.Size(115, 22)
        Me.lblInstruction.TabIndex = 183
        Me.lblInstruction.Text = "Scan a Truck"
        '
        'txtTRUCK
        '
        Me.Absx1.SetABSBindToTable(Me.txtTRUCK, False)
        Me.Absx1.SetABSColumnName(Me.txtTRUCK, "TRUCK_NO")
        Me.Absx1.SetABSHasButton(Me.txtTRUCK, True)
        Me.txtTRUCK.Location = New System.Drawing.Point(75, 47)
        Me.txtTRUCK.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtTRUCK.Name = "txtTRUCK"
        Me.txtTRUCK.ReadOnly = True
        Me.txtTRUCK.Size = New System.Drawing.Size(125, 29)
        Me.txtTRUCK.TabIndex = 180
        '
        'SOFPICKU
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1435, 456)
        Me.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.Name = "SOFPICKU"
        Me.Text = "Build a Truck"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.grdSOTORDQ1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTSCAN1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkCustomTruck, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSCAN, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTRUCK, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents chkCustomTruck As ABSCS.ABSCheckBox
    Friend WithEvents UltraLabel1 As Misc.UltraLabel
    Friend WithEvents cmdCancel As Misc.UltraButton
    Friend WithEvents lblInstruction2 As Misc.UltraLabel
    Friend WithEvents txtSCAN As UltraWinEditors.UltraTextEditor
    Friend WithEvents lblPreConfigured As Misc.UltraLabel
    Friend WithEvents cmdUpdate As Misc.UltraButton
    Friend WithEvents lblTruckIsReady As Misc.UltraLabel
    Friend WithEvents btnSimulateScan As Misc.UltraButton
    Friend WithEvents UltraLabel2 As Misc.UltraLabel
    Friend WithEvents lblInstruction As Misc.UltraLabel
    Friend WithEvents txtTRUCK As UltraWinEditors.UltraTextEditor
    Friend WithEvents grdSOTSCAN1 As UltraWinGrid.UltraGrid
    Friend WithEvents grdSOTORDQ1 As UltraWinGrid.UltraGrid
    Friend WithEvents UltraLabel3 As Misc.UltraLabel
    Friend WithEvents lblInstruction1 As Misc.UltraLabel
End Class
