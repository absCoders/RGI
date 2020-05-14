<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SOFRUSSE
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE1
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
        Dim UltraExplorerBarGroup1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup()
        Dim UltraExplorerBarItem2 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim UltraExplorerBarItem3 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim UltraExplorerBarItem4 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim UltraExplorerBarItem6 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDR1", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NO")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NAME")
        Dim UltraGridColumn23 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_FOB")
        Dim UltraGridColumn24 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO")
        Dim UltraGridColumn25 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_DATE")
        Dim UltraGridColumn26 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CANCEL_DATE")
        Dim UltraGridColumn27 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("POST_CODE")
        Dim UltraGridColumn28 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_VIA_CODE")
        Dim UltraGridColumn29 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_INSTR")
        Dim UltraGridColumn30 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TERM_CODE")
        Dim UltraGridColumn31 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SREP_CODE")
        Dim UltraGridColumn32 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn33 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_PICK_SEQ")
        Dim UltraGridColumn34 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REASON_CODE")
        Dim UltraGridColumn35 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SALES_DIVISION_CODE")
        Dim UltraGridColumn36 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn37 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn38 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn39 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn40 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_RECD")
        Dim UltraGridColumn41 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SOURCE")
        Dim UltraGridColumn42 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DEPT")
        Dim UltraGridColumn43 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FRT_TERMS")
        Dim UltraGridColumn44 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ADDR_TYPE_ST")
        Dim UltraGridColumn45 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_BOOKED")
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_YYYYPP_BOOKED")
        Dim UltraGridColumn47 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_CLOSED")
        Dim UltraGridColumn48 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_YYYYPP_CLOSED")
        Dim UltraGridColumn49 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_PRIORITY")
        Dim UltraGridColumn50 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_STATUS")
        Dim UltraGridColumn51 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_JRNL_NO")
        Dim UltraGridColumn52 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_GROUP_NO")
        Dim UltraGridColumn53 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_HOLD")
        Dim UltraGridColumn54 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_HOLD_CODES")
        Dim UltraGridColumn55 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_BILL_TO_CUST")
        Dim UltraGridColumn56 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_DC_NO")
        Dim UltraGridColumn57 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_BATCH_NO")
        Dim UltraGridColumn58 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_DOC_SEQ_NO")
        Dim UltraGridColumn59 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_APPOINTMENT")
        Dim UltraGridColumn60 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_INV_COMMENT")
        Dim UltraGridColumn61 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_FACTOR_IND")
        Dim UltraGridColumn62 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_PRE_ALLOC")
        Dim UltraGridColumn63 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_MERCH_TYPE")
        Dim UltraGridColumn64 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SREP2_CODE")
        Dim UltraGridColumn65 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CURR_CODE")
        Dim UltraGridColumn66 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CURR_EXCH_RATE")
        Dim UltraGridColumn67 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_VALUE_CHANGE_REMARK")
        Dim UltraGridColumn68 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_VALUE_CHANGE_OPER")
        Dim UltraGridColumn69 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_VALUE_CHANGE_DATE")
        Dim UltraGridColumn70 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_DURING_PHY")
        Dim UltraGridColumn71 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ORIG_SHIP_DATE")
        Dim UltraGridColumn72 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ORIG_CANCEL_DATE")
        Dim UltraGridColumn73 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_HOLD_REASON")
        Dim UltraGridColumn74 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SYNC_BATCH")
        Dim UltraGridColumn75 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TORDR")
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn76 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("YEAR")
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance33 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTRUSSE", -1)
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn77 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_LNO")
        Dim UltraGridColumn78 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE")
        Dim UltraGridColumn79 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn80 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_DTL_SEQ")
        Dim UltraGridColumn81 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY")
        Dim UltraGridColumn166 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_DOC_SEQ_NO")
        Dim UltraGridColumn82 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_UPC")
        Dim UltraGridColumn83 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_SKU")
        Dim UltraGridColumn84 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NEW_QTY")
        Dim UltraGridColumn86 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NEW_UPC")
        Dim UltraGridColumn87 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NEW_SKU")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NEW_COLOR_CODE")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NEW_ORDR_UNIT_PRICE")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_UNIT_PRICE")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RANGE_STYLE_CODE")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RANGE_STYLE_QTY_PP", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim Appearance37 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance38 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance39 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance42 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance43 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance44 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance45 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance46 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance47 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraTab5 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab6 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Me.UltraTabPageControl12 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdSOTORDRX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl13 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grdSOTRUSSE = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.chkCreateRanges = New System.Windows.Forms.CheckBox()
        Me.lblNewRange = New System.Windows.Forms.Label()
        Me.txtNewRange = New System.Windows.Forms.TextBox()
        Me.chkUpdateRange = New System.Windows.Forms.CheckBox()
        Me.chkIsRANGE = New System.Windows.Forms.CheckBox()
        Me.spl = New System.Windows.Forms.SplitContainer()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor5 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel14 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor3 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtCUST_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.tab = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage4 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl12.SuspendLayout()
        CType(Me.grdSOTORDRX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl13.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grdSOTRUSSE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.spl.Panel1.SuspendLayout()
        Me.spl.Panel2.SuspendLayout()
        Me.spl.SuspendLayout()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tab, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraExplorerBar1
        '
        UltraExplorerBarItem2.Key = "Edit"
        UltraExplorerBarItem2.Text = "Edit"
        UltraExplorerBarItem3.Key = "Update"
        UltraExplorerBarItem3.Text = "Update"
        UltraExplorerBarItem4.Key = "Cancel"
        UltraExplorerBarItem4.Text = "Cancel"
        UltraExplorerBarItem6.Key = "Done"
        UltraExplorerBarItem6.Text = "Done"
        UltraExplorerBarGroup1.Items.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem() {UltraExplorerBarItem2, UltraExplorerBarItem3, UltraExplorerBarItem4, UltraExplorerBarItem6})
        UltraExplorerBarGroup1.Key = "Screen Control"
        UltraExplorerBarGroup1.Text = "Screen Control"
        Me.UltraExplorerBar1.Groups.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup() {UltraExplorerBarGroup1})
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Margins.Bottom = 0
        Me.UltraExplorerBar1.Margins.Left = 0
        Me.UltraExplorerBar1.Margins.Right = 0
        Me.UltraExplorerBar1.Margins.Top = 0
        Me.UltraExplorerBar1.ShowDefaultContextMenu = False
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 554)
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.spl)
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(777, 574)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.grdASFBASEX, 0)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.spl, 0)
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
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Appearance10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance10
        Appearance11.BackColor = System.Drawing.SystemColors.Highlight
        Appearance11.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance11
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance12.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance12
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Appearance14.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance14
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
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(990, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 574)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(990, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 574)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(990, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraTabPageControl12
        '
        Me.UltraTabPageControl12.Controls.Add(Me.grdSOTORDRX)
        Me.UltraTabPageControl12.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl12.Name = "UltraTabPageControl12"
        Me.UltraTabPageControl12.Size = New System.Drawing.Size(773, 474)
        '
        'grdSOTORDRX
        '
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTORDRX.DisplayLayout.Appearance = Appearance2
        UltraGridColumn1.Header.Caption = "Order No"
        UltraGridColumn1.Header.VisiblePosition = 1
        UltraGridColumn1.Width = 99
        UltraGridColumn2.Header.Caption = "Order Date"
        UltraGridColumn2.Header.VisiblePosition = 3
        UltraGridColumn2.Width = 90
        UltraGridColumn3.Header.Caption = "Cust Code"
        UltraGridColumn3.Header.VisiblePosition = 5
        UltraGridColumn3.Width = 88
        UltraGridColumn4.Header.Caption = "Cust Name"
        UltraGridColumn4.Header.VisiblePosition = 6
        UltraGridColumn4.Width = 193
        UltraGridColumn5.Header.Caption = "Store no"
        UltraGridColumn5.Header.VisiblePosition = 7
        UltraGridColumn5.Width = 85
        UltraGridColumn6.Header.Caption = "Store Name"
        UltraGridColumn6.Header.VisiblePosition = 8
        UltraGridColumn6.Width = 178
        UltraGridColumn23.Header.VisiblePosition = 9
        UltraGridColumn23.Hidden = True
        UltraGridColumn24.Header.Caption = "Cust PO"
        UltraGridColumn24.Header.VisiblePosition = 10
        UltraGridColumn24.Width = 121
        UltraGridColumn25.Header.Caption = "Ship Date"
        UltraGridColumn25.Header.VisiblePosition = 11
        UltraGridColumn25.Width = 90
        UltraGridColumn26.Header.Caption = "Cancel Date"
        UltraGridColumn26.Header.VisiblePosition = 12
        UltraGridColumn26.Width = 105
        UltraGridColumn27.Header.VisiblePosition = 13
        UltraGridColumn27.Hidden = True
        UltraGridColumn28.Header.Caption = "Ship Via"
        UltraGridColumn28.Header.VisiblePosition = 17
        UltraGridColumn28.Width = 68
        UltraGridColumn29.Header.VisiblePosition = 14
        UltraGridColumn29.Hidden = True
        UltraGridColumn30.Header.VisiblePosition = 15
        UltraGridColumn30.Hidden = True
        UltraGridColumn31.Header.VisiblePosition = 16
        UltraGridColumn31.Hidden = True
        UltraGridColumn32.Header.Caption = "Whse"
        UltraGridColumn32.Header.VisiblePosition = 18
        UltraGridColumn32.Width = 46
        UltraGridColumn33.Header.VisiblePosition = 19
        UltraGridColumn33.Hidden = True
        UltraGridColumn34.Header.VisiblePosition = 20
        UltraGridColumn34.Hidden = True
        UltraGridColumn35.Header.VisiblePosition = 21
        UltraGridColumn35.Hidden = True
        UltraGridColumn36.Header.VisiblePosition = 22
        UltraGridColumn36.Hidden = True
        UltraGridColumn37.Header.VisiblePosition = 23
        UltraGridColumn37.Hidden = True
        UltraGridColumn38.Header.VisiblePosition = 24
        UltraGridColumn38.Hidden = True
        UltraGridColumn39.Header.VisiblePosition = 25
        UltraGridColumn39.Hidden = True
        UltraGridColumn40.Header.VisiblePosition = 26
        UltraGridColumn40.Hidden = True
        UltraGridColumn41.Header.VisiblePosition = 27
        UltraGridColumn41.Hidden = True
        UltraGridColumn42.Header.VisiblePosition = 28
        UltraGridColumn42.Hidden = True
        UltraGridColumn43.Header.VisiblePosition = 29
        UltraGridColumn43.Hidden = True
        UltraGridColumn44.Header.VisiblePosition = 30
        UltraGridColumn44.Hidden = True
        UltraGridColumn45.Header.VisiblePosition = 31
        UltraGridColumn45.Hidden = True
        UltraGridColumn46.Header.VisiblePosition = 32
        UltraGridColumn46.Hidden = True
        UltraGridColumn47.Header.VisiblePosition = 33
        UltraGridColumn47.Hidden = True
        UltraGridColumn48.Header.VisiblePosition = 34
        UltraGridColumn48.Hidden = True
        UltraGridColumn49.Header.VisiblePosition = 35
        UltraGridColumn49.Hidden = True
        UltraGridColumn50.Header.Caption = "Status"
        UltraGridColumn50.Header.VisiblePosition = 2
        UltraGridColumn50.Width = 82
        UltraGridColumn51.Header.VisiblePosition = 36
        UltraGridColumn51.Hidden = True
        UltraGridColumn52.Header.Caption = "Group"
        UltraGridColumn52.Header.VisiblePosition = 0
        UltraGridColumn52.Width = 98
        UltraGridColumn53.Header.VisiblePosition = 37
        UltraGridColumn53.Hidden = True
        UltraGridColumn54.Header.VisiblePosition = 38
        UltraGridColumn54.Hidden = True
        UltraGridColumn55.Header.VisiblePosition = 39
        UltraGridColumn55.Hidden = True
        UltraGridColumn56.Header.VisiblePosition = 40
        UltraGridColumn56.Hidden = True
        UltraGridColumn57.Header.VisiblePosition = 41
        UltraGridColumn57.Hidden = True
        UltraGridColumn58.Header.VisiblePosition = 42
        UltraGridColumn58.Hidden = True
        UltraGridColumn59.Header.VisiblePosition = 43
        UltraGridColumn59.Hidden = True
        UltraGridColumn60.Header.VisiblePosition = 44
        UltraGridColumn60.Hidden = True
        UltraGridColumn61.Header.VisiblePosition = 45
        UltraGridColumn61.Hidden = True
        UltraGridColumn62.Header.VisiblePosition = 46
        UltraGridColumn62.Hidden = True
        UltraGridColumn63.Header.VisiblePosition = 47
        UltraGridColumn63.Hidden = True
        UltraGridColumn64.Header.VisiblePosition = 48
        UltraGridColumn64.Hidden = True
        UltraGridColumn65.Header.VisiblePosition = 49
        UltraGridColumn65.Hidden = True
        UltraGridColumn66.Header.VisiblePosition = 50
        UltraGridColumn66.Hidden = True
        UltraGridColumn67.Header.VisiblePosition = 51
        UltraGridColumn67.Hidden = True
        UltraGridColumn68.Header.VisiblePosition = 52
        UltraGridColumn68.Hidden = True
        UltraGridColumn69.Header.VisiblePosition = 53
        UltraGridColumn69.Hidden = True
        UltraGridColumn70.Header.VisiblePosition = 54
        UltraGridColumn70.Hidden = True
        UltraGridColumn71.Header.VisiblePosition = 55
        UltraGridColumn71.Hidden = True
        UltraGridColumn72.Header.VisiblePosition = 56
        UltraGridColumn72.Hidden = True
        UltraGridColumn73.Header.VisiblePosition = 57
        UltraGridColumn73.Hidden = True
        UltraGridColumn74.Header.VisiblePosition = 58
        UltraGridColumn74.Hidden = True
        Appearance3.TextHAlignAsString = "Right"
        UltraGridColumn75.CellAppearance = Appearance3
        Appearance22.TextHAlignAsString = "Right"
        UltraGridColumn75.Header.Appearance = Appearance22
        UltraGridColumn75.Header.Caption = "Total"
        UltraGridColumn75.Header.VisiblePosition = 59
        UltraGridColumn75.Width = 117
        UltraGridColumn76.Header.Caption = "Year"
        UltraGridColumn76.Header.VisiblePosition = 4
        UltraGridColumn76.Width = 68
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn23, UltraGridColumn24, UltraGridColumn25, UltraGridColumn26, UltraGridColumn27, UltraGridColumn28, UltraGridColumn29, UltraGridColumn30, UltraGridColumn31, UltraGridColumn32, UltraGridColumn33, UltraGridColumn34, UltraGridColumn35, UltraGridColumn36, UltraGridColumn37, UltraGridColumn38, UltraGridColumn39, UltraGridColumn40, UltraGridColumn41, UltraGridColumn42, UltraGridColumn43, UltraGridColumn44, UltraGridColumn45, UltraGridColumn46, UltraGridColumn47, UltraGridColumn48, UltraGridColumn49, UltraGridColumn50, UltraGridColumn51, UltraGridColumn52, UltraGridColumn53, UltraGridColumn54, UltraGridColumn55, UltraGridColumn56, UltraGridColumn57, UltraGridColumn58, UltraGridColumn59, UltraGridColumn60, UltraGridColumn61, UltraGridColumn62, UltraGridColumn63, UltraGridColumn64, UltraGridColumn65, UltraGridColumn66, UltraGridColumn67, UltraGridColumn68, UltraGridColumn69, UltraGridColumn70, UltraGridColumn71, UltraGridColumn72, UltraGridColumn73, UltraGridColumn74, UltraGridColumn75, UltraGridColumn76})
        Me.grdSOTORDRX.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTORDRX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance23.TextHAlignAsString = "Left"
        Me.grdSOTORDRX.DisplayLayout.CaptionAppearance = Appearance23
        Me.grdSOTORDRX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance24.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance24.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance24.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDRX.DisplayLayout.GroupByBox.Appearance = Appearance24
        Appearance25.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDRX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance25
        Me.grdSOTORDRX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTORDRX.DisplayLayout.GroupByBox.Hidden = True
        Appearance26.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance26.BackColor2 = System.Drawing.SystemColors.Control
        Appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance26.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDRX.DisplayLayout.GroupByBox.PromptAppearance = Appearance26
        Me.grdSOTORDRX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTORDRX.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTORDRX.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance27.BackColor = System.Drawing.SystemColors.Window
        Appearance27.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTORDRX.DisplayLayout.Override.ActiveCellAppearance = Appearance27
        Me.grdSOTORDRX.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTORDRX.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDRX.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDRX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTORDRX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance28.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTORDRX.DisplayLayout.Override.CardAreaAppearance = Appearance28
        Appearance29.BorderColor = System.Drawing.Color.Silver
        Appearance29.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTORDRX.DisplayLayout.Override.CellAppearance = Appearance29
        Me.grdSOTORDRX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdSOTORDRX.DisplayLayout.Override.CellPadding = 0
        Appearance30.BackColor = System.Drawing.SystemColors.Control
        Appearance30.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance30.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance30.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDRX.DisplayLayout.Override.GroupByRowAppearance = Appearance30
        Appearance31.TextHAlignAsString = "Left"
        Me.grdSOTORDRX.DisplayLayout.Override.HeaderAppearance = Appearance31
        Me.grdSOTORDRX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTORDRX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance32.BackColor = System.Drawing.SystemColors.Window
        Appearance32.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTORDRX.DisplayLayout.Override.RowAppearance = Appearance32
        Me.grdSOTORDRX.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance33.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTORDRX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance33
        Me.grdSOTORDRX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTORDRX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTORDRX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTORDRX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTORDRX.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTORDRX.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTORDRX.Name = "grdSOTORDRX"
        Me.grdSOTORDRX.Size = New System.Drawing.Size(773, 474)
        Me.grdSOTORDRX.TabIndex = 13
        Me.grdSOTORDRX.Text = "Orders"
        '
        'UltraTabPageControl13
        '
        Me.UltraTabPageControl13.Controls.Add(Me.SplitContainer1)
        Me.UltraTabPageControl13.Location = New System.Drawing.Point(1, 25)
        Me.UltraTabPageControl13.Name = "UltraTabPageControl13"
        Me.UltraTabPageControl13.Size = New System.Drawing.Size(773, 474)
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdSOTRUSSE)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Size = New System.Drawing.Size(773, 474)
        Me.SplitContainer1.SplitterDistance = 257
        Me.SplitContainer1.TabIndex = 0
        '
        'grdSOTRUSSE
        '
        Appearance36.BackColor = System.Drawing.SystemColors.Window
        Appearance36.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTRUSSE.DisplayLayout.Appearance = Appearance36
        UltraGridColumn7.Header.Caption = "Order No"
        UltraGridColumn7.Header.VisiblePosition = 0
        UltraGridColumn7.Hidden = True
        UltraGridColumn7.Width = 99
        UltraGridColumn77.Header.Caption = "Line"
        UltraGridColumn77.Header.VisiblePosition = 1
        UltraGridColumn77.Width = 44
        UltraGridColumn78.Header.Caption = "Style Code"
        UltraGridColumn78.Header.VisiblePosition = 2
        UltraGridColumn78.Width = 128
        UltraGridColumn79.Header.Caption = "Color Code"
        UltraGridColumn79.Header.VisiblePosition = 3
        UltraGridColumn80.Header.VisiblePosition = 4
        UltraGridColumn80.Hidden = True
        UltraGridColumn81.Header.Caption = "Order Qty"
        UltraGridColumn81.Header.VisiblePosition = 5
        UltraGridColumn81.Width = 87
        UltraGridColumn166.Header.VisiblePosition = 12
        UltraGridColumn166.Hidden = True
        UltraGridColumn82.Header.Caption = "Customer UPC"
        UltraGridColumn82.Header.VisiblePosition = 7
        UltraGridColumn83.Header.Caption = "Customer SKU"
        UltraGridColumn83.Header.VisiblePosition = 8
        UltraGridColumn84.Header.Caption = "New Qty"
        UltraGridColumn84.Header.VisiblePosition = 9
        UltraGridColumn84.Width = 112
        UltraGridColumn86.Header.Caption = "New UPC"
        UltraGridColumn86.Header.VisiblePosition = 10
        UltraGridColumn87.Header.Caption = "New SKU"
        UltraGridColumn87.Header.VisiblePosition = 11
        UltraGridColumn8.Header.Caption = "New Color Code"
        UltraGridColumn8.Header.VisiblePosition = 13
        UltraGridColumn9.Header.Caption = "New Price"
        UltraGridColumn9.Header.VisiblePosition = 14
        UltraGridColumn10.Header.Caption = "Price"
        UltraGridColumn10.Header.VisiblePosition = 6
        UltraGridColumn11.Header.Caption = "Range Style Code"
        UltraGridColumn11.Header.VisiblePosition = 15
        UltraGridColumn11.Hidden = True
        UltraGridColumn12.Header.Caption = "Pack/Range"
        UltraGridColumn12.Header.VisiblePosition = 16
        UltraGridColumn12.Hidden = True
        UltraGridColumn12.Width = 99
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn7, UltraGridColumn77, UltraGridColumn78, UltraGridColumn79, UltraGridColumn80, UltraGridColumn81, UltraGridColumn166, UltraGridColumn82, UltraGridColumn83, UltraGridColumn84, UltraGridColumn86, UltraGridColumn87, UltraGridColumn8, UltraGridColumn9, UltraGridColumn10, UltraGridColumn11, UltraGridColumn12})
        Me.grdSOTRUSSE.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.grdSOTRUSSE.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance37.TextHAlignAsString = "Left"
        Me.grdSOTRUSSE.DisplayLayout.CaptionAppearance = Appearance37
        Me.grdSOTRUSSE.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance38.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance38.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance38.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTRUSSE.DisplayLayout.GroupByBox.Appearance = Appearance38
        Appearance39.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTRUSSE.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance39
        Me.grdSOTRUSSE.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTRUSSE.DisplayLayout.GroupByBox.Hidden = True
        Appearance40.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance40.BackColor2 = System.Drawing.SystemColors.Control
        Appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance40.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTRUSSE.DisplayLayout.GroupByBox.PromptAppearance = Appearance40
        Me.grdSOTRUSSE.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTRUSSE.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTRUSSE.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance41.BackColor = System.Drawing.SystemColors.Window
        Appearance41.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTRUSSE.DisplayLayout.Override.ActiveCellAppearance = Appearance41
        Me.grdSOTRUSSE.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTRUSSE.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTRUSSE.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTRUSSE.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTRUSSE.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance42.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTRUSSE.DisplayLayout.Override.CardAreaAppearance = Appearance42
        Appearance43.BorderColor = System.Drawing.Color.Silver
        Appearance43.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTRUSSE.DisplayLayout.Override.CellAppearance = Appearance43
        Me.grdSOTRUSSE.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdSOTRUSSE.DisplayLayout.Override.CellPadding = 0
        Appearance44.BackColor = System.Drawing.SystemColors.Control
        Appearance44.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance44.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance44.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTRUSSE.DisplayLayout.Override.GroupByRowAppearance = Appearance44
        Appearance45.TextHAlignAsString = "Left"
        Me.grdSOTRUSSE.DisplayLayout.Override.HeaderAppearance = Appearance45
        Me.grdSOTRUSSE.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTRUSSE.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance46.BackColor = System.Drawing.SystemColors.Window
        Appearance46.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTRUSSE.DisplayLayout.Override.RowAppearance = Appearance46
        Me.grdSOTRUSSE.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance47.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTRUSSE.DisplayLayout.Override.TemplateAddRowAppearance = Appearance47
        Me.grdSOTRUSSE.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTRUSSE.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTRUSSE.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTRUSSE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTRUSSE.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTRUSSE.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTRUSSE.Name = "grdSOTRUSSE"
        Me.grdSOTRUSSE.Size = New System.Drawing.Size(773, 257)
        Me.grdSOTRUSSE.TabIndex = 15
        Me.grdSOTRUSSE.Text = "Order Details"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkCreateRanges)
        Me.Panel1.Controls.Add(Me.lblNewRange)
        Me.Panel1.Controls.Add(Me.txtNewRange)
        Me.Panel1.Controls.Add(Me.chkUpdateRange)
        Me.Panel1.Controls.Add(Me.chkIsRANGE)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(773, 213)
        Me.Panel1.TabIndex = 0
        '
        'chkCreateRanges
        '
        Me.chkCreateRanges.AutoSize = True
        Me.chkCreateRanges.Location = New System.Drawing.Point(283, 17)
        Me.chkCreateRanges.Name = "chkCreateRanges"
        Me.chkCreateRanges.Size = New System.Drawing.Size(195, 20)
        Me.chkCreateRanges.TabIndex = 4
        Me.chkCreateRanges.Text = "Create New Range Styles"
        Me.chkCreateRanges.UseVisualStyleBackColor = True
        Me.chkCreateRanges.Visible = False
        '
        'lblNewRange
        '
        Me.lblNewRange.AutoSize = True
        Me.lblNewRange.Location = New System.Drawing.Point(51, 66)
        Me.lblNewRange.Name = "lblNewRange"
        Me.lblNewRange.Size = New System.Drawing.Size(171, 16)
        Me.lblNewRange.TabIndex = 3
        Me.lblNewRange.Text = "To The New Range Style"
        Me.lblNewRange.Visible = False
        '
        'txtNewRange
        '
        Me.txtNewRange.Location = New System.Drawing.Point(54, 85)
        Me.txtNewRange.Name = "txtNewRange"
        Me.txtNewRange.Size = New System.Drawing.Size(168, 23)
        Me.txtNewRange.TabIndex = 2
        Me.txtNewRange.Visible = False
        '
        'chkUpdateRange
        '
        Me.chkUpdateRange.AutoSize = True
        Me.chkUpdateRange.Location = New System.Drawing.Point(30, 43)
        Me.chkUpdateRange.Name = "chkUpdateRange"
        Me.chkUpdateRange.Size = New System.Drawing.Size(212, 20)
        Me.chkUpdateRange.TabIndex = 1
        Me.chkUpdateRange.Text = "Update Current Range Style"
        Me.chkUpdateRange.UseVisualStyleBackColor = True
        Me.chkUpdateRange.Visible = False
        '
        'chkIsRANGE
        '
        Me.chkIsRANGE.AutoSize = True
        Me.chkIsRANGE.Enabled = False
        Me.chkIsRANGE.Location = New System.Drawing.Point(30, 17)
        Me.chkIsRANGE.Name = "chkIsRANGE"
        Me.chkIsRANGE.Size = New System.Drawing.Size(166, 20)
        Me.chkIsRANGE.TabIndex = 0
        Me.chkIsRANGE.Text = "Order Is Range Order"
        Me.chkIsRANGE.UseVisualStyleBackColor = True
        '
        'spl
        '
        Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.spl.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.spl.Location = New System.Drawing.Point(0, 0)
        Me.spl.Name = "spl"
        Me.spl.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'spl.Panel1
        '
        Me.spl.Panel1.Controls.Add(Me.UltraGroupBox1)
        '
        'spl.Panel2
        '
        Me.spl.Panel2.Controls.Add(Me.tab)
        Me.spl.Size = New System.Drawing.Size(777, 574)
        Me.spl.SplitterDistance = 68
        Me.spl.TabIndex = 4
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.UltraGroupBox1.Controls.Add(Me.UltraTextEditor1)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel4)
        Me.UltraGroupBox1.Controls.Add(Me.UltraTextEditor5)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel14)
        Me.UltraGroupBox1.Controls.Add(Me.UltraTextEditor3)
        Me.UltraGroupBox1.Controls.Add(Me.txtCUST_CODE)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(777, 68)
        Me.UltraGroupBox1.TabIndex = 6
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "ORDR_NO")
        Me.UltraTextEditor1.Location = New System.Drawing.Point(466, 36)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.ReadOnly = True
        Me.UltraTextEditor1.Size = New System.Drawing.Size(150, 25)
        Me.UltraTextEditor1.TabIndex = 107
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(466, 12)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(66, 18)
        Me.UltraLabel4.TabIndex = 108
        Me.UltraLabel4.Text = "Order No"
        '
        'UltraTextEditor5
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor5, "ORDR_CUST_PO")
        Me.UltraTextEditor5.Location = New System.Drawing.Point(622, 36)
        Me.UltraTextEditor5.Name = "UltraTextEditor5"
        Me.UltraTextEditor5.ReadOnly = True
        Me.UltraTextEditor5.Size = New System.Drawing.Size(115, 25)
        Me.UltraTextEditor5.TabIndex = 2
        Me.UltraTextEditor5.Visible = False
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(622, 12)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(93, 18)
        Me.UltraLabel1.TabIndex = 106
        Me.UltraLabel1.Text = "Customer PO"
        Me.UltraLabel1.Visible = False
        '
        'UltraLabel14
        '
        Me.UltraLabel14.AutoSize = True
        Me.UltraLabel14.Location = New System.Drawing.Point(128, 12)
        Me.UltraLabel14.Name = "UltraLabel14"
        Me.UltraLabel14.Size = New System.Drawing.Size(44, 18)
        Me.UltraLabel14.TabIndex = 102
        Me.UltraLabel14.Text = "Name"
        '
        'UltraTextEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor3, "CUST_NAME")
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor3, "CUST_CODE")
        Me.Absx1.SetABSTableName(Me.UltraTextEditor3, "SOTORDRB")
        Me.UltraTextEditor3.Location = New System.Drawing.Point(128, 36)
        Me.UltraTextEditor3.Name = "UltraTextEditor3"
        Me.UltraTextEditor3.ReadOnly = True
        Me.UltraTextEditor3.Size = New System.Drawing.Size(332, 25)
        Me.UltraTextEditor3.TabIndex = 5
        Me.UltraTextEditor3.TabStop = False
        Me.UltraTextEditor3.Text = "TARGET NORTHERN OPERATION"
        '
        'txtCUST_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtCUST_CODE, "CUST_CODE")
        Me.Absx1.SetABSHasButton(Me.txtCUST_CODE, True)
        Me.Absx1.SetABSTableName(Me.txtCUST_CODE, "SOTORDRB")
        Me.txtCUST_CODE.Location = New System.Drawing.Point(13, 36)
        Me.txtCUST_CODE.Name = "txtCUST_CODE"
        Me.txtCUST_CODE.Size = New System.Drawing.Size(109, 25)
        Me.txtCUST_CODE.TabIndex = 0
        Me.txtCUST_CODE.Text = "TARGET"
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(13, 12)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(70, 18)
        Me.UltraLabel2.TabIndex = 3
        Me.UltraLabel2.Text = "Customer"
        '
        'tab
        '
        Me.tab.Controls.Add(Me.UltraTabSharedControlsPage4)
        Me.tab.Controls.Add(Me.UltraTabPageControl12)
        Me.tab.Controls.Add(Me.UltraTabPageControl13)
        Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tab.Location = New System.Drawing.Point(0, 0)
        Me.tab.Name = "tab"
        Me.tab.SharedControlsPage = Me.UltraTabSharedControlsPage4
        Me.tab.Size = New System.Drawing.Size(777, 502)
        Me.tab.TabIndex = 1
        UltraTab5.TabPage = Me.UltraTabPageControl12
        UltraTab5.Text = "0"
        UltraTab6.TabPage = Me.UltraTabPageControl13
        UltraTab6.Text = "1"
        Me.tab.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab5, UltraTab6})
        '
        'UltraTabSharedControlsPage4
        '
        Me.UltraTabSharedControlsPage4.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage4.Name = "UltraTabSharedControlsPage4"
        Me.UltraTabSharedControlsPage4.Size = New System.Drawing.Size(773, 474)
        '
        'SOFRUSSE
        '
        Me.Absx1.SetABSTableName(Me, "SOTORDR0")
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(990, 574)
        Me.Name = "SOFRUSSE"
        Me.Text = "SOFRUSSE"
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl12.ResumeLayout(False)
        CType(Me.grdSOTORDRX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl13.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grdSOTRUSSE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.spl.Panel1.ResumeLayout(False)
        Me.spl.Panel2.ResumeLayout(False)
        CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.ResumeLayout(False)
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.UltraGroupBox1.PerformLayout()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tab, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents spl As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraTextEditor5 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel14 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor3 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtCUST_CODE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents tab As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage4 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl12 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents grdSOTORDRX As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraTabPageControl13 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdSOTRUSSE As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblNewRange As System.Windows.Forms.Label
    Friend WithEvents txtNewRange As System.Windows.Forms.TextBox
    Friend WithEvents chkUpdateRange As System.Windows.Forms.CheckBox
    Friend WithEvents chkIsRANGE As System.Windows.Forms.CheckBox
    Friend WithEvents chkCreateRanges As System.Windows.Forms.CheckBox
End Class
