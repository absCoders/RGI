<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SOFCARTP
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
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDR0", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_TYPE")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_GROUP_NO")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_DC_NO")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DEPT")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_MERCH_TYPE")
        Dim UltraGridColumn15 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SALES_DIVISION_CODE")
        Dim UltraGridColumn16 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE")
        Dim UltraGridColumn23 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_DATE")
        Dim UltraGridColumn24 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CANCEL_DATE")
        Dim UltraGridColumn25 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ORIG_SHIP_DATE")
        Dim UltraGridColumn26 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ORIG_CANCEL_DATE")
        Dim UltraGridColumn27 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn28 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SREP_CODE")
        Dim UltraGridColumn29 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_TYPE_CODE")
        Dim UltraGridColumn30 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SOURCE")
        Dim UltraGridColumn31 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT")
        Dim UltraGridColumn32 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT_OPEN")
        Dim UltraGridColumn33 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT_PICK")
        Dim UltraGridColumn34 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT_SHIP")
        Dim UltraGridColumn35 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT_CANC")
        Dim UltraGridColumn36 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY")
        Dim UltraGridColumn37 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_OPEN")
        Dim UltraGridColumn38 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PICK")
        Dim UltraGridColumn39 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_SHIP")
        Dim UltraGridColumn40 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_CANC")
        Dim UltraGridColumn41 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CNT")
        Dim UltraGridColumn42 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CNT_OPEN")
        Dim UltraGridColumn43 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CNT_PICK")
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE_RECD")
        Dim UltraGridColumn47 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_PRIORITY")
        Dim UltraGridColumn48 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ARRIVAL_DATE")
        Dim UltraGridColumn49 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_LAST_ARRIVAL_DATE")
        Dim UltraGridColumn50 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO_MIN")
        Dim UltraGridColumn51 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO_MAX")
        Dim UltraGridColumn52 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_RELEASE_AVAIL_MIN")
        Dim UltraGridColumn53 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_RELEASE_AVAIL_MAX")
        Dim UltraGridColumn54 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_SHORT")
        Dim UltraGridColumn56 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_SHORT_OPER")
        Dim UltraGridColumn57 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_ACTION_DATE")
        Dim UltraGridColumn59 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_REL_ACTION_OPER")
        Dim UltraGridColumn60 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_CONS_NO")
        Dim UltraGridColumn61 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TERM_CODE", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn62 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn63 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn64 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_INSTR")
        Dim UltraGridColumn65 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_MESSAGE")
        Dim UltraGridColumn66 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_EXP_DATE")
        Dim UltraGridColumn67 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_LAST4")
        Dim UltraGridColumn68 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME")
        Dim UltraGridColumn69 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CITY")
        Dim UltraGridColumn70 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STATE")
        Dim UltraGridColumn71 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_COUNTRY")
        Dim UltraGridColumn72 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WAVE_NO")
        Dim UltraGridColumn74 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_LOAD_ID")
        Dim UltraGridColumn75 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT_ALLO_CUR")
        Dim UltraGridColumn76 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT_ALLO_FUT")
        Dim UltraGridColumn77 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_AMT_ALLO_CXL")
        Dim UltraGridColumn78 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PCT_ALLO_CUR")
        Dim UltraGridColumn79 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PCT_ALLO_FUT")
        Dim UltraGridColumn80 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PCT_ALLO_CXL")
        Dim UltraGridColumn81 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EDI_PO_TYPE")
        Dim UltraGridColumn90 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_NO")
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
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTPCKC1", -1)
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_NO")
        Dim UltraGridColumn154 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_CONFIG_NO")
        Dim UltraGridColumn155 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn113 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_OPEN")
        Dim UltraGridColumn156 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLES")
        Dim UltraGridColumn157 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDERS")
        Dim UltraGridColumn44 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACKS")
        Dim UltraGridColumn73 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_TOTAL")
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
        Dim UltraGridBand3 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDRO", -1)
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_GROUP_NO")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NO")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_DC_NO")
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_CONFIG_NO")
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance33 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance34 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance37 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance38 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand4 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTPCKC2", -1)
        Dim UltraGridColumn18 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_NO")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_CONFIG_NO")
        Dim UltraGridColumn19 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_LNO")
        Dim UltraGridColumn20 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE")
        Dim UltraGridColumn21 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn22 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_OPEN")
        Dim UltraGridColumn45 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_1")
        Dim UltraGridColumn55 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_2")
        Dim UltraGridColumn58 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_3")
        Dim UltraGridColumn82 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_4")
        Dim UltraGridColumn83 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_5")
        Dim UltraGridColumn84 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_6")
        Dim UltraGridColumn85 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_7")
        Dim UltraGridColumn86 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_8")
        Dim UltraGridColumn87 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK_9")
        Dim UltraGridColumn88 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_PACK")
        Dim UltraGridColumn89 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_QTY_LEFT")
        Dim Appearance39 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance42 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance43 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance44 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance45 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance46 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance47 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance48 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance49 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.UltraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.cmdDelete = New Infragistics.Win.Misc.UltraButton()
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton()
        Me.UltraLabel14 = New Infragistics.Win.Misc.UltraLabel()
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton()
        Me.UltraTextEditor3 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraTextEditor6 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.grdSOTORDR0 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.grdSOTPCKC1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdSOTPCKC4 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdSOTPCKC2 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox2.SuspendLayout()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.grdSOTORDR0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.grdSOTPCKC1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTPCKC4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTPCKC2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(1184, 661)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 661)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(1184, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 661)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(1184, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 661)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(1184, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.UltraGroupBox2)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(1184, 661)
        Me.SplitContainer2.SplitterDistance = 54
        Me.SplitContainer2.TabIndex = 1
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor1)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox2.Controls.Add(Me.cmdDelete)
        Me.UltraGroupBox2.Controls.Add(Me.cmdCancel)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel14)
        Me.UltraGroupBox2.Controls.Add(Me.cmdUpdate)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor3)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor6)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel4)
        Me.UltraGroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(1184, 54)
        Me.UltraGroupBox2.TabIndex = 0
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor1, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "PACK_NO")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(1099, 23)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(74, 25)
        Me.UltraTextEditor1.TabIndex = 108
        '
        'UltraLabel1
        '
        Me.UltraLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(1099, 5)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(59, 18)
        Me.UltraLabel1.TabIndex = 109
        Me.UltraLabel1.Text = "Pack No"
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance1.ForeColor = System.Drawing.Color.Red
        Me.cmdDelete.Appearance = Appearance1
        Me.cmdDelete.Location = New System.Drawing.Point(680, 21)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(83, 30)
        Me.cmdDelete.TabIndex = 107
        Me.cmdDelete.Text = "Delete"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(947, 21)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(83, 30)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "Cancel"
        '
        'UltraLabel14
        '
        Me.UltraLabel14.AutoSize = True
        Me.UltraLabel14.Location = New System.Drawing.Point(184, 5)
        Me.UltraLabel14.Name = "UltraLabel14"
        Me.UltraLabel14.Size = New System.Drawing.Size(44, 18)
        Me.UltraLabel14.TabIndex = 106
        Me.UltraLabel14.Text = "Name"
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Location = New System.Drawing.Point(863, 21)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(83, 30)
        Me.cmdUpdate.TabIndex = 15
        Me.cmdUpdate.Text = "Update"
        '
        'UltraTextEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor3, "CUST_NAME")
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor3, "CUST_CODE")
        Me.UltraTextEditor3.Location = New System.Drawing.Point(184, 23)
        Me.UltraTextEditor3.Name = "UltraTextEditor3"
        Me.UltraTextEditor3.ReadOnly = True
        Me.UltraTextEditor3.Size = New System.Drawing.Size(290, 25)
        Me.UltraTextEditor3.TabIndex = 105
        Me.UltraTextEditor3.TabStop = False
        '
        'UltraTextEditor6
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor6, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor6, "CUST_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor6, True)
        Me.UltraTextEditor6.Location = New System.Drawing.Point(6, 23)
        Me.UltraTextEditor6.Name = "UltraTextEditor6"
        Me.UltraTextEditor6.Size = New System.Drawing.Size(161, 25)
        Me.UltraTextEditor6.TabIndex = 103
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(6, 5)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(70, 18)
        Me.UltraLabel4.TabIndex = 104
        Me.UltraLabel4.Text = "Customer"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.grdSOTORDR0)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer5)
        Me.SplitContainer3.Size = New System.Drawing.Size(1184, 603)
        Me.SplitContainer3.SplitterDistance = 191
        Me.SplitContainer3.TabIndex = 15
        '
        'grdSOTORDR0
        '
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTORDR0.DisplayLayout.Appearance = Appearance2
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Hidden = True
        UltraGridColumn2.Header.Caption = "Group"
        UltraGridColumn2.Header.VisiblePosition = 6
        UltraGridColumn2.Width = 102
        UltraGridColumn3.Header.Caption = "Customer"
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Width = 93
        UltraGridColumn4.Header.Caption = "Customer PO"
        UltraGridColumn4.Header.VisiblePosition = 7
        UltraGridColumn4.Width = 107
        UltraGridColumn6.Header.Caption = "DC"
        UltraGridColumn6.Header.VisiblePosition = 12
        UltraGridColumn6.Width = 72
        UltraGridColumn7.Header.Caption = "Dept"
        UltraGridColumn7.Header.VisiblePosition = 10
        UltraGridColumn7.Width = 66
        UltraGridColumn14.Header.Caption = "EDI Type"
        UltraGridColumn14.Header.VisiblePosition = 11
        UltraGridColumn14.Width = 70
        UltraGridColumn15.Header.Caption = "Div"
        UltraGridColumn15.Header.VisiblePosition = 21
        UltraGridColumn15.Hidden = True
        UltraGridColumn15.Width = 36
        UltraGridColumn16.Format = "MM/dd"
        UltraGridColumn16.Header.Caption = "Date"
        UltraGridColumn16.Header.VisiblePosition = 13
        UltraGridColumn16.Width = 70
        UltraGridColumn23.Header.Caption = "Ship"
        UltraGridColumn23.Header.VisiblePosition = 14
        UltraGridColumn23.Width = 100
        UltraGridColumn24.Header.Caption = "Cancel"
        UltraGridColumn24.Header.VisiblePosition = 15
        UltraGridColumn24.Width = 100
        UltraGridColumn25.Header.Caption = "Orig Ship"
        UltraGridColumn25.Header.VisiblePosition = 19
        UltraGridColumn25.Width = 110
        UltraGridColumn26.Header.Caption = "Orig Cancel"
        UltraGridColumn26.Header.VisiblePosition = 20
        UltraGridColumn26.Width = 110
        UltraGridColumn27.Header.Caption = "Whse"
        UltraGridColumn27.Header.VisiblePosition = 9
        UltraGridColumn27.Width = 60
        UltraGridColumn28.Header.Caption = "SRep"
        UltraGridColumn28.Header.VisiblePosition = 8
        UltraGridColumn28.Width = 60
        UltraGridColumn29.Header.Caption = "Type"
        UltraGridColumn29.Header.VisiblePosition = 16
        UltraGridColumn29.Width = 57
        UltraGridColumn30.Header.Caption = "Src"
        UltraGridColumn30.Header.VisiblePosition = 18
        UltraGridColumn30.Width = 43
        UltraGridColumn31.Format = "#,##0"
        UltraGridColumn31.Header.Caption = "$Ordr"
        UltraGridColumn31.Header.VisiblePosition = 22
        UltraGridColumn31.Width = 75
        UltraGridColumn32.Format = "#,##0"
        UltraGridColumn32.Header.Caption = "$Open"
        UltraGridColumn32.Header.VisiblePosition = 23
        UltraGridColumn32.Width = 75
        UltraGridColumn33.Format = "#,##0"
        UltraGridColumn33.Header.Caption = "$Pick"
        UltraGridColumn33.Header.VisiblePosition = 24
        UltraGridColumn33.Width = 75
        UltraGridColumn34.Format = "#,##0"
        UltraGridColumn34.Header.Caption = "$Ship"
        UltraGridColumn34.Header.VisiblePosition = 25
        UltraGridColumn34.Width = 75
        UltraGridColumn35.Format = "#,##0"
        UltraGridColumn35.Header.Caption = "$Canc"
        UltraGridColumn35.Header.VisiblePosition = 26
        UltraGridColumn35.Width = 75
        UltraGridColumn36.Header.Caption = "#Ordr"
        UltraGridColumn36.Header.VisiblePosition = 30
        UltraGridColumn36.Width = 65
        UltraGridColumn37.Header.Caption = "#Open"
        UltraGridColumn37.Header.VisiblePosition = 31
        UltraGridColumn37.Width = 65
        UltraGridColumn38.Header.Caption = "#Pick"
        UltraGridColumn38.Header.VisiblePosition = 32
        UltraGridColumn38.Width = 65
        UltraGridColumn39.Header.Caption = "#Ship"
        UltraGridColumn39.Header.VisiblePosition = 33
        UltraGridColumn39.Width = 65
        UltraGridColumn40.Header.Caption = "#Canc"
        UltraGridColumn40.Header.VisiblePosition = 35
        UltraGridColumn40.Width = 65
        UltraGridColumn41.Header.Caption = "#"
        UltraGridColumn41.Header.VisiblePosition = 27
        UltraGridColumn41.Width = 50
        UltraGridColumn42.Header.Caption = "Open"
        UltraGridColumn42.Header.VisiblePosition = 28
        UltraGridColumn42.Width = 46
        UltraGridColumn43.Header.Caption = "Pick"
        UltraGridColumn43.Header.VisiblePosition = 29
        UltraGridColumn43.Width = 40
        UltraGridColumn46.Header.Caption = "Recd"
        UltraGridColumn46.Header.VisiblePosition = 36
        UltraGridColumn46.Width = 110
        UltraGridColumn47.Header.Caption = "Pri"
        UltraGridColumn47.Header.VisiblePosition = 40
        UltraGridColumn47.Width = 40
        UltraGridColumn48.Header.VisiblePosition = 34
        UltraGridColumn48.Hidden = True
        UltraGridColumn49.Header.VisiblePosition = 37
        UltraGridColumn49.Hidden = True
        UltraGridColumn50.Header.VisiblePosition = 38
        UltraGridColumn50.Hidden = True
        UltraGridColumn51.Header.VisiblePosition = 39
        UltraGridColumn51.Hidden = True
        UltraGridColumn52.Format = "MM/dd/yy"
        UltraGridColumn52.Header.Caption = "Avail 1st"
        UltraGridColumn52.Header.VisiblePosition = 46
        UltraGridColumn52.Width = 80
        UltraGridColumn53.Format = "MM/dd/yy"
        UltraGridColumn53.Header.Caption = "Avail Last"
        UltraGridColumn53.Header.VisiblePosition = 47
        UltraGridColumn53.Width = 80
        UltraGridColumn54.Header.Caption = "Rel"
        UltraGridColumn54.Header.VisiblePosition = 41
        UltraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn54.Width = 40
        UltraGridColumn56.Header.Caption = "Rel By"
        UltraGridColumn56.Header.VisiblePosition = 42
        UltraGridColumn56.Width = 75
        UltraGridColumn57.Header.Caption = "Action Date"
        UltraGridColumn57.Header.VisiblePosition = 44
        UltraGridColumn57.Width = 110
        UltraGridColumn59.Header.Caption = "Action By"
        UltraGridColumn59.Header.VisiblePosition = 45
        UltraGridColumn59.Width = 74
        UltraGridColumn60.Header.Caption = "Multi-PO Grp"
        UltraGridColumn60.Header.VisiblePosition = 43
        UltraGridColumn61.Header.Caption = "Terms"
        UltraGridColumn61.Header.VisiblePosition = 54
        UltraGridColumn61.Width = 60
        UltraGridColumn62.Format = "MM/dd/yy HH:mm"
        UltraGridColumn62.Header.Caption = "Last Change"
        UltraGridColumn62.Header.VisiblePosition = 48
        UltraGridColumn62.Width = 161
        UltraGridColumn63.Header.Caption = "By"
        UltraGridColumn63.Header.VisiblePosition = 49
        UltraGridColumn63.Width = 60
        UltraGridColumn64.Header.Caption = "Ship Instr"
        UltraGridColumn64.Header.VisiblePosition = 52
        UltraGridColumn64.Width = 70
        UltraGridColumn65.Header.Caption = "Message"
        UltraGridColumn65.Header.VisiblePosition = 53
        UltraGridColumn65.Width = 70
        UltraGridColumn66.Header.Caption = "CC Exp"
        UltraGridColumn66.Header.VisiblePosition = 50
        UltraGridColumn66.Width = 60
        UltraGridColumn67.Header.Caption = "Last4"
        UltraGridColumn67.Header.VisiblePosition = 51
        UltraGridColumn67.Width = 60
        UltraGridColumn68.Header.Caption = "Customer Name"
        UltraGridColumn68.Header.VisiblePosition = 2
        UltraGridColumn69.Header.Caption = "Cust City"
        UltraGridColumn69.Header.VisiblePosition = 3
        UltraGridColumn69.Width = 140
        UltraGridColumn70.Header.Caption = "Cust State"
        UltraGridColumn70.Header.VisiblePosition = 4
        UltraGridColumn70.Width = 88
        UltraGridColumn71.Header.Caption = "Cust Country"
        UltraGridColumn71.Header.VisiblePosition = 5
        UltraGridColumn71.Width = 160
        UltraGridColumn72.Header.Caption = "Wave No"
        UltraGridColumn72.Header.VisiblePosition = 55
        UltraGridColumn72.Width = 110
        UltraGridColumn74.Header.Caption = "Load ID"
        UltraGridColumn74.Header.VisiblePosition = 56
        UltraGridColumn75.Format = "#,##0"
        UltraGridColumn75.Header.Caption = "$Cur"
        UltraGridColumn75.Header.VisiblePosition = 57
        UltraGridColumn75.Width = 75
        UltraGridColumn76.Format = "#,##0"
        UltraGridColumn76.Header.Caption = "$Fut"
        UltraGridColumn76.Header.VisiblePosition = 58
        UltraGridColumn76.Width = 75
        UltraGridColumn77.Format = "#,##0"
        UltraGridColumn77.Header.Caption = "$Cxl"
        UltraGridColumn77.Header.VisiblePosition = 59
        UltraGridColumn77.Width = 75
        UltraGridColumn78.Format = "#0.0"
        UltraGridColumn78.Header.Caption = "Cur%"
        UltraGridColumn78.Header.VisiblePosition = 60
        UltraGridColumn78.Width = 55
        UltraGridColumn79.Format = "#0.0"
        UltraGridColumn79.Header.Caption = "Fut%"
        UltraGridColumn79.Header.VisiblePosition = 61
        UltraGridColumn79.Width = 55
        UltraGridColumn80.Format = "#0.0"
        UltraGridColumn80.Header.Caption = "Cxl%"
        UltraGridColumn80.Header.VisiblePosition = 62
        UltraGridColumn80.Width = 55
        UltraGridColumn81.Header.Caption = "EDI PO Type"
        UltraGridColumn81.Header.VisiblePosition = 17
        UltraGridColumn81.Width = 101
        UltraGridColumn90.Header.Caption = "Pack"
        UltraGridColumn90.Header.VisiblePosition = 63
        UltraGridColumn90.Width = 70
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn6, UltraGridColumn7, UltraGridColumn14, UltraGridColumn15, UltraGridColumn16, UltraGridColumn23, UltraGridColumn24, UltraGridColumn25, UltraGridColumn26, UltraGridColumn27, UltraGridColumn28, UltraGridColumn29, UltraGridColumn30, UltraGridColumn31, UltraGridColumn32, UltraGridColumn33, UltraGridColumn34, UltraGridColumn35, UltraGridColumn36, UltraGridColumn37, UltraGridColumn38, UltraGridColumn39, UltraGridColumn40, UltraGridColumn41, UltraGridColumn42, UltraGridColumn43, UltraGridColumn46, UltraGridColumn47, UltraGridColumn48, UltraGridColumn49, UltraGridColumn50, UltraGridColumn51, UltraGridColumn52, UltraGridColumn53, UltraGridColumn54, UltraGridColumn56, UltraGridColumn57, UltraGridColumn59, UltraGridColumn60, UltraGridColumn61, UltraGridColumn62, UltraGridColumn63, UltraGridColumn64, UltraGridColumn65, UltraGridColumn66, UltraGridColumn67, UltraGridColumn68, UltraGridColumn69, UltraGridColumn70, UltraGridColumn71, UltraGridColumn72, UltraGridColumn74, UltraGridColumn75, UltraGridColumn76, UltraGridColumn77, UltraGridColumn78, UltraGridColumn79, UltraGridColumn80, UltraGridColumn81, UltraGridColumn90})
        Me.grdSOTORDR0.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTORDR0.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance3.TextHAlignAsString = "Left"
        Me.grdSOTORDR0.DisplayLayout.CaptionAppearance = Appearance3
        Me.grdSOTORDR0.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance4.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance4.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance4.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDR0.DisplayLayout.GroupByBox.Appearance = Appearance4
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDR0.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance5
        Me.grdSOTORDR0.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = True
        Appearance6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance6.BackColor2 = System.Drawing.SystemColors.Control
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance6.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDR0.DisplayLayout.GroupByBox.PromptAppearance = Appearance6
        Me.grdSOTORDR0.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTORDR0.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTORDR0.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Appearance7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTORDR0.DisplayLayout.Override.ActiveCellAppearance = Appearance7
        Me.grdSOTORDR0.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTORDR0.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDR0.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDR0.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTORDR0.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance8.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTORDR0.DisplayLayout.Override.CardAreaAppearance = Appearance8
        Appearance9.BorderColor = System.Drawing.Color.Silver
        Appearance9.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTORDR0.DisplayLayout.Override.CellAppearance = Appearance9
        Me.grdSOTORDR0.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTORDR0.DisplayLayout.Override.CellPadding = 0
        Appearance10.BackColor = System.Drawing.SystemColors.Control
        Appearance10.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance10.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance10.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDR0.DisplayLayout.Override.GroupByRowAppearance = Appearance10
        Appearance11.TextHAlignAsString = "Left"
        Me.grdSOTORDR0.DisplayLayout.Override.HeaderAppearance = Appearance11
        Me.grdSOTORDR0.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTORDR0.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance12.BackColor = System.Drawing.SystemColors.Window
        Appearance12.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTORDR0.DisplayLayout.Override.RowAppearance = Appearance12
        Me.grdSOTORDR0.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance13.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTORDR0.DisplayLayout.Override.TemplateAddRowAppearance = Appearance13
        Me.grdSOTORDR0.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTORDR0.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTORDR0.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTORDR0.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTORDR0.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTORDR0.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTORDR0.Name = "grdSOTORDR0"
        Me.grdSOTORDR0.Size = New System.Drawing.Size(1184, 191)
        Me.grdSOTORDR0.TabIndex = 14
        Me.grdSOTORDR0.Text = "Open Orders"
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer5.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer5.Name = "SplitContainer5"
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.SplitContainer4)
        '
        'SplitContainer5.Panel2
        '
        Me.SplitContainer5.Panel2.Controls.Add(Me.grdSOTPCKC2)
        Me.SplitContainer5.Size = New System.Drawing.Size(1184, 408)
        Me.SplitContainer5.SplitterDistance = 585
        Me.SplitContainer5.TabIndex = 1
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.grdSOTPCKC1)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.grdSOTPCKC4)
        Me.SplitContainer4.Size = New System.Drawing.Size(585, 408)
        Me.SplitContainer4.SplitterDistance = 194
        Me.SplitContainer4.TabIndex = 0
        '
        'grdSOTPCKC1
        '
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTPCKC1.DisplayLayout.Appearance = Appearance14
        UltraGridColumn8.Header.VisiblePosition = 0
        UltraGridColumn8.Hidden = True
        UltraGridColumn154.Header.Caption = "Config"
        UltraGridColumn154.Header.VisiblePosition = 1
        UltraGridColumn154.Width = 72
        UltraGridColumn155.Header.Caption = "Order No"
        UltraGridColumn155.Header.VisiblePosition = 2
        UltraGridColumn155.Width = 116
        UltraGridColumn113.Header.Caption = "#Units"
        UltraGridColumn113.Header.VisiblePosition = 3
        UltraGridColumn113.Width = 70
        UltraGridColumn156.Header.Caption = "#Styles"
        UltraGridColumn156.Header.VisiblePosition = 4
        UltraGridColumn156.Width = 70
        UltraGridColumn157.Header.Caption = "#Orders"
        UltraGridColumn157.Header.VisiblePosition = 5
        UltraGridColumn157.Width = 70
        UltraGridColumn44.Header.Caption = "#Packs"
        UltraGridColumn44.Header.VisiblePosition = 6
        UltraGridColumn44.Width = 70
        UltraGridColumn73.Header.Caption = "Total Qty"
        UltraGridColumn73.Header.VisiblePosition = 7
        UltraGridColumn73.Width = 73
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn8, UltraGridColumn154, UltraGridColumn155, UltraGridColumn113, UltraGridColumn156, UltraGridColumn157, UltraGridColumn44, UltraGridColumn73})
        Me.grdSOTPCKC1.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.grdSOTPCKC1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance15.TextHAlignAsString = "Left"
        Me.grdSOTPCKC1.DisplayLayout.CaptionAppearance = Appearance15
        Me.grdSOTPCKC1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance16.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance16.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance16.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTPCKC1.DisplayLayout.GroupByBox.Appearance = Appearance16
        Appearance17.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTPCKC1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance17
        Me.grdSOTPCKC1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTPCKC1.DisplayLayout.GroupByBox.Hidden = True
        Appearance18.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance18.BackColor2 = System.Drawing.SystemColors.Control
        Appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance18.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTPCKC1.DisplayLayout.GroupByBox.PromptAppearance = Appearance18
        Me.grdSOTPCKC1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTPCKC1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTPCKC1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance19.BackColor = System.Drawing.SystemColors.Window
        Appearance19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTPCKC1.DisplayLayout.Override.ActiveCellAppearance = Appearance19
        Me.grdSOTPCKC1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTPCKC1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTPCKC1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTPCKC1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTPCKC1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance20.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTPCKC1.DisplayLayout.Override.CardAreaAppearance = Appearance20
        Appearance21.BorderColor = System.Drawing.Color.Silver
        Appearance21.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTPCKC1.DisplayLayout.Override.CellAppearance = Appearance21
        Me.grdSOTPCKC1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTPCKC1.DisplayLayout.Override.CellPadding = 0
        Appearance22.BackColor = System.Drawing.SystemColors.Control
        Appearance22.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance22.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance22.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTPCKC1.DisplayLayout.Override.GroupByRowAppearance = Appearance22
        Appearance23.TextHAlignAsString = "Left"
        Me.grdSOTPCKC1.DisplayLayout.Override.HeaderAppearance = Appearance23
        Me.grdSOTPCKC1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTPCKC1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance24.BackColor = System.Drawing.SystemColors.Window
        Appearance24.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTPCKC1.DisplayLayout.Override.RowAppearance = Appearance24
        Me.grdSOTPCKC1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance25.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTPCKC1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance25
        Me.grdSOTPCKC1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTPCKC1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTPCKC1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTPCKC1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTPCKC1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTPCKC1.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTPCKC1.Name = "grdSOTPCKC1"
        Me.grdSOTPCKC1.Size = New System.Drawing.Size(585, 194)
        Me.grdSOTPCKC1.TabIndex = 15
        Me.grdSOTPCKC1.Text = "Order Configurations"
        '
        'grdSOTPCKC4
        '
        Appearance26.BackColor = System.Drawing.SystemColors.Window
        Appearance26.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTPCKC4.DisplayLayout.Appearance = Appearance26
        UltraGridColumn9.Header.Caption = "Order No"
        UltraGridColumn9.Header.VisiblePosition = 0
        UltraGridColumn9.Width = 114
        UltraGridColumn10.Header.Caption = "Customer PO"
        UltraGridColumn10.Header.VisiblePosition = 1
        UltraGridColumn10.Width = 107
        UltraGridColumn11.Header.Caption = "Group"
        UltraGridColumn11.Header.VisiblePosition = 2
        UltraGridColumn11.Width = 102
        UltraGridColumn12.Header.Caption = "Store"
        UltraGridColumn12.Header.VisiblePosition = 3
        UltraGridColumn12.Width = 81
        UltraGridColumn13.Header.Caption = "DC"
        UltraGridColumn13.Header.VisiblePosition = 4
        UltraGridColumn13.Width = 72
        UltraGridColumn17.Header.VisiblePosition = 5
        UltraGridColumn17.Hidden = True
        UltraGridBand3.Columns.AddRange(New Object() {UltraGridColumn9, UltraGridColumn10, UltraGridColumn11, UltraGridColumn12, UltraGridColumn13, UltraGridColumn17})
        Me.grdSOTPCKC4.DisplayLayout.BandsSerializer.Add(UltraGridBand3)
        Me.grdSOTPCKC4.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance27.TextHAlignAsString = "Left"
        Me.grdSOTPCKC4.DisplayLayout.CaptionAppearance = Appearance27
        Me.grdSOTPCKC4.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance28.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance28.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance28.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTPCKC4.DisplayLayout.GroupByBox.Appearance = Appearance28
        Appearance29.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTPCKC4.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance29
        Me.grdSOTPCKC4.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTPCKC4.DisplayLayout.GroupByBox.Hidden = True
        Appearance30.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance30.BackColor2 = System.Drawing.SystemColors.Control
        Appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance30.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTPCKC4.DisplayLayout.GroupByBox.PromptAppearance = Appearance30
        Me.grdSOTPCKC4.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTPCKC4.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTPCKC4.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance31.BackColor = System.Drawing.SystemColors.Window
        Appearance31.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTPCKC4.DisplayLayout.Override.ActiveCellAppearance = Appearance31
        Me.grdSOTPCKC4.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTPCKC4.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTPCKC4.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTPCKC4.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTPCKC4.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance32.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTPCKC4.DisplayLayout.Override.CardAreaAppearance = Appearance32
        Appearance33.BorderColor = System.Drawing.Color.Silver
        Appearance33.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTPCKC4.DisplayLayout.Override.CellAppearance = Appearance33
        Me.grdSOTPCKC4.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTPCKC4.DisplayLayout.Override.CellPadding = 0
        Appearance34.BackColor = System.Drawing.SystemColors.Control
        Appearance34.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance34.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance34.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTPCKC4.DisplayLayout.Override.GroupByRowAppearance = Appearance34
        Appearance35.TextHAlignAsString = "Left"
        Me.grdSOTPCKC4.DisplayLayout.Override.HeaderAppearance = Appearance35
        Me.grdSOTPCKC4.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTPCKC4.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance36.BackColor = System.Drawing.SystemColors.Window
        Appearance36.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTPCKC4.DisplayLayout.Override.RowAppearance = Appearance36
        Me.grdSOTPCKC4.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance37.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTPCKC4.DisplayLayout.Override.TemplateAddRowAppearance = Appearance37
        Me.grdSOTPCKC4.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTPCKC4.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTPCKC4.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTPCKC4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTPCKC4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTPCKC4.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTPCKC4.Name = "grdSOTPCKC4"
        Me.grdSOTPCKC4.Size = New System.Drawing.Size(585, 210)
        Me.grdSOTPCKC4.TabIndex = 16
        Me.grdSOTPCKC4.Text = "Orders with Configuration"
        '
        'grdSOTPCKC2
        '
        Appearance38.BackColor = System.Drawing.SystemColors.Window
        Appearance38.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTPCKC2.DisplayLayout.Appearance = Appearance38
        UltraGridColumn18.Header.VisiblePosition = 0
        UltraGridColumn18.Hidden = True
        UltraGridColumn5.Header.VisiblePosition = 4
        UltraGridColumn5.Hidden = True
        UltraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn19.Header.Caption = "Ln"
        UltraGridColumn19.Header.VisiblePosition = 1
        UltraGridColumn19.Width = 36
        UltraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn20.Header.Caption = "Style"
        UltraGridColumn20.Header.VisiblePosition = 2
        UltraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn21.Header.Caption = "Color"
        UltraGridColumn21.Header.VisiblePosition = 3
        UltraGridColumn21.Width = 52
        UltraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn22.Header.Caption = "Open"
        UltraGridColumn22.Header.VisiblePosition = 5
        UltraGridColumn22.Width = 55
        UltraGridColumn45.Header.Caption = "1"
        UltraGridColumn45.Header.VisiblePosition = 6
        UltraGridColumn45.Width = 55
        UltraGridColumn55.Header.Caption = "2"
        UltraGridColumn55.Header.VisiblePosition = 7
        UltraGridColumn55.Width = 55
        UltraGridColumn58.Header.Caption = "3"
        UltraGridColumn58.Header.VisiblePosition = 8
        UltraGridColumn58.Width = 55
        UltraGridColumn82.Header.Caption = "4"
        UltraGridColumn82.Header.VisiblePosition = 9
        UltraGridColumn82.Width = 55
        UltraGridColumn83.Header.Caption = "5"
        UltraGridColumn83.Header.VisiblePosition = 10
        UltraGridColumn83.Width = 55
        UltraGridColumn84.Header.Caption = "6"
        UltraGridColumn84.Header.VisiblePosition = 11
        UltraGridColumn84.Width = 55
        UltraGridColumn85.Header.Caption = "7"
        UltraGridColumn85.Header.VisiblePosition = 12
        UltraGridColumn85.Width = 55
        UltraGridColumn86.Header.Caption = "8"
        UltraGridColumn86.Header.VisiblePosition = 13
        UltraGridColumn86.Width = 55
        UltraGridColumn87.Header.Caption = "9"
        UltraGridColumn87.Header.VisiblePosition = 14
        UltraGridColumn87.Width = 55
        UltraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn88.Header.Caption = "Packed"
        UltraGridColumn88.Header.VisiblePosition = 15
        UltraGridColumn88.Width = 60
        UltraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn89.Header.Caption = "NotPck"
        UltraGridColumn89.Header.VisiblePosition = 16
        UltraGridColumn89.Width = 60
        UltraGridBand4.Columns.AddRange(New Object() {UltraGridColumn18, UltraGridColumn5, UltraGridColumn19, UltraGridColumn20, UltraGridColumn21, UltraGridColumn22, UltraGridColumn45, UltraGridColumn55, UltraGridColumn58, UltraGridColumn82, UltraGridColumn83, UltraGridColumn84, UltraGridColumn85, UltraGridColumn86, UltraGridColumn87, UltraGridColumn88, UltraGridColumn89})
        Me.grdSOTPCKC2.DisplayLayout.BandsSerializer.Add(UltraGridBand4)
        Me.grdSOTPCKC2.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance39.TextHAlignAsString = "Left"
        Me.grdSOTPCKC2.DisplayLayout.CaptionAppearance = Appearance39
        Me.grdSOTPCKC2.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance40.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance40.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance40.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTPCKC2.DisplayLayout.GroupByBox.Appearance = Appearance40
        Appearance41.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTPCKC2.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance41
        Me.grdSOTPCKC2.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTPCKC2.DisplayLayout.GroupByBox.Hidden = True
        Appearance42.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance42.BackColor2 = System.Drawing.SystemColors.Control
        Appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance42.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTPCKC2.DisplayLayout.GroupByBox.PromptAppearance = Appearance42
        Me.grdSOTPCKC2.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTPCKC2.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTPCKC2.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance43.BackColor = System.Drawing.SystemColors.Window
        Appearance43.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTPCKC2.DisplayLayout.Override.ActiveCellAppearance = Appearance43
        Me.grdSOTPCKC2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTPCKC2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTPCKC2.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTPCKC2.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance44.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTPCKC2.DisplayLayout.Override.CardAreaAppearance = Appearance44
        Appearance45.BorderColor = System.Drawing.Color.Silver
        Appearance45.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTPCKC2.DisplayLayout.Override.CellAppearance = Appearance45
        Me.grdSOTPCKC2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSOTPCKC2.DisplayLayout.Override.CellPadding = 0
        Appearance46.BackColor = System.Drawing.SystemColors.Control
        Appearance46.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance46.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance46.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTPCKC2.DisplayLayout.Override.GroupByRowAppearance = Appearance46
        Appearance47.TextHAlignAsString = "Left"
        Me.grdSOTPCKC2.DisplayLayout.Override.HeaderAppearance = Appearance47
        Me.grdSOTPCKC2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTPCKC2.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance48.BackColor = System.Drawing.SystemColors.Window
        Appearance48.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTPCKC2.DisplayLayout.Override.RowAppearance = Appearance48
        Me.grdSOTPCKC2.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance49.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTPCKC2.DisplayLayout.Override.TemplateAddRowAppearance = Appearance49
        Me.grdSOTPCKC2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTPCKC2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTPCKC2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTPCKC2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTPCKC2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTPCKC2.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTPCKC2.Name = "grdSOTPCKC2"
        Me.grdSOTPCKC2.Size = New System.Drawing.Size(595, 408)
        Me.grdSOTPCKC2.TabIndex = 17
        Me.grdSOTPCKC2.Text = "Style/Color/Qty Details for Configuration"
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Panel2Collapsed = True
        Me.SplitContainer1.Size = New System.Drawing.Size(1184, 661)
        Me.SplitContainer1.SplitterDistance = 605
        Me.SplitContainer1.TabIndex = 16
        '
        'SOFCARTP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 661)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "SOFCARTP"
        Me.Text = "Pack Configuration"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox2.ResumeLayout(False)
        Me.UltraGroupBox2.PerformLayout()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        CType(Me.grdSOTORDR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer5.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        CType(Me.grdSOTPCKC1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTPCKC4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTPCKC2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

End Sub
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraTextEditor6 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cmdUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grdSOTORDR0 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdSOTPCKC1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grdSOTPCKC4 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraLabel14 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor3 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents SplitContainer5 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdSOTPCKC2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmdDelete As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
End Class
