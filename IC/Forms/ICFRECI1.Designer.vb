<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ICFRECI1
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
        Dim UltraExplorerBarItem1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim UltraExplorerBarItem2 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim UltraExplorerBarGroup2 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup()
        Dim Appearance61 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance62 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance63 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance64 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance65 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance66 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance67 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance68 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance69 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIP3", -1)
        Dim UltraGridColumn230 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn231 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_LNO")
        Dim UltraGridColumn232 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_NO")
        Dim UltraGridColumn233 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_LNO")
        Dim UltraGridColumn234 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SHP")
        Dim UltraGridColumn235 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_REC")
        Dim UltraGridColumn236 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CLOSE_PO")
        Dim UltraGridColumn237 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn238 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn239 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn240 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn241 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST")
        Dim UltraGridColumn242 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DUTY_RATE_CODE")
        Dim UltraGridColumn243 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DUTY_RATE")
        Dim UltraGridColumn244 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_FREIGHT_IN")
        Dim UltraGridColumn245 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_TRUCKING")
        Dim UltraGridColumn246 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_DUTY")
        Dim UltraGridColumn247 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_CUSTOMS")
        Dim UltraGridColumn248 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WEIGHT_FACTOR")
        Dim UltraGridColumn249 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_LANDED")
        Dim UltraGridColumn250 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_VCOST")
        Dim UltraGridColumn251 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_MATLS")
        Dim UltraGridColumn252 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FOB_CMT")
        Dim UltraGridColumn253 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_VCOST_UM")
        Dim UltraGridColumn254 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_MATLS_UM")
        Dim UltraGridColumn255 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_VCOST_DZ")
        Dim UltraGridColumn256 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_MATLS_DZ")
        Dim UltraGridColumn257 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_OTHER")
        Dim UltraGridColumn258 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_COMM")
        Dim UltraGridColumn259 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_OTHER_DZ")
        Dim UltraGridColumn260 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_CHANGED")
        Dim UltraGridColumn261 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_BUFFER")
        Dim UltraGridColumn262 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA")
        Dim UltraGridColumn263 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA_DZ")
        Dim UltraGridColumn264 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA_DF")
        Dim UltraGridColumn265 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA_DF_DZ")
        Dim UltraGridColumn208 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VEND_CODE")
        Dim UltraGridColumn266 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE")
        Dim UltraGridColumn267 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn268 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_OPN")
        Dim UltraGridColumn269 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_UOM")
        Dim UltraGridColumn270 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR2_COST")
        Dim UltraGridColumn271 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_DESC")
        Dim UltraGridColumn272 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SUB_BODY_CODE")
        Dim UltraGridColumn273 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SUB_UNIT_PACK_QTY")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_PACK_QTY")
        Dim UltraGridColumn274 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_REFERENCE")
        Dim UltraGridColumn95 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_SHIP_BY")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_REC_OLD")
        Dim UltraGridColumn119 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_VAR")
        Dim UltraGridColumn275 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_OPN_PRE")
        Dim UltraGridColumn276 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SHP_DZ")
        Dim UltraGridColumn277 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_REC_DZ")
        Dim UltraGridColumn278 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_STATUS")
        Dim UltraGridColumn279 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SR")
        Dim UltraGridColumn280 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SR_DZ")
        Dim UltraGridColumn281 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL_DUTY")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CONTAINER_NO")
        Dim UltraGridColumn282 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EXT_WEIGHT_FACTOR")
        Dim UltraGridColumn283 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NET_OPEN")
        Dim UltraGridColumn284 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("NET_OPEN_DZ")
        Dim UltraGridColumn15 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_AMT_REC")
        Dim UltraGridColumn285 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FIRST_COST_TOTAL")
        Dim UltraGridColumn286 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FIRST_COST_TOTAL_DZ", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn287 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COMMISSION_COST")
        Dim UltraGridColumn288 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COMMISSION_COST_DZ")
        Dim UltraGridColumn16 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LINE_EXACT")
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LINE_OVER")
        Dim UltraGridColumn18 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LINE_SHORT")
        Dim UltraGridColumn19 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LINE_ZERO")
        Dim UltraGridColumn162 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CBM")
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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ICTRECI0", -1)
        Dim UltraGridColumn33 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE")
        Dim UltraGridColumn34 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn35 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CLASS_CODE")
        Dim UltraGridColumn36 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_DESC")
        Dim UltraGridColumn37 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("BOMQTY")
        Dim UltraGridColumn38 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("BOMCST")
        Dim UltraGridColumn39 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RECQTY")
        Dim UltraGridColumn40 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RECCST")
        Dim UltraGridColumn41 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHPQTY")
        Dim UltraGridColumn42 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHPCST")
        Dim UltraGridColumn43 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ADJQTY")
        Dim UltraGridColumn44 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ADJCST")
        Dim UltraGridColumn45 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EOMQTY")
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EOMCST")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OOBQTY")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OOBCST")
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
        Dim UltraTab8 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand3 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ICTRECIG", -1)
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("JOURNAL_TYPE")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GL_AMT")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GL_AMT_C")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GL_AMT_R")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IC_AMT")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IC_AMT_C")
        Dim UltraGridColumn55 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IC_AMT_R")
        Dim UltraGridColumn56 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DIFF")
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
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
        Dim UltraGridBand4 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIPX", -1)
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_VESSEL")
        Dim UltraGridColumn20 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_RECEIVED")
        Dim UltraGridColumn21 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_ETA")
        Dim UltraGridColumn22 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY")
        Dim UltraGridColumn23 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("AMT")
        Dim UltraGridColumn86 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CHECK_DATE", 0)
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
        Dim Appearance48 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance49 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand5 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIPH", -1)
        Dim UltraGridColumn24 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OPS_YYYYPP")
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn25 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_LNO")
        Dim UltraGridColumn26 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_NO", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn27 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_LNO")
        Dim UltraGridColumn28 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SHP")
        Dim UltraGridColumn29 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_REC")
        Dim UltraGridColumn30 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CLOSE_PO")
        Dim UltraGridColumn31 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn32 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn52 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn53 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn54 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST")
        Dim UltraGridColumn57 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DUTY_RATE_CODE")
        Dim UltraGridColumn58 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DUTY_RATE")
        Dim UltraGridColumn59 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_FREIGHT_IN")
        Dim UltraGridColumn60 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_TRUCKING")
        Dim UltraGridColumn61 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_DUTY")
        Dim UltraGridColumn62 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_CUSTOMS")
        Dim UltraGridColumn63 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WEIGHT_FACTOR")
        Dim UltraGridColumn64 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_LANDED")
        Dim UltraGridColumn65 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_VCOST")
        Dim UltraGridColumn66 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_MATLS")
        Dim UltraGridColumn67 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FOB_CMT")
        Dim UltraGridColumn68 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_VCOST_UM")
        Dim UltraGridColumn69 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_MATLS_UM")
        Dim UltraGridColumn70 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_VCOST_DZ")
        Dim UltraGridColumn71 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_MATLS_DZ")
        Dim UltraGridColumn72 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_OTHER")
        Dim UltraGridColumn73 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_COMM")
        Dim UltraGridColumn74 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_OTHER_DZ")
        Dim UltraGridColumn75 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_CHANGED")
        Dim UltraGridColumn76 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_BUFFER")
        Dim UltraGridColumn77 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA")
        Dim UltraGridColumn78 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA_DZ")
        Dim UltraGridColumn79 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA_DF")
        Dim UltraGridColumn80 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_QUOTA_DF_DZ")
        Dim UltraGridColumn81 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COST_MISC")
        Dim UltraGridColumn82 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SHP_EXT")
        Dim UltraGridColumn47 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACCRUAL_STATUS", 0)
        Dim UltraGridColumn48 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VOUCHER_NO", 1)
        Dim UltraGridColumn49 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VEND_CODE", 2)
        Dim UltraGridColumn50 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VEND_NAME", 3)
        Dim UltraGridColumn51 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_SHIPPED", 4)
        Dim UltraGridColumn83 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_NUM", 5)
        Dim UltraGridColumn84 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_DATE", 6)
        Dim UltraGridColumn85 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CHECK_DATE", 7)
        Dim Appearance50 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance51 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance52 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance53 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance54 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance55 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance56 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance57 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance58 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance59 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance60 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraTab4 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab5 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab11 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab1 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab6 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab7 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim Appearance70 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance71 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance72 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance73 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance74 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance75 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance76 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance77 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance78 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance79 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance80 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance81 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.UltraExplorerBarContainerControl1 = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarContainerControl()
        Me.UltraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.chkOOB = New ABSCS.ABSCheckBox()
        Me.UltraTabPageControl8 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdPOTSHIP3 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl1 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grdICTRECI0 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabControl3 = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage5 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.UltraTabPageControl2 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdICTRECIG = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl3 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdPOTSHIPX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl4 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdPOTSHIPH = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl6 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.UltraTabPageControl7 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.tabMain = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage1 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.UltraTabPageControl5 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.cbeOPS_YYYYPP = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.tab = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage2 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.UltraCombo1 = New Infragistics.Win.UltraWinGrid.UltraCombo()
        Me.spl = New System.Windows.Forms.SplitContainer()
        Me.UltraTabControl1 = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage3 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.tabBOL = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage7 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.UltraTabControl2 = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage4 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraExplorerBar1.SuspendLayout()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraExplorerBarContainerControl1.SuspendLayout()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox2.SuspendLayout()
        CType(Me.chkOOB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl8.SuspendLayout()
        CType(Me.grdPOTSHIP3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grdICTRECI0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTabControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabControl3.SuspendLayout()
        Me.UltraTabPageControl2.SuspendLayout()
        CType(Me.grdICTRECIG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl3.SuspendLayout()
        CType(Me.grdPOTSHIPX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl4.SuspendLayout()
        CType(Me.grdPOTSHIPH, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl7.SuspendLayout()
        CType(Me.tabMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabMain.SuspendLayout()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.cbeOPS_YYYYPP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tab, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab.SuspendLayout()
        CType(Me.UltraCombo1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.spl.Panel1.SuspendLayout()
        Me.spl.Panel2.SuspendLayout()
        Me.spl.SuspendLayout()
        CType(Me.UltraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tabBOL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTabControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.Controls.Add(Me.UltraExplorerBarContainerControl1)
        UltraExplorerBarItem1.Key = "View"
        UltraExplorerBarItem1.Text = "View"
        UltraExplorerBarItem2.Key = "Done"
        UltraExplorerBarItem2.Text = "Done"
        UltraExplorerBarGroup1.Items.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem() {UltraExplorerBarItem1, UltraExplorerBarItem2})
        UltraExplorerBarGroup1.Key = "Screen Control"
        UltraExplorerBarGroup1.Text = "Screen Control"
        UltraExplorerBarGroup2.Container = Me.UltraExplorerBarContainerControl1
        UltraExplorerBarGroup2.Settings.ContainerHeight = 35
        UltraExplorerBarGroup2.Settings.Style = Infragistics.Win.UltraWinExplorerBar.GroupStyle.ControlContainer
        UltraExplorerBarGroup2.Text = "Options"
        Me.UltraExplorerBar1.Groups.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup() {UltraExplorerBarGroup1, UltraExplorerBarGroup2})
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Location = New System.Drawing.Point(0, 24)
        Me.UltraExplorerBar1.Margins.Bottom = 0
        Me.UltraExplorerBar1.Margins.Left = 0
        Me.UltraExplorerBar1.Margins.Right = 0
        Me.UltraExplorerBar1.Margins.Top = 0
        Me.UltraExplorerBar1.ShowDefaultContextMenu = False
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 669)
        Me.UltraExplorerBar1.Tag = "CLICK"
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.spl)
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(1027, 693)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.grdASFBASEX, 0)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.spl, 0)
        '
        'grdASFBASEX
        '
        Appearance61.BackColor = System.Drawing.SystemColors.Window
        Appearance61.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance61
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance62.BackColor = System.Drawing.SystemColors.Window
        Appearance62.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance62
        Appearance63.BackColor = System.Drawing.SystemColors.Highlight
        Appearance63.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance63
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance64.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance64
        Appearance65.BorderColor = System.Drawing.Color.Silver
        Appearance65.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance65
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance66.BackColor = System.Drawing.SystemColors.Control
        Appearance66.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance66.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance66.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance66
        Appearance67.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance67
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance68.BackColor = System.Drawing.SystemColors.Window
        Appearance68.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance68
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance69.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance69
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraExplorerBarContainerControl1
        '
        Me.UltraExplorerBarContainerControl1.Controls.Add(Me.UltraGroupBox2)
        Me.UltraExplorerBarContainerControl1.Location = New System.Drawing.Point(13, 144)
        Me.UltraExplorerBarContainerControl1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraExplorerBarContainerControl1.Name = "UltraExplorerBarContainerControl1"
        Me.UltraExplorerBarContainerControl1.Size = New System.Drawing.Size(189, 35)
        Me.UltraExplorerBarContainerControl1.TabIndex = 0
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.Controls.Add(Me.chkOOB)
        Me.UltraGroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(189, 35)
        Me.UltraGroupBox2.TabIndex = 0
        '
        'chkOOB
        '
        Me.chkOOB.Location = New System.Drawing.Point(29, 4)
        Me.chkOOB.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.chkOOB.Name = "chkOOB"
        Me.chkOOB.Size = New System.Drawing.Size(150, 22)
        Me.chkOOB.TabIndex = 0
        Me.chkOOB.Text = "OOB Only"
        '
        'UltraTabPageControl8
        '
        Me.UltraTabPageControl8.Controls.Add(Me.grdPOTSHIP3)
        Me.UltraTabPageControl8.Location = New System.Drawing.Point(1, 29)
        Me.UltraTabPageControl8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabPageControl8.Name = "UltraTabPageControl8"
        Me.UltraTabPageControl8.Size = New System.Drawing.Size(1015, 113)
        '
        'grdPOTSHIP3
        '
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Appearance13.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIP3.DisplayLayout.Appearance = Appearance13
        UltraGridColumn230.Header.VisiblePosition = 0
        UltraGridColumn230.Hidden = True
        UltraGridColumn231.Header.VisiblePosition = 1
        UltraGridColumn231.Hidden = True
        UltraGridColumn232.Header.Caption = "PO No"
        UltraGridColumn232.Header.VisiblePosition = 2
        UltraGridColumn232.Width = 69
        UltraGridColumn233.Header.Caption = "Ln"
        UltraGridColumn233.Header.VisiblePosition = 3
        UltraGridColumn233.Width = 29
        UltraGridColumn234.Header.Caption = "Ship"
        UltraGridColumn234.Header.VisiblePosition = 11
        UltraGridColumn234.Width = 70
        UltraGridColumn235.Header.Caption = "Recd"
        UltraGridColumn235.Header.VisiblePosition = 12
        UltraGridColumn235.Width = 70
        UltraGridColumn236.Header.Caption = "Close PO"
        UltraGridColumn236.Header.VisiblePosition = 21
        UltraGridColumn236.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn236.Width = 75
        UltraGridColumn237.Header.VisiblePosition = 22
        UltraGridColumn237.Hidden = True
        UltraGridColumn238.Header.VisiblePosition = 23
        UltraGridColumn238.Hidden = True
        UltraGridColumn239.Header.VisiblePosition = 24
        UltraGridColumn239.Hidden = True
        UltraGridColumn240.Header.VisiblePosition = 25
        UltraGridColumn240.Hidden = True
        UltraGridColumn241.Format = "#.000000"
        UltraGridColumn241.Header.Caption = "Cost"
        UltraGridColumn241.Header.VisiblePosition = 20
        UltraGridColumn241.Width = 94
        UltraGridColumn242.Header.Caption = "Duty Code"
        UltraGridColumn242.Header.VisiblePosition = 45
        UltraGridColumn242.Width = 83
        UltraGridColumn243.Header.Caption = "Duty Rate"
        UltraGridColumn243.Header.VisiblePosition = 46
        UltraGridColumn243.Width = 86
        UltraGridColumn244.Header.Caption = "Freight"
        UltraGridColumn244.Header.VisiblePosition = 47
        UltraGridColumn244.Width = 74
        UltraGridColumn245.Header.Caption = "Truck"
        UltraGridColumn245.Header.VisiblePosition = 50
        UltraGridColumn245.Width = 79
        UltraGridColumn246.Header.Caption = "Duty"
        UltraGridColumn246.Header.VisiblePosition = 48
        UltraGridColumn246.Width = 82
        UltraGridColumn247.Header.Caption = "Customs"
        UltraGridColumn247.Header.VisiblePosition = 49
        UltraGridColumn247.Width = 84
        UltraGridColumn248.Header.Caption = "WgtFct"
        UltraGridColumn248.Header.VisiblePosition = 43
        UltraGridColumn248.Width = 61
        UltraGridColumn249.Header.Caption = "Landed"
        UltraGridColumn249.Header.VisiblePosition = 51
        UltraGridColumn249.Width = 87
        UltraGridColumn250.Header.Caption = "VCost"
        UltraGridColumn250.Header.VisiblePosition = 52
        UltraGridColumn250.Hidden = True
        UltraGridColumn251.Header.Caption = "Matls"
        UltraGridColumn251.Header.VisiblePosition = 53
        UltraGridColumn251.Hidden = True
        UltraGridColumn252.Header.Caption = "FOB CMT"
        UltraGridColumn252.Header.VisiblePosition = 55
        UltraGridColumn252.Hidden = True
        UltraGridColumn252.Width = 87
        UltraGridColumn253.Header.Caption = "Vend/UM"
        UltraGridColumn253.Header.VisiblePosition = 27
        UltraGridColumn253.Width = 84
        UltraGridColumn254.Header.Caption = "Matls/UM"
        UltraGridColumn254.Header.VisiblePosition = 29
        UltraGridColumn254.Width = 76
        UltraGridColumn255.Header.Caption = "Vend/Dz"
        UltraGridColumn255.Header.VisiblePosition = 26
        UltraGridColumn255.Width = 85
        UltraGridColumn256.Header.Caption = "Matls/Dz"
        UltraGridColumn256.Header.VisiblePosition = 28
        UltraGridColumn256.Width = 77
        UltraGridColumn257.Header.Caption = "Other/UM"
        UltraGridColumn257.Header.VisiblePosition = 31
        UltraGridColumn257.Width = 77
        UltraGridColumn258.Header.Caption = "Comm%"
        UltraGridColumn258.Header.VisiblePosition = 38
        UltraGridColumn258.Width = 75
        UltraGridColumn259.Header.Caption = "Other/Dz"
        UltraGridColumn259.Header.VisiblePosition = 30
        UltraGridColumn259.Width = 82
        UltraGridColumn260.Header.Caption = "Cst Chg"
        UltraGridColumn260.Header.VisiblePosition = 54
        UltraGridColumn260.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn260.Width = 60
        UltraGridColumn261.Header.Caption = "Cost Inc%"
        UltraGridColumn261.Header.VisiblePosition = 39
        UltraGridColumn261.Width = 84
        UltraGridColumn262.Header.Caption = "Quota"
        UltraGridColumn262.Header.VisiblePosition = 35
        UltraGridColumn262.Width = 62
        UltraGridColumn263.Header.Caption = "Quota/Dz"
        UltraGridColumn263.Header.VisiblePosition = 34
        UltraGridColumn263.Width = 80
        UltraGridColumn264.Header.Caption = "Quota DF"
        UltraGridColumn264.Header.VisiblePosition = 37
        UltraGridColumn264.Width = 76
        UltraGridColumn265.Header.Caption = "Quota DF Dz"
        UltraGridColumn265.Header.VisiblePosition = 36
        UltraGridColumn265.Width = 97
        UltraGridColumn208.Header.Caption = "Supplier"
        UltraGridColumn208.Header.VisiblePosition = 4
        UltraGridColumn208.Width = 71
        UltraGridColumn266.Header.Caption = "Style"
        UltraGridColumn266.Header.VisiblePosition = 7
        UltraGridColumn266.Hidden = True
        UltraGridColumn266.Width = 109
        UltraGridColumn267.Header.Caption = "Color"
        UltraGridColumn267.Header.VisiblePosition = 9
        UltraGridColumn267.Hidden = True
        UltraGridColumn267.Width = 54
        UltraGridColumn268.Header.VisiblePosition = 57
        UltraGridColumn268.Hidden = True
        UltraGridColumn269.Header.Caption = "UM"
        UltraGridColumn269.Header.VisiblePosition = 17
        UltraGridColumn269.Width = 38
        UltraGridColumn270.Header.VisiblePosition = 58
        UltraGridColumn270.Hidden = True
        UltraGridColumn271.Header.Caption = "Description"
        UltraGridColumn271.Header.VisiblePosition = 8
        UltraGridColumn271.Width = 203
        UltraGridColumn272.Header.Caption = "SBC"
        UltraGridColumn272.Header.VisiblePosition = 42
        UltraGridColumn272.Width = 66
        UltraGridColumn273.Header.Caption = "Pcs/u"
        UltraGridColumn273.Header.VisiblePosition = 18
        UltraGridColumn273.Width = 51
        UltraGridColumn2.Header.Caption = "u/Cs"
        UltraGridColumn2.Header.VisiblePosition = 19
        UltraGridColumn2.Width = 47
        UltraGridColumn274.Header.Caption = "PO Reference"
        UltraGridColumn274.Header.VisiblePosition = 5
        UltraGridColumn274.Width = 109
        UltraGridColumn95.Header.Caption = "Ship By"
        UltraGridColumn95.Header.VisiblePosition = 6
        UltraGridColumn95.Width = 104
        UltraGridColumn14.Header.VisiblePosition = 56
        UltraGridColumn14.Hidden = True
        UltraGridColumn119.Header.Caption = "Var"
        UltraGridColumn119.Header.VisiblePosition = 13
        UltraGridColumn119.Width = 70
        UltraGridColumn275.Header.VisiblePosition = 59
        UltraGridColumn275.Hidden = True
        UltraGridColumn276.Header.Caption = "Dz Ship"
        UltraGridColumn276.Header.VisiblePosition = 15
        UltraGridColumn276.Width = 64
        UltraGridColumn277.Header.Caption = "Dz Recd"
        UltraGridColumn277.Header.VisiblePosition = 16
        UltraGridColumn277.Width = 62
        UltraGridColumn278.Header.VisiblePosition = 60
        UltraGridColumn278.Hidden = True
        UltraGridColumn279.Header.VisiblePosition = 62
        UltraGridColumn279.Hidden = True
        UltraGridColumn280.Header.VisiblePosition = 63
        UltraGridColumn280.Hidden = True
        UltraGridColumn281.Header.VisiblePosition = 64
        UltraGridColumn281.Hidden = True
        UltraGridColumn9.Header.VisiblePosition = 61
        UltraGridColumn9.Hidden = True
        UltraGridColumn282.Header.VisiblePosition = 66
        UltraGridColumn282.Hidden = True
        UltraGridColumn283.Format = "#,##0"
        UltraGridColumn283.Header.Caption = "Open"
        UltraGridColumn283.Header.VisiblePosition = 10
        UltraGridColumn283.Width = 70
        UltraGridColumn284.Format = "#,##0"
        UltraGridColumn284.Header.Caption = "Dz Open"
        UltraGridColumn284.Header.VisiblePosition = 14
        UltraGridColumn284.Width = 65
        UltraGridColumn15.Header.VisiblePosition = 65
        UltraGridColumn15.Hidden = True
        UltraGridColumn285.Header.Caption = "First/UM"
        UltraGridColumn285.Header.VisiblePosition = 33
        UltraGridColumn285.Width = 70
        UltraGridColumn286.Header.Caption = "First/Dz"
        UltraGridColumn286.Header.VisiblePosition = 32
        UltraGridColumn286.Width = 82
        UltraGridColumn287.Header.Caption = "Comm/Unit"
        UltraGridColumn287.Header.VisiblePosition = 41
        UltraGridColumn287.Width = 94
        UltraGridColumn288.Header.Caption = "Comm/Dz"
        UltraGridColumn288.Header.VisiblePosition = 40
        UltraGridColumn288.Width = 76
        UltraGridColumn16.Header.VisiblePosition = 67
        UltraGridColumn16.Hidden = True
        UltraGridColumn17.Header.VisiblePosition = 68
        UltraGridColumn17.Hidden = True
        UltraGridColumn18.Header.VisiblePosition = 69
        UltraGridColumn18.Hidden = True
        UltraGridColumn19.Header.VisiblePosition = 70
        UltraGridColumn19.Hidden = True
        UltraGridColumn162.Header.VisiblePosition = 44
        UltraGridColumn162.Width = 70
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn230, UltraGridColumn231, UltraGridColumn232, UltraGridColumn233, UltraGridColumn234, UltraGridColumn235, UltraGridColumn236, UltraGridColumn237, UltraGridColumn238, UltraGridColumn239, UltraGridColumn240, UltraGridColumn241, UltraGridColumn242, UltraGridColumn243, UltraGridColumn244, UltraGridColumn245, UltraGridColumn246, UltraGridColumn247, UltraGridColumn248, UltraGridColumn249, UltraGridColumn250, UltraGridColumn251, UltraGridColumn252, UltraGridColumn253, UltraGridColumn254, UltraGridColumn255, UltraGridColumn256, UltraGridColumn257, UltraGridColumn258, UltraGridColumn259, UltraGridColumn260, UltraGridColumn261, UltraGridColumn262, UltraGridColumn263, UltraGridColumn264, UltraGridColumn265, UltraGridColumn208, UltraGridColumn266, UltraGridColumn267, UltraGridColumn268, UltraGridColumn269, UltraGridColumn270, UltraGridColumn271, UltraGridColumn272, UltraGridColumn273, UltraGridColumn2, UltraGridColumn274, UltraGridColumn95, UltraGridColumn14, UltraGridColumn119, UltraGridColumn275, UltraGridColumn276, UltraGridColumn277, UltraGridColumn278, UltraGridColumn279, UltraGridColumn280, UltraGridColumn281, UltraGridColumn9, UltraGridColumn282, UltraGridColumn283, UltraGridColumn284, UltraGridColumn15, UltraGridColumn285, UltraGridColumn286, UltraGridColumn287, UltraGridColumn288, UltraGridColumn16, UltraGridColumn17, UltraGridColumn18, UltraGridColumn19, UltraGridColumn162})
        Me.grdPOTSHIP3.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.grdPOTSHIP3.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance14.TextHAlignAsString = "Left"
        Me.grdPOTSHIP3.DisplayLayout.CaptionAppearance = Appearance14
        Appearance15.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance15.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance15.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.Appearance = Appearance15
        Appearance16.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance16
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.Hidden = True
        Appearance17.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance17.BackColor2 = System.Drawing.SystemColors.Control
        Appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance17.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.PromptAppearance = Appearance17
        Me.grdPOTSHIP3.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIP3.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdPOTSHIP3.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance18.BackColor = System.Drawing.SystemColors.Window
        Appearance18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIP3.DisplayLayout.Override.ActiveCellAppearance = Appearance18
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP3.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIP3.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance19.BackColor = System.Drawing.Color.Transparent
        Me.grdPOTSHIP3.DisplayLayout.Override.CardAreaAppearance = Appearance19
        Appearance20.BorderColor = System.Drawing.Color.Silver
        Appearance20.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIP3.DisplayLayout.Override.CellAppearance = Appearance20
        Me.grdPOTSHIP3.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
        Me.grdPOTSHIP3.DisplayLayout.Override.CellPadding = 0
        Appearance21.BackColor = System.Drawing.SystemColors.Control
        Appearance21.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance21.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance21.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP3.DisplayLayout.Override.GroupByRowAppearance = Appearance21
        Appearance22.TextHAlignAsString = "Left"
        Me.grdPOTSHIP3.DisplayLayout.Override.HeaderAppearance = Appearance22
        Me.grdPOTSHIP3.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIP3.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance23.BackColor = System.Drawing.SystemColors.Window
        Appearance23.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIP3.DisplayLayout.Override.RowAppearance = Appearance23
        Me.grdPOTSHIP3.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance24.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIP3.DisplayLayout.Override.TemplateAddRowAppearance = Appearance24
        Me.grdPOTSHIP3.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIP3.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIP3.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIP3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIP3.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIP3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdPOTSHIP3.Name = "grdPOTSHIP3"
        Me.grdPOTSHIP3.Size = New System.Drawing.Size(1015, 113)
        Me.grdPOTSHIP3.TabIndex = 11
        Me.grdPOTSHIP3.Text = "Bill of Lading Details"
        '
        'UltraTabPageControl1
        '
        Me.UltraTabPageControl1.Controls.Add(Me.SplitContainer1)
        Me.UltraTabPageControl1.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabPageControl1.Name = "UltraTabPageControl1"
        Me.UltraTabPageControl1.Size = New System.Drawing.Size(1019, 550)
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdICTRECI0)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.UltraTabControl3)
        Me.SplitContainer1.Size = New System.Drawing.Size(1019, 550)
        Me.SplitContainer1.SplitterDistance = 401
        Me.SplitContainer1.TabIndex = 165
        '
        'grdICTRECI0
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdICTRECI0.DisplayLayout.Appearance = Appearance1
        UltraGridColumn33.Header.Caption = "Style"
        UltraGridColumn33.Header.VisiblePosition = 0
        UltraGridColumn33.Width = 111
        UltraGridColumn34.Header.Caption = "Color"
        UltraGridColumn34.Header.VisiblePosition = 1
        UltraGridColumn34.Width = 51
        UltraGridColumn35.Header.Caption = "Class"
        UltraGridColumn35.Header.VisiblePosition = 2
        UltraGridColumn35.Width = 61
        UltraGridColumn36.Header.Caption = "Description"
        UltraGridColumn36.Header.VisiblePosition = 3
        UltraGridColumn37.Header.Caption = "#Beg"
        UltraGridColumn37.Header.VisiblePosition = 4
        UltraGridColumn38.Header.Caption = "$Beg"
        UltraGridColumn38.Header.VisiblePosition = 10
        UltraGridColumn39.Header.Caption = "#Rec"
        UltraGridColumn39.Header.VisiblePosition = 5
        UltraGridColumn40.Header.Caption = "$Rec"
        UltraGridColumn40.Header.VisiblePosition = 11
        UltraGridColumn41.Header.Caption = "#Shp"
        UltraGridColumn41.Header.VisiblePosition = 6
        UltraGridColumn42.Header.Caption = "$Shp"
        UltraGridColumn42.Header.VisiblePosition = 12
        UltraGridColumn43.Header.Caption = "#Adj"
        UltraGridColumn43.Header.VisiblePosition = 7
        UltraGridColumn44.Header.Caption = "$Adj"
        UltraGridColumn44.Header.VisiblePosition = 13
        UltraGridColumn45.Header.Caption = "#EOM"
        UltraGridColumn45.Header.VisiblePosition = 8
        UltraGridColumn1.Header.Caption = "$EOM"
        UltraGridColumn1.Header.VisiblePosition = 14
        UltraGridColumn3.Header.Caption = "#OOB"
        UltraGridColumn3.Header.VisiblePosition = 9
        UltraGridColumn4.Header.Caption = "$OOB"
        UltraGridColumn4.Header.VisiblePosition = 15
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn33, UltraGridColumn34, UltraGridColumn35, UltraGridColumn36, UltraGridColumn37, UltraGridColumn38, UltraGridColumn39, UltraGridColumn40, UltraGridColumn41, UltraGridColumn42, UltraGridColumn43, UltraGridColumn44, UltraGridColumn45, UltraGridColumn1, UltraGridColumn3, UltraGridColumn4})
        Me.grdICTRECI0.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdICTRECI0.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdICTRECI0.DisplayLayout.CaptionAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECI0.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECI0.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdICTRECI0.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdICTRECI0.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECI0.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdICTRECI0.DisplayLayout.MaxColScrollRegions = 1
        Me.grdICTRECI0.DisplayLayout.MaxRowScrollRegions = 1
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdICTRECI0.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdICTRECI0.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdICTRECI0.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECI0.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECI0.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdICTRECI0.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.grdICTRECI0.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdICTRECI0.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdICTRECI0.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdICTRECI0.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECI0.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdICTRECI0.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdICTRECI0.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdICTRECI0.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdICTRECI0.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdICTRECI0.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdICTRECI0.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdICTRECI0.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdICTRECI0.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdICTRECI0.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdICTRECI0.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTRECI0.Location = New System.Drawing.Point(0, 0)
        Me.grdICTRECI0.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdICTRECI0.Name = "grdICTRECI0"
        Me.grdICTRECI0.Size = New System.Drawing.Size(1019, 401)
        Me.grdICTRECI0.TabIndex = 164
        Me.grdICTRECI0.Text = "Inventory Roll Forward"
        '
        'UltraTabControl3
        '
        Me.UltraTabControl3.Controls.Add(Me.UltraTabSharedControlsPage5)
        Me.UltraTabControl3.Controls.Add(Me.UltraTabPageControl8)
        Me.UltraTabControl3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraTabControl3.Location = New System.Drawing.Point(0, 0)
        Me.UltraTabControl3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabControl3.Name = "UltraTabControl3"
        Me.UltraTabControl3.SharedControlsPage = Me.UltraTabSharedControlsPage5
        Me.UltraTabControl3.Size = New System.Drawing.Size(1019, 145)
        Me.UltraTabControl3.TabIndex = 176
        Me.UltraTabControl3.TabOrientation = Infragistics.Win.UltraWinTabs.TabOrientation.TopLeft
        UltraTab8.TabPage = Me.UltraTabPageControl8
        UltraTab8.Text = "PO Details"
        Me.UltraTabControl3.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab8})
        '
        'UltraTabSharedControlsPage5
        '
        Me.UltraTabSharedControlsPage5.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabSharedControlsPage5.Name = "UltraTabSharedControlsPage5"
        Me.UltraTabSharedControlsPage5.Size = New System.Drawing.Size(1015, 113)
        '
        'UltraTabPageControl2
        '
        Me.UltraTabPageControl2.Controls.Add(Me.grdICTRECIG)
        Me.UltraTabPageControl2.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabPageControl2.Name = "UltraTabPageControl2"
        Me.UltraTabPageControl2.Size = New System.Drawing.Size(1019, 550)
        '
        'grdICTRECIG
        '
        Appearance25.BackColor = System.Drawing.SystemColors.Window
        Appearance25.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdICTRECIG.DisplayLayout.Appearance = Appearance25
        UltraGridColumn5.Header.Caption = "Type"
        UltraGridColumn5.Header.VisiblePosition = 0
        UltraGridColumn5.Width = 77
        UltraGridColumn6.Header.Caption = "GL Amt"
        UltraGridColumn6.Header.VisiblePosition = 1
        UltraGridColumn7.Header.Caption = "GL Amt C"
        UltraGridColumn7.Header.VisiblePosition = 2
        UltraGridColumn8.Header.Caption = "GL Amt R"
        UltraGridColumn8.Header.VisiblePosition = 3
        UltraGridColumn10.Header.Caption = "IC Amt"
        UltraGridColumn10.Header.VisiblePosition = 4
        UltraGridColumn11.Header.Caption = "IC Amt C"
        UltraGridColumn11.Header.VisiblePosition = 5
        UltraGridColumn55.Header.Caption = "IC Amt R"
        UltraGridColumn55.Header.VisiblePosition = 6
        UltraGridColumn56.Header.Caption = "Difference"
        UltraGridColumn56.Header.VisiblePosition = 7
        UltraGridBand3.Columns.AddRange(New Object() {UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn10, UltraGridColumn11, UltraGridColumn55, UltraGridColumn56})
        Me.grdICTRECIG.DisplayLayout.BandsSerializer.Add(UltraGridBand3)
        Me.grdICTRECIG.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance26.TextHAlignAsString = "Left"
        Me.grdICTRECIG.DisplayLayout.CaptionAppearance = Appearance26
        Appearance27.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance27.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance27.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECIG.DisplayLayout.GroupByBox.Appearance = Appearance27
        Appearance28.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECIG.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance28
        Me.grdICTRECIG.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdICTRECIG.DisplayLayout.GroupByBox.Hidden = True
        Appearance29.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance29.BackColor2 = System.Drawing.SystemColors.Control
        Appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance29.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECIG.DisplayLayout.GroupByBox.PromptAppearance = Appearance29
        Me.grdICTRECIG.DisplayLayout.MaxColScrollRegions = 1
        Me.grdICTRECIG.DisplayLayout.MaxRowScrollRegions = 1
        Appearance30.BackColor = System.Drawing.SystemColors.Window
        Appearance30.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdICTRECIG.DisplayLayout.Override.ActiveCellAppearance = Appearance30
        Me.grdICTRECIG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdICTRECIG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECIG.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECIG.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdICTRECIG.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance31.BackColor = System.Drawing.SystemColors.Window
        Me.grdICTRECIG.DisplayLayout.Override.CardAreaAppearance = Appearance31
        Appearance32.BorderColor = System.Drawing.Color.Silver
        Appearance32.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdICTRECIG.DisplayLayout.Override.CellAppearance = Appearance32
        Me.grdICTRECIG.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdICTRECIG.DisplayLayout.Override.CellPadding = 0
        Appearance33.BackColor = System.Drawing.SystemColors.Control
        Appearance33.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance33.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance33.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECIG.DisplayLayout.Override.GroupByRowAppearance = Appearance33
        Appearance34.TextHAlignAsString = "Left"
        Me.grdICTRECIG.DisplayLayout.Override.HeaderAppearance = Appearance34
        Me.grdICTRECIG.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdICTRECIG.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance35.BackColor = System.Drawing.SystemColors.Window
        Appearance35.BorderColor = System.Drawing.Color.Silver
        Me.grdICTRECIG.DisplayLayout.Override.RowAppearance = Appearance35
        Me.grdICTRECIG.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance36.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdICTRECIG.DisplayLayout.Override.TemplateAddRowAppearance = Appearance36
        Me.grdICTRECIG.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdICTRECIG.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdICTRECIG.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdICTRECIG.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTRECIG.Location = New System.Drawing.Point(0, 0)
        Me.grdICTRECIG.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdICTRECIG.Name = "grdICTRECIG"
        Me.grdICTRECIG.Size = New System.Drawing.Size(1019, 550)
        Me.grdICTRECIG.TabIndex = 165
        '
        'UltraTabPageControl3
        '
        Me.UltraTabPageControl3.Controls.Add(Me.grdPOTSHIPX)
        Me.UltraTabPageControl3.Location = New System.Drawing.Point(1, 1)
        Me.UltraTabPageControl3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabPageControl3.Name = "UltraTabPageControl3"
        Me.UltraTabPageControl3.Size = New System.Drawing.Size(1019, 550)
        '
        'grdPOTSHIPX
        '
        Appearance37.BackColor = System.Drawing.SystemColors.Window
        Appearance37.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIPX.DisplayLayout.Appearance = Appearance37
        UltraGridColumn12.Header.Caption = "Shipment No"
        UltraGridColumn12.Header.VisiblePosition = 0
        UltraGridColumn13.Header.Caption = "Vessel"
        UltraGridColumn13.Header.VisiblePosition = 1
        UltraGridColumn13.Width = 153
        UltraGridColumn20.Header.Caption = "Received"
        UltraGridColumn20.Header.VisiblePosition = 2
        UltraGridColumn20.Width = 93
        UltraGridColumn21.Header.Caption = "Ship ETA"
        UltraGridColumn21.Header.VisiblePosition = 3
        UltraGridColumn21.Width = 90
        UltraGridColumn22.Header.Caption = "Qty"
        UltraGridColumn22.Header.VisiblePosition = 4
        UltraGridColumn22.Width = 97
        UltraGridColumn23.Header.Caption = "Amt"
        UltraGridColumn23.Header.VisiblePosition = 5
        UltraGridColumn23.Width = 105
        UltraGridColumn86.Header.Caption = "Check Date"
        UltraGridColumn86.Header.VisiblePosition = 6
        UltraGridBand4.Columns.AddRange(New Object() {UltraGridColumn12, UltraGridColumn13, UltraGridColumn20, UltraGridColumn21, UltraGridColumn22, UltraGridColumn23, UltraGridColumn86})
        Me.grdPOTSHIPX.DisplayLayout.BandsSerializer.Add(UltraGridBand4)
        Me.grdPOTSHIPX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance38.TextHAlignAsString = "Left"
        Me.grdPOTSHIPX.DisplayLayout.CaptionAppearance = Appearance38
        Appearance39.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance39.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance39.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.Appearance = Appearance39
        Appearance40.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance40
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.Hidden = True
        Appearance41.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance41.BackColor2 = System.Drawing.SystemColors.Control
        Appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance41.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.PromptAppearance = Appearance41
        Me.grdPOTSHIPX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIPX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance42.BackColor = System.Drawing.SystemColors.Window
        Appearance42.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIPX.DisplayLayout.Override.ActiveCellAppearance = Appearance42
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIPX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance43.BackColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.Override.CardAreaAppearance = Appearance43
        Appearance44.BorderColor = System.Drawing.Color.Silver
        Appearance44.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIPX.DisplayLayout.Override.CellAppearance = Appearance44
        Me.grdPOTSHIPX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdPOTSHIPX.DisplayLayout.Override.CellPadding = 0
        Appearance45.BackColor = System.Drawing.SystemColors.Control
        Appearance45.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance45.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance45.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.Override.GroupByRowAppearance = Appearance45
        Appearance46.TextHAlignAsString = "Left"
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderAppearance = Appearance46
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance47.BackColor = System.Drawing.SystemColors.Window
        Appearance47.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIPX.DisplayLayout.Override.RowAppearance = Appearance47
        Me.grdPOTSHIPX.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance48.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIPX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance48
        Me.grdPOTSHIPX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIPX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIPX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIPX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIPX.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIPX.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdPOTSHIPX.Name = "grdPOTSHIPX"
        Me.grdPOTSHIPX.Size = New System.Drawing.Size(1019, 550)
        Me.grdPOTSHIPX.TabIndex = 165
        Me.grdPOTSHIPX.Text = "Receipts"
        '
        'UltraTabPageControl4
        '
        Me.UltraTabPageControl4.Controls.Add(Me.grdPOTSHIPH)
        Me.UltraTabPageControl4.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl4.Name = "UltraTabPageControl4"
        Me.UltraTabPageControl4.Size = New System.Drawing.Size(1019, 550)
        '
        'grdPOTSHIPH
        '
        Appearance49.BackColor = System.Drawing.SystemColors.Window
        Appearance49.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIPH.DisplayLayout.Appearance = Appearance49
        UltraGridColumn24.Header.VisiblePosition = 0
        UltraGridColumn24.Hidden = True
        UltraGridColumn46.Header.Caption = "Ship No"
        UltraGridColumn46.Header.VisiblePosition = 1
        UltraGridColumn46.RowLayoutColumnInfo.OriginX = 0
        UltraGridColumn46.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn46.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(65, 0)
        UltraGridColumn46.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn46.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn46.Width = 65
        UltraGridColumn25.Header.Caption = "Ship L"
        UltraGridColumn25.Header.VisiblePosition = 2
        UltraGridColumn25.RowLayoutColumnInfo.OriginX = 2
        UltraGridColumn25.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn25.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(58, 0)
        UltraGridColumn25.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn25.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn25.Width = 55
        UltraGridColumn26.Header.Caption = "PO Ord No"
        UltraGridColumn26.Header.VisiblePosition = 3
        UltraGridColumn26.RowLayoutColumnInfo.OriginX = 4
        UltraGridColumn26.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn26.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(83, 0)
        UltraGridColumn26.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn26.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn26.Width = 82
        UltraGridColumn27.Header.Caption = "PO Ln"
        UltraGridColumn27.Header.VisiblePosition = 4
        UltraGridColumn27.RowLayoutColumnInfo.OriginX = 6
        UltraGridColumn27.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn27.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(78, 0)
        UltraGridColumn27.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn27.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn27.Width = 47
        UltraGridColumn28.Header.Caption = "Qty In Transit"
        UltraGridColumn28.Header.VisiblePosition = 5
        UltraGridColumn28.RowLayoutColumnInfo.OriginX = 8
        UltraGridColumn28.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn28.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(105, 0)
        UltraGridColumn28.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn28.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn28.Width = 105
        UltraGridColumn29.Header.VisiblePosition = 8
        UltraGridColumn29.Hidden = True
        UltraGridColumn30.Header.VisiblePosition = 9
        UltraGridColumn30.Hidden = True
        UltraGridColumn31.Header.VisiblePosition = 10
        UltraGridColumn31.Hidden = True
        UltraGridColumn32.Header.VisiblePosition = 11
        UltraGridColumn32.Hidden = True
        UltraGridColumn52.Header.VisiblePosition = 12
        UltraGridColumn52.Hidden = True
        UltraGridColumn53.Header.VisiblePosition = 13
        UltraGridColumn53.Hidden = True
        UltraGridColumn54.Header.Caption = "PO Cost"
        UltraGridColumn54.Header.VisiblePosition = 7
        UltraGridColumn54.RowLayoutColumnInfo.OriginX = 14
        UltraGridColumn54.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn54.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(89, 0)
        UltraGridColumn54.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn54.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn54.Width = 64
        UltraGridColumn57.Header.VisiblePosition = 14
        UltraGridColumn57.Hidden = True
        UltraGridColumn58.Header.VisiblePosition = 15
        UltraGridColumn58.Hidden = True
        UltraGridColumn59.Header.VisiblePosition = 16
        UltraGridColumn59.Hidden = True
        UltraGridColumn60.Header.VisiblePosition = 17
        UltraGridColumn60.Hidden = True
        UltraGridColumn61.Header.VisiblePosition = 18
        UltraGridColumn61.Hidden = True
        UltraGridColumn62.Header.VisiblePosition = 19
        UltraGridColumn62.Hidden = True
        UltraGridColumn63.Header.VisiblePosition = 20
        UltraGridColumn63.Hidden = True
        UltraGridColumn64.Header.VisiblePosition = 21
        UltraGridColumn64.Hidden = True
        UltraGridColumn65.Header.Caption = "Cost Vcost"
        UltraGridColumn65.Header.VisiblePosition = 22
        UltraGridColumn65.RowLayoutColumnInfo.OriginX = 16
        UltraGridColumn65.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn65.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(73, 0)
        UltraGridColumn65.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn65.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn65.Width = 81
        UltraGridColumn66.Header.Caption = "Cost Matls"
        UltraGridColumn66.Header.VisiblePosition = 23
        UltraGridColumn66.RowLayoutColumnInfo.OriginX = 18
        UltraGridColumn66.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn66.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(142, 0)
        UltraGridColumn66.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn66.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn66.Width = 78
        UltraGridColumn67.Header.VisiblePosition = 24
        UltraGridColumn67.Hidden = True
        UltraGridColumn68.Header.VisiblePosition = 25
        UltraGridColumn68.Hidden = True
        UltraGridColumn69.Header.VisiblePosition = 26
        UltraGridColumn69.Hidden = True
        UltraGridColumn70.Header.VisiblePosition = 27
        UltraGridColumn70.Hidden = True
        UltraGridColumn71.Header.VisiblePosition = 28
        UltraGridColumn71.Hidden = True
        UltraGridColumn72.Header.Caption = "Cost Other"
        UltraGridColumn72.Header.VisiblePosition = 29
        UltraGridColumn72.RowLayoutColumnInfo.OriginX = 20
        UltraGridColumn72.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn72.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(112, 0)
        UltraGridColumn72.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn72.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn72.Width = 72
        UltraGridColumn73.Header.VisiblePosition = 30
        UltraGridColumn73.Hidden = True
        UltraGridColumn74.Header.VisiblePosition = 31
        UltraGridColumn74.Hidden = True
        UltraGridColumn75.Header.VisiblePosition = 32
        UltraGridColumn75.Hidden = True
        UltraGridColumn76.Header.VisiblePosition = 33
        UltraGridColumn76.Hidden = True
        UltraGridColumn77.Header.VisiblePosition = 34
        UltraGridColumn77.Hidden = True
        UltraGridColumn78.Header.VisiblePosition = 35
        UltraGridColumn78.Hidden = True
        UltraGridColumn79.Header.VisiblePosition = 36
        UltraGridColumn79.Hidden = True
        UltraGridColumn80.Header.VisiblePosition = 37
        UltraGridColumn80.Hidden = True
        UltraGridColumn81.Header.VisiblePosition = 38
        UltraGridColumn81.Hidden = True
        UltraGridColumn82.Header.Caption = "Qty In Transit Amt"
        UltraGridColumn82.Header.VisiblePosition = 6
        UltraGridColumn82.RowLayoutColumnInfo.OriginX = 10
        UltraGridColumn82.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn82.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(127, 0)
        UltraGridColumn82.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn82.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn82.Width = 138
        UltraGridColumn47.Header.Caption = "Status"
        UltraGridColumn47.Header.VisiblePosition = 39
        UltraGridColumn47.RowLayoutColumnInfo.OriginX = 12
        UltraGridColumn47.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn47.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(107, 0)
        UltraGridColumn47.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn47.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn47.Width = 111
        UltraGridColumn48.Header.Caption = "Voucher No"
        UltraGridColumn48.Header.VisiblePosition = 40
        UltraGridColumn48.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(104, 0)
        UltraGridColumn48.Width = 93
        UltraGridColumn49.Header.Caption = "Vend Cd"
        UltraGridColumn49.Header.VisiblePosition = 44
        UltraGridColumn49.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(73, 0)
        UltraGridColumn49.Width = 80
        UltraGridColumn50.Header.Caption = "Vend Name"
        UltraGridColumn50.Header.VisiblePosition = 45
        UltraGridColumn50.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(142, 0)
        UltraGridColumn50.Width = 147
        UltraGridColumn51.Header.Caption = "Date Shipped"
        UltraGridColumn51.Header.VisiblePosition = 41
        UltraGridColumn51.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(112, 0)
        UltraGridColumn51.Width = 108
        UltraGridColumn83.Header.Caption = "Inv Num"
        UltraGridColumn83.Header.VisiblePosition = 42
        UltraGridColumn83.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(71, 0)
        UltraGridColumn83.Width = 71
        UltraGridColumn84.Header.Caption = "Inv Date"
        UltraGridColumn84.Header.VisiblePosition = 43
        UltraGridColumn84.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(95, 0)
        UltraGridColumn84.Width = 88
        UltraGridColumn85.Header.Caption = "Check Date"
        UltraGridColumn85.Header.VisiblePosition = 46
        UltraGridBand5.Columns.AddRange(New Object() {UltraGridColumn24, UltraGridColumn46, UltraGridColumn25, UltraGridColumn26, UltraGridColumn27, UltraGridColumn28, UltraGridColumn29, UltraGridColumn30, UltraGridColumn31, UltraGridColumn32, UltraGridColumn52, UltraGridColumn53, UltraGridColumn54, UltraGridColumn57, UltraGridColumn58, UltraGridColumn59, UltraGridColumn60, UltraGridColumn61, UltraGridColumn62, UltraGridColumn63, UltraGridColumn64, UltraGridColumn65, UltraGridColumn66, UltraGridColumn67, UltraGridColumn68, UltraGridColumn69, UltraGridColumn70, UltraGridColumn71, UltraGridColumn72, UltraGridColumn73, UltraGridColumn74, UltraGridColumn75, UltraGridColumn76, UltraGridColumn77, UltraGridColumn78, UltraGridColumn79, UltraGridColumn80, UltraGridColumn81, UltraGridColumn82, UltraGridColumn47, UltraGridColumn48, UltraGridColumn49, UltraGridColumn50, UltraGridColumn51, UltraGridColumn83, UltraGridColumn84, UltraGridColumn85})
        Me.grdPOTSHIPH.DisplayLayout.BandsSerializer.Add(UltraGridBand5)
        Me.grdPOTSHIPH.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance50.TextHAlignAsString = "Left"
        Me.grdPOTSHIPH.DisplayLayout.CaptionAppearance = Appearance50
        Appearance51.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance51.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance51.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPH.DisplayLayout.GroupByBox.Appearance = Appearance51
        Appearance52.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPH.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance52
        Me.grdPOTSHIPH.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIPH.DisplayLayout.GroupByBox.Hidden = True
        Appearance53.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance53.BackColor2 = System.Drawing.SystemColors.Control
        Appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance53.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPH.DisplayLayout.GroupByBox.PromptAppearance = Appearance53
        Me.grdPOTSHIPH.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIPH.DisplayLayout.MaxRowScrollRegions = 1
        Appearance54.BackColor = System.Drawing.SystemColors.Window
        Appearance54.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIPH.DisplayLayout.Override.ActiveCellAppearance = Appearance54
        Me.grdPOTSHIPH.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIPH.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPH.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPH.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIPH.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance55.BackColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPH.DisplayLayout.Override.CardAreaAppearance = Appearance55
        Appearance56.BorderColor = System.Drawing.Color.Silver
        Appearance56.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIPH.DisplayLayout.Override.CellAppearance = Appearance56
        Me.grdPOTSHIPH.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdPOTSHIPH.DisplayLayout.Override.CellPadding = 0
        Appearance57.BackColor = System.Drawing.SystemColors.Control
        Appearance57.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance57.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance57.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPH.DisplayLayout.Override.GroupByRowAppearance = Appearance57
        Appearance58.TextHAlignAsString = "Left"
        Me.grdPOTSHIPH.DisplayLayout.Override.HeaderAppearance = Appearance58
        Me.grdPOTSHIPH.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIPH.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance59.BackColor = System.Drawing.SystemColors.Window
        Appearance59.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIPH.DisplayLayout.Override.RowAppearance = Appearance59
        Me.grdPOTSHIPH.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance60.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIPH.DisplayLayout.Override.TemplateAddRowAppearance = Appearance60
        Me.grdPOTSHIPH.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIPH.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIPH.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIPH.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIPH.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIPH.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdPOTSHIPH.Name = "grdPOTSHIPH"
        Me.grdPOTSHIPH.Size = New System.Drawing.Size(1019, 550)
        Me.grdPOTSHIPH.TabIndex = 166
        Me.grdPOTSHIPH.Text = "In Transit"
        '
        'UltraTabPageControl6
        '
        Me.UltraTabPageControl6.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabPageControl6.Name = "UltraTabPageControl6"
        Me.UltraTabPageControl6.Size = New System.Drawing.Size(1023, 582)
        '
        'UltraTabPageControl7
        '
        Me.UltraTabPageControl7.Controls.Add(Me.tabMain)
        Me.UltraTabPageControl7.Location = New System.Drawing.Point(1, 29)
        Me.UltraTabPageControl7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabPageControl7.Name = "UltraTabPageControl7"
        Me.UltraTabPageControl7.Size = New System.Drawing.Size(1023, 582)
        '
        'tabMain
        '
        Me.tabMain.Controls.Add(Me.UltraTabSharedControlsPage1)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl1)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl2)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl3)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl4)
        Me.tabMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabMain.Location = New System.Drawing.Point(0, 0)
        Me.tabMain.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SharedControlsPage = Me.UltraTabSharedControlsPage1
        Me.tabMain.Size = New System.Drawing.Size(1023, 582)
        Me.tabMain.TabIndex = 165
        Me.tabMain.TabOrientation = Infragistics.Win.UltraWinTabs.TabOrientation.BottomLeft
        UltraTab4.TabPage = Me.UltraTabPageControl1
        UltraTab4.Text = "Inventory Roll Forward"
        UltraTab5.TabPage = Me.UltraTabPageControl2
        UltraTab5.Text = "GL Inventory"
        UltraTab11.TabPage = Me.UltraTabPageControl3
        UltraTab11.Text = "Receipts"
        UltraTab1.TabPage = Me.UltraTabPageControl4
        UltraTab1.Text = "In Transit Eom"
        Me.tabMain.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab4, UltraTab5, UltraTab11, UltraTab1})
        '
        'UltraTabSharedControlsPage1
        '
        Me.UltraTabSharedControlsPage1.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabSharedControlsPage1.Name = "UltraTabSharedControlsPage1"
        Me.UltraTabSharedControlsPage1.Size = New System.Drawing.Size(1019, 550)
        '
        'UltraTabPageControl5
        '
        Me.UltraTabPageControl5.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl5.Name = "UltraTabPageControl5"
        Me.UltraTabPageControl5.Size = New System.Drawing.Size(765, 281)
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.cbeOPS_YYYYPP)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(1027, 75)
        Me.UltraGroupBox1.TabIndex = 2
        '
        'cbeOPS_YYYYPP
        '
        Me.Absx1.SetABSColumnName(Me.cbeOPS_YYYYPP, "OPS_YYYYPP")
        Me.cbeOPS_YYYYPP.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cbeOPS_YYYYPP.Location = New System.Drawing.Point(15, 40)
        Me.cbeOPS_YYYYPP.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cbeOPS_YYYYPP.Name = "cbeOPS_YYYYPP"
        Me.cbeOPS_YYYYPP.Size = New System.Drawing.Size(205, 29)
        Me.cbeOPS_YYYYPP.TabIndex = 116
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(15, 14)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(58, 22)
        Me.UltraLabel2.TabIndex = 2
        Me.UltraLabel2.Text = "Period"
        '
        'tab
        '
        Me.tab.Controls.Add(Me.UltraTabSharedControlsPage2)
        Me.tab.Controls.Add(Me.UltraTabPageControl6)
        Me.tab.Controls.Add(Me.UltraTabPageControl7)
        Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tab.Location = New System.Drawing.Point(0, 0)
        Me.tab.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.tab.Name = "tab"
        Me.tab.SharedControlsPage = Me.UltraTabSharedControlsPage2
        Me.tab.Size = New System.Drawing.Size(1027, 614)
        Me.tab.TabIndex = 166
        UltraTab6.TabPage = Me.UltraTabPageControl6
        UltraTab6.Text = "0"
        UltraTab7.TabPage = Me.UltraTabPageControl7
        UltraTab7.Text = "1"
        Me.tab.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab6, UltraTab7})
        '
        'UltraTabSharedControlsPage2
        '
        Me.UltraTabSharedControlsPage2.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTabSharedControlsPage2.Name = "UltraTabSharedControlsPage2"
        Me.UltraTabSharedControlsPage2.Size = New System.Drawing.Size(1023, 582)
        '
        'UltraCombo1
        '
        Appearance70.BackColor = System.Drawing.SystemColors.Window
        Appearance70.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.UltraCombo1.DisplayLayout.Appearance = Appearance70
        Me.UltraCombo1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.UltraCombo1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance71.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance71.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance71.BorderColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.GroupByBox.Appearance = Appearance71
        Appearance72.ForeColor = System.Drawing.SystemColors.GrayText
        Me.UltraCombo1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance72
        Me.UltraCombo1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance73.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance73.BackColor2 = System.Drawing.SystemColors.Control
        Appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance73.ForeColor = System.Drawing.SystemColors.GrayText
        Me.UltraCombo1.DisplayLayout.GroupByBox.PromptAppearance = Appearance73
        Me.UltraCombo1.DisplayLayout.MaxColScrollRegions = 1
        Me.UltraCombo1.DisplayLayout.MaxRowScrollRegions = 1
        Appearance74.BackColor = System.Drawing.SystemColors.Window
        Appearance74.ForeColor = System.Drawing.SystemColors.ControlText
        Me.UltraCombo1.DisplayLayout.Override.ActiveCellAppearance = Appearance74
        Appearance75.BackColor = System.Drawing.SystemColors.Highlight
        Appearance75.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.UltraCombo1.DisplayLayout.Override.ActiveRowAppearance = Appearance75
        Me.UltraCombo1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.UltraCombo1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance76.BackColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.Override.CardAreaAppearance = Appearance76
        Appearance77.BorderColor = System.Drawing.Color.Silver
        Appearance77.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.UltraCombo1.DisplayLayout.Override.CellAppearance = Appearance77
        Me.UltraCombo1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.UltraCombo1.DisplayLayout.Override.CellPadding = 0
        Appearance78.BackColor = System.Drawing.SystemColors.Control
        Appearance78.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance78.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance78.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance78.BorderColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.Override.GroupByRowAppearance = Appearance78
        Appearance79.TextHAlignAsString = "Left"
        Me.UltraCombo1.DisplayLayout.Override.HeaderAppearance = Appearance79
        Me.UltraCombo1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.UltraCombo1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance80.BackColor = System.Drawing.SystemColors.Window
        Appearance80.BorderColor = System.Drawing.Color.Silver
        Me.UltraCombo1.DisplayLayout.Override.RowAppearance = Appearance80
        Me.UltraCombo1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance81.BackColor = System.Drawing.SystemColors.ControlLight
        Me.UltraCombo1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance81
        Me.UltraCombo1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.UltraCombo1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.UltraCombo1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.UltraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList
        Me.UltraCombo1.Location = New System.Drawing.Point(0, 5)
        Me.UltraCombo1.Name = "UltraCombo1"
        Me.UltraCombo1.Size = New System.Drawing.Size(174, 22)
        Me.UltraCombo1.TabIndex = 0
        '
        'spl
        '
        Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.spl.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.spl.Location = New System.Drawing.Point(0, 0)
        Me.spl.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
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
        Me.spl.Size = New System.Drawing.Size(1027, 693)
        Me.spl.SplitterDistance = 75
        Me.spl.TabIndex = 168
        '
        'UltraTabControl1
        '
        Me.UltraTabControl1.Location = New System.Drawing.Point(0, 0)
        Me.UltraTabControl1.Name = "UltraTabControl1"
        Me.UltraTabControl1.SharedControlsPage = Me.UltraTabSharedControlsPage3
        Me.UltraTabControl1.Size = New System.Drawing.Size(200, 100)
        Me.UltraTabControl1.TabIndex = 0
        '
        'UltraTabSharedControlsPage3
        '
        Me.UltraTabSharedControlsPage3.Location = New System.Drawing.Point(1, 24)
        Me.UltraTabSharedControlsPage3.Name = "UltraTabSharedControlsPage3"
        Me.UltraTabSharedControlsPage3.Size = New System.Drawing.Size(196, 73)
        '
        'tabBOL
        '
        Me.tabBOL.Location = New System.Drawing.Point(0, 0)
        Me.tabBOL.Name = "tabBOL"
        Me.tabBOL.SharedControlsPage = Me.UltraTabSharedControlsPage7
        Me.tabBOL.Size = New System.Drawing.Size(200, 100)
        Me.tabBOL.TabIndex = 0
        '
        'UltraTabSharedControlsPage7
        '
        Me.UltraTabSharedControlsPage7.Location = New System.Drawing.Point(1, 24)
        Me.UltraTabSharedControlsPage7.Name = "UltraTabSharedControlsPage7"
        Me.UltraTabSharedControlsPage7.Size = New System.Drawing.Size(196, 73)
        '
        'UltraTabControl2
        '
        Me.UltraTabControl2.Location = New System.Drawing.Point(0, 0)
        Me.UltraTabControl2.Name = "UltraTabControl2"
        Me.UltraTabControl2.SharedControlsPage = Me.UltraTabSharedControlsPage4
        Me.UltraTabControl2.Size = New System.Drawing.Size(200, 100)
        Me.UltraTabControl2.TabIndex = 0
        '
        'UltraTabSharedControlsPage4
        '
        Me.UltraTabSharedControlsPage4.Location = New System.Drawing.Point(1, 24)
        Me.UltraTabSharedControlsPage4.Name = "UltraTabSharedControlsPage4"
        Me.UltraTabSharedControlsPage4.Size = New System.Drawing.Size(196, 73)
        '
        'ICFRECI1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1240, 693)
        Me.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.Name = "ICFRECI1"
        Me.Text = "ICFRECI1"
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraExplorerBar1.ResumeLayout(False)
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraExplorerBarContainerControl1.ResumeLayout(False)
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox2.ResumeLayout(False)
        CType(Me.chkOOB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl8.ResumeLayout(False)
        CType(Me.grdPOTSHIP3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grdICTRECI0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTabControl3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabControl3.ResumeLayout(False)
        Me.UltraTabPageControl2.ResumeLayout(False)
        CType(Me.grdICTRECIG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl3.ResumeLayout(False)
        CType(Me.grdPOTSHIPX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl4.ResumeLayout(False)
        CType(Me.grdPOTSHIPH, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl7.ResumeLayout(False)
        CType(Me.tabMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabMain.ResumeLayout(False)
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.UltraGroupBox1.PerformLayout()
        CType(Me.cbeOPS_YYYYPP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tab, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab.ResumeLayout(False)
        CType(Me.UltraCombo1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.Panel1.ResumeLayout(False)
        Me.spl.Panel2.ResumeLayout(False)
        CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.ResumeLayout(False)
        CType(Me.UltraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tabBOL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTabControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents grdICTRECI0 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraTabPageControl5 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents tab As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage2 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl6 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTabPageControl7 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraCombo1 As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents tabMain As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage1 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl1 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTabPageControl2 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents grdICTRECIG As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cbeOPS_YYYYPP As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents spl As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraExplorerBarContainerControl1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarContainerControl
    Friend WithEvents UltraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents chkOOB As ABSCS.ABSCheckBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraTabPageControl3 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents grdPOTSHIPX As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraTabControl1 As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage3 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents tabBOL As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage7 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabControl2 As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage4 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabControl3 As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage5 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl8 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents grdPOTSHIP3 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraTabPageControl4 As UltraWinTabControl.UltraTabPageControl
    Friend WithEvents grdPOTSHIPH As UltraWinGrid.UltraGrid
End Class
