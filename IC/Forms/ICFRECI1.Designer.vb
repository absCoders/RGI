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
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance70 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance62 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance63 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance64 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance130 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
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
        Dim Appearance132 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance133 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance134 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance135 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance136 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance137 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance138 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance139 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance140 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance141 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance142 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
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
        Dim Appearance38 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance39 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance42 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance43 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance44 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance46 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance47 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance48 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraTab8 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim Appearance49 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand3 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ICTRECIG", -1)
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("JOURNAL_TYPE")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GL_AMT")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GL_AMT_C")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GL_AMT_R")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IC_AMT")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IC_AMT_C")
        Dim UltraGridColumn55 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IC_AMT_R")
        Dim UltraGridColumn56 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DIFF")
        Dim Appearance50 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance51 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance52 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance53 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance54 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance55 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance56 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance57 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance58 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand4 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIPX", -1)
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn47 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_VESSEL")
        Dim UltraGridColumn48 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_RECEIVED")
        Dim UltraGridColumn49 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_ETA")
        Dim UltraGridColumn50 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY")
        Dim UltraGridColumn51 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("AMT")
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
        Dim Appearance37 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraTab4 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab5 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab11 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab6 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab7 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance109 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance110 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance111 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance65 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance66 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance67 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance68 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance69 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
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
        UltraExplorerBarGroup2.Settings.ContainerHeight = 31
        UltraExplorerBarGroup2.Settings.Style = Infragistics.Win.UltraWinExplorerBar.GroupStyle.ControlContainer
        UltraExplorerBarGroup2.Text = "Options"
        Me.UltraExplorerBar1.Groups.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup() {UltraExplorerBarGroup1, UltraExplorerBarGroup2})
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Margins.Bottom = 0
        Me.UltraExplorerBar1.Margins.Left = 0
        Me.UltraExplorerBar1.Margins.Right = 0
        Me.UltraExplorerBar1.Margins.Top = 0
        Me.UltraExplorerBar1.ShowDefaultContextMenu = False
        Me.UltraExplorerBar1.Tag = "CLICK"
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.spl)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.grdASFBASEX, 0)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.spl, 0)
        '
        'grdASFBASEX
        '
        Appearance15.BackColor = System.Drawing.SystemColors.Window
        Appearance15.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance15
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
        Appearance13.BorderColor = System.Drawing.Color.Silver
        Appearance13.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance13
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance70.BackColor = System.Drawing.SystemColors.Control
        Appearance70.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance70.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance70.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance70
        Appearance62.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance62
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance63.BackColor = System.Drawing.SystemColors.Window
        Appearance63.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance63
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance64.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance64
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraExplorerBarContainerControl1
        '
        Me.UltraExplorerBarContainerControl1.Controls.Add(Me.UltraGroupBox2)
        Me.UltraExplorerBarContainerControl1.Location = New System.Drawing.Point(13, 126)
        Me.UltraExplorerBarContainerControl1.Name = "UltraExplorerBarContainerControl1"
        Me.UltraExplorerBarContainerControl1.Size = New System.Drawing.Size(189, 31)
        Me.UltraExplorerBarContainerControl1.TabIndex = 0
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.Controls.Add(Me.chkOOB)
        Me.UltraGroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(189, 31)
        Me.UltraGroupBox2.TabIndex = 0
        '
        'chkOOB
        '
        Me.chkOOB.Location = New System.Drawing.Point(23, 4)
        Me.chkOOB.Name = "chkOOB"
        Me.chkOOB.Size = New System.Drawing.Size(120, 20)
        Me.chkOOB.TabIndex = 0
        Me.chkOOB.Text = "OOB Only"
        '
        'UltraTabPageControl8
        '
        Me.UltraTabPageControl8.Controls.Add(Me.grdPOTSHIP3)
        Me.UltraTabPageControl8.Location = New System.Drawing.Point(1, 25)
        Me.UltraTabPageControl8.Name = "UltraTabPageControl8"
        Me.UltraTabPageControl8.Size = New System.Drawing.Size(767, 97)
        '
        'grdPOTSHIP3
        '
        Appearance130.BackColor = System.Drawing.SystemColors.Window
        Appearance130.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIP3.DisplayLayout.Appearance = Appearance130
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
        Appearance132.TextHAlignAsString = "Left"
        Me.grdPOTSHIP3.DisplayLayout.CaptionAppearance = Appearance132
        Appearance133.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance133.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance133.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance133.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.Appearance = Appearance133
        Appearance134.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance134
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.Hidden = True
        Appearance135.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance135.BackColor2 = System.Drawing.SystemColors.Control
        Appearance135.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance135.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP3.DisplayLayout.GroupByBox.PromptAppearance = Appearance135
        Me.grdPOTSHIP3.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIP3.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdPOTSHIP3.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance136.BackColor = System.Drawing.SystemColors.Window
        Appearance136.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIP3.DisplayLayout.Override.ActiveCellAppearance = Appearance136
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP3.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIP3.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance137.BackColor = System.Drawing.Color.Transparent
        Me.grdPOTSHIP3.DisplayLayout.Override.CardAreaAppearance = Appearance137
        Appearance138.BorderColor = System.Drawing.Color.Silver
        Appearance138.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIP3.DisplayLayout.Override.CellAppearance = Appearance138
        Me.grdPOTSHIP3.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
        Me.grdPOTSHIP3.DisplayLayout.Override.CellPadding = 0
        Appearance139.BackColor = System.Drawing.SystemColors.Control
        Appearance139.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance139.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance139.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance139.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP3.DisplayLayout.Override.GroupByRowAppearance = Appearance139
        Appearance140.TextHAlignAsString = "Left"
        Me.grdPOTSHIP3.DisplayLayout.Override.HeaderAppearance = Appearance140
        Me.grdPOTSHIP3.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIP3.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance141.BackColor = System.Drawing.SystemColors.Window
        Appearance141.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIP3.DisplayLayout.Override.RowAppearance = Appearance141
        Me.grdPOTSHIP3.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance142.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIP3.DisplayLayout.Override.TemplateAddRowAppearance = Appearance142
        Me.grdPOTSHIP3.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIP3.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIP3.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIP3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIP3.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIP3.Name = "grdPOTSHIP3"
        Me.grdPOTSHIP3.Size = New System.Drawing.Size(767, 97)
        Me.grdPOTSHIP3.TabIndex = 11
        Me.grdPOTSHIP3.Text = "Bill of Lading Details"
        '
        'UltraTabPageControl1
        '
        Me.UltraTabPageControl1.Controls.Add(Me.SplitContainer1)
        Me.UltraTabPageControl1.Location = New System.Drawing.Point(1, 1)
        Me.UltraTabPageControl1.Name = "UltraTabPageControl1"
        Me.UltraTabPageControl1.Size = New System.Drawing.Size(771, 481)
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdICTRECI0)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.UltraTabControl3)
        Me.SplitContainer1.Size = New System.Drawing.Size(771, 481)
        Me.SplitContainer1.SplitterDistance = 352
        Me.SplitContainer1.TabIndex = 165
        '
        'grdICTRECI0
        '
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdICTRECI0.DisplayLayout.Appearance = Appearance2
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
        Appearance38.TextHAlignAsString = "Left"
        Me.grdICTRECI0.DisplayLayout.CaptionAppearance = Appearance38
        Appearance39.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance39.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance39.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECI0.DisplayLayout.GroupByBox.Appearance = Appearance39
        Appearance40.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECI0.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance40
        Me.grdICTRECI0.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdICTRECI0.DisplayLayout.GroupByBox.Hidden = True
        Appearance41.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance41.BackColor2 = System.Drawing.SystemColors.Control
        Appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance41.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECI0.DisplayLayout.GroupByBox.PromptAppearance = Appearance41
        Me.grdICTRECI0.DisplayLayout.MaxColScrollRegions = 1
        Me.grdICTRECI0.DisplayLayout.MaxRowScrollRegions = 1
        Appearance42.BackColor = System.Drawing.SystemColors.Window
        Appearance42.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdICTRECI0.DisplayLayout.Override.ActiveCellAppearance = Appearance42
        Me.grdICTRECI0.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdICTRECI0.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECI0.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECI0.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdICTRECI0.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance43.BackColor = System.Drawing.SystemColors.Window
        Me.grdICTRECI0.DisplayLayout.Override.CardAreaAppearance = Appearance43
        Appearance44.BorderColor = System.Drawing.Color.Silver
        Appearance44.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdICTRECI0.DisplayLayout.Override.CellAppearance = Appearance44
        Me.grdICTRECI0.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdICTRECI0.DisplayLayout.Override.CellPadding = 0
        Appearance3.BackColor = System.Drawing.SystemColors.Control
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECI0.DisplayLayout.Override.GroupByRowAppearance = Appearance3
        Appearance46.TextHAlignAsString = "Left"
        Me.grdICTRECI0.DisplayLayout.Override.HeaderAppearance = Appearance46
        Me.grdICTRECI0.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdICTRECI0.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance47.BackColor = System.Drawing.SystemColors.Window
        Appearance47.BorderColor = System.Drawing.Color.Silver
        Me.grdICTRECI0.DisplayLayout.Override.RowAppearance = Appearance47
        Me.grdICTRECI0.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance48.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdICTRECI0.DisplayLayout.Override.TemplateAddRowAppearance = Appearance48
        Me.grdICTRECI0.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdICTRECI0.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdICTRECI0.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdICTRECI0.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTRECI0.Location = New System.Drawing.Point(0, 0)
        Me.grdICTRECI0.Name = "grdICTRECI0"
        Me.grdICTRECI0.Size = New System.Drawing.Size(771, 352)
        Me.grdICTRECI0.TabIndex = 164
        Me.grdICTRECI0.Text = "Inventory Roll Forward"
        '
        'UltraTabControl3
        '
        Me.UltraTabControl3.Controls.Add(Me.UltraTabSharedControlsPage5)
        Me.UltraTabControl3.Controls.Add(Me.UltraTabPageControl8)
        Me.UltraTabControl3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraTabControl3.Location = New System.Drawing.Point(0, 0)
        Me.UltraTabControl3.Name = "UltraTabControl3"
        Me.UltraTabControl3.SharedControlsPage = Me.UltraTabSharedControlsPage5
        Me.UltraTabControl3.Size = New System.Drawing.Size(771, 125)
        Me.UltraTabControl3.TabIndex = 176
        Me.UltraTabControl3.TabOrientation = Infragistics.Win.UltraWinTabs.TabOrientation.TopLeft
        UltraTab8.TabPage = Me.UltraTabPageControl8
        UltraTab8.Text = "PO Details"
        Me.UltraTabControl3.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab8})
        '
        'UltraTabSharedControlsPage5
        '
        Me.UltraTabSharedControlsPage5.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage5.Name = "UltraTabSharedControlsPage5"
        Me.UltraTabSharedControlsPage5.Size = New System.Drawing.Size(767, 97)
        '
        'UltraTabPageControl2
        '
        Me.UltraTabPageControl2.Controls.Add(Me.grdICTRECIG)
        Me.UltraTabPageControl2.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl2.Name = "UltraTabPageControl2"
        Me.UltraTabPageControl2.Size = New System.Drawing.Size(771, 481)
        '
        'grdICTRECIG
        '
        Appearance49.BackColor = System.Drawing.SystemColors.Window
        Appearance49.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdICTRECIG.DisplayLayout.Appearance = Appearance49
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
        Appearance50.TextHAlignAsString = "Left"
        Me.grdICTRECIG.DisplayLayout.CaptionAppearance = Appearance50
        Appearance51.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance51.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance51.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECIG.DisplayLayout.GroupByBox.Appearance = Appearance51
        Appearance52.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECIG.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance52
        Me.grdICTRECIG.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdICTRECIG.DisplayLayout.GroupByBox.Hidden = True
        Appearance53.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance53.BackColor2 = System.Drawing.SystemColors.Control
        Appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance53.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTRECIG.DisplayLayout.GroupByBox.PromptAppearance = Appearance53
        Me.grdICTRECIG.DisplayLayout.MaxColScrollRegions = 1
        Me.grdICTRECIG.DisplayLayout.MaxRowScrollRegions = 1
        Appearance54.BackColor = System.Drawing.SystemColors.Window
        Appearance54.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdICTRECIG.DisplayLayout.Override.ActiveCellAppearance = Appearance54
        Me.grdICTRECIG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdICTRECIG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECIG.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTRECIG.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdICTRECIG.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance55.BackColor = System.Drawing.SystemColors.Window
        Me.grdICTRECIG.DisplayLayout.Override.CardAreaAppearance = Appearance55
        Appearance56.BorderColor = System.Drawing.Color.Silver
        Appearance56.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdICTRECIG.DisplayLayout.Override.CellAppearance = Appearance56
        Me.grdICTRECIG.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdICTRECIG.DisplayLayout.Override.CellPadding = 0
        Appearance57.BackColor = System.Drawing.SystemColors.Control
        Appearance57.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance57.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance57.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTRECIG.DisplayLayout.Override.GroupByRowAppearance = Appearance57
        Appearance58.TextHAlignAsString = "Left"
        Me.grdICTRECIG.DisplayLayout.Override.HeaderAppearance = Appearance58
        Me.grdICTRECIG.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdICTRECIG.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance8.BackColor = System.Drawing.SystemColors.Window
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Me.grdICTRECIG.DisplayLayout.Override.RowAppearance = Appearance8
        Me.grdICTRECIG.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance9.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdICTRECIG.DisplayLayout.Override.TemplateAddRowAppearance = Appearance9
        Me.grdICTRECIG.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdICTRECIG.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdICTRECIG.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdICTRECIG.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTRECIG.Location = New System.Drawing.Point(0, 0)
        Me.grdICTRECIG.Name = "grdICTRECIG"
        Me.grdICTRECIG.Size = New System.Drawing.Size(771, 481)
        Me.grdICTRECIG.TabIndex = 165
        '
        'UltraTabPageControl3
        '
        Me.UltraTabPageControl3.Controls.Add(Me.grdPOTSHIPX)
        Me.UltraTabPageControl3.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl3.Name = "UltraTabPageControl3"
        Me.UltraTabPageControl3.Size = New System.Drawing.Size(771, 481)
        '
        'grdPOTSHIPX
        '
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIPX.DisplayLayout.Appearance = Appearance14
        UltraGridColumn46.Header.Caption = "Shipment No"
        UltraGridColumn46.Header.VisiblePosition = 0
        UltraGridColumn47.Header.Caption = "Vessel"
        UltraGridColumn47.Header.VisiblePosition = 1
        UltraGridColumn47.Width = 153
        UltraGridColumn48.Header.Caption = "Received"
        UltraGridColumn48.Header.VisiblePosition = 2
        UltraGridColumn48.Width = 93
        UltraGridColumn49.Header.Caption = "Ship ETA"
        UltraGridColumn49.Header.VisiblePosition = 3
        UltraGridColumn49.Width = 90
        UltraGridColumn50.Header.Caption = "Qty"
        UltraGridColumn50.Header.VisiblePosition = 4
        UltraGridColumn50.Width = 97
        UltraGridColumn51.Header.Caption = "Amt"
        UltraGridColumn51.Header.VisiblePosition = 5
        UltraGridColumn51.Width = 105
        UltraGridBand4.Columns.AddRange(New Object() {UltraGridColumn46, UltraGridColumn47, UltraGridColumn48, UltraGridColumn49, UltraGridColumn50, UltraGridColumn51})
        Me.grdPOTSHIPX.DisplayLayout.BandsSerializer.Add(UltraGridBand4)
        Me.grdPOTSHIPX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance16.TextHAlignAsString = "Left"
        Me.grdPOTSHIPX.DisplayLayout.CaptionAppearance = Appearance16
        Appearance17.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance17.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance17.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.Appearance = Appearance17
        Appearance18.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance18
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.Hidden = True
        Appearance19.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance19.BackColor2 = System.Drawing.SystemColors.Control
        Appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance19.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.PromptAppearance = Appearance19
        Me.grdPOTSHIPX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIPX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Appearance20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIPX.DisplayLayout.Override.ActiveCellAppearance = Appearance20
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIPX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance21.BackColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.Override.CardAreaAppearance = Appearance21
        Appearance22.BorderColor = System.Drawing.Color.Silver
        Appearance22.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIPX.DisplayLayout.Override.CellAppearance = Appearance22
        Me.grdPOTSHIPX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdPOTSHIPX.DisplayLayout.Override.CellPadding = 0
        Appearance23.BackColor = System.Drawing.SystemColors.Control
        Appearance23.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance23.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance23.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.Override.GroupByRowAppearance = Appearance23
        Appearance24.TextHAlignAsString = "Left"
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderAppearance = Appearance24
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance25.BackColor = System.Drawing.SystemColors.Window
        Appearance25.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIPX.DisplayLayout.Override.RowAppearance = Appearance25
        Me.grdPOTSHIPX.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance37.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIPX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance37
        Me.grdPOTSHIPX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIPX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIPX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIPX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIPX.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIPX.Name = "grdPOTSHIPX"
        Me.grdPOTSHIPX.Size = New System.Drawing.Size(771, 481)
        Me.grdPOTSHIPX.TabIndex = 165
        Me.grdPOTSHIPX.Text = "Receipts"
        '
        'UltraTabPageControl6
        '
        Me.UltraTabPageControl6.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl6.Name = "UltraTabPageControl6"
        Me.UltraTabPageControl6.Size = New System.Drawing.Size(775, 509)
        '
        'UltraTabPageControl7
        '
        Me.UltraTabPageControl7.Controls.Add(Me.tabMain)
        Me.UltraTabPageControl7.Location = New System.Drawing.Point(1, 25)
        Me.UltraTabPageControl7.Name = "UltraTabPageControl7"
        Me.UltraTabPageControl7.Size = New System.Drawing.Size(775, 509)
        '
        'tabMain
        '
        Me.tabMain.Controls.Add(Me.UltraTabSharedControlsPage1)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl1)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl2)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl3)
        Me.tabMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabMain.Location = New System.Drawing.Point(0, 0)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SharedControlsPage = Me.UltraTabSharedControlsPage1
        Me.tabMain.Size = New System.Drawing.Size(775, 509)
        Me.tabMain.TabIndex = 165
        Me.tabMain.TabOrientation = Infragistics.Win.UltraWinTabs.TabOrientation.BottomLeft
        UltraTab4.TabPage = Me.UltraTabPageControl1
        UltraTab4.Text = "Inventory Roll Forward"
        UltraTab5.TabPage = Me.UltraTabPageControl2
        UltraTab5.Text = "GL Inventory"
        UltraTab11.TabPage = Me.UltraTabPageControl3
        UltraTab11.Text = "Receipts"
        Me.tabMain.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab4, UltraTab5, UltraTab11})
        '
        'UltraTabSharedControlsPage1
        '
        Me.UltraTabSharedControlsPage1.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage1.Name = "UltraTabSharedControlsPage1"
        Me.UltraTabSharedControlsPage1.Size = New System.Drawing.Size(771, 481)
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
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(779, 75)
        Me.UltraGroupBox1.TabIndex = 2
        '
        'cbeOPS_YYYYPP
        '
        Me.Absx1.SetABSColumnName(Me.cbeOPS_YYYYPP, "OPS_YYYYPP")
        Me.cbeOPS_YYYYPP.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cbeOPS_YYYYPP.Location = New System.Drawing.Point(12, 36)
        Me.cbeOPS_YYYYPP.Name = "cbeOPS_YYYYPP"
        Me.cbeOPS_YYYYPP.Size = New System.Drawing.Size(164, 25)
        Me.cbeOPS_YYYYPP.TabIndex = 116
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(12, 12)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(47, 18)
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
        Me.tab.Name = "tab"
        Me.tab.SharedControlsPage = Me.UltraTabSharedControlsPage2
        Me.tab.Size = New System.Drawing.Size(779, 537)
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
        Me.UltraTabSharedControlsPage2.Name = "UltraTabSharedControlsPage2"
        Me.UltraTabSharedControlsPage2.Size = New System.Drawing.Size(775, 509)
        '
        'UltraCombo1
        '
        Me.UltraCombo1.CheckedListSettings.CheckStateMember = ""
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Appearance5.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.UltraCombo1.DisplayLayout.Appearance = Appearance5
        Me.UltraCombo1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.UltraCombo1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance6.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance6.BorderColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.GroupByBox.Appearance = Appearance6
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.UltraCombo1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.UltraCombo1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance7.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance7.BackColor2 = System.Drawing.SystemColors.Control
        Appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance7.ForeColor = System.Drawing.SystemColors.GrayText
        Me.UltraCombo1.DisplayLayout.GroupByBox.PromptAppearance = Appearance7
        Me.UltraCombo1.DisplayLayout.MaxColScrollRegions = 1
        Me.UltraCombo1.DisplayLayout.MaxRowScrollRegions = 1
        Appearance109.BackColor = System.Drawing.SystemColors.Window
        Appearance109.ForeColor = System.Drawing.SystemColors.ControlText
        Me.UltraCombo1.DisplayLayout.Override.ActiveCellAppearance = Appearance109
        Appearance110.BackColor = System.Drawing.SystemColors.Highlight
        Appearance110.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.UltraCombo1.DisplayLayout.Override.ActiveRowAppearance = Appearance110
        Me.UltraCombo1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.UltraCombo1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance111.BackColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.Override.CardAreaAppearance = Appearance111
        Appearance65.BorderColor = System.Drawing.Color.Silver
        Appearance65.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.UltraCombo1.DisplayLayout.Override.CellAppearance = Appearance65
        Me.UltraCombo1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.UltraCombo1.DisplayLayout.Override.CellPadding = 0
        Appearance66.BackColor = System.Drawing.SystemColors.Control
        Appearance66.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance66.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance66.BorderColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.Override.GroupByRowAppearance = Appearance66
        Appearance67.TextHAlignAsString = "Left"
        Me.UltraCombo1.DisplayLayout.Override.HeaderAppearance = Appearance67
        Me.UltraCombo1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.UltraCombo1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance68.BackColor = System.Drawing.SystemColors.Window
        Appearance68.BorderColor = System.Drawing.Color.Silver
        Me.UltraCombo1.DisplayLayout.Override.RowAppearance = Appearance68
        Me.UltraCombo1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance69.BackColor = System.Drawing.SystemColors.ControlLight
        Me.UltraCombo1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance69
        Me.UltraCombo1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.UltraCombo1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.UltraCombo1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.UltraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList
        Me.UltraCombo1.Location = New System.Drawing.Point(0, 5)
        Me.UltraCombo1.Name = "UltraCombo1"
        Me.UltraCombo1.PreferredDropDownSize = New System.Drawing.Size(0, 0)
        Me.UltraCombo1.Size = New System.Drawing.Size(174, 22)
        Me.UltraCombo1.TabIndex = 0
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
        Me.spl.Size = New System.Drawing.Size(779, 616)
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
        Me.UltraTabSharedControlsPage3.Location = New System.Drawing.Point(1, 20)
        Me.UltraTabSharedControlsPage3.Name = "UltraTabSharedControlsPage3"
        Me.UltraTabSharedControlsPage3.Size = New System.Drawing.Size(196, 77)
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
        Me.UltraTabSharedControlsPage7.Location = New System.Drawing.Point(1, 20)
        Me.UltraTabSharedControlsPage7.Name = "UltraTabSharedControlsPage7"
        Me.UltraTabSharedControlsPage7.Size = New System.Drawing.Size(196, 77)
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
        Me.UltraTabSharedControlsPage4.Location = New System.Drawing.Point(1, 20)
        Me.UltraTabSharedControlsPage4.Name = "UltraTabSharedControlsPage4"
        Me.UltraTabSharedControlsPage4.Size = New System.Drawing.Size(196, 77)
        '
        'ICFRECI1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(992, 616)
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
End Class
