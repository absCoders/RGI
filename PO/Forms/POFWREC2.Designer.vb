<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class POFWREC2
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
        Dim Appearance85 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance86 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance87 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance88 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance89 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance90 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance91 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance92 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance93 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIP3", -1)
        Dim UltraGridColumn29 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn30 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_LNO")
        Dim UltraGridColumn31 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_NO")
        Dim UltraGridColumn32 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_LNO")
        Dim UltraGridColumn33 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SHP")
        Dim UltraGridColumn65 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE")
        Dim UltraGridColumn66 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn68 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_UOM")
        Dim UltraGridColumn34 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_REFERENCE")
        Dim UltraGridColumn35 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SPEC_ORDR_NO")
        Dim UltraGridColumn70 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_DESC")
        Dim UltraGridColumn72 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SUB_UNIT_PACK_QTY")
        Dim UltraGridColumn38 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_PACK_QTY")
        Dim UltraGridColumn122 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CASE_CUBE")
        Dim UltraGridColumn123 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_WEIGHT")
        Dim UltraGridColumn77 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_UOM")
        Dim UltraGridColumn78 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_BIN")
        Dim UltraGridColumn124 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("UPC_CODE")
        Dim UltraGridColumn125 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn101 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME")
        Dim UltraGridColumn74 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("UNITS")
        Dim UltraGridColumn79 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_NOS")
        Dim UltraGridColumn80 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTONS")
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
        Dim UltraGridBand3 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIP7", -1)
        Dim UltraGridColumn178 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn179 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_LNO")
        Dim UltraGridColumn180 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_NO")
        Dim UltraGridColumn181 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTONS")
        Dim UltraGridColumn182 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_COMMENTS")
        Dim UltraGridColumn183 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUSTOM_PPK")
        Dim UltraGridColumn184 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PPK_CODE")
        Dim UltraGridColumn185 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_PER_CTN")
        Dim UltraGridColumn186 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE")
        Dim UltraGridColumn187 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn188 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PPK_INNER_QTY")
        Dim UltraGridColumn157 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_DIMS")
        Dim UltraGridColumn161 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_VOLUME")
        Dim UltraGridColumn382 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_WEIGHT")
        Dim UltraGridColumn189 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLES")
        Dim UltraGridColumn190 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("UNITS")
        Dim UltraGridColumn195 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PPK_INNER_QTY_CALC")
        Dim UltraGridColumn191 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL_UNITS")
        Dim UltraGridColumn193 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE_1")
        Dim UltraGridColumn194 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE_1")
        Dim UltraGridColumn192 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_CODE")
        Dim UltraGridColumn544 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CBM")
        Dim UltraGridColumn383 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL_WEIGHT", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
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
        Dim UltraGridBand4 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIP8", -1)
        Dim UltraGridColumn196 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn197 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_LNO")
        Dim UltraGridColumn198 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_NO")
        Dim UltraGridColumn199 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn200 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn201 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY")
        Dim UltraGridColumn202 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DOZENS")
        Dim UltraGridColumn203 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PPK_INNER_QTY")
        Dim UltraGridColumn204 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("UNITS")
        Dim UltraGridColumn228 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTONS")
        Dim UltraGridColumn229 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL_UNITS")
        Dim UltraGridColumn163 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CBM")
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
        Dim UltraGridBand5 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTLPNL1", -1)
        Dim UltraGridColumn16 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("BARCODE")
        Dim UltraGridColumn26 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE")
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_REFERENCE")
        Dim UltraGridColumn18 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_LIST_DESC")
        Dim UltraGridColumn99 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_LIST_NO")
        Dim UltraGridColumn24 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_LIST_SHEET_NO")
        Dim UltraGridColumn25 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PACK_LIST_SHEET_LNO")
        Dim UltraGridColumn19 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("BARCODE_STATUS")
        Dim UltraGridColumn86 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn87 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_ID")
        Dim UltraGridColumn88 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SIZE_CODE")
        Dim UltraGridColumn89 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_REFERENCE2")
        Dim UltraGridColumn90 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE_PFX")
        Dim UltraGridColumn91 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE_PFX2")
        Dim UltraGridColumn92 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_CONF")
        Dim UltraGridColumn93 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn94 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn95 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn96 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn111 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_GRS_WGT")
        Dim UltraGridColumn112 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_NET_WGT")
        Dim UltraGridColumn113 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_DIMENSIONS")
        Dim UltraGridColumn114 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_PACK")
        Dim UltraGridColumn115 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PKG_CODE")
        Dim UltraGridColumn97 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CONF_SHP")
        Dim UltraGridColumn98 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CONF_REM")
        Dim UltraGridColumn100 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CONF_UNK")
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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTSHIPX", -1)
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_NO")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_VESSEL")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_ETA")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_LANDING_LEAD_DAYS")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_REF_NO")
        Dim UltraGridColumn63 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_ADV_DATE")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_SHIPPED")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PORT_CODE")
        Dim UltraGridColumn62 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_FREIGHT_IN")
        Dim UltraGridColumn15 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_TRUCKING")
        Dim UltraGridColumn116 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_DUTY")
        Dim UltraGridColumn117 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_CUSTOMS")
        Dim UltraGridColumn118 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_IND")
        Dim UltraGridColumn119 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FREIGHT_ENTERED_BY")
        Dim UltraGridColumn20 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_NOTES")
        Dim UltraGridColumn21 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REVIEW")
        Dim UltraGridColumn22 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("AIR_SHIP")
        Dim UltraGridColumn23 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_COMPLETE")
        Dim UltraGridColumn64 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LP_STATUS")
        Dim UltraGridColumn67 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LP_XNO")
        Dim UltraGridColumn50 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PORT_CODE_ORIG")
        Dim UltraGridColumn51 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PORT_CODE_DEST")
        Dim UltraGridColumn52 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_CODE")
        Dim UltraGridColumn53 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_COMPLETE_OPS_YYYYPP")
        Dim UltraGridColumn69 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_COMPLETE_INIT_OPER")
        Dim UltraGridColumn71 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_COMPLETE_INIT_DATE")
        Dim UltraGridColumn56 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_FRT_METHOD")
        Dim UltraGridColumn57 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VOYAGE_NO")
        Dim UltraGridColumn58 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUSTOMS_DUTY_AMT")
        Dim UltraGridColumn59 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUSTOMS_ENTRY_NO")
        Dim UltraGridColumn60 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_NO_DUTY")
        Dim UltraGridColumn120 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CONTAINER_NO")
        Dim UltraGridColumn73 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_STATUS")
        Dim UltraGridColumn75 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIPMENT_LNO")
        Dim UltraGridColumn76 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("BOL_NO")
        Dim UltraGridColumn121 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_CTNS")
        Dim UltraGridColumn61 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CBM")
        Dim UltraGridColumn36 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CONTAINER_TYPE_CODE")
        Dim UltraGridColumn37 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CONTAINER_SEAL_NO")
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
        Dim UltraTab6 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab7 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab8 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim Appearance61 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand6 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("POTORDRX", -1)
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_ORDER_NO")
        Dim UltraGridColumn126 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VEND_CODE")
        Dim UltraGridColumn27 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VEND_NAME")
        Dim UltraGridColumn41 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn102 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn47 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn48 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_ORDERED")
        Dim UltraGridColumn49 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_REFERENCE")
        Dim UltraGridColumn54 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn55 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_STATUS")
        Dim UltraGridColumn81 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_CANCELLED")
        Dim UltraGridColumn82 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_SHIP_BY")
        Dim UltraGridColumn83 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_ETA")
        Dim UltraGridColumn84 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SPEC_ORDR_NO")
        Dim UltraGridColumn85 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FOB_CMT")
        Dim UltraGridColumn127 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FACTORY_CODE")
        Dim UltraGridColumn128 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CONTACT")
        Dim UltraGridColumn103 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_HDR_CTR_REV")
        Dim UltraGridColumn281 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_NOTES")
        Dim UltraGridColumn282 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_XMIT_IND")
        Dim UltraGridColumn283 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_XMIT_BY")
        Dim UltraGridColumn284 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_XMIT_DATE")
        Dim UltraGridColumn104 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_XMIT_XNO")
        Dim UltraGridColumn105 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PORT_CODE_ORIG")
        Dim UltraGridColumn287 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PORT_CODE_DEST")
        Dim UltraGridColumn288 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COST_CODE")
        Dim UltraGridColumn106 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_CANCEL")
        Dim UltraGridColumn289 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_FOB_DESC")
        Dim UltraGridColumn290 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_SHIP_VIA")
        Dim UltraGridColumn291 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CARTON_MARKS")
        Dim UltraGridColumn129 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn292 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TERM_CODE")
        Dim UltraGridColumn293 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_PRINTED")
        Dim UltraGridColumn294 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_PRINTED_IND")
        Dim UltraGridColumn295 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_MESSAGE")
        Dim UltraGridColumn275 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn296 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_REVISION_NOTE")
        Dim UltraGridColumn297 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_BATCH_NO")
        Dim UltraGridColumn227 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LABEL_RESP_CODE")
        Dim UltraGridColumn298 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_APPR_NOTES")
        Dim UltraGridColumn299 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_APPR_BY")
        Dim UltraGridColumn300 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_APPR_DATE")
        Dim UltraGridColumn301 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_APPR_PENDING")
        Dim UltraGridColumn302 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_APPR_AMOUNT")
        Dim UltraGridColumn303 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COMM_PAYABLE_TO_BRKR")
        Dim UltraGridColumn304 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COMM_CHGBACK_TO_SUPP")
        Dim UltraGridColumn305 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_COMM_PCT")
        Dim UltraGridColumn306 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_HAS_PPK")
        Dim UltraGridColumn307 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_APPR_QUEUE")
        Dim UltraGridColumn109 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME")
        Dim UltraGridColumn110 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CANCEL_DATE")
        Dim UltraGridColumn130 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_ORD")
        Dim UltraGridColumn131 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_SHP")
        Dim UltraGridColumn132 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_REC")
        Dim UltraGridColumn133 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_QTY_OPN")
        Dim UltraGridColumn134 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_AMT_ORD")
        Dim UltraGridColumn135 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_AMT_SHP")
        Dim UltraGridColumn136 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_AMT_REC")
        Dim UltraGridColumn137 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_AMT_OPN")
        Dim UltraGridColumn107 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_LINES_CONF")
        Dim UltraGridColumn108 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_LINES")
        Dim UltraGridColumn265 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CTNS_ORD")
        Dim UltraGridColumn267 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CTNS_SHP")
        Dim UltraGridColumn268 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CTNS_OPN")
        Dim UltraGridColumn266 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CUBE_ORD")
        Dim UltraGridColumn269 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CUBE_SHP")
        Dim UltraGridColumn270 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_CUBE_OPN")
        Dim UltraGridColumn271 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_SHIP_BY_MIN")
        Dim UltraGridColumn272 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_ETA_MIN")
        Dim UltraGridColumn273 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_SHIP_BY_MAX")
        Dim UltraGridColumn274 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PO_DATE_ETA_MAX")
        Dim Appearance62 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance63 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance64 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance65 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance66 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance67 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance68 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance69 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance70 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance71 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance72 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance73 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand7 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ICTSTYL1_RECENT", -1)
        Dim UltraGridColumn28 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn351 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_STATUS")
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_DESC")
        Dim UltraGridColumn276 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_WEIGHT")
        Dim UltraGridColumn277 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FABRIC_CODE")
        Dim UltraGridColumn278 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SEASON_CODE")
        Dim UltraGridColumn279 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SUB_BODY_CODE")
        Dim UltraGridColumn280 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SALES_DIVISION_CODE")
        Dim UltraGridColumn285 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INNER_PACK_QTY")
        Dim UltraGridColumn286 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTON_PACK_QTY")
        Dim UltraGridColumn331 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn332 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FASHION_PROMO")
        Dim UltraGridColumn333 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CMT_NO")
        Dim UltraGridColumn334 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_UOM")
        Dim UltraGridColumn335 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_PRICE")
        Dim UltraGridColumn336 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_COST")
        Dim UltraGridColumn337 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SUB_UNIT_PACK_QTY")
        Dim UltraGridColumn338 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn339 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn340 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn342 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn343 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_RETAIL")
        Dim UltraGridColumn410 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DUTY_RATE_CODE")
        Dim UltraGridColumn411 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WEIGHT_CODE")
        Dim UltraGridColumn412 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SIZE_SCALE")
        Dim UltraGridColumn413 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IMAGE_NAME")
        Dim UltraGridColumn414 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REPLENISHMENT")
        Dim UltraGridColumn415 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CLASS_CODE")
        Dim UltraGridColumn416 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CASE_CUBE")
        Dim UltraGridColumn417 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_PO_QTY_MIN")
        Dim UltraGridColumn418 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PURCH_NOTES")
        Dim UltraGridColumn419 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FACTORY_CODE")
        Dim UltraGridColumn420 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_DESC2")
        Dim UltraGridColumn421 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_SO_QTY_MIN")
        Dim UltraGridColumn422 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE_ORIG")
        Dim UltraGridColumn423 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SUB_UNIT_BAG_QTY")
        Dim UltraGridColumn424 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LABEL_TYPE_CODE")
        Dim UltraGridColumn425 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_XMAS_DATE")
        Dim UltraGridColumn426 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_UPC_ON_PO")
        Dim UltraGridColumn427 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_COST_FOB_CF")
        Dim UltraGridColumn428 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("VEND_CODE")
        Dim UltraGridColumn429 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_MATL_DESC")
        Dim UltraGridColumn430 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CASE_WEIGHT_GRS")
        Dim UltraGridColumn431 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_GROUP_CODE")
        Dim UltraGridColumn432 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COUNTRY_CODE")
        Dim UltraGridColumn433 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ROYALTY_CODE")
        Dim UltraGridColumn434 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_PROMO_PRICE")
        Dim UltraGridColumn435 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SIZE_CODE")
        Dim UltraGridColumn436 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DEPT_CODE")
        Dim UltraGridColumn219 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_HIDE_FROM_CAT")
        Dim UltraGridColumn220 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_HIDE_FROM_3PL")
        Dim UltraGridColumn254 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_CODE_PLM")
        Dim UltraGridColumn39 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CARTONS_PER_UNIT")
        Dim UltraGridColumn40 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RESHIPBOX_CODE")
        Dim UltraGridColumn498 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("REQUIRES_EXP_DATE")
        Dim UltraGridColumn437 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("MASTER_BODY_CODE")
        Dim UltraGridColumn42 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_ORDR_DATE")
        Dim UltraGridColumn43 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_ORDR_NO")
        Dim UltraGridColumn44 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_ORDR_CUST_CODE")
        Dim UltraGridColumn45 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_ORDR_CUST_PO")
        Dim UltraGridColumn438 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_ONHD")
        Dim UltraGridColumn439 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_ONPO")
        Dim UltraGridColumn440 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_TRAN")
        Dim UltraGridColumn441 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_OPEN")
        Dim UltraGridColumn442 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_PICK")
        Dim UltraGridColumn443 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_COMM")
        Dim UltraGridColumn444 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_PROD")
        Dim UltraGridColumn445 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("QTY_NETA")
        Dim UltraGridColumn446 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_CODE")
        Dim UltraGridColumn447 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLOR_DESC")
        Dim UltraGridColumn448 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn449 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_DESC")
        Dim UltraGridColumn450 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("IMAGE")
        Dim UltraGridColumn499 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OPNQTY")
        Dim UltraGridColumn500 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OPNAMT")
        Dim UltraGridColumn501 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OPNCST")
        Dim UltraGridColumn502 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHPQTY")
        Dim UltraGridColumn503 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHPAMT")
        Dim UltraGridColumn504 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHPCST")
        Dim UltraGridColumn460 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_COST_LDP")
        Dim UltraGridColumn461 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_COST_LDP_CODE")
        Dim UltraGridColumn462 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_COST_ELC")
        Dim UltraGridColumn463 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_COST_CUM")
        Dim UltraGridColumn451 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STYLE_COST_EXT")
        Dim Appearance74 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance75 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance76 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance77 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance78 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance79 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance80 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance81 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance82 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance83 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance84 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraTab3 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab4 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab5 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab1 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab2 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Me.UltraTabPageControl6 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdPOTSHIP3 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl7 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.splPOTSHIP7 = New System.Windows.Forms.SplitContainer()
        Me.grdPOTSHIP7 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdPOTSHIP8 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl8 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grdPOTLPNL1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl3 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.splPOTSHIPX = New System.Windows.Forms.SplitContainer()
        Me.grdPOTSHIPX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.tabShipment = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage3 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.UltraTabPageControl4 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdPOTORDRX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl5 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdICTSTYL1_Recent = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl1 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.tabMain = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage2 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.UltraTabPageControl2 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.spl = New System.Windows.Forms.SplitContainer()
        Me.tab = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage1 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl6.SuspendLayout()
        CType(Me.grdPOTSHIP3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl7.SuspendLayout()
        CType(Me.splPOTSHIP7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splPOTSHIP7.Panel1.SuspendLayout()
        Me.splPOTSHIP7.Panel2.SuspendLayout()
        Me.splPOTSHIP7.SuspendLayout()
        CType(Me.grdPOTSHIP7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdPOTSHIP8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl8.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grdPOTLPNL1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl3.SuspendLayout()
        CType(Me.splPOTSHIPX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splPOTSHIPX.Panel1.SuspendLayout()
        Me.splPOTSHIPX.Panel2.SuspendLayout()
        Me.splPOTSHIPX.SuspendLayout()
        CType(Me.grdPOTSHIPX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tabShipment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabShipment.SuspendLayout()
        Me.UltraTabPageControl4.SuspendLayout()
        CType(Me.grdPOTORDRX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl5.SuspendLayout()
        CType(Me.grdICTSTYL1_Recent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl1.SuspendLayout()
        CType(Me.tabMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabMain.SuspendLayout()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.spl.Panel1.SuspendLayout()
        Me.spl.Panel2.SuspendLayout()
        Me.spl.SuspendLayout()
        CType(Me.tab, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraExplorerBar1
        '
        UltraExplorerBarItem1.Text = "Refresh"
        UltraExplorerBarItem2.Text = "Print"
        UltraExplorerBarGroup1.Items.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem() {UltraExplorerBarItem1, UltraExplorerBarItem2})
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
        Me.UltraExplorerBar1.Tag = "CLICK"
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.spl)
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(945, 574)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.spl, 0)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.grdASFBASEX, 0)
        '
        'grdASFBASEX
        '
        Appearance85.BackColor = System.Drawing.SystemColors.Window
        Appearance85.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance85
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance86.BackColor = System.Drawing.SystemColors.Window
        Appearance86.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance86
        Appearance87.BackColor = System.Drawing.SystemColors.Highlight
        Appearance87.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance87
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance88.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance88
        Appearance89.BorderColor = System.Drawing.Color.Silver
        Appearance89.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance89
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance90.BackColor = System.Drawing.SystemColors.Control
        Appearance90.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance90.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance90.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance90.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance90
        Appearance91.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance91
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance92.BackColor = System.Drawing.SystemColors.Window
        Appearance92.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance92
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance93.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance93
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
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(1158, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 574)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(1158, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 574)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(1158, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraTabPageControl6
        '
        Me.UltraTabPageControl6.Controls.Add(Me.grdPOTSHIP3)
        Me.UltraTabPageControl6.Location = New System.Drawing.Point(1, 25)
        Me.UltraTabPageControl6.Name = "UltraTabPageControl6"
        Me.UltraTabPageControl6.Size = New System.Drawing.Size(933, 254)
        '
        'grdPOTSHIP3
        '
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Appearance13.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIP3.DisplayLayout.Appearance = Appearance13
        UltraGridColumn29.Header.VisiblePosition = 0
        UltraGridColumn29.Hidden = True
        UltraGridColumn29.Width = 14
        UltraGridColumn30.Header.VisiblePosition = 1
        UltraGridColumn30.Hidden = True
        UltraGridColumn30.Width = 14
        UltraGridColumn31.Header.Caption = "POrder No"
        UltraGridColumn31.Header.VisiblePosition = 2
        UltraGridColumn31.Width = 87
        UltraGridColumn32.Header.Caption = "Ln"
        UltraGridColumn32.Header.VisiblePosition = 3
        UltraGridColumn32.Width = 35
        UltraGridColumn33.Header.Caption = "Qty Shp"
        UltraGridColumn33.Header.VisiblePosition = 7
        UltraGridColumn33.Width = 79
        UltraGridColumn65.Header.Caption = "Style"
        UltraGridColumn65.Header.VisiblePosition = 4
        UltraGridColumn65.Width = 94
        UltraGridColumn66.Header.Caption = "Color"
        UltraGridColumn66.Header.VisiblePosition = 5
        UltraGridColumn66.Width = 46
        UltraGridColumn68.Header.Caption = "UM"
        UltraGridColumn68.Header.VisiblePosition = 8
        UltraGridColumn68.Width = 51
        UltraGridColumn34.Header.VisiblePosition = 9
        UltraGridColumn34.Hidden = True
        UltraGridColumn35.Header.VisiblePosition = 10
        UltraGridColumn35.Hidden = True
        UltraGridColumn70.Header.Caption = "Description"
        UltraGridColumn70.Header.VisiblePosition = 6
        UltraGridColumn70.Width = 212
        UltraGridColumn72.Header.Caption = "Pack"
        UltraGridColumn72.Header.VisiblePosition = 11
        UltraGridColumn72.Width = 46
        UltraGridColumn38.Header.Caption = "#/Ctn"
        UltraGridColumn38.Header.VisiblePosition = 13
        UltraGridColumn38.Width = 57
        UltraGridColumn122.Header.Caption = "Cube/Cs"
        UltraGridColumn122.Header.VisiblePosition = 16
        UltraGridColumn122.Width = 72
        UltraGridColumn123.Header.VisiblePosition = 17
        UltraGridColumn123.Hidden = True
        UltraGridColumn77.Header.VisiblePosition = 18
        UltraGridColumn77.Hidden = True
        UltraGridColumn78.Header.VisiblePosition = 19
        UltraGridColumn78.Hidden = True
        UltraGridColumn124.Header.VisiblePosition = 20
        UltraGridColumn124.Hidden = True
        UltraGridColumn125.Header.VisiblePosition = 21
        UltraGridColumn125.Hidden = True
        UltraGridColumn101.Header.VisiblePosition = 22
        UltraGridColumn101.Hidden = True
        UltraGridColumn74.Format = "#,##0"
        UltraGridColumn74.Header.Caption = "Units"
        UltraGridColumn74.Header.VisiblePosition = 12
        UltraGridColumn74.Width = 94
        UltraGridColumn79.Header.Caption = "Carton Nos"
        UltraGridColumn79.Header.VisiblePosition = 15
        UltraGridColumn79.Width = 86
        UltraGridColumn80.Header.Caption = "Ctns"
        UltraGridColumn80.Header.VisiblePosition = 14
        UltraGridColumn80.Width = 50
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn29, UltraGridColumn30, UltraGridColumn31, UltraGridColumn32, UltraGridColumn33, UltraGridColumn65, UltraGridColumn66, UltraGridColumn68, UltraGridColumn34, UltraGridColumn35, UltraGridColumn70, UltraGridColumn72, UltraGridColumn38, UltraGridColumn122, UltraGridColumn123, UltraGridColumn77, UltraGridColumn78, UltraGridColumn124, UltraGridColumn125, UltraGridColumn101, UltraGridColumn74, UltraGridColumn79, UltraGridColumn80})
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
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIP3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
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
        Me.grdPOTSHIP3.Name = "grdPOTSHIP3"
        Me.grdPOTSHIP3.Size = New System.Drawing.Size(933, 254)
        Me.grdPOTSHIP3.TabIndex = 12
        Me.grdPOTSHIP3.Text = "Shipments"
        '
        'UltraTabPageControl7
        '
        Me.UltraTabPageControl7.Controls.Add(Me.splPOTSHIP7)
        Me.UltraTabPageControl7.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl7.Name = "UltraTabPageControl7"
        Me.UltraTabPageControl7.Size = New System.Drawing.Size(933, 254)
        '
        'splPOTSHIP7
        '
        Me.splPOTSHIP7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splPOTSHIP7.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.splPOTSHIP7.Location = New System.Drawing.Point(0, 0)
        Me.splPOTSHIP7.Name = "splPOTSHIP7"
        '
        'splPOTSHIP7.Panel1
        '
        Me.splPOTSHIP7.Panel1.Controls.Add(Me.grdPOTSHIP7)
        '
        'splPOTSHIP7.Panel2
        '
        Me.splPOTSHIP7.Panel2.Controls.Add(Me.grdPOTSHIP8)
        Me.splPOTSHIP7.Size = New System.Drawing.Size(933, 254)
        Me.splPOTSHIP7.SplitterDistance = 827
        Me.splPOTSHIP7.TabIndex = 14
        '
        'grdPOTSHIP7
        '
        Appearance25.BackColor = System.Drawing.SystemColors.Window
        Appearance25.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIP7.DisplayLayout.Appearance = Appearance25
        UltraGridColumn178.Header.VisiblePosition = 0
        UltraGridColumn178.Hidden = True
        UltraGridColumn179.Header.Caption = "Ln"
        UltraGridColumn179.Header.VisiblePosition = 1
        UltraGridColumn179.Hidden = True
        UltraGridColumn179.Width = 36
        UltraGridColumn180.Header.Caption = "#"
        UltraGridColumn180.Header.VisiblePosition = 2
        UltraGridColumn180.Width = 25
        UltraGridColumn181.Header.Caption = "Ctns"
        UltraGridColumn181.Header.VisiblePosition = 3
        UltraGridColumn181.Width = 46
        UltraGridColumn182.Header.Caption = "Notes"
        UltraGridColumn182.Header.VisiblePosition = 4
        UltraGridColumn182.Width = 75
        UltraGridColumn183.Header.Caption = "Custom"
        UltraGridColumn183.Header.VisiblePosition = 5
        UltraGridColumn183.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn183.Width = 67
        UltraGridColumn184.Header.Caption = "Pre-Pack Code"
        UltraGridColumn184.Header.VisiblePosition = 6
        UltraGridColumn184.Hidden = True
        UltraGridColumn185.Header.VisiblePosition = 8
        UltraGridColumn185.Hidden = True
        UltraGridColumn186.Header.VisiblePosition = 11
        UltraGridColumn186.Hidden = True
        UltraGridColumn187.Header.VisiblePosition = 19
        UltraGridColumn187.Hidden = True
        UltraGridColumn188.Header.VisiblePosition = 16
        UltraGridColumn188.Hidden = True
        UltraGridColumn157.Header.Caption = "Ctn Dims"
        UltraGridColumn157.Header.VisiblePosition = 12
        UltraGridColumn157.Width = 75
        UltraGridColumn161.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn161.Header.Caption = "Ctn Vol"
        UltraGridColumn161.Header.VisiblePosition = 17
        UltraGridColumn161.Hidden = True
        UltraGridColumn161.Width = 75
        UltraGridColumn382.Header.Caption = "Ctn Wgt"
        UltraGridColumn382.Header.VisiblePosition = 13
        UltraGridColumn382.Width = 75
        UltraGridColumn189.Header.Caption = "Styles"
        UltraGridColumn189.Header.VisiblePosition = 9
        UltraGridColumn189.Width = 56
        UltraGridColumn190.Header.Caption = "Units"
        UltraGridColumn190.Header.VisiblePosition = 18
        UltraGridColumn190.Width = 60
        UltraGridColumn195.Header.Caption = "Inner"
        UltraGridColumn195.Header.VisiblePosition = 10
        UltraGridColumn195.Width = 49
        UltraGridColumn191.Header.Caption = "Total"
        UltraGridColumn191.Header.VisiblePosition = 20
        UltraGridColumn191.Width = 75
        UltraGridColumn193.Header.VisiblePosition = 21
        UltraGridColumn193.Hidden = True
        UltraGridColumn194.Header.VisiblePosition = 22
        UltraGridColumn194.Hidden = True
        UltraGridColumn192.Header.Caption = "Item Code"
        UltraGridColumn192.Header.VisiblePosition = 7
        UltraGridColumn192.Width = 109
        UltraGridColumn544.Header.VisiblePosition = 15
        UltraGridColumn544.Width = 65
        UltraGridColumn383.Header.Caption = "Tot Wgt"
        UltraGridColumn383.Header.VisiblePosition = 14
        UltraGridColumn383.Width = 80
        UltraGridBand3.Columns.AddRange(New Object() {UltraGridColumn178, UltraGridColumn179, UltraGridColumn180, UltraGridColumn181, UltraGridColumn182, UltraGridColumn183, UltraGridColumn184, UltraGridColumn185, UltraGridColumn186, UltraGridColumn187, UltraGridColumn188, UltraGridColumn157, UltraGridColumn161, UltraGridColumn382, UltraGridColumn189, UltraGridColumn190, UltraGridColumn195, UltraGridColumn191, UltraGridColumn193, UltraGridColumn194, UltraGridColumn192, UltraGridColumn544, UltraGridColumn383})
        Me.grdPOTSHIP7.DisplayLayout.BandsSerializer.Add(UltraGridBand3)
        Me.grdPOTSHIP7.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance26.TextHAlignAsString = "Left"
        Me.grdPOTSHIP7.DisplayLayout.CaptionAppearance = Appearance26
        Appearance27.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance27.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance27.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP7.DisplayLayout.GroupByBox.Appearance = Appearance27
        Appearance28.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP7.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance28
        Me.grdPOTSHIP7.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIP7.DisplayLayout.GroupByBox.Hidden = True
        Appearance29.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance29.BackColor2 = System.Drawing.SystemColors.Control
        Appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance29.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP7.DisplayLayout.GroupByBox.PromptAppearance = Appearance29
        Me.grdPOTSHIP7.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIP7.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdPOTSHIP7.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdPOTSHIP7.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance30.BackColor = System.Drawing.SystemColors.Window
        Appearance30.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIP7.DisplayLayout.Override.ActiveCellAppearance = Appearance30
        Me.grdPOTSHIP7.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIP7.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP7.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP7.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIP7.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance31.BackColor = System.Drawing.Color.Transparent
        Me.grdPOTSHIP7.DisplayLayout.Override.CardAreaAppearance = Appearance31
        Appearance32.BorderColor = System.Drawing.Color.Silver
        Appearance32.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIP7.DisplayLayout.Override.CellAppearance = Appearance32
        Me.grdPOTSHIP7.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
        Me.grdPOTSHIP7.DisplayLayout.Override.CellPadding = 0
        Appearance33.BackColor = System.Drawing.SystemColors.Control
        Appearance33.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance33.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance33.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP7.DisplayLayout.Override.GroupByRowAppearance = Appearance33
        Appearance34.TextHAlignAsString = "Left"
        Me.grdPOTSHIP7.DisplayLayout.Override.HeaderAppearance = Appearance34
        Me.grdPOTSHIP7.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIP7.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance35.BackColor = System.Drawing.SystemColors.Window
        Appearance35.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIP7.DisplayLayout.Override.RowAppearance = Appearance35
        Me.grdPOTSHIP7.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance36.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIP7.DisplayLayout.Override.TemplateAddRowAppearance = Appearance36
        Me.grdPOTSHIP7.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIP7.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIP7.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIP7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIP7.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIP7.Name = "grdPOTSHIP7"
        Me.grdPOTSHIP7.Size = New System.Drawing.Size(827, 254)
        Me.grdPOTSHIP7.TabIndex = 12
        Me.grdPOTSHIP7.Text = "Carton Types"
        '
        'grdPOTSHIP8
        '
        Appearance37.BackColor = System.Drawing.SystemColors.Window
        Appearance37.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIP8.DisplayLayout.Appearance = Appearance37
        UltraGridColumn196.Header.VisiblePosition = 0
        UltraGridColumn196.Hidden = True
        UltraGridColumn197.Header.VisiblePosition = 1
        UltraGridColumn197.Hidden = True
        UltraGridColumn198.Header.Caption = "#"
        UltraGridColumn198.Header.VisiblePosition = 2
        UltraGridColumn198.Hidden = True
        UltraGridColumn198.Width = 25
        UltraGridColumn199.Header.Caption = "Style"
        UltraGridColumn199.Header.VisiblePosition = 3
        UltraGridColumn199.Width = 95
        UltraGridColumn200.Header.Caption = "Clr"
        UltraGridColumn200.Header.VisiblePosition = 4
        UltraGridColumn200.Width = 40
        UltraGridColumn201.Header.Caption = "Qty"
        UltraGridColumn201.Header.VisiblePosition = 5
        UltraGridColumn201.Width = 60
        UltraGridColumn202.Header.Caption = "Dz"
        UltraGridColumn202.Header.VisiblePosition = 6
        UltraGridColumn202.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn202.Width = 31
        UltraGridColumn203.Header.Caption = "Inner"
        UltraGridColumn203.Header.VisiblePosition = 7
        UltraGridColumn203.Width = 52
        UltraGridColumn204.Header.Caption = "Units"
        UltraGridColumn204.Header.VisiblePosition = 8
        UltraGridColumn204.Width = 60
        UltraGridColumn228.Header.VisiblePosition = 9
        UltraGridColumn228.Hidden = True
        UltraGridColumn229.Header.Caption = "Total"
        UltraGridColumn229.Header.VisiblePosition = 10
        UltraGridColumn229.Width = 75
        UltraGridColumn163.Header.VisiblePosition = 11
        UltraGridColumn163.Width = 64
        UltraGridBand4.Columns.AddRange(New Object() {UltraGridColumn196, UltraGridColumn197, UltraGridColumn198, UltraGridColumn199, UltraGridColumn200, UltraGridColumn201, UltraGridColumn202, UltraGridColumn203, UltraGridColumn204, UltraGridColumn228, UltraGridColumn229, UltraGridColumn163})
        Me.grdPOTSHIP8.DisplayLayout.BandsSerializer.Add(UltraGridBand4)
        Me.grdPOTSHIP8.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance38.TextHAlignAsString = "Left"
        Me.grdPOTSHIP8.DisplayLayout.CaptionAppearance = Appearance38
        Appearance39.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance39.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance39.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP8.DisplayLayout.GroupByBox.Appearance = Appearance39
        Appearance40.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP8.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance40
        Me.grdPOTSHIP8.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIP8.DisplayLayout.GroupByBox.Hidden = True
        Appearance41.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance41.BackColor2 = System.Drawing.SystemColors.Control
        Appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance41.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIP8.DisplayLayout.GroupByBox.PromptAppearance = Appearance41
        Me.grdPOTSHIP8.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIP8.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdPOTSHIP8.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance42.BackColor = System.Drawing.SystemColors.Window
        Appearance42.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIP8.DisplayLayout.Override.ActiveCellAppearance = Appearance42
        Me.grdPOTSHIP8.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIP8.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP8.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdPOTSHIP8.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIP8.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance43.BackColor = System.Drawing.Color.Transparent
        Me.grdPOTSHIP8.DisplayLayout.Override.CardAreaAppearance = Appearance43
        Appearance44.BorderColor = System.Drawing.Color.Silver
        Appearance44.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIP8.DisplayLayout.Override.CellAppearance = Appearance44
        Me.grdPOTSHIP8.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
        Me.grdPOTSHIP8.DisplayLayout.Override.CellPadding = 0
        Appearance45.BackColor = System.Drawing.SystemColors.Control
        Appearance45.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance45.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance45.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIP8.DisplayLayout.Override.GroupByRowAppearance = Appearance45
        Appearance46.TextHAlignAsString = "Left"
        Me.grdPOTSHIP8.DisplayLayout.Override.HeaderAppearance = Appearance46
        Me.grdPOTSHIP8.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIP8.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance47.BackColor = System.Drawing.SystemColors.Window
        Appearance47.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIP8.DisplayLayout.Override.RowAppearance = Appearance47
        Me.grdPOTSHIP8.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance48.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIP8.DisplayLayout.Override.TemplateAddRowAppearance = Appearance48
        Me.grdPOTSHIP8.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIP8.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIP8.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIP8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIP8.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIP8.Name = "grdPOTSHIP8"
        Me.grdPOTSHIP8.Size = New System.Drawing.Size(102, 254)
        Me.grdPOTSHIP8.TabIndex = 12
        Me.grdPOTSHIP8.Text = "Carton Configuration by Style/Color"
        '
        'UltraTabPageControl8
        '
        Me.UltraTabPageControl8.Controls.Add(Me.SplitContainer1)
        Me.UltraTabPageControl8.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl8.Name = "UltraTabPageControl8"
        Me.UltraTabPageControl8.Size = New System.Drawing.Size(933, 254)
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.grdPOTLPNL1)
        Me.SplitContainer1.Size = New System.Drawing.Size(933, 254)
        Me.SplitContainer1.SplitterDistance = 105
        Me.SplitContainer1.TabIndex = 25
        '
        'grdPOTLPNL1
        '
        Appearance49.BackColor = System.Drawing.SystemColors.Window
        Appearance49.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTLPNL1.DisplayLayout.Appearance = Appearance49
        UltraGridColumn16.Header.Caption = "LPN"
        UltraGridColumn16.Header.VisiblePosition = 0
        UltraGridColumn16.Width = 105
        UltraGridColumn26.Header.Caption = "Style Code"
        UltraGridColumn26.Header.VisiblePosition = 5
        UltraGridColumn26.Width = 120
        UltraGridColumn17.Header.VisiblePosition = 3
        UltraGridColumn17.Hidden = True
        UltraGridColumn18.Header.VisiblePosition = 6
        UltraGridColumn18.Hidden = True
        UltraGridColumn99.Header.VisiblePosition = 1
        UltraGridColumn99.Hidden = True
        UltraGridColumn24.Header.VisiblePosition = 2
        UltraGridColumn24.Hidden = True
        UltraGridColumn25.Header.Caption = "Ln"
        UltraGridColumn25.Header.VisiblePosition = 4
        UltraGridColumn25.Hidden = True
        UltraGridColumn25.Width = 40
        UltraGridColumn19.Header.Caption = "Status"
        UltraGridColumn19.Header.VisiblePosition = 7
        UltraGridColumn19.Hidden = True
        UltraGridColumn86.Header.Caption = "Color"
        UltraGridColumn86.Header.VisiblePosition = 8
        UltraGridColumn86.Width = 67
        UltraGridColumn87.Header.Caption = "Carton ID"
        UltraGridColumn87.Header.VisiblePosition = 9
        UltraGridColumn87.Width = 87
        UltraGridColumn88.Header.Caption = "Size"
        UltraGridColumn88.Header.VisiblePosition = 10
        UltraGridColumn88.Width = 78
        UltraGridColumn89.Header.VisiblePosition = 11
        UltraGridColumn89.Hidden = True
        UltraGridColumn90.Header.VisiblePosition = 12
        UltraGridColumn90.Hidden = True
        UltraGridColumn91.Header.VisiblePosition = 13
        UltraGridColumn91.Hidden = True
        UltraGridColumn92.Header.Caption = "Confirmation"
        UltraGridColumn92.Header.VisiblePosition = 14
        UltraGridColumn92.Width = 116
        UltraGridColumn93.Format = "MM/dd HH:mm"
        UltraGridColumn93.Header.Caption = "Created"
        UltraGridColumn93.Header.VisiblePosition = 15
        UltraGridColumn93.Hidden = True
        UltraGridColumn93.Width = 129
        UltraGridColumn94.Header.Caption = "By"
        UltraGridColumn94.Header.VisiblePosition = 16
        UltraGridColumn94.Hidden = True
        UltraGridColumn94.Width = 107
        UltraGridColumn95.Format = "MM/dd HH:mm"
        UltraGridColumn95.Header.Caption = "Confirmed"
        UltraGridColumn95.Header.VisiblePosition = 17
        UltraGridColumn95.Hidden = True
        UltraGridColumn95.Width = 125
        UltraGridColumn96.Header.Caption = "By"
        UltraGridColumn96.Header.VisiblePosition = 18
        UltraGridColumn96.Hidden = True
        UltraGridColumn96.Width = 98
        UltraGridColumn111.Header.Caption = "Grs Wgt"
        UltraGridColumn111.Header.VisiblePosition = 23
        UltraGridColumn111.Width = 60
        UltraGridColumn112.Header.Caption = "Net Wgt"
        UltraGridColumn112.Header.VisiblePosition = 24
        UltraGridColumn112.Width = 60
        UltraGridColumn113.Header.Caption = "Dimensions"
        UltraGridColumn113.Header.VisiblePosition = 25
        UltraGridColumn113.Width = 160
        UltraGridColumn114.Header.Caption = "Qty/Ctn"
        UltraGridColumn114.Header.VisiblePosition = 22
        UltraGridColumn114.Width = 65
        UltraGridColumn115.Header.VisiblePosition = 26
        UltraGridColumn115.Hidden = True
        UltraGridColumn97.Header.Caption = "To Ship"
        UltraGridColumn97.Header.VisiblePosition = 19
        UltraGridColumn97.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn97.Width = 100
        UltraGridColumn98.Header.Caption = "Remove"
        UltraGridColumn98.Header.VisiblePosition = 20
        UltraGridColumn98.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn98.Width = 100
        UltraGridColumn100.Header.Caption = "Unknown"
        UltraGridColumn100.Header.VisiblePosition = 21
        UltraGridColumn100.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn100.Width = 100
        UltraGridBand5.Columns.AddRange(New Object() {UltraGridColumn16, UltraGridColumn26, UltraGridColumn17, UltraGridColumn18, UltraGridColumn99, UltraGridColumn24, UltraGridColumn25, UltraGridColumn19, UltraGridColumn86, UltraGridColumn87, UltraGridColumn88, UltraGridColumn89, UltraGridColumn90, UltraGridColumn91, UltraGridColumn92, UltraGridColumn93, UltraGridColumn94, UltraGridColumn95, UltraGridColumn96, UltraGridColumn111, UltraGridColumn112, UltraGridColumn113, UltraGridColumn114, UltraGridColumn115, UltraGridColumn97, UltraGridColumn98, UltraGridColumn100})
        Me.grdPOTLPNL1.DisplayLayout.BandsSerializer.Add(UltraGridBand5)
        Me.grdPOTLPNL1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance50.TextHAlignAsString = "Left"
        Me.grdPOTLPNL1.DisplayLayout.CaptionAppearance = Appearance50
        Me.grdPOTLPNL1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance51.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance51.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance51.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTLPNL1.DisplayLayout.GroupByBox.Appearance = Appearance51
        Appearance52.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTLPNL1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance52
        Me.grdPOTLPNL1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTLPNL1.DisplayLayout.GroupByBox.Hidden = True
        Appearance53.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance53.BackColor2 = System.Drawing.SystemColors.Control
        Appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance53.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTLPNL1.DisplayLayout.GroupByBox.PromptAppearance = Appearance53
        Me.grdPOTLPNL1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTLPNL1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdPOTLPNL1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance54.BackColor = System.Drawing.SystemColors.Window
        Appearance54.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTLPNL1.DisplayLayout.Override.ActiveCellAppearance = Appearance54
        Me.grdPOTLPNL1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTLPNL1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTLPNL1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTLPNL1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTLPNL1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance55.BackColor = System.Drawing.Color.Transparent
        Me.grdPOTLPNL1.DisplayLayout.Override.CardAreaAppearance = Appearance55
        Appearance56.BorderColor = System.Drawing.Color.Silver
        Appearance56.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTLPNL1.DisplayLayout.Override.CellAppearance = Appearance56
        Me.grdPOTLPNL1.DisplayLayout.Override.CellPadding = 0
        Appearance57.BackColor = System.Drawing.SystemColors.Control
        Appearance57.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance57.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance57.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTLPNL1.DisplayLayout.Override.GroupByRowAppearance = Appearance57
        Appearance58.TextHAlignAsString = "Left"
        Me.grdPOTLPNL1.DisplayLayout.Override.HeaderAppearance = Appearance58
        Me.grdPOTLPNL1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTLPNL1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance59.BackColor = System.Drawing.SystemColors.Window
        Appearance59.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTLPNL1.DisplayLayout.Override.RowAppearance = Appearance59
        Me.grdPOTLPNL1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.grdPOTLPNL1.DisplayLayout.Override.RowSpacingAfter = 1
        Appearance60.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTLPNL1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance60
        Me.grdPOTLPNL1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTLPNL1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTLPNL1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTLPNL1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTLPNL1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdPOTLPNL1.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTLPNL1.Name = "grdPOTLPNL1"
        Me.grdPOTLPNL1.Size = New System.Drawing.Size(824, 254)
        Me.grdPOTLPNL1.TabIndex = 24
        Me.grdPOTLPNL1.Text = "LPNs"
        '
        'UltraTabPageControl3
        '
        Me.UltraTabPageControl3.Controls.Add(Me.splPOTSHIPX)
        Me.UltraTabPageControl3.Location = New System.Drawing.Point(1, 1)
        Me.UltraTabPageControl3.Name = "UltraTabPageControl3"
        Me.UltraTabPageControl3.Size = New System.Drawing.Size(937, 518)
        '
        'splPOTSHIPX
        '
        Me.splPOTSHIPX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splPOTSHIPX.Location = New System.Drawing.Point(0, 0)
        Me.splPOTSHIPX.Name = "splPOTSHIPX"
        Me.splPOTSHIPX.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splPOTSHIPX.Panel1
        '
        Me.splPOTSHIPX.Panel1.Controls.Add(Me.grdPOTSHIPX)
        '
        'splPOTSHIPX.Panel2
        '
        Me.splPOTSHIPX.Panel2.Controls.Add(Me.tabShipment)
        Me.splPOTSHIPX.Size = New System.Drawing.Size(937, 518)
        Me.splPOTSHIPX.SplitterDistance = 232
        Me.splPOTSHIPX.TabIndex = 12
        '
        'grdPOTSHIPX
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTSHIPX.DisplayLayout.Appearance = Appearance1
        UltraGridColumn9.Header.Caption = "Shipment No"
        UltraGridColumn9.Header.VisiblePosition = 0
        UltraGridColumn9.Width = 77
        UltraGridColumn2.Header.Caption = "Vessel"
        UltraGridColumn2.Header.VisiblePosition = 1
        UltraGridColumn2.Width = 138
        UltraGridColumn3.Header.Caption = "ETA"
        UltraGridColumn3.Header.VisiblePosition = 2
        UltraGridColumn3.Width = 94
        UltraGridColumn4.Header.VisiblePosition = 3
        UltraGridColumn4.Hidden = True
        UltraGridColumn4.Width = 96
        UltraGridColumn5.Header.Caption = "Shipper Ref No"
        UltraGridColumn5.Header.VisiblePosition = 4
        UltraGridColumn5.Width = 147
        UltraGridColumn63.Header.VisiblePosition = 5
        UltraGridColumn63.Hidden = True
        UltraGridColumn63.Width = 74
        UltraGridColumn7.Header.Caption = "Shipped"
        UltraGridColumn7.Header.VisiblePosition = 6
        UltraGridColumn7.Width = 92
        UltraGridColumn8.Header.Caption = "Port"
        UltraGridColumn8.Header.VisiblePosition = 7
        UltraGridColumn8.Width = 41
        UltraGridColumn62.Header.Caption = "Whse"
        UltraGridColumn62.Header.VisiblePosition = 8
        UltraGridColumn62.Width = 45
        UltraGridColumn10.Header.VisiblePosition = 9
        UltraGridColumn10.Hidden = True
        UltraGridColumn10.Width = 25
        UltraGridColumn11.Header.VisiblePosition = 10
        UltraGridColumn11.Hidden = True
        UltraGridColumn11.Width = 25
        UltraGridColumn12.Header.VisiblePosition = 11
        UltraGridColumn12.Hidden = True
        UltraGridColumn12.Width = 23
        UltraGridColumn13.Header.VisiblePosition = 12
        UltraGridColumn13.Hidden = True
        UltraGridColumn13.Width = 23
        UltraGridColumn14.Header.VisiblePosition = 13
        UltraGridColumn14.Hidden = True
        UltraGridColumn14.Width = 30
        UltraGridColumn15.Header.VisiblePosition = 14
        UltraGridColumn15.Hidden = True
        UltraGridColumn15.Width = 28
        UltraGridColumn116.Header.VisiblePosition = 15
        UltraGridColumn116.Hidden = True
        UltraGridColumn116.Width = 21
        UltraGridColumn117.Header.VisiblePosition = 16
        UltraGridColumn117.Hidden = True
        UltraGridColumn117.Width = 27
        UltraGridColumn118.Header.VisiblePosition = 17
        UltraGridColumn118.Hidden = True
        UltraGridColumn118.Width = 25
        UltraGridColumn119.Header.VisiblePosition = 18
        UltraGridColumn119.Hidden = True
        UltraGridColumn119.Width = 35
        UltraGridColumn20.Header.VisiblePosition = 19
        UltraGridColumn20.Hidden = True
        UltraGridColumn20.Width = 25
        UltraGridColumn21.Header.VisiblePosition = 20
        UltraGridColumn21.Hidden = True
        UltraGridColumn21.Width = 25
        UltraGridColumn22.Header.VisiblePosition = 21
        UltraGridColumn22.Hidden = True
        UltraGridColumn22.Width = 25
        UltraGridColumn23.Header.VisiblePosition = 22
        UltraGridColumn23.Hidden = True
        UltraGridColumn23.Width = 68
        UltraGridColumn64.Header.VisiblePosition = 23
        UltraGridColumn64.Hidden = True
        UltraGridColumn67.Header.VisiblePosition = 25
        UltraGridColumn67.Hidden = True
        UltraGridColumn50.Header.VisiblePosition = 27
        UltraGridColumn50.Hidden = True
        UltraGridColumn51.Header.VisiblePosition = 29
        UltraGridColumn51.Hidden = True
        UltraGridColumn52.Header.VisiblePosition = 31
        UltraGridColumn52.Hidden = True
        UltraGridColumn53.Header.VisiblePosition = 33
        UltraGridColumn53.Hidden = True
        UltraGridColumn69.Header.VisiblePosition = 35
        UltraGridColumn69.Hidden = True
        UltraGridColumn71.Header.VisiblePosition = 37
        UltraGridColumn71.Hidden = True
        UltraGridColumn56.Header.VisiblePosition = 38
        UltraGridColumn56.Hidden = True
        UltraGridColumn57.Header.VisiblePosition = 39
        UltraGridColumn57.Hidden = True
        UltraGridColumn58.Header.VisiblePosition = 40
        UltraGridColumn58.Hidden = True
        UltraGridColumn59.Header.VisiblePosition = 41
        UltraGridColumn59.Hidden = True
        UltraGridColumn60.Header.VisiblePosition = 43
        UltraGridColumn60.Hidden = True
        UltraGridColumn120.Header.Caption = "Container No"
        UltraGridColumn120.Header.VisiblePosition = 24
        UltraGridColumn120.Width = 130
        UltraGridColumn73.Header.VisiblePosition = 26
        UltraGridColumn73.Hidden = True
        UltraGridColumn73.Width = 74
        UltraGridColumn75.Header.VisiblePosition = 28
        UltraGridColumn75.Hidden = True
        UltraGridColumn75.Width = 82
        UltraGridColumn76.Header.Caption = "BOL No"
        UltraGridColumn76.Header.VisiblePosition = 34
        UltraGridColumn76.Width = 148
        UltraGridColumn121.Header.Caption = "Ctns"
        UltraGridColumn121.Header.VisiblePosition = 36
        UltraGridColumn121.Width = 61
        UltraGridColumn61.Header.VisiblePosition = 42
        UltraGridColumn61.Width = 80
        UltraGridColumn36.Header.Caption = "Type"
        UltraGridColumn36.Header.VisiblePosition = 30
        UltraGridColumn36.Width = 69
        UltraGridColumn37.Header.Caption = "Seal No"
        UltraGridColumn37.Header.VisiblePosition = 32
        UltraGridColumn37.Width = 90
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn9, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn63, UltraGridColumn7, UltraGridColumn8, UltraGridColumn62, UltraGridColumn10, UltraGridColumn11, UltraGridColumn12, UltraGridColumn13, UltraGridColumn14, UltraGridColumn15, UltraGridColumn116, UltraGridColumn117, UltraGridColumn118, UltraGridColumn119, UltraGridColumn20, UltraGridColumn21, UltraGridColumn22, UltraGridColumn23, UltraGridColumn64, UltraGridColumn67, UltraGridColumn50, UltraGridColumn51, UltraGridColumn52, UltraGridColumn53, UltraGridColumn69, UltraGridColumn71, UltraGridColumn56, UltraGridColumn57, UltraGridColumn58, UltraGridColumn59, UltraGridColumn60, UltraGridColumn120, UltraGridColumn73, UltraGridColumn75, UltraGridColumn76, UltraGridColumn121, UltraGridColumn61, UltraGridColumn36, UltraGridColumn37})
        Me.grdPOTSHIPX.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdPOTSHIPX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdPOTSHIPX.DisplayLayout.CaptionAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTSHIPX.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdPOTSHIPX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTSHIPX.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdPOTSHIPX.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTSHIPX.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPX.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTSHIPX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTSHIPX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.Color.Transparent
        Me.grdPOTSHIPX.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTSHIPX.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdPOTSHIPX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
        Me.grdPOTSHIPX.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTSHIPX.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTSHIPX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTSHIPX.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdPOTSHIPX.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTSHIPX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdPOTSHIPX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTSHIPX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTSHIPX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTSHIPX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTSHIPX.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTSHIPX.Name = "grdPOTSHIPX"
        Me.grdPOTSHIPX.Size = New System.Drawing.Size(937, 232)
        Me.grdPOTSHIPX.TabIndex = 11
        Me.grdPOTSHIPX.Text = "Shipments"
        '
        'tabShipment
        '
        Me.tabShipment.Controls.Add(Me.UltraTabSharedControlsPage3)
        Me.tabShipment.Controls.Add(Me.UltraTabPageControl6)
        Me.tabShipment.Controls.Add(Me.UltraTabPageControl7)
        Me.tabShipment.Controls.Add(Me.UltraTabPageControl8)
        Me.tabShipment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabShipment.Location = New System.Drawing.Point(0, 0)
        Me.tabShipment.Name = "tabShipment"
        Me.tabShipment.SharedControlsPage = Me.UltraTabSharedControlsPage3
        Me.tabShipment.Size = New System.Drawing.Size(937, 282)
        Me.tabShipment.TabIndex = 13
        UltraTab6.TabPage = Me.UltraTabPageControl6
        UltraTab6.Text = "PO Details"
        UltraTab7.TabPage = Me.UltraTabPageControl7
        UltraTab7.Text = "Cartons"
        UltraTab8.TabPage = Me.UltraTabPageControl8
        UltraTab8.Text = "LPNs"
        Me.tabShipment.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab6, UltraTab7, UltraTab8})
        '
        'UltraTabSharedControlsPage3
        '
        Me.UltraTabSharedControlsPage3.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage3.Name = "UltraTabSharedControlsPage3"
        Me.UltraTabSharedControlsPage3.Size = New System.Drawing.Size(933, 254)
        '
        'UltraTabPageControl4
        '
        Me.UltraTabPageControl4.Controls.Add(Me.grdPOTORDRX)
        Me.UltraTabPageControl4.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl4.Name = "UltraTabPageControl4"
        Me.UltraTabPageControl4.Size = New System.Drawing.Size(937, 518)
        '
        'grdPOTORDRX
        '
        Appearance61.BackColor = System.Drawing.SystemColors.Window
        Appearance61.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdPOTORDRX.DisplayLayout.Appearance = Appearance61
        UltraGridColumn6.Header.Caption = "PO No"
        UltraGridColumn6.Header.VisiblePosition = 0
        UltraGridColumn6.Width = 63
        UltraGridColumn126.Header.Caption = "Supplier"
        UltraGridColumn126.Header.VisiblePosition = 2
        UltraGridColumn126.Width = 83
        UltraGridColumn27.Header.Caption = "Name"
        UltraGridColumn27.Header.VisiblePosition = 3
        UltraGridColumn41.Header.Caption = "By"
        UltraGridColumn41.Header.VisiblePosition = 20
        UltraGridColumn41.Width = 75
        UltraGridColumn102.Header.Caption = "By"
        UltraGridColumn102.Header.VisiblePosition = 24
        UltraGridColumn102.Width = 75
        UltraGridColumn46.Format = "MM/dd/yyyy HH:mm"
        UltraGridColumn46.Header.Caption = "Entered"
        UltraGridColumn46.Header.VisiblePosition = 18
        UltraGridColumn46.Width = 150
        UltraGridColumn47.Format = "MM/dd/yyyy HH:mm"
        UltraGridColumn47.Header.Caption = "Modified"
        UltraGridColumn47.Header.VisiblePosition = 22
        UltraGridColumn47.Width = 150
        UltraGridColumn48.Header.Caption = "Order Date"
        UltraGridColumn48.Header.VisiblePosition = 4
        UltraGridColumn48.Width = 100
        UltraGridColumn49.Header.Caption = "PO Reference"
        UltraGridColumn49.Header.VisiblePosition = 1
        UltraGridColumn49.Width = 113
        UltraGridColumn54.Header.Caption = "Whse"
        UltraGridColumn54.Header.VisiblePosition = 16
        UltraGridColumn54.Width = 60
        UltraGridColumn55.Header.Caption = "Status"
        UltraGridColumn55.Header.VisiblePosition = 5
        UltraGridColumn55.Width = 70
        UltraGridColumn81.Header.Caption = "Cancelled"
        UltraGridColumn81.Header.VisiblePosition = 17
        UltraGridColumn81.Width = 100
        UltraGridColumn82.Header.Caption = "Ship By"
        UltraGridColumn82.Header.VisiblePosition = 6
        UltraGridColumn82.Width = 100
        UltraGridColumn83.Header.Caption = "ETA"
        UltraGridColumn83.Header.VisiblePosition = 8
        UltraGridColumn83.Width = 100
        UltraGridColumn84.Header.Caption = "Special Order"
        UltraGridColumn84.Header.VisiblePosition = 14
        UltraGridColumn85.Header.Caption = "Type"
        UltraGridColumn85.Header.VisiblePosition = 13
        UltraGridColumn85.Width = 66
        UltraGridColumn127.Header.Caption = "Factory"
        UltraGridColumn127.Header.VisiblePosition = 15
        UltraGridColumn127.Width = 77
        UltraGridColumn128.Header.Caption = "Contact"
        UltraGridColumn128.Header.VisiblePosition = 26
        UltraGridColumn103.Header.VisiblePosition = 19
        UltraGridColumn103.Hidden = True
        UltraGridColumn281.Header.VisiblePosition = 21
        UltraGridColumn281.Hidden = True
        UltraGridColumn282.Header.VisiblePosition = 23
        UltraGridColumn282.Hidden = True
        UltraGridColumn283.Header.VisiblePosition = 25
        UltraGridColumn283.Hidden = True
        UltraGridColumn284.Header.VisiblePosition = 27
        UltraGridColumn284.Hidden = True
        UltraGridColumn104.Header.VisiblePosition = 28
        UltraGridColumn104.Hidden = True
        UltraGridColumn105.Header.VisiblePosition = 30
        UltraGridColumn105.Hidden = True
        UltraGridColumn287.Header.VisiblePosition = 32
        UltraGridColumn287.Hidden = True
        UltraGridColumn288.Header.VisiblePosition = 34
        UltraGridColumn288.Hidden = True
        UltraGridColumn106.Header.Caption = "Cancel"
        UltraGridColumn106.Header.VisiblePosition = 7
        UltraGridColumn106.Width = 98
        UltraGridColumn289.Header.Caption = "FOB"
        UltraGridColumn289.Header.VisiblePosition = 45
        UltraGridColumn289.Width = 76
        UltraGridColumn290.Header.Caption = "Via"
        UltraGridColumn290.Header.VisiblePosition = 46
        UltraGridColumn290.Width = 62
        UltraGridColumn291.Header.Caption = "Ctn Marks"
        UltraGridColumn291.Header.VisiblePosition = 47
        UltraGridColumn291.Width = 85
        UltraGridColumn129.Header.Caption = "Order No"
        UltraGridColumn129.Header.VisiblePosition = 31
        UltraGridColumn129.Width = 117
        UltraGridColumn292.Header.VisiblePosition = 48
        UltraGridColumn292.Hidden = True
        UltraGridColumn293.Header.Caption = "Date Prt"
        UltraGridColumn293.Header.VisiblePosition = 49
        UltraGridColumn294.Header.Caption = "Prt"
        UltraGridColumn294.Header.VisiblePosition = 50
        UltraGridColumn294.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn294.Width = 39
        UltraGridColumn295.Header.VisiblePosition = 51
        UltraGridColumn295.Hidden = True
        UltraGridColumn275.Header.Caption = "Cust Code"
        UltraGridColumn275.Header.VisiblePosition = 33
        UltraGridColumn275.Width = 89
        UltraGridColumn296.Header.Caption = "Revision Note"
        UltraGridColumn296.Header.VisiblePosition = 53
        UltraGridColumn297.Header.VisiblePosition = 55
        UltraGridColumn297.Hidden = True
        UltraGridColumn227.Header.Caption = "Label Resp"
        UltraGridColumn227.Header.VisiblePosition = 29
        UltraGridColumn227.Width = 102
        UltraGridColumn298.Header.Caption = "Appr Notes"
        UltraGridColumn298.Header.VisiblePosition = 62
        UltraGridColumn299.Header.Caption = "Appr By"
        UltraGridColumn299.Header.VisiblePosition = 63
        UltraGridColumn299.Width = 83
        UltraGridColumn300.Header.Caption = "Appr Date"
        UltraGridColumn300.Header.VisiblePosition = 64
        UltraGridColumn301.Header.Caption = "Appr Pend"
        UltraGridColumn301.Header.VisiblePosition = 65
        UltraGridColumn301.Width = 87
        UltraGridColumn302.Header.Caption = "Appr Amt"
        UltraGridColumn302.Header.VisiblePosition = 66
        UltraGridColumn302.Width = 96
        UltraGridColumn303.Header.VisiblePosition = 67
        UltraGridColumn303.Hidden = True
        UltraGridColumn304.Header.VisiblePosition = 68
        UltraGridColumn304.Hidden = True
        UltraGridColumn305.Header.VisiblePosition = 69
        UltraGridColumn305.Hidden = True
        UltraGridColumn306.Header.VisiblePosition = 70
        UltraGridColumn306.Hidden = True
        UltraGridColumn307.Header.Caption = "Appr Q"
        UltraGridColumn307.Header.VisiblePosition = 71
        UltraGridColumn307.Width = 81
        UltraGridColumn109.Header.Caption = "Customer Name"
        UltraGridColumn109.Header.VisiblePosition = 35
        UltraGridColumn110.Header.Caption = "Order Cancel"
        UltraGridColumn110.Header.VisiblePosition = 36
        UltraGridColumn110.Width = 106
        UltraGridColumn130.Header.Caption = "Qty Ord"
        UltraGridColumn130.Header.VisiblePosition = 37
        UltraGridColumn130.Width = 100
        UltraGridColumn131.Header.Caption = "Qty Shp"
        UltraGridColumn131.Header.VisiblePosition = 39
        UltraGridColumn131.Width = 100
        UltraGridColumn132.Header.Caption = "Qty Rec"
        UltraGridColumn132.Header.VisiblePosition = 40
        UltraGridColumn132.Width = 100
        UltraGridColumn133.Header.Caption = "Qty Opn"
        UltraGridColumn133.Header.VisiblePosition = 38
        UltraGridColumn133.Width = 100
        UltraGridColumn134.Header.Caption = "Amt Ord"
        UltraGridColumn134.Header.VisiblePosition = 41
        UltraGridColumn134.Width = 100
        UltraGridColumn135.Header.Caption = "Amt Shp"
        UltraGridColumn135.Header.VisiblePosition = 42
        UltraGridColumn135.Width = 100
        UltraGridColumn136.Header.Caption = "Amt Rec"
        UltraGridColumn136.Header.VisiblePosition = 44
        UltraGridColumn136.Width = 100
        UltraGridColumn137.Header.Caption = "Amt Opn"
        UltraGridColumn137.Header.VisiblePosition = 43
        UltraGridColumn137.Width = 100
        UltraGridColumn107.Header.VisiblePosition = 52
        UltraGridColumn107.Hidden = True
        UltraGridColumn108.Header.Caption = "Lines"
        UltraGridColumn108.Header.VisiblePosition = 54
        UltraGridColumn108.Width = 50
        UltraGridColumn265.Header.Caption = "Ctns Ord"
        UltraGridColumn265.Header.VisiblePosition = 56
        UltraGridColumn265.Width = 85
        UltraGridColumn267.Header.Caption = "Ctns Shp"
        UltraGridColumn267.Header.VisiblePosition = 58
        UltraGridColumn267.Width = 85
        UltraGridColumn268.Header.Caption = "Ctns Opn"
        UltraGridColumn268.Header.VisiblePosition = 57
        UltraGridColumn268.Width = 85
        UltraGridColumn266.Header.Caption = "Cube Ord"
        UltraGridColumn266.Header.VisiblePosition = 59
        UltraGridColumn266.Width = 85
        UltraGridColumn269.Header.Caption = "Cube Shp"
        UltraGridColumn269.Header.VisiblePosition = 61
        UltraGridColumn269.Width = 85
        UltraGridColumn270.Header.Caption = "Cube Opn"
        UltraGridColumn270.Header.VisiblePosition = 60
        UltraGridColumn270.Width = 85
        UltraGridColumn271.Header.Caption = "Ship By Min"
        UltraGridColumn271.Header.VisiblePosition = 9
        UltraGridColumn271.Width = 100
        UltraGridColumn272.Header.Caption = "ETA Min"
        UltraGridColumn272.Header.VisiblePosition = 10
        UltraGridColumn272.Width = 100
        UltraGridColumn273.Header.Caption = "Ship By Max"
        UltraGridColumn273.Header.VisiblePosition = 11
        UltraGridColumn273.Width = 100
        UltraGridColumn274.Header.Caption = "ETA Max"
        UltraGridColumn274.Header.VisiblePosition = 12
        UltraGridColumn274.Width = 100
        UltraGridBand6.Columns.AddRange(New Object() {UltraGridColumn6, UltraGridColumn126, UltraGridColumn27, UltraGridColumn41, UltraGridColumn102, UltraGridColumn46, UltraGridColumn47, UltraGridColumn48, UltraGridColumn49, UltraGridColumn54, UltraGridColumn55, UltraGridColumn81, UltraGridColumn82, UltraGridColumn83, UltraGridColumn84, UltraGridColumn85, UltraGridColumn127, UltraGridColumn128, UltraGridColumn103, UltraGridColumn281, UltraGridColumn282, UltraGridColumn283, UltraGridColumn284, UltraGridColumn104, UltraGridColumn105, UltraGridColumn287, UltraGridColumn288, UltraGridColumn106, UltraGridColumn289, UltraGridColumn290, UltraGridColumn291, UltraGridColumn129, UltraGridColumn292, UltraGridColumn293, UltraGridColumn294, UltraGridColumn295, UltraGridColumn275, UltraGridColumn296, UltraGridColumn297, UltraGridColumn227, UltraGridColumn298, UltraGridColumn299, UltraGridColumn300, UltraGridColumn301, UltraGridColumn302, UltraGridColumn303, UltraGridColumn304, UltraGridColumn305, UltraGridColumn306, UltraGridColumn307, UltraGridColumn109, UltraGridColumn110, UltraGridColumn130, UltraGridColumn131, UltraGridColumn132, UltraGridColumn133, UltraGridColumn134, UltraGridColumn135, UltraGridColumn136, UltraGridColumn137, UltraGridColumn107, UltraGridColumn108, UltraGridColumn265, UltraGridColumn267, UltraGridColumn268, UltraGridColumn266, UltraGridColumn269, UltraGridColumn270, UltraGridColumn271, UltraGridColumn272, UltraGridColumn273, UltraGridColumn274})
        Me.grdPOTORDRX.DisplayLayout.BandsSerializer.Add(UltraGridBand6)
        Me.grdPOTORDRX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance62.TextHAlignAsString = "Left"
        Me.grdPOTORDRX.DisplayLayout.CaptionAppearance = Appearance62
        Appearance63.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance63.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance63.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance63.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTORDRX.DisplayLayout.GroupByBox.Appearance = Appearance63
        Appearance64.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTORDRX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance64
        Me.grdPOTORDRX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdPOTORDRX.DisplayLayout.GroupByBox.Hidden = True
        Appearance65.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance65.BackColor2 = System.Drawing.SystemColors.Control
        Appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance65.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdPOTORDRX.DisplayLayout.GroupByBox.PromptAppearance = Appearance65
        Me.grdPOTORDRX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdPOTORDRX.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdPOTORDRX.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance66.BackColor = System.Drawing.SystemColors.Window
        Appearance66.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdPOTORDRX.DisplayLayout.Override.ActiveCellAppearance = Appearance66
        Me.grdPOTORDRX.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdPOTORDRX.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTORDRX.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdPOTORDRX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdPOTORDRX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance67.BackColor = System.Drawing.Color.Transparent
        Me.grdPOTORDRX.DisplayLayout.Override.CardAreaAppearance = Appearance67
        Appearance68.BorderColor = System.Drawing.Color.Silver
        Appearance68.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdPOTORDRX.DisplayLayout.Override.CellAppearance = Appearance68
        Me.grdPOTORDRX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdPOTORDRX.DisplayLayout.Override.CellPadding = 0
        Appearance69.BackColor = System.Drawing.SystemColors.Control
        Appearance69.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance69.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance69.BorderColor = System.Drawing.SystemColors.Window
        Me.grdPOTORDRX.DisplayLayout.Override.GroupByRowAppearance = Appearance69
        Appearance70.TextHAlignAsString = "Left"
        Me.grdPOTORDRX.DisplayLayout.Override.HeaderAppearance = Appearance70
        Me.grdPOTORDRX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdPOTORDRX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance71.BackColor = System.Drawing.SystemColors.Window
        Appearance71.BorderColor = System.Drawing.Color.Silver
        Me.grdPOTORDRX.DisplayLayout.Override.RowAppearance = Appearance71
        Me.grdPOTORDRX.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance72.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdPOTORDRX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance72
        Me.grdPOTORDRX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdPOTORDRX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdPOTORDRX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdPOTORDRX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPOTORDRX.Location = New System.Drawing.Point(0, 0)
        Me.grdPOTORDRX.Name = "grdPOTORDRX"
        Me.grdPOTORDRX.Size = New System.Drawing.Size(937, 518)
        Me.grdPOTORDRX.TabIndex = 12
        Me.grdPOTORDRX.Text = "Purchase Orders"
        '
        'UltraTabPageControl5
        '
        Me.UltraTabPageControl5.Controls.Add(Me.grdICTSTYL1_Recent)
        Me.UltraTabPageControl5.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl5.Name = "UltraTabPageControl5"
        Me.UltraTabPageControl5.Size = New System.Drawing.Size(937, 518)
        '
        'grdICTSTYL1_Recent
        '
        Appearance73.BackColor = System.Drawing.SystemColors.Window
        Appearance73.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdICTSTYL1_Recent.DisplayLayout.Appearance = Appearance73
        UltraGridColumn28.Header.Caption = "Style Code"
        UltraGridColumn28.Header.VisiblePosition = 0
        UltraGridColumn28.Width = 107
        UltraGridColumn351.Header.Caption = "Status"
        UltraGridColumn351.Header.VisiblePosition = 1
        UltraGridColumn351.Width = 65
        UltraGridColumn1.Header.Caption = "Description"
        UltraGridColumn1.Header.VisiblePosition = 2
        UltraGridColumn1.Width = 191
        UltraGridColumn276.Header.VisiblePosition = 35
        UltraGridColumn276.Hidden = True
        UltraGridColumn277.Header.Caption = "Fabric"
        UltraGridColumn277.Header.VisiblePosition = 15
        UltraGridColumn277.Width = 90
        UltraGridColumn278.Header.Caption = "Season"
        UltraGridColumn278.Header.VisiblePosition = 16
        UltraGridColumn278.Width = 74
        UltraGridColumn279.Header.Caption = "Sub-Body"
        UltraGridColumn279.Header.VisiblePosition = 17
        UltraGridColumn279.Width = 89
        UltraGridColumn280.Header.Caption = "SDiv"
        UltraGridColumn280.Header.VisiblePosition = 19
        UltraGridColumn280.Width = 49
        UltraGridColumn285.Header.Caption = "Inner"
        UltraGridColumn285.Header.VisiblePosition = 20
        UltraGridColumn285.Width = 50
        UltraGridColumn286.Header.Caption = "Case Pack"
        UltraGridColumn286.Header.VisiblePosition = 21
        UltraGridColumn286.Width = 50
        UltraGridColumn331.Header.Caption = "Customer"
        UltraGridColumn331.Header.VisiblePosition = 22
        UltraGridColumn331.Width = 95
        UltraGridColumn332.Header.Caption = "FP"
        UltraGridColumn332.Header.VisiblePosition = 23
        UltraGridColumn332.Width = 30
        UltraGridColumn333.Header.VisiblePosition = 24
        UltraGridColumn333.Hidden = True
        UltraGridColumn334.Header.VisiblePosition = 25
        UltraGridColumn334.Hidden = True
        UltraGridColumn335.Header.VisiblePosition = 26
        UltraGridColumn335.Hidden = True
        UltraGridColumn336.Header.Caption = "Ldd Cost"
        UltraGridColumn336.Header.VisiblePosition = 64
        UltraGridColumn336.Width = 71
        UltraGridColumn337.Header.Caption = "Unit Pack"
        UltraGridColumn337.Header.VisiblePosition = 27
        UltraGridColumn337.Width = 50
        UltraGridColumn338.Header.VisiblePosition = 28
        UltraGridColumn338.Hidden = True
        UltraGridColumn339.Header.VisiblePosition = 29
        UltraGridColumn339.Hidden = True
        UltraGridColumn340.Header.VisiblePosition = 30
        UltraGridColumn340.Hidden = True
        UltraGridColumn342.Header.VisiblePosition = 31
        UltraGridColumn342.Hidden = True
        UltraGridColumn343.Header.Caption = "Retail"
        UltraGridColumn343.Header.VisiblePosition = 32
        UltraGridColumn343.Width = 68
        UltraGridColumn410.Header.Caption = "Duty Code"
        UltraGridColumn410.Header.VisiblePosition = 33
        UltraGridColumn410.Hidden = True
        UltraGridColumn410.Width = 70
        UltraGridColumn411.Header.VisiblePosition = 34
        UltraGridColumn411.Hidden = True
        UltraGridColumn412.Header.VisiblePosition = 36
        UltraGridColumn412.Hidden = True
        UltraGridColumn413.Header.VisiblePosition = 37
        UltraGridColumn413.Hidden = True
        UltraGridColumn414.Header.Caption = "Repl"
        UltraGridColumn414.Header.VisiblePosition = 48
        UltraGridColumn414.Hidden = True
        UltraGridColumn414.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn414.Width = 45
        UltraGridColumn415.Header.Caption = "Class"
        UltraGridColumn415.Header.VisiblePosition = 39
        UltraGridColumn415.Width = 91
        UltraGridColumn416.Header.Caption = "Case Cube"
        UltraGridColumn416.Header.VisiblePosition = 51
        UltraGridColumn416.Width = 87
        UltraGridColumn417.Header.Caption = "Min PO Qty"
        UltraGridColumn417.Header.VisiblePosition = 49
        UltraGridColumn417.Width = 93
        UltraGridColumn418.Header.Caption = "Purchase Notes"
        UltraGridColumn418.Header.VisiblePosition = 54
        UltraGridColumn419.Header.Caption = "Factory"
        UltraGridColumn419.Header.VisiblePosition = 63
        UltraGridColumn419.Width = 119
        UltraGridColumn420.Header.Caption = "Description 2"
        UltraGridColumn420.Header.VisiblePosition = 42
        UltraGridColumn421.Header.Caption = "Min SO Qty"
        UltraGridColumn421.Header.VisiblePosition = 69
        UltraGridColumn421.Width = 96
        UltraGridColumn422.Header.VisiblePosition = 70
        UltraGridColumn422.Hidden = True
        UltraGridColumn423.Header.VisiblePosition = 71
        UltraGridColumn423.Hidden = True
        UltraGridColumn424.Header.VisiblePosition = 72
        UltraGridColumn424.Hidden = True
        UltraGridColumn425.Header.Caption = "XMas Date"
        UltraGridColumn425.Header.VisiblePosition = 74
        UltraGridColumn425.Width = 109
        UltraGridColumn426.Header.VisiblePosition = 76
        UltraGridColumn426.Hidden = True
        UltraGridColumn427.Header.VisiblePosition = 78
        UltraGridColumn427.Hidden = True
        UltraGridColumn428.Header.Caption = "Vendor"
        UltraGridColumn428.Header.VisiblePosition = 61
        UltraGridColumn429.Header.Caption = "Style Matl Desc"
        UltraGridColumn429.Header.VisiblePosition = 43
        UltraGridColumn430.Header.Caption = "Cs Wgt"
        UltraGridColumn430.Header.VisiblePosition = 47
        UltraGridColumn430.Width = 60
        UltraGridColumn431.Header.Caption = "Group"
        UltraGridColumn431.Header.VisiblePosition = 40
        UltraGridColumn431.Width = 58
        UltraGridColumn432.Header.Caption = "Origin"
        UltraGridColumn432.Header.VisiblePosition = 45
        UltraGridColumn432.Width = 59
        UltraGridColumn433.Header.Caption = "Royalty"
        UltraGridColumn433.Header.VisiblePosition = 41
        UltraGridColumn433.Width = 70
        UltraGridColumn434.Header.VisiblePosition = 44
        UltraGridColumn434.Hidden = True
        UltraGridColumn435.Header.Caption = "Size"
        UltraGridColumn435.Header.VisiblePosition = 18
        UltraGridColumn435.Width = 60
        UltraGridColumn436.Header.Caption = "Dept"
        UltraGridColumn436.Header.VisiblePosition = 46
        UltraGridColumn436.Width = 60
        UltraGridColumn219.Header.VisiblePosition = 50
        UltraGridColumn219.Hidden = True
        UltraGridColumn220.Header.VisiblePosition = 52
        UltraGridColumn220.Hidden = True
        UltraGridColumn254.Header.VisiblePosition = 58
        UltraGridColumn254.Hidden = True
        UltraGridColumn39.Header.VisiblePosition = 53
        UltraGridColumn39.Hidden = True
        UltraGridColumn40.Header.VisiblePosition = 56
        UltraGridColumn40.Hidden = True
        UltraGridColumn498.Header.VisiblePosition = 55
        UltraGridColumn498.Hidden = True
        UltraGridColumn437.Header.Caption = "Body"
        UltraGridColumn437.Header.VisiblePosition = 38
        UltraGridColumn437.Width = 79
        UltraGridColumn42.Header.Caption = "LstOrd Date"
        UltraGridColumn42.Header.VisiblePosition = 57
        UltraGridColumn42.Width = 100
        UltraGridColumn43.Header.Caption = "LstOrd No"
        UltraGridColumn43.Header.VisiblePosition = 59
        UltraGridColumn43.Width = 100
        UltraGridColumn44.Header.Caption = "LstOrd Cust"
        UltraGridColumn44.Header.VisiblePosition = 60
        UltraGridColumn44.Width = 100
        UltraGridColumn45.Header.Caption = "LstOrd PO"
        UltraGridColumn45.Header.VisiblePosition = 62
        UltraGridColumn45.Width = 100
        UltraGridColumn438.Header.Caption = "On Hand"
        UltraGridColumn438.Header.VisiblePosition = 5
        UltraGridColumn438.Width = 70
        UltraGridColumn439.Header.Caption = "On PO"
        UltraGridColumn439.Header.VisiblePosition = 8
        UltraGridColumn439.Width = 70
        UltraGridColumn440.Header.Caption = "In Xit"
        UltraGridColumn440.Header.VisiblePosition = 9
        UltraGridColumn440.Width = 70
        UltraGridColumn441.Header.Caption = "Open"
        UltraGridColumn441.Header.VisiblePosition = 10
        UltraGridColumn441.Width = 70
        UltraGridColumn442.Header.Caption = "In Pick"
        UltraGridColumn442.Header.VisiblePosition = 11
        UltraGridColumn442.Width = 70
        UltraGridColumn443.Header.Caption = "Comm"
        UltraGridColumn443.Header.VisiblePosition = 12
        UltraGridColumn443.Width = 70
        UltraGridColumn444.Header.Caption = "Prod"
        UltraGridColumn444.Header.VisiblePosition = 13
        UltraGridColumn444.Width = 70
        UltraGridColumn445.Header.Caption = "Net Pos"
        UltraGridColumn445.Header.VisiblePosition = 14
        UltraGridColumn445.Width = 70
        UltraGridColumn446.Header.Caption = "Clr"
        UltraGridColumn446.Header.VisiblePosition = 3
        UltraGridColumn446.Hidden = True
        UltraGridColumn446.Width = 46
        UltraGridColumn447.Header.Caption = "Color Desc"
        UltraGridColumn447.Header.VisiblePosition = 4
        UltraGridColumn447.Hidden = True
        UltraGridColumn447.Width = 116
        UltraGridColumn448.Header.Caption = "Whse"
        UltraGridColumn448.Header.VisiblePosition = 6
        UltraGridColumn448.Width = 60
        UltraGridColumn449.Header.Caption = "Whse Desc"
        UltraGridColumn449.Header.VisiblePosition = 7
        UltraGridColumn449.Width = 100
        UltraGridColumn450.Header.Caption = "Image"
        UltraGridColumn450.Header.VisiblePosition = 80
        UltraGridColumn499.Header.VisiblePosition = 73
        UltraGridColumn499.Hidden = True
        UltraGridColumn500.Header.VisiblePosition = 75
        UltraGridColumn500.Hidden = True
        UltraGridColumn501.Header.VisiblePosition = 77
        UltraGridColumn501.Hidden = True
        UltraGridColumn502.Header.VisiblePosition = 79
        UltraGridColumn502.Hidden = True
        UltraGridColumn503.Header.VisiblePosition = 81
        UltraGridColumn503.Hidden = True
        UltraGridColumn504.Header.VisiblePosition = 82
        UltraGridColumn504.Hidden = True
        UltraGridColumn460.Header.Caption = "LDP"
        UltraGridColumn460.Header.VisiblePosition = 66
        UltraGridColumn460.Width = 67
        UltraGridColumn461.Header.VisiblePosition = 83
        UltraGridColumn461.Hidden = True
        UltraGridColumn462.Header.Caption = "ELC"
        UltraGridColumn462.Header.VisiblePosition = 67
        UltraGridColumn462.Width = 59
        UltraGridColumn463.Header.Caption = "Ext Cost"
        UltraGridColumn463.Header.VisiblePosition = 68
        UltraGridColumn463.Width = 97
        UltraGridColumn451.Header.Caption = "Ext Ldd Cost"
        UltraGridColumn451.Header.VisiblePosition = 65
        UltraGridColumn451.Width = 111
        UltraGridBand7.Columns.AddRange(New Object() {UltraGridColumn28, UltraGridColumn351, UltraGridColumn1, UltraGridColumn276, UltraGridColumn277, UltraGridColumn278, UltraGridColumn279, UltraGridColumn280, UltraGridColumn285, UltraGridColumn286, UltraGridColumn331, UltraGridColumn332, UltraGridColumn333, UltraGridColumn334, UltraGridColumn335, UltraGridColumn336, UltraGridColumn337, UltraGridColumn338, UltraGridColumn339, UltraGridColumn340, UltraGridColumn342, UltraGridColumn343, UltraGridColumn410, UltraGridColumn411, UltraGridColumn412, UltraGridColumn413, UltraGridColumn414, UltraGridColumn415, UltraGridColumn416, UltraGridColumn417, UltraGridColumn418, UltraGridColumn419, UltraGridColumn420, UltraGridColumn421, UltraGridColumn422, UltraGridColumn423, UltraGridColumn424, UltraGridColumn425, UltraGridColumn426, UltraGridColumn427, UltraGridColumn428, UltraGridColumn429, UltraGridColumn430, UltraGridColumn431, UltraGridColumn432, UltraGridColumn433, UltraGridColumn434, UltraGridColumn435, UltraGridColumn436, UltraGridColumn219, UltraGridColumn220, UltraGridColumn254, UltraGridColumn39, UltraGridColumn40, UltraGridColumn498, UltraGridColumn437, UltraGridColumn42, UltraGridColumn43, UltraGridColumn44, UltraGridColumn45, UltraGridColumn438, UltraGridColumn439, UltraGridColumn440, UltraGridColumn441, UltraGridColumn442, UltraGridColumn443, UltraGridColumn444, UltraGridColumn445, UltraGridColumn446, UltraGridColumn447, UltraGridColumn448, UltraGridColumn449, UltraGridColumn450, UltraGridColumn499, UltraGridColumn500, UltraGridColumn501, UltraGridColumn502, UltraGridColumn503, UltraGridColumn504, UltraGridColumn460, UltraGridColumn461, UltraGridColumn462, UltraGridColumn463, UltraGridColumn451})
        Me.grdICTSTYL1_Recent.DisplayLayout.BandsSerializer.Add(UltraGridBand7)
        Me.grdICTSTYL1_Recent.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance74.TextHAlignAsString = "Left"
        Me.grdICTSTYL1_Recent.DisplayLayout.CaptionAppearance = Appearance74
        Appearance75.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance75.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance75.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTSTYL1_Recent.DisplayLayout.GroupByBox.Appearance = Appearance75
        Appearance76.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTSTYL1_Recent.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance76
        Me.grdICTSTYL1_Recent.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdICTSTYL1_Recent.DisplayLayout.GroupByBox.Hidden = True
        Appearance77.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance77.BackColor2 = System.Drawing.SystemColors.Control
        Appearance77.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance77.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTSTYL1_Recent.DisplayLayout.GroupByBox.PromptAppearance = Appearance77
        Me.grdICTSTYL1_Recent.DisplayLayout.MaxColScrollRegions = 1
        Me.grdICTSTYL1_Recent.DisplayLayout.MaxRowScrollRegions = 1
        Appearance78.BackColor = System.Drawing.SystemColors.Window
        Appearance78.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.ActiveCellAppearance = Appearance78
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance79.BackColor = System.Drawing.SystemColors.Window
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.CardAreaAppearance = Appearance79
        Appearance80.BorderColor = System.Drawing.Color.Silver
        Appearance80.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.CellAppearance = Appearance80
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.CellPadding = 0
        Appearance81.BackColor = System.Drawing.SystemColors.Control
        Appearance81.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance81.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance81.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance81.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.GroupByRowAppearance = Appearance81
        Appearance82.TextHAlignAsString = "Left"
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.HeaderAppearance = Appearance82
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance83.BackColor = System.Drawing.SystemColors.Window
        Appearance83.BorderColor = System.Drawing.Color.Silver
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.RowAppearance = Appearance83
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance84.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdICTSTYL1_Recent.DisplayLayout.Override.TemplateAddRowAppearance = Appearance84
        Me.grdICTSTYL1_Recent.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdICTSTYL1_Recent.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdICTSTYL1_Recent.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdICTSTYL1_Recent.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTSTYL1_Recent.Location = New System.Drawing.Point(0, 0)
        Me.grdICTSTYL1_Recent.Name = "grdICTSTYL1_Recent"
        Me.grdICTSTYL1_Recent.Size = New System.Drawing.Size(937, 518)
        Me.grdICTSTYL1_Recent.TabIndex = 167
        Me.grdICTSTYL1_Recent.Text = "Recently Viewed Styles"
        '
        'UltraTabPageControl1
        '
        Me.UltraTabPageControl1.Controls.Add(Me.tabMain)
        Me.UltraTabPageControl1.Location = New System.Drawing.Point(1, 25)
        Me.UltraTabPageControl1.Name = "UltraTabPageControl1"
        Me.UltraTabPageControl1.Size = New System.Drawing.Size(941, 546)
        '
        'tabMain
        '
        Me.tabMain.Controls.Add(Me.UltraTabSharedControlsPage2)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl3)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl4)
        Me.tabMain.Controls.Add(Me.UltraTabPageControl5)
        Me.tabMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabMain.Location = New System.Drawing.Point(0, 0)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SharedControlsPage = Me.UltraTabSharedControlsPage2
        Me.tabMain.Size = New System.Drawing.Size(941, 546)
        Me.tabMain.TabIndex = 13
        Me.tabMain.TabOrientation = Infragistics.Win.UltraWinTabs.TabOrientation.BottomLeft
        UltraTab3.TabPage = Me.UltraTabPageControl3
        UltraTab3.Text = "Shipments"
        UltraTab4.TabPage = Me.UltraTabPageControl4
        UltraTab4.Text = "Open POs"
        UltraTab5.TabPage = Me.UltraTabPageControl5
        UltraTab5.Text = "Inventory Status"
        Me.tabMain.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab3, UltraTab4, UltraTab5})
        '
        'UltraTabSharedControlsPage2
        '
        Me.UltraTabSharedControlsPage2.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage2.Name = "UltraTabSharedControlsPage2"
        Me.UltraTabSharedControlsPage2.Size = New System.Drawing.Size(937, 518)
        '
        'UltraTabPageControl2
        '
        Me.UltraTabPageControl2.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl2.Name = "UltraTabPageControl2"
        Me.UltraTabPageControl2.Size = New System.Drawing.Size(941, 546)
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(150, 75)
        Me.UltraGroupBox1.TabIndex = 9
        '
        'spl
        '
        Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.spl.Location = New System.Drawing.Point(0, 0)
        Me.spl.Name = "spl"
        Me.spl.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'spl.Panel1
        '
        Me.spl.Panel1.Controls.Add(Me.UltraGroupBox1)
        Me.spl.Panel1Collapsed = True
        '
        'spl.Panel2
        '
        Me.spl.Panel2.Controls.Add(Me.tab)
        Me.spl.Size = New System.Drawing.Size(945, 574)
        Me.spl.SplitterDistance = 75
        Me.spl.TabIndex = 12
        '
        'tab
        '
        Me.tab.Controls.Add(Me.UltraTabSharedControlsPage1)
        Me.tab.Controls.Add(Me.UltraTabPageControl1)
        Me.tab.Controls.Add(Me.UltraTabPageControl2)
        Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tab.Location = New System.Drawing.Point(0, 0)
        Me.tab.Name = "tab"
        Me.tab.SharedControlsPage = Me.UltraTabSharedControlsPage1
        Me.tab.Size = New System.Drawing.Size(945, 574)
        Me.tab.TabIndex = 12
        UltraTab1.TabPage = Me.UltraTabPageControl1
        UltraTab1.Text = "0"
        UltraTab2.TabPage = Me.UltraTabPageControl2
        UltraTab2.Text = "1"
        Me.tab.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab1, UltraTab2})
        '
        'UltraTabSharedControlsPage1
        '
        Me.UltraTabSharedControlsPage1.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage1.Name = "UltraTabSharedControlsPage1"
        Me.UltraTabSharedControlsPage1.Size = New System.Drawing.Size(941, 546)
        '
        'POFWREC2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1158, 574)
        Me.Name = "POFWREC2"
        Me.Text = "POFWREC2"
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl6.ResumeLayout(False)
        CType(Me.grdPOTSHIP3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl7.ResumeLayout(False)
        Me.splPOTSHIP7.Panel1.ResumeLayout(False)
        Me.splPOTSHIP7.Panel2.ResumeLayout(False)
        CType(Me.splPOTSHIP7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splPOTSHIP7.ResumeLayout(False)
        CType(Me.grdPOTSHIP7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdPOTSHIP8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl8.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grdPOTLPNL1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl3.ResumeLayout(False)
        Me.splPOTSHIPX.Panel1.ResumeLayout(False)
        Me.splPOTSHIPX.Panel2.ResumeLayout(False)
        CType(Me.splPOTSHIPX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splPOTSHIPX.ResumeLayout(False)
        CType(Me.grdPOTSHIPX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tabShipment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabShipment.ResumeLayout(False)
        Me.UltraTabPageControl4.ResumeLayout(False)
        CType(Me.grdPOTORDRX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl5.ResumeLayout(False)
        CType(Me.grdICTSTYL1_Recent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl1.ResumeLayout(False)
        CType(Me.tabMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabMain.ResumeLayout(False)
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.Panel1.ResumeLayout(False)
        Me.spl.Panel2.ResumeLayout(False)
        CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.ResumeLayout(False)
        CType(Me.tab, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents spl As System.Windows.Forms.SplitContainer
    Friend WithEvents grdPOTSHIPX As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents tab As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage1 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl1 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTabPageControl2 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents splPOTSHIPX As System.Windows.Forms.SplitContainer
    Friend WithEvents grdPOTSHIP3 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents tabMain As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage2 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl3 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTabPageControl4 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTabPageControl5 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents grdPOTORDRX As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grdICTSTYL1_Recent As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents tabShipment As UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage3 As UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl6 As UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTabPageControl7 As UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraTabPageControl8 As UltraWinTabControl.UltraTabPageControl
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents grdPOTLPNL1 As UltraWinGrid.UltraGrid
    Friend WithEvents splPOTSHIP7 As SplitContainer
    Friend WithEvents grdPOTSHIP8 As UltraWinGrid.UltraGrid
    Friend WithEvents grdPOTSHIP7 As UltraWinGrid.UltraGrid
End Class
