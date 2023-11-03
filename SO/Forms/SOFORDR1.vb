Imports Infragistics.Win.UltraWinGrid

Public Class SOFORDR1

    'Stop ' Use Copy-From Order Values for SREP_CODE, SREP2_CODE, SHIP_VIA_CODE
    'Stop ' ORDR_ADDR_TYPE_ST LOGIC NEEDED
    'Stop ' LOAD UP COMBO THAT OFFERS SELECTION IF BILLTO ADDRESSES
    'Stop ' Sql = "Select SOTORDR2.*, 1 RANGE_STYLE_QTY_PER_PP from SOTORDR2 where ORDR_NO = '" & ORDR_NO_x & "'"
#Region "Declarations"
    Dim CUST_CODE As String
    Dim WHSE_CODE As String
    Dim REV_NO As Int32
    Dim ORDR_NO As String
    Dim ORDR_NO_to_copy As String
    Dim ORDR_GROUP_NO As String     ' ORDR_GROUP_NO for Order currently in process
    Dim ORDR_GROUP_NOs As New List(Of String)    ' ORDR_GROUP_NOs for Order(s) currently in process

    Dim EDI_JRNL_NO As String       ' EDI_JRNL_NO for Order currently in process
    Dim ORDR_CUST_PO As String      ' Customer's PO No
    Dim CUST_STORE_NO As String     ' Store Related to this Order
    Dim SREP_CODE As String         ' Orders Sales Rep Code
    Dim SREP2_CODE As String        ' Orders Sales Rep2 Code

    Dim rowSOTORDR1 As DataRow
    Dim rowSOTORDR0 As DataRow
    Dim rowARTCUST1 As DataRow      ' ARTCUST1 for the Sold-To
    Dim rowARTCUST1_BT As DataRow   ' ARTCUST1 for the Bill-To
    Dim rowARTCUST2 As DataRow      ' ARTCUST2 for Store
    Dim dvwSOTORDR5 As DataView
    Dim sqlSOTORDRB As String = ""

    Dim rowSOTORDR1po As DataRow
    Dim rowICTSTYL1 As DataRow

    Dim disable_update As Boolean   ' flag indicating that the Update Button should be disabled
    Dim TOTAL_ORDR_AMT As Decimal ' Total Order Amt from EDI Header
    Dim TOTAL_ORDR_QTY As Decimal ' Total Order Qty from EDI Header

    Dim ORDR_LNOs As New List(Of Int64) ' list of ORDR_LNOs that are deleted

    'Dim MASS_CHANGE_TEXT As String  ' Value to Propagate in Multiple Order Maintenance
    'Dim MASS_CHANGE_VALUE As String ' Value of Cloned Field Prior to Change

    'Dim do_not_update               ' to stop grdSOWORDR2_LostFocus from Updating

    Dim CUST_NAME As String         ' Sold-To Customer Name
    Dim CUST_BILL_TO_CUST As String ' Bill-To Customer Code
    Dim CUST_DC_NO As String        ' DC Related to the CUST_STORE_NO of this Order

    Dim COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE

    Dim restore_reservation As Boolean
    Dim sub_grid As String

    Dim multi_store_is_active As Boolean            ' Multi-Store Order mode
    Dim multi_store_changes_made_to_SOTORDRS As Boolean           ' Changes were made in grdSOTORDRS and need to get to grdSOWORDR2 without recursively going back to grdSOWORDRS
    Dim CUST_STORE_NOs_multi_store As New List(Of String)
    Dim ORDR_NOs_to_maintain As New List(Of String)

    Dim multiple_order_maintenance As Boolean = False
    Dim multiple_order_type As String = ""

    Dim msqty As Int64               ' Repeat MS qty when entering stores blindly
    Dim msqty_col As Int64        ' ORDR_LNO for MS qty
    Dim RANGE_TYPE As String
    Dim ALLOW_CHANGE_RANGE As String
    Dim STYLE_CODE_last_entry As String = ""
    Dim sqlSOTORDRH As String = ""

    Dim CURR_CODE As String = "USD"
    Dim CURR_EXCH_RATE As Decimal = 1

    Private clsShip As New TAC.WHCSHIP1
    Private shipPackageDetailList As New List(Of nsoftware.InShip.PackageDetail)
    Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

    Dim clsASCBASE1_allo As ASCBASE1

    Dim multistore_changes_were_made_to_qty As Boolean
    Dim multistore_OK_TO_UPDATE As Boolean

    Dim BTB_TYPE As String
    Dim ORDR_LNO_ctr As Integer = 0

    Dim SOTORDPX As String = ""
    Dim PO_ORDER_NOs As New List(Of String)

    Private clsSOCORDR1 As TAC.SOCORDR1
    Dim blnAutomatic As Boolean = False

    Dim COLUMN_NAMEs_All As New List(Of String)
    Dim COLUMN_NAMEs_Short As New List(Of String)
    Private clsTACENCRY As TAC.ASCENCRY

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("POTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("EDTPARM1")

        clsTACENCRY = New TAC.ASCENCRY
        Dim rowASTPARMP As DataRow = ASCDATA1.GetDataRow("Select * from ASTPARMP WHERE AS_PARM_KEY = 'Z'")
        If rowASTPARMP Is Nothing OrElse Not rowASTPARMP.Table.Columns.Contains("AS_PARM_USE_ENCRYPTION") OrElse rowASTPARMP.Item("AS_PARM_USE_ENCRYPTION") & String.Empty <> "1" Then
            clsTACENCRY.UseEncryption = False
        Else
            clsTACENCRY.UseEncryption = True
        End If

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            cmdAutoPO.Visible = True
        End If

        With dst

            ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 1)
            .Tables("SOTORDRX").Columns.Add("CUST_CITY", GetType(System.String))
            .Tables("SOTORDRX").Columns.Add("CUST_STATE", GetType(System.String))
            .Tables("SOTORDRX").Columns.Add("CUST_COUNTRY", GetType(System.String))
            'ORDR_QTY_OPEN
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_OPEN", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_PICK", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_SHIP", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_CANC", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_TOTAL", GetType(System.Decimal), "ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP")

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)


            'If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
            With .Tables("SOTORDR1").Columns
                .Add("PF_WEIGHT", GetType(System.Decimal))
                .Add("PF_WEIGHT_UOM", GetType(System.String))
                .Add("PO_SHIPMENT_NO", GetType(System.String))
                .Add("PF_CARTONS", GetType(System.Int64))
                .Add("PF_NOTE", GetType(System.String))
                .Add("PF_OVERSEAS_DOMESTIC", GetType(System.String))
                .Add("PF_INV_DATE", GetType(System.DateTime))
                .Add("PF_INV_NO", GetType(System.String))
                .Add("PF_VIA", GetType(System.String))
            End With
            With .Tables("SOTORDR1")
                .Columns("PF_WEIGHT_UOM").MaxLength = 3
                .Columns("PO_SHIPMENT_NO").MaxLength = 6
                .Columns("PF_NOTE").MaxLength = 90
                .Columns("PF_OVERSEAS_DOMESTIC").MaxLength = 1
            End With
            'End If


            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & " from SOTORDR2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2_ORIG", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTSTYL1.CASE_CUBE, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_STATUS, ICTSTYL1.STYLE_ASST_QTY, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
                & ", ICTSTDQ3.DATE_1, ICTSTDQ3.QTY_1, ICTSTDQ3.DATE_2, ICTSTDQ3.QTY_2, ICTSTDQ3.DATE_3, ICTSTDQ3.QTY_3, ICTSTDQ3.DATE_4, ICTSTDQ3.QTY_4" & vbCrLf _
                & " from SOTORDR2,ICTCOLR1,ICTSTYL1,ICTSTYC1,ICTSTDQ3" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = :PARM1" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and ICTSTDQ3.ORDR_GROUP_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and ICTSTDQ3.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTDQ3.COLOR_CODE (+) = SOTORDR2.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2)
            .Tables("SOTORDR2").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(ORDR_QTY,0) / ISNULL(CARTON_PACK_QTY,0))")
            .Tables("SOTORDR2").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")
            .Tables("SOTORDR2").Columns.Add("MU_PCT", GetType(System.Decimal), "IIF(ISNULL(PO_COST,0) = 0, 0, 100 * (ISNULL(ORDR_UNIT_PRICE,0) - ISNULL(PO_COST,0)) / ISNULL(PO_COST,0))")
            .Tables("SOTORDR2").Columns.Add("ORDR_AMT_CURR", GetType(System.Decimal), "ISNULL(ORDR_QTY,0)*ISNULL(ORDR_UNIT_PRICE_CURR,0)")


            If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
                .Tables("SOTORDR2").Columns.Add("PF_QTY", GetType(System.Int64))
                .Tables("SOTORDR2").Columns.Add("PF_DUTY_HTS_CODE", GetType(System.String))
                .Tables("SOTORDR2").Columns("PF_DUTY_HTS_CODE").MaxLength = 20
                .Tables("SOTORDR2").Columns.Add("PF_ORDER_NO", GetType(System.String))
                .Tables("SOTORDR2").Columns("PF_ORDER_NO").MaxLength = 20
            End If


            ' currently this is RGI specific.
            If Not dst.Tables("SOTORDR2").Columns.Contains("ORDR_LINE_CANC") Then
                .Tables("SOTORDR2").Columns.Add("ORDR_LINE_CANC", GetType(System.String))
                .Tables("SOTORDR2").Columns("ORDR_LINE_CANC").MaxLength = 1
            End If

            With .Tables("SOTORDR2").Columns
                .Add("RANGE_STYLE_QTY_PER_PP", GetType(System.Int64))
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_ALLO", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("DISC_AMT", GetType(System.Decimal), "ISNULL(STYLE_PRICE,0)-ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("DISC_PCT", GetType(System.Decimal), "IIF(ISNULL(STYLE_PRICE,0)=0,0,100*DISC_AMT/ISNULL(STYLE_PRICE,0))")

                ' .Add("ORDR_QTY_ALLO_CUR", GetType(System.Int64), "IIF(ORDR_RELEASE_AVAIL IS NULL,ORDR_QTY_ALLO,0)")
                .Add("ORDR_QTY_ALLO_CUR", GetType(System.Int64), "ISNULL(QTY_1,0)")
                .Add("ORDR_AMT_ALLO_CUR", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO_CUR,0) * ISNULL(ORDR_UNIT_PRICE,0)")

                .Add("ORDR_UNIT_COST", GetType(System.Decimal))
                .Add("CGS", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_COST,0)")
                .Add("GP_AMT", GetType(System.Decimal), "ISNULL(ORDR_AMT_SHIP,0) - ISNULL(CGS,0)")
                .Add("GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ORDR_AMT_SHIP,0)=0,0,100 * ISNULL(GP_AMT,0) / ISNULL(ORDR_AMT_SHIP,0))")
            End With
            .Tables("SOTORDR2").Columns("ORDR_UNIT_PRICE_MANUAL").DefaultValue = "0"
            '.Tables("SOTORDR2").Columns("DUTY_RATE_CODE").DataType = GetType(System.Double)
            ' .Tables("SOTORDR2").Columns("ORDR_UNIT_PRICE").DataType = GetType(System.Double)

            'If ASCMAIN1.CLIENT = "RGI" Then
            With .Tables("SOTORDR2").Columns
                .Add("AMT_1", GetType(System.Decimal), "ISNULL(QTY_1,0)*ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("AMT_2", GetType(System.Decimal), "ISNULL(QTY_2,0)*ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("AMT_3", GetType(System.Decimal), "ISNULL(QTY_3,0)*ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("AMT_4", GetType(System.Decimal), "ISNULL(QTY_4,0)*ISNULL(ORDR_UNIT_PRICE,0)")
            End With
            'End If

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_TYPE_CODE" _
                & ", SOTORDR2.STYLE_DESC, SOTORDR2.STYLE_UOM" _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC" _
                & ", SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.ORDR_UNIT_PRICE_MANUAL, ICTCOLR1.COLOR_DESC" _
                & ",SOTORDR1.ORDR_DATE,SOTORDR1.ORDR_CUST_PO,SOTORDR1.ORDR_SHIP_DATE,SOTORDR1.ORDR_CANCEL_DATE" _
                & " from SOTORDR2,SOTORDR1,ICTCOLR1" _
                & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE"
            sqlSOTORDRH = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTORDRH", "**", 0, False, "", 2)
            With .Tables("SOTORDRH").Columns
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            End With

            Create_TDA(.Tables.Add, "SOTORDC1", "*")
            With .Tables("SOTORDC1").Columns
                .Add("CUST_CREDIT_CARD_LAST4", GetType(System.String))
                .Add("CCPA_DATE_VOID", GetType(System.String))
            End With
            Create_TDA(.Tables.Add, "SOTORDC2", "*", 1)
            Create_Relation("SOTORDC1", "SOTORDC2", "ORDR_NO,TRANS_NO")

            ASCMAIN1.sql = "Select * from ICTSTAT2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 2, False)

            '''TEMPDGJ TO BUILD SOTCSTY1 RECORDS FOR VAN AMAZONFBA FROM AMAZON SPREADSEET
            '''Create_TDA(.Tables.Add("SOTCSTY1"), "SOTCSTY1", "*")

            ASCMAIN1.sql = "Select * from ICTDISC1"
            Create_TDA(.Tables.Add, "ICTDISC1", "**", 0, False)
            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            With .Tables.Add("ICTPRICX")
                .Columns.Add("TIER", GetType(System.Int32))
                .Columns.Add("PCT", GetType(System.Int32))
                .Columns.Add("DESC")
                .Columns.Add("CASES", GetType(System.Decimal))
                .Columns.Add("ABBR")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("PRICE", GetType(System.Decimal))
                .Columns.Add("CUST_PRICE", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "Select * from ICTSTDQ2"
            Create_TDA(.Tables.Add, "ICTSTDQ2", "**", 0, False)

            'ASCMAIN1.sql = "Select * from ICTSTDQ3"
            Create_TDA(.Tables.Add, "ICTSTDQ3", "*", 3, False)

            Create_TDA(.Tables.Add, "SOTORDR3", "*", 1)
            .Tables("SOTORDR3").Columns.Add("ORDR_AMT", GetType(System.Decimal), "ORDR_QTY * ORDR_UNIT_PRICE")
            Dim szt As String = ""
            For I As Integer = 1 To 12
                szt &= "+ISNULL(SIZE_QTY_" & Format(I, "00") & ",0)"
            Next
            .Tables("SOTORDR3").Columns.Add("TOTAL_SIZE_QTY", GetType(System.Decimal), szt)

            Create_TDA(.Tables.Add, "SOTORDR4", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)
            .Tables("SOTORDR5").Columns("CUST_ADDR_CODE").MaxLength = 10

            'T = .Tables.Add("SOTORDR5_BT") : T = .Tables("SOTORDR5").Clone
            'T = .Tables.Add("SOTORDR5_ST") : T = .Tables("SOTORDR5").Clone
            'T = .Tables.Add("SOTORDR5_MK") : T = .Tables("SOTORDR5").Clone
            'T = .Tables.Add("SOTORDR5_DC") : T = .Tables("SOTORDR5").Clone
            'T = .Tables.Add("SOTORDR5_BY") : T = .Tables("SOTORDR5").Clone

            Create_TDA(.Tables.Add, "SOTORDR9", "*", 1)

            With .Tables.Add("SOTORDRT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("KEY")}
            End With

            ASCMAIN1.sql = "Select SOTPICK1.*, SOTSHIP1.SHIP_DATE_SHIPPED, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & " from SOTPICK1,SOTSHIP1,SOTINVH1 " & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE (+) = 'I'" & vbCrLf _
                & "   and SOTINVH1.INV_NO (+) = SOTPICK1.INV_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS <> 'D'" & vbCrLf _
                & "   and SOTPICK1.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*" & vbCrLf _
                & " from SOTPICK2,SOTPICK1 " & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO and SOTPICK1.ORDR_NO = :PARM1" _
                & "   and SOTPICK1.PICK_STATUS <> 'D'"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)


            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            Create_Relation("SOTORDR2", "SOTPICK2", "ORDR_NO,ORDR_LNO")
            With .Tables("SOTPICK2").Columns
                .Add("STYLE_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).STYLE_CODE")
                .Add("COLOR_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).COLOR_CODE")
                .Add("STYLE_DESC", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).STYLE_DESC")
                .Add("COLOR_DESC", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).COLOR_DESC")
                .Add("RANGE_STYLE_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).RANGE_STYLE_CODE")
                .Add("PICK_STATUS", GetType(System.String), "PARENT(SOTPICK1_SOTPICK2).PICK_STATUS")
                .Add("PICK_TOTAL", GetType(System.Decimal), "IIF(PICK_STATUS = 'P',ISNULL(PICK_QTY,0),ISNULL(PICK_QTY_CONF,0)) * ISNULL(PICK_UNIT_PRICE,0)")
            End With

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from SOTCART1 " & vbCrLf _
                & " where SOTCART1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
                & " from SOTCART2,SOTCART1 " & vbCrLf _
                & " where SOTCART2.CART_NO = SOTCART1.CART_NO" _
                & "   and SOTCART1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "V", 2)

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")


            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.CUST_SKU, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SIZE_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE_SUB" & vbCrLf _
                & ", NVL(SOTORDR2.STYLE_CODE_SUB,SOTORDR2.STYLE_CODE) || ':' || SOTORDR2.COLOR_CODE || ':' || NVL(SOTORDR2.CUST_STYLE_CODE,'') ||  ':' || NVL(SOTORDR2.CUST_COLOR_CODE,'') ||  ':' || NVL(SOTORDR2.CUST_SKU,'') ||  ':' || NVL(SOTORDR2.CUST_UPC,'') ||  ':' || NVL(SOTORDR2.CUST_SIZE_CODE,'') ||  ':' || NVL(SOTORDR2.STYLE_CODE,'') STYLE_KEY" & vbCrLf _
                & ", NVL(SOTORDR2.STYLE_CODE_SUB,SOTORDR2.STYLE_CODE) STYLE_CODE_ORIG" & vbCrLf _
                & " from SOTORDR2 " & vbCrLf _
                & " where SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRQ", "**", 0, False, "V", 2)

            With .Tables.Add("SOTORDRQ_KEY")
                With .Columns
                    .Add("ORDR_NO")
                    .Add("STYLE_KEY")
                    .Add("ORDR_LNO", GetType(System.Int32))
                End With
                .PrimaryKey = New DataColumn() { .Columns("ORDR_NO"), .Columns("STYLE_KEY")}
            End With

            'ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE STYLE_CODE_ORIG, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.ORDR_LNO, SOTORDR2.STYLE_CODE_SUB, SOTORDR2.ORDR_UNIT_PRICE" _
            '    & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC, SOTORDR2.ORDR_QTY_ORIG" _
            '    & ", SOTORDR2.STYLE_UOM, SOTORDR2.STYLE_DESC, SOTORDR2.CARTON_PACK_QTY" _
            '    & " from SOTORDR2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE"



            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE STYLE_CODE_ORIG, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_CODE_SUB, SOTORDR2.ORDR_UNIT_PRICE" _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_COLOR_CODE" _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC, SOTORDR2.ORDR_QTY_ORIG" _
                & ", SOTORDR2.STYLE_UOM, SOTORDR2.STYLE_DESC, SOTORDR2.CARTON_PACK_QTY" _
                & " from SOTORDR2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE"
            Create_TDA(.Tables.Add, "SOTORDRI", "**", 0, False, "", 0) ' 3)
            With .Tables("SOTORDRI")
                With .Columns
                    .Add("STYLE_KEY")
                    .Add("COL", GetType(System.Int32))
                    .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                    .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                    .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                    .Add("STYLE_KEY_CLONED_FROM")
                End With
                ' .PrimaryKey = New DataColumn() {.Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
                .PrimaryKey = New DataColumn() { .Columns("STYLE_KEY")}
            End With

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO" & vbCrLf _
              & ", SOTORDR1.CUST_STORE_NO" & vbCrLf _
              & ", SOTORDR1.ORDR_SHIP_DATE" & vbCrLf _
              & ", SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
              & ", SOTORDR1.ORDR_CUST_PO" & vbCrLf _
              & ", SOTORDR1.ORDR_DEPT" & vbCrLf _
              & ", SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
              & ", SOTORDR1.ORDR_HOLD" & vbCrLf _
              & ", SOTORDR1.ORDR_STATUS" & vbCrLf _
              & ", SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
              & ", ARTCUST2.CUST_NAME CUST_STORE_NAME" & vbCrLf _
              & ", ARTCUST2.CUST_ADDR_NAME CUST_STORE_LOCATION" & vbCrLf _
              & " from SOTORDR1,ARTCUST2 " & vbCrLf _
              & " where SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
              & "   and ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
              & "   and ARTCUST2.CUST_ADDR_TYPE = 'MK'" & vbCrLf _
              & "   and ARTCUST2.CUST_ADDR_CODE = SOTORDR1.CUST_STORE_NO" & vbCrLf
            sqlSOTORDRB = ASCMAIN1.sql
            ASCMAIN1.sql &= "   and SOTORDR1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRB", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "SOTORDR0", "*")
            With .Tables("SOTORDR0")
                .Columns.Add("ORDR_HOLD")
                .Columns.Add("REASON_CODE")
                .Columns.Add("ORDR_ADDR_TYPE_ST")

                .Columns.Add("ORDR_HOLD_REASON")
                '.Columns.Add("ORDR_ARRIVAL_DATE", GetType(System.DateTime))
                '.Columns.Add("ORDR_LAST_ARRIVAL_DATE", GetType(System.DateTime))
                .Columns.Add("ORDR_SHIP_INSTR")
                .Columns.Add("ORDR_INV_COMMENT")
                .Columns.Add("TERM_CODE")
                '.Columns.Add("WHSE_CODE")
                .Columns.Add("SHIP_VIA_CODE")
                .Columns.Add("FRT_TERMS")
                '.Columns.Add("REASON_CODE")
                .Columns.Add("SREP2_CODE")
                .Columns.Add("CUST_FACTOR_IND")
            End With

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTORDR1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V")
            .Tables("TATEVNT1").Columns.Add("ATTACHMENT_EXT")

            ASCMAIN1.sql = "Select * from SOTORDXR where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDXR", "**", 0, True, "V")

            ASCMAIN1.sql = "Select CUST_STORE_NO, ORDR_CUST_PO from SOTORDR1 where ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRS", "**", 0, False, "V", 1)
            With .Tables("SOTORDRS")
                .Columns.Add("TOTAL_QTY", GetType(System.Int64))
                .Columns.Add("TOTAL_AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("CUST_STORE_NO")}
            End With

            ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC from SOTORDR2,ICTCOLR1" _
                & " where SOTORDR2.ORDR_NO = :PARM1" _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTORDRR", "**", 0, False, "V", 0)
            .Tables("SOTORDRR").Columns.Add("RANGE_STYLE_QTY_PER_PP", GetType(System.Int64))
            .Tables("SOTORDRR").Columns.Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")

            Create_Relation("SOTORDR2", "SOTORDR3", "ORDR_NO,ORDR_LNO")
            Create_Relation("SOTORDR2", "SOTORDRR", "ORDR_NO,ORDR_LNO")

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from SOTORDR7 where ORDR_GROUP_NO = :PARM1 " & vbCrLf _
                & "   and SOTORDR7.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTORDR7.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTORDR7.PICK_BATCH_NO is Null"
            Create_TDA(.Tables.Add, "SOTORDR7", "**", 0, True, "VVV", 1)

            Create_TDA(.Tables.Add, "SOTRSRV1", "*")
            Create_TDA(.Tables.Add, "SOTRSRV2", "*")

            ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV2,SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.CUST_CODE = :PARM1 " & vbCrLf _
                & "   and SOTRSRV2.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTRSRV2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_QTY_OPEN > 0" & vbCrLf
            Create_TDA(.Tables.Add, "SOTRSRVX", "**", 0, False, "VVV", 0)


            ASCMAIN1.sql = "Select ARTCUST2.* from ARTCUST2" & vbCrLf _
                 & " where CUST_CODE = :PARM1" & vbCrLf _
                 & "   and CUST_ADDR_TYPE = :PARM2" & vbCrLf _
                 & "   and CUST_ADDR_CODE = :PARM3" & vbCrLf
            Create_TDA(.Tables.Add, "ARTCUST2_BT", "**", 0, False, "VVV", 0)
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "VVV", 3)

            ASCMAIN1.sql = "Select SOTORDR5.* from SOTORDR5" & vbCrLf _
                 & " where ORDR_NO = :PARM1" & vbCrLf _
                 & "   and CUST_ADDR_TYPE = :PARM2" & vbCrLf
            Create_TDA(.Tables.Add, "SOTORDR5_BT", "**", 0, False, "VV", 0)

            With .Tables.Add("EDTDOCS1")
                .Columns.Add("EDI_DOC_ID")
                .Columns.Add("EDI_DOC_DATE", GetType(System.DateTime))
                .Columns.Add("EDI_DOC_SEQ_NO")
                .Columns.Add("FILENAME_GEN")
                .Columns.Add("FILENAME_ABS")
                .Columns.Add("EDI_DOC_TEXT")
                .Columns.Add("EDI_DOC_STATUS")
                .Columns.Add("EDI_DOC_DESC")
                .Columns.Add("EDI_ISA_NO")
            End With

            With .Tables.Add("SOTORDR1_HOLDS")
                .Columns.Add("ORDR_NO")
                .Columns.Add("ORDR_REL_HOLD_CODE")
            End With

            ASCMAIN1.sql = "Select SIZE_CODE from ICTSIZE1"
            Create_TDA(.Tables.Add, "ICTSIZE1", "**", 0, False, "", 1)
            .Tables("ICTSIZE1").Columns.Add("SEL")

            ASCMAIN1.sql = "Select * from SOTWORK1 where WO_REF_TYPE = 'S' and WO_REF_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTWORK1", "**", 0, , "V", 1)
            ASCMAIN1.sql = "Select * from SOTWORK2 where WO_NO in " _
                & " (Select WO_NO from SOTWORK1 where WO_REF_TYPE = 'S' and WO_REF_NO = :PARM1)"
            Create_TDA(.Tables.Add, "SOTWORK2", "**", 0, , "V", 1)

            'ASCMAIN1.sql = "Select STYLE_STATUS, STYLE_COLOR_STATUS" _
            '    & " from ICTSTYL1,ICTSTYC1" _
            '    & " where ICTSTYC1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" _
            '    & "   and ICTSTYL1.STYLE_CODE = :PARM1" _
            '    & "   and ICTSTYC1.COLOR_CODE (+) = :PARM2"
            'Create_TDA(.Tables.Add, "ICTSTYL1_VAR", "**", 0, False, "VV", 0)

            If ASCMAIN1.CLIENT = "VAN" And Not InquiryMode Then
                With .Tables.Add("ERROR_TBL")
                    .Columns.Add("SKU", GetType(System.String))
                    .Columns.Add("ERROR_DETAIL", GetType(System.String))
                End With
            End If

            Create_TDA(.Tables.Add, "SOTORDRG", "*")

            Create_TDA(.Tables.Add, "POTORDR1", "*", 1)

            ASCMAIN1.sql = "Select POTORDR2.*,ICTSTYL1.CASE_CUBE" _
                    & " from POTORDR2,ICTSTYL1" _
                    & " where ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" _
                    & "   and POTORDR2.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, True, "V", 2)
            .Tables("POTORDR2").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(PO_QTY_ORD,0) / ISNULL(CARTON_PACK_QTY,0))")
            .Tables("POTORDR2").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")

            Create_Relation("POTORDR1", "POTORDR2", "PO_ORDER_NO")

            .Tables("POTORDR1").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "SUM(CHILD(POTORDR1_POTORDR2).TOTAL_CUBE)")


            ASCMAIN1.sql = "Select SOTORDP1.*, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                    & " from SOTORDP1,SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDP1.ORDR_NO and SOTORDP1.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDP1", "**", 0, True, "V")

            ASCMAIN1.sql = "Select * from SOTORDP2 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDP2", "**", 0, True, "V")

            Create_Relation("SOTORDR2", "SOTORDP2", "ORDR_NO,ORDR_LNO")
            With .Tables("SOTORDP2").Columns
                .Add("STYLE_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTORDP2).STYLE_CODE")
                .Add("COLOR_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTORDP2).COLOR_CODE")
                .Add("ORDR_QTY", GetType(System.Int64), "PARENT(SOTORDR2_SOTORDP2).ORDR_QTY")
                .Add("ORDR_UNIT_PRICE", GetType(System.Decimal), "PARENT(SOTORDR2_SOTORDP2).ORDR_UNIT_PRICE")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            End With

            Create_Relation("SOTORDP1", "SOTORDP2", "ORDR_NO,INV_NO")
            With .Tables("SOTORDP1").Columns
                .Add("INV_TOTAL_AMOUNT", GetType(System.Decimal), "SUM(CHILD(SOTORDP1_SOTORDP2).ORDR_AMT_SHIP)")
            End With

            ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTINVH1", "**", 0, False, "V")


            With .Tables.Add("SOTCORDR")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("STYLE_DESC")
                .Columns.Add("CUST_STYLE_CODE")
                .Columns.Add("CUST_COLOR_CODE")
                .Columns.Add("CUST_SIZE_CODE")
                .Columns.Add("CUST_UPC")
                .Columns.Add("CUST_SKU")
                .Columns.Add("ORDR", GetType(System.Int64))
                .Columns.Add("OPEN", GetType(System.Int64))
                .Columns.Add("PICK", GetType(System.Int64))
                .Columns.Add("SHIP", GetType(System.Int64))
                .Columns.Add("CANC", GetType(System.Int64))
                .Columns.Add("ALLO", GetType(System.Int64))
                .Columns.Add("ORDR_AMT", GetType(System.Decimal))


                .Columns("ORDR").DefaultValue = 0
                .Columns("OPEN").DefaultValue = 0
                .Columns("PICK").DefaultValue = 0
                .Columns("SHIP").DefaultValue = 0
                .Columns("CANC").DefaultValue = 0
                .Columns("ALLO").DefaultValue = 0
                .Columns("ORDR_AMT").DefaultValue = 0

                .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}

            End With

            Dim TBL As DataTable = .Tables("SOTORDR0").Clone
            TBL.TableName = "SOTCORDG"
            .Tables.Add(TBL)
        End With

        If SOTORDPX = "" Then
            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.VEND_CODE from SOTORDR2,ICTSTYL1 where ROWNUM < 1"
            SOTORDPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDPX & " add Primary Key (ORDR_NO, ORDR_LNO)")
        End If

        ASCMAIN1.sql = "Select SIZE_CODE from ICTSIZE1 order by SIZE_CODE"
        cbeICTSIZE1.DataSource = ASCDATA1.GetDataTable

        Fill_Records("ICTSIZE1")
        Fill_Records("ICTDISC1")
        Fill_Records("ICTCLAS1")

        grdSOTORDRB.DataSource = dst.Tables("SOTORDRB")
        grdSOTORDRI.DataSource = dst.Tables("SOTORDRI")

        grdEDTDOCS1.DataSource = dst.Tables("EDTDOCS1")
        '   grdEDTDOCS1.Visible = False

        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        grdSOTORDRH.DataSource = dst.Tables("SOTORDRH")
        grdICTPRICX.DataSource = dst.Tables("ICTPRICX")
        grdICTSTDQ2.DataSource = dst.Tables("ICTSTDQ2")
        grdICTSTDQ3.DataSource = dst.Tables("ICTSTDQ3")
        grdSOTORDR3.DataSource = dst.Tables("SOTORDR3")
        grdSOTORDR4.DataSource = dst.Tables("SOTORDR4")
        grdSOTORDRR.DataSource = dst.Tables("SOTORDRR")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdSOTORDXR.DataSource = dst.Tables("SOTORDXR")
        grdSOTORDRS.DataSource = dst.Tables("SOTORDRS")
        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")

        grdSOTORDRT.DataSource = dst.Tables("SOTORDRT")
        grdSOTORDR1_HOLDS.DataSource = dst.Tables("SOTORDR1_HOLDS")
        grdPOTORDR1.DataSource = dst.Tables("POTORDR1")

        grdSOTORDP1.DataSource = dst.Tables("SOTORDP1")
        grdSOTORDP2.DataSource = dst.Tables("SOTORDP2")

        grdICTSTAT2.DataSource = dst.Tables("ICTSTAT2")

        grdSOTORDC1.DataSource = dst.Tables("SOTORDC1")
        ASCMAIN1.Add_Value_List(grdSOTORDC1, "TRANS_TYPE", Nothing, New String() {":", "C:Credit Card", "D:Deposit", "O:On Account", "A:Additional Funds"})
        ASCMAIN1.Add_Value_List(grdSOTORDC1, "CCPA_STATUS", Nothing, New String() {":", "T:Authorization", "E:Error", "S:Sale/Deposit", "A:Pre-Auth Sale"})

        grdSOTORDRX.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDRX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_DATE", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdICTSTDQ2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"DATE_1", "ADD_1", "DATE_2", "ADD_2", "DATE_3", "ADD_3", "DATE_4", "ADD_4"}
                Dim C As Integer = Val(Mid(COLUMN_NAME, Len(COLUMN_NAME), 1))
                With .Columns(COLUMN_NAME).Header
                    .Appearance.BackColor = Drawing.Color.White
                    .Appearance.BackColor2 = New System.Drawing.Color() {Drawing.Color.LightBlue, Drawing.Color.Pink, Drawing.Color.LightGreen, Drawing.Color.Orange}(C - 1)
                    .Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
            Next
        End With

        With grdICTSTDQ3.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"DATE_1", "QTY_1", "DATE_2", "QTY_2", "DATE_3", "QTY_3", "DATE_4", "QTY_4"}
                Dim C As Integer = Val(Mid(COLUMN_NAME, Len(COLUMN_NAME), 1))
                With .Columns(COLUMN_NAME).Header
                    .Appearance.BackColor = Drawing.Color.White
                    .Appearance.BackColor2 = New System.Drawing.Color() {Drawing.Color.LightBlue, Drawing.Color.Pink, Drawing.Color.LightGreen, Drawing.Color.Orange}(C - 1)
                    .Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
            Next
        End With


        With grdSOTORDRS.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "ORDR_CUST_PO", "TOTAL_AMT", "TOTAL_QTY"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        With grdICTPRICX.DisplayLayout.Bands(0)
            'For Each COLUMN_NAME As String In New String() {"PRICE", "CUST_PRICE"}

            'Next
            .Columns("PRICE").Width = 50
            .Columns("CUST_PRICE").Width = 50
            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                .Columns("CUST_PRICE").Hidden = False
                .Columns("CUST_PRICE").Header.Caption = "Cust"
            Else
                .Columns("CUST_PRICE").Hidden = True
            End If
        End With

        With grdSOTORDRB.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "CUST_STORE_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_CUST_PO"}
                With .Columns(COLUMN_NAME)
                    .Header.Fixed = True
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
        End With

        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_LNO", "STYLE_CODE_SUB", "STYLE_CODE", "STYLE_DESC", "RANGE_STYLE_CODE", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            .Columns("ORDR_UNIT_PRICE").MaskInput = "nnnn.nnnn"
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                .Columns("PO_COST").MaskInput = "nnnn.nnnnnn"
                .Columns("PO_COST").Format = "#.000000"
            End If

            For Each COLUMN_NAME As String In New String() {"ORDR_UNIT_PRICE_CURR", "ORDR_AMT_CURR"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Lime
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next

        End With

        If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
            Create_Summary(grdSOTORDR2, "PF_QTY")
        End If



        If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then

            With grdSOTORDR2.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If gcol.Key = "PF_QTY" Or gcol.Key = "PF_DUTY_HTS_CODE" Or gcol.Key = "PF_ORDER_NO" Then
                        gcol.CellAppearance.BackColor = Drawing.Color.AliceBlue
                        With gcol.Header
                            .Appearance.BackColor = Drawing.Color.White
                            .Appearance.BackColor2 = Drawing.Color.Orange
                            .Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        End With

                    End If
                Next
            End With

        End If


        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"DATE_1", "QTY_1", "AMT_1", "DATE_2", "QTY_2", "AMT_2", "DATE_3", "QTY_3", "AMT_3", "DATE_4", "QTY_4", "AMT_4"}
                If ASCMAIN1.CLIENT = "RGI" Then
                    Dim C As Integer = Val(Mid(COLUMN_NAME, Len(COLUMN_NAME), 1))

                    If COLUMN_NAME.StartsWith("QTY") Then
                        .Columns(COLUMN_NAME).Format = "#,##0"
                        Create_Summary(grdSOTORDR2, COLUMN_NAME)
                    End If
                    If COLUMN_NAME.StartsWith("AMT") Then
                        .Columns(COLUMN_NAME).Format = "#,##0.00"
                        Create_Summary(grdSOTORDR2, COLUMN_NAME)
                    End If
                    With .Columns(COLUMN_NAME).Header
                        .Appearance.BackColor = Drawing.Color.White
                        .Appearance.BackColor2 = New System.Drawing.Color() {Drawing.Color.LightBlue, Drawing.Color.Pink, Drawing.Color.LightGreen, Drawing.Color.Orange}(C - 1)
                        .Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    End With
                Else

                    .Columns(COLUMN_NAME).Hidden = True

                End If
            Next
        End With



        grdSOTORDRH.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSOTORDRH.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdSOTORDRH.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

        With grdSOTORDRH.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC", "ORDR_UNIT_PRICE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_LNO", "STYLE_DESC", "COLOR_DESC"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If Not New String() {"STYLE_CODE", "COLOR_CODE", "ORDR_QTY", "ORDR_UNIT_PRICE", "ORDR_UNIT_PRICE_MANUAL", "RANGE_STYLE_CODE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"CUST_UPC", "CUST_SKU", "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "STYLE_RETAIL"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightPink
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() _
                {"ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC", "RANGE_STYLE_CODE", "STYLE_CODE_SUB",
                 "INNER_PACK_QTY", "CARTON_PACK_QTY", "TOTAL_CARTONS", "CASE_CUBE", "TOTAL_CUBE", "STYLE_UOM", "STYLE_CLASS_CODE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"STYLE_PRICE", "ORDR_UNIT_PRICE", "ORDR_UNIT_PRICE_CALC", "ORDR_UNIT_PRICE_MANUAL", "ORDR_PRICE_SOURCE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"PO_COST", "MU_PCT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_LNO", "STYLE_DESC", "COLOR_DESC"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        With grdSOTORDRI.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If Not New String() {"ORDR_UNIT_PRICE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next
            grdSOTORDRI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            .Columns("COL").Hidden = True

            .Columns("ORDR_QTY").Hidden = True
            .Columns("ORDR_QTY_CANC").Hidden = True
            .Columns("ORDR_AMT").Hidden = True
            .Columns("ORDR_AMT_CANC").Hidden = True

            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE_SUB", "STYLE_KEY", "STYLE_UOM", "CARTON_PACK_QTY", "STYLE_CODE", "COLOR_CODE", "STYLE_DESC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_UPC", "CUST_SKU"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_UNIT_PRICE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With


        With grdSOTORDR3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "ORDR_UNIT_PRICE" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key.StartsWith("SIZE_QTY") Or New String() {"CUST_STYLE_CODE", "CUST_COLOR_CODE", "ORDR_UNIT_PRICE", "ORDR_QTY"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If

                If gcol.Key.StartsWith("SIZE_QTY") Then
                    Create_Summary(grdSOTORDR3, gcol.Key)
                    With gcol
                        .Width = grdSOTORDR3.DisplayLayout.Bands(0).Columns("ORDR_QTY").Width
                        .CellAppearance.TextHAlign = HAlign.Right
                        .Header.Appearance.TextHAlign = HAlign.Right
                    End With
                End If
            Next
        End With

        With grdSOTORDRR.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If
                If gcol.Key = "ORDR_UNIT_PRICE" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_UPC", "CUST_SKU"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.HotPink
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"STYLE_CODE", "COLOR_CODE", "ORDR_QTY", "ORDR_QTY_OPEN"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                ElseIf New String() {"CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_UPC", "CUST_SKU"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next
        End With


        With grdSOTPICK1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "INV_NO" Or gcol.Key = "PICK_SHIPPED" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        With grdSOTPICK1.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "" Or gcol.Key = "" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" <> "1" Then
                    If gcol.Key = "RANGE_STYLE_CODE" Then
                        gcol.Hidden = True
                    End If
                End If
            Next
        End With


        With grdSOTORDP1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "INV_REF" Or gcol.Key = "INV_DATE" Or gcol.Key = "INV_COMMENT" Or gcol.Key = "PICK_CNT_CARTONS" Or gcol.Key = "BILL_OF_LADING_NO" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        With grdSOTORDP2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "ORDR_QTY_SHIP" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                .Columns("ORDR_UNIT_PRICE").Format = "#.000000"
            End If
        End With

        Create_Summary(grdSOTORDRX, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRX, "ORDR_QTY_OPEN", "Sum")
        Create_Summary(grdSOTORDRX, "ORDR_QTY_PICK", "Sum")
        Create_Summary(grdSOTORDRX, "ORDR_QTY_SHIP", "Sum")
        Create_Summary(grdSOTORDRX, "ORDR_TOTAL", "Sum")

        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY",
                                                  "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC",
                                                  "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC",
                                                  "TOTAL_CARTONS", "TOTAL_CUBE"})
        Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY_ALLO_CUR", "ORDR_AMT_ALLO_CUR"})
        Create_Summary(grdSOTORDR2, New String() {"CGS", "GP_AMT"})
        Create_Summary(grdSOTORDR2, New String() {"ORDR_AMT_CURR"})
        Create_Summary(grdSOTORDR2, "GP_PCT", "Custom")
        Create_Summary(grdSOTORDR2, "MU_PCT", "Custom")

        Create_Summary(grdSOTORDRB, "CUST_STORE_NO", "Count")

        Create_Summary(grdSOTORDRI, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRI, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_CANC", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_CANC"})

        Create_Summary(grdSOTORDRR, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRR, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})

        Create_Summary(grdSOTORDR3, "ORDR_SUB_LNO", "Count")
        Create_Summary(grdSOTORDR3, New String() {"ORDR_QTY", "ORDR_AMT"})

        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
        Create_Summary(grdSOTPICK1, New String() {"PICK_FREIGHT", "INV_FREIGHT", "INV_TOTAL_AMOUNT"})
        Create_Summary(grdSOTPICK1, "PICK_LNO", "Count", "SOTPICK1_SOTPICK2")
        Create_Summary(grdSOTPICK1, New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", "PICK_QTY_CANC_REL", "PICK_QTY_BACK_REL", "PICK_TOTAL"}, , "SOTPICK1_SOTPICK2")
        grdSOTPICK1.DisplayLayout.Bands(1).SummaryFooterCaption = "Pick Ticket Totals"

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL", "CART_TOTAL_WGT_CALC"})

        Create_Summary(grdSOTORDRS, "CUST_STORE_NO", "Count")
        Create_Summary(grdSOTORDRS, "TOTAL_AMT", "TOTAL_QTY")

        Create_Summary(grdSOTORDP1, "INV_NO", "Count")
        Create_Summary(grdSOTORDP1, New String() {"INV_TOTAL_AMOUNT"})

        Create_Summary(grdSOTORDP2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDP2, New String() {"ORDR_QTY_SHIP", "ORDR_AMT_SHIP"})



        Bind_Controls(grpBILLTO, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'BT'", "", DataViewRowState.CurrentRows))
        Bind_Controls(grpSOLDTO, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'BY'", "", DataViewRowState.CurrentRows))
        dvwSOTORDR5 = New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'ST'", "", DataViewRowState.CurrentRows)
        Bind_Controls(grpSHIPTO, "SOTORDR5", dvwSOTORDR5)
        Bind_Controls(grpSTORE, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'MK'", "", DataViewRowState.CurrentRows))
        Bind_Controls(grpDC, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'DC'", "", DataViewRowState.CurrentRows))
        Bind_Controls(frmShipToOption, "SOTORDR1")
        Bind_Controls(grpSOLDTO_Attributes, "ARTCUST1")
        Bind_Controls(frmSOTORDRD, "SOTORDR0")

        'Set_Read_Only(grpStatus, True)
        Set_Read_Only(grpBILLTO, True)
        Set_Read_Only(grpSOLDTO, True)
        Set_Read_Only(grpSHIPTO, True)
        Set_Read_Only(grpSTORE, True)
        Set_Read_Only(grpDC, True)

        Set_Read_Only(frmPPQTY, True)

        Set_Read_Only_for_ctl(Absx1.optFor("ORDR_SOURCE"), True)
        Set_Read_Only_for_ctl(optB, True)

        'Absx1.txtFor("ORDR_PRIORITY").Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("X2")
        'Absx1.txtFor("ORDR_PRIORITY").ReadOnly = Not ASCMAIN1.USER_SECURITY_CODEs.Contains("X2")

        With dst.Tables("SOTORDRT").Rows
            .Add(New Object() {1, "Order", 0, 0})
            .Add(New Object() {2, "Open", 0, 0})
            .Add(New Object() {3, "Allo", 0, 0})
            .Add(New Object() {4, "Pick", 0, 0})
            .Add(New Object() {5, "Ship", 0, 0})
            .Add(New Object() {6, "Canc", 0, 0})
        End With
        Sort_grdColumns(grdSOTORDRT, "KEY", True)

        Show_Filter(grdSOTORDRX, True)
        grdSOTORDRX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTORDRB, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTORDR1_HOLDS, "ORDR_REL_HOLD_CODE")

        Check_InquiryMode()

        ASCMAIN1.sql = "Select T_CODE, T_DESC from ASTCODE1" _
            & " where TABLE_NAME = 'SOTORDR1' and COLUMN_NAME = 'ORDR_SOURCE'" _
            & " and T_CODE Not in ('E','K','W','S')"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim T_CODE As String = row.Item("T_CODE")
            Dim T_DESC As String = row.Item("T_DESC")
            Absx1.optFor("ORDR_SOURCE").ValueList.ValueListItems.Add(New ValueListItem(T_CODE, T_DESC))
        Next

        With grdSOTORDR2.DisplayLayout.Bands(0)

            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                .Columns("DUTY_RATE_CODE").Hidden = False
            Else
                .Columns("DUTY_RATE_CODE").Hidden = True
                .Columns("STYLE_UOM").Hidden = True
                .Columns("STYLE_CLASS_CODE").Hidden = True
                .Columns("STYLE_PRICE").Hidden = True
                .Columns("ORDR_UNIT_PRICE_CALC").Hidden = True
                .Columns("ORDR_UNIT_PRICE_MANUAL").Hidden = True
                .Columns("ORDR_QTY_ALLO_CUR").Hidden = True
                .Columns("ORDR_AMT_ALLO_CUR").Hidden = True
            End If
        End With

        If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
            splStyleStatus.Panel2Collapsed = True
        ElseIf (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then
            tabDetails.Tabs("Pricing && Availability").Text = "Available to Sell"
            splPA.Panel1Collapsed = True
        Else
            tabDetails.Tabs("Pricing && Availability").Visible = False
            splStyleStatus.Panel2Collapsed = True
        End If

        dteSearchS.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dteSearchE.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)

        dteSearchS.DateTime = DateAdd(DateInterval.Month, -1, DateTime.Now)
        dteSearchE.DateTime = DateTime.Now

        clsSOCORDR1 = New TAC.SOCORDR1(Me)

        grpBILLTO_Attributes.Visible = False ' UNTIL WE HAVE SOMETHING TO PUT IN THERE
        chkCUST_ORDR_CALL_B4_SHIPPING.Visible = (ASCMAIN1.CLIENT = "RGI")
        lblORDR_NO_WEB.Visible = (ASCMAIN1.CLIENT = "RGI")
        txtORDR_NO_WEB.Visible = (ASCMAIN1.CLIENT = "RGI")
        chkORDR_INCL_VAS.Visible = (ASCMAIN1.CLIENT = "NYA")

        dteORDR_REL_ACTION_DATE.Visible = (ASCMAIN1.CLIENT = "RGI")
        lblORDR_REL_ACTION_DATE.Visible = (ASCMAIN1.CLIENT = "RGI")
        SplitContainer4.Panel2Collapsed = Not (ASCMAIN1.CLIENT = "RGI")

        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
            tabSOTORDRX.Tabs("Pro-Forma Invoices").Visible = False
        End If

        chkShortView.Visible = (ASCMAIN1.CLIENT = "RGI")
        MakeTransparent(chkShortView)

        If ASCMAIN1.CLIENT = "RGI" Then
            chkCUST_FACTOR_IND.Enabled = False
            grpBuyerInfo.Visible = True
        Else
            grpBuyerInfo.Visible = False
        End If
    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFORDRI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Dim rowARTCUST2 As DataRow = Nothing
                multiple_order_maintenance = False
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
                        If CUST_BILL_TO_CUST = "" Then
                            CUST_BILL_TO_CUST = CUST_CODE
                        End If

                        Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
                        If rowARTCUST1_BT Is Nothing Then
                            EMsg &= vbCr & "Unable to determine Bill-To Customer"
                        Else
                            If rowARTCUST1_BT.Item("POST_CODE") & "" = "" Then
                                EMsg &= vbCr & "No value specified for the Post Code for Bill-To Customer " & CUST_BILL_TO_CUST
                            Else
                                If LookUp("ARTPOST1", rowARTCUST1_BT.Item("POST_CODE")) Is Nothing Then
                                    EMsg &= vbCr & "Invalide AR Post Code specified for Bill-To Customer " & CUST_BILL_TO_CUST
                                End If
                            End If
                        End If

                        ' apostrophe in Cust PO causes ABSolution to crash when lookig to see if it is a duplicate PO entry
                        Absx1.txtFor("ORDR_CUST_PO").Text = Absx1.txtFor("ORDR_CUST_PO").Text.Trim.Replace("'", "")
                        ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                        If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                            EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                        End If

                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If


                If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "B2C" Then
                    ' disabling this block email Jennifer 9/4/19
                    ' EMsg &= vbCr & "You may not manually enter an eCommerce order"
                    ' see wjz email to ed on 04/07/2019, lock was put in place because SOFORDR1 does not price with ecom rules
                    ' changes have been made in SOFORDR1 for B2C orders to not change price, disable adding & deleting lines
                End If

                If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR" Then
                    ' Absx1.txtFor("CUST_STORE_NO").Text = "000000"
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE_TO").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & "No value specified for Transfer-To Warehouse"
                    Else
                        If rowICTWHSE1.Item("WHSE_TYPE") & "" <> "W" Then
                            EMsg &= vbCr & "Invalid Warehouse Type for Warehouse Transfer (" & rowICTWHSE1.Item("WHSE_TYPE") & ")"
                        End If

                        If ASCMAIN1.CLIENT = "NYA" Then
                            If Absx1.txtFor("WHSE_CODE_TO").Text = "21" Then
                                EMsg &= vbCr & "Warehouse Transfers not allowed for NYAG Candada - need to set up Intercompany Sale"
                            End If
                        End If

                    End If
                    If Absx1.txtFor("CUST_STORE_NO").Text & "" = "" Then
                        Absx1.txtFor("CUST_STORE_NO").Text = "000000"
                    End If
                End If
                If Absx1.txtFor("CUST_STORE_NO").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Store (Mark-For)"
                Else
                    rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", Absx1.txtFor("CUST_STORE_NO").Text})
                    If rowARTCUST2 Is Nothing AndAlso Absx1.txtFor("CUST_STORE_NO").Text = "000000" Then
                        rowARTCUST2 = Make_ARTCUST2_BTST()
                    End If
                    If rowARTCUST2 IsNot Nothing Then
                        CUST_STORE_NO = Absx1.txtFor("CUST_STORE_NO").Text
                    Else
                        EMsg &= vbCr & "No Record of Customer Store " & Absx1.txtFor("CUST_STORE_NO").Text
                    End If
                End If


                If EMsg = "" Then
                    ' Check Copy From Order to verify that it exists, and that it is with same customer
                    If Absx1.txtFor("ORDR_NO").Text <> "" Then
                        Dim rowSOTORDR1_CopyFrom As DataRow = LookUp("SOTORDR1", Absx1.txtFor("ORDR_NO").Text)
                        If rowSOTORDR1_CopyFrom Is Nothing Then
                            EMsg &= vbCr & "No Record of (Copy From) Order No " & Absx1.txtFor("ORDR_NO").Text
                        ElseIf rowSOTORDR1_CopyFrom.Item("CUST_CODE") <> CUST_CODE Then
                            EMsg &= vbCr & "Copy Order Feature works with Same Customer Only"
                        End If
                    End If

                    ' Customer must have a Sales Rep assigned
                    SREP_CODE = rowARTCUST1.Item("SREP_CODE") & ""
                    Dim rowSOTSREP1 As DataRow = Nothing
                    If SREP_CODE <> "" Then rowSOTSREP1 = LookUp("SOTSREP1", SREP_CODE)
                    If rowSOTSREP1 Is Nothing Then
                        EMsg &= vbCr & "This Customer Has No Sales Rep Assigned"
                    End If
                    SREP2_CODE = rowARTCUST1.Item("SREP2_CODE") & ""
                End If


                If EMsg = "" Then
                    ' Load Default values in for Selected Fields if we have seen this Customer PO before
                    ASCMAIN1.sql = "Select ORDR_GROUP_NO, CUST_STORE_NO, ORDR_SHIP_DATE, ORDR_CANCEL_DATE" & vbCrLf _
                        & ", ORDR_DATE, ORDR_DEPT, ORDR_SHIP_INSTR, FRT_TERMS, SALES_DIVISION_CODE" & vbCrLf _
                        & " from SOTORDR1 where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & " and ORDR_CUST_PO = '" & ORDR_CUST_PO & "'" & vbCrLf _
                        & " order by ORDR_SHIP_DATE DESC"
                    ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where ROWNUM < 2"
                    rowSOTORDR1po = ASCDATA1.GetDataRow

                    If rowSOTORDR1po IsNot Nothing Then
                        ASCMAIN1.sql = "Select ORDR_NO, ORDR_DATE from SOTORDR1 " & vbCrLf _
                            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                            & "   and CUST_STORE_NO = '" & CUST_STORE_NO & "'" & vbCrLf _
                            & "   and ORDR_CUST_PO = '" & ORDR_CUST_PO & "'" & vbCrLf _
                            & "   and ORDR_STATUS in ('O','P','F')"
                        Dim rowDup As DataRow = ASCDATA1.GetDataRow
                        If rowDup IsNot Nothing Then
                            If MsgBox("Same Customer PO has already been entered for Store " & CUST_STORE_NO _
                                      & vbCrLf & " (See Sales Order " & rowDup.Item("ORDR_NO") & " dated " & Format(rowDup.Item("ORDR_DATE"), "MM/dd/yyyy") & ")" _
                                      & vbCrLf & vbCrLf & "Are You Sure that you want to Proceed?", MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                                      "Possible Order Duplication") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If


                If EMsg = "" Then
                    If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                        If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "BTB" Then
                            Dim iBTB_TYPE As Integer = -1
                            Dim WHSE_CODEs() As String = {"FE", "FD", "SP", "NY", "ZZ", "NC"}
                            Using frmASFMSGBF As New ASFMSGBF

                                iBTB_TYPE = frmASFMSGBF.Get_opt_from_User("Select Type of Back-to-Back Order", WHSE_CODEs, 0, "Once Selected, this may not be changed")
                            End Using
                            If iBTB_TYPE = -1 Then Exit Sub
                            BTB_TYPE = WHSE_CODEs(iBTB_TYPE)
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                        ' NO RESERVATIONS
                    Else
                        If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                    End If
                    'If Not ASCMAIN1.Logical_Lock("SOTORDR1", CUST_CODE) Then Exit Sub

                End If

            Case "Edit", "View"
                multiple_order_maintenance = False

                CUST_CODE = ""
                ORDR_NO = ""
                If Absx1.txtFor("ORDR_NO").Text = "" Then
                    EMsg &= vbCr & "No Order No Specified"
                Else
                    ORDR_NO = Absx1.txtFor("ORDR_NO").Text
                    rowSOTORDR1 = LookUp("SOTORDR1", ORDR_NO)
                    If rowSOTORDR1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Sales Order No " & ORDR_NO
                    Else
                        ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO")
                        ORDR_GROUP_NOs.Clear()
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                        CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
                        EDI_JRNL_NO = rowSOTORDR1.Item("EDI_JRNL_NO") & ""

                        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                            If Not TAC.TACMAIN1.NyaCanadaWhseList.Contains(rowSOTORDR1.Item("WHSE_CODE") & "") Then ' <> "18" Then
                                EMsg &= vbCr & "Invalid Order, must be in warehouse(s) (" & TAC.TACMAIN1.NyaCanadaWhseCommaSeparatedString & ")"
                            End If

                        End If

                        ASCMAIN1.sql = "Select Count (*) ORDR_CNT" & vbCrLf _
                            & ", SUM(DECODE(ORDR_STATUS,'O',1,0)) O" _
                            & ", SUM(DECODE(ORDR_STATUS,'P',1,0)) P" _
                            & ", SUM(DECODE(ORDR_STATUS,'C',1,0)) C" _
                            & ", SUM(DECODE(ORDR_STATUS,'F',1,0)) F" _
                            & " from SOTORDR1" & vbCrLf _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                            & "   and ORDR_STATUS in ('O','C','P','F')"
                        Dim rowSTATS As DataRow = ASCDATA1.GetDataRow
                        If Val(rowSTATS.Item("ORDR_CNT") & "") > 1 Then
                            multiple_order_maintenance = True
                            multiple_order_type = "ORDR_GROUP_NO"
                        End If

                        If Not multiple_order_maintenance And EDI_JRNL_NO <> "" Then
                            ASCMAIN1.sql = "Select Count (*) ORDR_CNT" & vbCrLf _
                               & ", SUM(DECODE(ORDR_STATUS,'O',1,0)) O" _
                               & ", SUM(DECODE(ORDR_STATUS,'P',1,0)) P" _
                               & ", SUM(DECODE(ORDR_STATUS,'C',1,0)) C" _
                               & ", SUM(DECODE(ORDR_STATUS,'F',1,0)) F" _
                               & " from SOTORDR1" & vbCrLf _
                               & " where EDI_JRNL_NO = '" & EDI_JRNL_NO & "'" & vbCrLf _
                               & "   and ORDR_STATUS in ('O','C','P','F')"
                            rowSTATS = ASCDATA1.GetDataRow
                            If Val(rowSTATS.Item("ORDR_CNT") & "") > 1 Then
                                multiple_order_maintenance = True
                                multiple_order_type = "EDI_JRNL_NO"
                            End If
                        End If

                        'multiple_order_maintenance = False ' UNTIL WE GET SOTORDRB/Q ROUTINES DONE RIGHT

                        If multiple_order_maintenance Then
                            ' PROBABLY WILL NEED A REVERSE CANCELLATION OPTION FOR ORDERS IN A MOG
                            If Not InquiryMode Then
                                If Val(rowSTATS.Item("O") & "") = 0 Then
                                    EMsg &= vbCr & "Sales Order No " & ORDR_NO & " belongs to a Multiple-Order Group" _
                                        & vbCr & "- No Orders are Open in that group"
                                End If
                            End If
                        Else
                            If rowSOTORDR1.Item("ORDR_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                                Select Case rowSOTORDR1.Item("ORDR_STATUS")
                                    Case "C" ' for cancelled orders, the Edit command is not enabled (currently), so we will never use this procedure, which appears to be useful for multiple store orders as well.  For now, we will just use the Reverse_Cancel method for a singel order in the Reinstate Cancelled command.
                                        MsgBox("Sales Order No " & ORDR_NO & " has been Cancelled", MsgBoxStyle.OkOnly,
                                                "Cannot Edit Order")
                                        If MsgBox("Re-Open Order for Processing", MsgBoxStyle.YesNo,
                                                    "Answer Yes to Reverse Cancellation of this Order") = MsgBoxResult.No Then
                                            Exit Sub
                                        Else
                                            Select Case MsgBox("Reverse Cancellation of All Orders in this Group",
                                                                MsgBoxStyle.YesNoCancel,
                                                                "Yes for All Orders in Group, No for this order only")
                                                Case MsgBoxResult.Yes
                                                    ASCMAIN1.sql = "Select ORDR_NO, CUST_STORE_NO from SOTORDR1" _
                                                        & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                                                        & "   and ORDR_STATUS = 'C'"
                                                    Dim dt As DataTable = ASCDATA1.GetDataTable
                                                    Using F As New ASFMSGBF
                                                        F.Show_grd(dt, Me, "The following Orders will be Restored")
                                                        If F.user_option = -1 Then
                                                            MsgBox("Restoration of Cancelled Orders was NOT Performed", MsgBoxStyle.OkOnly, "Please Note")
                                                        Else
                                                            If ASCMAIN1.Running_in_VS Then Stop ' send in list of orders
                                                            Reverse_Cancel("", ORDR_GROUP_NO, dt)
                                                            MsgBox("Restoration of Cancelled Orders is Complete", MsgBoxStyle.OkOnly, "Please Note")
                                                        End If
                                                    End Using

                                                Case MsgBoxResult.No
                                                    Reverse_Cancel(ORDR_NO, "")

                                                Case MsgBoxResult.Cancel
                                                    Exit Sub
                                            End Select

                                        End If

                                    Case "D"
                                        EMsg &= vbCr & "Sales Order No " & ORDR_NO & " has been Deleted"
                                    Case "P"
                                        EMsg &= vbCr & "Sales Order No " & ORDR_NO & " has been Completely Released for Picking"
                                    Case Else ' such as "F"
                                        EMsg &= vbCr & "Sales Order No " & ORDR_NO & " is No Longer Open"
                                End Select
                            End If
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then

                    If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "B2C" Then
                        'EMsg &= vbCr & "You may not manually edit an eCommerce order"
                    End If

                    'ASCMAIN1.sql = "Select Distinct PO_ORDER_NO from POTORDR2" _
                    '    & " where ORDR_NO = '" & ORDR_NO & "' and (PO_QTY_SHP <> 0 or PO_QTY_REC <> 0)"
                    'For Each rowPOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
                    '    Dim PO_ORDER_NO As String = rowPOTORDR2.Item("PO_ORDER_NO")
                    '    EMsg &= vbCr & "Purchase Order No " & PO_ORDER_NO & " has been Shipped"
                    'Next

                    disable_update = False
                    'ASCMAIN1.sql = "Select Count (*) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS in ('O','P','F')"
                    'If Val(ASCDATA1.GetDataValue) > 1 Then
                    '    If MsgBox("You Cannot Make Changes to an Order belonging to a Multiple Order Group" _
                    '              & vbCr & vbCr & "However, you Can Delete or Cancel this Order (Store) using this Option.", _
                    '              MsgBoxStyle.OkCancel, _
                    '              "Note: Order " & ORDR_NO & " is part of Order Group (" & ORDR_GROUP_NO & ")") = MsgBoxResult.Cancel Then
                    '        Exit Sub
                    '    End If
                    '    disable_update = True
                    'End If
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub

                    If multiple_order_maintenance Then
                        ASCMAIN1.sql = "Select ORDR_NO" & vbCrLf _
                           & " from SOTORDR1" & vbCrLf _
                           & IIf(multiple_order_type = "ORDR_GROUP_NO",
                                 " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'",
                                 " where EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") & vbCrLf _
                           & "   and ORDR_STATUS in ('O','C','P','F')"
                        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                            If Not ASCMAIN1.Logical_Lock("SOTORDR1", row.Item("ORDR_NO")) Then
                                multiple_order_maintenance = False
                                Exit Sub
                            End If
                        Next
                    End If

                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE) Then Exit Sub
                        'AndAlso (rowSOTORDR1.Item("WHSE_CODE") & "" <> "NY") 
                        If ASCMAIN1.CLIENT = "RGI" AndAlso (rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB") Then
                            ' no need to MT lock release if BTB and FD/FE
                        Else
                            If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub
                        End If

                        If ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then
                            If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB" Then
                                If Not ASCMAIN1.Logical_Open("F", "POFCENT1") Then Exit Sub
                                ASCMAIN1.sql = "Select PO_ORDER_NO from POTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
                                For Each rowPOTORDR1 As DataRow In ASCDATA1.GetDataTable.Rows
                                    Dim PO_ORDER_NO As String = rowPOTORDR1.Item("PO_ORDER_NO")
                                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                                Next
                            End If
                        End If

                    End If
                End If


            Case "Pro-Forma"

                If ASCMAIN1.CLIENT = "VAN" And InquiryMode And Absx1.txtFor("PO_SHIPMENT_NO").Text <> "" Then '  DGJ AndAlso PROFORMA
                    If LookUp("POTSHIP1", Absx1.txtFor("PO_SHIPMENT_NO").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Shipment Number for Pro-Forma"
                    End If
                End If
                If ASCMAIN1.CLIENT = "VAN" Then
                    If Absx1.numFor("PF_CARTONS").Value = 0 Then
                        EMsg &= vbCr & "Must enter Cartons for a Pro-Forma to print"
                    End If
                    If Absx1.numFor("PF_WEIGHT").Text = 0 Then
                        EMsg &= vbCr & "Must enter Weight for a Pro-Forma to print"
                    End If
                    If Absx1.optFor("PF_WEIGHT_UOM").Value & "" = "" Then
                        EMsg &= vbCr & "Must have a Pro-Forma Weight Unit of Measure"
                    End If
                    If Absx1.optFor("PF_OVERSEAS_DOMESTIC").Value & "" = "" Then
                        EMsg &= vbCr & "Must have Selected Overseas or Domestic for a Pro-Forma to print"
                    End If
                    If Absx1.optFor("PF_WHSE_FACTORY").Value & "" = "" Then
                        EMsg &= vbCr & "Must have Selected From Whse or Factory for a Pro-Forma to print"
                    End If
                    If Absx1.optFor("PF_OVERSEAS_DOMESTIC").Value & "" = "O" And Absx1.optFor("PF_WHSE_FACTORY").Value & "" = "F" And Absx1.txtFor("PO_SHIPMENT_NO").Text = "" Then
                        EMsg &= vbCr & "You must Select a Shipment No for an Overseas Pro-Forma to print"
                    End If


                End If

            Case "Update"

                Dim TERM_TYPE As String = String.Empty

                ' Need to see if the customer requires details ship complete or not at all.
                ' Currently Regency - QVC.com (310921)
                If EntryMode = "E" Then
                    If rowARTCUST1 IsNot Nothing Then
                        If rowARTCUST1.Table.Columns.Contains("CUST_SHIP_COMPLETE_DETAIL") Then
                            If rowARTCUST1.Item("CUST_SHIP_COMPLETE_DETAIL") & String.Empty = "1" Then
                                'ORDR_QTY, ORDR_QTY_ORIG
                                For Each row As DataRow In dst.Tables("SOTORDR2").Select("ISNULL(ORDR_QTY, 0) <> ISNULL(ORDR_QTY_ORIG, 0) and ISNULL(ORDR_QTY, 0) <> 0")
                                    EMsg &= vbCr & "Item (" & row.Item("STYLE_CODE") & " / " & row.Item("COLOR_CODE") & ") must ship complete or not at all. Customer does not allow partial detail line shipments." _
                                        & " Original quantity is " & row.Item("ORDR_QTY_ORIG") & "."
                                Next
                            End If
                            If EMsg.Length > 0 Then Exit Select
                        End If
                    End If
                End If

                If EntryMode = "M" Or multiple_order_maintenance Then
                    '   If ASCMAIN1.Running_in_VS Then Stop
                    Dim ORDR_SHIP_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value
                    Dim ORDR_CANCEL_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_CANCEL_DATE").Value
                    'Dim ORDR_HOLD As String = IIf(Absx1.chkFor("SOTORDR0.ORDR_HOLD").Checked, "1", "0")
                    'Dim WHSE_CODE As String = Absx1.dteFor("SOTORDR0.WHSE_CODE").Value

                    If Format(ORDR_SHIP_DATE, "yyyyMMdd") > Format(ORDR_CANCEL_DATE, "yyyyMMdd") Then
                        EMsg &= vbCr & "Ship Date Cannot be Later than Cancel Date"
                    End If

                    If Not multistore_OK_TO_UPDATE Then
                        EMsg &= vbCr & "This Order may not be updated using Multiple Order Maintenance"
                    End If

                    If Absx1.dteFor("SOTORDR0.ORDR_ARRIVAL_DATE").Value & "" <> "" And Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value & "" <> "" Then
                        If Format(Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value, "yyyyMMdd") > Format(Absx1.dteFor("SOTORDR0.ORDR_ARRIVAL_DATE").Value, "yyyyMMdd") Then
                            EMsg &= vbCr & "Arrival date must not be prior to Ship-By Date"
                        End If
                    End If

                    Dim rowICTWHSE1 As DataRow = Nothing
                    Dim WHSE_CODE As String = Absx1.txtFor("SOTORDR0.WHSE_CODE").Text
                    If WHSE_CODE = "" Then
                        EMsg &= vbCr & "Warehouse is required"
                    Else
                        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                        If rowICTWHSE1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Warehouse Code"
                        End If
                    End If

                    Dim FRT_TERMS As String = Absx1.txtFor("SOTORDR0.FRT_TERMS").Text
                    If FRT_TERMS = "" Then
                        EMsg &= vbCr & "Freight Terms are Mandatory"
                    Else
                        Dim row As DataRow = LookUp("ASTCODE1", New String() {"SOTORDR1", "FRT_TERMS", FRT_TERMS})
                        If row Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Freight Terms"
                        Else

                        End If
                    End If


                    Dim TERM_CODE As String = Absx1.txtFor("SOTORDR0.TERM_CODE").Text
                    If TERM_CODE = "" Then
                        EMsg &= vbCr & "Terms Code is required"
                    Else
                        LookUp("TATTERM1", TERM_CODE)
                        If cdr Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Terms Code"
                        Else
                            TERM_TYPE = cdr.Item("TERM_TYPE") & String.Empty
                        End If
                    End If

                    If Absx1.chkFor("SOTORDR0.CUST_FACTOR_IND").Checked Then
                        If TERM_TYPE = "C" Then EMsg &= vbCr & "Cannot Factor with Terms Code " & TERM_CODE
                        'If ORDR_AMT = 0 Then EMsg &= vbCr & "Cannot Factor with $0 Order"
                    End If

                    Dim SREP_CODE As String = Absx1.txtFor("SOTORDR0.SREP_CODE").Text
                    If SREP_CODE = "" Then
                        EMsg &= vbCr & "Sales Rep is required"
                    Else
                        LookUp("SOTSREP1", SREP_CODE)
                        If cdr Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Sales Rep Code"
                        End If
                    End If

                    Dim SREP2_CODE As String = Absx1.txtFor("SOTORDR0.SREP2_CODE").Text
                    If SREP2_CODE = "" Then
                    Else
                        LookUp("SOTSREP1", SREP2_CODE)
                        If cdr Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Sales Rep2 Code"
                        End If
                    End If

                    Dim SHIP_VIA_CODE As String = Absx1.txtFor("SOTORDR0.SHIP_VIA_CODE").Text
                    If SHIP_VIA_CODE = "" Then
                        EMsg &= vbCr & "Ship Via is required"
                    Else
                        Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                        If rowSOTSVIA1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Ship Via Code"
                        Else
                            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                                If rowSOTSVIA1.Item("SHIP_VIA_STATUS") & "" <> "A" Then
                                    EMsg &= vbCr & "Ship Via Code specified is Inactive"
                                End If

                                If rowICTWHSE1.Item("WHSE_EDI_ID") & "" <> "" And SHIP_VIA_CODE <> "ROUT" Then
                                    ASCMAIN1.sql = "Select * from EDTXREF3" _
                                            & " where SENDER_ID_QUAL = :PARM1 and SENDER_ID = :PARM2 and SHIP_VIA_CODE = :PARM3"
                                    Dim rowEDTXREF3 As DataRow = ASCDATA1.GetDataRow _
                                                            (ASCMAIN1.sql, "VVV", New String() _
                                                             {rowICTWHSE1.Item("WHSE_EDI_QUAL"),
                                                              rowICTWHSE1.Item("WHSE_EDI_ID"),
                                                              SHIP_VIA_CODE})
                                    If rowEDTXREF3 Is Nothing Then
                                        EMsg &= vbCr & "Ship Via Code does not translate to a Valid Service Level for Whse " & WHSE_CODE
                                    End If

                                    If FRT_TERMS = "COL" Or FRT_TERMS = "3PY" Then
                                        Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & ""
                                        If CARRIER_CODE = "UPS" AndAlso rowARTCUST1.Item("UPS_ACCT_NO") & "" = "" Then EMsg &= vbCr & "No UPS Account set up for Customer"
                                        If CARRIER_CODE = "FEDEX" AndAlso rowARTCUST1.Item("FDX_ACCT_NO") & "" = "" Then EMsg &= vbCr & "No Fedex Account set up for Customer"
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                        EMsg &= vbCr & "Customer PO is required"
                    End If


                Else

                    If ASCMAIN1.CLIENT = "NYA" And EntryMode = "E" And Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", txtCUST_STORE_NO.Text})
                        If rowARTCUST2 Is Nothing Then
                            EMsg &= vbCr & "Invalid Customer Store"
                        End If
                        If txtCUST_STORE_NO.Text <> CUST_STORE_NO Then
                            EMsg &= vbCr & "Issue with Customer Store - internal Memory does not agree with Control Value"
                        End If
                    End If

                    Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

                    If FRT_TERMS = "" Then
                        EMsg &= vbCr & "Freight Terms are Mandatory"
                    Else
                        Dim row As DataRow = LookUp("ASTCODE1", New String() {"SOTORDR1", "FRT_TERMS", FRT_TERMS})
                        If row Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Freight Terms"
                        Else

                            If FRT_TERMS = "PPA" Then
                                If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR" Then
                                    EMsg &= vbCr & "Invalid Value Specified for Freight Terms for a Transfer Order"
                                End If
                            End If

                            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                                If FRT_TERMS = "PPD" Or FRT_TERMS = "PPA" Then
                                    If Absx1.dteFor("ORDR_ARRIVAL_DATE").Value & "" = "" Then
                                        Absx1.dteFor("ORDR_ARRIVAL_DATE").Value = Absx1.dteFor("ORDR_CANCEL_DATE").Value
                                    End If
                                    If Absx1.dteFor("ORDR_ARRIVAL_DATE").Value & "" = "" Then
                                        EMsg &= vbCr & "Arrival Date is Mandatory for Frt Terms " & FRT_TERMS
                                    End If
                                End If

                                'If WHSE_CODE = "FE" And FRT_TERMS <> "COL" Then EMsg &= vbCr & "Frt Terms should be COL for FE Orders"
                                If WHSE_CODE = "FD" And FRT_TERMS <> "PPD" Then EMsg &= vbCr & "Frt Terms should be PPD for FD Orders"
                                If WHSE_CODE = "SP" And FRT_TERMS <> "PPD" Then EMsg &= vbCr & "Frt Terms should be PPD for SP Orders"
                                If WHSE_CODE = "FA" And FRT_TERMS <> "PPA" Then EMsg &= vbCr & "Frt Terms should be PPA for FA Orders"
                            End If
                        End If
                    End If

                    Validate_Code("ORDR_PRIORITY")
                    'If Not Validate_Code("ORDR_PRIORITY") Then
                    '    EMsg &= vbCr & "Invalid Value Specified for Order Priority"
                    'End If

                    If Absx1.dteFor("ORDR_ARRIVAL_DATE").Value & "" <> "" And Absx1.dteFor("ORDR_SHIP_DATE").Value & "" <> "" Then
                        If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") > Format(Absx1.dteFor("ORDR_ARRIVAL_DATE").Value, "yyyyMMdd") Then
                            EMsg &= vbCr & "Arrival date must not be prior to Ship-By Date"
                        End If
                    End If

                    ' WHAT IS THIS SECTION DOING?
                    Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
                    If grdSOTORDR2.Rows.Count = 0 Then
                        Setup_SubGrid(False, True)
                    Else
                        If grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                            Setup_SubGrid(True, False)
                        Else
                            Setup_SubGrid(False, False)
                        End If
                    End If

                    If Absx1.chkFor("ORDR_HOLD").Checked Then
                        If Absx1.txtFor("ORDR_HOLD_REASON").Text = "" Then
                            EMsg &= vbCr & "You must specify a reason why you are placing this order On Hold"
                        End If
                    End If

                    If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                        Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                    Else
                        If EMsg = "" And Absx1.optFor("ORDR_SOURCE").Value <> "E" Then
                            If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                                 > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                                EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                            End If
                        End If
                    End If

                    Dim ORDR_AMT As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
                    Dim ORDR_QTY_ORIG As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "") & "")

                    If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR" Then
                        If ORDR_AMT <> 0 Then
                            EMsg &= vbCr & "Transfer Order may NOT have Revenue Associated"
                        End If

                        If ASCMAIN1.CLIENT = "NYA" Then
                            Dim rowFrom As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                            Dim rowTo As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE_TO").Text)
                            If rowFrom Is Nothing Then
                                EMsg &= vbCr & "Please check Transfer From Warehouse Codes - invalid"
                            ElseIf rowTo Is Nothing Then
                                EMsg &= vbCr & "Please check Transfer To Warehouse Codes - invalid"
                            Else
                                If rowFrom.Item("SEG4_CODE") & "" <> rowTo.Item("SEG4_CODE") & "" Then
                                    EMsg &= vbCr & "Warehouse Transfers not allowed between Warehouses from Different Companies"
                                End If
                            End If
                        End If
                    End If

                    If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                        If System.Math.Round(TOTAL_ORDR_AMT, 2) <> System.Math.Round(ORDR_AMT, 2) Then
                            Dim EDI_Value_Remark As String
                            Using F As New ASFMSGBF
                                EDI_Value_Remark = F.Get_txt_from_User(String.Format(
                                        "Total Value Of Order ({0}) Does Not Equal Original EDI Amount ({1})",
                                            Format(ORDR_AMT, "#,###0.00"),
                                            Format(TOTAL_ORDR_AMT, "#,###0.00")),
                                        "Please Provide A Reason For This Change In Order To Proceed", False, 50)
                            End Using

                            If EDI_Value_Remark = "" Then
                                EMsg &= vbCr & "EDI Value Change Canceled Or Reason Not Provided."
                            Else
                                rowSOTORDR1.Item("EDI_VALUE_CHANGE_REMARK") = EDI_Value_Remark
                                rowSOTORDR1.Item("EDI_VALUE_CHANGE_OPER") = ASCMAIN1.USER_ID
                                rowSOTORDR1.Item("EDI_VALUE_CHANGE_DATE") = DATETIME_STAMP
                            End If

                        ElseIf System.Math.Round(TOTAL_ORDR_QTY, 2) <> System.Math.Round(ORDR_QTY_ORIG, 2) Then
                            'Dim EDI_Value_Remark As String
                            'Using F As New ASFMSGBF
                            '    EDI_Value_Remark = F.Get_txt_from_User(String.Format( _
                            '        "Total Qty Of Order ({0}) Does Not Equal Original EDI Qty ({1})", _
                            '            Format(ORDR_QTY_ORIG, "#,###0.00"), _
                            '            Format(TOTAL_ORDR_QTY, "#,###0.00")), _
                            '        "Please Provide A Reason For This Change In Order To Proceed", False, 50)
                            'End Using

                            'If EDI_Value_Remark = "" Then
                            '    EMsg &= vbCr & "EDI Value Change Canceled Or Reason Not Provided."
                            'Else
                            '    rowSOTORDR1.Item("EDI_VALUE_CHANGE_REMARK") = EDI_Value_Remark
                            '    rowSOTORDR1.Item("EDI_VALUE_CHANGE_OPER") = ASCMAIN1.USER_ID
                            '    rowSOTORDR1.Item("EDI_VALUE_CHANGE_DATE") = DATETIME_STAMP
                            'End If
                        End If

                        Using dt As New DataTable
                            dt.Columns.Add("LNO", GetType(System.Int64))
                            dt.Columns.Add("TYPE")
                            dt.Columns.Add("ORDR_QTY", GetType(System.Int64))
                            dt.Columns.Add("ORDR_QTY_ORIG", GetType(System.Int64))
                            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select _
                                        ("ISNULL(ORDR_QTY,0) <> ISNULL(ORDR_QTY_ORIG,0) and RANGE_STYLE_CODE is Null")
                                dt.Rows.Add(New Object() {rowSOTORDR2.Item("ORDR_LNO"),
                                                                "Style " & rowSOTORDR2.Item("STYLE_CODE"),
                                                                rowSOTORDR2.Item("ORDR_QTY"),
                                                                rowSOTORDR2.Item("ORDR_QTY_ORIG")})
                            Next
                            For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select _
                                        ("ISNULL(ORDR_QTY,0) <> ISNULL(ORDR_QTY_ORIG,0)")
                                dt.Rows.Add(New Object() {rowSOTORDRR.Item("ORDR_LNO"),
                                                                "Range " & rowSOTORDRR.Item("RANGE_STYLE_CODE"),
                                                                rowSOTORDRR.Item("ORDR_QTY"),
                                                                rowSOTORDRR.Item("ORDR_QTY_ORIG")})
                            Next
                            If dt.Rows.Count <> 0 Then
                                Using F As New ASFMSGBF
                                    F.Show_grd(dt, Me,
                                                   "The following lines have an Order Qty that has changed from the Original Qty",
                                                   "Please Verify that it is OK to Continue")
                                    If F.user_option = -1 Then
                                        Exit Sub
                                    End If
                                End Using
                            End If
                        End Using
                    End If

                    If multi_store_is_active Then
                        If Absx1.optFor("ORDR_ADDR_TYPE_ST").Value = "DC" Then
                            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                                Dim CUST_STORE_NO As String = rowSOTORDRS.Item("CUST_STORE_NO") & ""
                                Dim rowARTCUST3 As DataRow = LookUp _
                                        ("ARTCUST3", New String() {CUST_CODE, "MK", CUST_STORE_NO, "DC"})
                                If rowARTCUST3 Is Nothing Then
                                    EMsg &= vbCr & "No DC set up for Store " & CUST_STORE_NO
                                End If
                            Next
                        End If
                    End If

                    If Absx1.txtFor("TERM_CODE").Text = "" Then
                        EMsg &= vbCr & "Terms Code is required"
                    Else
                        Validate_Code("TERM_CODE")
                        If cdr IsNot Nothing Then
                            TERM_TYPE = cdr.Item("TERM_TYPE") & String.Empty
                        End If
                    End If

                    If Absx1.chkFor("CUST_FACTOR_IND").Checked Then
                        If TERM_TYPE = "C" Then EMsg &= vbCr & "Cannot Factor with Terms Code " & Absx1.txtFor("TERM_CODE").Text
                        If ORDR_AMT = 0 Then EMsg &= vbCr & "Cannot Factor with $0 Order"
                    End If

                    If Absx1.txtFor("SREP_CODE").Text = "" Then
                        EMsg &= vbCr & "Sales Rep is required"
                    Else
                        Validate_Code("SREP_CODE")
                    End If

                    Validate_Code("SREP2_CODE", False, True)

                    Dim rowICTWHSE1 As DataRow = Nothing
                    If WHSE_CODE = "" Then
                        EMsg &= vbCr & "Whse is required"
                    Else
                        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                        If rowICTWHSE1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Whse Code"
                        End If
                    End If

                    If ASCMAIN1.CLIENT = "NYA" Then
                        If rowICTWHSE1 IsNot Nothing AndAlso rowICTWHSE1.Item("LP_CODE") & "" = "TSI" Then
                            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, "ST"})
                            If rowSOTORDR5 Is Nothing Then
                                EMsg &= vbCr & "Invalid Ship-To"
                            Else
                                If rowSOTORDR5.Item("CUST_ADDR3") & "" <> "" Then
                                    EMsg &= vbCr & "TSI cannot handle Address Line 3"
                                End If
                            End If
                        End If
                    End If


                    Dim SHIP_VIA_CODE As String = Absx1.txtFor("SHIP_VIA_CODE").Text
                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        If optORDR_SOURCE.Value & "" = "E" And SHIP_VIA_CODE = "" Then
                            SHIP_VIA_CODE = "ROUT"
                        End If
                    End If
                    If SHIP_VIA_CODE = "" Then
                        EMsg &= vbCr & "Ship Via is required"
                    Else
                        Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                        If rowSOTSVIA1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Ship Via Code"
                        Else
                            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                                If rowSOTSVIA1.Item("SHIP_VIA_STATUS") & "" <> "A" Then
                                    EMsg &= vbCr & "Ship Via Code specified is Inactive"
                                End If
                                If rowICTWHSE1.Item("WHSE_EDI_ID") & "" <> "" And SHIP_VIA_CODE <> "ROUT" Then
                                    ASCMAIN1.sql = "Select * from EDTXREF3" _
                                            & " where SENDER_ID_QUAL = :PARM1 and SENDER_ID = :PARM2 and SHIP_VIA_CODE = :PARM3"
                                    Dim rowEDTXREF3 As DataRow = ASCDATA1.GetDataRow _
                                                            (ASCMAIN1.sql, "VVV", New String() _
                                                             {rowICTWHSE1.Item("WHSE_EDI_QUAL"),
                                                              rowICTWHSE1.Item("WHSE_EDI_ID"),
                                                              SHIP_VIA_CODE})
                                    If rowEDTXREF3 Is Nothing Then
                                        EMsg &= vbCr & "Ship Via Code does not translate to a Valid Service Level for Whse " & WHSE_CODE
                                    End If

                                    If FRT_TERMS = "COL" Or FRT_TERMS = "3PY" Then
                                        Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & ""
                                        If CARRIER_CODE = "UPS" AndAlso rowARTCUST1.Item("UPS_ACCT_NO") & "" = "" Then EMsg &= vbCr & "No UPS Account set up for Customer"
                                        If CARRIER_CODE = "FEDEX" AndAlso rowARTCUST1.Item("FDX_ACCT_NO") & "" = "" Then EMsg &= vbCr & "No Fedex Account set up for Customer"
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                        EMsg &= vbCr & "Customer PO is required"
                    End If

                    If grdSOTORDR2.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Items on Order"
                    Else
                        If Val(dst.Tables("SOTORDR2").Compute("COUNT(ORDR_LNO)", "ORDR_QTY > 0") & "") = 0 Then
                            EMsg &= vbCr & "No Items on Order with Qty >0"
                        End If

                        For Each row As DataRow In ASCDATA1.SelectDistinct _
                                (dst.Tables("SOTORDR2").Select("RANGE_STYLE_CODE is Not Null"), "RANGE_STYLE_CODE").Rows
                            Dim RANGE_STYLE_CODE As String = row.Item("RANGE_STYLE_CODE")
                            If Val(dst.Tables("SOTORDR2").Select("RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'").Length) > 1 Then
                                EMsg &= vbCr & "Range Style " & RANGE_STYLE_CODE & " occurs on this Order More than Once"
                            End If
                        Next

                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select _
                                ("ISNULL(RANGE_STYLE_CODE,'') <> '' and RANGE_STYLE_QTY_PER_PP = 1")
                            Dim RANGE_STYLE_CODE As String = rowSOTORDR2.Item("RANGE_STYLE_CODE") & ""
                            Dim ORDR_LNO As Int32 = Val(rowSOTORDR2.Item("ORDR_LNO"))
                            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
                            Dim LINES As Int64 = Val(dst.Tables("SOTORDRR").Compute("COUNT(ORDR_LNO)", sqlw) & "")
                            '   Dim TOTAL As Decimal = Val(dst.Tables("SOTORDRR").Compute("SUM(ORDR_AMT)", sqlw) & "")
                            Dim ORDR_QTY As Int64 = Val(dst.Tables("SOTORDRR").Compute("SUM(ORDR_QTY)", sqlw) & "")
                            If Val(rowSOTORDR2.Item("ORDR_QTY") & "") <> ORDR_QTY _
                                Or LINES = 0 Then
                                EMsg &= vbCr & "Line " & CStr(ORDR_LNO) & ": Range Style Qty Out of Balance w/Components"
                            End If
                        Next

                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                            Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")

                            Dim CUST_UPC As String = rowSOTORDR2.Item("CUST_UPC") & ""
                            If CUST_UPC <> "" Then
                                If (Len(CUST_UPC) <> 11 And Len(CUST_UPC) <> 12 And Len(CUST_UPC) <> 13) Or Format(Val(CUST_UPC), "".PadLeft(Len(CUST_UPC), "0")) <> CUST_UPC Then
                                    EMsg &= vbCr & "Invalid Customer UPC/EAN on Line " & CStr(ORDR_LNO) & " (" & CUST_UPC _
                                            & "); must be 11 or 12 or 13 numeric digits"
                                End If
                            End If

                            Dim rowSOTORDR3s() As DataRow = rowSOTORDR2.GetChildRows("SOTORDR2_SOTORDR3")
                            If rowSOTORDR3s.Length > 0 Then
                                Dim ORDR_QTY_2 As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                                Dim unique_QTYs_by_Store As New Dictionary(Of Int64, String)
                                For Each row As DataRow In dst.Tables("SOTORDRS").Select("")
                                    Dim QTY As Int64 = Val(row.Item("QTY_" & Format(ORDR_LNO, "000")) & "")
                                    If Not unique_QTYs_by_Store.ContainsKey(QTY) Then
                                        Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                                        unique_QTYs_by_Store.Add(QTY, CUST_STORE_NO)
                                    End If
                                Next

                                Dim ORDR_QTY_3 As Int64 = 0
                                For Each rowSOTORDR3 As DataRow In rowSOTORDR3s
                                    Dim ORDR_QTY As Int64 = Val(rowSOTORDR3.Item("ORDR_QTY") & "")
                                    ORDR_QTY_3 += ORDR_QTY

                                    For Each QTY As Int64 In unique_QTYs_by_Store.Keys
                                        If QTY Mod ORDR_QTY <> 0 Then
                                            EMsg &= vbCr & "Rounding Problem with Sub-Details on Line " _
                                                    & CStr(ORDR_LNO) & " with Store " & unique_QTYs_by_Store(QTY)
                                            Exit For
                                        End If
                                    Next
                                    If rowSOTORDR2.Item("SIZE_DESC_01") & "" <> "" Then
                                        Dim TOTAL_SIZE_QTY As Int64 = Val(rowSOTORDR3.Item("TOTAL_SIZE_QTY") & "")
                                        If TOTAL_SIZE_QTY <> ORDR_QTY Then
                                            EMsg = EMsg & vbCr & "Size Distribution out of Balance with a Component on Line " & CStr(ORDR_LNO)
                                        End If
                                    End If
                                Next
                                If ORDR_QTY_2 <> ORDR_QTY_3 Then
                                    EMsg &= vbCr & "Sub-Details out of Balance with Total Amount for Style on Line " & CStr(ORDR_LNO)
                                End If
                            End If
                        Next

                        Dim STYLE_CODEs As String = ""
                        For Each TABLE_NAME As String In New String() {"SOTORDR2", "SOTORDRR"}
                            For Each row As DataRow In ASCDATA1.SelectDistinct _
                                        (dst.Tables("SOTORDR2").Select("STYLE_CODE is Not Null"), "STYLE_CODE").Rows
                                Dim STYLE_CODE As String = row.Item("STYLE_CODE") & ""
                                STYLE_CODEs &= ",'" & STYLE_CODE & "'"
                            Next
                        Next

                        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                            ' EACH STYLE IN THIS TABLE MUST HAVE AN INNER PACK
                            '  OR ELSE BAD THINGS HAPPEN IN SEND PT TO 3PL
                            '  IN THESE LINES OF CODE:
                            'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            'rowEDT940O2.Item("PACK_SIZE") = rowICTSTYL1.Item("INNER_PACK_QTY")
                            'Dim PACK_SIZE As Int64 = Val(rowEDT940O2.Item("PACK_SIZE"))
                            'Dim PICK_QTY As Int64 = Val(rowEDT940O2.Item("PICK_QTY"))

                            Dim STLE_CODEs_on_order As New List(Of String)
                            For Each row As DataRow In dst.Tables("SOTORDRR").Select("")
                                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                                If STLE_CODEs_on_order.Contains(STYLE_CODE) Then
                                    EMsg &= vbCr & "Duplicate Style in Range Components: " & STYLE_CODE
                                Else
                                    STLE_CODEs_on_order.Add(STYLE_CODE)
                                End If
                                Dim row2 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                                If Val(row2.Item("INNER_PACK_QTY") & "") = 0 Then
                                    EMsg &= vbCr & "Inner Pack Qty missing for Range Style Component " & STYLE_CODE
                                End If
                            Next
                        End If


                        'If STYLE_CODEs <> "" Then
                        '    ASCMAIN1.sql = "Select Distinct SALES_DIVISION_CODE from ICTSTYL1 where STYLE_CODE IN (" & Mid$(STYLE_CODEs, 2) & ")"
                        '    Dim rows() As DataRow = ASCDATA1.GetDataTable.Select
                        '    If rows.Length > 1 Then
                        '        'If MsgBox("Order Contains Styles From Different Divisions" _
                        '        '          & vbCrLf & "Are You Sure You Want To Continue?", MsgBoxStyle.YesNo, "Mixed Styles") = MsgBoxResult.No Then
                        '        '    Exit Sub
                        '        'End If
                        '    Else
                        '        'If rows(0).Item("SALES_DIVISION_CODE") & "" <> Absx1.txtFor("SALES_DIVISION_CODE").Text And Absx1.txtFor("SALES_DIVISION_CODE").Text <> "" Then
                        '        '    If MsgBox("Order Contains Styles From a Sales Divison Other Than " & Absx1.txtFor("SALES_DIVISION_CODE").Text _
                        '        '              & vbCrLf & "Are You Sure You Want To Continue?", MsgBoxStyle.YesNo, "Mixed Styles") = MsgBoxResult.No Then
                        '        '        Exit Sub
                        '        '    End If
                        '        'End If
                        '    End If
                        'End If
                    End If

                    If Absx1.chkFor("ORDR_SHIP_COMPLETE").Checked And chkReleaseNow.Checked Then
                        EMsg &= "You have opted to Ship Short on Next Release," _
                                      & vbCrLf & " but the Order won't ship short" _
                                      & vbCrLf & " because it is indicated to Ship Complete"
                    End If

                    If EMsg = "" Then
                        '  Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                        '    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

                        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                            Dim ORDR_INV_COMMENT As String = Absx1.txtFor("ORDR_INV_COMMENT").Text
                            If Split(ORDR_INV_COMMENT, vbCrLf).Length > 4 Then
                                If MsgBox("Your Invoice Comment will not print correctly on the printed invoice." _
                                               & vbCrLf & vbCrLf & "Continue with Update anyway?",
                                              MsgBoxStyle.YesNo, "Warning - Print area is only 4 lines") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            End If
                        End If


                        If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB" Then
                            If rowICTWHSE1.Item("WHSE_TYPE") & "" <> "P" Then

                                If (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") And (WHSE_CODE = "NC" Or WHSE_CODE = "NY" Or WHSE_CODE = "ZZ" Or WHSE_CODE = "SP") Then
                                    'If (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") And (WHSE_CODE = "NY" Or WHSE_CODE = "ZZ" Or WHSE_CODE = "SP") Then
                                    ' NY IS OK FOR A BTB ORDER - PROBABLY NEED AN ATTRIBUTE IN ICTWHSE1 - WAITING ON WHR FOR DDL OK
                                Else
                                    EMsg &= vbCr & "Invalid Warehouse Code (" & WHSE_CODE & ") for a Back-to-Back Order"
                                    End If
                                End If

                                If dst.Tables("SOTORDP1").Select("INV_DATE IS NULL").Length <> 0 Then
                                EMsg &= vbCr & "Invalid Invoice Date for a Pro-Forma Invoice on a Back-to-Back Order"
                            End If

                            If multi_store_is_active Then
                                EMsg &= vbCr & "Cannot Do Multi-Store Back-to-Back Orders"
                            End If

                            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                                If STYLE_CODE = "" Then
                                    EMsg &= vbCr & "No Range Styles on Back-to-Back Orders (See Line " & rowSOTORDR2.Item("ORDR_LNO") & ")"
                                    Exit For
                                Else
                                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                                    Dim VEND_CODE As String = rowICTSTYL1.Item("VEND_CODE") & ""
                                    If VEND_CODE = "" Then
                                        EMsg &= vbCr & "All Styles on a Back-to-Back Order must have a Primary Supplier (See Style " & STYLE_CODE & ")"
                                        Exit For
                                    Else
                                        If Val(rowSOTORDR2.Item("PO_COST") & "") <= 0 Then
                                            EMsg &= vbCr & "All Styles on a Back-to-Back Order must have a PO Cost defined (See Style " & STYLE_CODE & ")"
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            If rowICTWHSE1.Item("WHSE_TYPE") & "" = "P" Then
                                EMsg &= vbCr & "Invalid Warehouse Code (" & WHSE_CODE & ") for this type of Orders"
                            End If
                        End If
                    End If
                End If

                If ASCMAIN1.CLIENT = "NYA" Then
                    If EMsg = "" Then
                        Dim SEG4_CODE As String = TAC.TACMAIN1.Check_Division_MixMatch(Me, EMsg, "SOTORDR2", Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("WHSE_CODE").Text)
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                              "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTPICK1 where PICK_STATUS <> 'D' "
                If EntryMode = "M" Or multiple_order_maintenance Then
                    ASCMAIN1.sql &= " and ORDR_NO in (Select ORDR_NO from SOTORDR1 " _
                            & IIf(multiple_order_type = "ORDR_GROUP_NO",
                                  "where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'",
                                  "where EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") _
                            & ")"
                Else
                    ASCMAIN1.sql &= " and ORDR_NO = '" & ORDR_NO & "'"
                End If

                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Order has been Released and/or Partially Shipped"
                Else
                    If EntryMode = "M" Or multiple_order_maintenance Then
                        If Absx1.txtFor("SOTORDR0.REASON_CODE").Text = "" Then
                            EMsg &= vbCr & "You must Specify a Reason Code when Deleting a Group of Orders"
                        Else
                            If LookUp("SOTREAS1", Absx1.txtFor("SOTORDR0.REASON_CODE").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                            End If
                        End If
                        If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'P'").Length <> 0 Then
                            EMsg &= vbCr & "No Orders May be in Pick when Deleting an Orders in a Group"
                        End If
                        If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'O'").Length = 0 Then
                            EMsg &= vbCr & "No Open Orders to Delete"
                        End If
                    Else
                        If Absx1.txtFor("REASON_CODE").Text = "" Then
                            EMsg &= vbCr & "You must Specify a Reason Code when Deleting an Order"
                        Else
                            If LookUp("SOTREAS1", Absx1.txtFor("REASON_CODE").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                            End If
                        End If
                    End If

                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Order as Deleted",
                                      MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If


            Case "Start Multi-Store"
                If grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.DataChanged Then
                    grdSOTORDR2.ActiveRow.Update()
                End If

                If rowSOTORDR1.Item("CCPA_NO") & "" <> "" Then
                    EMsg &= vbCr & "Cannot do a Multi-Store Order which is associated with a Credit Card"
                End If

                If multi_store_is_active Then
                    If MsgBox("Are You Sure that you want to Clear All Multi-Store Entries",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Setup_MS(False)
                    End If
                    Exit Sub
                End If

                If grdSOTORDR2.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must First Add Line Items to This Order"
                Else
                    If dst.Tables("SOTORDR2").Select("ISNULL(RANGE_STYLE_CODE,'') <> ''").Length <> 0 Then
                        EMsg &= vbCr & "Cannot Have Range Styles on a Multi-Store Order"
                    End If
                End If

            Case "Cancel Order"

                If EntryMode = "M" Or multiple_order_maintenance Then
                    If Absx1.txtFor("SOTORDR0.REASON_CODE").Text = "" Then
                        EMsg &= vbCr & "You must Specify a Reason Code when Cancelling an Order"
                    Else
                        If LookUp("SOTREAS1", Absx1.txtFor("SOTORDR0.REASON_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                        End If
                    End If
                    If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'P'").Length <> 0 Then
                        EMsg &= vbCr & "No Orders May be in Pick when Cancelling an Orders in a Group"
                    End If
                    If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'O'").Length = 0 Then
                        EMsg &= vbCr & "No Open Orders to Cancel"
                    End If
                Else
                    If Absx1.txtFor("REASON_CODE").Text = "" Then
                        EMsg &= vbCr & "You must specify a Reason Code when Cancelling an Order"
                    Else
                        If LookUp("SOTREAS1", Absx1.txtFor("REASON_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                        End If
                    End If
                End If
                If EMsg = "" Then
                    If MsgBox("Do you want to Cancel (the remaining open balance on) this Order" _
                                   & vbCrLf & "(Lost Sales will be charged)",
                                   vbYesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Credit Card"
                If Absx1.txtFor("ORDR_NO").Text = "" Then
                    EMsg &= vbCr & "No Order No Specified"
                Else
                    ASCMAIN1.sql = "Select * from ARTCCPA1" _
                        & " where ORDR_NO = '" & HFs("ORDR_NO") & "'" _
                        & " and CCPA_STATUS IN ('C','T')"
                    Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow
                    If rowARTCCPA1 IsNot Nothing Then
                        Dim dispMessage = True
                        ' Allow RGI to do a second, third, ... Authorization since they make multi shipments and only authorize what they need to ship
                        ' the avaiable product
                        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                            Dim row As DataRow = ASCDATA1.GetDataRow("select * from artccpa1 where CCPA_NO_AUTH = '" & rowARTCCPA1.Item("CCPA_NO") & "'")
                            If row IsNot Nothing Then
                                dispMessage = False
                            End If
                        End If

                        If dispMessage Then
                            EMsg &= "CC Authorization for " & Format(Val(rowARTCCPA1.Item("CCPA_AMT") & ""), "$#.00") & " already recorded for this order"
                            Exit Select
                        End If
                    End If

                    If multi_store_is_active Then
                        EMsg &= vbCr & "Cannot do a Multi-Store Order which is associated with a Credit Card"
                    End If
                End If

            Case "Re-Queue for Credit"
                Dim CUST_FACTOR_IND As String = rowSOTORDR1.Item("CUST_FACTOR_IND") & ""
                Dim ORDR_HOLD As String = rowSOTORDR1.Item("ORDR_HOLD") & ""
                Dim rowSOTORDR0 = LookUp("SOTORDR0", rowSOTORDR1.Item("ORDR_GROUP_NO"))
                Dim ORDR_AMT_NOW As Decimal = Val(rowSOTORDR0.Item("ORDR_AMT") & "")
                ' CUST_FACTOR_IND = "1" SHOULD NOT BE ALLOWED IN PREREQ IF ORDR_AMT_NOW = 0 OR TERM_TYPE = "C"
                'If ORDR_AMT_NOW <> 0 And CUST_FACTOR_IND = "1" And ORDR_HOLD <> "1" And TERM_TYPE <> "C" Then
                If CUST_FACTOR_IND = "1" And ORDR_HOLD <> "1" And ORDR_AMT_NOW > 0 Then
                Else
                    EMsg &= vbCr & "Order must be Factored, not on hold, and >$0 in order to generate a Credit Approval Request"
                End If

                If EMsg = "" Then
                    If blnAutomatic Then
                        ' no prompt - automated queue
                    Else
                        If MsgBox("Do you want to submit a request to the Factor for Credit Approval on this order?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
                    End If
                End If

            Case "Convert to Regular"

                If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, , , , 2) Then
                    Exit Sub
                Else
                    If MsgBox("This option will convert" & vbCrLf & " Sales Order " & ORDR_NO & vbCrLf & " from a BTB Order to a Regular Order." _
                                  & vbCrLf & vbCrLf & "OK to Proceed?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        ASCMAIN1.MultiTask_Release(, , 2)
                        Exit Sub
                    End If
                End If


            Case "Reinstate Cancelled"


                If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "B2C" Then
                    MsgBox("Cannot re-instate cancelled eCommerce orders", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, , , , 2) Then
                    Exit Sub
                Else
                    If MsgBox("This option will re-open" & vbCrLf & " Sales Order " & ORDR_NO & " using the Qtys Cancelled as Open Order Qtys." _
                                  & vbCrLf & vbCrLf & "OK to Proceed?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        ASCMAIN1.MultiTask_Release(, , 2)
                        Exit Sub
                    End If
                End If

            Case "Cancel && Clone"

                ASCMAIN1.sql = "Select Count (*) from POTORDR2 where ORDR_NO = '" & ORDR_NO & "' and PO_QTY_OPN <> 0"
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Cancel & Clone option is valid only for BTB Orders" & vbCr & " whose associated POs have nothing open"
                Else
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, , , , 2) Then
                        Exit Sub
                    Else
                        Dim row As DataRow = LookUp("SOTORDR1", ORDR_NO)
                        If row.Item("ORDR_STATUS") <> "O" Then
                            EMsg &= "Order Is No Longer Open"
                            ASCMAIN1.MultiTask_Release(, , 2)
                        Else
                            If MsgBox("This option will cancel the remaining balance open" & vbCrLf & " on BTB Sales Order " & ORDR_NO & vbCrLf & " and clone it to a new Regular Sales Order with those Qtys as Open Order Qtys." _
                                          & vbCrLf & vbCrLf & "OK to Proceed?",
                                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                ASCMAIN1.MultiTask_Release(, , 2)
                                Exit Sub
                            End If
                        End If
                    End If
                End If


            Case "Cancel-Keep PO"
                ' PO cannot be shipped
                ' SO cannot be released or shipped - must be entirely open

                ASCMAIN1.sql = "Select Count (*) from POTORDR2 where ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                        & " and (NVL(PO_QTY_OPN,0) = 0 or NVL(PO_QTY_OPN,0) <> NVL(PO_QTY_ORD,0) or NVL(PO_QTY_SHP,0) <> 0 or NVL(PO_QTY_SHP,0) <> 0)"
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Cancel / Keep PO option is valid only for BTB Orders" _
                            & vbCr & " whose associated POs have Qty open" _
                            & vbCr & "   and have NOT been Shipped nor Received"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, , , , 2) Then
                        Exit Sub
                    Else
                        Dim row As DataRow = LookUp("SOTORDR1", ORDR_NO)
                        If row.Item("ORDR_STATUS") <> "O" Then
                            EMsg &= "Order Is No Longer Open"
                            ASCMAIN1.MultiTask_Release(, , 2)
                        Else
                            ASCMAIN1.sql = "Select Count (*) from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'" _
                                    & " and (NVL(ORDR_QTY_PICK,0) <> 0 or NVL(ORDR_QTY_SHIP,0) <> 0)"
                            If Val(ASCDATA1.GetDataValue) <> 0 Then
                                EMsg &= vbCr & "Cancel / Keep PO option is valid only for BTB Orders" _
                                        & vbCr & " which are not Released and not already Billed" _
                                        & vbCr & " (ie, they must be 100% open)"
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    Dim WHSE_CODE_keep As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                    If MsgBox("This option will cancel the remaining balance open" _
                                  & vbCrLf & " on BTB Sales Order " & ORDR_NO &
                                  vbCrLf & " and keep the Original POs open with Destination " & WHSE_CODE_keep & "." _
                                  & vbCrLf & vbCrLf & "OK to Proceed?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        ASCMAIN1.MultiTask_Release(, , 2)
                        Exit Sub
                    End If
                End If

            Case "Confirm (855)"
                If ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, , , , 2) Then
                    Exit Sub
                Else
                    Dim row As DataRow = LookUp("SOTORDR1", ORDR_NO)
                    If row.Item("ORDR_STATUS") <> "O" And row.Item("ORDR_STATUS") <> "P" Then
                        EMsg &= "Order Is No Longer Open"
                        ASCMAIN1.MultiTask_Release(, , 2)
                    Else
                        If MsgBox("This option will confirm Order Group " & ORDR_GROUP_NO & vbCrLf & " by sending an EDI Document 855." _
                                      & vbCrLf & vbCrLf & "OK to Proceed?",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            ASCMAIN1.MultiTask_Release(, , 2)
                            Exit Sub
                        End If
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

                If ORDR_NO_to_copy <> "" Then
                    ' need to use grids
                    ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO_to_copy & "'"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "ORDR_LNO")
                        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = row.Item("COLOR_CODE")

                        Add_grdSOTORDR2(STYLE_CODE, COLOR_CODE, Val(row.Item("ORDR_QTY") & ""), Val(row.Item("ORDR_UNIT_PRICE") & ""))

                        'For Each TABLE_NAME As String In New String() _
                        '        {"SOTORDR2", "SOTORDR3", "SOTORDR4", "SOTORDR9"}
                        '    '  "SOTORDR5",
                        '    Fill_Records(TABLE_NAME, ORDR_NO_to_copy)
                        '    For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
                        '        row.Item("ORDR_NO") = ORDR_NO
                        '        If TABLE_NAME = "SOTORDR2" Then
                        '            row.Item("ORDR_QTY_ALLO") = 0
                        '            row.Item("ORDR_QTY_OPEN") = row.Item("ORDR_QTY")
                        '            row.Item("ORDR_QTY_PICK") = 0
                        '            row.Item("ORDR_QTY_SHIP") = 0
                        '            row.Item("ORDR_QTY_CANC") = 0
                        '            row.Item("ORDR_STATUS") = "O"
                        '            row.Item("RSRV_NO") = DBNull.Value
                        '            row.Item("RSRV_LNO") = DBNull.Value
                        '            row.Item("ORDR_QTY_PRE_ALLO") = 0
                        '        End If
                        '    Next
                    Next

                    If grdSOTORDR2.ActiveRow IsNot Nothing Then
                        grdSOTORDR2.ActiveRow.CancelUpdate()
                    End If
                    Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

                End If


            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                If EntryMode = "M" Or multiple_order_maintenance Then
                    Update_Record_multiple_order_group()
                Else
                    Update_Record()
                End If

                If chkAllocate.Checked Then
                    Dim ORDR_NO_now As String = ORDR_NO
                    Mode_Settings(False)
                    Allocate_Order(ORDR_NO_now)
                    Absx1.txtFor("ORDR_NO").Text = ORDR_NO_now
                    Click_Command("View")
                    tabMain.SelectedTab = tabMain.Tabs("Line Items")
                Else
                    Mode_Settings(False)
                End If

            Case "Delete"
                'Delete_Record()
                Delete_Order()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Pro-Forma"

                Using F As New ASFMSGBF

                    If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
                        Print_ProForma("ORDR_QTY")
                    Else
                        Dim ORDR_QTY_fields() As String = New String() {"Qty Ordered", "Qty Open", "Qty Allocated", "Qty In Pick", "Qty Available"}
                        'Dim ORDR_QTY_fields() As String = New String() {"Qty Ordered", "Qty Open", "Qty Allocated", "Qty Allocated Current", "Qty In Pick"}
                        Dim i As Integer = F.Get_opt_from_User("Which Qty Field should be Used", ORDR_QTY_fields, 0, "Pro-Forma Qty Option")
                        If i <> -1 Then
                            Dim ORDR_QTY_field As String = New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_PICK", "ORDR_QTY_ALLO_X"}(i)
                            'Dim ORDR_QTY_field As String = New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_ALLO_CURR", "ORDR_QTY_PICK"}(i)
                            ' BLOWS UP BECAUSE ORDR_QTY_ALLO_CUR IS NOT IN ORACLE SOTORDR2
                            Print_ProForma(ORDR_QTY_field)
                        End If
                    End If

                End Using

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Cancel Order"
                Cancel_Order()
                Mode_Settings(False)

            Case "Start Multi-Store"
                ReSelect_Stores()

            Case "Reset Qty's"
                If MsgBox("Are You Sure that you want to Reset the Qty's to Store " & Absx1.txtFor("CUST_STORE_NO").Text & "'s Values",
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Init_MultiStore()
                End If

            Case "Re-Select Stores"
                ReSelect_Stores()

            Case "Clear Zeroes"
                Clear_Zeroes()

            Case "Work Orders"
                Using F As New TAC.SOFWORK1(Me, "S", ORDR_NO, (EntryMode = "V" Or InquiryMode),
                                            Absx1.txtFor("CUST_CODE").Text,
                                            Absx1.txtFor("ORDR_CUST_PO").Text,
                                            Absx1.dteFor("ORDR_SHIP_DATE").Value,
                                             Absx1.dteFor("ORDR_CANCEL_DATE").Value,
                                            "Work Orders relating to Sales Order " & ORDR_NO)
                    F.ShowDialog()
                End Using


            Case "Credit Card"
                Credit_Card()
                Update_Record_TDA("SOTORDC1")
                Update_Record_TDA("SOTORDC2")

            Case "Re-Queue for Credit"
                BeginTrans()
                Credit_Request()
                Dim msg As String = "Request for Credit Approval has been Re-Queued"
                If blnAutomatic Then
                    msg = ""
                End If
                CommitTrans(msg)

            Case "Convert to Regular"
                Dim ORDR_NO_to_convert As String = ORDR_NO
                BeginTrans()
                Dim ORDR_FOB As String = Get_ORDR_FOB("REG", rowSOTORDR1.Item("WHSE_CODE"))
                ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_TYPE_CODE = 'REG', ORDR_FOB = :PARM1 where ORDR_NO = '" & ORDR_NO_to_convert & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {ORDR_FOB})
                ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_PRICE_SOURCE = NULL where ORDR_NO = '" & ORDR_NO_to_convert & "'"
                ASCDATA1.ExecuteSQL()
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    ' IF WE EVER DO MULTIPLE ORDERS IN A GROUP - WE WILL NEED TO CALL THIS FOR EACH ORDER
                    ASCDATA1.ExecuteSP("SOPORDR1_COMM", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
                End If
                CommitTrans()
                Click_Command("Done")
                Absx1.txtFor("ORDR_NO").Text = ORDR_NO_to_convert
                Click_Command("View")

            Case "Reinstate Cancelled"
                Dim ORDR_NO_to_reinstate As String = ORDR_NO
                Reverse_Cancel(ORDR_NO, "")
                Click_Command("Done")
                Absx1.txtFor("ORDR_NO").Text = ORDR_NO_to_reinstate
                Click_Command("View")

            Case "Cancel && Clone"
                Dim ORDR_NO_to_clone As String = ORDR_NO
                Absx1.txtFor("REASON_CODE").Text = "CLONE"

                Dim TBL2 As DataTable = dst.Tables("SOTORDR2").Copy

                Cancel_Order()
                Click_Command("Done")
                Absx1.txtFor("ORDR_NO").Text = ORDR_NO_to_clone
                Click_Command("View")
                Dim ORDR_NO_new As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                ORDR_NO = ORDR_NO_new
                rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
                rowSOTORDR1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTORDR1.Item("INIT_DATE") = DATETIME_STAMP
                rowSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTORDR1.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTORDR1.Item("ORDR_STATUS") = "O"
                rowSOTORDR1.Item("WHSE_CODE") = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    rowSOTORDR1.Item("WHSE_CODE") = "NY"
                End If
                rowSOTORDR1.Item("REASON_CODE") = DBNull.Value
                rowSOTORDR1.Item("ORDR_GROUP_NO") = DBNull.Value
                rowSOTORDR1.Item("ORDR_DATE_CLOSED") = DBNull.Value
                rowSOTORDR1.Item("ORDR_YYYYPP_CLOSED") = DBNull.Value
                rowSOTORDR1.AcceptChanges()
                rowSOTORDR1.SetAdded()

                EntryMode = "N"

                dst.Tables("SOTORDP2").Rows.Clear()
                dst.Tables("SOTORDP1").Rows.Clear()

                For Each TABLE_NAME As String In New String() {"SOTORDR2", "SOTORDR5"}
                    For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                        row.Item("ORDR_NO") = ORDR_NO
                        If TABLE_NAME = "SOTORDR2" Then
                            Dim row2 As DataRow = TBL2.Rows.Find(New Object() {ORDR_NO_to_clone, row.Item("ORDR_LNO")})
                            row.Item("ORDR_QTY") = row2.Item("ORDR_QTY_OPEN")
                            row.Item("ORDR_QTY_OPEN") = row2.Item("ORDR_QTY_OPEN")
                            row.Item("ORDR_QTY_ALLO") = 0
                            row.Item("ORDR_QTY_PICK") = 0
                            row.Item("ORDR_QTY_SHIP") = 0
                            row.Item("ORDR_QTY_CANC") = 0
                            row.Item("ORDR_STATUS") = "O"
                            row.Item("ORDR_QTY_ORIG") = row2.Item("ORDR_QTY_OPEN")
                            row.Item("PO_COST") = DBNull.Value
                        End If
                        row.AcceptChanges()
                        row.SetAdded()
                    Next
                Next

                For Each TABLE_NAME As String In New String() {"SOTORDR3", "SOTORDR4", "SOTORDR9", "SOTORDXR",
                                                               "TATEVNT1", "SOTWORK1", "SOTWORK2", "SOTORDP1", "SOTORDP2"}
                    dst.Tables(TABLE_NAME).Rows.Clear()
                Next

                Click_Command("Update")
                If ScreenMode Then
                    MsgBox("An unexpected error occurred during the Cancel & Clone", MsgBoxStyle.OkOnly, "Please Call ABS")
                Else
                    MsgBox("BTB Sales Order " & ORDR_NO_to_clone & " has been successfully Cancelled and Cloned" & vbCrLf & " to Regular Sales Order " & ORDR_NO_new, MsgBoxStyle.OkOnly, "Verification")
                End If
                Absx1.txtFor("ORDR_NO").Text = ORDR_NO_new
                Click_Command("View")

            Case "Cancel-Keep PO"
                Absx1.txtFor("REASON_CODE").Text = "CXLKPO"

                PO_ORDER_NOs.Clear()
                Cancel_Order(True)

                MsgBox("BTB Sales Order " & ORDR_NO & " has been successfully Cancelled" _
                       & vbCrLf & " and its POs have been kept Open", MsgBoxStyle.OkOnly, "Verification")

                Click_Command("Done")

            Case "Confirm (855)"
                TAC.SOCMAIN1.Generate_855(Me.clsASCBASE1, ORDR_GROUP_NO)
                Dim msg As String = "" 'TAC.SOCMAIN1.Generate_855(Me.clsASCBASE1, ORDR_GROUP_NO)
                If msg <> "" Then
                    MsgBox("Error Occurred sending 855:" _
                           & vbCrLf & vbCrLf & msg _
                           & vbCrLf & vbCrLf & "Please take a screenshot and email to ABS", MsgBoxStyle.OkOnly, "Please Report this Error to ABS")
                Else
                    MsgBox("A Confirmation (855) has been queued up for Transmission to the Customer", MsgBoxStyle.OkOnly, "Verification")
                    UltraExplorerBar1.Groups("Screen Control").Items("Confirm (855)").Visible = False
                End If

            Case "Create POs"
                BeginTrans()
                Update_POs(False)
                CommitTrans()

                MsgBox("POs have been Created", MsgBoxStyle.OkOnly, "Verification")

                Set_Control_POs(False)
                Mode_Settings(False)

            Case "Cancel POs"
                Set_Control_POs(False)
                Mode_Settings(False)

            Case "Import XFR File"
                Import_XFR_File()
            Case "Import Amazon File"
                Import_Amazon_File()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        If Not ASCMAIN1.Running_in_VS Then
            'If ASCMAIN1.DBS_COMPANY = "VAN" Then
            '    MsgBox("This program is not ready for Production Use", MsgBoxStyle.OkOnly, "Please call ABS")
            '    Me.Close()
            'End If
        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            lblORDR_CATEGORY.Visible = True
            txtORDR_CATEGORY.Visible = True
        Else
            lblORDR_CATEGORY.Visible = False
            txtORDR_CATEGORY.Visible = False
        End If

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowSOTORDR1.Item("ORDR_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If

                    ' Credit Cards can be Authorized when in Open and Pick.
                    If ",O,P,".Contains(rowSOTORDR1.Item("ORDR_STATUS") & String.Empty) Then
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "E" Or EntryMode = "N") Then
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.False
                    End If
                End If

                .Items("Update").Settings.Enabled = iScreenMode
                If disable_update Then
                    .Items("Update").Settings.Enabled = DefaultableBoolean.False
                End If
                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode
                .Items("Pro-Forma").Settings.Enabled = iScreenMode

                .Items("Cancel Order").Settings.Enabled = iScreenMode

                '.Items("Pro-Rate").Visible = (EntryMode <> "V")

                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode
                .Items("Credit Card").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                .Items("Print").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                .Items("Pro-Forma").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode

                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Delete").Visible = (EntryMode = "E") And Not InquiryMode AndAlso (rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" <> "BTB")
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode

                .Items("Cancel Order").Visible = (EntryMode = "E")

                .Items("Work Orders").Text = "Work Orders" & IIf(dst.Tables("SOTWORK1").Rows.Count = 0, "", " (" & CStr(dst.Tables("SOTWORK1").Rows.Count) & ")")
                .Items("Work Orders").Visible = ScreenMode And (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")

                If (EntryMode = "N") Then
                    .Items("Credit Card").Visible = (MENU_ITEM_OBJECT = "SOFORDR1") And Not disable_update
                    .Items("Credit Card").Settings.Enabled = iScreenMode
                End If

                If disable_update Or multiple_order_maintenance Or InquiryMode Or Not ScreenMode Then
                    .Items("Credit Card").Visible = False
                Else
                    .Items("Credit Card").Visible = (Absx1.optFor("ORDR_TYPE_CODE").Value = "REG")
                End If

                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    .Items("Re-Queue for Credit").Visible = ScreenMode And EntryMode = "V" And (rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("ORDR_STATUS") & "" = "O")
                Else
                    .Items("Re-Queue for Credit").Visible = False
                End If

                .Items("Convert to Regular").Visible = False
                If Not InquiryMode And (EntryMode = "V") And (Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB") And dst.Tables("POTORDR1").Rows.Count = 0 Then
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1.Item("WHSE_TYPE") & "" = "W" Then
                        .Items("Convert to Regular").Visible = True
                    End If
                End If

                .Items("Reinstate Cancelled").Visible = False
                If Not InquiryMode And ScreenMode Then
                    Dim cancelled_line_items As Boolean = (dst.Tables("SOTORDR2").Select("ORDR_STATUS = 'C'").Length <> 0)
                    If (EntryMode = "V") And (Absx1.optFor("ORDR_TYPE_CODE").Value = "REG") _
                        AndAlso rowSOTORDR1.Item("ORDR_STATUS") & "" = "C" Or (rowSOTORDR1.Item("ORDR_STATUS") & "" = "F" And cancelled_line_items) Then
                        .Items("Reinstate Cancelled").Visible = True
                    End If
                End If

                .Items("Cancel && Clone").Visible = False
                .Items("Cancel-Keep PO").Visible = False
                If Not InquiryMode And (EntryMode = "V") And (Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB") AndAlso rowSOTORDR1.Item("ORDR_STATUS") & "" = "O" Then
                    .Items("Cancel && Clone").Visible = True
                    .Items("Cancel-Keep PO").Visible = True
                End If

                .Items("Confirm (855)").Visible = False
                If Not InquiryMode And (EntryMode = "V") And (Absx1.optFor("ORDR_TYPE_CODE").Value = "REG") Then
                    If rowSOTORDR1.Item("ORDR_SOURCE") & "" = "E" And (rowSOTORDR1.Item("ORDR_STATUS") & "" = "O" Or rowSOTORDR1.Item("ORDR_STATUS") & "" = "P") Then
                        ASCMAIN1.sql = "Select * from EDTTRPM1 where EDI_DOC_NO = '855' and CUST_CODE = '" & rowSOTORDR1.Item("CUST_CODE") & "'"
                        Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow
                        If rowEDTTRPM1 IsNot Nothing Then
                            ASCMAIN1.sql = "Select EDT855O1.* from EDT855O1,EDTSYSIH" _
                                & " where EDT855O1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" _
                                & "   and EDT855O1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                                & "   and EDTSYSIH.COMPANY_CODE = EDT855O1.COMPANY_CODE" _
                                & "   and EDTSYSIH.EDI_OUTBOUND_DOC_NO = EDT855O1.EDI_OUTBOUND_DOC_NO" _
                                & "   and TRIM(EDTSYSIH.EDI_TP_ID) = '" & Trim(rowEDTTRPM1.Item("EDI_TP_ID")) & "'"
                            Dim rowEDT855O1 As DataRow = ASCDATA1.GetDataRow
                            If rowEDT855O1 Is Nothing Then
                                .Items("Confirm (855)").Visible = True
                            End If
                        End If
                    End If
                End If


                If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                    .Items("Re-Queue for Credit").Visible = False
                    .Items("Work Orders").Visible = False
                End If
            End With

            .Groups("Create POs from SOs").Visible = False
            .Groups("Show Orders").Visible = Not ScreenMode And (tabSOTORDRX.SelectedTab.Key = "Open Orders")

            .Groups("Screen Control").Items("Set Range Style PPK = 1").Visible = New String() {"wayne", "angela"}.Contains(ASCMAIN1.USER_ID)

            .Groups("Multi-Store").Visible = False
            .Groups("Order Details").Visible = False
            .Groups("Totals").Visible = ScreenMode
            .Groups("Copy").Visible = False '  Not ScreenMode And Not InquiryMode
            .Groups("Release Holds").Visible = ScreenMode And (EntryMode = "V") AndAlso (rowSOTORDR1.Item("ORDR_STATUS") = "O")
            .Groups("Pro Forma Invoice").Visible = False
            .Groups("Special Functions").Visible = (ASCMAIN1.CLIENT = "VAN" And (ASCMAIN1.USER_ID = "dgj" Or ASCMAIN1.USER_ID = "wendy")) And Not ScreenMode And Not InquiryMode


            With .Groups("Multi-Store")
                .Visible = (EntryMode = "N") And dst.Tables("ARTCUST2").Rows.Count > 1
                If ScreenMode AndAlso rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "XFR" Then
                    .Visible = False
                End If
                .Items("Start Multi-Store").Settings.Enabled = DefaultableBoolean.True
                .Items("Reset Qty's").Settings.Enabled = DefaultableBoolean.False
                .Items("Re-Select Stores").Settings.Enabled = DefaultableBoolean.False
                .Items("Clear Zeroes").Settings.Enabled = DefaultableBoolean.False
            End With

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                .Groups("Multi-Store").Visible = False
            End If
        End With


        chkAllocate.Visible = (EntryMode = "N" Or EntryMode = "E") AndAlso Absx1.optFor("ORDR_TYPE_CODE").Value <> "BTB"
        chkReleaseNow.Visible = (EntryMode = "N" Or EntryMode = "E") AndAlso Absx1.optFor("ORDR_TYPE_CODE").Value <> "BTB"

        lblCURR_CODE.Visible = ScreenMode

        lblINV_NO.Visible = InquiryMode And Not ScreenMode
        txtINV_NO.Visible = InquiryMode And Not ScreenMode
        lblPICK_NO.Visible = InquiryMode And Not ScreenMode
        txtPICK_NO.Visible = InquiryMode And Not ScreenMode
        If InquiryMode Then
            Absx1.optFor("ORDR_TYPE_CODE").Visible = ScreenMode
        End If

        lblMultiStore.Visible = False
        lblStatus.Visible = ScreenMode

        tabDetails.Visible = ScreenMode And (EntryMode <> "N")

        Set_Read_Only(grpSHIPTO, True)
        Set_Read_Only(splComments.Panel1, True)

        If ScreenMode And Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
            'Absx1.txtFor("ORDR_PRIORITY").ReadOnly = Not ASCMAIN1.USER_SECURITY_CODEs.Contains("X2")
            Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_PRIORITY"), Not ASCMAIN1.USER_SECURITY_CODEs.Contains("X2"))
        End If

        tabMain.Tabs("Multi-Store").Visible = False
        tabMain.Tabs("Shipments").Visible = (EntryMode = "V") And dst.Tables("SOTPICK1").Rows.Count > 0
        tabMain.Tabs("Back-to-Back").Visible = (EntryMode = "V" Or EntryMode = "E") And dst.Tables("POTORDR1").Rows.Count > 0

        tabSOTORDRX.Visible = Not tf
        splPOs.Visible = False

        tabMain.Visible = tf And Not ((EntryMode = "E") And multiple_order_maintenance)
        frmSOTORDRD.Visible = tf And ((EntryMode = "E") And multiple_order_maintenance)

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ASCMAIN1.CLIENT = "NYA" AndAlso (EntryMode = "E") AndAlso dst.Tables("POTORDR1").Rows.Count > 0 Then
            Set_Read_Only_for_ctl(Absx1.txtFor("CUST_STORE_NO"), False)
        End If

        With grdSOTORDP1.DisplayLayout.Bands(0)
            For Each C As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}
                .Columns(C).Hidden = ScreenMode
            Next
        End With
        splSOTORDPX.Panel2Collapsed = Not ScreenMode
        splProForma.Panel2Collapsed = Not ScreenMode
        If ScreenMode Then
            splProForma.Parent = splBTB.Panel2
            Show_Filter(grdSOTORDP1, False)
        Else
            splProForma.Parent = tabSOTORDRX.Tabs("Pro-Forma Invoices").TabPage
            Show_Filter(grdSOTORDP1, True)
            tabSOTORDRX.SelectedTab = tabSOTORDRX.Tabs("Open Orders")
        End If

        chkShortView.Visible = (ASCMAIN1.CLIENT = "RGI") And EntryMode = "V"

        If ScreenMode Then

            grdSOTORDR2.DisplayLayout.Bands(0).Columns("PO_COST").Hidden = Not (rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB")
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("MU_PCT").Hidden = Not (rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB") Or Not (rowSOTORDR1.Item("WHSE_CODE") & "" = "FE")

            With grdSOTORDR2.DisplayLayout.Bands(0)
                .Columns("ORDR_UNIT_COST").Hidden = Not (EntryMode = "V") Or Not (rowSOTORDR1.Item("ORDR_STATUS") = "F")
                .Columns("CGS").Hidden = Not (EntryMode = "V") Or Not (rowSOTORDR1.Item("ORDR_STATUS") = "F")
                .Columns("GP_AMT").Hidden = Not (EntryMode = "V") Or Not (rowSOTORDR1.Item("ORDR_STATUS") = "F")
                .Columns("GP_PCT").Hidden = Not (EntryMode = "V") Or Not (rowSOTORDR1.Item("ORDR_STATUS") = "F")
            End With

            If EntryMode = "V" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdSOTORDR2, grdSOTORDR3, grdSOTORDR4, grdSOTORDRR, grdSOTORDRS, grdSOTORDP1, grdSOTORDP2}
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                Next
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdSOTORDR2, grdSOTORDR3, grdSOTORDR4, grdSOTORDRR, grdSOTORDRS, grdSOTORDP1, grdSOTORDP2}
                    If grd.Name = "grdSOTORDR3" Or grd.Name = "grdSOTORDRS" Or grd.Name = "grdSOTORDP1" Or grd.Name = "grdSOTORDP2" Then
                        grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    Else
                        grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    End If
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                Next

                If Absx1.optFor("ORDR_TYPE_CODE").Value = "B2C" Then
                    grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grdSOTORDR2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                End If

                If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                    grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                End If

                If EntryMode = "E" Then
                    '0000366445
                    If Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                        ASCMAIN1.sql = "Select Distinct PO_ORDER_NO from POTORDR2" _
                            & " where ORDR_NO = '" & ORDR_NO & "' and (PO_QTY_SHP <> 0 or PO_QTY_REC <> 0)"
                        For Each rowPOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
                            Dim PO_ORDER_NO As String = rowPOTORDR2.Item("PO_ORDER_NO")
                            grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

                            MsgBox("There are restrictions when editing a BTB Order which has been Shipped" _
                                   & vbCrLf & " (ie, no Adding Lines)" _
                                   & vbCrLf & "and Purchase Order No " & PO_ORDER_NO & " has been Shipped", MsgBoxStyle.OkOnly, "Verification")
                            'Exit Sub
                            Exit For
                            'EMsg &= vbCr & "Purchase Order No " & PO_ORDER_NO & " has been Shipped"
                        Next
                    End If
                End If


                Set_Read_Only(splComments.Panel1, False)

                If EntryMode <> "E" Then
                    Set_Read_Only(grpSHIPTO, False)
                    If CUST_DC_NO = "" Then
                        Absx1.optFor("ORDR_ADDR_TYPE_ST").Value = "MK"
                    Else
                        Absx1.optFor("ORDR_ADDR_TYPE_ST").Value = "DC"
                    End If
                Else
                End If

                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    Set_Read_Only(grpSHIPTO, False)
                    '     Set_Read_Only_for_ctl("", True)
                End If

            End If

            Set_Read_Only_for_ctl(optShipTo, (CUST_DC_NO = ""))
            'If CUST_DC_NO = "" Then
            '    frmShipToOption.Enabled = False
            'Else
            '    frmShipToOption.Enabled = True
            'End If
            Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), (EntryMode = "V"))
            Set_Read_Only_for_ctl(Absx1.optFor("ORDR_SOURCE"), True)
            tabHeader.Tabs("EDI").Visible = Absx1.optFor("ORDR_SOURCE").Value = "E"
            tabHeader.Tabs("Audit Trail").Visible = Not (EntryMode = "N")

            If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                lblCURR_CODE.Text = CURR_CODE
                lblCURR_CODE.Visible = True
            Else
                lblCURR_CODE.Visible = False
            End If

        Else
            Clear_Record()
        End If

        If ScreenMode Then

            Set_XFR_Visibility(Absx1.optFor("ORDR_TYPE_CODE").Value = "XFR")

            With grdSOTORDRR.DisplayLayout.Bands(0)

                For Each COLUMN_NAME As String In New String() _
              {"ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_ALLO", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                    .Columns(COLUMN_NAME).Hidden = (EntryMode = "N")
                Next
            End With

            With grdSOTORDR2.DisplayLayout.Bands(0)
                Dim sample_or_transfer As Boolean = (Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "SAM" Or Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR")

                '.Columns("COMM_RATE").Hidden = (EntryMode <> "V") Or Not (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
                '.Columns("DISC_AMT").Hidden = (EntryMode <> "V") Or Not (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
                '.Columns("DISC_PCT").Hidden = (EntryMode <> "V") Or Not (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")

                For Each COLUMN_NAME As String In New String() _
              {"ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_ALLO", "ORDR_QTY_SHIP", "ORDR_QTY_CANC",
               "ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}
                    .Columns(COLUMN_NAME).Hidden = (EntryMode = "N") Or (COLUMN_NAME.StartsWith("ORDR_AMT") And sample_or_transfer)
                Next
                .Columns("ORDR_RELEASE_AVAIL").Hidden = (EntryMode = "N") ' Or sample_or_transfer
                .Columns("ORDR_UNIT_PRICE").Hidden = sample_or_transfer
                .Columns("ORDR_UNIT_PRICE_CALC").Hidden = sample_or_transfer
                .Columns("ORDR_UNIT_PRICE_MANUAL").Hidden = sample_or_transfer
                .Columns("ORDR_PRICE_SOURCE").Hidden = sample_or_transfer

                .Columns("RANGE_STYLE_CODE").Hidden = (ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" <> "1")
                .Columns("STYLE_CODE_SUB").Hidden = (EntryMode = "N") Or (ROWs("SOTPARM1").Item("SO_PARM_SUB_STYLES") & "" <> "1")
                .Columns("X").Hidden = InquiryMode Or (EntryMode <> "E")

                '  splOrderDetails.Panel2Collapsed = (EntryMode = "N" Or EntryMode = "E")
                If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "XFR" Then
                    If ASCMAIN1.CLIENT = "VAN" Then
                        splOrderDetails.Panel2Collapsed = False
                    Else
                        splOrderDetails.Panel2Collapsed = True
                    End If
                    .Columns("RANGE_STYLE_CODE").Hidden = True
                    .Columns("STYLE_CODE_SUB").Hidden = True
                    .Columns("X").Hidden = True

                    .Columns("ORDR_AMT").Hidden = True
                    .Columns("ORDR_AMT_OPEN").Hidden = True
                    .Columns("ORDR_AMT_ALLO").Hidden = True
                    .Columns("ORDR_AMT_PICK").Hidden = True
                    .Columns("ORDR_AMT_SHIP").Hidden = True
                    .Columns("ORDR_AMT_CANC").Hidden = True
                Else
                    'splOrderDetails.Panel2Collapsed = True
                End If

            End With


            If InquiryMode Or (EntryMode <> "N" And EntryMode <> "E") Then
                chkKeepSupplier.Visible = False
            Else
                If ASCMAIN1.CLIENT = "NYA" Then
                    chkKeepSupplier.Visible = True
                    chkKeepSupplier.Checked = True
                Else
                    chkKeepSupplier.Visible = False
                End If
            End If

        Else
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If EntryMode = "E" Or EntryMode = "N" Then
                Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "BTB")
            End If
        End If

        grpICTSIZE1.Visible = (EntryMode = "N" Or EntryMode = "E")

        If ASCMAIN1.CLIENT = "NYA" Then
            If ScreenMode Then

                Dim isUSD As Boolean = (CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"))
                With grdSOTORDR2.DisplayLayout.Bands(0)
                    .Columns("ORDR_UNIT_PRICE_CURR").Hidden = isUSD
                    .Columns("ORDR_AMT_CURR").Hidden = isUSD
                    If Not isUSD Then
                        .Columns("ORDR_UNIT_PRICE_CURR").Header.Caption = CURR_CODE & " Net"
                        .Columns("ORDR_AMT_CURR").Header.Caption = CURR_CODE & " Amt"
                        .Columns("ORDR_AMT_CURR").Header.SetVisiblePosition(.Columns("ORDR_UNIT_PRICE_CURR").Header.VisiblePosition + 1, False)
                    End If
                End With

            End If
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            chkCUST_FACTOR_IND.Enabled = False
            chkCUST_FACTOR_IND.Checked = False
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("ORDR_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.optFor("ORDR_TYPE_CODE").Value = "REG"
        Absx1.txtFor("INV_NO").Text = ""
        CUST_CODE = ""
        ORDR_NO = ""

        txtPICK_NO.Text = ""
        txtINV_NO.Text = ""
        ORDR_NO_to_copy = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDR2", "SOTORDR3", "SOTORDR4", "SOTORDR5", "TATEVNT1", "SOTORDXR",
             "SOTORDRR", "SOTORDR9", "SOTPICK1", "SOTPICK2", "SOTORDRG",
             "SOTORDP1", "SOTORDP2", "SOTORDRB", "SOTORDRI", "SOTORDC1", "SOTORDC2",
             "SOTCART1", "SOTCART2", "SOTWORK1", "SOTWORK2", "SOTORDR1_HOLDS", "POTORDR1", "POTORDR2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Toggle_Customer_Style_Fields(False)
        Toggle_Disc_Comm_Fields(False)
        multiple_order_maintenance = False
        lblORDR_NO.Text = "Order No"

        chkExportInfo.Checked = False
        SET_PROFORMA_CONTROLS()
        chkReleaseNow.Checked = False
        chkAllocate.Checked = False

        'Load_SOTORDPX()
        tabSOTORDRX.Tabs("Pro-Forma Invoices").Tag = "*"

        Load_SOTORDRX()
        If multi_store_is_active Then Setup_MS(False)
        multistore_changes_were_made_to_qty = False
        chkShortView.Checked = False

        PO_ORDER_NOs.Clear()

        lblPromo.Visible = False
        lblPromo.Text = ""
        btnShowPromo.Visible = False
    End Sub

    Sub Load_Record()

        grdPOTORDR1.Parent = splBTB.Panel1

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ORDR_LNO_ctr = 0

        If EntryMode = "N" Then
            Init_Record()
            dst.Tables("SOTORDR2_ORIG").Rows.Clear()

        Else
            rowSOTORDR1 = Fill_Record("SOTORDR1", ORDR_NO)
            Fill_Records("SOTORDR2", ORDR_NO)

            If ASCMAIN1.CLIENT = "RGIx" Then
                ' this code might belong in SOR routines instead of just when calling up an order in SOI
                ASCMAIN1.sql = $"Select * from ICTSTDQ3 where ORDR_GROUP_NO = '{ORDR_GROUP_NO}'"
                For Each rowICTSTDQ3 As DataRow In ASCDATA1.GetDataTable().Select("")
                    Dim STYLE_CODE As String = rowICTSTDQ3.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowICTSTDQ3.Item("COLOR_CODE")

                    Dim ORDR_RELEASE_AVAIL As Date
                    Dim gotone As Boolean = False
                    For I As Integer = 1 To 4
                        If Val(rowICTSTDQ3.Item($"QTY_{CStr(I)}") & "") > 0 Then
                            ORDR_RELEASE_AVAIL = rowICTSTDQ3.Item($"DATE_{CStr(I)}")
                            gotone = True
                            Exit For
                        End If
                    Next
                    If gotone Then
                        Dim sqlw As String = $"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'"
                        For Each row2 As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
                            If row2.Item("ORDR_RELEASE_AVAIL") & "" <> "" _
                                AndAlso Format(row2.Item("ORDR_RELEASE_AVAIL") & "", "yyyyMMdd") = Format(ORDR_RELEASE_AVAIL, "yyyyMMdd") Then
                                ' same date
                            Else
                                row2.Item("ORDR_RELEASE_AVAIL") = ORDR_RELEASE_AVAIL
                                If ASCMAIN1.Running_in_VS Then
                                End If
                            End If
                        Next
                    End If
                Next
            End If


            If EntryMode = "V" Or EntryMode = "E" Then
                TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "ORDR" & EntryMode, "Order Called up to " & IIf(EntryMode = "V", "View", "Edit"))
            End If
            '        .Columns("ORDR_UNIT_COST").Hidden = Not (EntryMode = "V")

            If rowSOTORDR1.Item("ORDR_STATUS") = "F" Then
                ASCMAIN1.sql = "Select INV_LNO, Sum (ORDR_UNIT_COST * ORDR_QTY_SHIP) CGS, Sum (ORDR_QTY_SHIP) QTY from SOTINVH1,SOTINVH2 where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO and SOTINVH1.ORDR_NO = '" & ORDR_NO & "' group by SOTINVH2.INV_LNO"
                Dim tbl As DataTable = ASCDATA1.GetDataTable
                For Each row As DataRow In tbl.Rows
                    Dim INV_LNO As Integer = Val(row.Item("INV_LNO") & "")
                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, INV_LNO})
                    If rowSOTORDR2 IsNot Nothing Then
                        Dim QTY As Int64 = Val(row.Item("QTY") & "")
                        Dim CGS As Decimal = Val(row.Item("CGS") & "")
                        Dim ORDR_UNIT_COST As Decimal = 0
                        If QTY <> 0 Then ORDR_UNIT_COST = CGS / QTY
                        rowSOTORDR2.Item("ORDR_UNIT_COST") = ORDR_UNIT_COST
                    End If
                Next
            End If

            'If ASCMAIN1.Running_in_VS Then Stop ' Sql = "Select SOTORDR2.*, 1 RANGE_STYLE_QTY_PER_PP from SOTORDR2 where ORDR_NO = '" & ORDR_NO_x & "'"
            ORDR_LNO_ctr = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "")

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                Dim M As String = "###,##0.00"
                For Each row As DataRow In dst.Tables("SOTORDR2").Select("")
                    Dim ORDR_UNIT_PRICE As Decimal = Val(row.Item("ORDR_UNIT_PRICE") & "")
                    If Format(ORDR_UNIT_PRICE, "###.00") & "00" <> Format(ORDR_UNIT_PRICE, "###.0000") Then
                        M = "###.0000"
                        Exit For
                    End If
                Next
                grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = M
            End If

            If EntryMode = "V" Then
                grdSOTORDR2.Text = "Order " & ORDR_NO & " Details for Customer PO " & Absx1.txtFor("ORDR_CUST_PO").Text & " - " & CUST_CODE & "-" & rowSOTORDR1.Item("CUST_NAME")
            Else
                grdSOTORDR2.Text = "Order Details"

            End If

            Fill_Records("SOTORDR2_ORIG", ORDR_NO)
            Fill_Records("SOTORDR3", ORDR_NO)
            Fill_Records("SOTORDR4", ORDR_NO)
            Fill_Records("SOTORDR5", ORDR_NO)
            Fill_Records("SOTORDR9", ORDR_NO)

            ASCMAIN1.sql = "Select SOTORDC1.*, ARTCCPA1.CUST_CREDIT_CARD_LAST4, ARTCCPA1.CCPA_DATE_VOID" _
                 & " from SOTORDC1, ARTCCPA1 " _
                 & " where SOTORDC1.ccpa_no = ARTCCPA1.ccpa_no (+)" _
                 & " and SOTORDC1.ORDR_NO = '" & ORDR_NO & "'"
            Fill_Records("SOTORDC1", String.Empty, True, ASCMAIN1.sql)
            Fill_Records("SOTORDC2", ORDR_NO)

            If multiple_order_maintenance Then
                If EntryMode = "E" Then
                    multistore_OK_TO_UPDATE = True
                    dst.Tables("SOTORDRQ_KEY").Rows.Clear()

                    dst.Tables("SOTORDRB").Rows.Clear()
                    grdSOTORDRB.DisplayLayout.Bands(0).Summaries.Clear()
                    If dst.Tables("SOTORDRB").Columns.Contains("QTY_000") Then
                        dst.Tables("SOTORDRB").Columns("QTY_000").Expression = ""
                    End If

                    With dst.Tables("SOTORDRB")
                        For DCOLi As Integer = .Columns.Count - 1 To 0 Step -1
                            If .Columns(DCOLi).ColumnName.StartsWith("QTY_") Then

                                .Columns.Remove(.Columns(DCOLi).ColumnName)
                            End If
                        Next
                    End With

                    Create_Summary(grdSOTORDRB, "CUST_STORE_NO", "Count")

                    grdSOTORDRB.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTORDRB.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

                    Dim COL As Integer = 0

                    Dim C As DataColumn = dst.Tables("SOTORDRB").Columns.Add("QTY_" & Format(COL, "000"), GetType(System.Int64))
                    With grdSOTORDRB.DisplayLayout.Bands(0).Columns("QTY_" & Format(COL, "000"))
                        .Header.Caption = "Total"
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Width = 80
                        .Hidden = False
                        .Format = "#,##0"
                        .CellAppearance.BackColor = Drawing.Color.Beige
                        .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                        .Header.Appearance.BackColor = Drawing.Color.White
                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        .Header.Appearance.TextHAlign = HAlign.Right
                        .CellAppearance.TextHAlign = HAlign.Right
                        Create_Summary(grdSOTORDRB, "QTY_" & Format(COL, "000"))
                    End With

                    If multiple_order_type = "ORDR_GROUP_NO" Then
                        Fill_Records("SOTORDRB", ORDR_GROUP_NO)
                    Else
                        ASCMAIN1.sql = sqlSOTORDRB & " and SOTORDR1.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
                        Fill_Records("SOTORDRB", "", True, ASCMAIN1.sql)
                        For Each rowSOTORDRB As DataRow In dst.Tables("SOTORDRB").Select("")
                            Dim ORDR_GROUP_NO As String = rowSOTORDRB.Item("ORDR_GROUP_NO")
                            If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                        Next
                    End If
                    rowSOTORDR0 = Fill_Record("SOTORDR0", ORDR_GROUP_NO)
                    rowSOTORDR0.Item("ORDR_HOLD") = rowSOTORDR1.Item("ORDR_HOLD")
                    rowSOTORDR0.Item("REASON_CODE") = rowSOTORDR1.Item("REASON_CODE")
                    rowSOTORDR0.Item("ORDR_ADDR_TYPE_ST") = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")

                    rowSOTORDR0.Item("ORDR_HOLD_REASON") = rowSOTORDR1.Item("ORDR_HOLD_REASON")
                    rowSOTORDR0.Item("ORDR_ARRIVAL_DATE") = rowSOTORDR1.Item("ORDR_ARRIVAL_DATE")
                    rowSOTORDR0.Item("ORDR_LAST_ARRIVAL_DATE") = rowSOTORDR1.Item("ORDR_LAST_ARRIVAL_DATE")
                    rowSOTORDR0.Item("ORDR_SHIP_INSTR") = rowSOTORDR1.Item("ORDR_SHIP_INSTR")
                    rowSOTORDR0.Item("ORDR_INV_COMMENT") = rowSOTORDR1.Item("ORDR_INV_COMMENT")
                    rowSOTORDR0.Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
                    rowSOTORDR0.Item("SHIP_VIA_CODE") = rowSOTORDR1.Item("SHIP_VIA_CODE")
                    rowSOTORDR0.Item("FRT_TERMS") = rowSOTORDR1.Item("FRT_TERMS")
                    rowSOTORDR0.Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE")
                    rowSOTORDR0.Item("CUST_FACTOR_IND") = rowSOTORDR1.Item("CUST_FACTOR_IND")

                    dst.Tables("SOTORDRQ").Rows.Clear()
                    For Each row As DataRow In dst.Tables("SOTORDRB").Select("")
                        Dim ORDR_NO_MS As String = row.Item("ORDR_NO")
                        Fill_Records("SOTORDRQ", ORDR_NO_MS, False)
                    Next

                    Dim STYLE_CODE As String = ""
                    Dim COLOR_CODE As String = ""

                    ' Dim ORDR_LNO As Integer = 0
                    Dim STYLE_CODE_SUB As String = ""
                    Dim STYLE_CODE_ORIG As String = ""
                    Dim STYLE_KEY As String = ""
                    Dim T As String = ""
                    Dim CN As String = ""

                    Dim rowSOTORDRI As DataRow

                    Dim SRT As String = "STYLE_KEY" ' "STYLE_CODE_ORIG,STYLE_CODE,COLOR_CODE"
                    For Each rowSOTORDRQ As DataRow In dst.Tables("SOTORDRQ").Select("", SRT)
                        Dim ORDR_NO As String = rowSOTORDRQ.Item("ORDR_NO")
                        Dim ORDR_LNO As Int32 = Val(rowSOTORDRQ.Item("ORDR_LNO") & "")

                        If rowSOTORDRQ.Item("STYLE_KEY") <> STYLE_KEY Then
                            STYLE_KEY = rowSOTORDRQ.Item("STYLE_KEY")
                            STYLE_CODE_ORIG = rowSOTORDRQ.Item("STYLE_CODE_ORIG")
                            STYLE_CODE = rowSOTORDRQ.Item("STYLE_CODE")
                            COLOR_CODE = rowSOTORDRQ.Item("COLOR_CODE")
                            COL += 1

                            'ORDR_LNO = 0 ' Val(rowSOTORDRQ.Item("ORDR_LNO") & "") - cannot rely on consistent line no
                            STYLE_CODE_SUB = rowSOTORDRQ.Item("STYLE_CODE_SUB") & ""

                            rowSOTORDRI = dst.Tables("SOTORDRI").NewRow
                            rowSOTORDRI.Item("STYLE_KEY") = STYLE_KEY
                            rowSOTORDRI.Item("STYLE_CODE_ORIG") = STYLE_CODE_ORIG
                            rowSOTORDRI.Item("STYLE_CODE") = STYLE_CODE

                            rowSOTORDRI.Item("CUST_STYLE_CODE") = rowSOTORDRQ.Item("CUST_STYLE_CODE")
                            rowSOTORDRI.Item("CUST_UPC") = rowSOTORDRQ.Item("CUST_UPC")
                            rowSOTORDRI.Item("CUST_SKU") = rowSOTORDRQ.Item("CUST_SKU")
                            rowSOTORDRI.Item("CUST_SIZE_CODE") = rowSOTORDRQ.Item("CUST_SIZE_CODE")

                            rowSOTORDRI.Item("COLOR_CODE") = COLOR_CODE
                            rowSOTORDRI.Item("CUST_COLOR_CODE") = rowSOTORDRQ.Item("CUST_COLOR_CODE")

                            '   rowSOTORDRI.Item("ORDR_LNO") = ORDR_LNO
                            rowSOTORDRI.Item("STYLE_CODE_SUB") = STYLE_CODE_SUB
                            rowSOTORDRI.Item("COL") = COL

                            rowSOTORDRI.Item("ORDR_UNIT_PRICE") = rowSOTORDRQ.Item("ORDR_UNIT_PRICE")

                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            rowSOTORDRI.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                            rowSOTORDRI.Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
                            rowSOTORDRI.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")

                            rowSOTORDRI.Item("ORDR_QTY") = 0
                            rowSOTORDRI.Item("ORDR_QTY_OPEN") = 0
                            rowSOTORDRI.Item("ORDR_QTY_PICK") = 0
                            rowSOTORDRI.Item("ORDR_QTY_SHIP") = 0
                            rowSOTORDRI.Item("ORDR_QTY_CANC") = 0
                            rowSOTORDRI.Item("ORDR_QTY_ORIG") = 0

                            dst.Tables("SOTORDRI").Rows.Add(rowSOTORDRI)

                            CN = "QTY_" & Format(COL, "000")
                            C = dst.Tables("SOTORDRB").Columns.Add(CN, GetType(System.Int64))
                            With grdSOTORDRB.DisplayLayout.Bands(0).Columns(CN)
                                .Header.Caption = STYLE_CODE & ":" & COLOR_CODE
                                .Width = 80
                                .Hidden = False
                                .Format = "#,##0"
                                .Header.Appearance.TextHAlign = HAlign.Right
                                .CellAppearance.TextHAlign = HAlign.Right
                                ' Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                                .Header.ToolTipText = STYLE_CODE & ":" & COLOR_CODE & vbCrLf & rowICTSTYL1.Item("STYLE_DESC") _
                                    & "Customer Style: " & rowSOTORDRI.Item("CUST_STYLE_CODE") _
                                    & "Customer UPC: " & rowSOTORDRI.Item("CUST_UPC") _
                                    & " Customer SKU: " & rowSOTORDRI.Item("CUST_SKU") _
                                    & " Customer Size: " & rowSOTORDRI.Item("CUST_SIZE_CODE")

                                rowSOTORDRI.Item("COLOR_CODE") = COLOR_CODE
                                rowSOTORDRI.Item("CUST_COLOR_CODE") = rowSOTORDRQ.Item("CUST_COLOR_CODE")

                                Create_Summary(grdSOTORDRB, CN)
                                T &= " + ISNULL(" & CN & ",0)"
                                .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                                .Header.Appearance.BackColor = Drawing.Color.White
                                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            End With
                        Else
                            rowSOTORDRI = dst.Tables("SOTORDRI").Rows.Find(New String() {STYLE_KEY}) ' STYLE_CODE_ORIG ' STYLE_CODE, COLOR_CODE})
                        End If

                        If dst.Tables("SOTORDRQ_KEY").Rows.Find(New Object() {ORDR_NO, STYLE_KEY}) Is Nothing Then
                            dst.Tables("SOTORDRQ_KEY").Rows.Add(New Object() {ORDR_NO, STYLE_KEY, ORDR_LNO})
                        Else
                            If multistore_OK_TO_UPDATE Then
                                MsgBox("Problem with Order " & ORDR_NO & ", Line " & ORDR_LNO, MsgBoxStyle.OkOnly, "Integrity Issue with Multi-Store Edit")
                            End If
                            multistore_OK_TO_UPDATE = False
                        End If

                        rowSOTORDRI.Item("ORDR_QTY") += Val(rowSOTORDRQ.Item("ORDR_QTY") & "")
                        rowSOTORDRI.Item("ORDR_QTY_OPEN") += Val(rowSOTORDRQ.Item("ORDR_QTY_OPEN") & "")
                        rowSOTORDRI.Item("ORDR_QTY_PICK") += Val(rowSOTORDRQ.Item("ORDR_QTY_PICK") & "")
                        rowSOTORDRI.Item("ORDR_QTY_SHIP") += Val(rowSOTORDRQ.Item("ORDR_QTY_SHIP") & "")
                        rowSOTORDRI.Item("ORDR_QTY_CANC") += Val(rowSOTORDRQ.Item("ORDR_QTY_CANC") & "")
                        '  rowSOTORDRI.Item("ORDR_QTY_ORIG") += Val(rowSOTORDRQ.Item("ORDR_QTY_ORIG") & "")

                        Dim rowSOTORDRB As DataRow = dst.Tables("SOTORDRB").Rows.Find(ORDR_NO)
                        rowSOTORDRB.Item(CN) = Val(rowSOTORDRB.Item(CN) & "") + Val(rowSOTORDRQ.Item("ORDR_QTY_OPEN") & "")
                    Next
                    Setup_Multiple_Order_Grid(False)
                    '  Fill_Records("SOTORDRQ", ORDR_GROUP_NO)
                    dst.Tables("SOTORDRB").Columns("QTY_000").Expression = Mid(T, 4)

                    Sort_grdColumns(grdSOTORDRB, "CUST_STORE_NO")
                    Sort_grdColumns(grdSOTORDRI, "STYLE_KEY")

                    If Not multistore_OK_TO_UPDATE Then
                        MsgBox("There are integrity issues with the details of this order and the functional capabilities of Multiple Order Maintenance" & vbCrLf & vbCrLf & "Update will not be permitted", MsgBoxStyle.OkOnly, "Warning")
                    End If
                End If


            Else
                Dim ORDR_REL_HOLD_CODES As String = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & ""
                If ORDR_REL_HOLD_CODES <> "" Then
                    For I As Integer = 1 To ORDR_REL_HOLD_CODES.Length
                        dst.Tables("SOTORDR1_HOLDS").Rows.Add(New String() {ORDR_NO, Mid(ORDR_REL_HOLD_CODES, I, 1)})
                    Next
                End If
            End If

        End If


        Fill_Records("SOTORDXR", ORDR_NO)
        Sort_grdColumns(grdSOTORDXR, "INIT_DATE".ToLower)

        Fill_Records("SOTWORK1", ORDR_NO)
        Fill_Records("SOTWORK2", ORDR_NO)

        If Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
            Fill_Records("SOTORDP1", ORDR_NO)
            Fill_Records("SOTORDP2", ORDR_NO)
            Setup_grdSOTORDP2()

            Fill_Records("SOTINVH1", ORDR_NO)
        Else
            ' IF COMING IN FROM THE PRO-FORMA INVOICES TAB, SOTORDP2 MAY HAVE DATA IN OT
            dst.Tables("SOTORDP1").Rows.Clear()
            dst.Tables("SOTORDP2").Rows.Clear()
        End If

        Sort_grdColumns(grdSOTORDP1, "INV_NO")

        'lblCURR_CODE.Text = rowSOTORDR1.Item("CURR_CODE")
        'grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Header.Caption = "Price " & rowSOTORDR1.Item("CURR_CODE")

        'If rowSOTORDR1.Item("ORDR_SOURCE") & "" = "E" Then
        Load_EDI_Documents()
        'End If


        ASCMAIN1.sql = "Select POTORDR1.* from POTORDR1 where POTORDR1.ORDR_NO = '" & ORDR_NO & "'"
        Fill_Records("POTORDR1", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select POTORDR2.*,ICTSTYL1.CASE_CUBE from POTORDR2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE and POTORDR2.PO_ORDER_NO in (Select PO_ORDER_NO from POTORDR1 where ORDR_NO = '" & ORDR_NO & "')"
        Fill_Records("POTORDR2", "", True, ASCMAIN1.sql)

        Dim rowSOTORDRG As DataRow = Fill_Record("SOTORDRG", ORDR_GROUP_NO)
        If rowSOTORDRG IsNot Nothing AndAlso rowSOTORDRG.Item("ORDR_REL_SHORT") & "" = "1" Then
            chkReleaseNow.Checked = True
        Else
            chkReleaseNow.Checked = False
        End If

        If rowSOTORDRG IsNot Nothing AndAlso rowSOTORDRG.Item("ORDR_REL_ACTION_DATE") & "" <> "" Then
            dteORDR_REL_ACTION_DATE.Value = CDate(rowSOTORDRG.Item("ORDR_REL_ACTION_DATE"))
        Else
            dteORDR_REL_ACTION_DATE.Value = Nothing
        End If

        CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
        CUST_STORE_NO = rowSOTORDR1.Item("CUST_STORE_NO") & ""
        CUST_DC_NO = rowSOTORDR1.Item("CUST_DC_NO") & ""
        ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO") & ""
        ORDR_GROUP_NOs.Clear()
        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        CUST_BILL_TO_CUST = rowSOTORDR1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST <> CUST_CODE Then
            rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
        Else
            rowARTCUST1_BT = rowARTCUST1
        End If

        'Absx1.txtFor("CUST_ROUTING_INST").Text = rowARTCUST1.Item("CUST_ROUTING_INST") & ""

        'rowARTCUST2 = Fill_Record("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})

        ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "' and CUST_ADDR_TYPE = 'MK'"
        Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)
        rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_STORE_NO})

        'ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "' and CUST_ADDR_TYPE = 'DC'"
        'Fill_Records("ARTCUST2", "", False, ASCMAIN1.sql)

        'Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
        ' Sort_grdColumns(grdSOTORDR3, "ORDR_DIV_CODE")
        Sort_grdColumns(grdSOTORDR4, "ORDR_CLNO")
        Sort_grdColumns(grdSOTORDRR, "ORDR_LNO")

        Setup_grdSOTORDR2()
        'Setup_SOTORDR3()

        If EntryMode = "N" Then
            lblStatus.Text = "New Order"
        Else
            Select Case rowSOTORDR1.Item("ORDR_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "P"
                    lblStatus.Text = "In Pick"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
                Case "F"
                    lblStatus.Text = "Shipped"
            End Select
        End If

        sub_grid = ""

        If ASCMAIN1.CLIENT = "NYA" Then
            CURR_CODE = rowSOTORDR1.Item("CURR_CODE") & ""
            If CURR_CODE = "" Then
                CURR_CODE = "USD"
                CURR_EXCH_RATE = 1
            Else
                If CURR_CODE = "USD" Then
                    CURR_EXCH_RATE = 1
                Else
                    'CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, Now.Date)
                    CURR_EXCH_RATE = Val(rowSOTORDR1.Item("CURR_EXCH_RATE") & "")

                End If
            End If
        End If


        If EntryMode = "N" Then
            Absx1.optFor("ORDR_ADDR_TYPE_ST").Value = "DC"
        End If

        CUST_STORE_NOs_multi_store.Clear()

        ' Segregate Range Styles into Separate Table
        Dim RANGE_STYLE_CODE As String
        Dim RANGE_STYLE_LNO As Int32 = 0
        Dim RANGE_STYLE_QTY As Int64
        Dim RANGE_STYLE_PRICE As Decimal
        Dim RANGE_INNER_PACK_QTY As Int64
        Dim f As Int64
        Dim rowSOTORDR2_T As DataRow = Nothing

        dst.Tables("SOTORDRR").Rows.Clear()


        If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                Dim rowICTDUTY1 As DataRow = clsASCBASE1.LookUp("ICTDUTY1", rowICTSTYL1.Item("DUTY_RATE_CODE"))
                rowSOTORDR2.Item("PF_DUTY_HTS_CODE") = rowICTDUTY1.Item("DUTY_HTS_CODE")
                rowSOTORDR2.Item("PF_QTY") = rowSOTORDR2.Item("ORDR_QTY")
            Next

        End If



        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("RANGE_STYLE_CODE is Not Null and RANGE_STYLE_LNO is Not Null and RANGE_STYLE_LNO > 0", "RANGE_STYLE_LNO")
            Dim rowSOTORDRR As DataRow = dst.Tables("SOTORDRR").NewRow
            For i As Integer = 0 To dst.Tables("SOTORDRR").Columns.Count - 1
                Dim COLUMN_NAME As String = dst.Tables("SOTORDRR").Columns(i).ColumnName
                rowSOTORDRR.Item(COLUMN_NAME) = rowSOTORDR2.Item(COLUMN_NAME)
            Next i

            rowSOTORDRR.Item("RANGE_STYLE_LNO") = rowSOTORDRR.Item("ORDR_LNO")

            If RANGE_STYLE_LNO <> Val(rowSOTORDR2.Item("RANGE_STYLE_LNO") & "") Then
                RANGE_STYLE_CODE = rowSOTORDR2.Item("RANGE_STYLE_CODE")
                RANGE_STYLE_LNO = rowSOTORDR2.Item("RANGE_STYLE_LNO")
                Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").Rows.Find(New Object() {ORDR_NO, RANGE_STYLE_LNO})
                If Val(rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") & "") > 1 Then 'Put here on 12/23/2004 by WR.  Needs to be tested for Manual entrerd orders.
                    RANGE_STYLE_QTY = rowSOTORDR9.Item("RANGE_STYLE_PP_QTY")
                Else
                    RANGE_STYLE_QTY = rowSOTORDR9.Item("RANGE_STYLE_QTY")
                End If
                RANGE_STYLE_PRICE = rowSOTORDR9.Item("RANGE_STYLE_PRICE")
                RANGE_INNER_PACK_QTY = rowSOTORDR9.Item("RANGE_INNER_PACK_QTY")

                rowSOTORDR2.Item("ORDR_LNO") = RANGE_STYLE_LNO
                rowSOTORDR2.Item("STYLE_CODE") = ""
                rowSOTORDR2.Item("COLOR_CODE") = ""
                If Val(rowSOTORDR2.Item("ORDR_QTY") & "") <> 0 Then
                    If rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") > 1 Then
                        f = 1
                    Else
                        f = RANGE_STYLE_QTY / Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                    End If
                Else
                    f = 1
                End If
                '  If f <> 1 And ASCMAIN1.Running_in_VS Then Stop ' i THINK F SHOULD BE 1 BECAUSE THE 5 ORDR_QTY_XXXX FIELDS WERE AGGREGATED STRAIGHT FROM VALUES IN SOTORDR2 RANGE STYLE COMPONENT RECORDS IN VB6

                f = 1
                rowSOTORDR2.Item("ORDR_QTY") = RANGE_STYLE_QTY

                'dynSOWORDR2.ITEM("ORDR_QTY").Value = f * dynSOWORDR2.ITEM("ORDR_QTY").Value

                rowSOTORDR2.Item("ORDR_QTY_OPEN") = f * Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                rowSOTORDR2.Item("ORDR_QTY_ALLO") = f * Val(rowSOTORDR2.Item("ORDR_QTY_ALLO") & "")
                rowSOTORDR2.Item("ORDR_QTY_PICK") = f * Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                rowSOTORDR2.Item("ORDR_QTY_SHIP") = f * Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & "")
                rowSOTORDR2.Item("ORDR_QTY_CANC") = f * Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")

                rowSOTORDR2.Item("STYLE_UOM") = rowSOTORDR9.Item("RANGE_STYLE_UOM")
                rowSOTORDR2.Item("STYLE_DESC") = rowSOTORDR9.Item("RANGE_STYLE_DESC")
                rowSOTORDR2.Item("RANGE_STYLE_QTY_PER_PP") = rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP")
                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = RANGE_STYLE_PRICE
                rowSOTORDR2.Item("INNER_PACK_QTY") = RANGE_INNER_PACK_QTY
                rowSOTORDR2.Item("CARTON_PACK_QTY") = 0
                rowSOTORDR2_T = rowSOTORDR2
            Else

                rowSOTORDR2_T.Item("ORDR_QTY_OPEN") += f * Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                rowSOTORDR2_T.Item("ORDR_QTY_ALLO") += f * Val(rowSOTORDR2.Item("ORDR_QTY_ALLO") & "")
                rowSOTORDR2_T.Item("ORDR_QTY_PICK") += f * Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                rowSOTORDR2_T.Item("ORDR_QTY_SHIP") += f * Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & "")
                rowSOTORDR2_T.Item("ORDR_QTY_CANC") += f * Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")

                rowSOTORDR2.Item("RANGE_STYLE_CODE") = "~DELETE~"
            End If

            rowSOTORDRR.Item("ORDR_LNO") = RANGE_STYLE_LNO
            dst.Tables("SOTORDRR").Rows.Add(rowSOTORDRR)
        Next

        ' ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), "RANGE_STYLE_CODE is Not Null")
        ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), "RANGE_STYLE_CODE = '~DELETE~'")
        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

        '---------------------------

        If InquiryMode Or EntryMode = "V" Then
            Fill_Records("SOTPICK1", ORDR_NO)
            Fill_Records("SOTPICK2", ORDR_NO)


            For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select("")
                Dim ORDR_NO As String = rowSOTORDRR.Item("ORDR_NO")
                Dim RANGE_STYLE_LNOX As Integer = Val(rowSOTORDRR.Item("RANGE_STYLE_LNO") & "")
                Dim ORDR_LNO As Integer = Val(rowSOTORDRR.Item("ORDR_LNO") & "")
                Dim sql As String = "ORDR_LNO = " & CStr(RANGE_STYLE_LNOX)
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sql)
                    rowSOTPICK2.Item("ORDR_LNO") = ORDR_LNO
                Next
            Next
            'Set_SOTPICK2()
            Set_SOTCART1()
        End If

        Dim order_is_in_pick As Boolean = (dst.Tables("SOTORDR2").Select("ORDR_QTY_PICK <> 0").Length <> 0)
        Dim order_was_shipped As Boolean = (dst.Tables("SOTORDR2").Select("ORDR_QTY_SHIP <> 0").Length <> 0)

        With grdSOTORDR2.DisplayLayout.Bands(0)
            ' If (EntryMode = "E" Or EntryMode = "N") And Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") = 0 Then
            If (EntryMode = "E" Or EntryMode = "N") And Not order_is_in_pick And Not order_was_shipped Then

                .Columns("PO_COST").CellActivation = UltraWinGrid.Activation.AllowEdit

                ' all of these are handled in Setup_grdSOTORDR2
                'If EntryMode = "E" And Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                '    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                'Else
                '    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                'End If
                'If EntryMode = "E" Then
                '    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                '    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                'End If

                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                Else
                    For Each COLUMN_NAME As String In New String() {"CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_UPC", "CUST_SKU", "STYLE_RETAIL"}
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    Next
                End If
            Else
                .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("PO_COST").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                Else
                    For Each COLUMN_NAME As String In New String() {"CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_UPC", "CUST_SKU", "STYLE_RETAIL"}
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                    Next
                End If
            End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then  ' IF ENTRYMODE = 'E' AFTERROWACTIVATE WILL TAKE CARE OF THINGS (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            'grdSOTORDR2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, False)
        Else
            'grdSOTORDR2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            Set_Read_Only(splHeader, True)
        End If

        If grdSOTORDR2.Rows.Count = 0 Then
            Setup_SubGrid(False, True)

        Else
            If grdSOTORDR2.ActiveRow Is Nothing Then
                grdSOTORDR2.ActiveRow = grdSOTORDR2.Rows(0)
            End If
            If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                Setup_SubGrid(True, False)
            Else
                Setup_SubGrid(False, False)
            End If

        End If

        If EntryMode = "E" Or EntryMode = "N" Then
            If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                UltraExplorerBar1.Groups("Order Details").Items("Add Substitution").Visible = True
                Set_Read_Only_for_ctl(Absx1.txtFor("EDI_APPOINTMENT"), True)
                'frmShipToOption.Enabled = False
                Set_Read_Only_for_ctl(optShipTo, True)
            Else
                grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                UltraExplorerBar1.Groups("Order Details").Items("Add Substitution").Visible = False
                Set_Read_Only_for_ctl(Absx1.txtFor("EDI_APPOINTMENT"), False)
                'frmShipToOption.Enabled = True
                Set_Read_Only_for_ctl(optShipTo, False)
            End If
        Else
            Set_Read_Only_for_ctl(Absx1.txtFor("EDI_APPOINTMENT"), True)
        End If

        Dim rowSOTORDR5 As DataRow = Nothing
        Dim CUST_ADDR_CODEs() As String = {"BY", "BT", "MK", "DC"}
        Dim CUST_ADDR_CODE As String = ""
        If EntryMode = "N" Then
            CUST_ADDR_CODEs = {"BY", "BT", "MK", "DC", "ST"}
        End If
        For Each CUST_ADDR_TYPE As String In CUST_ADDR_CODEs
            Dim row As DataRow = Nothing
            If CUST_ADDR_TYPE = "BY" Then
                row = rowARTCUST1
                CUST_ADDR_CODE = CUST_CODE
            ElseIf CUST_ADDR_TYPE = "BT" Then
                row = rowARTCUST1_BT
                CUST_ADDR_CODE = CUST_BILL_TO_CUST
            ElseIf CUST_ADDR_TYPE = "MK" Then
                row = rowARTCUST2
                CUST_ADDR_CODE = CUST_STORE_NO
            ElseIf CUST_ADDR_TYPE = "DC" Then
                row = LookUp("ARTCUST2", New String() {CUST_CODE, "DC", CUST_DC_NO}, True)
                CUST_ADDR_CODE = CUST_DC_NO
            ElseIf CUST_ADDR_TYPE = "ST" Then

                Dim ORDR_ADDR_TYPE_ST As String = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")

                If Absx1.optFor("ORDR_TYPE_CODE").Value = "XFR" Then
                    If Absx1.txtFor("CUST_STORE_NO").Text & "" <> "000000" Then
                        CUST_ADDR_CODE = IIf(ORDR_ADDR_TYPE_ST = "DC", CUST_DC_NO, CUST_STORE_NO)
                        row = LookUp("ARTCUST2", New String() {CUST_CODE, ORDR_ADDR_TYPE_ST, CUST_ADDR_CODE}, True)
                    Else
                        CUST_ADDR_CODE = Absx1.txtFor("WHSE_CODE_TO").Text
                        row = LookUp("ICTWHSE1", New String() {CUST_ADDR_CODE})
                    End If
                Else
                    CUST_ADDR_CODE = IIf(ORDR_ADDR_TYPE_ST = "DC", CUST_DC_NO, CUST_STORE_NO)
                    row = LookUp("ARTCUST2", New String() {CUST_CODE, ORDR_ADDR_TYPE_ST, CUST_ADDR_CODE}, True)
                End If

            End If

            rowSOTORDR5 = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, CUST_ADDR_TYPE})
            If rowSOTORDR5 Is Nothing Then
                rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                With rowSOTORDR5
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
                    .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE

                    If row IsNot Nothing Then
                        For Each COLUMN_NAME As String In New String() _
                            {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE",
                             "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}
                            Dim COLUMN_NAME_ST As String = COLUMN_NAME
                            If Absx1.optFor("ORDR_TYPE_CODE").Value = "XFR" And CUST_ADDR_TYPE = "ST" And Absx1.txtFor("CUST_STORE_NO").Text & "" = "000000" Then
                                COLUMN_NAME_ST = Replace(COLUMN_NAME, "CUST", "WHSE")
                                COLUMN_NAME_ST = Replace(COLUMN_NAME_ST, "WHSE_NAME", "WHSE_DESC")
                            Else

                            End If
                            .Item(COLUMN_NAME) = row.Item(COLUMN_NAME_ST)
                        Next
                        ' EVENTUALLY, EVERYONE SHOULD HAVE THIS, SO WE WOULD JUST PUT CUST_ADDR3 IN THE COLUMN_NAMEs LIST ABOVE
                        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                            If Absx1.optFor("ORDR_TYPE_CODE").Value = "XFR" And CUST_ADDR_TYPE = "ST" And Absx1.txtFor("CUST_STORE_NO").Text & "" = "000000" Then
                                .Item("CUST_ADDR3") = row.Item("WHSE_ADDR3")
                            Else
                                .Item("CUST_ADDR3") = row.Item("CUST_ADDR3")
                            End If
                        End If
                    End If

                End With
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
            End If
        Next
        txtCUST_ADDR_CODE_BT.Text = "000000"

        If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
            TOTAL_ORDR_AMT = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
            ASCMAIN1.sql = "Select ALLOW_CHANGE_RANGE from EDTSLSP1" _
             & " WHERE CUST_CODE = '" & CUST_CODE & "'"
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                ALLOW_CHANGE_RANGE = "0" ' THIS FIELD IS NOT SET UP YET AT VAN.EDTSLSP1
            Else
                ALLOW_CHANGE_RANGE = ASCDATA1.GetDataValue
            End If
        Else
            TOTAL_ORDR_AMT = 0
            ALLOW_CHANGE_RANGE = ""
        End If

        If chkCUST_FACTOR_IND.Checked Or rowARTCUST1.Item("CUST_FACTOR_IND") & "" = "1" Then
            chkCUST_FACTOR_IND.Visible = True
        Else
            chkCUST_FACTOR_IND.Visible = False
        End If
        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            ' WE SHOULD HAVE "FACTORING" AS A PARAMETER
            chkCUST_FACTOR_IND.Visible = False
        End If

        Display_Totals()

        ' tabMain.SelectedTab = tabMain.Tabs("Names && Addresses")
        tabMain.SelectedTab = tabMain.Tabs("Order Header Info")

        EnforceConstraints(True)

        ASCMAIN1.sql = "Select Max (REV_NO) from SOTORDXR where ORDR_NO = '" & ORDR_NO & "'"
        REV_NO = Val(ASCDATA1.GetDataValue & "")
        If EntryMode = "N" Then
            lblORDR_NO.Text = "Order No - New"

        ElseIf EntryMode = "E" Then
            lblORDR_NO.Text = "Order No - Rev#" & CStr(REV_NO + 1)

        Else
            lblORDR_NO.Text = "Order No - Rev#" & CStr(REV_NO)
        End If

        Load_Events()

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Setup_DISC_AMT(Absx1.txtFor("WHSE_CODE").Text)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Init_Record()
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ORDR_NO = ASCMAIN1.Next_Control_No("ORDR_NO")
        Else
            ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
        End If
        ORDR_GROUP_NO = ""
        ORDR_GROUP_NOs.Clear()

        rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
        With rowSOTORDR1
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_STORE_NO") = CUST_STORE_NO
            .Item("ORDR_CUST_PO") = ORDR_CUST_PO
            .Item("ORDR_DATE") = DATETIME_STAMP.Date
            .Item("ORDR_SOURCE") = "K"
            .Item("ORDR_STATUS") = "O"
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
            .Item("ORDR_TYPE_CODE") = HFs("ORDR_TYPE_CODE")
            If .Item("ORDR_TYPE_CODE") & "" = "" Then
                .Item("ORDR_TYPE_CODE") = "REG"
            ElseIf .Item("ORDR_TYPE_CODE") = "XFR" Then
                .Item("WHSE_CODE_TO") = HFs("WHSE_CODE_TO")
            End If

            Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
            If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
            If WHSE_CODE = "" Then WHSE_CODE = ""
            .Item("WHSE_CODE") = WHSE_CODE

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                If HFs("ORDR_TYPE_CODE") = "BTB" Then
                    WHSE_CODE = BTB_TYPE
                    .Item("WHSE_CODE") = WHSE_CODE
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1.Item("WHSE_TYPE") & "" <> "P" Then
                        If (WHSE_CODE = "NY" Or WHSE_CODE = "NC") Then
                            ' NY IS OK FOR A BTB ORDER - PROBABLY NEED AN ATTRIBUTE IN ICTWHSE1 - WAITING ON WHR FOR DDL OK
                        Else
                            .Item("WHSE_CODE") = DBNull.Value
                        End If
                    End If
                End If
            End If

            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
            Dim rowARTCUST3 As DataRow = LookUp("ARTCUST3", New String() {CUST_CODE, "MK", CUST_STORE_NO, "DC"})
            If rowARTCUST3 IsNot Nothing AndAlso rowARTCUST3.Item("CUST_ADDR_CODE2") & "" <> "" Then
                .Item("ORDR_ADDR_TYPE_ST") = "DC"
                .Item("CUST_DC_NO") = rowARTCUST3.Item("CUST_ADDR_CODE2") & ""
            Else
                .Item("ORDR_ADDR_TYPE_ST") = "MK"
            End If

            .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date

            ' Sold To
            .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
            .Item("SREP_CODE") = SREP_CODE
            .Item("SREP2_CODE") = SREP2_CODE
            .Item("ORDR_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
            If .Item("ORDR_PRIORITY") & "" = "" Then
                .Item("ORDR_PRIORITY") = ROWs("SOTPARM1").Item("SO_PARM_ORDR_PRIORITY")
            End If

            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST

            .Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS") & ""
            .Item("ORDR_SHIP_INSTR") = rowARTCUST1.Item("CUST_SPECIAL_INST") & ""
            .Item("ORDR_INV_COMMENT") = rowARTCUST1.Item("CUST_INV_COMMENT") & ""
            .Item("SHIP_VIA_CODE") = rowARTCUST1.Item("SHIP_VIA_CODE") & ""
            .Item("ORDR_SHIP_COMPLETE") = rowARTCUST1.Item("CUST_SHIP_COMPLETE") & ""
            .Item("CUST_ORDR_CALL_B4_SHIPPING") = rowARTCUST1.Item("CUST_ORDR_CALL_B4_SHIPPING") & ""

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                .Item("ORDR_MESSAGE") = rowARTCUST1.Item("CUST_ROUTING_INST") & ""
            End If

            .Item("CURR_CODE") = "USD"
            .Item("CURR_EXCH_RATE") = 1

            If ASCMAIN1.CLIENT = "NYA" Then
                CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
                If CURR_CODE = "" Then
                    CURR_CODE = "USD"
                    CURR_EXCH_RATE = 1
                Else
                    If CURR_CODE = "USD" Then
                        CURR_EXCH_RATE = 1
                    Else
                        'CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, Now.Date)
                        CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me.ROWs("GLTPARM1"), CURR_CODE, Now.Date)

                    End If
                End If

                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE

            End If

            ' Bill To
            If CUST_BILL_TO_CUST <> CUST_CODE Then
                rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
            Else
                rowARTCUST1_BT = rowARTCUST1
            End If

            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE") & ""
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE") & ""
            .Item("CUST_FACTOR_IND") = rowARTCUST1_BT.Item("CUST_FACTOR_IND") & ""

            If Absx1.optFor("ORDR_TYPE_CODE").Value = "XFR" Then
                .Item("CUST_FACTOR_IND") = "0"
                .Item("FRT_TERMS") = "COL"
            ElseIf Absx1.optFor("ORDR_TYPE_CODE").Value = "SAM" Or Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                If (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") AndAlso Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                    ' LEAVE FACTOR SETTING ALONE FOR NYA BTB
                Else
                    .Item("CUST_FACTOR_IND") = "0"
                End If
            End If

            ' Store
            If rowARTCUST2 IsNot Nothing Then
                .Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_NAME") & ""
            End If

            ' Use Values from Previously Entered Order
            ' NOT A GOOD IDEA WZ 03/21/13
            'If rowSOTORDR1po IsNot Nothing Then
            '    .Item("ORDR_GROUP_NO") = rowSOTORDR1po.Item("ORDR_GROUP_NO")
            '    .Item("ORDR_SHIP_DATE") = rowSOTORDR1po.Item("ORDR_SHIP_DATE")
            '    .Item("ORDR_CANCEL_DATE") = rowSOTORDR1po.Item("ORDR_CANCEL_DATE")
            '    .Item("ORDR_DATE") = rowSOTORDR1po.Item("ORDR_DATE")
            '    .Item("ORDR_DEPT") = rowSOTORDR1po.Item("ORDR_DEPT")
            '    .Item("ORDR_SHIP_INSTR") = rowSOTORDR1po.Item("ORDR_SHIP_INSTR")
            '    .Item("FRT_TERMS") = rowSOTORDR1po.Item("FRT_TERMS")
            'End If

            'If ASCMAIN1.Running_in_VS Then Stop ' Use Copy-From Order Values for SREP_CODE, SREP2_CODE, SHIP_VIA_CODE

        End With
        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

    End Sub

    'Function Add_grdSOTORDR2(STYLE_CODE As String, COLOR_CODE As String, ORDR_UNIT_PRICE As Decimal, ORDR_QTY As Int64) As UltraWinGrid.UltraGridRow

    '    If grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.IsAddRow Then
    '        grdSOTORDR2.ActiveRow = Nothing
    '    End If
    '    grdSOTORDR2.DisplayLayout.Bands(0).AddNew()
    '    With grdSOTORDR2.ActiveRow
    '        .Cells("STYLE_CODE").Value = STYLE_CODE
    '        .Cells("COLOR_CODE").Value = COLOR_CODE
    '        .Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE
    '        .Cells("ORDR_QTY").Value = ORDR_QTY
    '        .Update()
    '    End With
    '    Return grdSOTORDR2.ActiveRow
    'End Function

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDR2", "SOTORDR3", "SOTORDR4", "SOTORDR5", "SOTORDR9"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record_multiple_order_group()

        BeginTrans()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        ' Re-Evaluate Order Status after making changes
        restore_reservation = True

        Dim ALL_ORDERS As New List(Of String)

        For Each rowSOTORDRB As DataRow In dst.Tables("SOTORDRB").Select("")
            Dim ORDR_NO As String = rowSOTORDRB.Item("ORDR_NO")
            ALL_ORDERS.Add(ORDR_NO)

            TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "Change", "Multi-Order")

            Dim changes As Boolean = False

            Fill_Records("SOTORDR2", ORDR_NO)

            For Each rowSOTORDRQ_KEY As DataRow In dst.Tables("SOTORDRQ_KEY").Select("ORDR_NO = '" & ORDR_NO & "'")
                Dim ORDR_LNO As Int32 = Val(rowSOTORDRQ_KEY.Item("ORDR_LNO") & "")
                Dim STYLE_KEY As String = rowSOTORDRQ_KEY.Item("STYLE_KEY")
                Dim rowSOTORDRI As DataRow = dst.Tables("SOTORDRI").Rows.Find(New Object() {STYLE_KEY}) ' STYLE_CODE_ORIG ' {STYLE_CODE, COLOR_CODE})
                Dim COLI As Integer = Val(rowSOTORDRI.Item("COL") & "")

                Dim rowSOTORDR2 As DataRow = Nothing

                If rowSOTORDRI.Item("STYLE_KEY_CLONED_FROM") & "" <> "" Then
                    Dim row() As DataRow = dst.Tables("SOTORDRQ").Select("ORDR_NO = '" & ORDR_NO & "' and STYLE_KEY = '" & rowSOTORDRI.Item("STYLE_KEY_CLONED_FROM") & "'")
                    Dim rowSOTORDR2_cloned_from As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, row(0).Item("ORDR_LNO")})
                    rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                    With rowSOTORDR2
                        .ItemArray = rowSOTORDR2_cloned_from.ItemArray
                        .Item("ORDR_LNO") = ORDR_LNO
                        .Item("STYLE_CODE") = rowSOTORDRI.Item("STYLE_CODE")
                        .Item("STYLE_CODE_SUB") = rowSOTORDRI.Item("STYLE_CODE_SUB")
                        .Item("ORDR_QTY_PICK") = 0
                        .Item("ORDR_QTY_ALLO") = 0
                        .Item("ORDR_QTY_SHIP") = 0
                        .Item("ORDR_QTY_CANC") = 0
                        .Item("RSRV_NO") = DBNull.Value
                        .Item("RSRV_LNO") = DBNull.Value
                    End With
                    dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
                    changes = True
                Else
                    rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                End If

                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim STYLE_CODE_SUB As String = rowSOTORDR2.Item("STYLE_CODE_SUB") & ""
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")

                Dim STYLE_CODE_ORIG As String = STYLE_CODE
                If STYLE_CODE_SUB <> "" Then STYLE_CODE_ORIG = STYLE_CODE_SUB

                Dim ORDR_QTY As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                Dim ORDR_QTY_PICK As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & "")
                Dim ORDR_QTY_CANC As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")

                Dim ORDR_QTY_OPEN_NEW As Int64 = Val(rowSOTORDRB.Item("QTY_" & Format(COLI, "000")) & "")

                If ORDR_QTY_OPEN_NEW <> ORDR_QTY_OPEN Then
                    'ORDR_QTY_CANC = ORDR_QTY_CANC + (ORDR_QTY_OPEN - QTY)
                    ORDR_QTY_OPEN = ORDR_QTY_OPEN_NEW
                    ORDR_QTY_CANC = ORDR_QTY - ORDR_QTY_PICK - ORDR_QTY_SHIP - ORDR_QTY_OPEN
                    If ORDR_QTY_CANC < 0 Then ORDR_QTY_CANC = 0
                    rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = ORDR_QTY_CANC
                    changes = True
                End If

                Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                If ORDR_UNIT_PRICE <> Val(rowSOTORDRI.Item("ORDR_UNIT_PRICE") & "") Then
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Val(rowSOTORDRI.Item("ORDR_UNIT_PRICE") & "")
                    changes = True
                End If

                If (rowSOTORDRI.Item("STYLE_CODE_SUB") & "" <> rowSOTORDR2.Item("STYLE_CODE_SUB") & "") _
                Or (rowSOTORDRI.Item("STYLE_CODE") & "" <> rowSOTORDR2.Item("STYLE_CODE") & "") Then
                    rowSOTORDR2.Item("STYLE_CODE") = rowSOTORDRI.Item("STYLE_CODE")
                    rowSOTORDR2.Item("STYLE_CODE_SUB") = rowSOTORDRI.Item("STYLE_CODE_SUB")
                    changes = True
                End If
            Next

            Dependent_Updates(-1, ORDR_NO)

            If changes Then
                'Dependent_Updates(-1, ORDR_NO)
                Update_Record_TDA("SOTORDR2")
                'Dependent_Updates(1, ORDR_NO)
            End If
        Next

        Dim ORDR_SHIP_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value
        Dim ORDR_CANCEL_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_CANCEL_DATE").Value
        Dim ORDR_ARRIVAL_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_ARRIVAL_DATE").Value
        Dim ORDR_LAST_ARRIVAL_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_LAST_ARRIVAL_DATE").Value

        If Format(ORDR_ARRIVAL_DATE, "MM/dd/yyyy") = "01/01/0001" Then
            ORDR_ARRIVAL_DATE = Nothing
        End If
        If Format(ORDR_LAST_ARRIVAL_DATE, "MM/dd/yyyy") = "01/01/0001" Then
            ORDR_LAST_ARRIVAL_DATE = Nothing
        End If

        Dim ORDR_DEPT As String = Absx1.txtFor("SOTORDR0.ORDR_DEPT").Text
        Dim CUST_FACTOR_IND As String = IIf(Absx1.chkFor("SOTORDR0.CUST_FACTOR_IND").Checked, "1", "0")
        Dim ORDR_HOLD As String = IIf(Absx1.chkFor("SOTORDR0.ORDR_HOLD").Checked, "1", "0")
        Dim ORDR_HOLD_REASON As String = Absx1.txtFor("SOTORDR0.ORDR_HOLD_REASON").Text

        'ORDR_ADDR_TYPE_ST
        Dim ORDR_SHIP_INSTR As String = Absx1.txtFor("SOTORDR0.ORDR_SHIP_INSTR").Text
        Dim ORDR_INV_COMMENT As String = Absx1.txtFor("SOTORDR0.ORDR_INV_COMMENT").Text

        Dim SREP_CODE As String = Absx1.txtFor("SOTORDR0.SREP_CODE").Text
        Dim SREP2_CODE As String = Absx1.txtFor("SOTORDR0.SREP2_CODE").Text
        Dim TERM_CODE As String = Absx1.txtFor("SOTORDR0.TERM_CODE").Text
        Dim WHSE_CODE As String = Absx1.txtFor("SOTORDR0.WHSE_CODE").Text
        Dim SHIP_VIA_CODE As String = Absx1.txtFor("SOTORDR0.SHIP_VIA_CODE").Text
        Dim FRT_TERMS As String = Absx1.txtFor("SOTORDR0.FRT_TERMS").Text
        Dim REASON_CODE As String = Absx1.txtFor("SOTORDR0.REASON_CODE").Text

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCMAIN1.sql = "Update SOTORDR1 " & vbCrLf _
                & " Set ORDR_SHIP_DATE = :PARM1, ORDR_CANCEL_DATE = :PARM2, ORDR_ARRIVAL_DATE = :PARM3, ORDR_LAST_ARRIVAL_DATE = :PARM4" & vbCrLf _
                & ", ORDR_DEPT = :PARM5, CUST_FACTOR_IND = :PARM6, ORDR_HOLD = :PARM7, ORDR_HOLD_REASON = :PARM8" & vbCrLf _
                & ", ORDR_SHIP_INSTR = :PARM9, ORDR_INV_COMMENT = :PARM10" & vbCrLf _
                & ", SREP_CODE = :PARM11, SREP2_CODE = :PARM12, TERM_CODE = :PARM13, WHSE_CODE = :PARM14" & vbCrLf _
                & ", SHIP_VIA_CODE = :PARM15, FRT_TERMS = :PARM16, REASON_CODE = :PARM17" & vbCrLf _
                & ", LAST_DATE = :PARM18, LAST_OPER = :PARM19" & vbCrLf _
                & " where ORDR_GROUP_NO = :PARM20"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DDDDVVVVVVVVVVVVVDVV",
                                New Object() {ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_ARRIVAL_DATE, ORDR_LAST_ARRIVAL_DATE,
                                              ORDR_DEPT, CUST_FACTOR_IND, ORDR_HOLD, ORDR_HOLD_REASON,
                                              ORDR_SHIP_INSTR, ORDR_INV_COMMENT,
                                              SREP_CODE, SREP2_CODE, TERM_CODE, WHSE_CODE,
                                              SHIP_VIA_CODE, FRT_TERMS, REASON_CODE,
                                              DATETIME_STAMP, ASCMAIN1.USER_ID, ORDR_GROUP_NO})

            ASCMAIN1.sql = "Select ORDR_NO from (" & vbCrLf _
                & "Select SOTORDR2.ORDR_NO" & vbCrLf _
                & ", SUM (ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & " from SOTORDR2,SOTORDR1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "  and SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
                & "  and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                & " group by SOTORDR2.ORDR_NO)" & vbCrLf _
                & " where ORDR_QTY_OPEN = 0 AND ORDR_QTY = ORDR_QTY_CANC"
            ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_STATUS = 'C' where ORDR_NO in (" & vbCrLf & ASCMAIN1.sql & ")"
            ASCDATA1.ExecuteSQL()

            ' THIS BLOCK IS NOT NEC SINCE WE DO NOT BRING UP CANCELLED ORDERS FOR EDIT
            ' ALSO NEED TO BE CAREFUL NOT TO MAKE AN ORDER OPEN THAT WAS PURPOSEFULLY CANCLLED
            'ASCMAIN1.sql = "Select ORDR_NO from (" & vbCrLf _
            '    & "Select SOTORDR2.ORDR_NO" & vbCrLf _
            '    & ", SUM (ORDR_QTY) ORDR_QTY" & vbCrLf _
            '    & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            '    & ", SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            '    & " from SOTORDR2,SOTORDR1" & vbCrLf _
            '    & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            '    & "  and SOTORDR1.ORDR_STATUS = 'C'" & vbCrLf _
            '    & "  and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            '    & " group by SOTORDR2.ORDR_NO)" & vbCrLf _
            '    & " where ORDR_QTY_OPEN <> 0"
            'ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_STATUS = 'O' where ORDR_NO in (" & vbCrLf & ASCMAIN1.sql & ")"
            'ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

        Next

        For Each ORDR_NOx As String In ALL_ORDERS
            Dependent_Updates(1, ORDR_NOx)
        Next

        CommitTrans("Update Complete")
    End Sub

    Sub Update_Record()

        BeginTrans()

        restore_reservation = True

        Dim ALL_ORDERS As New List(Of String)
        ALL_ORDERS.Add(ORDR_NO)

        Dim ORDR_NO_ORIG As String = ORDR_NO
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        If rowSOTORDR1.Item("ORDR_GROUP_NO") & "" = "" Then
            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                ORDR_GROUP_NO = ORDR_NO
            Else
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("ORDR_GROUP_NO")
                Else
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                End If

            End If

            rowSOTORDR1.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            ORDR_GROUP_NOs.Clear()
            ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
        End If

        ' Load up SOTORDR5 with Bill-To and Ship-To Address

        If EntryMode = "N" Or EntryMode = "E" Then
            ASCDATA1.DeleteRows("SOTORDR5", "CUST_ADDR_TYPE <> 'BT' and CUST_ADDR_TYPE <> 'ST'")
            'Update_SOTORDR5(ORDR_NO, "BT", txtCUST_ADDR_CODE_BT.Text, "")
            'Update_SOTORDR5(ORDR_NO, rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & "", CUST_DC_NO, CUST_STORE_NO)
        End If

        ' Re-Evaluate Order Status after making changes

        If EntryMode = "E" Then
            Dim ORDR_STATUS As String
            Dim OPEN As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", "") & "")
            Dim PICK As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_PICK)", "") & "")
            Dim SHIP As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_SHIP)", "") & "")
            If OPEN <> 0 Then
                ORDR_STATUS = "O"
            ElseIf PICK <> 0 Then
                ORDR_STATUS = "P"
            ElseIf SHIP <> 0 Then
                ORDR_STATUS = "F"
            Else
                ORDR_STATUS = "C"
            End If
            rowSOTORDR1.Item("ORDR_STATUS") = ORDR_STATUS

            'Dim clsSOCORDR1 As New TAC.SOCORDR1(Me)
            'clsSOCORDR1.EvaluateCancelledDetailLines(dst.Tables("SOTORDR2"))
        End If

        ' Copy Range Style Components to End of Order

        dst.Tables("SOTORDR9").Rows.Clear()

        Dim RANGE_STYLE_CODE As String = ""
        Dim RANGE_STYLE_QTY As Int64
        Dim RANGE_STYLE_PRICE As Decimal
        Dim RANGE_STYLE_PRICE_A As Decimal
        Dim RANGE_INNER_PACK_QTY As Int64
        Dim RANGE_STYLE_TYPE As String = ""
        Dim RANGE_STYLE_QTY_PER_PP As Int64
        Dim EDI_DOC_SEQ_NO As String = ""
        Dim EDI_DTL_SEQ As Int64

        For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select("", "ORDR_LNO")
            Dim ORDR_NO As String = rowSOTORDRR.Item("ORDR_NO")
            Dim ORDR_LNO As Integer = Val(rowSOTORDRR.Item("ORDR_LNO") & "")
            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})

            If RANGE_STYLE_CODE <> rowSOTORDR2.Item("RANGE_STYLE_CODE") Then
                RANGE_STYLE_CODE = rowSOTORDR2.Item("RANGE_STYLE_CODE")
                RANGE_STYLE_QTY = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                RANGE_INNER_PACK_QTY = Val(rowSOTORDR2.Item("INNER_PACK_QTY") & "")
                ' RANGE_CARTON_PACK_QTY = Val(rowSOTORDR2.Item("CARTON_PACK_QTY") & "")
                RANGE_STYLE_PRICE = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                RANGE_STYLE_PRICE_A = Val(rowSOTORDRR.Item("ORDR_UNIT_PRICE") & "")

                RANGE_STYLE_QTY_PER_PP = Val(rowSOTORDR2.Item("RANGE_STYLE_QTY_PER_PP") & "")
                EDI_DOC_SEQ_NO = rowSOTORDR2.Item("EDI_DOC_SEQ_NO") & ""
                EDI_DTL_SEQ = Val(rowSOTORDR2.Item("EDI_DTL_SEQ") & "")

                If RANGE_STYLE_QTY_PER_PP = 0 Then
                    RANGE_STYLE_QTY_PER_PP = 1
                End If
                If RANGE_STYLE_QTY_PER_PP = 1 Then
                    RANGE_STYLE_TYPE = "R" ' RANGE STYLE
                Else
                    RANGE_STYLE_TYPE = "A" ' ASSORTMENT
                End If

                Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").NewRow
                rowSOTORDR9.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR9.Item("RANGE_STYLE_LNO") = ORDR_LNO
                rowSOTORDR9.Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
                rowSOTORDR9.Item("RANGE_STYLE_QTY") = RANGE_STYLE_QTY * RANGE_STYLE_QTY_PER_PP
                rowSOTORDR9.Item("RANGE_STYLE_PRICE") = RANGE_STYLE_PRICE / RANGE_STYLE_QTY_PER_PP
                rowSOTORDR9.Item("RANGE_INNER_PACK_QTY") = RANGE_INNER_PACK_QTY
                'rowSOTORDR9.Item("RANGE_CARTON_PACK_QTY") = RANGE_CARTON_PACK_QTY
                rowSOTORDR9.Item("RANGE_STYLE_DESC") = rowSOTORDR2.Item("STYLE_DESC")
                rowSOTORDR9.Item("RANGE_STYLE_UOM") = rowSOTORDR2.Item("STYLE_UOM")
                rowSOTORDR9.Item("RANGE_STYLE_PP_QTY") = RANGE_STYLE_QTY
                rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE") = RANGE_STYLE_PRICE
                rowSOTORDR9.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                rowSOTORDR9.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") = RANGE_STYLE_QTY_PER_PP
                dst.Tables("SOTORDR9").Rows.Add(rowSOTORDR9)
            End If

            rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
            For i As Integer = 0 To dst.Tables("SOTORDRR").Columns.Count - 1
                rowSOTORDR2.Item(i) = rowSOTORDRR.Item(i)
            Next i
            'rowSOTORDR2.Item("RANGE_STYLE_QTY") = RANGE_STYLE_QTY
            'rowSOTORDR2.Item("RANGE_STYLE_PRICE") = RANGE_STYLE_PRICE
            'rowSOTORDR2.Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
            'rowSOTORDR2.Item("RANGE_INNER_PACK_QTY") = RANGE_INNER_PACK_QTY
            'rowSOTORDR2.Item("RANGE_STYLE_DESC") = rowSOTORDR2.Item("STYLE_DESC")
            'rowSOTORDR2.Item("RANGE_STYLE_UOM") = rowSOTORDR2.Item("STYLE_UOM")

            If Val(rowSOTORDR2.Item("RANGE_STYLE_LNO") & "") <> 0 Then
                rowSOTORDR2.Item("ORDR_LNO") = rowSOTORDR2.Item("RANGE_STYLE_LNO")
            Else
                rowSOTORDR2.Item("ORDR_LNO") = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
            End If

            rowSOTORDR2.Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
            rowSOTORDR2.Item("RANGE_STYLE_LNO") = rowSOTORDRR.Item("ORDR_LNO")
            If EDI_DOC_SEQ_NO <> "" Then
                rowSOTORDR2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
            End If
            If EDI_DTL_SEQ <> 0 Then
                rowSOTORDR2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
            End If
            If RANGE_STYLE_TYPE <> "A" Then
                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = RANGE_STYLE_PRICE
            Else
                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Val(rowSOTORDRR.Item("ORDR_UNIT_PRICE") & "")
            End If
            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
        Next

        ' Delete Range Style Header Records from SOTORDR2
        Dim sqldelete As String = "RANGE_STYLE_CODE is Not Null and ISNULL(STYLE_CODE,'') = ''"
        ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), sqldelete)

        ' Set ORDR_STATUS and ORDR_QTY_ORIG

        Dim SALES_DIVISION_CODE As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            rowSOTORDR2.Item("ORDR_STATUS") = rowSOTORDR1.Item("ORDR_STATUS")
            If EntryMode = "N" Then
                rowSOTORDR2.Item("ORDR_QTY_ORIG") = rowSOTORDR2.Item("ORDR_QTY")
                If SALES_DIVISION_CODE = "" Then
                    Dim ROWICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTORDR2.Item("STYLE_CODE"))
                    SALES_DIVISION_CODE = ROWICTSTYL1.Item("SALES_DIVISION_CODE") & ""
                End If
            End If
        Next

        ' Double-Check SALES_DIVISION_CODE

        If EntryMode = "N" Then
            If SALES_DIVISION_CODE <> rowSOTORDR1.Item("SALES_DIVISION_CODE") & "" And SALES_DIVISION_CODE <> "" Then
                rowSOTORDR1.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
            End If

            rowSOTORDR1.Item("ORDR_ORIG_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
            rowSOTORDR1.Item("ORDR_ORIG_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
        ElseIf EntryMode = "E" Then
            'If rowSOTORDR1.Item("ORDR_PRIORITY") & "" <> rowSOTORDR1.Item("ORDR_PRIORITY", DataRowVersion.Original) & "" Then

            'End If
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Dim ORDR_FOB As String = Get_ORDR_FOB(Absx1.optFor("ORDR_TYPE_CODE").Value, Absx1.txtFor("WHSE_CODE").Text)
            rowSOTORDR1.Item("ORDR_FOB") = ORDR_FOB
        End If

        ' Copy Order to Multiple Stores if Multi-Store Mode is True

        If multi_store_is_active Then
            Clear_Zeroes()

            ' Traverse the Multi-Store grid, and Write ORDR1/2'S for each Store

            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                Dim CUST_STORE_NO As String = rowSOTORDRS.Item("CUST_STORE_NO")
                Dim ORDR_CUST_PO As String = rowSOTORDRS.Item("ORDR_CUST_PO")

                If CUST_STORE_NO = Me.CUST_STORE_NO Then
                    rowSOTORDR1.Item("ORDR_CUST_PO") = ORDR_CUST_PO
                    Update_SOTORDR5(ORDR_NO, rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & "", rowSOTORDR1.Item("CUST_DC_NO") & "", CUST_STORE_NO)
                Else
                    Dim ORDR_NO As String = ""
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        ORDR_NO = ASCMAIN1.Next_Control_No("ORDR_NO")
                    Else
                        ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                    End If

                    ALL_ORDERS.Add(ORDR_NO)
                    ASCMAIN1.Progress("-", ORDR_NO)
                    Dim row As DataRow = dst.Tables("SOTORDR1").NewRow
                    row.ItemArray = rowSOTORDR1.ItemArray
                    row.Item("ORDR_NO") = ORDR_NO
                    dst.Tables("SOTORDR1").Rows.Add(row)

                    Dim rowARTCUST2_MS As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO}, True)
                    Dim rowARTCUST3_MS As DataRow = LookUp("ARTCUST3", New String() {CUST_CODE, "MK", CUST_STORE_NO, "DC"}, True)
                    Dim CUST_DC_NO As String = rowARTCUST3_MS.Item("CUST_ADDR_CODE2") & ""
                    row.Item("ORDR_NO") = ORDR_NO
                    row.Item("CUST_STORE_NO") = CUST_STORE_NO
                    row.Item("ORDR_CUST_PO") = ORDR_CUST_PO
                    row.Item("CUST_DC_NO") = CUST_DC_NO

                    row.Item("CUST_STORE_NAME") = rowARTCUST2_MS.Item("CUST_NAME")

                    Update_SOTORDR5(ORDR_NO, row.Item("ORDR_ADDR_TYPE_ST") & "", CUST_DC_NO, CUST_STORE_NO)

                    Dim ORDR_QTY_orig_store As Int64

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & Me.ORDR_NO & "'")
                        Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim ORDR_QTY As Int64 = Val(rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) & "")
                        If ORDR_QTY <> 0 Then
                            Dim row2 As DataRow = dst.Tables("SOTORDR2").NewRow
                            row2.ItemArray = rowSOTORDR2.ItemArray
                            ORDR_QTY_orig_store = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                            row2.Item("ORDR_NO") = ORDR_NO
                            row2.Item("ORDR_QTY") = ORDR_QTY
                            row2.Item("ORDR_QTY_OPEN") = ORDR_QTY
                            row2.Item("ORDR_QTY_ORIG") = ORDR_QTY
                            dst.Tables("SOTORDR2").Rows.Add(row2)

                            For Each rowSOTORDR3 As DataRow In dst.Tables("SOTORDR3").Select _
                                                                   ("ORDR_NO = '" & Me.ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO))
                                Dim row3 As DataRow = dst.Tables("SOTORDR3").NewRow
                                row3.ItemArray = rowSOTORDR3.ItemArray
                                row3.Item("ORDR_NO") = ORDR_NO
                                Dim factor As Decimal = Val(row3.Item("ORDR_QTY") & "") / Val(row2.Item("ORDR_QTY") & "")
                                row3.Item("ORDR_QTY") = Val(row3.Item("ORDR_QTY") & "") * factor
                                For i3 As Integer = 1 To 12
                                    Dim C As String = "SIZE_QTY_" & Format$(i3, "00")
                                    row3.Item(C) = Val(row3.Item(C) & "") * factor
                                Next i3
                                dst.Tables("SOTORDR3").Rows.Add(row3)
                            Next
                        End If
                    Next

                End If
            Next

            Copy_Records(ALL_ORDERS)
        End If


        ' Remove all Order Details where no qty ordered - SOTORDR2 & SOTORDR3

        ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), "ISNULL(ORDR_QTY,0) = 0 and ISNULL(ORDR_QTY_OPEN,0) = 0 and ISNULL(ORDR_QTY_SHIP,0) = 0 and ISNULL(ORDR_QTY_PICK,0) = 0 and ISNULL(ORDR_QTY_CANC,0) = 0")

        ' Update all Currency Fields

        If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "") / 1 ' CURR_EXCH_RATE
            Next
            For Each rowSOTORDR9 As DataRow In dst.Tables("SOTORDR9").Select("")
                rowSOTORDR9.Item("RANGE_STYLE_PRICE_CURR") = Val(rowSOTORDR9.Item("RANGE_STYLE_PRICE") & "") / 1 ' CURR_EXCH_RATE
                rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE_CURR") = Val(rowSOTORDR9.Item("RANGE_STYLE_PRICE") & "") / 1 ' CURR_EXCH_RATE
            Next
        End If

        'If "USD" <> "USD" Then
        '    Stop
        '    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
        '        rowSOTORDR1.Item("CURR_CODE") = "USD" ' CURR_CODE
        '        rowSOTORDR1.Item("CURR_EXCH_RATE") = 1 ' CURR_EXCH_RATE
        '    Next
        '    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
        '        rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2.Item("ORDR_UNIT_PRICE") / 1 ' CURR_EXCH_RATE
        '    Next
        '    For Each rowSOTORDR9 As DataRow In dst.Tables("SOTORDR9").Select("")
        '        rowSOTORDR9.Item("RANGE_STYLE_PRICE_CURR") = rowSOTORDR9.Item("RANGE_STYLE_PRICE") / 1 ' CURR_EXCH_RATE
        '        rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE_CURR") = rowSOTORDR9.Item("RANGE_STYLE_PRICE") / 1 ' CURR_EXCH_RATE
        '    Next
        'End If

        Record_Event("UPDT", "Sales Order Updated")
        If EntryMode = "E" Then Check_Changed_Fields()
        If EntryMode <> "N" Then Delete_Records()
        Dim SQLD As String = "ORDR_NO = '" & ORDR_NO & "'"
        INIT_LAST("SOTORDR1", False, , True)
        Update_Record_TDA("SOTORDR1", SQLD)
        Update_Record_TDA("SOTORDR2", SQLD)
        Update_Record_TDA("SOTORDR3", SQLD)
        Update_Record_TDA("SOTORDR4", SQLD)
        Update_Record_TDA("SOTORDR5", SQLD)
        Update_Record_TDA("SOTORDR9", SQLD)

        Update_Record_TDA("SOTORDXR")
        Update_Record_TDA("TATEVNT1")

        Update_Record_TDA("SOTWORK1")
        Update_Record_TDA("SOTWORK2")

        If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB" Then
            Update_Record_TDA("SOTORDP1", SQLD)
            Update_Record_TDA("SOTORDP2", SQLD)
        End If

        If EntryMode = "E" Then
            If Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") <> 0 Then
                ASCMAIN1.sql = "Delete from SOTPICK2 where (PICK_NO, PICK_LNO) in (" & vbCrLf _
                    & "Select SOTPICK2.PICK_NO, SOTPICK2.PICK_LNO from SOTPICK1,SOTPICK2" & vbCrLf _
                    & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                    & "   and NVL(SOTPICK1.PICK_STATUS,'F') <> 'P' " & vbCrLf _
                    & "   and NVL(SOTPICK2.PICK_QTY,0) = 0" & vbCrLf _
                    & "   and NVL(SOTPICK2.PICK_QTY_CONF,0) = 0" & vbCrLf _
                    & "   and NVL(SOTPICK2.PICK_QTY_CANC,0) = 0" & vbCrLf _
                    & "   and SOTPICK2.ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                    & "   and (SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO) in " & vbCrLf _
                    & "(Select ORDR_NO, ORDR_LNO from SOTPICK2 where ORDR_NO = '" & ORDR_NO & "' minus" & vbCrLf _
                    & " Select ORDR_NO, ORDR_LNO from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'))"
                ASCDATA1.ExecuteSQL()
            End If
        End If

        For Each ORDR_NOx As String In ALL_ORDERS
            Dependent_Updates(1, ORDR_NOx)
        Next

        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            ' IF WE EVER DO MULTIPLE ORDERS IN A GROUP - WE WILL NEED TO CALL THIS FOR EACH ORDER
            ASCDATA1.ExecuteSP("SOPORDR1_COMM", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Credit_Request()
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else
            If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB" Then
                Create_PO()
            End If
        End If

        Dim rowSOTORDRG As DataRow = Fill_Record("SOTORDRG", ORDR_GROUP_NO)

        If chkReleaseNow.Checked Then
            If rowSOTORDRG Is Nothing Then
                rowSOTORDRG = dst.Tables("SOTORDRG").NewRow
                rowSOTORDRG.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                dst.Tables("SOTORDRG").Rows.Add(rowSOTORDRG)
            End If
            If rowSOTORDRG.Item("ORDR_REL_SHORT") & "" = "1" Then ' Already queued to release short
            Else
                rowSOTORDRG.Item("ORDR_REL_SHORT") = "1"
                rowSOTORDRG.Item("ORDR_REL_SHORT_OPER") = ASCMAIN1.USER_ID
                rowSOTORDRG.Item("ORDR_REL_SHORT_DATE") = DATETIME_STAMP
                rowSOTORDRG.Item("ORDR_REL_SHORT_MIN") = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT_ALLO)", "") & "")
            End If

            Update_Record_TDA("SOTORDRG")
        Else
            If rowSOTORDRG IsNot Nothing AndAlso rowSOTORDRG.Item("ORDR_REL_SHORT") & "" = "1" Then
                rowSOTORDRG.Item("ORDR_REL_SHORT") = "0"
                ' LEAVE OPER AND DATE ALONE - TO DOCUMENT WHO SET THE FLAG LAST - NOT SURE THAT THIS IS SO IMPORTANT
                Update_Record_TDA("SOTORDRG")
            End If
        End If

        If dteORDR_REL_ACTION_DATE.Value & "" <> "" Then
            If rowSOTORDRG Is Nothing Then
                rowSOTORDRG = dst.Tables("SOTORDRG").NewRow
                rowSOTORDRG.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                dst.Tables("SOTORDRG").Rows.Add(rowSOTORDRG)
            End If
            rowSOTORDRG.Item("ORDR_REL_ACTION_OPER") = ASCMAIN1.USER_ID
            rowSOTORDRG.Item("ORDR_REL_ACTION_DATE") = dteORDR_REL_ACTION_DATE.Value

            Update_Record_TDA("SOTORDRG")
        Else
            If rowSOTORDRG IsNot Nothing AndAlso rowSOTORDRG.Item("ORDR_REL_ACTION_DATE") & "" <> "" Then
                rowSOTORDRG.Item("ORDR_REL_ACTION_DATE") = DBNull.Value
                rowSOTORDRG.Item("ORDR_REL_ACTION_OPER") = DBNull.Value
                Update_Record_TDA("SOTORDRG")
            End If
        End If

        If chkAllocate.Checked Then
            CommitTrans()
        Else
            CommitTrans("Update Complete")
        End If

    End Sub

    Sub Credit_Request()
        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
        Dim TERM_CODE As String = rowSOTORDR1.Item("TERM_CODE") & ""
        Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE)
        Dim TERM_TYPE As String = rowTATTERM1.Item("TERM_TYPE") & ""
        Dim CUST_FACTOR_IND As String = rowSOTORDR1.Item("CUST_FACTOR_IND") & ""
        Dim ORDR_HOLD As String = rowSOTORDR1.Item("ORDR_HOLD") & ""
        Dim ORDR_AMT_NOW As Decimal = Val(rowSOTORDR0.Item("ORDR_AMT") & "")
        Dim ORDR_CANCEL_DATE As Date = rowSOTORDR1.Item("ORDR_CANCEL_DATE")

        ' CUST_FACTOR_IND = "1" SHOULD NOT BE ALLOWED IN PREREQ IF ORDR_AMT_NOW = 0 OR TERM_TYPE = "C"
        'If ORDR_AMT_NOW <> 0 And CUST_FACTOR_IND = "1" And ORDR_HOLD <> "1" And TERM_TYPE <> "C" Then
        If CUST_FACTOR_IND = "1" And ORDR_HOLD <> "1" Then
            Dim ORDR_AMT_PRIOR As Decimal = 0
            Dim ORDR_CANCEL_DATE_PRIOR As Date = Now
            If EntryMode = "N" Or EntryMode = "V" Then
                ' do nothing, we are going to process a request for credit
            Else
                ASCMAIN1.sql = "Select Max (EDI_OUTBOUND_DOC_NO) EDI_OUTBOUND_DOC_NO from EDT855O1" & vbCrLf _
                    & " where COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" & vbCrLf _
                    & "   and ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                Dim EDI_OUTBOUND_DOC_NO As String = ASCDATA1.GetDataValue
                If EDI_OUTBOUND_DOC_NO <> "" Then
                    Dim rowEDT855O1 As DataRow = LookUp("EDT855O1", New String() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO})
                    ORDR_AMT_PRIOR = Val(rowEDT855O1.Item("ORDR_AMT") & "")
                    ORDR_CANCEL_DATE_PRIOR = rowEDT855O1.Item("ORDR_CANCEL_DATE")
                End If
            End If
            If ORDR_AMT_NOW > ORDR_AMT_PRIOR OrElse (Format(ORDR_CANCEL_DATE_PRIOR, "yyyyMMdd") <> Format(ORDR_CANCEL_DATE, "yyyyMMdd")) Then
                TAC.SOCMAIN1.Credit_Request(TERM_CODE, rowSOTORDR0)

                Load_EDI_Documents()
                Load_Events()
            End If
        End If
    End Sub

    Sub Create_PO(Optional BTB As Boolean = True, Optional ORDR_NOs As List(Of String) = Nothing)
        PO_ORDER_NOs.Clear()
        Dim WHSE_CODE As String = ""
        Dim sqlORDR_NOs As String = ""

        If BTB Then
            Cancel_POs(False)
            ORDR_NOs = New List(Of String)
            ORDR_NOs.Add(ORDR_NO)
            sqlORDR_NOs = "'" & ORDR_NO & "'"
        Else
            rowSOTORDR1 = LookUp("SOTORDR1", ORDR_NOs(0))
            CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
            sqlORDR_NOs = "'" & Join(ORDR_NOs.ToArray, "','") & "'"
        End If

        WHSE_CODE = rowSOTORDR1.Item("WHSE_CODE")
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

        ASCDATA1.ExecuteSQL("Delete from " & SOTORDPX)

        If BTB Then
            ASCMAIN1.sql = "Insert into " & SOTORDPX & vbCrLf _
                & " Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.VEND_CODE from ICTSTYL1,SOTORDR2" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE and SOTORDR2.ORDR_NO = '" & ORDR_NO & "'"
            ASCDATA1.ExecuteSQL()

            Dim sqlKeepSupplierIf As String = "     and (POTORDR2.PO_QTY_SHP <> 0 or POTORDR2.PO_QTY_REC <> 0)"
            If chkKeepSupplier.Checked Then
                sqlKeepSupplierIf = "" ' maybe needs to be -> "     and (POTORDR2.PO_QTY_OPN <> 0 or POTORDR2.PO_QTY_SHP <> 0 or POTORDR2.PO_QTY_REC <> 0)"

            End If
            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & "  Select POTORDR1.VEND_CODE,POTORDR2.ORDR_NO,POTORDR2.ORDR_LNO from POTORDR1,POTORDR2 " & vbCrLf _
                & "   where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "     and POTORDR2.ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                & sqlKeepSupplierIf & ";" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDPX & vbCrLf _
                & "    Set VEND_CODE = R1.VEND_CODE where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
                & "  End Loop; " & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

        Else
            ASCMAIN1.sql = "Insert into " & SOTORDPX & vbCrLf _
                & " Select ORDR_NO, ROWNUM ORDR_LNO, STYLE_CODE, COLOR_CODE, VEND_CODE from (" & vbCrLf _
                & " Select '0000000000' ORDR_NO, 0 ORDR_LNO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.VEND_CODE from ICTSTYL1,SOTORDR2" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE and SOTORDR2.ORDR_NO in (" & sqlORDR_NOs & ")" & vbCrLf _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.VEND_CODE" & vbCrLf _
                & " order by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.VEND_CODE" & vbCrLf _
                & ")"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Select Distinct VEND_CODE from " & SOTORDPX
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim VEND_CODE As String = row.Item("VEND_CODE")
            Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
            Dim PO_ORDER_NO As String = ""
            Dim rowPOTORDR1s() As DataRow = dst.Tables("POTORDR1").Select("VEND_CODE = '" & VEND_CODE & "'")
            Dim rowPOTORDR1 As DataRow = Nothing
            If rowPOTORDR1s.Length <> 0 Then
                If rowPOTORDR1s.Length <> 1 Then
                    MsgBox("More than 1 PO defined for Vendor " & VEND_CODE, MsgBoxStyle.OkOnly, "Please Call ABS")
                    Stop
                End If
                rowPOTORDR1 = rowPOTORDR1s(0)
                PO_ORDER_NO = rowPOTORDR1.Item("PO_ORDER_NO")

                TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "AUTO", "Auto-Edit Update", "", Me.Name)

            Else
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    PO_ORDER_NO = ASCMAIN1.Next_Control_No("PO_ORDER_NO")
                Else
                    PO_ORDER_NO = ASCMAIN1.Next_Control_No("POTORDR1.PO_ORDER_NO")
                End If

                PO_ORDER_NOs.Add(PO_ORDER_NO)
                rowPOTORDR1 = dst.Tables("POTORDR1").NewRow
                With rowPOTORDR1
                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("VEND_CODE") = VEND_CODE
                    .Item("VEND_NAME") = rowAPTVEND1.Item("VEND_NAME")
                    .Item("PO_REFERENCE") = "."
                    .Item("PO_STATUS") = "O"
                    .Item("PO_XMIT_IND") = "0"
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("TERM_CODE") = rowAPTVEND1.Item("TERM_CODE")
                    .Item("PO_CARTON_MARKS") = rowARTCUST1.Item("PO_CARTON_MARKS")
                    .Item("PORT_CODE_ORIG") = rowAPTVEND1.Item("PORT_CODE")
                    .Item("COST_CODE") = rowAPTVEND1.Item("COST_CODE")
                    .Item("PO_FOB_DESC") = rowAPTVEND1.Item("VEND_PURCH_FOB_DESC")
                    .Item("PO_SHIP_VIA") = rowAPTVEND1.Item("VEND_PURCH_SHIP_VIA")
                    .Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
                    .Item("PO_FOB_DESC") = rowSOTORDR1.Item("ORDR_FOB")

                    .Item("PO_DATE_ORDERED") = rowSOTORDR1.Item("ORDR_DATE")
                    .Item("PO_DATE_SHIP_BY") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                    .Item("PO_DATE_CANCEL") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")

                    ' If rowSOTORDR1.Item("ORDR_DATE") = "COL" Then
                    '    .Item("PO_DATE_ETA") = .Item("PO_DATE_SHIP_BY")
                    ' Else
                    .Item("PO_DATE_ETA") = CDate(rowSOTORDR1.Item("ORDR_SHIP_DATE")).AddDays(Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETD_TO_ETA") & ""))
                    ' End If

                    If BTB Then
                        .Item("FOB_CMT") = "B"
                        .Item("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO")
                    Else
                        .Item("FOB_CMT") = "F"
                        .Item("ORDR_NO") = DBNull.Value
                    End If

                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        .Item("PO_COMM_PAYABLE_TO_BRKR") = "1"
                        .Item("PO_COMM_CHGBACK_TO_SUPP") = "1"
                        .Item("PO_COMM_PCT") = 2 ' BE CAREFUL - IF YOU PUT THIS INTO PARAMETER FILE IT MIGHT BE ADDED TO PO_COST
                        .Item("PO_WEB_VISIBLE") = "1"
                    End If

                End With


                dst.Tables("POTORDR1").Rows.Add(rowPOTORDR1)

                TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "AUTO", "Auto-Create", "", Me.Name)
            End If

            With rowPOTORDR1
                .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
                .Item("PO_SPEC_ORDR_NO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                .Item("PORT_CODE_DEST") = rowICTWHSE1.Item("PORT_CODE")

                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", rowSOTORDR1.Item("SHIP_VIA_CODE") & "")
                If rowSOTSVIA1 IsNot Nothing Then
                    .Item("PO_SHIP_VIA") = rowSOTSVIA1.Item("SHIP_VIA_DESC")
                End If
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
            End With

            Dim PO_ORDER_LNO As Int32 = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", "PO_ORDER_NO = '" & PO_ORDER_NO & "'") & "")

            Dim new_line_added As Boolean = False
            Dim change_made_to_qty_or_PO_Cost As Boolean = False
            Dim customer_info_changed As Boolean = False

            If BTB Then
                ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                    & " from " & SOTORDPX & " SOTORDPX,SOTORDR2" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO = SOTORDPX.ORDR_NO" & vbCrLf _
                    & "   and SOTORDR2.ORDR_LNO = SOTORDPX.ORDR_LNO" & vbCrLf _
                    & "   and SOTORDPX.VEND_CODE = '" & VEND_CODE & "'"
            Else
                ASCMAIN1.sql = "Select SOTORDPX.ORDR_NO, SOTORDPX.ORDR_LNO" & vbCrLf _
                    & ", SOTORDPX.STYLE_CODE, SOTORDPX.COLOR_CODE" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY, SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, 0 PO_COST" & vbCrLf _
                    & " from " & SOTORDPX & " SOTORDPX,SOTORDR2" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO in (" & sqlORDR_NOs & ")" & vbCrLf _
                    & "   and SOTORDR2.STYLE_CODE = SOTORDPX.STYLE_CODE" & vbCrLf _
                    & "   and SOTORDR2.COLOR_CODE = SOTORDPX.COLOR_CODE" & vbCrLf _
                    & "   and SOTORDPX.VEND_CODE = '" & VEND_CODE & "'" & vbCrLf _
                    & " group by SOTORDPX.ORDR_NO, SOTORDPX.ORDR_LNO, SOTORDPX.STYLE_CODE, SOTORDPX.COLOR_CODE"
            End If

            For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Select("", "ORDR_LNO")
                Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                Dim rowPOTORDR2s() As DataRow = dst.Tables("POTORDR2").Select("PO_ORDER_NO = '" & PO_ORDER_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO))
                Dim rowPOTORDR2 As DataRow = Nothing
                If rowPOTORDR2s.Length = 1 Then
                    rowPOTORDR2 = rowPOTORDR2s(0)
                Else
                    new_line_added = True
                    rowPOTORDR2 = dst.Tables("POTORDR2").NewRow
                    With rowPOTORDR2
                        .Item("PO_ORDER_NO") = PO_ORDER_NO
                        PO_ORDER_LNO += 1
                        .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        .Item("PO_DATE_SHIP_BY") = rowPOTORDR1.Item("PO_DATE_SHIP_BY")
                        .Item("PO_DATE_ETA") = rowPOTORDR1.Item("PO_DATE_ETA")
                        .Item("PO_ORIG_DATE_SHIP_BY") = .Item("PO_DATE_SHIP_BY")
                        .Item("PO_ORIG_DATE_ETA") = .Item("PO_DATE_ETA")

                        .Item("PO_QTY_UOM") = 1
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("PO_STATUS") = "O" ' DOES NOT ALLOW NULLS

                        If BTB Then
                            .Item("ORDR_NO") = rowSOTORDR2.Item("ORDR_NO")
                            .Item("ORDR_LNO") = ORDR_LNO
                        Else
                            .Item("ORDR_NO") = DBNull.Value
                            .Item("ORDR_LNO") = DBNull.Value
                        End If
                    End With
                    dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)
                End If

                If BTB Then
                    With rowSOTORDR2
                        Dim row_orig As DataRow = dst.Tables("SOTORDR2_ORIG").Rows.Find(New Object() {rowSOTORDR2.Item("ORDR_NO"), rowSOTORDR2.Item("ORDR_LNO")})

                        If row_orig IsNot Nothing _
                            AndAlso (
                               .Item("CUST_STYLE_CODE") & "" <> row_orig.Item("CUST_STYLE_CODE") & "" _
                            Or .Item("CUST_COLOR_CODE") & "" <> row_orig.Item("CUST_COLOR_CODE") & "" _
                            Or .Item("CUST_SIZE_CODE") & "" <> row_orig.Item("CUST_SIZE_CODE") & "" _
                            Or .Item("CUST_UPC") & "" <> row_orig.Item("CUST_UPC") & "" _
                            Or .Item("CUST_SKU") & "" <> row_orig.Item("CUST_SKU") & "" _
                            Or Val(.Item("STYLE_RETAIL") & "") <> Val(row_orig.Item("STYLE_RETAIL") & "")
                            ) Then
                            customer_info_changed = True
                        End If
                    End With

                    With rowPOTORDR2
                        If Not .RowState = DataRowState.Added Then
                            If .Item("STYLE_CODE", DataRowVersion.Original) & "" <> STYLE_CODE _
                            Or .Item("COLOR_CODE", DataRowVersion.Original) & "" <> rowSOTORDR2.Item("COLOR_CODE") & "" _
                            Or Val(.Item("PO_QTY_ORD", DataRowVersion.Original) & "") <> Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "") _
                            Or Val(.Item("PO_QTY_OPN", DataRowVersion.Original) & "") <> Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "") _
                            Or Val(.Item("PO_COST", DataRowVersion.Original) & "") <> Val(rowSOTORDR2.Item("PO_COST") & "") _
                            Or Val(.Item("PO_COST_VCOST", DataRowVersion.Original) & "") <> Val(rowSOTORDR2.Item("PO_COST") & "") _
                            Or Val(.Item("CARTON_PACK_QTY", DataRowVersion.Original) & "") <> Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "") _
                            Then
                                change_made_to_qty_or_PO_Cost = True
                            End If
                        End If
                    End With
                End If

                With rowPOTORDR2
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = rowSOTORDR2.Item("COLOR_CODE")
                    .Item("PO_QTY_ORD") = rowSOTORDR2.Item("ORDR_QTY")
                    .Item("PO_QTY_OPN") = Val(.Item("PO_QTY_ORD") & "") - Val(.Item("PO_QTY_SHP") & "")
                    If Val(.Item("PO_QTY_OPN") & "") <= 0 Then
                        .Item("PO_QTY_OPN") = 0
                        .Item("PO_STATUS") = "C"
                    Else
                        .Item("PO_STATUS") = "O"
                    End If

                    Dim PO_COST As Decimal
                    If BTB Then
                        PO_COST = rowSOTORDR2.Item("PO_COST")
                    Else
                        PO_COST = 0
                        Dim rowICTSTYV1 As DataRow = LookUp("ICTSTYV1", New String() {STYLE_CODE, VEND_CODE})
                        If rowICTSTYV1 IsNot Nothing Then
                            If rowICTSTYV1.Item("NEW_PO_COST_DATE") & "" <> "" AndAlso Format(rowPOTORDR1.Item("PO_DATE_ORDERED"), "yyyyMMdd") >= Format(rowICTSTYV1.Item("NEW_PO_COST_DATE"), "yyyyMMdd") Then
                                PO_COST = Val(rowICTSTYV1.Item("NEW_PO_COST") & "")
                            Else
                                PO_COST = Val(rowICTSTYV1.Item("PO_COST") & "")
                            End If
                        End If
                    End If

                    Dim PO_COST_OTHER As Decimal = Val(.Item("PO_COST_OTHER") & "")

                    .Item("PO_COST") = PO_COST + PO_COST_OTHER
                    .Item("PO_COST_VCOST") = PO_COST

                    .Item("SUB_UNIT_PACK_QTY") = rowICTSTYL1.Item("SUB_UNIT_PACK_QTY")
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("PO_COST_VCOST_DZ") = PO_COST * 12
                    .Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
                    .Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
                End With
            Next

            If Not rowPOTORDR1.RowState = DataRowState.Added Then

                Dim PO_QTY_OPN_orig As Int64 = Val(ASCDATA1.GetDataValue _
                    ("Select Sum (PO_QTY_OPN) from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"))

                Dim PO_QTY_OPN_curr As Int64 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", "PO_ORDER_NO = '" & PO_ORDER_NO & "'") & "")

                If new_line_added Or change_made_to_qty_or_PO_Cost Or customer_info_changed Or PO_QTY_OPN_orig <> PO_QTY_OPN_curr Then
                    If rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1" Then
                        Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
                        rowPOTORDR1.Item("PO_PRINTED_IND") = "0"
                        rowPOTORDR1.Item("PO_XMIT_IND") = "0"
                        PO_HDR_CTR_REV += 1
                        rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV
                        rowPOTORDR1.Item("PO_REVISION_NOTE") = "Modified in Sales Order Entry"
                        MsgBox("PO Revision Counter has been Updated for PO " & PO_ORDER_NO _
                               & vbCrLf & vbCrLf & "Re-Transmission will be Required.", MsgBoxStyle.OkOnly, "Verification")
                    End If

                    If PO_QTY_OPN_curr <> 0 And rowPOTORDR1.Item("PO_STATUS") & "" <> "O" Then
                        rowPOTORDR1.Item("PO_STATUS") = "O"

                        TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID,
                                                  "REO-PO",
                                                  "PO Re-Opened from BTB Update",
                                                  ORDR_NO, "POFORDR1")

                    End If
                End If
            End If
        Next

        If BTB Then
            Update_POs()
        End If

    End Sub

    Sub Update_POs(Optional BTB As Boolean = True)

        Update_Record_TDA("POTORDR1")
        Update_Record_TDA("POTORDR2")

        If BTB Then
            ASCMAIN1.sql = "Delete from POTORDR2" & vbCrLf _
                & " where ORDR_NO = '" & ORDR_NO & "'" _
                & "   and NVL(PO_QTY_OPN,0) = 0" & vbCrLf _
                & "   and NVL(PO_QTY_SHP,0) = 0" & vbCrLf _
                & "   and NVL(PO_QTY_REC,0) = 0"
            ASCDATA1.ExecuteSQL()

        End If

        For Each PO_ORDER_NO As String In PO_ORDER_NOs
            Dim cancelPO As Boolean = False
            If BTB Then

                ASCMAIN1.sql = "Delete from POTORDR6 where (PO_ORDER_NO, PO_ORDER_LNO) in (" & vbCrLf _
                    & "Select Distinct PO_ORDER_NO, PO_ORDER_LNO from POTORDR6 where PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                    & " minus" & vbCrLf _
                    & "Select Distinct PO_ORDER_NO, PO_ORDER_LNO from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "')"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Select Count (*) LINES" & vbCrLf _
                    & ", Sum (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
                    & ", Sum (PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
                    & ", Sum (PO_QTY_REC) PO_QTY_REC from POTORDR2" & vbCrLf _
                    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                    & "   and (NVL(PO_QTY_OPN,0) <> 0 or NVL(PO_QTY_SHP,0) <> 0 or NVL(PO_QTY_REC,0) <> 0)"
                Dim row As DataRow = ASCDATA1.GetDataRow
                Dim LINES As Int64 = Val(row.Item("LINES") & "")
                Dim PO_QTY_OPN As Int64 = Val(row.Item("PO_QTY_OPN") & "")
                Dim PO_QTY_SHP As Int64 = Val(row.Item("PO_QTY_SHP") & "")
                Dim PO_QTY_REC As Int64 = Val(row.Item("PO_QTY_REC") & "")
                Dim row1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                Dim PO_STATUS As String = row1.Item("PO_STATUS") & ""
                If LINES = 0 And PO_STATUS = "O" Then
                    cancelPO = True

                    TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID,
                                              "CXL-PO",
                                              "PO Cancelled from BTB Update",
                                              ORDR_NO, "POFORDR1")
                    'Else
                    '    If PO_QTY_OPN <> 0 And PO_STATUS = "C" Then
                    '        ASCMAIN1.sql = "Update POTORDR1 Set PO_STATUS = 'O', PO_DATE_CANCELLED = NULL" & vbCrLf _
                    '            & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                    '        ASCDATA1.ExecuteSQL()
                    '    End If
                End If
            End If
            TAC.POCMAIN1.Dependent_Updates(1, PO_ORDER_NO, cancelPO)
        Next
    End Sub

    Sub Cancel_POs(cancel_PO_and_write_POs_to_database As Boolean)

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
            Dim PO_ORDER_NO As String = rowPOTORDR2.Item("PO_ORDER_NO")
            If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then
                PO_ORDER_NOs.Add(PO_ORDER_NO)
                If cancel_PO_and_write_POs_to_database Then
                    Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(PO_ORDER_NO)
                    rowPOTORDR1.Item("PO_STATUS") = "C"
                End If
            End If
            rowPOTORDR2.Item("PO_QTY_OPN") = 0
            rowPOTORDR2.Item("PO_STATUS") = "C"
        Next

        For Each PO_ORDER_NO As String In PO_ORDER_NOs
            TAC.POCMAIN1.Dependent_Updates(-1, PO_ORDER_NO, cancel_PO_and_write_POs_to_database)
        Next

        If cancel_PO_and_write_POs_to_database Then
            Update_Record_TDA("POTORDR1")
            Update_Record_TDA("POTORDR2")
        End If
    End Sub

    Sub Copy_Records(ORDR_NOs_to_copy_to As List(Of String))

        For Each TABLE_NAME As String In New String() {"SOTORDR4", "SOTORDR5", "SOTORDR9"}
            '  Sql = "Select * from " & z & " where ORDR_NO = '" & Mid$(ALL_ORDERS, 1, 10) & "'"

            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "'"
            If TABLE_NAME = "SOTORDR5" Then sqlw &= " and CUST_ADDR_TYPE = 'BT'"
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlw)
                For Each ORDR_NO_to_copy_to As String In ORDR_NOs_to_copy_to
                    Dim row2 As DataRow = dst.Tables(TABLE_NAME).NewRow
                    row2.ItemArray = row.ItemArray
                    row2.Item("ORDR_NO") = ORDR_NO_to_copy_to
                    If dst.Tables(TABLE_NAME).Rows.Find(New String() {row2.Item("ORDR_NO"), row2.Item("CUST_ADDR_TYPE")}) Is Nothing Then
                        dst.Tables(TABLE_NAME).Rows.Add(row2)
                    End If
                Next
            Next
        Next
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "CUST_STORE_NO"
                sql_where = "CUST_ADDR_TYPE = 'MK' and CUST_ADDR_STATUS = 'A'"

            Case "SOTORDR5_BT.CUST_ADDR_CODE"
                If ASCMAIN1.Running_in_VS Then Stop
                Stop
                ' Sql = "Select CUST_ADDR_CODE, CUST_NAME FROM ARTCUST2 WHERE CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                ' Sql = Sql & " AND CUST_ADDR_TYPE = 'BT'"


            Case "ORDR_NO", "INV_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                If Absx1.txtFor("CUST_STORE_NO").Text <> "" And Absx1.txtFor("CUST_CODE").Text = "" Then
                    MsgBox("You must enter a Customer Code", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""


                If COLUMN_NAME = "ORDR_NO" Then
                    If InquiryMode Then
                    Else
                        sql_where &= " and SOTORDR1.ORDR_STATUS = 'O' "
                    End If
                End If
                If COLUMN_NAME = "INV_NO" Then
                    sql_where &= " and SOTORDR1.ORDR_STATUS = 'F' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("CUST_STORE_NO").Text <> "" Then
                    sql_where &= " and SOTORDR1.CUST_STORE_NO = '" & Absx1.txtFor("CUST_STORE_NO").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTORDR1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If

            Case "SHIP_VIA_CODE"
                If Not InquiryMode Then
                    If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                        sql_where &= "NVL(SHIP_VIA_STATUS,'A') = 'A'"
                    End If
                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        sql_where &= "SHIP_VIA_STATUS = 'A'"

                        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

                        If rowICTWHSE1.Item("WHSE_EDI_ID") & "" <> "" Then
                            sql_where = " and (SHIP_VIA_CODE = 'ROUT' or SHIP_VIA_CODE in (" & vbCrLf _
                                & "Select EDTXREF3.SHIP_VIA_CODE from EDTXREF3,ICTWHSE1" & vbCrLf _
                                & " where EDTXREF3.SENDER_ID_QUAL = ICTWHSE1.WHSE_EDI_QUAL" & vbCrLf _
                                & "   and EDTXREF3.SENDER_ID = ICTWHSE1.WHSE_EDI_ID" & vbCrLf _
                                & "   and ICTWHSE1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                                & "))"
                        End If
                    End If
                End If

            Case "SREP_CODE"
                sql_where &= " and SOTSREP1.SREP_STATUS = 'A'"

            Case "SREP2_CODE"
                sql_where &= " and SOTSREP1.SREP_STATUS = 'A'"
        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("ORDR_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTORDR1"
            E.COLUMN_NAME = "ORDR_NO"
            E.CODE_VALUE = Absx1.txtFor("ORDR_NO").Text
            E.DESC_VALUE = "Sales Order"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDRX, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Refresh", "Create POs", "Copy Order", "Customer Order Status")
        Load_Popup_Menu(grdSOTORDR2, "BBBBSBBSBBBB", "Style Status Inquiry", "Style Master File", "Get PO Cost if 0", "Style Multi-Color", "Show UPC/SKU", "Copy from Reservation", "Sub Style", "Show Disc/Comm", "Clone Line", "Group as Pre-Pack", "Customer Order Status", "Import Details From Excel", "Show Import Template")
        Load_Popup_Menu(grdSOTORDR3, "B", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTORDRS, "BB", "Set Customer PO to Value in Header", "Update Qty to All Stores")
        Load_Popup_Menu(grdSOTORDXR, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTORDRB, "S", "Show Additional Header Fields")
        Load_Popup_Menu(grdSOTCART1, "B", "Track Shipment")
        Load_Popup_Menu(grdSOTPICK1, "BBBB", "Show Invoice", "email Invoice", "Track Shipment", "Show Invoice as Report", "Show Raw EDI 940")
        Load_Popup_Menu(grdSOTORDP1, "BBB", "Generate Pro-Forma Invoice", "Show Pro-Forma Invoice", "email Pro-Forma Invoice")
        Load_Popup_Menu(grdTATEVNT1, "BB", "Show email", "Show Document")
        Load_Popup_Menu(grdSOTORDRI, "BBBBBB", "Style Status Inquiry", "Style Master File", "Sub Style", "Cancel Qty for All Stores", "Reset Qty to Original", "Clone Line")
        Load_Popup_Menu(grdPOTORDR1, "B", "PO Inquiry")
        Load_Popup_Menu(grdSOTINVH1, "BBB", "Show Invoice", "email Invoice", "Show Invoice as Report")
        Load_Popup_Menu(grdEDTDOCS1, "B", "Show Document")

        Check_InquiryMode()
        ' If InquiryMode Then
        Load_Popup_Menu(grdSOTORDC1, "BBB", "CC Deposit", "Add On Account", "Additional Funds", "De-Activate", "CC Authorization", "Void Authorization", "Charge Against Auth")
        ' End If

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If
        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

            Case "grdSOTORDRX"
                tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                tlb_btn = DirectCast(tlb_pop.Tools("Create POs"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And (grdSOTORDRX.ActiveRow IsNot Nothing Or grdSOTORDRX.Selected.Rows.Count <> 0) And (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
                tlb_btn = DirectCast(tlb_pop.Tools("Copy Order"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And (grdSOTORDRX.ActiveRow IsNot Nothing Or grdSOTORDRX.Selected.Rows.Count <> 0) And Not InquiryMode And (ASCMAIN1.CLIENT = "VAN")
                tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Status"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And (grdSOTORDRX.ActiveRow IsNot Nothing Or grdSOTORDRX.Selected.Rows.Count <> 0) And (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")


            Case "grdSOTORDRI"
                tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                tlb_btn = DirectCast(tlb_pop.Tools("Cancel Qty for All Stores"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And grdSOTORDRI.ActiveRow IsNot Nothing
                tlb_btn = DirectCast(tlb_pop.Tools("Reset Qty to Original"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And grdSOTORDRI.ActiveRow IsNot Nothing

            Case "grdSOTORDP1"
                tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

                tlb_btn = DirectCast(tlb_pop.Tools("Generate Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E")

                tlb_btn = DirectCast(tlb_pop.Tools("Show Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") Or (EntryMode = "")

                tlb_btn = DirectCast(tlb_pop.Tools("email Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V")

            Case "grdTATEVNT1"
                tlb_btn = DirectCast(tlb_pop.Tools("Show email"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("EVENT_TYPE").Value = "EML" Or grd.ActiveRow.Cells("EVENT_TYPE").Value = "CORSTA" Or grd.ActiveRow.Cells("EVENT_TYPE").Value = "VASEML"))

                tlb_btn = DirectCast(tlb_pop.Tools("Show Document"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("EVENT_TYPE").Value = "RELORD" Or grd.ActiveRow.Cells("EVENT_TYPE").Value = "PF" Or grd.ActiveRow.Cells("EVENT_TYPE").Value = "RELSHP"))


            Case "grdSOTORDC1"
                '"CC Deposit", "Add On Account", "Additional Funds", "De-Activate", "CC Authorization", "Charge Against Auth"

                ' get updated order data.
                Dim row As DataRow = ASCDATA1.GetDataRow("Select * from SOTORDR1 where ORDR_NO = :PARM1", "V", New Object() {ORDR_NO})
                Dim orderInPickOrOpen As Boolean = row.Item("ORDR_STATUS") = "O" OrElse row.Item("ORDR_STATUS") = "P"
                Dim hasCreditCardCharge As Boolean = (row.Item("CCPA_NO") & String.Empty).ToString.Trim.Length > 0
                Dim hasPickTickets As Boolean = ASCDATA1.GetDataTable("Select * from SOTPICK1 where ORDR_NO = :PARM1 AND PICK_STATUS = 'P'", "", "V", New Object() {ORDR_NO}).Rows.Count > 0


                tlb_btn = DirectCast(tlb_pop.Tools("De-Activate"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") AndAlso orderInPickOrOpen AndAlso hasCreditCardCharge _
                        AndAlso (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "SOTORDC1" AndAlso grd.ActiveRow.Cells("ACTIVE_IND").Value = "1")

                tlb_btn = DirectCast(tlb_pop.Tools("CC Deposit"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") AndAlso orderInPickOrOpen AndAlso hasCreditCardCharge

                tlb_btn = DirectCast(tlb_pop.Tools("Add On Account"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") AndAlso orderInPickOrOpen AndAlso hasCreditCardCharge

                tlb_btn = DirectCast(tlb_pop.Tools("Additional Funds"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") AndAlso orderInPickOrOpen AndAlso hasCreditCardCharge

                tlb_btn = DirectCast(tlb_pop.Tools("CC Authorization"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") AndAlso orderInPickOrOpen 'AndAlso hasCreditCardCharge
                '   tlb_btn.SharedProps.Visible = (InquiryMode Or EntryMode = "E") AndAlso orderInPickOrOpen 'AndAlso hasCreditCardCharge

                '"Void Authorization"
                tlb_btn = DirectCast(tlb_pop.Tools("Void Authorization"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") AndAlso Not hasPickTickets _
                        AndAlso (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "SOTORDC1" AndAlso grd.ActiveRow.Cells("ACTIVE_IND").Value = "1")

                '"Charge Against Auth"
                tlb_btn = DirectCast(tlb_pop.Tools("Charge Against Auth"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "V") AndAlso 1 = 2 _
                    AndAlso hasCreditCardCharge _
                    AndAlso grd.ActiveRow.Band.Key = "SOTORDC1" _
                    AndAlso (grd.ActiveRow.Cells("CCPA_STATUS").Value & String.Empty).ToString.Trim = "T" _
                    AndAlso (grd.ActiveRow.Cells("ACTIVE_IND").Value & String.Empty).ToString.Trim = "1"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
            ' NEED TO GET PAST HERE FOR STYLE MULTI-COLOR WHEN THERE ARE NO ROWS IN THE GRID
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdSOTORDRS"
                    Dim show_qty_copy_option As Boolean = False
                    Dim ORDR_QTY As Int64 = 0
                    If grdSOTORDRS.ActiveCell IsNot Nothing Then
                        Dim COLUMN_NAME As String = grdSOTORDRS.ActiveCell.Column.Key
                        If COLUMN_NAME.StartsWith("QTY_") Then
                            show_qty_copy_option = True
                            ORDR_QTY = Val(grdSOTORDRS.ActiveCell.Value & "")
                        End If
                    End If
                    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = show_qty_copy_option
                    tlb_btn.SharedProps.Caption = "Update Qty to " & CStr(ORDR_QTY) & " for All Stores"

                Case "grdSOTORDR2"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_sbt = DirectCast(tlb_pop.Tools("Show UPC/SKU"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.SharedProps.Visible = True ' (Absx1.optFor("ORDR_SOURCE").Value = "K")
                    tlb_sbt.Tag = "X"
                    tlb_sbt.Checked = Not grdSOTORDR2.DisplayLayout.Bands(0).Columns("CUST_UPC").Hidden
                    tlb_sbt.Tag = ""

                    tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                    tlb_btn = DirectCast(tlb_pop.Tools("Copy from Reservation"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And Not (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")

                    tlb_btn = DirectCast(tlb_pop.Tools("Sub Style"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E") And (ROWs("SOTPARM1").Item("SO_PARM_SUB_STYLES") & "" = "1")

                    tlb_btn = DirectCast(tlb_pop.Tools("Get PO Cost if 0"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E") And (Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB")

                    tlb_sbt = DirectCast(tlb_pop.Tools("Show Disc/Comm"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.SharedProps.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")
                    tlb_sbt.Tag = "X"
                    tlb_sbt.Checked = Not grdSOTORDR2.DisplayLayout.Bands(0).Columns("DISC_AMT").Hidden
                    tlb_sbt.Tag = ""

                    tlb_btn = DirectCast(tlb_pop.Tools("Clone Line"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") AndAlso rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"

                    tlb_btn = DirectCast(tlb_pop.Tools("Group as Pre-Pack"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") AndAlso rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG" AndAlso ASCMAIN1.Running_in_VS

                    tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Status"), UltraWinToolbars.ButtonTool)
                    tlb_sbt.SharedProps.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")

                    tlb_btn = DirectCast(tlb_pop.Tools("Import Details From Excel"), UltraWinToolbars.ButtonTool)
                    tlb_sbt.SharedProps.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") And (ASCMAIN1.Running_in_VS)

                    tlb_btn = DirectCast(tlb_pop.Tools("Show Import Template"), UltraWinToolbars.ButtonTool)
                    tlb_sbt.SharedProps.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") And (ASCMAIN1.Running_in_VS)

                Case "grdSOTORDR3"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_btn = DirectCast(tlb_pop.Tools("Add Sizes"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Case "grdSOTORDRI"
                    tlb_btn = DirectCast(tlb_pop.Tools("Clone Line"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") AndAlso rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Import Details From Excel"
                ImportDetailsFromExcel()
            Case "Show Import Template"
                Dim FName As String = "SOUpload.xlsx"
                Dim FLDName As String = "templates"
                Dim ROOTName As String = ASCMAIN1.Folders("Archive")
                If Not ROOTName.EndsWith("\") Then
                    ROOTName = ROOTName & "\"
                End If
                Dim FILENAME As String = $"{ROOTName}{FLDName}\{FName}"
                Show_Document(FILENAME)
            Case "Set Customer PO to Value in Header"
                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                    rowSOTORDRS.Item("ORDR_CUST_PO") = Absx1.txtFor("ORDR_CUST_PO").Text
                Next

            Case "Show UPC/SKU"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    Toggle_Customer_Style_Fields(tlb_sbt.Checked)
                End If

            Case "Show Disc/Comm"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    Toggle_Disc_Comm_Fields(tlb_sbt.Checked)
                End If

            Case "Style Multi-Color"
                Using F As New TAC.ICFSTYCX
                    F.STYLE_CODE = ""
                    F.WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    F.Price_Caption = "Unit Price"
                    F.ShowDialog()
                    If F.STYLE_CODE <> "" Then
                        Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
                    End If
                End Using

            Case "Copy from Reservation"
                Dim sql_where As String = ""
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("RSRV_NO", , sql_where)

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreFilter.Add("CUST_CODE", CUST_CODE)
                    ASCMAIN1.CodeSelector.PreFilter.Add("RSRV_STATUS", "O")
                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Loading")
                        Dim RSRV_NO As String = ASCMAIN1.CodeSelector.SelectedCode

                        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, RSRV_QTY, ORDR_UNIT_PRICE from SOTRSRV2 where RSRV_NO = '" & RSRV_NO & "'"
                        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                            Add_grdSOTORDR2(row.Item("STYLE_CODE"), row.Item("COLOR_CODE"), row.Item("RSRV_QTY"), row.Item("ORDR_UNIT_PRICE"))
                        Next

                        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                    End If
                End If

            Case "Add Sizes"
                cmdSizes_Click()

            Case "Refresh"
                Load_SOTORDRX()

            Case "Show Additional Header Fields"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Setup_Multiple_Order_Grid(tlb_sbt.Checked)

            Case "Group as Pre-Pack"
                ' NOTE THIS IS ONLY AVAILABLE WHEN RUNNING FROM VS - WJZ ONLY - TESTING MANUALLY ENTERED ORDERS 
                ' INTENDED FOR NON-EDI ORDERS ONLY

                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso Not grd.ActiveRow.IsAddRow AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.Selected.Rows.Count < 2 Then
                    MsgBox("You must select at least 2 styles that you want to Group as a Pre-Pack" & vbCrLf & " prior to selecting this option", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                Dim STYLE_CODE_ppk As String = ""
                Dim QTY_TOTAL As Int64 = 0
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If STYLE_CODE_ppk = "" Then
                        STYLE_CODE_ppk = grow.Cells("STYLE_CODE").Value
                    Else
                        If STYLE_CODE_ppk <> grow.Cells("STYLE_CODE").Value Then
                            MsgBox("Cannot group styles on order lines with Different Styles",
                                                         MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                            Exit Sub
                        End If
                    End If

                    If Val(grow.Cells("ORDR_QTY_PICK").Value & "") <> 0 Or
                         Val(grow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 Then
                        MsgBox("Cannot group styles on order lines that have been released or shipped",
                               MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Exit Sub
                    End If

                    QTY_TOTAL += Val(grow.Cells("ORDR_QTY_OPEN").Value & "")
                Next

                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_ppk)
                Dim INNER_PACK_QTY As Int64 = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")

                If INNER_PACK_QTY = 0 Then
                    MsgBox("Pre-Pack Styles Must have a Non-Zero Value for Inner Pack (in the Style Master)",
                        MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                Else
                    If QTY_TOTAL Mod INNER_PACK_QTY <> 0 Then
                        MsgBox("Total Open Qty (" & CStr(QTY_TOTAL) & " is not evenly Divisible by the Style's Inner Pack Qty (" & INNER_PACK_QTY & ")",
                            MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Exit Sub
                    End If
                End If

            Case "Sub Style"

                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso Not grd.ActiveRow.IsAddRow AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("You must select the styles that you want to sub" & vbCrLf & " prior to selecting this option", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                Dim COLOR_CODE As String = ""
                Dim COLOR_CODEs As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If Val(grow.Cells("ORDR_QTY_PICK").Value & "") <> 0 Or
                         Val(grow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 Then
                        MsgBox("Cannot sub styles on order lines that have been released or shipped",
                               MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Exit Sub
                    Else
                        If COLOR_CODE = "" Then COLOR_CODE = grow.Cells("COLOR_CODE").Value
                        If Not COLOR_CODEs.Contains(grow.Cells("COLOR_CODE").Value) Then
                            COLOR_CODEs.Add(grow.Cells("COLOR_CODE").Value)
                        End If
                    End If
                Next

                Dim STYLE_CODE_SUB As String = Select_Style(COLOR_CODE)
                If STYLE_CODE_SUB <> "" Then
                    ' SHOULD TEST THIS NEW STYLE AGAINST ALL COLORS IN COLOR_CODEs
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_SUB)
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        If grow.Cells("STYLE_CODE_SUB").Value & "" = "" Then
                            grow.Cells("STYLE_CODE_SUB").Value = grow.Cells("STYLE_CODE").Value
                        ElseIf grow.Cells("STYLE_CODE_SUB").Value & "" = STYLE_CODE_SUB Then
                            grow.Cells("STYLE_CODE_SUB").Value = ""
                        End If
                        grow.Cells("STYLE_CODE").Value = STYLE_CODE_SUB
                        '  grow.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        ' ROBIN/JOHN/IAN WANTS THE ORIGINAL STYLE DESC TO PERSIST - RECALL 812172 ON 12/04/13

                        If grd.Name = "grdSOTORDRI" Then
                            Dim COL As Integer = Val(grow.Cells("COL").Value & "")
                            Dim CN As String = "QTY_" & Format(COL, "000")
                            With grdSOTORDRB.DisplayLayout.Bands(0).Columns(CN).Header
                                .Caption = STYLE_CODE_SUB & ":" & COLOR_CODE
                                .ToolTipText = STYLE_CODE_SUB & ":" & COLOR_CODE & vbCrLf & rowICTSTYL1.Item("STYLE_DESC")
                            End With
                        End If
                        grow.Update()
                    Next


                End If

            Case "Multiple Order Maintenance"
                If grdSOTORDRX.Selected.Rows.Count <> 0 Then
                    Dim ORDR_SHIP_DATE As Date = Nothing
                    Dim ORDR_CANCEL_DATE As Date = Nothing
                    Dim ORDR_HOLD As String = ""
                    ORDR_NOs_to_maintain.Clear()
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRX.Selected.Rows
                        Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then
                            ORDR_NOs_to_maintain.Clear()
                            Exit For
                        Else
                            ORDR_NOs_to_maintain.Add(ORDR_NO)
                        End If
                        If grow.Cells("CUST_CODE").Value <> Absx1.txtFor("CUST_CODE").Text Then
                            MsgBox("Customer Code of each order selected" & vbCrLf & " must be the same as Customer Entered Above (" & Absx1.txtFor("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Proceed with Multiple Order Maintenance")
                            ORDR_NOs_to_maintain.Clear()
                            Exit For
                        End If
                        If ORDR_HOLD = "" Then
                            ORDR_HOLD = IIf(grow.Cells("ORDR_HOLD").Value & "" = "1", "1", "0")
                            ORDR_SHIP_DATE = grow.Cells("ORDR_SHIP_DATE").Value
                            ORDR_CANCEL_DATE = grow.Cells("ORDR_CANCEL_DATE").Value
                        Else
                            If ORDR_HOLD <> IIf(grow.Cells("ORDR_HOLD").Value & "" = "1", "1", "0") _
                            Or ORDR_SHIP_DATE <> grow.Cells("ORDR_SHIP_DATE").Value _
                            Or ORDR_CANCEL_DATE <> grow.Cells("ORDR_CANCEL_DATE").Value Then
                                MsgBox("On Hold Status and Shipping Window of all orders selected must be the same", MsgBoxStyle.OkOnly, "Cannot Proceed with Multiple Order Maintenance")
                                ORDR_NOs_to_maintain.Clear()
                                Exit For
                            End If
                        End If
                    Next

                    If ORDR_NOs_to_maintain.Count = 0 Then
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    Else
                        Absx1.txtFor("ORDR_NO").Text = ORDR_NOs_to_maintain(0)
                        Click_Command("Edit")
                    End If
                End If

            Case "Create POs"
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    If grdSOTORDRX.ActiveRow IsNot Nothing Then
                        grdSOTORDRX.ActiveRow.Selected = True
                    End If
                End If
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    MsgBox("No Sales Orders Selected", MsgBoxStyle.OkOnly, "Cannot Create POs until you Select Sales Orders")
                Else
                    Dim ORDR_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRX.Selected.Rows
                        Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
                        ORDR_NOs.Add(ORDR_NO)
                    Next

                    If MsgBox("You have selected " & ORDR_NOs.Count & " Sales Orders." _
                              & vbCrLf & vbCrLf & "Continue to Generate POs to Cover the Demand" _
                              & vbCrLf & " represented by these Sales Orders?" _
                              & vbCrLf & vbCrLf & "(You will be given a chance to Review before Actually Generating POs)",
                              MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Create_POs(ORDR_NOs)
                    End If
                End If
            Case "Customer Order Status"
                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                Dim ORDR_GROUP_NOs_to_batch As New List(Of String)
                '   For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_NO").Text
                    ORDR_GROUP_NOs_to_batch.Add(ORDR_GROUP_NO)
                Next

                Using FRM As New SOFCORS1
                    FRM.ORDR_GROUP_NOs = ORDR_GROUP_NOs_to_batch
                    FRM.CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    FRM.ShowDialog()
                End Using
                If grd.Name = "grdSOTORDR2" Then
                    Fill_Records("SOTORDR4", ORDR_NO)
                End If

               ' grdSOTORDR0.Selected.Rows.Clear()
            Case "Get PO Cost if 0"
                Get_PO_Cost()

            Case "Generate Pro-Forma Invoice"
                Generate_Pro_Forma_Invoice()

            Case "CC Authorization", "CC Deposit", "Add On Account", "Additional Funds", "De-Activate", "Charge Against Auth", "Void Authorization"
                Dim trans_no As String = String.Empty
                If grd.ActiveRow IsNot Nothing Then
                    trans_no = grd.ActiveRow.Cells("TRANS_NO").Value
                End If
                ProcessCreditCardDeposit(e.Tool.Key, trans_no)

                EnforceConstraints(False)
                ASCMAIN1.sql = "Select SOTORDC1.*, ARTCCPA1.CUST_CREDIT_CARD_LAST4, ARTCCPA1.CCPA_DATE_VOID" _
                 & " from SOTORDC1, ARTCCPA1 " _
                 & " where SOTORDC1.ccpa_no = ARTCCPA1.ccpa_no (+)" _
                 & " and SOTORDC1.ORDR_NO = '" & ORDR_NO & "'"
                Fill_Records("SOTORDC1", String.Empty, True, ASCMAIN1.sql)
                Fill_Records("SOTORDC2", Absx1.txtFor("ORDR_NO").Text)
                EnforceConstraints(True)
                Exit Sub

            Case "Cancel Qty for All Stores"

                Dim COL As Integer = grd.ActiveRow.Cells("COL").Value
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim ORDR_QTY_OPEN As Integer = Val(grd.ActiveRow.Cells("ORDR_QTY_OPEN").Value & "")

                If MsgBox("Are you sure that you want to Cancel " & CStr(ORDR_QTY_OPEN) & " units of Style " & STYLE_CODE & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Dim CN As String = "QTY_" & Format(COL, "000")
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRB.Rows
                        grow.Cells(CN).Value = 0
                        grow.Update()
                    Next
                    grd.ActiveRow.Cells("ORDR_QTY_OPEN").Value = 0
                    grd.ActiveRow.Update()
                    MsgBox("Style " & STYLE_CODE & " has been Cancelled")
                End If

            Case "Reset Qty to Original"

                Dim COL As Integer = grd.ActiveRow.Cells("COL").Value
                Dim STYLE_KEY As String = grd.ActiveRow.Cells("STYLE_KEY").Value
                ' Dim ORDR_LNO As Integer = dst.Tables("SOTORDRQ_KEY").Rows.Find(New Object() {"", ""})(0) '  grd.ActiveRow.Cells("ORDR_LNO").Value
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                '   Dim ORDR_QTY_OPEN As Integer = Val(grd.ActiveRow.Cells("ORDR_QTY_OPEN").Value & "")

                If MsgBox("Are you sure that you want to Reset Style " & STYLE_KEY & " to Original Qtys?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Dim CN As String = "QTY_" & Format(COL, "000")
                    Dim Q As Int64 = 0
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRB.Rows
                        Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
                        'Dim rowSOTORDRQ As DataRow = dst.Tables("SOTORDRQ").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                        Dim rowSOTORDRQs() As DataRow = dst.Tables("SOTORDRQ").Select("ORDR_NO = '" & ORDR_NO & "' and STYLE_KEY = '" & STYLE_KEY & "'")
                        If rowSOTORDRQs.Length = 1 Then
                            grow.Cells(CN).Value = rowSOTORDRQs(0).Item("ORDR_QTY")
                            Q += Val(rowSOTORDRQs(0).Item("ORDR_QTY") & "")
                            grow.Update()
                        End If
                    Next
                    grd.ActiveRow.Cells("ORDR_QTY_OPEN").Value = Q
                    grd.ActiveRow.Update()
                    MsgBox("Style " & STYLE_KEY & " has been Reset to Original Order Qtys")
                End If

            Case "Copy Order"
                If grdSOTORDRX.ActiveRow Is Nothing Then Exit Sub

                Dim ORDR_NO_to_copy_maybe As String = grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO_to_copy_maybe)
                If rowSOTORDR1 Is Nothing Then
                    Exit Sub
                Else
                    If Absx1.txtFor("CUST_CODE").Text <> rowSOTORDR1.Item("CUST_CODE") Then
                        MsgBox("Start the order by entering the Customer and Store", MsgBoxStyle.OkOnly, "Cannot Proceed")
                        Exit Sub
                    End If

                    ORDR_NO_to_copy = ORDR_NO_to_copy_maybe
                    If Absx1.txtFor("CUST_STORE_NO").Text = "" Then
                        Absx1.txtFor("CUST_STORE_NO").Text = rowSOTORDR1.Item("CUST_STORE_NO")
                    End If
                    Absx1.txtFor("ORDR_CUST_PO").Text = rowSOTORDR1.Item("ORDR_CUST_PO") & ""
                    Absx1.optFor("ORDR_TYPE_CODE").Value = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                    Click_Command("New")

                    Absx1.txtFor("ORDR_DEPT").Text = rowSOTORDR1.Item("ORDR_DEPT") & ""
                    '       Absx1.txtFor("ORDR_FOB").Text = rowSOTORDR1.Item("ORDR_FOB") & ""
                    Absx1.dteFor("ORDR_SHIP_DATE").Value = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                    Absx1.dteFor("ORDR_CANCEL_DATE").Value = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
                    Absx1.dteFor("ORDR_ARRIVAL_DATE").Value = rowSOTORDR1.Item("ORDR_ARRIVAL_DATE")
                    Absx1.txtFor("SALES_DIVISION_CODE").Text = rowSOTORDR1.Item("SALES_DIVISION_CODE") & ""

                    Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, "BT"})
                    rowSOTORDR5.Item("CUST_ADDR_CODE") = "000000"

                    Record_Event("COPY", "Order Copied from " & ORDR_NO_to_copy)
                    ORDR_NO_to_copy = ""
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Show Document"
                If grd.Name = "grdEDTDOCS1" Then
                    If grdEDTDOCS1.ActiveRow IsNot Nothing Then
                        Dim EDI_DOC_SEQ_NO As String = grdEDTDOCS1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                        Dim EDI_DOC_ID As String = grdEDTDOCS1.ActiveRow.Cells("EDI_DOC_ID").Value & ""
                        Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), EDI_DOC_ID, ORDR_GROUP_NO)
                        Using frm As New ASFTEXT1
                            frm.t = RAW_EDI
                            frm.Text = "Raw EDI for " & EDI_DOC_ID
                            frm.ShowDialog()
                        End Using
                    End If
                ElseIf grd.Name = "grdTATEVNT1" Then

                    Dim EVENT_TYPE As String = grd.ActiveRow.Cells("EVENT_TYPE").Value & ""

                    If EVENT_TYPE = "RELORD" Then

                        Dim EDI_DOC_SEQ_NO As String = grdTATEVNT1.ActiveRow.Cells("EVENT_KEY").Value & ""
                        Dim EDI_DOC_ID As String = "940"
                        Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), EDI_DOC_ID)
                        Using frm As New ASFTEXT1
                            frm.t = RAW_EDI
                            frm.Text = "Raw EDI for " & EDI_DOC_ID
                            frm.ShowDialog()
                        End Using

                    ElseIf EVENT_TYPE = "PF" Then

                        Dim PF As String = grdTATEVNT1.ActiveRow.Cells("EVENT_KEY").Value & ""
                        Dim FILENAME As String = ASCMAIN1.Folders("Archive") & "\PF\" & PF & ".PDF"
                        Show_Document(FILENAME)

                    ElseIf EVENT_TYPE = "RELSHP" Then

                        Dim EDI_DOC_SEQ_NO As String = grdTATEVNT1.ActiveRow.Cells("EVENT_KEY").Value & ""
                        Dim EDI_DOC_ID As String = "945"
                        Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), EDI_DOC_ID)
                        Using frm As New ASFTEXT1
                            frm.t = RAW_EDI
                            frm.Text = "Raw EDI for " & EDI_DOC_ID
                            frm.ShowDialog()
                        End Using

                    End If
                End If


            Case "Clone Line"

                If grd.Name = "grdSOTORDR2" Then

                    Dim ORDR_NO_original As String = grdSOTORDR2.ActiveRow.Cells("ORDR_NO").Value
                    Dim ORDR_LNO_original As Integer = grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value

                    Dim rowSOTORDR2_original As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO_original, ORDR_LNO_original})
                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow

                    rowSOTORDR2.ItemArray = rowSOTORDR2_original.ItemArray

                    Dim ORDR_LNO As Integer = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
                    With rowSOTORDR2
                        .Item("ORDR_LNO") = ORDR_LNO
                        '.ITEM("ORDR_QTY") = DBNULL.VALUE
                        If rowSOTORDR2_original.Item("STYLE_CODE_SUB") & "" <> "" Then
                            .Item("STYLE_CODE") = rowSOTORDR2_original.Item("STYLE_CODE_SUB")
                        End If
                        .Item("ORDR_QTY_ALLO") = DBNull.Value
                        .Item("ORDR_QTY_OPEN") = .Item("ORDR_QTY")
                        .Item("ORDR_QTY_PICK") = DBNull.Value
                        .Item("ORDR_QTY_SHIP") = DBNull.Value
                        .Item("ORDR_QTY_CANC") = DBNull.Value
                        .Item("ORDR_STATUS") = "O"
                        .Item("ORDR_RELEASE") = DBNull.Value
                        .Item("ORDR_RELEASE_AVAIL") = DBNull.Value
                        .Item("ORDR_QTY_ORIG") = .Item("ORDR_QTY")
                        .Item("RSRV_NO") = DBNull.Value
                        .Item("RSRV_LNO") = DBNull.Value
                        .Item("STYLE_CODE_SUB") = DBNull.Value
                        .Item("ORDR_QTY_PRE_ALLO") = DBNull.Value
                    End With
                    dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                ElseIf grd.Name = "grdSOTORDRI" Then

                    Dim STYLE_KEY_original As String = grd.ActiveRow.Cells("STYLE_KEY").Value
                    Dim rowSOTORDRI_original As DataRow = dst.Tables("SOTORDRI").Select("STYLE_KEY = '" & STYLE_KEY_original & "'")(0)

                    Dim COLOR_CODE As String = rowSOTORDRI_original.Item("COLOR_CODE")
                    Dim STYLE_CODE As String = Select_Style(COLOR_CODE)
                    If STYLE_CODE = "" Then
                        Exit Sub
                    End If

                    '   Dim ORDR_LNO_original As Int32 = Val(rowSOTORDRI_original.Item("ORDR_LNO") & "")
                    Dim ORDR_LNO As Int32 = Val(dst.Tables("SOTORDRQ").Compute("MAX(ORDR_LNO)", "") & "") + 1

                    '812-015VC:AST:::102367831:025925479663::812-015VC
                    '812-015VC:AST:::102367831:025925479663::18R
                    ' & ", NVL(SOTORDR2.STYLE_CODE_SUB,SOTORDR2.STYLE_CODE) || ':' || SOTORDR2.COLOR_CODE || ':' || NVL(SOTORDR2.CUST_STYLE_CODE,'') ||  ':' || NVL(SOTORDR2.CUST_COLOR_CODE,'') ||  ':' || NVL(SOTORDR2.CUST_SKU,'') ||  ':' || NVL(SOTORDR2.CUST_UPC,'') ||  ':' || NVL(SOTORDR2.CUST_SIZE_CODE,'') ||  ':' || NVL(SOTORDR2.STYLE_CODE,'') STYLE_KEY" & vbCrLf _
                    Dim SK() As String = Split(STYLE_KEY_original, ":")
                    SK(SK.Length - 1) = STYLE_CODE
                    Dim STYLE_KEY As String = Join(SK, ":")

                    Dim COL As Integer = -1
                    For I As Integer = 1 To 99
                        If Not dst.Tables("SOTORDRB").Columns.Contains("QTY_" & Format(I, "000")) Then
                            COL = I
                            Exit For
                        End If
                    Next

                    If COL = -1 Then Exit Sub

                    Dim rowSOTORDRI As DataRow = dst.Tables("SOTORDRI").NewRow
                    With rowSOTORDRI
                        .Item("STYLE_KEY") = STYLE_KEY
                        .Item("STYLE_CODE_ORIG") = rowSOTORDRI_original.Item("STYLE_CODE_ORIG")
                        .Item("STYLE_CODE") = STYLE_CODE

                        .Item("CUST_STYLE_CODE") = rowSOTORDRI_original.Item("CUST_STYLE_CODE")
                        .Item("CUST_UPC") = rowSOTORDRI_original.Item("CUST_UPC")
                        .Item("CUST_SKU") = rowSOTORDRI_original.Item("CUST_SKU")
                        .Item("CUST_SIZE_CODE") = rowSOTORDRI_original.Item("CUST_SIZE_CODE")

                        .Item("COLOR_CODE") = rowSOTORDRI_original.Item("COLOR_CODE")
                        .Item("CUST_COLOR_CODE") = rowSOTORDRI_original.Item("CUST_COLOR_CODE")

                        '    .Item("ORDR_LNO") = ORDR_LNO
                        .Item("STYLE_CODE_SUB") = rowSOTORDRI_original.Item("STYLE_CODE_ORIG")
                        .Item("COL") = COL

                        .Item("ORDR_UNIT_PRICE") = rowSOTORDRI_original.Item("ORDR_UNIT_PRICE")

                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        '.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                        .Item("STYLE_DESC") = rowSOTORDRI_original.Item("STYLE_DESC")

                        .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
                        .Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")

                        .Item("ORDR_QTY") = 0
                        If Val(rowSOTORDRI_original.Item("ORDR_QTY") & "") <> 0 Then
                            .Item("ORDR_QTY") = rowSOTORDRI_original.Item("ORDR_QTY")
                            .Item("ORDR_QTY_OPEN") = rowSOTORDRI_original.Item("ORDR_QTY_OPEN")
                        End If
                        '.Item("ORDR_QTY_PICK") = 0
                        '.Item("ORDR_QTY_SHIP") = 0
                        '.Item("ORDR_QTY_CANC") = 0
                        '.Item("ORDR_QTY_ORIG") = 0
                        .Item("STYLE_KEY_CLONED_FROM") = STYLE_KEY_original
                    End With

                    dst.Tables("SOTORDRI").Rows.Add(rowSOTORDRI)

                    Dim CN As String = "QTY_" & Format(COL, "000")
                    Dim C As DataColumn = dst.Tables("SOTORDRB").Columns.Add(CN, GetType(System.Int64))
                    C.DefaultValue = 0
                    With grdSOTORDRB.DisplayLayout.Bands(0).Columns(CN)
                        .Header.Caption = STYLE_CODE & ":" & COLOR_CODE
                        .Width = 80
                        .Hidden = False
                        .Format = "#,##0"

                        .Header.Appearance.TextHAlign = HAlign.Right
                        .CellAppearance.TextHAlign = HAlign.Right
                        ' Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        .Header.ToolTipText = STYLE_CODE & ":" & COLOR_CODE & vbCrLf & rowICTSTYL1.Item("STYLE_DESC") _
                            & "Customer Style: " & rowSOTORDRI.Item("CUST_STYLE_CODE") _
                            & "Customer UPC: " & rowSOTORDRI.Item("CUST_UPC") _
                            & " Customer SKU: " & rowSOTORDRI.Item("CUST_SKU") _
                            & " Customer Size: " & rowSOTORDRI.Item("CUST_SIZE_CODE")

                        Create_Summary(grdSOTORDRB, CN)
                        dst.Tables("SOTORDRB").Columns("QTY_000").Expression &= " + ISNULL(" & CN & ",0)"

                        .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                        .Header.Appearance.BackColor = Drawing.Color.White
                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    End With

                    For Each rowSOTORDRQ_original As DataRow In dst.Tables("SOTORDRQ").Select("STYLE_KEY = '" & STYLE_KEY_original & "'")

                        Dim rowSOTORDRQ As DataRow = dst.Tables("SOTORDRQ").NewRow
                        With rowSOTORDRQ
                            .ItemArray = rowSOTORDRQ_original.ItemArray
                            .Item("STYLE_KEY") = STYLE_KEY
                            .Item("ORDR_LNO") = ORDR_LNO
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("STYLE_CODE_SUB") = rowSOTORDRI_original.Item("STYLE_CODE_ORIG")
                            .Item("STYLE_CODE_ORIG") = rowSOTORDRI_original.Item("STYLE_CODE_ORIG")

                            .Item("ORDR_QTY") = 0
                            If Val(rowSOTORDRQ_original.Item("ORDR_QTY") & "") <> 0 Then
                                .Item("ORDR_QTY") = rowSOTORDRQ_original.Item("ORDR_QTY")
                                .Item("ORDR_QTY_OPEN") = rowSOTORDRQ_original.Item("ORDR_QTY_OPEN")
                                Dim rowSOTORDRB As DataRow = dst.Tables("SOTORDRB").Rows.Find(rowSOTORDRQ_original.Item("ORDR_NO"))
                                rowSOTORDRB.Item("QTY_" & Format(COL, "000")) = rowSOTORDRQ_original.Item("ORDR_QTY_OPEN")
                            End If
                            .Item("ORDR_QTY_PICK") = 0
                            .Item("ORDR_QTY_SHIP") = 0
                            .Item("ORDR_QTY_CANC") = 0
                            '  .Item("ORDR_QTY_ORIG") = 0
                        End With
                        dst.Tables("SOTORDRQ").Rows.Add(rowSOTORDRQ)

                        Dim ORDR_NO As String = rowSOTORDRQ.Item("ORDR_NO")
                        If dst.Tables("SOTORDRQ_KEY").Rows.Find(New Object() {ORDR_NO, STYLE_KEY}) Is Nothing Then
                            dst.Tables("SOTORDRQ_KEY").Rows.Add(New Object() {ORDR_NO, STYLE_KEY, ORDR_LNO})
                        Else
                            If multistore_OK_TO_UPDATE Then
                                MsgBox("Problem with Order " & ORDR_NO & ", Line " & ORDR_LNO, MsgBoxStyle.OkOnly, "Integrity Issue with Multi-Store Edit")
                            End If

                            multistore_OK_TO_UPDATE = False
                        End If
                    Next
                End If

            Case "Show Raw EDI 940"

                'If grdSOTORDR1.ActiveRow IsNot Nothing Then
                '    Dim EDI_DOC_SEQ_NO As String = grdSOTORDR1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                'End If
                '  Display_Raw(grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value & "")

                If grd.ActiveRow IsNot Nothing Then
                    ASCMAIN1.sql = "SELECT * FROM EDT940O1 WHERE PICK_NO = '" & grd.ActiveRow.Cells("PICK_NO").Value & "'"
                    Dim rowEDT940O1 As DataRow = ASCDATA1.GetDataRow

                    If rowEDT940O1 IsNot Nothing Then
                        Dim EDI_DOC_SEQ_NO As String = rowEDT940O1.Item("EDI_OUTBOUND_DOC_NO")
                        Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), "940")
                        Using frm As New ASFTEXT1
                            frm.t = RAW_EDI
                            frm.Text = "Raw EDI 940 for " & CUST_CODE & " Pick Ticket No " & grd.ActiveRow.Cells("PICK_NO").Value
                            frm.ShowDialog()
                        End Using
                    End If

                End If

            Case "Sales Order Inquiry"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", SO_ORDER_NO)
                If rowSOTINVH1 IsNot Nothing Then
                    Context_Launch("View", SO_ORDER_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                    Exit Sub
                End If

                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

                'Case "Style Master"
                '    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                '    If rowICTSTYL1 IsNot Nothing Then
                '        Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                '    End If

            Case "Style Master File"
                If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                    Exit Sub
                End If

                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If

            Case "Update Qty to All Stores"
                Dim ORDR_QTY As Int64 = Val(grdSOTORDRS.ActiveCell.Value & "")
                Dim COLUMN_NAME As String = grdSOTORDRS.ActiveCell.Column.Key
                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                    rowSOTORDRS.Item(COLUMN_NAME) = ORDR_QTY
                Next

            Case "Track Shipment"
                Dim CART_TRACKING_NO As String = ""
                Dim PICK_NO As String = ""
                If grd.ActiveRow.Band.Key = "SOTCART1" Then
                    CART_TRACKING_NO = grd.ActiveRow.Cells("CART_TRACKING_NO").Value & ""
                    PICK_NO = grd.ActiveRow.Cells("PICK_NO").Value & ""
                ElseIf grd.ActiveRow.Band.Key = "SOTPICK1" Then
                    PICK_NO = grd.ActiveRow.Cells("PICK_NO").Value & ""
                    ASCMAIN1.sql = "Select MIN(CART_TRACKING_NO) CART_TRACKING_NO from SOTCART1 where PICK_NO = '" & PICK_NO & "' and CART_TRACKING_NO is Not Null"
                    CART_TRACKING_NO = ASCDATA1.GetDataValue
                Else
                    CART_TRACKING_NO = grd.ActiveRow.ParentRow.Cells("CART_TRACKING_NO").Value & ""
                    PICK_NO = grd.ActiveRow.ParentRow.Cells("PICK_NO").Value & ""
                End If

                Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                If rowSOTPICK1 Is Nothing Then
                    MsgBox("Could not locate Pick Ticket (" & PICK_NO & ") for this shipment", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                Else
                    If rowSOTPICK1.Item("PICK_STATUS") <> "F" Then
                        MsgBox("Pick Ticket (" & PICK_NO & ") has not been Shipped yet", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                    Else
                        If CART_TRACKING_NO = "" Then
                            MsgBox("Could not locate Carton Tracking Number for this shipment", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                        Else
                            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_VIA_CODE from SOTSHIP1,SOTPICK1 where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO and SOTPICK1.PICK_NO = '" & PICK_NO & "'"
                            Dim SHIP_VIA_CODE As String = ASCDATA1.GetDataValue
                            TAC.SOCMAIN1.Track_Shipment(SHIP_VIA_CODE, CART_TRACKING_NO)
                        End If
                    End If
                End If

            Case "Show Pro-Forma Invoice", "email Pro-Forma Invoice"

                ASCMAIN1.Progress("Now Preparing Document")

                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value

                Dim REPORTFILE As String = "SORINVP1"
                If Not REPORTS.ContainsKey(REPORTFILE) Then
                    REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
                    REPORTS(REPORTFILE).Prepare_dst(False, "")
                End If

                Dim RPT As String = "SORINVP1"
                Dim AR_PARM_INVOICE_RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
                If AR_PARM_INVOICE_RPT <> "" Then RPT = AR_PARM_INVOICE_RPT ' "SORINVP1"

                REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and SOTORDR1.ORDR_NO = '" & ORDR_NO & "'", "1", "O"})

                Dim rowSOTINVH1 As DataRow = REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTINVH1").Rows(0)
                rowSOTINVH1.Item("INV_NO") = INV_NO
                rowSOTINVH1.Item("INV_TYPE") = "I"
                rowSOTINVH1.Item("PICK_NO") = "BTB"

                Dim ORDR_CUST_PO As String = Absx1.txtFor("ORDR_CUST_PO").Text
                If Not ScreenMode Then
                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                    ORDR_CUST_PO = rowSOTORDR1.Item("ORDR_CUST_PO") & ""
                End If
                rowSOTINVH1.Item("ORDR_CUST_PO") = ORDR_CUST_PO ' grd.ActiveRow.Cells("INV_REF").Value
                rowSOTINVH1.Item("INV_DATE") = grd.ActiveRow.Cells("INV_DATE").Value
                rowSOTINVH1.Item("INV_SALES") = grd.ActiveRow.Cells("INV_TOTAL_AMOUNT").Value
                rowSOTINVH1.Item("INV_SALES_CURR") = grd.ActiveRow.Cells("INV_TOTAL_AMOUNT").Value
                rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = grd.ActiveRow.Cells("INV_TOTAL_AMOUNT").Value
                rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR") = grd.ActiveRow.Cells("INV_TOTAL_AMOUNT").Value
                rowSOTINVH1.Item("INV_TOTAL_AMT_CURR") = grd.ActiveRow.Cells("INV_TOTAL_AMOUNT").Value
                rowSOTINVH1.Item("INV_COMMENT") = grd.ActiveRow.Cells("INV_COMMENT").Value

                If rowSOTINVH1.Item("SHIP_BOL_NO") & "" = "" Then
                    rowSOTINVH1.Item("SHIP_BOL_NO") = "0000000000"
                    REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTSHIP1").Rows.Add(New Object() {"0000000000"})
                End If
                Dim rowSOTSHIP1 As DataRow = REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTSHIP1").Rows(0)

                'rowSOTSHIP1.Item("INV_NO") = INV_NO
                'rowSOTSHIP1.Item("INV_TYPE") = "I"
                'rowSOTSHIP1.Item("PICK_NO") = "BTB"
                ' rowSOTSHIP1.Item("SHIP_CNT_CARTONS") = 99
                '  Dim rowPOTSHIP1 As DataRow = LookUp("POTSHIP1", "")
                rowSOTSHIP1.Item("BILL_OF_LADING_NO") = grd.ActiveRow.Cells("BILL_OF_LADING_NO").Value
                rowSOTSHIP1.Item("SHIP_VIA_CODE") = Absx1.txtFor("SHIP_VIA_CODE").Text
                rowSOTSHIP1.Item("FRT_TERMS") = Absx1.txtFor("FRT_TERMS").Text

                If rowSOTINVH1.Item("PICK_NO") & "" = "" Or rowSOTINVH1.Item("PICK_NO") & "" = "BTB" Then
                    rowSOTINVH1.Item("PICK_NO") = "0000000000"
                    Dim rowSOTPICK1_0 As DataRow = REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTPICK1").NewRow
                    rowSOTPICK1_0.Item("PICK_NO") = "0000000000"
                    rowSOTPICK1_0.Item("PICK_STATUS") = "F"
                    REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1_0)
                    'REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTPICK1").Rows.Add(New Object() {"0000000000"})
                End If
                Dim rowSOTPICK1 As DataRow = REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTPICK1").Rows(0)
                rowSOTPICK1.Item("PICK_CNT_CARTONS") = Val(grd.ActiveRow.Cells("PICK_CNT_CARTONS").Value & "")
                rowSOTPICK1.Item("SHIP_BOL_NO") = rowSOTINVH1.Item("SHIP_BOL_NO")
                rowSOTPICK1.Item("INV_NO") = rowSOTINVH1.Item("INV_NO")

                'Dim rowSOTSHIP1 As DataRow = REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTSHIP1").Rows(0)
                'rowSOTSHIP1.Item("SHIP_REF") = grd.ActiveRow.Cells("INV_REF").Value

                For Each rowSOTINVH2 As DataRow In REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTINVH2").Select("")
                    Dim INV_LNO As Int32 = Val(rowSOTINVH2.Item("INV_LNO") & "")
                    Dim rowSOTORDP2 As DataRow = dst.Tables("SOTORDP2").Rows.Find(New Object() {ORDR_NO, INV_NO, INV_LNO})
                    If rowSOTORDP2 Is Nothing Then
                        rowSOTINVH2.Item("ORDR_QTY_SHIP") = 0
                    Else
                        rowSOTINVH2.Item("ORDR_QTY_SHIP") = rowSOTORDP2.Item("ORDR_QTY_SHIP")
                    End If
                Next
                ASCDATA1.DeleteRows(REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTINVH2"), "ORDR_QTY_SHIP = 0")


                With REPORTS(REPORTFILE).clsASCBASE1
                    .Print_Report_Begin()
                    .CR_params.Add("SUBT", "")
                    .CR_params.Add("CONS_INV", "0")
                    .CR_params.Add("EXPORT_INFO", IIf(chkExportInfo.Checked, "1", "0"))

                    If e.Tool.Key = "email Pro-Forma Invoice" Then
                        Dim SUBJECT As String = "Invoice " & INV_NO

                        Dim tempFileName As String = INV_NO
                        Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
                        Dim FILENAME As String = .F.REPORT_FILENAMES(REPORT_NO)
                        .Print_Report_End(, True)

                        Dim SEND_NO As String = TAC.SOCMAIN1.email_Invoice(Me,
                             CUST_CODE,
                             CUST_NAME,
                             rowARTCUST1.Item("CUST_EMAIL") & "",
                             rowARTCUST1.Item("CUST_CONTACT") & "",
                             FILENAME, FILENAME, SUBJECT, INV_NO, ORDR_NO)
                        If SEND_NO <> "" Then
                            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTORDR1' and TABLE_KEY = '" & ORDR_NO & "' and EVENT_KEY = '" & SEND_NO & "'"
                            Fill_Records("TATEVNT1", "", False, ASCMAIN1.sql)
                            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
                        End If

                    Else
                        .Generate_Report(RPT, "Pro-Forma Invoice", , True, , , , , False)
                        .Print_Report_End()
                    End If
                End With

            Case "Show email"
                If grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "EML" _
                    Or grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "CORSTA" Or grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "VASEML" Then
                    Dim FILENAME As String = grd.ActiveRow.Cells("EVENT_KEY").Value & ".EML"
                    Show_Document(ASCMAIN1.Folders("Archive") & "\email\Sent\" & FILENAME)
                End If

            Case "Show Invoice", "email Invoice", "Show Invoice as Report"

                ASCMAIN1.Progress("Now Preparing Document")

                Dim ATTACHMENT As String = ""
                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value & ""
                Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value & ""
                Dim return_PDF As Boolean = Not (e.Tool.Key = "Show Invoice as Report")
                Dim FILENAME As String = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO, return_PDF, , , , IIf(chkExportInfo.Checked, "1", "0"))
                Dim SUBJECT As String = "Invoice " & INV_NO

                If INV_NO = "" Then
                    MsgBox("Pick Ticket (" & PICK_NO & ") has not been Shipped yet", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                    Exit Sub
                End If

                If e.Tool.Key = "email Invoice" Then
                    Dim SEND_NO As String = TAC.SOCMAIN1.email_Invoice(Me,
                      CUST_CODE,
                      CUST_NAME,
                      rowARTCUST1.Item("CUST_EMAIL") & "",
                      rowARTCUST1.Item("CUST_CONTACT") & "",
                      FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT, INV_NO, ORDR_NO)

                    If SEND_NO <> "" Then
                        ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTORDR1' and TABLE_KEY = '" & ORDR_NO & "' and EVENT_KEY = '" & SEND_NO & "'"
                        Fill_Records("TATEVNT1", "", False, ASCMAIN1.sql)

                        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
                    End If

                ElseIf e.Tool.Key = "Show Invoice" Then
                    Show_Document(FILENAME)
                End If

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")


        End Select
    End Sub

    Private Sub ImportDetailsFromExcel()
        Dim tstMsg As String = vbCrLf & "This Feature Is Under Test." & vbCrLf & "Please Review Your Data."
        Dim Results As Text.StringBuilder = ImportDetailsToGrid()
        If Results.Length > 0 Then
            MsgBox(Results.ToString & tstMsg, vbCritical, "Import Errors")
        Else
            MsgBox("Import Complete." & tstMsg, vbOK, "Import Errors")
        End If
    End Sub
    Private Sub ExcelProcessKill()
        Dim oProcesses() As Process
        Dim bFound As Boolean

        Try
            'Get all currently running process Ids for Excel applications
            oProcesses = Process.GetProcessesByName("Excel")

            If oProcesses.Length > 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    bFound = False

                    'For j As Integer = 0 To mExcelProcesses.Length - 1
                    '    If oProcesses(i).Id = mExcelProcesses(j).Id Then
                    '        bFound = True
                    '        Exit For
                    '    End If
                    'Next

                    If Not bFound Then
                        oProcesses(i).Kill()
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "CUST_NAME_SEARCH"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If optShowOrders.Value = "N" Then
                        Dim CUST_NAME_SEARCH As String = Absx1.txtFor("CUST_NAME_SEARCH").Text
                        CUST_NAME_SEARCH = CUST_NAME_SEARCH.Trim
                        If CUST_NAME_SEARCH.Length < 3 Then
                            MessageBox.Show("Strore Name requires a minimum of 3 characters.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                        Load_SOTORDRX()
                    End If
                End If

            Case "CUST_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If optShowOrders.Value = "C" Then
                        Load_SOTORDRX()
                    Else
                        optShowOrders.Value = "C"
                    End If
                End If

            Case "CUST_STORE_NO"
                If ScreenMode Then
                    e.SuppressKeyPress = True
                    e.Handled = True
                    Exit Sub
                End If

                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode _
                        And Absx1.txtFor("CUST_CODE").Text <> "" _
                        And Absx1.txtFor("CUST_STORE_NO").Text <> "" _
                        And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "ORDR_CUST_PO"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode _
                       And Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("CUST_STORE_NO").Text <> "" _
                       And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "ORDR_NO"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If

            Case "INV_NO"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim INV_NO As String = txtINV_NO.Text
                    Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"I", INV_NO})
                    If rowSOTINVH1 Is Nothing Then
                        MsgBox("No Record of Invoice " & INV_NO, MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Else
                        Absx1.txtFor("ORDR_NO").Text = rowSOTINVH1.Item("ORDR_NO")
                        Click_Command("View")
                    End If

                End If

            Case "PICK_NO"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim PICK_NO As String = txtPICK_NO.Text
                    Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", New String() {PICK_NO})
                    If rowSOTPICK1 Is Nothing Then
                        MsgBox("No Record of Pick Ticket " & PICK_NO, MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Else
                        Absx1.txtFor("ORDR_NO").Text = rowSOTPICK1.Item("ORDR_NO")
                        Click_Command("View")
                    End If

                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    If optShowOrders.Value = "C" Then
                        Load_SOTORDRX()
                    Else
                        optShowOrders.Value = "C"
                    End If

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 IsNot Nothing Then
                            ASCMAIN1.sql = "Select Count (*) STORES, Max (CUST_ADDR_CODE) CUST_STORE_NO from ARTCUST2" _
                                & " where CUST_CODE = :PARM1 and CUST_ADDR_TYPE = 'MK'"
                            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {CUST_CODE})
                            If Val(row.Item("STORES") & "") = 1 Then
                                Absx1.txtFor("CUST_STORE_NO").Text = row.Item("CUST_STORE_NO")
                                Absx1.txtFor("CUST_STORE_NO").Focus()
                            End If

                        End If
                    End If
                End If

            Case "WHSE_CODE"
                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    Setup_DISC_AMT(Absx1.txtFor("WHSE_CODE").Text)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If optShowOrders.Value = "C" Then
                    Load_SOTORDRX()
                Else
                    optShowOrders.Value = "C"
                End If

            Case "CUST_STORE_NO"
                If ScreenMode And ASCMAIN1.CLIENT = "NYA" And EntryMode = "E" And Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                    Dim CUST_STORE_NO As String = Absx1.txtFor("CUST_STORE_NO").Text
                    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
                    If rowARTCUST2 IsNot Nothing Then

                        For Each CUST_ADDR_TYPE As String In New String() {"MK", "ST"}
                            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, CUST_ADDR_TYPE})
                            With rowSOTORDR5
                                For Each COLUMN_NAME As String In New String() _
                                   {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE",
                                    "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}
                                    Dim COLUMN_NAME_ST As String = COLUMN_NAME
                                    .Item(COLUMN_NAME) = rowARTCUST2.Item(COLUMN_NAME_ST)
                                Next
                                .Item("CUST_ADDR_CODE") = CUST_STORE_NO
                            End With
                        Next

                        Synch_TABLE_NAME("SOTORDR1")

                        rowSOTORDR1.Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_NAME")
                        Me.CUST_STORE_NO = CUST_STORE_NO
                    End If
                End If

            Case "ORDR_NO"
                Click_Command("View")
            Case "ORDR_NO_WEB"
                ASCMAIN1.sql = String.Format("Select ORDR_NO from SOTORDR1 where ORDR_NO_WEB = '{0}'", txtORDR_NO_WEB.Text)
                Dim ORDR_NO As String = ASCDATA1.GetDataValue
                Absx1.txtFor("ORDR_NO").Text = ORDR_NO
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        '  If Me.IsClosing Then Exit Sub

        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ORDR_ADDR_TYPE_ST"
                If Not Me.IsLoading Then
                    'Synch_TABLE_NAME("SOTORDR5")
                    Dim X As CurrencyManager = Me.BindingContext(dvwSOTORDR5)
                    X.EndCurrentEdit()
                    Dim rowSOTORDR5_ST As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, "ST"})
                    Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, optShipTo.Value})
                    For Each dc As DataColumn In dst.Tables("SOTORDR5").Columns
                        If dc.ColumnName <> "ORDR_NO" And dc.ColumnName <> "CUST_ADDR_TYPE" Then
                            rowSOTORDR5_ST.Item(dc.ColumnName) = rowSOTORDR5.Item(dc.ColumnName)
                        End If
                    Next
                End If

            Case "ORDR_TYPE_CODE"
                '  lblCUST_STORE_NO.Visible = (Absx1.optFor("ORDR_TYPE_CODE").Value & "" <> "XFR")
                '  txtCUST_STORE_NO.Visible = (Absx1.optFor("ORDR_TYPE_CODE").Value & "" <> "XFR")
                lblWHSE_CODE_TO.Visible = (Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR")
                txtWHSE_CODE_TO.Visible = (Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR")

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "ORDR_CANCEL_DATE"
                If ScreenMode And grdSOTORDR2.Rows.Count <> 0 Then
                    grdSOTORDR2.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
                End If
        End Select

    End Sub

#End Region

    Sub Load_SOTORDPX()

        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        ASCMAIN1.sql = "Select SOTORDP1.*, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
            & " From SOTORDP1, SOTORDR1 Where SOTORDR1.ORDR_NO = SOTORDP1.ORDR_NO" & vbCrLf _
            & " and SOTORDR1.ORDR_STATUS IN ('O','P')"
        Fill_Records("SOTORDP1", "", , ASCMAIN1.sql)

        ASCMAIN1.sql = "Select Distinct SOTORDP2.*" & vbCrLf _
            & " From SOTORDP2, SOTORDP1, SOTORDR1 Where SOTORDP2.ORDR_NO = SOTORDP1.ORDR_NO" & vbCrLf _
            & "  and SOTORDR1.ORDR_NO = SOTORDP1.ORDR_NO" & vbCrLf _
            & "  and SOTORDR1.ORDR_STATUS In ('O','P')"
        Fill_Records("SOTORDP2", "", , ASCMAIN1.sql)

        dst.Tables("SOTORDR2").Rows.Clear()

        For Each row As DataRow In ASCDATA1.SelectDistinct("SOTORDP1", New String() {"ORDR_NO"}).Select("")
            Dim ORDR_NO As String = row.Item("ORDR_NO")
            Fill_Records("SOTORDR2", ORDR_NO, False)
        Next


        EnforceConstraints(True)

        Sort_grdColumns(grdSOTORDP1, "ORDR_NO".ToLower)

        tabSOTORDRX.Tabs("Pro-Forma Invoices").Tag = ""


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Load_SOTORDRX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        txtCustNameSearch.Left = dteSearchS.Left
        dst.Tables("SOTORDRX").Rows.Clear()

        'ASCMAIN1.sql = "Select SOTORDR1.*, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" _
        '    & " from SOTORDR1, ARTCUST1" _
        '    & " WHERE SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE"

        ASCMAIN1.sql = " Select SOTORDR1.*, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY, ARTCUST1.CUST_CREDIT_HOLD" _
         & " , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_OPEN, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_OPEN" _
         & " , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_PICK, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_PICK" _
         & " , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_SHIP, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_SHIP" _
         & " , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_CANC, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_CANC" _
         & "  from SOTORDR1, ARTCUST1" _
         & "  WHERE SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE"

        ASCMAIN1.Progress("Now Building List of Sales Orders", "")
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If optShowOrders.Value = "A" And CUST_CODE = "" Then
            'ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_STATUS = 'O'"
            ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = 'O'"
            grdSOTORDRX.Text = "All Open Sales Orders"
        ElseIf optShowOrders.Value = "M" Then
            'ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_STATUS = 'O' and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = 'O' and (SOTORDR1.INIT_OPER = '" & ASCMAIN1.USER_ID & "' or SOTORDR1.LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            grdSOTORDRX.Text = "Open Sales Orders entered or modified by Me"
        ElseIf optShowOrders.Value = "C" Or CUST_CODE <> "" Then
            'ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_STATUS = '" & optCustomerOrders.Value & "' and CUST_CODE = '" & CUST_CODE & "'"
            ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = '" & optCustomerOrders.Value & "' and SOTORDR1.CUST_CODE = '" & CUST_CODE & "'"
            grdSOTORDRX.Text = "Open Sales Orders associated with " & CUST_CODE
        ElseIf optShowOrders.Value = "D" Then
            ASCMAIN1.sql &= " and SOTORDR1.ORDR_DATE BETWEEN '" & dteSearchS.DateTime.ToString("dd-MMM-yyyy") & "' and '" & dteSearchE.DateTime.ToString("dd-MMM-yyyy") & "'"
            grdSOTORDRX.Text = "Sales Orders created between " & dteSearchS.DateTime.ToString("MM/dd/yyyy") & " and " & dteSearchE.DateTime.ToString("MM/dd/yyyy")
        ElseIf optShowOrders.Value = "N" Then
            txtCustNameSearch.Text = txtCustNameSearch.Text.Trim
            txtCustNameSearch.Text = txtCustNameSearch.Text.Replace("'", "")
            If txtCustNameSearch.TextLength = 0 Then
                ASCMAIN1.sql &= " and ROWNUM < 1"
            Else
                ASCMAIN1.sql &= " and UPPER(SOTORDR1.CUST_STORE_NAME) LIKE '%" & txtCustNameSearch.Text.ToUpper & "%'"
                ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = '" & optCustomerOrders.Value & "'"
                ASCMAIN1.sql &= " and SOTORDR1.ORDR_DATE BETWEEN '" & dteSearchS.DateTime.ToString("dd-MMM-yyyy") & "' and '" & dteSearchE.DateTime.ToString("dd-MMM-yyyy") & "'"
            End If
            grdSOTORDRX.Text = optCustomerOrders.Text & " Sales Orders for Customer Name like " & txtCustNameSearch.Text
        Else
            'ASCMAIN1.sql = "Select * from SOTORDR1 where ROWNUM < 1"
            ASCMAIN1.sql &= " and ROWNUM < 1"
        End If

        If chkOnHoldOnly.Checked Then
            ASCMAIN1.sql &= " and (nvl(SOTORDR1.ORDR_HOLD, '0') = '1' or nvl(ARTCUST1.CUST_CREDIT_HOLD, '0') = '1')"
        End If

        If ASCMAIN1.CLIENT = "NYA" And ASCMAIN1.USER_CODES = "CA" Then
            ASCMAIN1.sql &= " and SOTORDR1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")"
        End If

        Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)

        Sort_grdColumns(grdSOTORDRX, "ORDR_NO".ToLower)
        grdSOTORDRX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdSOTORDRX.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Order")

        Dim REPORT_NAME As String = "SORORDR1"
        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {" and SOTORDR1.ORDR_NO = '" & ORDR_NO & "'"})
        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            Dim SUBT As String = ""
            .CR_params.Add("SUBT", SUBT)
            .Generate_Report(REPORT_NAME, "Sales Order", SUBT, True, , , , , False)
            .Print_Report_End()
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_ProForma(Optional ORDR_QTY_field As String = "")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Invoice")
        Dim pfComment As String = ""
        Dim REPORT_NAME As String = "SORINVP1"
        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If RPT = "" Then
            RPT = REPORT_NAME

        End If

        Dim force_proforma = (ASCMAIN1.CLIENT = "VAN" And InquiryMode)
        If rowSOTORDR1.Item("ORDR_STATUS") = "F" And Not force_proforma Then
            If ASCMAIN1.CLIENT = "VAN" Then
                REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {" and SOTINVH1.ORDR_NO = '" & ORDR_NO & "'", "1", pfComment})
            Else
                REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {" and SOTINVH1.ORDR_NO = '" & ORDR_NO & "'"})
            End If
            With REPORTS(REPORT_NAME).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "0")
                .CR_params.Add("EXPORT_INFO", IIf(chkExportInfo.Checked, "1", "0"))
                .Generate_Report(RPT, "Sales Invoice", , True, , , , , False)
                .Print_Report_End()
            End With
        Else

            REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {" and SOTORDR1.ORDR_NO = '" & ORDR_NO & "'", "1", "O", ORDR_QTY_field})

            If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
                Dim rowSOTPICK1 As DataRow ' = REPORTS(REPORT_NAME).dst.Tables("SOTPICK1").Rows.Find(ORDR_NO)
                rowSOTPICK1 = REPORTS(REPORT_NAME).dst.Tables("SOTPICK1").NewRow
                rowSOTPICK1.Item("PICK_NO") = ORDR_NO
                rowSOTPICK1.Item("ORDR_NO") = ORDR_NO
                rowSOTPICK1.Item("PICK_CNT_CARTONS") = Absx1.numFor("PF_CARTONS").Value
                rowSOTPICK1.Item("PICK_TOTAL_WGT") = Absx1.numFor("PF_WEIGHT").Value
                rowSOTPICK1.Item("PF_WEIGHT_UOM") = Absx1.optFor("PF_WEIGHT_UOM").Text
                rowSOTPICK1.Item("PO_SHIPMENT_NO") = Absx1.txtFor("PO_SHIPMENT_NO").Text
                rowSOTPICK1.Item("PICK_STATUS") = "P"
                REPORTS(REPORT_NAME).dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

                Dim rowSOTINVH1 As DataRow = REPORTS(REPORT_NAME).dst.Tables("SOTINVH1").Rows.Find(New Object() {"P", ORDR_NO})
                rowSOTINVH1.Item("INV_COMMENT") = Absx1.txtFor("PF_NOTE").Text
                rowSOTINVH1.Item("PICK_NO") = ORDR_NO
                rowSOTINVH1.Item("PF_OVERSEAS_DOMESTIC") = Absx1.optFor("PF_OVERSEAS_DOMESTIC").Value
                rowSOTINVH1.Item("INV_DATE") = Absx1.dteFor("PF_INV_DATE").Value
                rowSOTINVH1.Item("PF_VIA") = Absx1.txtFor("PF_VIA").Text
                rowSOTINVH1.Item("PO_SHIPMENT_NO") = Absx1.txtFor("PO_SHIPMENT_NO").Text

                Dim rowSOTORDR1_PF As DataRow = REPORTS(REPORT_NAME).dst.Tables("SOTORDR1").Rows.Find(New Object() {ORDR_NO})
                rowSOTORDR1_PF.Item("ORDR_FOB") = Absx1.txtFor("ORDR_FOB").Text

                If Absx1.txtFor("PO_SHIPMENT_NO").Text <> "" Then
                    'ASCMAIN1.sql = "SELECT PO_ORDER_NO,PO_SPEC_ORDR_NO FROM POTORDR1 WHERE PO_ORDER_NO IN (Select DISTINCT PO_ORDER_NO FROM POTSHIP3" _
                    '        & " Where PO_SHIPMENT_NO = '" & Absx1.txtFor("PO_SHIPMENT_NO").Text & "')"
                    'ASCMAIN1.sql = "SELECT 1 SEQ, 'ORDER #s' TYPE,LISTAGG(PO_SPEC_ORDR_NO,',') WITHIN GROUP (ORDER BY PO_SPEC_ORDR_NO) AS PO_SPEC_ORDR_NO FROM (" _
                    '        & " SELECT DISTINCT PO_SPEC_ORDR_NO FROM POTORDR1 WHERE PO_ORDER_NO IN (" _
                    '        & " SELECT DISTINCT PO_ORDER_NO FROM POTSHIP3 WHERE PO_SHIPMENT_NO = '" & Absx1.txtFor("PO_SHIPMENT_NO").Text & "'))" _
                    '        & " UNION " _
                    '        & " SELECT 2 SEQ,'ASHLEY INV#s' TYPE,LISTAGG(COMM_INV_NO,',') WITHIN GROUP (ORDER BY COMM_INV_NO) AS PO_SPEC_ORDR_NO" _
                    '        & " FROM (SELECT DISTINCT COMM_INV_NO FROM POTSHIP2 WHERE PO_SHIPMENT_NO = '" & Absx1.txtFor("PO_SHIPMENT_NO").Text & "')"

                    'Dim COMMENTS As String = ""
                    'For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "SEQ")
                    '    COMMENTS = COMMENTS & row.Item(1) & " " & row.Item(2) & "     "

                    'Next
                    'rowSOTINVH1.Item("PF_SHIP_NOTES") = COMMENTS

                    rowSOTINVH1.Item("PF_SHIP_NOTES") = "Invoice No " & Absx1.txtFor("PF_INV_NO").Text
                End If


                Dim INV_SALES As Decimal = 0
                For Each rowSOTINVH2 As DataRow In REPORTS(REPORT_NAME).dst.Tables("SOTINVH2").Select()
                    '' Need to put in code to get DUTY_HTS_CODE AN DISPLAY SHIP VIA / INVOICE NUMBER DEPENDING ON Absx1.txtFor("PO_SHIPMENT_NO").Text
                    Dim ORDR_LNO As Integer = Val(rowSOTINVH2.Item("INV_LNO") & "")
                    Dim ROWSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                    rowSOTINVH2.Item("ORDR_QTY_SHIP") = ROWSOTORDR2.Item("PF_QTY")
                    rowSOTINVH2.Item("LIC_CODE") = ROWSOTORDR2.Item("PF_ORDER_NO")
                    rowSOTINVH2.Item("DUTY_HTS_CODE") = ROWSOTORDR2.Item("PF_DUTY_HTS_CODE")

                    Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & "")
                    Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE") & "")
                    INV_SALES += ORDR_QTY_SHIP * ORDR_UNIT_PRICE
                Next

                rowSOTINVH1.Item("INV_SALES") = INV_SALES
                rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = INV_SALES

            End If

            With REPORTS(REPORT_NAME).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "0")
                .CR_params.Add("EXPORT_INFO", IIf(chkExportInfo.Checked, "1", "0"))

                If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then

                    Dim PF As String = ASCMAIN1.Next_Control_No("SOTINVH1.PF")
                    ASCMAIN1.Record_Event("SOTORDR1", ORDR_NO, "", Now, ASCMAIN1.USER_ID, "PF", "Pro-Forma Invoice", PF)
                    ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTORDR1' and TABLE_KEY = '" & ORDR_NO & "' and EVENT_TYPE = 'PF' and EVENT_KEY = '" & PF & "'"
                    Fill_Records("TATEVNT1", "", False, ASCMAIN1.sql)
                    Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

                    .Generate_Report(RPT, "Pro-Forma Invoice - " & CUST_CODE & " PO " & Absx1.txtFor("ORDR_CUST_PO").Text, , True, , , , , True)

                    .CR_params.Add("SUBT", "")
                    .CR_params.Add("CONS_INV", "0")
                    .CR_params.Add("EXPORT_INFO", IIf(chkExportInfo.Checked, "1", "0"))
                    .Generate_Report(RPT, "Pro-Forma Invoice", , False, , , "PDF", PF, False)

                    My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & PF & ".PDF", ASCMAIN1.Folders("Archive") & "\PF\" & PF & ".PDF")

                Else

                    .Generate_Report(RPT, "Sales Order Confirmation", , True, , , , , False)
                End If

                .Print_Report_End()

            End With
        End If

        If ASCMAIN1.CLIENT = "VAN" Then ' Produce Export Documents

            If rowSOTORDR1.Item("ORDR_STATUS") = "F" Then
                ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_NO = '" & rowSOTORDR1.Item("ORDR_NO") & "'"
                Dim rowSOTINVH1 = ASCDATA1.GetDataRow
            End If

            Dim xls_filename As String = "ExportDocuments.xlsx"
            Dim filename As String = ASCMAIN1.Folders("SharedRoot") & "\Templates\" & xls_filename
            If ASCMAIN1.Running_in_VS Then
                filename = "C:\Share\VDI\Templates\" & xls_filename
            Else
                filename = "R:\VDI\Templates\" & xls_filename
            End If

            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(filename)

            Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
            Dim range As SpreadsheetGear.IRange = Nothing

            Dim RX As Int32 = 0
            Dim CX As Int32 = 0



            oSheet = oWB.Worksheets("Packing List")
            oSheet.Cells("F3").NumberFormat = "@"
            oSheet.Cells("F3").Value = rowSOTORDR1.Item("ORDR_NO")
            oSheet.Cells("H3").NumberFormat = "MM/dd/yyyy"
            oSheet.Cells("H3").Value = rowSOTORDR1.Item("ORDR_DATE") ' *******************

            oSheet.Cells("B14").NumberFormat = "@"
            oSheet.Cells("B14").Value = rowSOTORDR1.Item("CUST_CODE")
            oSheet.Cells("D14").NumberFormat = "@"
            oSheet.Cells("D14").Value = rowSOTORDR1.Item("CUST_STORE_NO")

            oSheet.Cells("A17").NumberFormat = "@"
            oSheet.Cells("A17").Value = rowSOTORDR1.Item("ORDR_NO")
            oSheet.Cells("B17").NumberFormat = "@"
            oSheet.Cells("B17").Value = rowSOTORDR1.Item("ORDR_CUST_PO")
            oSheet.Cells("C17").NumberFormat = "@"
            oSheet.Cells("C17").Value = rowSOTORDR1.Item("ORDR_DEPT")
            oSheet.Cells("D17").NumberFormat = "@"
            oSheet.Cells("D17").Value = rowSOTORDR1.Item("SREP_CODE")
            oSheet.Cells("E17").NumberFormat = "@"
            oSheet.Cells("E17").Value = rowSOTORDR1.Item("TERM_CODE")
            oSheet.Cells("F17").NumberFormat = "@"
            oSheet.Cells("F17").Value = rowSOTORDR1.Item("SHIP_VIA_CODE")

            oSheet.Cells("H17").NumberFormat = "@"
            oSheet.Cells("H17").Value = rowSOTORDR1.Item("ORDR_FOB")
            oSheet.Cells("I17").NumberFormat = "#,##0"
            oSheet.Cells("I17").Value = Absx1.numFor("PF_CARTONS").Value
            oSheet.Cells("J17").NumberFormat = "#,##0.0"
            oSheet.Cells("J17").Value = Absx1.numFor("PF_WEIGHT").Value & " " & Absx1.optFor("PF_WEIGHT_UOM").Text
            oSheet.Cells("L17").NumberFormat = "@"
            oSheet.Cells("L17").Value = ""

            For Each CUST_ADDR_TYPE As String In New String() {"BT", "ST"}
                Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {rowSOTORDR1.Item("ORDR_NO"), CUST_ADDR_TYPE})
                Dim CA As Integer = 5
                If CUST_ADDR_TYPE = "BT" Then
                    CA = 0
                End If
                Dim RA As Integer = 5

                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_NAME")
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR1")
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR2")
                'RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR3")
                Dim CSZ As String = rowSOTORDR5.Item("CUST_CITY") & "," & rowSOTORDR5.Item("CUST_STATE") & " " & rowSOTORDR5.Item("CUST_ZIP_CODE") & " " & rowSOTORDR5.Item("CUST_COUNTRY")
                RA += 1 : oSheet.Cells(RA, CA).Value = CSZ
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_CONTACT")
            Next


            RX = 18
            CX = 0

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "ORDR_LNO")
                RX += 1
                CX = -1
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE, True)
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE, True)
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_CODE")
                CX += 1 : oSheet.Cells(RX, CX).Value = "'" & rowSOTORDR2.Item("COLOR_CODE")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_DESC")
                CX += 2
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("PF_QTY")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("PF_QTY")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowICTCOLR1.Item("COLOR_DESC") & ""

            Next





            oSheet = oWB.Worksheets("Supplier Info Sheet")
            oSheet.Cells("B8").NumberFormat = "@"
            oSheet.Cells("B8").Value = rowSOTORDR1.Item("ORDR_CUST_PO")
            oSheet.Cells("D9").NumberFormat = "@"
            oSheet.Cells("D9").Value = rowSOTORDR1.Item("ORDR_NO")

            For Each CUST_ADDR_TYPE As String In New String() {"BT", "ST"}
                Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {rowSOTORDR1.Item("ORDR_NO"), CUST_ADDR_TYPE})
                Dim CA As Integer = 6
                If CUST_ADDR_TYPE = "BT" Then
                    CA = 4
                End If
                Dim RA As Integer = 0

                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_NAME")
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR1")
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR2")
                'RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR3")
                Dim CSZ As String = rowSOTORDR5.Item("CUST_CITY") & "," & rowSOTORDR5.Item("CUST_STATE") & " " & rowSOTORDR5.Item("CUST_ZIP_CODE") & " " & rowSOTORDR5.Item("CUST_COUNTRY")
                RA += 1 : oSheet.Cells(RA, CA).Value = CSZ
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_CONTACT")
            Next


            RX = 14
            CX = 0

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "ORDR_LNO")
                RX += 1
                CX = -1
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE, True)
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_DESC")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowICTSTYL1.Item("STYLE_MATL_DESC") & ""
                CX += 1 : oSheet.Cells(RX, CX).Value = rowICTSTYL1.Item("COUNTRY_CODE") & ""
                CX += 1 : oSheet.Cells(RX, CX).Value = "?"
                CX += 1 : oSheet.Cells(RX, CX).Value = "?"
                CX += 1 : oSheet.Cells(RX, CX).Value = "?"
            Next




            oSheet = oWB.Worksheets("Commercial Invoice")
            oSheet.Cells("F3").NumberFormat = "@"
            oSheet.Cells("F3").Value = rowSOTORDR1.Item("ORDR_NO")
            oSheet.Cells("H3").NumberFormat = "MM/dd/yyyy"
            oSheet.Cells("H3").Value = rowSOTORDR1.Item("ORDR_DATE") ' *******************

            oSheet.Cells("B12").NumberFormat = "@"
            oSheet.Cells("B12").Value = rowSOTORDR1.Item("CUST_CODE")
            oSheet.Cells("D12").NumberFormat = "@"
            oSheet.Cells("D12").Value = rowSOTORDR1.Item("CUST_STORE_NO")

            oSheet.Cells("A15").NumberFormat = "@"
            oSheet.Cells("A15").Value = rowSOTORDR1.Item("ORDR_NO")
            oSheet.Cells("B15").NumberFormat = "@"
            oSheet.Cells("B15").Value = rowSOTORDR1.Item("ORDR_CUST_PO")
            oSheet.Cells("C15").NumberFormat = "@"
            oSheet.Cells("C15").Value = rowSOTORDR1.Item("ORDR_DEPT")
            oSheet.Cells("D15").NumberFormat = "@"
            oSheet.Cells("D15").Value = rowSOTORDR1.Item("SREP_CODE")
            oSheet.Cells("H15").NumberFormat = "@"
            oSheet.Cells("H15").Value = rowSOTORDR1.Item("TERM_CODE")
            oSheet.Cells("F15").NumberFormat = "@"
            oSheet.Cells("F15").Value = rowSOTORDR1.Item("SHIP_VIA_CODE")

            oSheet.Cells("I15").NumberFormat = "@"
            oSheet.Cells("I15").Value = rowSOTORDR1.Item("ORDR_FOB")
            oSheet.Cells("E15").NumberFormat = "#,##0"
            oSheet.Cells("E15").Value = Absx1.numFor("PF_CARTONS").Value
            '    oSheet.Cells("F13").NumberFormat = "#,##0.0"
            '    oSheet.Cells("F13").Value = "LB"
            '   oSheet.Cells("G13").NumberFormat = "#,##0.0"
            '  oSheet.Cells("G13").Value = "#LBS"

            oSheet.Cells("I12").NumberFormat = "@"
            oSheet.Cells("I12").Value = ""

            For Each CUST_ADDR_TYPE As String In New String() {"BT", "ST"}
                Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {rowSOTORDR1.Item("ORDR_NO"), CUST_ADDR_TYPE})
                Dim CA As Integer = 5
                If CUST_ADDR_TYPE = "BT" Then
                    CA = 0
                End If
                Dim RA As Integer = 5

                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_NAME")
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR1")
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR2")
                'RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_ADDR3")
                Dim CSZ As String = rowSOTORDR5.Item("CUST_CITY") & "," & rowSOTORDR5.Item("CUST_STATE") & " " & rowSOTORDR5.Item("CUST_ZIP_CODE") & " " & rowSOTORDR5.Item("CUST_COUNTRY")
                RA += 1 : oSheet.Cells(RA, CA).Value = CSZ
                RA += 1 : oSheet.Cells(RA, CA).Value = rowSOTORDR5.Item("CUST_CONTACT")
            Next


            RX = 16
            CX = 0

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "ORDR_LNO")
                RX += 1
                CX = -1

                If RX > 21 Then
                    oSheet.Cells(RX & ":" & RX).Insert()
                End If


                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE, True)
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE, True)
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_CODE")
                CX += 1 : oSheet.Cells(RX, CX).Value = "'" & rowSOTORDR2.Item("COLOR_CODE")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_DESC")
                CX += 2
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("PF_QTY")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_UOM")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_AMT")

            Next






            Dim SFX As String = ASCMAIN1.Next_Control_No("ExportDocuments")
            Dim XLS_FILE As String = Replace(xls_filename, "ExportDocuments", "ExportDocuments" & "_" & SFX)
            oWB.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILE, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
            oWB.Close()
            range = Nothing
            oSheet = Nothing
            oWB = Nothing
            Dim p As Process = Process.Start(ASCMAIN1.Folders("Temp") & XLS_FILE)


        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSOTORDRX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDRX.DoubleClickRow
        Absx1.txtFor("ORDR_NO").Text = e.Row.Cells("ORDR_NO").Value & ""
        Click_Command("View")
    End Sub

    Sub Setup_SOTORDR3()
        If grdSOTORDR2.ActiveRow Is Nothing OrElse (Not grdSOTORDR2.ActiveRow.IsDataRow Or grdSOTORDR2.ActiveRow.IsAddRow) Then
            grpSOTORDR3.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdSOTORDR3.DataSource, DataTable).DefaultView
            Dim ORDR_LNO As Integer = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value)
            dvw.RowFilter = "ORDR_LNO = " & CStr(ORDR_LNO)
            grdSOTORDR3.Text = "Customer Style / Color Details for Order Line " & CStr(ORDR_LNO)
            grpSOTORDR3.Visible = True
        End If

        cbeICTSIZE1.Value = DBNull.Value
    End Sub


#Region "VB6"

    Sub Set_SOTPICK2()
        'If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
        '    'grdSOTPICK2.Visible = False
        'Else
        '    Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
        '    Fill_Records("SOTPICK2", PICK_NO)
        '    'Sort_grdColumns(grdSOTPICK2, "PICK_LNO")
        '    'grdSOTPICK2.Text = "Pick Ticket Details for Pick Ticket " & PICK_NO
        '    'grdSOTPICK2.Visible = True
        'End If
    End Sub

    Sub Set_SOTCART1()
        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTCART1.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value

            EnforceConstraints(False)
            dst.Tables("SOTCART2").Rows.Clear()
            dst.Tables("SOTCART1").Rows.Clear()
            Fill_Records("SOTCART1", PICK_NO)
            Fill_Records("SOTCART2", PICK_NO)

            For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select("")
                Dim ORDR_NO As String = rowSOTORDRR.Item("ORDR_NO")
                Dim RANGE_STYLE_LNO As Integer = Val(rowSOTORDRR.Item("RANGE_STYLE_LNO") & "")
                Dim ORDR_LNO As Integer = Val(rowSOTORDRR.Item("ORDR_LNO") & "")
                Dim sql As String = "ORDR_LNO = " & CStr(RANGE_STYLE_LNO)
                For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select(sql)
                    rowSOTCART2.Item("ORDR_LNO") = ORDR_LNO
                Next
            Next

            EnforceConstraints(True)

            Sort_grdColumns(grdSOTCART1, "CART_NO")
            Set_SOTCART2()
            grdSOTCART1.Text = "Cartons for Pick Ticket " & PICK_NO
            grdSOTCART1.Visible = True
        End If
    End Sub

    Sub Set_SOTCART2()
        'If grdSOTCART1.ActiveRow Is Nothing OrElse Not grdSOTCART1.ActiveRow.IsDataRow Then
        '    'grdSOTCART2.Visible = False
        'Else
        '    Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value
        '    Fill_Records("SOTCART2", CART_NO)
        '    'Sort_grdColumns(grdSOTCART2, "CART_LNO")
        '    'grdSOTCART2.Text = "Cartons Details for Carton " & CART_NO
        '    'grdSOTCART2.Visible = True
        'End If
    End Sub

    Sub Update_SOTORDR5 _
    (ORDR_NO As String, ORDR_ADDR_TYPE_ST As String, CUST_DC_NO As String, CUST_STORE_NO As String)

        Dim CUST_NAME As String
        Dim CUST_EMAIL As String
        Dim CUST_ADDR_TYPE As String
        Dim CUST_ADDR_CODE As String

        Dim rowC As DataRow

        If ORDR_ADDR_TYPE_ST = "BT" Then
            Dim CUST_ADDR_CODE_BT As String = txtCUST_ADDR_CODE_BT.Text
            Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
            Dim rowARTCUST2_BT As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, ORDR_ADDR_TYPE_ST, CUST_ADDR_CODE_BT})
            If rowARTCUST2_BT IsNot Nothing Then
                rowC = rowARTCUST2_BT
            Else
                rowC = rowARTCUST1_BT
            End If

            CUST_NAME = rowC.Item("CUST_NAME") & ""
            CUST_EMAIL = rowC.Item("CUST_EMAIL") & ""
            CUST_ADDR_TYPE = "BT"
            CUST_ADDR_CODE = CUST_DC_NO
            If ASCMAIN1.Running_in_VS Then Stop ' WHY DO WE PUT CUST_DC_NO INTO CUST_ADDR_CODE FOR THE BT ADDRESS RECORD?
            ' THIS FIELD SHOULD CONTAIN NULL OR ALL 0'S
        Else
            Dim CUST_NAME_default As String = rowSOTORDR1.Item("CUST_NAME") & " "
            If ORDR_ADDR_TYPE_ST = "DC" Then
                CUST_ADDR_CODE = CUST_DC_NO
                CUST_NAME_default &= "DC #" & CUST_DC_NO
            Else
                CUST_ADDR_CODE = CUST_STORE_NO
                CUST_NAME_default &= "#" & CUST_STORE_NO
            End If

            rowC = LookUp("ARTCUST2", New String() {CUST_CODE, ORDR_ADDR_TYPE_ST, CUST_ADDR_CODE})
            Dim rowARTCUST2_DC As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "DC", CUST_DC_NO})

            If ORDR_ADDR_TYPE_ST = "DC" Then
                If rowARTCUST2_DC IsNot Nothing AndAlso rowARTCUST2_DC.Item("CUST_NAME") & "" <> "" Then
                    CUST_NAME = rowARTCUST2_DC.Item("CUST_NAME")
                Else
                    CUST_NAME = CUST_NAME_default
                End If
            Else
                If rowC IsNot Nothing AndAlso rowC.Item("CUST_NAME") & "" <> "" Then
                    CUST_NAME = rowC.Item("CUST_NAME")
                Else
                    CUST_NAME = CUST_NAME_default
                End If
                'CUST_NAME = CUST_NAME_default
            End If

            CUST_EMAIL = ""

            CUST_ADDR_TYPE = "ST"
            'CUST_ADDR_CODE = rowARTCUST2_DC.Item("CUST_ADDR_CODE").Value
        End If

        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, CUST_ADDR_TYPE})
        If rowSOTORDR5 Is Nothing Then
            rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
            rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
            rowSOTORDR5.Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
            dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
        End If
        With rowSOTORDR5
            .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
            .Item("CUST_NAME") = CUST_NAME
            .Item("CUST_EMAIL") = CUST_EMAIL
            For Each COLUMN_NAME As String In New String() _
                {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE",
                 "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX"}
                .Item(COLUMN_NAME) = rowC.Item(COLUMN_NAME)
            Next
        End With
        ' dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
    End Sub

    Sub Cancel_Order(Optional KeepPO As Boolean = False)
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        restore_reservation = False

        Dim EMsg As String

        If (EntryMode = "M" Or multiple_order_maintenance) Then
            ASCMAIN1.sql = "Select ORDR_NO from SOTORDR1 " & vbCrLf _
               & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
               & IIf(multiple_order_type = "ORDR_GROUP_NO",
                     " and ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'",
                     " and EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") & vbCrLf _
               & " and ORDR_STATUS = 'O'"
            Dim dt As DataTable = ASCDATA1.GetDataTable
            For Each row As DataRow In dt.Rows
                Cancel_Order_1(row.Item("ORDR_NO"))
            Next
            EMsg = CStr(dt.Rows.Count) & IIf(multiple_order_type = "ORDR_GROUP_NO",
                                             " Orders from Order Group " & ORDR_GROUP_NO,
                                             " Orders from EDI Journal " & EDI_JRNL_NO) & " have been Cancelled"
        Else
            Cancel_Order_1(ORDR_NO)
            EMsg = "Order " & ORDR_NO & " has been Cancelled"
        End If

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        Next

        Cancel_POs(True)

        If ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "RGI" Then
            If (ASCMAIN1.CLIENT = "RGI" And CUST_CODE = "031013" And Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR") Then
                'Wafyair xfr orders come in as 855 can't reply to an 855 with another 855
            ElseIf optORDR_SOURCE.Value = "E" Then
                ASCMAIN1.sql = "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '855' AND CUST_CODE = '" & CUST_CODE & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing Then
                    TAC.EDC855O1.Generate_855(clsASCBASE1, ORDR_GROUP_NO)
                End If
            End If
        End If

        If KeepPO Then
            ' wipe out POs association with Sales Order
            ' change PO whse & restore to Open

            For Each PO_ORDER_NO As String In PO_ORDER_NOs
                ASCMAIN1.sql = "Update POTORDR1 Set ORDR_NO = NULL, CUST_CODE = NULL" _
                    & ", FOB_CMT = 'F', PO_STATUS = 'O', WHSE_CODE = '" & ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & "'" _
                    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update POTORDR2 Set ORDR_NO = NULL, ORDR_LNO = NULL" _
                    & ", PO_STATUS = 'O', PO_QTY_OPN = PO_QTY_ORD" _
                    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                ASCDATA1.ExecuteSQL()

                TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID,
                                          "CXLKPO",
                                          "Sales Order " & ORDR_NO & " Cancelled, PO Kept Open",
                                          ORDR_NO, "POFORDR1")

                TAC.POCMAIN1.Dependent_Updates(1, PO_ORDER_NO)
            Next
        End If

        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Cancel_Order_1(ORDR_NO As String)
        Dependent_Updates(-1, ORDR_NO)

        ASCMAIN1.sql = "Select Sum (ORDR_QTY_PICK) from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        Dim ORDR_STATUS As String = ""
        If Val(ASCDATA1.GetDataValue) <> 0 Then
            ORDR_STATUS = "P"
        Else
            Dim rowSOTORDR1_cancel = LookUp("SOTORDR1", ORDR_NO)
            If Val(rowSOTORDR1_cancel.ITEM("ORDR_PICK_SEQ") & "") = 0 Then
                ORDR_STATUS = "C"
            Else
                ORDR_STATUS = "F"
            End If
        End If

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2" _
            & "    Set ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + NVL(R1.ORDR_QTY_OPEN,0)" _
            & "      , ORDR_QTY_OPEN = 0" _
            & "    where Current of C1;" _
            & "   Update SOTORDR2" _
            & "    Set ORDR_STATUS = CASE WHEN ORDR_QTY_PICK <> 0 THEN 'P' ELSE CASE WHEN ORDR_QTY_SHIP <> 0 THEN 'F' ELSE 'C' END END" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "ORDCXL", "Order Cancelled")

        ASCMAIN1.sql = "Update SOTORDR1 Set REASON_CODE = :PARM1" _
            & ", ORDR_STATUS = :PARM2, ORDR_DATE_CLOSED = TRUNC(SYSDATE), ORDR_YYYYPP_CLOSED = :PARM3" _
            & " where ORDR_NO = :PARM4"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {Absx1.txtFor("REASON_CODE").Text, ORDR_STATUS, ASCMAIN1.CYP, ORDR_NO})
    End Sub

    Sub Reverse_Cancel(ORDR_NO As String, ORDR_GROUP_NO As String, Optional dt As DataTable = Nothing)
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        restore_reservation = False
        Dim EMsg As String

        If ORDR_NO <> "" Then
            Reverse_Cancel_1(ORDR_NO)
            EMsg = "Order " & ORDR_NO & " has been Re-Opened"
        Else
            If ASCMAIN1.Running_in_VS Then Stop ' do we look at a selected indicator?
            For Each row As DataRow In dt.Rows
                Reverse_Cancel_1(row.Item("ORDR_NO"))
            Next
            EMsg = CStr(dt.Rows.Count) & " Orders from Order Group " & ORDR_GROUP_NO & " have been Re-Opened"
        End If

        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        CommitTrans("Process Complete")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Reverse_Cancel_1(ORDR_NO As String)
        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is " _
            & "  Select * from SOTORDR2" _
            & "   where ORDR_NO = '" & ORDR_NO & "'" _
            & "     and ORDR_STATUS = 'C'" _
            & "     and NVL(ORDR_QTY_CANC,0) <> 0 for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2" _
            & "    Set ORDR_QTY_CANC = 0" _
            & "      , ORDR_QTY_OPEN = NVL(R1.ORDR_QTY_CANC,0), ORDR_STATUS = 'O'" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR1 Set REASON_CODE = Null" _
            & ", ORDR_STATUS = 'O', ORDR_DATE_CLOSED = Null, ORDR_YYYYPP_CLOSED = Null" _
            & " where ORDR_NO = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {ORDR_NO})

        Dim row As DataRow = LookUp("SOTORDR1", ORDR_NO)
        Dim ORDR_GROUP_NO_X As String = row.Item("ORDR_GROUP_NO")
        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO_X}, New String() {"ORDR_GROUP_NO_IN"})

        Dependent_Updates(1, ORDR_NO)
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        restore_reservation = False
        Dim EMsg As String

        If (EntryMode = "M" Or multiple_order_maintenance) Then
            ASCMAIN1.sql = "Select ORDR_NO from SOTORDR1 " & vbCrLf _
               & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
               & IIf(multiple_order_type = "ORDR_GROUP_NO",
                     " and ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'",
                     " and EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") & vbCrLf _
               & " and ORDR_STATUS = 'O'"
            Dim dt As DataTable = ASCDATA1.GetDataTable
            For Each row As DataRow In dt.Rows
                Delete_Order_1(row.Item("ORDR_NO"))
            Next
            EMsg = CStr(dt.Rows.Count) & IIf(multiple_order_type = "ORDR_GROUP_NO",
                                  " Orders from Order Group " & ORDR_GROUP_NO,
                                  " Orders from EDI Journal " & EDI_JRNL_NO) & " have been marked as Deleted"
        Else
            Delete_Order_1(ORDR_NO)
            EMsg = "Order " & ORDR_NO & " has been marked as Deleted"
        End If

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        Next

        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(ORDR_NO As String)
        Dependent_Updates(-1, ORDR_NO)

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 Is Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2" _
            & "    Set ORDR_QTY_OPEN = 0, ORDR_STATUS = '" & "D" & "'" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR1 Set REASON_CODE = :PARM1" _
            & ", ORDR_STATUS = :PARM2, ORDR_DATE_CLOSED = TRUNC(SYSDATE), ORDR_YYYYPP_CLOSED = :PARM3" _
            & " where ORDR_NO = :PARM4"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {Absx1.txtFor("REASON_CODE").Text, "D", ASCMAIN1.CYP, ORDR_NO})
    End Sub

    Sub Dependent_Updates(S As Integer, ORDR_NO As String)

        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")
            Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")

            If S = -1 Then
                If rowSOTORDR2.Item("RSRV_NO") & "" <> "" And restore_reservation Then
                    'Only restore this reservation line if it hasn't been substitutioned.  Per Gabe 07/30/02 - WR.
                    Dim row As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, rowSOTORDR2.Item("ORDR_LNO")})
                    If row IsNot Nothing Then  'Added for Angela. 1/24/05.  She was adding styles to range that had pulled from reservation already.
                        If row.Item("STYLE_CODE_SUB") & "" = "" Then
                            Update_SOTRSRVx(rowSOTORDR2, S)
                        End If
                    End If
                End If
            Else
                Dim rowSOTRSRVX As DataRow = Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                '& " order by SOTRSRV1.ORDR_CANCEL_DATE"

                Dim Ps() As Object

                If rowSOTRSRVX IsNot Nothing Then
                    rowSOTORDR2.Item("RSRV_NO") = rowSOTRSRVX.Item("RSRV_NO")
                    rowSOTORDR2.Item("RSRV_LNO") = rowSOTRSRVX.Item("RSRV_LNO")
                    Ps = {rowSOTRSRVX.Item("RSRV_NO"), rowSOTRSRVX.Item("RSRV_LNO")}
                    Update_SOTRSRVx(rowSOTORDR2, S)
                Else
                    rowSOTORDR2.Item("RSRV_NO") = DBNull.Value
                    rowSOTORDR2.Item("RSRV_LNO") = DBNull.Value
                    Ps = {DBNull.Value, DBNull.Value}
                End If

                'Update_Record_TDA("SOTORDR2")
                ASCMAIN1.sql = "Update SOTORDR2 Set RSRV_NO = :PARM1, RSRV_LNO = :PARM2" _
                    & " where ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VN", Ps)
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", S * QTY_TO_COMMIT)
            End If
        Next

    End Sub

    Sub Update_SOTRSRVx(rowSOTORDR2 As DataRow, S As Integer)
        Dim RSRV_NO As String = rowSOTORDR2.Item("RSRV_NO") & ""
        Dim RSRV_LNO As Int64 = Val(rowSOTORDR2.Item("RSRV_LNO") & "")

        Dim rowSOTRSRV1 As DataRow = Fill_Record("SOTRSRV1", RSRV_NO)
        Dim WHSE_CODE As String = rowSOTRSRV1.Item("WHSE_CODE")

        Dim rowSOTRSRV2 As DataRow = Fill_Record("SOTRSRV2", New String() {RSRV_NO, RSRV_LNO})
        With rowSOTRSRV2
            Dim RSRV_QTY As Int64 = .Item("RSRV_QTY")
            Dim RSRV_QTY_OPEN As Int64 = Val(.Item("RSRV_QTY_OPEN") & "")
            Dim RSRV_QTY_CANC As Int64 = Val(.Item("RSRV_QTY_CANC") & "")
            Dim RSRV_QTY_USED As Int64 = Val(.Item("RSRV_QTY_USED") & "") _
                          + S * Val(rowSOTORDR2.Item("ORDR_QTY") & "")

            '  + S * Val(rowSOTORDR2.Item("ORDR_QTY_ORIG") & "") - USING ORDR_QTY_ORIG WILL ALWAYS HAVE 0 IMPACT WHEN CHANGING THE ORDER
            Dim RSRV_QTY_OPEN_OLD As Int64 = RSRV_QTY_OPEN
            RSRV_QTY_OPEN = RSRV_QTY - RSRV_QTY_CANC - RSRV_QTY_USED
            If RSRV_QTY_OPEN < 0 Then
                RSRV_QTY_OPEN = 0
            End If
            Dim RSRV_QTY_OPEN_NEW As Int64 = RSRV_QTY_OPEN
            .Item("RSRV_QTY_USED") = RSRV_QTY_USED
            .Item("RSRV_QTY_OPEN") = RSRV_QTY_OPEN

            Dim QTY_TO_COMMIT As Int64 = RSRV_QTY_OPEN_NEW - RSRV_QTY_OPEN_OLD
            If QTY_TO_COMMIT <> 0 Then
                Dim STYLE_CODE As String = .Item("STYLE_CODE")
                Dim COLOR_CODE As String = .Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY_TO_COMMIT)
            End If

        End With

        If S = -1 Then
        Else

            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim rowSOTORDR7 As DataRow = Fill_Record("SOTORDR7", New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})

            If rowSOTORDR7 Is Nothing Then
                rowSOTORDR7 = dst.Tables("SOTORDR7").NewRow
                rowSOTORDR7.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                rowSOTORDR7.Item("STYLE_CODE") = STYLE_CODE
                rowSOTORDR7.Item("COLOR_CODE") = COLOR_CODE
                dst.Tables("SOTORDR7").Rows.Add(rowSOTORDR7)
            End If
            If rowSOTRSRV2.Item("RSRV_PRIORITY_DATE") & "" = "" Then
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV1.Item("INIT_DATE")).Date ' DateValue(Format(rowSOTRSRV1.Item("INIT_DATE"), "MM/dd/yyyy"))
            Else
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV2.Item("RSRV_PRIORITY_DATE")).Date '  DateValue(Format$(rowSOTRSRV2.Item("RSRV_PRIORITY_DATE"), "MM/DD/YYYY"))
            End If
            rowSOTORDR7.Item("ORDR_PRIORITY") = rowSOTRSRV2.Item("RSRV_PRIORITY")
            Update_Record_TDA("SOTORDR7", "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
        End If
        Update_Record_TDA("SOTRSRV2")

        ASCMAIN1.sql = "Select Sum (RSRV_QTY_OPEN) from SOTRSRV2 where RSRV_NO = :PARM1"
        Dim RSRV_QTY_OPEN_total As Int64 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {RSRV_NO}))

        If RSRV_QTY_OPEN_total = 0 Then
            rowSOTRSRV1.Item("RSRV_STATUS") = "F"
        Else
            rowSOTRSRV1.Item("RSRV_STATUS") = "O"
        End If
        Update_Record_TDA("SOTRSRV1")
    End Sub

    Sub Display_Totals()

        Dim KEY As Int32 = 0
        For Each SFX As String In New String() {"", "OPEN", "ALLO", "PICK", "SHIP", "CANC"}
            If SFX <> "" Then SFX = "_" & SFX
            KEY += 1
            Dim rowSOTORDRT As DataRow = dst.Tables("SOTORDRT").Rows.Find(KEY)
            rowSOTORDRT.Item("QTY") = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY" & SFX & ")", "") & "")
            rowSOTORDRT.Item("AMT") = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT" & SFX & ")", "") & "")
        Next

        If multi_store_is_active Then
            Set_MS_TOTAL_AMT_Expression()
        End If
    End Sub

    Sub Display_Totals_R(ORDR_LNO As Int32)

        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
        Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)

        Select Case RANGE_TYPE
            Case Is = "R"
                rowSOTORDR2.Item("ORDR_QTY") = Val(dst.Tables("SOTORDRR").Compute("SUM(ORDR_QTY)", sqlw) & "")
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = Val(dst.Tables("SOTORDRR").Compute("SUM(ORDR_QTY_OPEN)", sqlw) & "")
                grdSOTORDRR.Text = "Range Style Components for Line " & CStr(ORDR_LNO)

            Case Is = "A"
                Dim ORDR_AMT As Decimal = Val(dst.Tables("SOTORDRR").Compute("SUM(ORDR_AMT)", sqlw) & "")
                Dim ORDR_QTY As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = ORDR_AMT / ORDR_QTY
                ' Sql = "Select SUM (SOWORDRR.ORDR_UNIT_PRICE * SOWORDRR.QTY_PER_PP) as ORDR_UNIT_PRICE" & vbCrLf
                grdSOTORDRR.Text = "Assortment Style Components for Line " & CStr(ORDR_LNO)
                If ASCMAIN1.Running_in_VS Then Stop ' grdSOWORDR2_AfterColUpdate(grdSOWORDR2.Columns("ORDR_QTY").position)

            Case Else
        End Select
    End Sub

    Function Select_Style(ByRef COLOR_CODE As String) As String

        Dim STYLE_CODE As String = ""

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            STYLE_CODE = ASCMAIN1.CodeSelector.SelectedCode
        End If

        If COLOR_CODE <> "" Then
            If STYLE_CODE <> "" Then
                Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 Is Nothing Then
                    MsgBox("Color Code '" & COLOR_CODE & "' is not Associated with Style " & STYLE_CODE)
                    STYLE_CODE = ""
                End If
            End If
        Else
            ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = :PARM1"
            Dim rows() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {STYLE_CODE}).Select
            If rows.Length = 1 Then
                COLOR_CODE = rows(0).Item("COLOR_CODE")
            Else
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("COLOR_CODE")
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    ASCMAIN1.CodeSelector.SQL = "Select * from (" & ASCMAIN1.CodeSelector.SQL & ")" _
                        & " where COLOR_CODE in " _
                        & " (Select COLOR_CODE from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "')"
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    COLOR_CODE = ASCMAIN1.CodeSelector.SelectedCode
                    If COLOR_CODE = "" Then STYLE_CODE = ""
                End If
            End If
        End If

        Return STYLE_CODE
    End Function

    Private Sub cmdAddLine_Click()
        If multi_store_is_active Then
            Exit Sub
        End If

        Dim ORDR_LNO As Int32 = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value)
        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
        If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" Then
            If Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "") = 0 Then
                MsgBox("Qty Open is Zero", MsgBoxStyle.OkOnly, "Cannot Sub a Line w/0 Qty Open")
                Exit Sub
            End If
            If MsgBox("Please Note: When Adding a Sub to a Range-Style," _
                      & vbCrLf & "   you will be prompted for a Style to Substitute for the Selected Component" _
                      & vbCrLf & vbCrLf & "   Style: " & grdSOTORDRR.ActiveRow.Cells("STYLE_CODE").Value & ", " & grdSOTORDRR.ActiveRow.Cells("ORDR_QTY_OPEN").Value & " units" _
                      & vbCrLf & vbCrLf & "Continue w/Substitution",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                      "Verify to Continue") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        Dim COLOR_CODE As String = ""
        Dim STYLE_CODE As String = Select_Style(COLOR_CODE)
        If STYLE_CODE = "" Then
            Exit Sub
        End If
        Dim STYLE_CODE_SUB As String = STYLE_CODE
        Dim COLOR_CODE_SUB As String = COLOR_CODE
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_SUB)

        If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" Then
            Add_Range_Style_Component(STYLE_CODE_SUB, COLOR_CODE_SUB)
            Exit Sub
        End If

        Dim rowSOTORDR2_SUB As DataRow = dst.Tables("SOTORDR2").NewRow
        With rowSOTORDR2_SUB
            For i As Integer = 0 To dst.Tables("SOTORDR2").Columns.Count - 1
                rowSOTORDR2_SUB.Item(i) = rowSOTORDR2.Item(i)
            Next i
            Dim ORDR_LNO_ctr As Int32 = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
            .Item("ORDR_LNO") = ORDR_LNO_ctr
            .Item("ORDR_QTY") = 0
            .Item("ORDR_QTY_ORIG") = 0
            .Item("ORDR_QTY_ALLO") = 0
            .Item("ORDR_QTY_OPEN") = 0
            .Item("ORDR_QTY_PICK") = 0
            .Item("ORDR_QTY_SHIP") = 0
            .Item("ORDR_QTY_CANC") = 0
            .Item("RSRV_NO") = DBNull.Value
            .Item("RSRV_LNO") = DBNull.Value
        End With
        dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2_SUB)

        grdSOTORDR2.ActiveRow = grdSOTORDR2.Rows.GetItem(grdSOTORDR2.Rows.Count - 1)
        'datSOWORDR2.Recordset.MoveLast()

    End Sub

    Sub Add_Range_Style_Component(STYLE_CODE_SUB As String, COLOR_CODE_SUB As String)
        If grdSOTORDRR.ActiveRow Is Nothing Then Exit Sub

        Dim ORDR_LNO As Int32 = Val(grdSOTORDRR.ActiveRow.Cells("ORDR_LNO").Value & "")
        Dim RANGE_STYLE_LNO As Int32 = Val(grdSOTORDRR.ActiveRow.Cells("RANGE_STYLE_LNO").Value & "")
        Dim rowSOTORDRR As DataRow = dst.Tables("SOTORDRR").Rows.Find _
                                     (New Object() {ORDR_NO, RANGE_STYLE_LNO})

        Dim rowSOTORDRR_ADD As DataRow = dst.Tables("SOTORDRR").NewRow
        With rowSOTORDRR_ADD
            For i As Integer = 0 To dst.Tables("SOTORDRR").Columns.Count - 1
                .Item(i) = rowSOTORDRR.Item(i)
            Next i
            .Item("ORDR_QTY_OPEN") = 0
            .Item("ORDR_QTY_CANC") = Val(.Item("ORDR_QTY") & "") _
                                   - Val(.Item("ORDR_QTY_SHIP") & "") _
                                   - Val(.Item("ORDR_QTY_OPEN") & "") _
                                   - Val(.Item("ORDR_QTY_PICK") & "")
            If Val(.Item("ORDR_QTY_CANC") & "") < 0 Then
                .Item("ORDR_QTY_CANC") = 0
            End If
        End With
        dst.Tables("SOTORDRR").Rows.Add(rowSOTORDRR_ADD)

        With rowSOTORDRR_ADD
            .Item("RANGE_STYLE_LNO") = Val(dst.Tables("SOTORDRR").Compute("MAX(RANGE_STYLE_LNO)", "") & "") + 1
            .Item("ORDR_QTY") = 0
            .Item("ORDR_QTY_ORIG") = 0
            .Item("ORDR_QTY_ALLO") = 0
            '.ITEM("ORDR_QTY_OPEN") = 0
            .Item("ORDR_QTY_PICK") = 0
            .Item("ORDR_QTY_SHIP") = 0
            .Item("ORDR_QTY_CANC") = 0
            .Item("RSRV_NO") = ""
            .Item("RSRV_LNO") = 0
            .Item("ORDR_STATUS") = ""
            .Item("STYLE_CODE_SUB") = ""
            .Item("ORDR_RELEASE") = ""
            .Item("ORDR_RELEASE_AVAIL") = 0
            .Item("ORDR_EXTD_COST") = 0
            .Item("ORDR_QTY_PRE_ALLO") = 0
            .Item("STYLE_CODE_SUB") = .Item("STYLE_CODE")
            .Item("STYLE_CODE") = STYLE_CODE_SUB
            .Item("COLOR_CODE") = COLOR_CODE_SUB

            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", .Item("STYLE_CODE"))
            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC") & ""
            .Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY") & ""
            .Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY") & ""
            .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM") & ""
        End With

        Display_Totals_R(ORDR_LNO)

    End Sub

    Private Sub Set_Range_Style_PPK()
        Dim MSG As String = "Are You Sure You Want To Fix Order Number " & Absx1.txtFor("ORDR_NO").Text & " ?"
        If MsgBox(MSG, MsgBoxStyle.YesNo, "Thought I would Double Check..") = MsgBoxResult.Yes Then
            MSG = "Order " & ORDR_NO & " Has Been Fixed." & vbCrLf _
                & "You Will Need to Exit The Order Entry Screen" & vbCrLf _
                & "By Hitting The Cancel Button.  When You" & vbCrLf _
                & "Come Back In Life Should Be Good Again."
            ASCMAIN1.sql = "UPDATE SOTORDR9" & vbCrLf _
                & " Set RANGE_STYLE_QTY_PER_PP = 1" & vbCrLf _
                & " WHERE ORDR_NO = '" & ORDR_NO & "'"
            ASCDATA1.ExecuteSQL()
            MsgBox(MSG, MsgBoxStyle.Information, "Bug Squished")
        End If
    End Sub

    Sub Setup_MS(tf As Boolean)

        UltraExplorerBar1.Groups("Multi-Store").Items("Start Multi-Store").Text = IIf(tf, "Clear Multi-Store", "Start Multi-Store")

        tabMain.Tabs("Multi-Store").Visible = tf
        ' UltraExplorerBar1.Groups("Multi-Store").Visible = tf
        lblMultiStore.Visible = tf


        multi_store_is_active = tf

        msqty = 0
        msqty_col = 0

        If multi_store_is_active Then
            Init_MultiStore()

            With UltraExplorerBar1.Groups("Multi-Store")
                .Items("Reset Qty's").Settings.Enabled = DefaultableBoolean.True
                .Items("Re-Select Stores").Settings.Enabled = DefaultableBoolean.True
                .Items("Clear Zeroes").Settings.Enabled = DefaultableBoolean.True
            End With

        Else
            CUST_STORE_NOs_multi_store.Clear()

            With UltraExplorerBar1.Groups("Multi-Store")
                .Items("Reset Qty's").Settings.Enabled = DefaultableBoolean.False
                .Items("Re-Select Stores").Settings.Enabled = DefaultableBoolean.False
                .Items("Clear Zeroes").Settings.Enabled = DefaultableBoolean.False
            End With
        End If
    End Sub

    Sub Init_MultiStore()
        Me.Cursor = Cursors.WaitCursor

        dst.Tables("SOTORDRS").Rows.Clear()

        ' Add a column for each style referenced in grdSOWORDR2
        dst.Tables("SOTORDRS").Columns("TOTAL_QTY").Expression = ""
        dst.Tables("SOTORDRS").Columns("TOTAL_AMT").Expression = ""
        For i As Integer = dst.Tables("SOTORDRS").Columns.Count - 1 To 0 Step -1
            Dim DC As DataColumn = dst.Tables("SOTORDRS").Columns(i)
            If DC.ColumnName = "TOTAL_AMT" Or DC.ColumnName = "TOTAL_QTY" Then
                Exit For
            Else
                Dim summary As UltraWinGrid.SummarySettings = grdSOTORDRS.DisplayLayout.Bands(0).Summaries(DC.ColumnName)
                grdSOTORDRS.DisplayLayout.Bands(0).Summaries.Remove(summary)

                dst.Tables("SOTORDRS").Columns.Remove(DC)
            End If
        Next

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select()
            Fill_Styles(rowSOTORDR2.Item("ORDR_LNO"), rowSOTORDR2.Item("RANGE_STYLE_CODE") & "", rowSOTORDR2.Item("STYLE_CODE") & "", rowSOTORDR2.Item("COLOR_CODE") & "", False)
        Next
        Set_MS_TOTAL_QTY_Expression()

        ' Add a row for each store referenced in CUST_STORE_NOs_multi_store
        Fill_Stores(CUST_STORE_NOs_multi_store)

        tabMain.SelectedTab = tabMain.Tabs("Multi-Store")

        lblMSCopyToStore.Visible = False
        txtMSCopyToStore.Visible = False
        optMSCopyToStore.Visible = False
        optMSCopyToStore.Tag = ""
        lblMSCopyToStore.Tag = ""

        Me.Cursor = Cursors.Default
    End Sub

    Sub Fill_Styles(ORDR_LNO As Int64, RANGE_STYLE_CODE As String, STYLE_CODE As String, COLOR_CODE As String,
                    Optional reset_TOTAL_QTY As Boolean = True)

        Dim COLUMN_NAME As String = "QTY_" & Format(ORDR_LNO, "000")
        dst.Tables("SOTORDRS").Columns.Add(COLUMN_NAME, GetType(System.Int64))

        With grdSOTORDRS.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
            .Hidden = False
            .Width = grdSOTORDRS.DisplayLayout.Bands(0).Columns("TOTAL_QTY").Width
            .Format = "#,##0"
            .CellAppearance.TextHAlign = HAlign.Right
            .Header.Appearance.TextHAlign = HAlign.Right
            .Header.Caption = IIf(RANGE_STYLE_CODE <> "", RANGE_STYLE_CODE, STYLE_CODE & vbCrLf & COLOR_CODE)
            Create_Summary(grdSOTORDRS, COLUMN_NAME)
            .CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        If reset_TOTAL_QTY Then Set_MS_TOTAL_QTY_Expression(ORDR_LNO)
    End Sub

    Private Sub ReSelect_Stores()
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_STORE_NO")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True

            ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "'" '  and CUST_ADDR_STATUS = 'A'"
            Dim TBL As DataTable = ASCDATA1.GetDataTable
            ASCMAIN1.CodeSelector.UseDataFromTable = TBL

            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Join(CUST_STORE_NOs_multi_store.ToArray, Chr(0))
            ' If ASCMAIN1.Running_in_VS Then Stop ' Sql = Sql & " and CUST_ADDR_STATUS = 'A'"
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            If ASCMAIN1.CodeSelector.SelectedCodes.Count <> 0 Then
                CUST_STORE_NOs_multi_store.Clear()
                For Each CC As String In ASCMAIN1.CodeSelector.SelectedCodes
                    CUST_STORE_NOs_multi_store.Add(CC)
                Next
                If Not CUST_STORE_NOs_multi_store.Contains(Absx1.txtFor("CUST_STORE_NO").Text) Then
                    CUST_STORE_NOs_multi_store.Add(Absx1.txtFor("CUST_STORE_NO").Text)
                End If

                Dim C As New List(Of String)
                For Each CUST_STORE_NO As String In CUST_STORE_NOs_multi_store
                    C.Add(CUST_STORE_NO)
                Next

                If multi_store_is_active Then ' Delete and Add Stores to table as required
                    If dst.Tables("SOTORDRS").Rows.Count > 0 Then
                        For i As Integer = dst.Tables("SOTORDRS").Rows.Count - 1 To 0
                            Dim row As DataRow = dst.Tables("SOTORDRS").Rows(i)
                            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                            If Not CUST_STORE_NOs_multi_store.Contains(CUST_STORE_NO) Then
                                row.Delete()
                            Else
                                C.Remove(CUST_STORE_NO)
                            End If
                        Next
                    End If
                    If C.Count <> 0 Then
                        Fill_Stores(C)
                    End If
                Else
                    Setup_MS(True)
                End If
            End If
        End If
    End Sub

    Sub Fill_Stores(ByVal CUST_STORE_NOs As List(Of String))
        Dim QTYs As New Dictionary(Of Int64, Int64)
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            QTYs.Add(Val(rowSOTORDR2.Item("ORDR_LNO") & ""), Val(rowSOTORDR2.Item("ORDR_QTY") & ""))
        Next

        multi_store_changes_made_to_SOTORDRS = True

        For Each CUST_STORE_NO As String In CUST_STORE_NOs
            Dim rowSOTORDRS As DataRow = dst.Tables("SOTORDRS").Rows.Find(CUST_STORE_NO)
            If rowSOTORDRS Is Nothing Then
                rowSOTORDRS = dst.Tables("SOTORDRS").NewRow
                rowSOTORDRS.Item("CUST_STORE_NO") = CUST_STORE_NO
                rowSOTORDRS.Item("ORDR_CUST_PO") = Absx1.txtFor("ORDR_CUST_PO").Text
                For Each ORDR_LNO As Int64 In QTYs.Keys
                    If QTYs(ORDR_LNO) <> 0 Then
                        rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) = QTYs(ORDR_LNO)
                    End If
                Next
                dst.Tables("SOTORDRS").Rows.Add(rowSOTORDRS)
            End If
        Next

        With dst.Tables("SOTORDRS")
            .AcceptChanges()
            If .Rows.Count > 0 Then
                For R As Integer = .Rows.Count - 1 To 0 Step -1
                    Dim CUST_STORE_NO As String = .Rows(R).Item("CUST_STORE_NO")
                    If Not CUST_STORE_NOs.Contains(CUST_STORE_NO) Then
                        .Rows.Remove(.Rows(R))
                    End If
                Next
            End If
            .AcceptChanges()
        End With


        multi_store_changes_made_to_SOTORDRS = False

        Sort_grdColumns(grdSOTORDRS, "CUST_STORE_NO")
        Set_MS_TOTAL_AMT_Expression()

    End Sub

    Private Sub cmdSizes_Click()

        Dim SIZE_DESCs As String = ""
        For i As Integer = 1 To 12
            Dim SIZE_DESC As String = grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(i, "00")).Value & ""
            If SIZE_DESC <> "" Then
                SIZE_DESCs &= Chr(0) & SIZE_DESC
            Else
                Exit For
            End If
        Next i

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("RANGE_STYLE_CODE", "ICTRSTY1")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = SIZE_DESCs
            ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Add("CUST_CODE", Absx1.txtFor("CUST_CODE").Text)
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Dim SIZE_DESCs_new As String = ""
                For Each SIZE_DESC As String In ASCMAIN1.CodeSelector.SelectedCodes
                    If SIZE_DESC <> "" Then
                        SIZE_DESCs_new &= Chr(0) & SIZE_DESC
                    Else
                        Exit For
                    End If
                Next

                If SIZE_DESCs <> SIZE_DESCs_new Then
                    Dim SDs() As String = Split(Mid(SIZE_DESCs_new, 2), Chr(0))
                    For i As Integer = 1 To 12
                        Dim SIZE_DESC As String = ""
                        If SDs.Length >= i Then
                            SIZE_DESC = SDs(i - 1)
                        End If
                        grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(i, "00")).Value = SIZE_DESC
                    Next i
                    Setup_Sizes()
                End If
            End If
        End If
    End Sub

    Sub Setup_Sizes()
        With grdSOTORDR3.DisplayLayout.Bands(0)
            For i As Integer = 1 To 12
                Dim z As String = grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(i, "00")).Value & ""
                .Columns("SIZE_QTY_" & Format$(i, "00")).Header.Caption = z
                If z <> "" Then
                    .Columns("SIZE_QTY_" & Format(i, "00")).Hidden = False
                    ' .Columns("SIZE_QTY_" & Format(i, "00")).Width = .Columns("ORDR_QTY").Width
                    ' .Columns("SIZE_QTY_" & Format(i, "00")).Format = "#,##0"
                Else
                    .Columns("SIZE_QTY_" & Format(i, "00")).Hidden = True
                End If
            Next i
        End With
    End Sub

    Sub Set_MS_TOTAL_AMT_Expression()
        Dim EXP As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
            Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
            EXP &= "+ISNULL(QTY_" & Format(ORDR_LNO, "000") & ",0) * " & CStr(ORDR_UNIT_PRICE)
        Next
        dst.Tables("SOTORDRS").Columns("TOTAL_AMT").Expression = Mid(EXP, 2)
    End Sub

    Sub Set_MS_TOTAL_QTY_Expression(Optional ORDR_LNO As Int64 = 0)
        Dim EXP As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            EXP &= "+ISNULL(QTY_" & Format(Val(rowSOTORDR2.Item("ORDR_LNO") & ""), "000") & ",0)"
        Next
        If ORDR_LNO <> 0 Then
            If InStr(EXP, "QTY_" & Format(ORDR_LNO, "000")) = 0 Then
                EXP &= "+QTY_" & Format(ORDR_LNO, "000")
            End If
        End If

        dst.Tables("SOTORDRS").Columns("TOTAL_QTY").Expression = Mid(EXP, 2)
    End Sub

    Sub Setup_SubGrid(Range_tf As Boolean, AddNew_tf As Boolean)
        If grdSOTORDR2.ActiveRow Is Nothing OrElse grdSOTORDR2.ActiveRow.IsAddRow Then
            grpSOTORDR3.Visible = False
            grpSOTORDRR.Visible = False
            'tabDetails.Tabs("Range Style Components").Visible = False
            'tabDetails.Tabs("Customer Style Components").Visible = False
            'tabDetails.Tabs("Order History").Visible = False
            tabDetails.Visible = False

            sub_grid = ""
            Exit Sub
        End If


        Dim Z As String = IIf(Range_tf, "1", "0") & IIf(AddNew_tf, "1", "0") & Format(Val(grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value & ""), "000")
        If sub_grid = Z Then
            Exit Sub
        Else
            sub_grid = Z
        End If

        Me.Cursor = Cursors.WaitCursor
        'With grdSOTORDR2.DisplayLayout.Bands(0)
        '    .Columns("STYLE_DESC").CellActivation = IIf(Range_tf, UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
        '    If Range_tf Or Not AddNew_tf Then
        '        .Columns("STYLE_CODE").Style = UltraWinGrid.ColumnStyle.Default
        '    Else
        '        .Columns("STYLE_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
        '    End If
        '    .Columns("STYLE_CODE").CellActivation = IIf(Range_tf Or Not AddNew_tf, UltraWinGrid.Activation.NoEdit, UltraWinGrid.Activation.AllowEdit)
        '    If Range_tf Or Not AddNew_tf Then
        '        .Columns("COLOR_CODE").Style = UltraWinGrid.ColumnStyle.Default
        '    Else
        '        .Columns("COLOR_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
        '    End If

        '    .Columns("RANGE_STYLE_CODE").CellActivation = IIf(Not Range_tf Or Not AddNew_tf, UltraWinGrid.Activation.NoEdit, UltraWinGrid.Activation.AllowEdit)
        'End With

        'If ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" <> "1" Then
        '    tabDetails.Tabs("Range Style Components").Visible = False
        'End If
        'If ROWs("SOTPARM1").Item("SO_PARM_CUST_STYLE_INFO") & "" <> "1" Then
        '    tabDetails.Tabs("Customer Style Components").Visible = False
        'End If

        If ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" = "1" Then
            Dim RANGE_STYLE_CODE As String = grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & ""
            If RANGE_STYLE_CODE <> "" Then
                tabDetails.Tabs("Range Style Components").Visible = True
                tabDetails.SelectedTab = tabDetails.Tabs("Range Style Components")
            Else
                tabDetails.Tabs("Range Style Components").Visible = False
            End If
        Else
            tabDetails.Tabs("Range Style Components").Visible = False
        End If



        ' tabDetails.Tabs("Range Style Components").Visible = (ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" = "1")

        tabDetails.Tabs("Customer Style Components").Visible = (ROWs("SOTPARM1").Item("SO_PARM_CUST_STYLE_INFO") & "" = "1")

        If Range_tf And ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" = "1" Then
            Setup_Range()

            grpSOTORDR3.Visible = False
            tabDetails.Tabs("Customer Style Components").Visible = False

            UltraExplorerBar1.Groups("Order Details").Items("Define Size Scale").Visible = False
        ElseIf ROWs("SOTPARM1").Item("SO_PARM_CUST_STYLE_INFO") & "" = "1" Then
            Setup_CStyle()

            grpSOTORDRR.Visible = False
            tabDetails.Tabs("Range Style Components").Visible = False
        End If

        tabDetails.Visible = True

        Me.Cursor = Cursors.Default
    End Sub

    Sub Setup_Range()
        Dim ORDR_LNO As Int64 = 0
        If grdSOTORDR2.ActiveRow.IsAddRow Then
            ORDR_LNO = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
        Else
            ORDR_LNO = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value)
        End If
        Dim dvw As DataView = dst.Tables("SOTORDRR").DefaultView
        dvw.RowFilter = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)

        Select Case grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_QTY_PER_PP").Value & ""
            Case Is = "0", ""
                RANGE_TYPE = "U"
            Case Is = "1"
                RANGE_TYPE = "R"
            Case Else
                RANGE_TYPE = "A"
        End Select

        Display_Totals_R(ORDR_LNO)

        If EntryMode = "E" Or EntryMode = "N" Then
            With grdSOTORDRR.DisplayLayout.Override
                If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
                Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
                Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value & "") <> 0 Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End If
            End With
        End If

        grpSOTORDRR.Visible = True And (ROWs("SOTPARM1").Item("SO_PARM_RANGES") & "" = "1")
        fpPPQTY.Text = Val(grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_QTY_PER_PP").Value & "")
        SetRangeType(RANGE_TYPE)
    End Sub

    Sub Setup_CStyle()

        With grdSOTORDR2.ActiveRow
            Dim ORDR_LNO As Integer
            If .IsAddRow Then
                ORDR_LNO = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
            Else
                ORDR_LNO = Val(.Cells("ORDR_LNO").Value)
            End If
            Dim Sqlx As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
            Dim dvw As DataView = dst.Tables("SOTORDR3").DefaultView
            dvw.RowFilter = Sqlx

            If Val(.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
            Or Val(.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
            Or Val(.Cells("ORDR_QTY_CANC").Value & "") <> 0 _
            Or InquiryMode Or Not (EntryMode = "E" Or EntryMode = "N") _
            Then
                With grdSOTORDR3.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With
            Else
                With grdSOTORDR3.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
            End If
        End With

        Setup_Sizes()
        grpSOTORDR3.Visible = True And (ROWs("SOTPARM1").Item("SO_PARM_CUST_STYLE_INFO") & "" = "1")
        UltraExplorerBar1.Groups("Order Details").Items("Define Size Scale").Visible = Not InquiryMode
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTORDR2.ActiveRow
            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    If .Cells("STYLE_CODE").Text <> "" Then
                        Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value & "")
                        Cancel = (STYLE_CODE = "")
                    End If
                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Case "ORDR_QTY"
                    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(.Cells("ORDR_QTY").Value & "") = "" Then
                        MsgBox("Order Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY")
                        Exit Sub
                    End If
                    If Val(.Cells("ORDR_QTY").Value & "") < 0 Then
                        MsgBox("Order Qty May Not be Negative", vbOKOnly, "Invalid Order Quantity")
                        Cancel = True
                    End If
                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                        SQLS.AppendLine("SELECT NVL(STYLE_ASST_QTY,0) AS STYLE_ASST_QTY")
                        SQLS.AppendLine("FROM ICTSTYL1")
                        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", .Cells("STYLE_CODE").Value & ""))
                        ASCMAIN1.sql = SQLS.ToString()
                        Dim STYLE_ASST_QTY As Int16 = Val(ASCDATA1.GetDataValue)
                        If STYLE_ASST_QTY > 0 Then
                            If Val(.Cells("ORDR_QTY").Value & "") Mod STYLE_ASST_QTY <> 0 Then
                                'e.Row.Cells("ORDR_QTY").ToolTipText += " Order Qty not Divisible by Assortment of " & STYLE_ASST_QTY
                                Dim iResult As MsgBoxResult
                                Dim iTitle As String = "Assortment"
                                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                                iMSG.AppendLine("Order Qty not Divisible by Assortment of " & STYLE_ASST_QTY & ".")
                                iMSG.AppendLine("Is That OK With You?")
                                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                                If iResult <> MsgBoxResult.Yes Then
                                    Cancel = True
                                End If
                            End If
                        End If
                    End If
            End Select
        End With
    End Sub

    Sub Validate_Columns_R(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTORDRR.ActiveRow

            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    If .Cells("STYLE_CODE").Value & "" <> "" Then
                        Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Text)
                        Cancel = (STYLE_CODE = "")
                    End If
                Case "COLOR_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Text)
                    If STYLE_CODE = "" Then
                        Cancel = True
                    Else
                        If .Cells("COLOR_CODE").Value & "" <> "" Then
                            If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                                Cancel = True
                            End If
                        Else
                            Cancel = True
                        End If
                    End If
                Case "ORDR_QTY"
                    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(.Cells("ORDR_QTY").Value & "") = "" Then
                        MsgBox("Order Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdSOTORDRR.ActiveCell = grdSOTORDRR.ActiveRow.Cells("ORDR_QTY")
                        Exit Sub
                    End If
                    If Val(.Cells("ORDR_QTY").Value & "") < 0 Then
                        MsgBox("Order Qty May Not be Negative", vbOKOnly, "Invalid Order Quantity")
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Sub Set_PPQTY()
        If Not IsLoading AndAlso grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
            If Val(grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_QTY_PER_PP").Value & "") = 0 Then
                fpPPQTY.Value = 1
                grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_QTY_PER_PP").Value = 1
            Else
                fpPPQTY.Text = Val(grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_QTY_PER_PP").Value & "")
            End If
        End If
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim E As String = ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            If rowICTSTYL1 Is Nothing Then
                Dim PARTAIALSTYLE As String = PARTIALSTYLE(STYLE_CODE_z)
                If PARTAIALSTYLE.Length > 0 Then
                    STYLE_CODE_z = PARTAIALSTYLE
                    rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)
                End If
            End If
        End If

        If rowICTSTYL1 Is Nothing Then
            E = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                ' E = "Style Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                E = "Style does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                E = "Style does not have a valid Division Code" & vbCrLf
            End If
        End If

        If E = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If E <> "" And grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If E = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

    Sub SetRangeType(RangeType As String)
        'A = Assortment.
        'R = Range.
        'U = We don't know yet.
        Select Case RangeType
            Case Is = "R", "U"
                If EntryMode = "E" Or EntryMode = "N" Then
                    With grdSOTORDR2.DisplayLayout.Bands(0)
                        .Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                    End With
                    With grdSOTORDRR.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                End If

                With grdSOTORDRR.DisplayLayout.Bands(0)
                    .Columns("ORDR_UNIT_PRICE").Hidden = True
                    .Columns("QTY_PER_PP").Hidden = True
                End With
            Case Is = "A"
                If EntryMode = "E" Or EntryMode = "N" Then
                    With grdSOTORDR2.DisplayLayout.Bands(0)
                        .Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                    End With
                    With grdSOTORDRR.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                End If

                With grdSOTORDRR.DisplayLayout.Bands(0)
                    .Columns("ORDR_UNIT_PRICE").Hidden = False
                    .Columns("QTY_PER_PP").Hidden = False
                End With
        End Select
    End Sub

    Sub GET_RANGE(RANGE_STYLE_CODE As String, ORDR_LNO As Integer)

        'Dim RANGE_TYPE As String

        Dim rowICTRSTY1 As DataRow = LookUp("ICTRSTY1", New String() {CUST_CODE, RANGE_STYLE_CODE})
        If rowICTRSTY1 IsNot Nothing Then
            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
            If dst.Tables("SOTORDRR").Select(sqlw).Length <> 0 Then
                If MsgBox("This Will Erase All Range / Assortments For This Line Of The Order." _
                            & vbCrLf & "Are You Sure?", MsgBoxStyle.YesNo, "Erase") = vbYes Then
                    ASCDATA1.DeleteRows(dst.Tables("SOTORDRR"), sqlw)
                Else
                    Exit Sub
                End If
            End If
            fpPPQTY.Text = Val(rowICTRSTY1.Item("RANGE_STYLE_QTY_PER_PP") & "")
            grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_QTY_PER_PP").Value = rowICTRSTY1.Item("RANGE_STYLE_QTY_PER_PP")
            grdSOTORDR2.ActiveRow.Cells("STYLE_DESC").Value = rowICTRSTY1.Item("RANGE_STYLE_DESC")
            If Val(rowICTRSTY1.Item("RANGE_STYLE_QTY_PER_PP") & "") = 1 Then
                RANGE_TYPE = "R"
            Else
                RANGE_TYPE = "A"
            End If

            Dim RANGE_STYLE_LNO As Int32 = 0
            Dim RANGE_STYLE_PRICE As Decimal = 0

            ASCMAIN1.sql = "SELECT ICTRSTY2.*, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & " from ICTRSTY2, ICTSTYL1" & vbCrLf _
                & " where ICTRSTY2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTRSTY2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and ICTRSTY2.RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'" & vbCrLf

            For Each rowICTRSTY2 As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim rowSOTORDRR As DataRow = dst.Tables("SOTORDRR").NewRow
                rowSOTORDRR.Item("ORDR_NO") = ORDR_NO
                rowSOTORDRR.Item("ORDR_LNO") = ORDR_LNO
                rowSOTORDRR.Item("STYLE_CODE") = rowICTRSTY2.Item("STYLE_CODE")
                rowSOTORDRR.Item("COLOR_CODE") = rowICTRSTY2.Item("COLOR_CODE")
                rowSOTORDRR.Item("STYLE_DESC") = rowICTRSTY2.Item("STYLE_DESC")
                rowSOTORDRR.Item("RANGE_STYLE_CODE").Value = grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value
                rowSOTORDRR.Item("ORDR_EXTD_COST").Value = Val(rowICTRSTY2.Item("STYLE_PRICE") & "") * Val(rowICTRSTY2.Item("STYLE_QTY") & "")
                rowSOTORDRR.Item("ORDR_UNIT_PRICE").Value = rowICTRSTY2.Item("STYLE_PRICE")
                If RANGE_TYPE = "R" Then
                    RANGE_STYLE_PRICE = Val(rowICTRSTY2.Item("STYLE_PRICE") & "")
                Else
                    RANGE_STYLE_PRICE = RANGE_STYLE_PRICE + Val(rowICTRSTY2.Item("STYLE_PRICE") & "")
                End If
                rowSOTORDRR.Item("ORDR_QTY") = rowICTRSTY2.Item("STYLE_QTY")
                rowSOTORDRR.Item("ORDR_QTY_OPEN") = rowICTRSTY2.Item("STYLE_QTY")
                rowSOTORDRR.Item("ORDR_STATUS") = "O"
                rowSOTORDRR.Item("ORDR_QTY_ORIG") = rowICTRSTY2.Item("STYLE_QTY")
                rowSOTORDRR.Item("QTY_PER_PP") = rowICTRSTY2.Item("STYLE_QTY")
                dst.Tables("SOTORDRR").Rows.Add(rowSOTORDRR)
            Next

            grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE").Value = RANGE_STYLE_PRICE
        End If

        If RANGE_TYPE = "" Then
            RANGE_TYPE = "U"
        End If
        SetRangeType(RANGE_TYPE)

    End Sub

    Sub Update_PrePack(ORDR_QTY As Integer)

        Dim PRIOR_ORDERED As Int64
        Dim NEW_ORDERED As Int64
        Dim ORDR_LNO As Int32 = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value & "")
        Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)

        Dim rowSOTORDRRs() As DataRow = dst.Tables("SOTORDRR").Select(sqlw)
        'CALC OLD ORDERED
        If rowSOTORDRRs.Length <> 0 Then
            If ORDR_QTY = 0 Then
                NEW_ORDERED = 0
            Else
                PRIOR_ORDERED = Val(rowSOTORDRRs(0).Item("ORDR_QTY") & "") / Val(rowSOTORDRRs(0).Item("QTY_PER_PP") & "")
                NEW_ORDERED = ORDR_QTY - PRIOR_ORDERED
            End If
        Else
            NEW_ORDERED = 1
        End If
        For Each row As DataRow In rowSOTORDRRs
            row.Item("ORDR_QTY") = Val(row.Item("ORDR_QTY") & "") + (Val(row.Item("QTY_PER_PP") & "") * NEW_ORDERED)
            row.Item("ORDR_QTY_OPEN") = Val(row.Item("ORDR_QTY_OPEN") & "") + (Val(row.Item("QTY_PER_PP") & "") * NEW_ORDERED)
            'WE SHOULD LEAVE SHIP ALONE.  HANDLE THAT IN CONF. WITH INVH9 ONLY.
        Next
    End Sub

#End Region

#Region "grdSOTORDR2"

    Private Sub grdSOTORDR2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim STYLE_CODE As String = Validate_Style(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value)
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    If STYLE_CODE <> e.Cell.Value & "" Then
                        grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value = STYLE_CODE
                    End If
                    ShowPromo(STYLE_CODE)
                Else
                    lblPromo.Visible = False
                    lblPromo.Text = ""
                    btnShowPromo.Visible = False
                End If
                If STYLE_CODE <> "" Then
                    e.Cell.Row.Cells("STYLE_UOM").Value = rowICTSTYL1.Item("STYLE_UOM") & ""
                    e.Cell.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                    e.Cell.Row.Cells("INNER_PACK_QTY").Value = rowICTSTYL1.Item("INNER_PACK_QTY")
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
                    e.Cell.Row.Cells("RANGE_STYLE_CODE").Value = DBNull.Value
                    e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                    e.Cell.Row.Cells("STYLE_PRICE").Value = rowICTSTYL1.Item("STYLE_PRICE")
                    e.Cell.Row.Cells("CASE_CUBE").Value = rowICTSTYL1.Item("CASE_CUBE")

                    If Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                        e.Cell.Row.Cells("PO_COST").Value = TAC.SOCMAIN1.Get_PO_Cost(Me, STYLE_CODE, rowICTSTYL1.Item("VEND_CODE") & "", rowSOTORDR1)
                    End If

                    If COLOR_CODEs.Count = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                    End If
                    sub_grid = ""
                    Setup_SubGrid(False, e.Cell.Row.IsAddRow)

                    Order_History(STYLE_CODE, "")

                    Price_and_Availability(STYLE_CODE,
                                           e.Cell.Row.Cells("STYLE_CLASS_CODE").Value & "",
                                           "",
                                           Val(e.Cell.Row.Cells("CARTON_PACK_QTY").Value & ""),
                                           Val(e.Cell.Row.Cells("STYLE_PRICE").Value & ""))

                    If tabDetails.Tabs("Pricing && Availability").Visible Then
                        tabDetails.SelectedTab = tabDetails.Tabs("Pricing && Availability")
                    End If
                    tabDetails.Tabs("Range Style Components").Visible = False
                    tabDetails.Tabs("Customer Style Components").Visible = False
                    tabDetails.Visible = True


                    If ASCMAIN1.CLIENT = "NYA" Then

                        If CURR_CODE <> "" And CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then

                            Dim rowSOTPRIC2 As DataRow = Nothing
                            Dim PRICE_LIST_CODE As String = "" ' = HFs("CUST_CODE")
                            PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & ""
                            If PRICE_LIST_CODE = "" Then
                                PRICE_LIST_CODE = HFs("CUST_CODE")
                            End If
                            rowSOTPRIC2 = LookUp("SOTPRIC2", New String() {PRICE_LIST_CODE, STYLE_CODE})
                            If rowSOTPRIC2 Is Nothing Then
                                e.Cell.Row.Cells("ORDR_UNIT_PRICE_CURR").Value = 0
                                e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = e.Cell.Row.Cells("ORDR_UNIT_PRICE_CURR").Value * CURR_EXCH_RATE

                            Else
                                e.Cell.Row.Cells("ORDR_UNIT_PRICE_CURR").Value = rowSOTPRIC2.Item("STYLE_PRICE")
                                e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = e.Cell.Row.Cells("ORDR_UNIT_PRICE_CURR").Value * CURR_EXCH_RATE

                            End If
                        End If
                    End If


                Else

                End If

            Case "COLOR_CODE"
                Dim COLOR_CODE As String = e.Cell.Value & "" ' grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
                If COLOR_CODE <> "" Then
                    If COLOR_CODEs.Contains(COLOR_CODE) Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        e.Cell.Row.Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                    End If

                    Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value & ""
                    Order_History(STYLE_CODE, COLOR_CODE)
                    Price_and_Availability(STYLE_CODE,
                                           e.Cell.Row.Cells("STYLE_CLASS_CODE").Value & "",
                                           COLOR_CODE,
                                           Val(e.Cell.Row.Cells("CARTON_PACK_QTY").Value & ""),
                                           Val(e.Cell.Row.Cells("STYLE_PRICE").Value & ""))
                    If tabDetails.Tabs("Pricing && Availability").Visible Then
                        tabDetails.SelectedTab = tabDetails.Tabs("Pricing && Availability")
                    End If

                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR" Then
                            e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = 0
                        Else
                            If Val(e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value & "") = 0 Then
                                Dim rowSOTRSRVX As DataRow = Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                                If rowSOTRSRVX IsNot Nothing Then
                                    e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = rowSOTRSRVX.Item("ORDR_UNIT_PRICE")
                                End If
                            End If
                        End If
                    End If
                End If

            Case "RANGE_STYLE_CODE"
                If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                    Setup_SubGrid(True, False)
                    'grdSOWORDR2.Columns("STYLE_DESC").Text = "Style Range"
                    grdSOTORDR2.ActiveRow.Cells("STYLE_UOM").Value = "EA"
                    grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("STYLE_DESC")
                    grdSOTORDR2.ActiveRow.Update()
                End If

            Case "ORDR_QTY"
                grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value
                Set_PPQTY()
                If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" _
                    And Val(fpPPQTY.Value & "") > 1 Then 'We are updating an Pre-Pack.
                    Update_PrePack(Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value & ""))
                End If

                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim ORDR_PRICE_SOURCE As String = ""
                    Dim ORDR_UNIT_PRICE_CALC As Decimal = 0
                    Dim ORDR_UNIT_PRICE_STD As Decimal = 0

                    Dim STYLE_CODE As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & ""
                    Dim COLOR_CODE As String = grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""

                    Dim ORDR_QTY As Int32 = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "")

                    If Absx1.optFor("ORDR_TYPE_CODE").Value = "B2C" Then
                        ' DO NONE OF THIS SILLY PRICE STUFF FROR B2C ORDERS
                    Else
                        If STYLE_CODE <> "" And COLOR_CODE <> "" Then
                            If Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                                Dim BTB_Price As New FEFDPrice(Me, STYLE_CODE, 1)
                                If Absx1.txtFor("WHSE_CODE").Text = "FE" Then
                                    ORDR_UNIT_PRICE_CALC = BTB_Price.FEPrice
                                    ORDR_PRICE_SOURCE = "FE"
                                Else
                                    ORDR_UNIT_PRICE_CALC = BTB_Price.FDPrice
                                    ORDR_PRICE_SOURCE = "FD"
                                End If
                                ORDR_UNIT_PRICE_STD = BTB_Price.FEPrice
                                e.Cell.Row.Cells("ORDR_UNIT_PRICE_STD").Value = ORDR_UNIT_PRICE_STD
                                BTB_Price = Nothing
                            Else
                                ORDR_UNIT_PRICE_CALC = TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1,
                                       STYLE_CODE, COLOR_CODE, ORDR_QTY, ORDR_PRICE_SOURCE)
                            End If

                            e.Cell.Row.Cells("ORDR_UNIT_PRICE_CALC").Value = ORDR_UNIT_PRICE_CALC
                            e.Cell.Row.Cells("ORDR_PRICE_SOURCE").Value = ORDR_PRICE_SOURCE
                            If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "SAM" Or Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR" Then
                                e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = 0
                            Else
                                If e.Cell.Row.Cells("ORDR_UNIT_PRICE_MANUAL").Value & "" <> "1" Then
                                    e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE_CALC
                                End If
                            End If
                        Else
                            ' ?
                        End If
                    End If
                End If

            Case "ORDR_QTY_OPEN"
                grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") _
                        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") _
                        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value & "") _
                        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "")
                If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Text) < 0 Then
                    grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = 0
                End If

            Case "ORDR_UNIT_PRICE"
                Set_PPQTY()

                If Absx1.optFor("ORDR_TYPE_CODE").Value = "B2C" Then
                    ' DO NONE OF THIS SILLY PRICE STUFF FROR B2C ORDERS
                Else
                    If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE").Value & "") <> Val(grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE_CALC").Value & "") Then
                        grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE_MANUAL").Value = "1"
                    End If
                End If

            Case "ORDR_UNIT_PRICE_MANUAL"

                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    If Absx1.optFor("ORDR_TYPE_CODE").Value = "B2C" Then
                        ' DO NONE OF THIS SILLY PRICE STUFF FROR B2C ORDERS
                    Else

                        Dim ORDR_PRICE_SOURCE As String = ""
                        Dim ORDR_UNIT_PRICE_CALC As Decimal = 0
                        If Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Then
                            Dim BTB_Price As New FEFDPrice(Me, grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "", 1)
                            If Absx1.txtFor("WHSE_CODE").Text = "FE" Then
                                ORDR_UNIT_PRICE_CALC = BTB_Price.FEPrice
                                ORDR_PRICE_SOURCE = "FE"
                            Else
                                ORDR_UNIT_PRICE_CALC = BTB_Price.FDPrice
                                ORDR_PRICE_SOURCE = "FD"
                            End If
                            BTB_Price = Nothing
                        Else
                            ORDR_UNIT_PRICE_CALC = TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1,
                                       grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "",
                                       grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & "",
                                       Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & ""), ORDR_PRICE_SOURCE)
                        End If
                        e.Cell.Row.Cells("ORDR_UNIT_PRICE_CALC").Value = ORDR_UNIT_PRICE_CALC
                        If e.Cell.Row.Cells("ORDR_UNIT_PRICE_MANUAL").Value & "" <> "1" Then
                            e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE_CALC
                        End If
                    End If
                End If
        End Select
    End Sub

    Private Sub grdSOTORDR2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR2.AfterRowActivate

        Setup_grdSOTORDR2()

    End Sub

    Sub Setup_grdSOTORDR2()

        RANGE_TYPE = ""

        If grdSOTORDR2.ActiveRow Is Nothing Then
            lblPromo.Visible = False
            lblPromo.Text = ""
            btnShowPromo.Visible = False
            Exit Sub
        End If

        Setup_SOTORDR3()
        Set_PPQTY()

        If grdSOTORDR2.ActiveRow Is Nothing OrElse grdSOTORDR2.ActiveRow.IsAddRow Then
            lblPromo.Visible = False
            lblPromo.Text = ""
            btnShowPromo.Visible = False
        Else
            Dim STYLE_CODE As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
            Dim STYLE_CLASS_CODE As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CLASS_CODE").Value & ""
            Dim CARTON As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CLASS_CODE").Value & ""

            Order_History(STYLE_CODE, COLOR_CODE)
            Price_and_Availability(STYLE_CODE, STYLE_CLASS_CODE, COLOR_CODE,
                                   Val(grdSOTORDR2.ActiveRow.Cells("CARTON_PACK_QTY").Value & ""),
                                   Val(grdSOTORDR2.ActiveRow.Cells("STYLE_PRICE").Value & ""))
            If tabDetails.Tabs("Pricing && Availability").Visible Then
                tabDetails.SelectedTab = tabDetails.Tabs("Pricing && Availability")
            End If

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                ShowPromo(STYLE_CODE)
            End If

        End If

        If Trim(grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "") = "" And
           Trim(grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "") = "" And
            (grdSOTORDR2.ActiveCell Is Nothing OrElse
             (grdSOTORDR2.ActiveCell.Column.Key <> "STYLE_CODE" And
              grdSOTORDR2.ActiveCell.Column.Key <> "RANGE_STYLE_CODE")) _
        Then
            ' this line clears the screen and makes users think that their lines are gone
            '  grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE")
        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            With grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE_MANUAL")
                If grdSOTORDR2.ActiveRow.IsAddRow Then
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End With
        End If

        If grdSOTORDR2.ActiveRow.IsAddRow Then
            With grdSOTORDR2.DisplayLayout.Bands(0)
                .Columns("RANGE_STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End With

            ''- DOUBLE REMS ARE MINE
            If grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" _
                And grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" = "" Then

                If grdSOTORDR2.ActiveRow.IsAddRow Then
                    Setup_SubGrid(False, True)
                Else
                    If grdSOTORDR2.ActiveCell.Column.Key = "STYLE_CODE" Then
                        Setup_SubGrid(False, True)
                        '' grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE")
                    Else
                        Setup_SubGrid(True, True)
                        '' grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE")
                    End If

                End If
            Else
                ' i don't think I ever hit this code - it should be removed - DOUBLE REMS ARE MINE

                ''With grdSOTORDR2.DisplayLayout.Bands(0)
                ''    If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                ''        .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                ''        If RANGE_TYPE <> "A" Then
                ''            .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                ''        Else
                ''            .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                ''        End If
                ''    Else
                ''        .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                ''    End If
                ''End With
            End If

        Else
            If grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                Setup_SubGrid(True, False)
            Else
                Setup_SubGrid(False, False)
            End If

            ' THIS CODE WAS REMMED OUT IN LOAD RECORD
            'If EntryMode = "E" And Absx1.optFor("ORDR_SOURCE").Value = "E" Then
            '    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
            'Else
            '    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            'End If
            'If EntryMode = "E" Then
            '    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
            'Else
            '    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            'End If

            With grdSOTORDR2.DisplayLayout.Bands(0)
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RANGE_STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

                If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" = "" Then
                    Validate_Style(grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "")
                    .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit

                    ' Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value & "") <> 0 
                    If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
                    Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
                      Then
                        'Or absx1.txtfor("ORDR_SOURCE")).Text = "E" 'was also part of this
                        .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                    Else
                        .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                Else
                    .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit

                End If
            End With

            If grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                If Val(CInt(fpPPQTY.Text)) = 1 Or Val(fpPPQTY.Text) = 0 Then
                    SetRangeType("R")
                Else
                    SetRangeType("A")
                End If
            End If
        End If
    End Sub
    Private Sub grdSOTORDR2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTORDR2.AfterRowsDeleted
        If multi_store_is_active Then ' And Not multi_store_changes_made_to_SOTORDRS Then
            dst.Tables("SOTORDRS").Columns("TOTAL_AMT").Expression = ""
            dst.Tables("SOTORDRS").Columns("TOTAL_QTY").Expression = ""
            For Each ORDR_LNO As Int64 In ORDR_LNOs
                Dim summary As UltraWinGrid.SummarySettings = grdSOTORDRS.DisplayLayout.Bands(0).Summaries("QTY_" & Format(ORDR_LNO, "000"))
                grdSOTORDRS.DisplayLayout.Bands(0).Summaries.Remove(summary)
                dst.Tables("SOTORDRS").Columns.Remove("QTY_" & Format(ORDR_LNO, "000"))
            Next
        End If

        Display_Totals()

        If grdSOTORDR2.Rows.Count = 0 Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = ""
        End If

        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
        If grdSOTORDR2.Rows.Count = 0 Then
            Setup_SubGrid(False, False)
        End If
    End Sub

    Private Sub grdSOTORDR2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDR2.AfterRowUpdate

        If multi_store_is_active Then
            If multi_store_changes_made_to_SOTORDRS Then
                multi_store_changes_made_to_SOTORDRS = False ' Turn off multix from BeforeUpdate
            Else
                multi_store_changes_made_to_SOTORDRS = True
                Dim rowSOTORDRS As DataRow = dst.Tables("SOTORDRS").Rows.Find(Absx1.txtFor("CUST_STORE_NO").Text)
                Dim ORDR_LNO As Int64 = Val(e.Row.Cells("ORDR_LNO").Value & "")
                rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) = Val(e.Row.Cells("ORDR_QTY").Value)
                multi_store_changes_made_to_SOTORDRS = False
            End If
        End If

        Display_Totals()

        If e.Row.Cells("STYLE_CODE").Value & "" <> "" And Absx1.txtFor("SALES_DIVISION_CODE").Text = "" AndAlso rowICTSTYL1 IsNot Nothing Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTSTYL1.Item("SALES_DIVISION_CODE") & ""
        End If

        ' If e.Row.IsAddRow Then
        ' if we just added a row
        If EntryMode = "N" Or EntryMode = "E" Then
            If e.Row.Cells("ORDR_STATUS").Tag & "" = "Added" Then
                Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
                grdSOTORDR2.DisplayLayout.Bands(0).AddNew()
                e.Row.Cells("ORDR_STATUS").Tag = DBNull.Value
            End If
        End If
        ' End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Dim M As String = "###,##0.00"
            If grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = M Then
                Dim ORDR_UNIT_PRICE As Decimal = Val(e.Row.Cells("ORDR_UNIT_PRICE").Value & "")
                If Format(ORDR_UNIT_PRICE, "###.00") & "00" <> Format(ORDR_UNIT_PRICE, "###.0000") Then
                    M = "###.0000"
                    grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = M
                    grdSOTORDR2.Rows.Refresh(UltraWinGrid.RefreshRow.RefreshDisplay)
                End If
            End If
        End If

    End Sub

    Private Sub grdSOTORDR2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDR2.BeforeCellUpdate

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim STYLE_CODE As String = Validate_Style(e.NewValue & "")
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    If STYLE_CODE <> e.NewValue & "" Then
                        If Not IsNothing(grdSOTORDR2.ActiveCell) Then
                            'e.Cancel = True
                            'grdSOTORDR2.ActiveCell.Value = STYLE_CODE
                        End If
                    End If
                End If
                If STYLE_CODE = "" Then
                    e.Cancel = True
                End If

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    MsgBox("You Must Specify a Warehouse before entering Styles",
                           MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    e.Cancel = True
                    Exit Sub
                End If

            Case "RANGE_STYLE_CODE"
                If multi_store_is_active Then
                    If e.Cell.Value & "" <> "" Then
                        MsgBox("Cannot Have Range Styles on a Multi-Store Order",
                               MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
        End Select

    End Sub

    Private Sub grdSOTORDR2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTORDR2.BeforeExitEditMode
        If grdSOTORDR2.ActiveCell IsNot Nothing Then
            With grdSOTORDR2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE", "RANGE_STYLE_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        If .Column.Key = "RANGE_STYLE_CODE" And .EditorResolved.Value & "" <> CStr(.EditorResolved.Value & "").ToUpper Then
                            .EditorResolved.Value = CStr(.EditorResolved.Value & "").ToUpper
                        End If
                    Case "ORDR_QTY"
                        If .EditorResolved.Value & "" = "" _
                        Or Val(.EditorResolved.Value & "") < 0 _
                        Then
                            .EditorResolved.Value = 0
                        End If
                        If Val(.EditorResolved.Value & "") < 0 Then
                            .EditorResolved.Value = System.Math.Abs(Val(.EditorResolved.Value & ""))
                        End If

                        Dim STYLE_CODE As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & ""
                        Dim COLOR_CODE As String = grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
                        If STYLE_CODE = "" Or COLOR_CODE = "" Then
                            .EditorResolved.Value = DBNull.Value ' IF YOU SET THIS TO 0, THEN THE LINE WILL BE PERMITTED TO BE UPDATED
                        End If

                    Case "ORDR_UNIT_PRICE"

                        Dim STYLE_CODE As String = grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & ""
                        Dim COLOR_CODE As String = grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
                        Dim RANGE_STYLE_CODE As String = grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value & ""
                        If (STYLE_CODE = "" Or COLOR_CODE = "") And RANGE_STYLE_CODE = "" Then
                            .EditorResolved.Value = DBNull.Value ' IF YOU SET THIS TO 0, THEN THE LINE WILL BE PERMITTED TO BE UPDATED
                        End If

                    Case "ORDR_QTY_OPEN"
                        If .EditorResolved.Value & "" = "" _
                        Or Val(.EditorResolved.Value & "") < 0 _
                        Then
                            .EditorResolved.Value = 0
                        End If
                        Dim q As Int64 = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "") _
                                       + Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "")
                        '+ Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value & "")
                        If Val(.EditorResolved.Value & "") > Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") - q Then
                            .EditorResolved.Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") - q
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTORDR2_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDR2.BeforeRowActivate
        If grdSOTORDRR.ActiveRow IsNot Nothing AndAlso grdSOTORDRR.ActiveRow.DataChanged Then
            grdSOTORDRR.ActiveRow.Update()
        End If
        If grdSOTORDR3.ActiveRow IsNot Nothing AndAlso grdSOTORDR3.ActiveRow.DataChanged Then
            grdSOTORDR3.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdSOTORDR2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTORDR2.BeforeRowsDeleted

        ORDR_LNOs.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Not grow.IsAddRow Then
                If Val(grow.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
                Or Val(grow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
                Or Val(grow.Cells("ORDR_QTY_CANC").Value & "") <> 0 _
                Then
                    MsgBox("Cannot Delete a Line if it has ever been " & vbCr & "Picked, Shipped Or Cancelled" & vbCr & "Use the Cancel Button (x)")
                    e.Cancel = True
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTPICK1,SOTPICK2" _
                    & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" _
                    & "   and SOTPICK1.PICK_STATUS = 'P'" _
                    & "   and SOTPICK2.ORDR_NO = '" & ORDR_NO & "' and SOTPICK2.ORDR_LNO = " & grow.Cells("ORDR_LNO").Value
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    MsgBox("Cannot Delete a Line if there is an active Pick Ticket (In Pick) referring to the line")
                    e.Cancel = True
                    Exit Sub
                End If

                If grow.Cells("RSRV_NO").Value & "" <> "" Then
                    MsgBox("Cannot Delete a Line if it has ever been " _
                           & vbCrLf & "Used in a Reservation" & vbCrLf & "Use the Cancel Button (x)")
                    e.Cancel = True
                    Exit Sub
                End If

                If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                    If ALLOW_CHANGE_RANGE <> "1" Then
                        MsgBox("Cannot Delete a Line from an EDI Order" & vbCrLf & "Use the Cancel Button (x)")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
                ORDR_LNOs.Add(grow.Cells("ORDR_LNO").Value)
            End If
        Next
    End Sub

    Private Sub grdSOTORDR2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDR2.BeforeRowUpdate
        Dim iResult As String

        If e.Row.Cells("RANGE_STYLE_CODE").Value & "" = "" Then
            Validate_Columns("STYLE_CODE", e.Cancel)
            If e.Cancel Then MsgBox("Invalid Style")
            If Not e.Cancel Then
                Validate_Columns("COLOR_CODE", e.Cancel)
                If e.Cancel Then MsgBox("Invalid Color")
            Else

            End If
            If Not e.Cancel Then
                Validate_Columns("ORDR_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
                If e.Cancel Then MsgBox("Invalid Qty")
            End If
        Else
            If e.Row.Cells("STYLE_CODE").Value & "" <> "" Then
                MsgBox("Cannot Have Both a Style Code and a Range Style Code on the Same Line", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
                Exit Sub
            End If

            If multi_store_is_active Then
                MsgBox("Cannot Have Range Styles on a Multi-Store Order", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
                Exit Sub
            End If
        End If

        If Absx1.optFor("ORDR_SOURCE").Value = "K" Then 'Only Check KeyBrd Orders 11/24/04 W.R.
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                If e.Row.Cells("CUST_UPC").Value & "" <> "" Then
                    iResult = TAC.TACMAIN1.Validate_UPC(e.Row.Cells("CUST_UPC").Value & "", ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                    If iResult & "" <> "" Then
                        MsgBox(iResult, vbOKOnly, "UPC Error")
                        e.Cancel = True
                    End If
                End If
            End If
        End If


        If Not e.Cancel Then
            If e.Row.IsAddRow Then
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
                If dst.Tables("SOTORDR2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length <> 0 Then
                    'MsgBox("Style/Color " & "" & " already on Order")
                    'ASCMAIN1.Progress("Style/Color " & STYLE_CODE & "/" & COLOR_CODE & " is already on Order")
                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        ' PERMIT THIS CONDITION SO THAT JOHN CAN ENTER THE SAME STYLE ON MULTIPLE LINES SO THAT HE CAN SUB FROM MULTIPLE STYLES

                    Else
                        MsgBox("Style/Color " & STYLE_CODE & "/" & COLOR_CODE & " is already on Order", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        e.Cancel = True
                    End If
                End If
            End If
        End If


        If e.Cancel = True Then
            Exit Sub
        End If


        STYLE_CODE_last_entry = e.Row.Cells("STYLE_CODE").Value & ""

        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_NO").Value = ORDR_NO
            Dim ORDR_LNO As Int64 = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
            If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "BTB" And ORDR_LNO <= ORDR_LNO_ctr Then
                ORDR_LNO = ORDR_LNO_ctr + 1
            End If
            e.Row.Cells("ORDR_LNO").Value = ORDR_LNO
            e.Row.Cells("ORDR_QTY_ORIG").Value = e.Row.Cells("ORDR_QTY").Value
            e.Row.Cells("ORDR_STATUS").Value = "O"
            e.Row.Cells("ORDR_STATUS").Tag = "Added"
            If e.Row.Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                e.Row.Cells("STYLE_CODE").Value = DBNull.Value
                e.Row.Cells("COLOR_CODE").Value = DBNull.Value
                e.Row.Cells("COLOR_DESC").Value = DBNull.Value
            End If

            If multi_store_is_active Then
                multi_store_changes_made_to_SOTORDRS = True
                Dim RANGE_STYLE_CODE As String = e.Row.Cells("RANGE_STYLE_CODE").Value & ""
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value
                Fill_Styles(ORDR_LNO, RANGE_STYLE_CODE, STYLE_CODE, COLOR_CODE)
                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                    rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) = e.Row.Cells("ORDR_QTY").Value
                Next
                'Set_MS_TOTAL_AMT_Expression() ' gonna happen in afterrowupdate -> DisplayTotals
                'multi_store_changes_made_to_SOTORDRS = False 
                ' Turn multi_store_changes_made_to_SOTORDRS off in the AfterUpdate event
            End If
        End If
    End Sub

    Private Sub grdSOTORDR2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "X"
                    If Val(.Cells("ORDR_QTY_CANC").Value & "") <> 0 Then
                        If MsgBox("Restore Cancelled Qty of " & .Cells("ORDR_QTY_CANC").Value,
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select("ORDR_LNO = " & .Cells("ORDR_LNO").Value)
                                rowSOTORDRR.Item("ORDR_QTY_OPEN") = Val(rowSOTORDRR.Item("ORDR_QTY_OPEN") & "") + Val(rowSOTORDRR.Item("ORDR_QTY_CANC") & "")
                                rowSOTORDRR.Item("ORDR_QTY_CANC") = 0
                            Next
                        End If
                        .Cells("ORDR_QTY_OPEN").Value = Val(.Cells("ORDR_QTY_OPEN").Value & "") + Val(.Cells("ORDR_QTY_CANC").Value & "")
                        ' grdSOWORDR2_AfterColUpdate(.Cells("ORDR_QTY_OPEN").position)
                        .Update()
                    Else
                        If MsgBox("Cancel Remaining Qty Open of " & .Cells("ORDR_QTY_OPEN").Value,
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowSOTORDRR As DataRow In dst.Tables("SOTORDRR").Select("ORDR_LNO = " & .Cells("ORDR_LNO").Value)
                                rowSOTORDRR.Item("ORDR_QTY_OPEN") = 0
                                Dim ORDR_QTY_CANC As Int64 = Val(rowSOTORDRR.Item("ORDR_QTY") & "") _
                                                           - Val(rowSOTORDRR.Item("ORDR_QTY_SHIP") & "") _
                                                           - Val(rowSOTORDRR.Item("ORDR_QTY_PICK") & "")
                                rowSOTORDRR.Item("ORDR_QTY_CANC") = IIf(ORDR_QTY_CANC < 0, 0, ORDR_QTY_CANC)
                            Next
                        End If
                        .Cells("ORDR_QTY_OPEN").Value = "0"
                        ' grdSOWORDR2_AfterColUpdate(.Cells("ORDR_QTY_OPEN").position)
                        grdSOTORDR2.ActiveRow.Update()
                    End If

                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTORDR2, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                    grdClickCellButton(grdSOTORDR2, sql_where)

                Case "RANGE_STYLE_CODE"
                    If e.Cell.Value & "" = "" Then
                        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("RANGE_STYLE_CODE", "ICTRSTY1")
                        If ASCMAIN1.CodeSelector.SQL <> "" Then
                            ASCMAIN1.CodeSelector.MultipleSelections = False
                            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                            ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Add("CUST_CODE", Absx1.txtFor("CUST_CODE").Text)
                            Using F As New ASFCODE1
                                F.ShowDialog()
                            End Using
                            If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                                grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value = ASCMAIN1.CodeSelector.SelectedCode
                                Dim ORDR_LNO As Int64 = Val(e.Cell.Row.Cells("ORDR_LNO").Value & "")
                                GET_RANGE(grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value, ORDR_LNO)
                            End If
                        End If
                    End If

                Case "STYLE_CODE_SUB"
                    If .IsAddRow Then
                        MsgBox("You Must First Complete the Line, then you may Specify a Substitute",
                               MsgBoxStyle.OkOnly, "Cannot Specify a Substitute on a Line Not already Added to Order")
                        Exit Sub
                    End If
                    Dim STYLE_CODE_SUB As String
                    Dim COLOR_CODE_SUB As String
                    If Val(.Cells("ORDR_QTY_PICK").Value & "") = 0 And
                       Val(.Cells("ORDR_QTY_SHIP").Value & "") = 0 And
                       Val(.Cells("ORDR_QTY_CANC").Value & "") = 0 Then

                        COLOR_CODE_SUB = .Cells("COLOR_CODE").Value & ""
                        STYLE_CODE_SUB = Select_Style(COLOR_CODE_SUB)

                        If STYLE_CODE_SUB <> "" Then
                            Dim z As String = .Cells("STYLE_CODE").Value
                            Dim STYLE_CODE As String = Validate_Style(STYLE_CODE_SUB)
                            If STYLE_CODE = "" Then
                                STYLE_CODE = z
                                Validate_Style(z)
                            Else
                                If .Cells("STYLE_CODE_SUB").Value = "" Then
                                    .Cells("STYLE_CODE_SUB").Value = .Cells("STYLE_CODE").Value
                                End If
                                .Cells("STYLE_CODE").Value = STYLE_CODE_SUB
                                ' .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                                .Update()
                            End If
                        End If
                    Else
                        ' CANNOT SUB STYLE IF PICKED, SHIPPED OR CANCELLED
                    End If
            End Select
        End With

    End Sub

#End Region

#Region "grdSOTORDR3"
    Private Sub grdSOTORDR3_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDR3.BeforeCellUpdate

    End Sub

    Private Sub grdSOTORDR3_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTORDR3.BeforeExitEditMode
        If grdSOTORDR3.ActiveCell IsNot Nothing Then
            With grdSOTORDR3.ActiveCell
                Select Case .Column.Key
                    Case "CUST_STYLE_CODE", "CUST_COLOR_CODE"
                        If .EditorResolved.Value & "" <> CStr(.EditorResolved.Value & "").ToUpper Then
                            .EditorResolved.Value = CStr(.EditorResolved.Value & "").ToUpper
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTORDR3_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDR3.BeforeRowUpdate
        If Val(e.Row.Cells("ORDR_QTY").Value & "") = 0 Then
            e.Cancel = True
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_NO").Value = ORDR_NO
            Dim ORDR_LNO As Int64 = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value & "")
            e.Row.Cells("ORDR_LNO").Value = ORDR_LNO
            e.Row.Cells("ORDR_SUB_LNO").Value = Val(dst.Tables("SOTORDR3").Compute("MAX(ORDR_SUB_LNO)", "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)) & "") + 1
        End If
    End Sub
#End Region

#Region "grdSOTORDR4"

    Private Sub grdSOTORDR4_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDR4.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_CLNO").Value = Val(dst.Tables("SOTORDR4").Compute("MAX(ORDR_CLNO)", "") & "") + 1
            e.Row.Cells("ORDR_NO").Value = ORDR_NO
            e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
        Else
            e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        End If
    End Sub
#End Region

#Region "grdSOTORDRR"

    Private Sub grdSOTORDRR_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDRR.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim STYLE_CODE As String = Validate_Style(e.Cell.Value & "")
                If STYLE_CODE <> "" Then
                    e.Cell.Row.Cells("STYLE_UOM").Value = rowICTSTYL1.Item("STYLE_UOM")
                    e.Cell.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                    e.Cell.Row.Cells("INNER_PACK_QTY").Value = rowICTSTYL1.Item("INNER_PACK_QTY")
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
                    e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                    e.Cell.Row.Cells("STYLE_PRICE").Value = rowICTSTYL1.Item("STYLE_PRICE")
                    ' e.Cell.Row.Cells("CASE_CUBE").Value = rowICTSTYL1.Item("CASE_CUBE")
                    If COLOR_CODEs.Count = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                    End If
                End If

            Case "ORDR_UNIT_PRICE"
                'grdSOWORDRR.MoveNext()

            Case "ORDR_QTY"
                e.Cell.Row.Cells("ORDR_QTY_OPEN").Value = e.Cell.Row.Cells("ORDR_QTY").Value

            Case "ORDR_QTY_OPEN"
                e.Cell.Row.Cells("ORDR_QTY").Value = e.Cell.Row.Cells("ORDR_QTY_OPEN").Value
                Dim QTY As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY").Value & "") _
                                 - Val(e.Cell.Row.Cells("ORDR_QTY_PICK").Value & "") _
                                 - Val(e.Cell.Row.Cells("ORDR_QTY_SHIP").Value & "") _
                                 - Val(e.Cell.Row.Cells("ORDR_QTY_OPEN").Value & "")
                If QTY < 0 Then
                    QTY = 0
                End If
                e.Cell.Row.Cells("ORDR_QTY_CANC").Value = QTY
        End Select
    End Sub

    Private Sub grdSOTORDRR_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDRR.AfterRowActivate
        If grdSOTORDRR.ActiveRow.IsAddRow Then
            grdSOTORDRR.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSOTORDRR.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If

        If Trim(grdSOTORDRR.ActiveRow.Cells("STYLE_CODE").Value & "") & "" = "" _
            And (grdSOTORDRR.ActiveCell Is Nothing OrElse grdSOTORDRR.ActiveCell.Column.Key <> "STYLE_CODE") Then
            grdSOTORDRR.ActiveCell = grdSOTORDRR.ActiveRow.Cells("STYLE_CODE")
            grdSOTORDRR.ActiveColScrollRegion.Scroll(UltraWinGrid.ColScrollAction.Left)
            Exit Sub
        End If

        If grdSOTORDRR.ActiveRow.IsAddRow Then
        Else
            Validate_Style(grdSOTORDRR.ActiveRow.Cells("STYLE_CODE").Value & "")
        End If
    End Sub

    Private Sub grdSOTORDRR_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTORDRR.AfterRowsDeleted

        Dim ORDR_LNO As Integer = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value & "")
        Display_Totals_R(ORDR_LNO)
        SetRangeType("U")
    End Sub

    Private Sub grdSOTORDRR_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDRR.AfterRowUpdate
        Display_Totals_R(e.Row.Cells("ORDR_LNO").Value)

    End Sub

    Private Sub grdSOTORDRR_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDRR.BeforeCellUpdate
        If grdSOTORDRR.Rows.Count = 1 Then
            If Val(fpPPQTY.Text) = 1 Or Val(fpPPQTY.Text) = 0 Then
                SetRangeType("R")
            Else
                SetRangeType("A")
            End If
        End If
    End Sub

    Private Sub grdSOTORDRR_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTORDRR.BeforeExitEditMode
        If grdSOTORDRR.ActiveCell IsNot Nothing Then
            With grdSOTORDRR.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                        'If .EditorResolved.Value & "" <> CStr(.EditorResolved.Value & "").ToUpper Then
                        '    .EditorResolved.Value = CStr(.EditorResolved.Value & "").ToUpper
                        'End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTORDRR_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTORDRR.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
            Dim ORDR_LNO As Int32 = Val(grow.Cells("ORDR_LNO").Value & "")
            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
            sqlw &= " and RSRV_NO is not null"
            If dst.Tables("SOTORDRR").Select(sqlw).Length <> 0 Then
                MsgBox("Cannot Delete a Line if it has ever been " _
                       & vbCrLf & " used in a Reservation." _
                       & vbCrLf & vbCrLf & "Set Qty To Zero.")
                e.Cancel = True
                Exit Sub
            End If
        Next
    End Sub

    Private Sub grdSOTORDRR_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDRR.BeforeRowUpdate
        Validate_Columns_R("STYLE_CODE", e.Cancel)
        Validate_Columns_R("COLOR_CODE", e.Cancel)
        Validate_Columns_R("ORDR_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_LNO").Value = grdSOTORDR2.ActiveRow.Cells("ORDR_LNO").Value
            e.Row.Cells("ORDR_NO").Value = ORDR_NO
            e.Row.Cells("RANGE_STYLE_CODE").Value = grdSOTORDR2.ActiveRow.Cells("RANGE_STYLE_CODE").Value
            e.Row.Cells("ORDR_QTY_ORIG").Value = e.Row.Cells("ORDR_QTY").Value
            e.Row.Cells("ORDR_UNIT_PRICE").Value = grdSOTORDR2.ActiveRow.Cells("ORDR_UNIT_PRICE").Value
            e.Row.Cells("RANGE_STYLE_QTY_PER_PP").Value = fpPPQTY.Text
            e.Row.Cells("ORDR_STATUS").Value = "O"
        End If
    End Sub
#End Region

#Region "grdSOTORDRS"

    Private Sub grdSOTORDRS_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDRS.AfterCellUpdate
        msqty = Val(e.Cell.Value & "")
        msqty_col = Val(Split(e.Cell.Column.Key, "_")(1))
        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, msqty_col})

        With optMSCopyToStore.ValueList
            .ValueListItems(0).DisplayText = "Copy All Style/Colors from Store " & e.Cell.Row.Cells("CUST_STORE_NO").Value
            .ValueListItems(0).Tag = e.Cell.Row.Cells("CUST_STORE_NO").Value
            If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" Then
                .ValueListItems(1).DisplayText = "Copy Qty of " & CStr(msqty) & " to Range Style " & rowSOTORDR2.Item("RANGE_STYLE_CODE")
                .ValueListItems(1).Tag = rowSOTORDR2.Item("RANGE_STYLE_CODE")
            Else
                .ValueListItems(1).DisplayText = "Copy Qty of " & CStr(msqty) & " to Style/Color " & rowSOTORDR2.Item("STYLE_CODE") & "/" & rowSOTORDR2.Item("COLOR_CODE")
                .ValueListItems(1).Tag = rowSOTORDR2.Item("STYLE_CODE") & "/" & rowSOTORDR2.Item("COLOR_CODE")
            End If
            .ValueListItems(2).DisplayText = "Copy Qty of " & CStr(msqty) & " to All Styles"
        End With

        lblMSCopyToStore.Visible = True
        txtMSCopyToStore.Visible = True
        optMSCopyToStore.Visible = True
    End Sub

    Private Sub grdSOTORDRS_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTORDRS.AfterRowsDeleted
        ' Set_MS_TOTAL_AMT_Expression()
    End Sub

    Private Sub grdSOTORDRS_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDRS.AfterRowUpdate
        '  Set_MS_TOTAL_AMT_Expression()

        If e.Row.Cells("CUST_STORE_NO").Value = Absx1.txtFor("CUST_STORE_NO").Text And Not multi_store_changes_made_to_SOTORDRS Then
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR2.Rows
                If grow.IsAddRow Then
                Else
                    Dim ORDR_LNO As Int64 = Val(grow.Cells("ORDR_LNO").Value & "")
                    Dim ORDR_QTY As Int64 = Val(grow.Cells("ORDR_QTY").Value & "")
                    Dim QTY As Int64 = Val(e.Row.Cells("QTY_" & Format(ORDR_LNO, "000")).Value & "")
                    If Val(grow.Cells("ORDR_QTY").Value & "") <> QTY Then
                        multi_store_changes_made_to_SOTORDRS = True
                        grow.Cells("ORDR_QTY").Value = QTY
                        grow.Update()
                        multi_store_changes_made_to_SOTORDRS = False
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub grdSOTORDRS_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTORDRS.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("CUST_STORE_NO").Value = Absx1.txtFor("CUST_STORE_NO").Text Then
                MsgBox("You Cannot Delete the Main Store", MsgBoxStyle.OkOnly, "Cannot Delete")
                e.Cancel = True
                e.DisplayPromptMsg = False
            End If
        Next
    End Sub
#End Region

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCART1.AfterRowActivate
        Set_SOTCART2()
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        'Set_SOTPICK2()
        Set_SOTCART1()
    End Sub

    Sub Clear_Zeroes()
        Dim sqlw As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
            Dim COLUMN_NAME As String = "QTY_" & Format(ORDR_LNO, "000")
            sqlw &= " and (ISNULL(" & COLUMN_NAME & ",0) = 0)"
        Next
        sqlw = "CUST_STORE_NO <> '" & Absx1.txtFor("CUST_STORE_NO").Text & "'" & sqlw
        ASCDATA1.DeleteRows(dst.Tables("SOTORDRS"), sqlw)
    End Sub

    Sub Load_EDI_Documents()
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else

            dst.Tables("EDTDOCS1").Rows.Clear()

            Dim EDI_DOC_SEQ_NO As String = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
            If EDI_DOC_SEQ_NO <> "" Then
                ASCMAIN1.sql = "Select * from EDT850T1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                For Each rowEDT850T1 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim row As DataRow = TAC.SOCMAIN1.Get_EDI_row(EDI_DOC_SEQ_NO, "850")
                    If row IsNot Nothing Then
                        Dim DT As DataTable = row.Table
                        Dim rowEDTDOCS1 As DataRow = dst.Tables("EDTDOCS1").NewRow
                        With rowEDTDOCS1
                            .Item("EDI_DOC_ID") = "850"
                            .Item("EDI_DOC_DATE") = rowEDT850T1.Item("INIT_DATE")
                            .Item("EDI_DOC_SEQ_NO") = rowEDT850T1.Item("EDI_DOC_SEQ_NO")
                            .Item("FILENAME_GEN") = row.Item("DocumentBlobKEY")
                            .Item("FILENAME_ABS") = ""
                            .Item("EDI_DOC_TEXT") = ""
                            .Item("EDI_DOC_STATUS") = ""
                            .Item("EDI_DOC_DESC") = "Purchase Order"
                            .Item("EDI_ISA_NO") = ""
                        End With
                        dst.Tables("EDTDOCS1").Rows.Add(rowEDTDOCS1)
                    End If
                Next
            End If

            ASCMAIN1.sql = "Select * from EDT940O1 where PICK_NO in (Select PICK_NO from SOTPICK1 where ORDR_NO = '" & ORDR_NO & "')"
            For Each rowEDT940O1 As DataRow In ASCDATA1.GetDataTable.Rows
                Dim row As DataRow = TAC.SOCMAIN1.Get_EDI_row(rowEDT940O1.Item("EDI_OUTBOUND_DOC_NO"), "940")
                Dim rowEDTDOCS1 As DataRow = dst.Tables("EDTDOCS1").NewRow
                With rowEDTDOCS1
                    .Item("EDI_DOC_ID") = "940"
                    .Item("EDI_DOC_DATE") = rowEDT940O1.Item("INIT_DATE")
                    .Item("EDI_DOC_SEQ_NO") = rowEDT940O1.Item("EDI_OUTBOUND_DOC_NO")
                    If row IsNot Nothing Then .Item("FILENAME_GEN") = row.Item("DocumentBlobKEY")
                    .Item("FILENAME_ABS") = ""
                    .Item("EDI_DOC_TEXT") = ""
                    .Item("EDI_DOC_STATUS") = ""
                    .Item("EDI_DOC_DESC") = "Release Order"
                    .Item("EDI_ISA_NO") = ""
                End With
                dst.Tables("EDTDOCS1").Rows.Add(rowEDTDOCS1)
            Next

            ASCMAIN1.sql = "Select * from EDT855O1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            For Each rowEDT855O1 As DataRow In ASCDATA1.GetDataTable.Rows
                Dim row As DataRow = TAC.SOCMAIN1.Get_EDI_row(rowEDT855O1.Item("EDI_OUTBOUND_DOC_NO"), "855", ORDR_GROUP_NO)
                Dim rowEDTDOCS1 As DataRow = dst.Tables("EDTDOCS1").NewRow
                With rowEDTDOCS1
                    .Item("EDI_DOC_ID") = "855"
                    .Item("EDI_DOC_DATE") = rowEDT855O1.Item("INIT_DATE")
                    .Item("EDI_DOC_SEQ_NO") = rowEDT855O1.Item("EDI_OUTBOUND_DOC_NO")
                    If row IsNot Nothing Then .Item("FILENAME_GEN") = row.Item("DocumentBlobKEY")
                    .Item("FILENAME_ABS") = ""
                    .Item("EDI_DOC_TEXT") = ""
                    .Item("EDI_DOC_STATUS") = ""
                    Dim EDI_DOC_DESC As String = "Reverse EDI PO"
                    If ASCMAIN1.CLIENT = "NYA" Then EDI_DOC_DESC = "Credit Request"
                    .Item("EDI_DOC_DESC") = EDI_DOC_DESC
                    .Item("EDI_ISA_NO") = ""
                End With
                dst.Tables("EDTDOCS1").Rows.Add(rowEDTDOCS1)
            Next

        End If
    End Sub

    Sub Load_Events()
        Fill_Records("TATEVNT1", ORDR_NO)

        If EntryMode = "N" Then
            Record_Event("INIT", "Sales Order Entry Started")
        ElseIf EntryMode = "E" Then
            Record_Event("LAST", "Sales Order Edit Started")
        End If

        If ASCMAIN1.CLIENT = "NYA" Then
            ASCMAIN1.sql = "Select * from SOTAUTH1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable().Select("")
                Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
                With rowTATEVNT1
                    .Item("TABLE_NAME") = "SOTORDR0"
                    .Item("TABLE_KEY") = ORDR_GROUP_NO
                    .Item("EVENT_KEY") = row.Item("ORDR_CRED_CLR_AUTH_NO")
                    Dim ORDR_CRED_CLR_AUTH As String = row.Item("ORDR_CRED_CLR_AUTH") & ""
                    Dim z As String = ORDR_CRED_CLR_AUTH
                    If ORDR_CRED_CLR_AUTH = "A" Then z = "Appr"
                    If ORDR_CRED_CLR_AUTH = "H" Then z = "Hold"
                    If ORDR_CRED_CLR_AUTH = "D" Then z = "Decl"
                    .Item("EVENT_DESC") = "Credit Decision for " & Format(Val(row.Item("ORDR_CRED_CLR_AUTH_AMT") & ""), "$#,##0.00") & " - " & z
                    .Item("EVENT_TYPE") = "CR-DSP " & ORDR_CRED_CLR_AUTH
                    .Item("INIT_OPER") = row.Item("ORDR_CRED_CLR_BY")
                    .Item("INIT_DATE") = row.Item("INIT_DATE")
                End With
                dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
            Next
        End If

        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
    End Sub

    Private Sub txtMSStore_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtMSCopyToStore.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtMSCopyToStore.Text = "" Or msqty_col = 0 Then
                MsgBox("You Must First Enter a Store No")
                Exit Sub
            End If

            Dim CUST_STORE_NO_Copy_To As String = txtMSCopyToStore.Text.PadLeft(6, "0")
            If dst.Tables("SOTORDRS").Rows.Find(CUST_STORE_NO_Copy_To) Is Nothing Then
                MsgBox("Store " & CUST_STORE_NO_Copy_To & " Not in Multi-Store Grid" & vbCrLf & "You must enter a Store No to Copy To from the Stores listed above")
                txtMSCopyToStore.Text = ""
                Exit Sub
            End If

            Dim CUST_STORE_NO_Copy_From As String = optMSCopyToStore.ValueList.ValueListItems(0).Tag ' grdSOTORDRS.ActiveRow.Cells("CUST_STORE_NO").Value
            If grdSOTORDRS.ActiveRow Is Nothing OrElse
                grdSOTORDRS.ActiveRow.Cells("CUST_STORE_NO").Value <> CUST_STORE_NO_Copy_To Then
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRS.Rows
                    If grow.Cells("CUST_STORE_NO").Value = CUST_STORE_NO_Copy_To Then
                        grdSOTORDRS.ActiveRow = grow
                        Exit For
                    End If
                Next
            End If

            Dim rowSOTORDRS_Copy_From As DataRow = dst.Tables("SOTORDRS").Rows.Find(CUST_STORE_NO_Copy_From)
            lblMSCopyToStore.Tag = "x"

            ' Dim rowSOTORDRS As DataRow = dst.Tables("SOTORDRS").Rows.Find(txtMSCopyToStore.Text.PadLeft(6, "0"))
            If grdSOTORDRS.ActiveRow IsNot Nothing Then
                If optMSCopyToStore.Value = "Store" Then
                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                        Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim C As String = "QTY_" & Format(ORDR_LNO, "000")
                        grdSOTORDRS.ActiveRow.Cells(C).Value = rowSOTORDRS_Copy_From.Item(C)
                    Next
                    'For i As Integer = grdMS_Cols_Orig + 0 To dst.Tables("SOTORDRS").Columns.Count - 1
                    '    grdSOTORDRS.ActiveRow.Cells(i).Value = rowSOTORDRS_this.Item(i)
                    'Next i
                ElseIf optMSCopyToStore.Value = "Style" Then
                    grdSOTORDRS.ActiveRow.Cells("QTY_" & Format(msqty_col, "000")).Value = msqty
                ElseIf optMSCopyToStore.Value = "Qty" Then
                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                        Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim C As String = "QTY_" & Format(ORDR_LNO, "000")
                        grdSOTORDRS.ActiveRow.Cells(C).Value = msqty
                    Next
                    'For i As Integer = grdMS_Cols_Orig + 0 To dst.Tables("SOTORDRS").Columns.Count - 1
                    '    grdSOTORDRS.ActiveRow.Cells(i).Value = msqty
                    'Next i
                End If
                grdSOTORDRS.ActiveRow.Update()
                txtMSCopyToStore.Text = ""
            End If
        End If
    End Sub

    Sub Add_Colors(STYLE_CODE As String, tbl As DataTable, PRICE As Decimal)
        If tbl.Select("ISNULL(QTY,0)<>0").Length = 0 Then
            MsgBox("No Qty's Entered", MsgBoxStyle.OkOnly, "Cannot Add Colors")
            Exit Sub
        End If

        If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "SAM" Or Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR" Then
            PRICE = 0
        End If
        For Each rowICTCOLRM As DataRow In tbl.Select("ISNULL(QTY,0)<>0", "COLOR_CODE")
            Add_grdSOTORDR2(STYLE_CODE, rowICTCOLRM.Item("COLOR_CODE"), rowICTCOLRM.Item("QTY"), PRICE)
        Next
        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
    End Sub

    Sub Add_grdSOTORDR2(STYLE_CODE As String, COLOR_CODE As String,
                        ORDR_QTY As Int64, ORDR_UNIT_PRICE As Decimal)
        If grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.DataChanged Then
            grdSOTORDR2.ActiveRow.CancelUpdate()
        End If


        grdSOTORDR2.DisplayLayout.Bands(0).AddNew()
        With grdSOTORDR2.ActiveRow
            .Cells("STYLE_CODE").Value = STYLE_CODE
            .Cells("COLOR_CODE").Value = COLOR_CODE
            .Cells("ORDR_QTY").Value = ORDR_QTY
            If Absx1.optFor("ORDR_TYPE_CODE").Value = "XFR" Then
                .Cells("ORDR_UNIT_PRICE").Value = 0
            Else
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    If ORDR_UNIT_PRICE <> 0 Then
                        .Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE
                    End If
                Else
                    .Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE
                End If
            End If
            .Update()
        End With
    End Sub

    Function Check_Changed_Fields() As Boolean

        REV_NO += 1

        Dim LAST_DATE As Date = DATETIME_STAMP
        If EntryMode = "N" Then Stop
        Dim REV_LNO As Integer = 0

        Check_Changed_Fields = False

        dst.Tables("SOTORDXR").Rows.Clear()

        ASCMAIN1.Progress("Logging Header Changes")

        For i As Integer = 0 To rowSOTORDR1.Table.Columns.Count - 1
            Dim COLUMN_NAME As String = dst.Tables("SOTORDR1").Columns(i).ColumnName

            If rowSOTORDR1.Item(COLUMN_NAME) & "" _
            <> rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                Check_Changed_Fields = True
                ASCMAIN1.Progress("-", COLUMN_NAME)
                Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                With rowSOTORDXR
                    .Item("REV_NO") = REV_NO
                    REV_LNO += 1
                    .Item("REV_LNO") = REV_LNO
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = 0
                    .Item("INIT_DATE") = LAST_DATE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("COLUMN_NAME") = COLUMN_NAME
                    .Item("OLD_VALUE") = rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original)
                    .Item("NEW_VALUE") = rowSOTORDR1.Item(COLUMN_NAME)
                    .Item("EMODE") = EntryMode
                End With
                dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                Check_Changed_Fields = True
            End If
        Next i

        ASCMAIN1.Progress("Logging Detail Changes")

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each rowSOTORDR2_orig As DataRow In dt.Rows
            Dim ORDR_LNO As Int64 = rowSOTORDR2_orig.Item("ORDR_LNO")
            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
            If rowSOTORDR2 Is Nothing Then ' Line was Deleted
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowSOTORDR2_orig.Table.Columns(i).ColumnName
                    Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                    With rowSOTORDXR
                        .Item("REV_NO") = REV_NO
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("ORDR_NO") = ORDR_NO
                        .Item("ORDR_LNO") = ORDR_LNO
                        .Item("INIT_DATE") = LAST_DATE
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("OLD_VALUE") = rowSOTORDR2_orig.Item(COLUMN_NAME)
                        '.Item("NEW_VALUE") = ""
                        .Item("EMODE") = EntryMode
                        Dim CONTEXT As String
                        If rowSOTORDR2_orig.Item("RANGE_STYLE_CODE") & "" <> "" Then
                            CONTEXT = rowSOTORDR2_orig.Item("RANGE_STYLE_CODE")
                        Else
                            CONTEXT = rowSOTORDR2_orig.Item("STYLE_CODE") & "/" & rowSOTORDR2_orig.Item("COLOR_CODE")
                        End If
                        .Item("CONTEXT") = CONTEXT
                    End With
                    dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                Next

                Check_Changed_Fields = True
            Else
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowSOTORDR2_orig.Table.Columns(i).ColumnName
                    If rowSOTORDR2.Item(COLUMN_NAME) & "" <> rowSOTORDR2_orig.Item(COLUMN_NAME) & "" Then
                        ' Value in Column was Changed
                        Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                        With rowSOTORDXR
                            .Item("REV_NO") = REV_NO
                            REV_LNO += 1
                            .Item("REV_LNO") = REV_LNO
                            .Item("ORDR_NO") = ORDR_NO
                            .Item("ORDR_LNO") = ORDR_LNO
                            .Item("INIT_DATE") = LAST_DATE
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("COLUMN_NAME") = COLUMN_NAME
                            .Item("OLD_VALUE") = rowSOTORDR2_orig.Item(COLUMN_NAME)
                            .Item("NEW_VALUE") = rowSOTORDR2.Item(COLUMN_NAME)
                            .Item("EMODE") = EntryMode
                            Dim CONTEXT As String
                            If rowSOTORDR2_orig.Item("RANGE_STYLE_CODE") & "" <> "" Then
                                CONTEXT = rowSOTORDR2_orig.Item("RANGE_STYLE_CODE")
                            Else
                                CONTEXT = rowSOTORDR2_orig.Item("STYLE_CODE") & "/" & rowSOTORDR2_orig.Item("COLOR_CODE")
                            End If
                            '  .Item("CONTEXT") = CONTEXT
                        End With
                        dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                        Check_Changed_Fields = True
                    End If
                Next
            End If
        Next

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "", DataViewRowState.Added)
            Dim ORDR_LNO = rowSOTORDR2.Item("ORDR_LNO")
            ' For i As Integer = 0 To dt.Columns.Count - 1
            Dim COLUMN_NAME As String = "" ' dt.Columns(i).ColumnName
            Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
            With rowSOTORDXR
                .Item("REV_NO") = REV_NO
                REV_LNO += 1
                .Item("REV_LNO") = REV_LNO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = ORDR_LNO
                .Item("INIT_DATE") = LAST_DATE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("COLUMN_NAME") = COLUMN_NAME
                '.Item("OLD_VALUE") = ""
                .Item("NEW_VALUE") = "PO Line Added" ' rowSOTORDR2.Item(COLUMN_NAME)
                .Item("EMODE") = EntryMode
                Dim CONTEXT As String
                If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" Then
                    CONTEXT = rowSOTORDR2.Item("RANGE_STYLE_CODE")
                Else
                    CONTEXT = rowSOTORDR2.Item("STYLE_CODE") & "/" & rowSOTORDR2.Item("COLOR_CODE")
                End If
                .Item("CONTEXT") = CONTEXT
            End With
            dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
            Check_Changed_Fields = True
            'Next
        Next

        ASCMAIN1.Progress("")
        Return Check_Changed_Fields
    End Function

    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String)
        Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
        rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
        rowTATEVNT1.Item("TABLE_KEY") = ORDR_NO
        rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
        rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowTATEVNT1.Item("EVENT_TYPE") = EVENT_TYPE
        rowTATEVNT1.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
    End Sub

    Private Sub grdICTSIZE1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        Add_Selected_Size(e.Cell.Row.Cells("SIZE_CODE").Value)
    End Sub

    Sub Add_Selected_Size(SIZE_CODE As String)

        Dim SIZE_INDEX As Integer = 0

        For I As Integer = 1 To 12
            Dim SIZE_DESC As String = grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(I, "00")).Value & ""
            If SIZE_DESC = "" And SIZE_INDEX = 0 Then
                SIZE_INDEX = I
            End If
            If SIZE_DESC = SIZE_CODE Then
                If MsgBox("Do you want to Remove this Size from the Components Grid?",
                       MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Duplicate Size Detected") = MsgBoxResult.Yes Then
                    Dim SIZE_INDEX_MAX As Integer = I
                    For j As Integer = I + 1 To 12
                        Dim SIZE_CODE_j As String = grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(j, "00")).Value & ""
                        If SIZE_CODE_j <> "" Then
                            SIZE_INDEX_MAX = j
                        End If
                        grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(j - 1, "00")).Value = SIZE_CODE_j
                        grdSOTORDR3.DisplayLayout.Bands(0).Columns("SIZE_QTY_" & Format(j - 1, "00")).Header.Caption = SIZE_CODE_j
                    Next
                    grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(SIZE_INDEX_MAX, "00")).Value = DBNull.Value
                    With grdSOTORDR3.DisplayLayout.Bands(0).Columns("SIZE_QTY_" & Format(SIZE_INDEX_MAX, "00"))
                        .Hidden = True
                        .Header.Caption = ""
                    End With
                    For Each ROWSOTORDR3 As DataRow In dst.Tables("SOTORDR3").Select("")
                        For j As Integer = I + 1 To 12
                            ROWSOTORDR3.Item("SIZE_QTY_" & Format(j - 1, "00")) = ROWSOTORDR3.Item("SIZE_QTY_" & Format(j, "00"))
                        Next
                        ROWSOTORDR3.Item("SIZE_QTY_" & Format(SIZE_INDEX_MAX, "00")) = DBNull.Value
                    Next

                    cbeICTSIZE1.Value = DBNull.Value

                End If

                Exit Sub
            End If
        Next
        If SIZE_INDEX = 0 Then
            MsgBox("All 12 Size Slots are used", MsgBoxStyle.OkOnly, "Could not perform requested action")
            Exit Sub
        End If

        With grdSOTORDR3.DisplayLayout.Bands(0).Columns("SIZE_QTY_" & Format(SIZE_INDEX, "00"))
            .Hidden = False
            .Header.Caption = SIZE_CODE
        End With
        grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(SIZE_INDEX, "00")).Value = SIZE_CODE
        cbeICTSIZE1.Value = DBNull.Value
    End Sub

    Private Sub cbeICTSIZE1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles cbeICTSIZE1.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            If cbeICTSIZE1.Value & "" <> "" Then
                Add_Selected_Size(cbeICTSIZE1.Value)
            End If
        End If
    End Sub

    Private Sub cbeICTSIZE1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cbeICTSIZE1.ValueChanged
        Dim size_is_already_in_grid As Boolean = False
        Dim SIZE_CODE As String = cbeICTSIZE1.Value & ""
        If SIZE_CODE = "" Then
            cmdAddSize.Visible = False
        Else
            For i As Integer = 1 To 12
                If grdSOTORDR2.ActiveRow.Cells("SIZE_DESC_" & Format(i, "00")).Value & "" = SIZE_CODE Then
                    size_is_already_in_grid = True
                    Exit For
                End If
            Next

            If size_is_already_in_grid Then
                cmdAddSize.Text = "Remove Size " & SIZE_CODE
            Else
                cmdAddSize.Text = "Add Size " & SIZE_CODE
            End If
            cmdAddSize.Visible = True
        End If
    End Sub

    Private Sub cmdAddSize_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddSize.Click
        If cbeICTSIZE1.Value & "" <> "" Then
            Add_Selected_Size(cbeICTSIZE1.Value)
        End If
    End Sub

    Private Sub optShowOrders_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShowOrders.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        optCustomerOrders.Visible = (optShowOrders.Value = "C" OrElse optShowOrders.Value = "N")

        If optShowOrders.Value = "N" Then
            txtCustNameSearch.Enabled = True
        Else
            txtCustNameSearch.Enabled = False
            txtCustNameSearch.Clear()
        End If

        Load_SOTORDRX()
    End Sub

    Private Sub grdSOTORDR2_Error(sender As Object, e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSOTORDR2.Error
        MsgBox(e.ErrorText, MsgBoxStyle.OkOnly, "Cannot Update Order Detail")
        e.Cancel = True
        'grdSOTORDR2.ActiveRow.CancelUpdate()
        ' NASTY ERROR SHOWS HERE AND IS A PRECURSOR TO ERROR WHEN CLICKING UPDATE: Column 'ORDR_NO' does not allow nulls.
        ' CAN MAKE IT HAPPEN BY ENTERING A STYLE, COLOR, QTY, ND THEN CLICK INTO THE GRID
        ' NOT SURE WHY CLICKING INTO THE GRID CAUSES THIS ERROR
        ' TRIED OTHER FORMS (ICFIADJ1) AND DO NOT GET SAME BEHAVIOR
        ' MUST BE A PROPERTY, OR CODE IN ONE OF THESE EVENT PROCEDRES

        grdSOTORDR2.Rows.Refresh(RefreshRow.ReloadData)
    End Sub

    Private Sub grdSOTORDR2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR2.InitializeRow

        Dim CARTON_PACK_QTY As Integer = Val(e.Row.Cells("CARTON_PACK_QTY").Value & "")
        Dim INNER_PACK_QTY As Integer = Val(e.Row.Cells("INNER_PACK_QTY").Value & "")
        Dim ORDR_QTY As Int32 = Val(e.Row.Cells("ORDR_QTY").Value & "")
        ' Dim ORDR_QTY_OPEN As Int32 = Val(e.Row.Cells("ORDR_QTY_OPEN").Value & "")

        Dim MIN_ORD_QTY As Integer = INNER_PACK_QTY
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            ' MIN_ORD_QTY = CARTON_PACK_QTY
        End If

        Dim red_if_not_case_pack As Boolean = False
        If ASCMAIN1.CLIENT = "NYA" Then
            red_if_not_case_pack = True
        End If

        If (MIN_ORD_QTY <> 0 And ORDR_QTY < MIN_ORD_QTY) _
        Or (INNER_PACK_QTY <> 0 AndAlso ORDR_QTY Mod MIN_ORD_QTY <> 0) Then
            e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Red
            If ORDR_QTY < MIN_ORD_QTY Then
                e.Row.Cells("ORDR_QTY").ToolTipText = "Min Order Qty is " & CStr(MIN_ORD_QTY)
            Else
                e.Row.Cells("ORDR_QTY").ToolTipText = "Order Qty not Divisible by Inner Pack Qty"
            End If

        ElseIf (Absx1.optFor("ORDR_TYPE_CODE").Value = "BTB" Or red_if_not_case_pack) And
            ((CARTON_PACK_QTY <> 0 And ORDR_QTY < CARTON_PACK_QTY) Or
             (CARTON_PACK_QTY <> 0 AndAlso ORDR_QTY Mod CARTON_PACK_QTY <> 0)) Then
            e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Red
            If ORDR_QTY < CARTON_PACK_QTY Then
                e.Row.Cells("ORDR_QTY").ToolTipText = "Min Order Qty is " & CStr(CARTON_PACK_QTY)
            Else
                e.Row.Cells("ORDR_QTY").ToolTipText = "Order Qty not Divisible by Case Pack Qty"
            End If
        Else
            e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("ORDR_QTY").ToolTipText = ""
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            Dim STYLE_ASST_QTY As Integer = Val(e.Row.Cells("STYLE_ASST_QTY").Value & "")
            If STYLE_ASST_QTY > 0 Then
                If ORDR_QTY Mod STYLE_ASST_QTY <> 0 Then
                    e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Red
                    If e.Row.Cells("ORDR_QTY").ToolTipText.Length > 0 Then
                        e.Row.Cells("ORDR_QTY").ToolTipText += " And Order Qty not Divisible by Assortment of " & STYLE_ASST_QTY
                    Else
                        e.Row.Cells("ORDR_QTY").ToolTipText = "Order Qty not Divisible by Assortment of " & STYLE_ASST_QTY
                    End If
                End If
            End If
        End If

        Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""

        If STYLE_CODE <> "" Then
            Dim STYLE_STATUS As String = e.Row.Cells("STYLE_STATUS").Value & ""
            Dim STYLE_COLOR_STATUS As String = e.Row.Cells("STYLE_COLOR_STATUS").Value & ""
            If ASCMAIN1.CLIENT = "RGI" Then
                Dim ss As New Text.StringBuilder With {.Length = 0}
                If STYLE_STATUS.Length = 0 Then
                    ss.Length = 0
                    ss.AppendLine("SELECT STYLE_STATUS")
                    ss.AppendLine("FROM ICTSTYL1")
                    ss.AppendLine($"WHERE STYLE_CODE = '{STYLE_CODE}'")
                    ASCMAIN1.sql = ss.ToString()
                    STYLE_STATUS = ASCDATA1.GetDataValue
                End If
                If STYLE_COLOR_STATUS.Length = 0 Then
                    ss.Length = 0
                    ss.AppendLine("SELECT STYLE_COLOR_STATUS")
                    ss.AppendLine("FROM ICTSTYC1")
                    ss.AppendLine($"WHERE STYLE_CODE = '{STYLE_CODE}'")
                    ss.AppendLine($"AND COLOR_CODE = '{COLOR_CODE}'")
                    ASCMAIN1.sql = ss.ToString()
                    STYLE_COLOR_STATUS = ASCDATA1.GetDataValue
                End If
            End If
            If STYLE_STATUS = "D" Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("STYLE_CODE").ToolTipText = "Style is Discontinued"
            ElseIf STYLE_STATUS = "N" Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.DarkOrange
                e.Row.Cells("STYLE_CODE").ToolTipText = "Style is Do Not Re-Order"
            Else
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("STYLE_CODE").ToolTipText = ""
            End If
            If STYLE_COLOR_STATUS = "D" Then
                e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("COLOR_CODE").ToolTipText = "Color is Discontinued"
            ElseIf STYLE_COLOR_STATUS = "N" Then
                e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Drawing.Color.DarkOrange
                e.Row.Cells("COLOR_CODE").ToolTipText = "Color is Do Not Re-Order"
            Else
                e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("COLOR_CODE").ToolTipText = ""
            End If
        Else
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("STYLE_CODE").ToolTipText = ""
            e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("COLOR_CODE").ToolTipText = ""
        End If

        e.Row.Cells("ORDR_QTY_ALLO").Appearance.ForeColor = Drawing.Color.Empty
        e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.ForeColor = Drawing.Color.Empty
        e.Row.Cells("ORDR_QTY_ALLO").ToolTipText = ""
        e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = ""
        Dim ORDR_QTY_OPEN As Int64 = Val(e.Row.Cells("ORDR_QTY_OPEN").Value & "")
        Dim ORDR_QTY_ALLO As Int64 = Val(e.Row.Cells("ORDR_QTY_ALLO").Value & "")
        'If ORDR_QTY_OPEN = 100 And ORDR_QTY_ALLO = 0 Then
        '    Stop
        'End If
        If ORDR_QTY_OPEN <> 0 And ORDR_QTY_ALLO = 0 Then
            If EntryMode = "V" Then
                e.Row.Cells("ORDR_QTY_ALLO").Appearance.BackColor = Drawing.Color.Red
                e.Row.Cells("ORDR_QTY_ALLO").ToolTipText = "Inventory Shortage"
            End If
        Else

            'If ASCMAIN1.CLIENT = "RGI" Then
            '    Dim QTY_1 As Int32 = Val(e.Row.Cells("QTY_1").Value & "")
            '    Dim QTY_2 As Int32 = Val(e.Row.Cells("QTY_2").Value & "")
            '    Dim QTY_3 As Int32 = Val(e.Row.Cells("QTY_3").Value & "")
            '    Dim QTY_4 As Int32 = Val(e.Row.Cells("QTY_4").Value & "")
            '    If ORDR_QTY_OPEN > QTY_1 + QTY_2 + QTY_3 + QTY_4 Then
            '        e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.BackColor = Drawing.Color.Orange
            '    End If
            'End If

            If e.Row.Cells("ORDR_RELEASE_AVAIL").Value & "" <> "" Then
                Dim ORDR_RELEASE_AVAIL As Date = e.Row.Cells("ORDR_RELEASE_AVAIL").Value
                If Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" <> "" Then
                    If Format(ORDR_RELEASE_AVAIL, "yyyyMMdd") > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                        e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.ForeColor = Drawing.Color.Red
                        e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = "Availability date is > Cancel Date"
                    End If
                End If
            Else
                If ASCMAIN1.CLIENT = "RGI" Then
                    Dim QTY_1 As Int32 = Val(e.Row.Cells("QTY_1").Value & "")
                    Dim QTY_2 As Int32 = Val(e.Row.Cells("QTY_2").Value & "")
                    Dim QTY_3 As Int32 = Val(e.Row.Cells("QTY_3").Value & "")
                    Dim QTY_4 As Int32 = Val(e.Row.Cells("QTY_4").Value & "")

                    If ORDR_QTY_OPEN <> 0 Then
                        If QTY_1 <> 0 And QTY_2 = 0 And QTY_3 = 0 And QTY_4 = 0 Then
                            e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.BackColor = Drawing.Color.LightGreen
                            e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = "Green: Allocation is Drawing from Current Stock O/H"
                        ElseIf QTY_1 = 0 And QTY_2 = 0 And QTY_3 = 0 And QTY_4 = 0 Then
                            e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.BackColor = Drawing.Color.Orange
                            e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = "Orange: No Availability"
                        Else
                            e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.BackColor = Drawing.Color.Yellow
                            e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = "Yellow: Check Style/Color Allocation below"
                        End If
                    End If

                Else

                    If Val(e.Row.Cells("ORDR_QTY_ALLO").Value & "") <> 0 Then
                        e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.BackColor = Drawing.Color.LightGreen
                        e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = "Green: Allocation is Drawing from Current Stock O/H"
                    End If

                End If
            End If

        End If

        Dim MU_PCT As Decimal = Val(e.Row.Cells("MU_PCT").Value & "")
        If MU_PCT < 20 Then
            e.Row.Cells("MU_PCT").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("MU_PCT").ToolTipText = "MU is less than 20%"
        ElseIf MU_PCT > 40 Then
            e.Row.Cells("MU_PCT").Appearance.BackColor = Drawing.Color.LightGreen
            e.Row.Cells("MU_PCT").ToolTipText = "MU is greater than 40%"
        Else
            e.Row.Cells("MU_PCT").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("MU_PCT").ToolTipText = ""
        End If

    End Sub

    Private Sub grdSOTORDR2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles grdSOTORDR2.KeyDown
        With grdSOTORDR2
            Try
                If e.KeyCode = Keys.F5 Then
                    If .ActiveCell IsNot Nothing Then
                        Select Case .ActiveCell.Column.Key
                            Case "STYLE_CODE"
                                grdSOTORDR2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                                .ActiveCell.Value = STYLE_CODE_last_entry
                                .ActiveCell.SelStart = Len(grdSOTORDR2.ActiveCell.Text)
                        End Select

                    End If
                End If
            Catch ex As Exception

            End Try
        End With
    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow
        If e.Row.Cells("CUST_CREDIT_HOLD").Value & "" = "1" Then
            e.Row.Appearance.BackColor = Drawing.Color.Red
        ElseIf e.Row.Cells("ORDR_HOLD").Value & "" = "1" Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Order_History(STYLE_CODE As String, COLOR_CODE As String)
        ASCMAIN1.sql = sqlSOTORDRH _
                    & " and SOTORDR1.CUST_CODE = '" & CUST_CODE & "'" _
                    & " and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" _
                    & IIf(COLOR_CODE = "", "", " and SOTORDR2.COLOR_CODE = '" & Trim(COLOR_CODE) & "'")
        ' GETTING ORA-01756: quoted string not properly terminated
        ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ")"
        Try
            Fill_Records("SOTORDRH", "", True, ASCMAIN1.sql)
        Catch ex As Exception
            dst.Tables("SOTORDRH").Rows.Clear()
        End Try
        Sort_grdColumns(grdSOTORDRH, "ORDR_DATE".ToLower)
        grdSOTORDRH.Text = "Order History for Style " & STYLE_CODE & IIf(COLOR_CODE = "", "", ", Color " & COLOR_CODE)
    End Sub

    Sub Price_and_Availability(
        STYLE_CODE As String,
        STYLE_CLASS_CODE As String,
        COLOR_CODE As String,
        CARTON_PACK_QTY As Int64,
        STYLE_PRICE As Decimal)

        ASCMAIN1.sql = "Select * from ICTSTDQ2 " _
            & " where STYLE_CODE = '" & STYLE_CODE & "'" _
            & IIf(COLOR_CODE = "", "", " and COLOR_CODE = '" & COLOR_CODE & "'")
        Fill_Records("ICTSTDQ2", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdICTSTDQ2, "WHSE_CODE,COLOR_CODE")
        grdICTSTDQ2.Text = "Availability for Style " & STYLE_CODE


        If ASCMAIN1.CLIENT = "RGI" Then
            If grdSOTORDR2.ActiveRow Is Nothing OrElse grdSOTORDR2.ActiveRow.IsAddRow OrElse Not grdSOTORDR2.ActiveRow.IsDataRow Then
                grdICTSTDQ3.Visible = False
            Else
                ASCMAIN1.sql = "Select * from ICTSTDQ3 " _
                    & " where ORDG_GROUP_NO = STYLE_CODE = '" & STYLE_CODE & "'" _
                    & IIf(COLOR_CODE = "", "", " and COLOR_CODE = '" & COLOR_CODE & "'")
                Fill_Records("ICTSTDQ3", New String() {ORDR_NO, STYLE_CODE, COLOR_CODE})
                'Dim dvw As DataView = DirectCast(grdICTSTDQ3.DataSource, DataTable).DefaultView
                'dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
                'Sort_grdColumns(grdICTSTDQ3, "STYLE_CODE,COLOR_CODE")
                grdICTSTDQ3.Text = "Allocation for Style/Color " & STYLE_CODE & "/" & COLOR_CODE
                grdICTSTDQ3.Visible = True
            End If
        End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Fill_Records("ICTSTAT2", New String() {STYLE_CODE, "AST"})
            grdICTSTAT2.Text = "Style Status for " & STYLE_CODE & ":" & COLOR_CODE
            grdICTSTAT2.Visible = True
        End If

        grdICTPRICX.Text = "Price List for Style " & STYLE_CODE & IIf(COLOR_CODE = "", "", ", Color " & COLOR_CODE)
        dst.Tables("ICTPRICX").Rows.Clear()
        Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)
        If rowICTCLAS1 IsNot Nothing Then
            Dim rowICTDISC1 As DataRow = dst.Tables("ICTDISC1").Rows.Find(rowICTCLAS1.Item("DISC_CODE") & "")
            If rowICTDISC1 IsNot Nothing Then
                For I As Integer = 1 To 4
                    Dim rowICTPRICX As DataRow = dst.Tables("ICTPRICX").NewRow
                    With rowICTPRICX
                        .Item("TIER") = I
                        .Item("PCT") = rowICTDISC1.Item("DISC" & CStr(I) & "_PCT")
                        .Item("DESC") = rowICTDISC1.Item("DISC" & CStr(I) & "_DESC")
                        .Item("CASES") = rowICTDISC1.Item("DISC" & CStr(I) & "_CASES")
                        .Item("ABBR") = rowICTDISC1.Item("DISC" & CStr(I) & "_ABBR")
                        .Item("QTY") = CARTON_PACK_QTY * Val(rowICTDISC1.Item("DISC" & CStr(I) & "_CASES"))
                        .Item("PRICE") = STYLE_PRICE * (100 - Val(rowICTDISC1.Item("DISC" & CStr(I) & "_PCT"))) / 100

                        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                            If Absx1.optFor("ORDR_TYPE_CODE").Value = "B2C" Then
                                ' DO NONE OF THIS SILLY PRICE STUFF FROR B2C ORDERS
                            Else
                                Dim QTY As Integer = Val(.Item("QTY") & "")
                                If QTY = 0 Then QTY = 1
                                Dim ORDR_UNIT_PRICE_CALC As Decimal = TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1,
                                      grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "",
                                      grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & "",
                                      QTY, "")
                                .Item("CUST_PRICE") = ORDR_UNIT_PRICE_CALC
                            End If
                        End If
                    End With
                    dst.Tables("ICTPRICX").Rows.Add(rowICTPRICX)
                Next
            End If
        End If
    End Sub

    Sub Credit_Card(Optional TRAN_TYPE As String = "A", Optional AUTH_CCPA_NO As String = "", Optional TRANS_NO As Int16 = 0, Optional AdditionalAuthorization As Boolean = False)

        Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text.Trim
        Dim FRT_TERMS As String = MyBase.Absx1.txtFor("FRT_TERMS").Text.Trim
        Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text.Trim
        Dim freightCost As Decimal = 0
        Dim rowSOTORDR1X As DataRow = Nothing
        Dim chargeAgainstAuth As Boolean = False
        Dim rowARTCCPA1_AUTH As DataRow = Nothing
        Dim rowSOTORDC1 As DataRow = Nothing

        If TRAN_TYPE = "C" Then
            chargeAgainstAuth = True
            TRAN_TYPE = "S"
            rowARTCCPA1_AUTH = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = '" & AUTH_CCPA_NO & "'")
            rowSOTORDC1 = dst.Tables("SOTORDC1").Rows.Find(New Object() {ORDR_NO, TRANS_NO})
            If rowARTCCPA1_AUTH Is Nothing OrElse rowARTCCPA1_AUTH.Item("CCPA_STATUS") <> "T" Then
                MessageBox.Show("The selected CC line item is not an Authorization.", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        ElseIf AUTH_CCPA_NO.Length > 0 Then
            rowARTCCPA1_AUTH = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = '" & AUTH_CCPA_NO & "'")
        End If

        ' Here are some rules
        If dst.Tables("SOTORDR1").Rows.Count = 0 Then
            MessageBox.Show("Invalid or Missing Sales Order Number.", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        rowSOTORDR1X = dst.Tables("SOTORDR1").Rows(0)

        If Not AdditionalAuthorization Then
            If rowSOTORDR1X.Item("CCPA_NO") & String.Empty <> String.Empty AndAlso Not chargeAgainstAuth Then
                Dim dispMessage As Boolean = True

                ' Allow RGI to do a second, third, ... Authorization since they make multi shipments and only authorize what they need to ship
                ' the avaiable product
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim row As DataRow = ASCDATA1.GetDataRow("select * from artccpa1 where CCPA_NO_AUTH = '" & rowSOTORDR1X.Item("CCPA_NO") & "'")
                    If row IsNot Nothing Then
                        dispMessage = False
                    End If
                End If
                If dispMessage Then
                    MessageBox.Show("This sales order has an existing credit card authorization. You are not permitted to authorize additional funds.",
                        "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            End If
        End If

        'RGI Validates Credit Cards on the Web, no actual Credit Card Authorizations
        ' 11/22/2022
        'If rowSOTORDR1X.Item("ORDR_SOURCE") & String.Empty = "W" AndAlso ASCMAIN1.CLIENT = "RGI" Then
        '    MessageBox.Show("Web sales credit card authorization was processed on the website. You are not permitted to authorize additional funds.",
        '        "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Exit Sub
        'End If

        If Not ",O,P,".Contains(rowSOTORDR1X.Item("ORDR_STATUS") & String.Empty) Then
            MessageBox.Show("Only Open and In-Pick statuses can perform a credit card Authorization. If the order has been shipped, you may charge the credit card in Customer Inquiry.",
                "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If EntryMode = "V" AndAlso Not InquiryMode Then
            ASCMAIN1.sql = "Select Count (*) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS in ('O','P','F')"
            If Val(ASCDATA1.GetDataValue) > 1 Then
                MessageBox.Show("You Cannot perform credit card processing on a Multiple Order Group", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        EMsg = String.Empty
        If FRT_TERMS.Length > 0 Then
            If ASCDATA1.GetDataRow("select * from astcode1 where TABLE_NAME = 'SOTORDR1' AND COLUMN_NAME = 'FRT_TERMS' AND T_CODE = '" & FRT_TERMS & "'") Is Nothing Then
                EMsg &= vbCr & "Freight Terms are required to process a credit card."
            End If
        Else
            EMsg &= vbCr & "Freight Terms are required to process a credit card."
        End If

        If SHIP_VIA_CODE.Length > 0 Then
            If ASCDATA1.GetDataRow("SELECT * FROM SOTSVIA1 WHERE SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'") Is Nothing Then
                EMsg &= vbCr & "Ship Via Code is required for credit card processing."
            End If
        Else
            EMsg &= vbCr & "Ship Via Code is required for credit card processing."
        End If


        'If Not IsDate(MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value) Then
        '    EMsg &= vbCr & "Ship date is required for credit card processing."
        'End If

        'If CDate(MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value) < DateTime.Now Then
        '    If EntryMode <> "V" Then
        '        EMsg &= vbCr & "Ship date must be greater equal to today for credit card processing."
        '    End If
        'End If


        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow("select sotcarr1.carrier_type" _
                                                         & " from sotsvia1, sotcarr1" _
                                                         & " where sotsvia1.carrier_code = sotcarr1.carrier_code" _
                                                         & " and ship_via_code = :PARM1", "V", New Object() {SHIP_VIA_CODE})


        If rowSOTCARR1 Is Nothing Then
            EMsg &= vbCr & "Could not determine carrier for the Ship Via Code."
        End If

        If EMsg.Length > 0 Then
            MessageBox.Show(EMsg, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim ORDR_TOTAL_AMT As Decimal = 0

        If chargeAgainstAuth AndAlso rowSOTORDC1 IsNot Nothing Then
            ORDR_TOTAL_AMT = Val(rowSOTORDC1.Item("BALANCE") & String.Empty)
        Else
            ' Fedex, UPS and similar pay for freight when freight terms of PPA 
            If rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty = "U" AndAlso FRT_TERMS.ToUpper = "PPA" Then
                ' New Rule 1/24/2013. 20% or $20 the greater of the two
                freightCost = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "ORDR_STATUS <> 'F' AND ORDR_STATUS <> 'C'  AND ORDR_STATUS <> 'V'") & String.Empty) * 0.2
                If freightCost < 20 Then
                    freightCost = 20
                End If
            End If

            ORDR_TOTAL_AMT += Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "ORDR_STATUS <> 'F' AND ORDR_STATUS <> 'C'  AND ORDR_STATUS <> 'V'") & String.Empty)

            Dim MiscCalc As Decimal = 0
            If ASCMAIN1.CLIENT = "RGI" AndAlso TRAN_TYPE = "A" Then
                Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOROFSURCHG WHERE ORDR_NO = :PARM1", "V", New Object() {HFs("ORDR_NO")})
                If rowSOTORDR1 IsNot Nothing Then
                    'MiscCalc = Math.Round(ORDR_TOTAL_AMT * 0.045, 2)
                    'MiscCalc = Math.Round(ORDR_TOTAL_AMT * 0.1, 2) 'Changed to 10% Per Rich 6/18/21 W.R.
                    MiscCalc = Math.Round(ORDR_TOTAL_AMT * 0.25, 2) 'Changed to 25% Per Rich 8/22/21 W.R.
                End If
                Dim rowSOTORDR1_2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOROFSURCHG2 WHERE ORDR_NO = :PARM1", "V", New Object() {HFs("ORDR_NO")})
                If rowSOTORDR1_2 IsNot Nothing Then
                    MiscCalc = Math.Round(ORDR_TOTAL_AMT * 0.1, 2)
                End If
            End If

            ORDR_TOTAL_AMT += MiscCalc

            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
            If Not (rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1") Then
                ORDR_TOTAL_AMT += freightCost
            End If
        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            Dim zmsg As String = String.Empty
            Dim authorizeFunds As Decimal = 0
            Dim inPick As Decimal = Val(dst.Tables("SOTORDRT").Select("KEY = 4")(0).Item("AMT") & String.Empty)
            If inPick = 0 Then
                inPick = Val(dst.Tables("SOTORDRT").Select("KEY = 2")(0).Item("AMT") & String.Empty)
            End If
            Dim inputValue As String = InputBox("Enter the amount of merchandise funds you want to authorize.", "Additional Funds", Format(inPick, "#,##0.00")) & String.Empty
            inputValue = inputValue.Trim
            If inputValue.Length = 0 Then inputValue = 0

            authorizeFunds = CDec(inputValue)

            Select Case authorizeFunds
                Case Is < 0
                    MessageBox.Show("On Account amount must be greater than $0.00", "On Account", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                Case 0
                    ' Assume the user clicked cancel
                    Exit Sub
                Case Is > 0
                    ORDR_TOTAL_AMT = authorizeFunds
                    zmsg = "The merchandise amount you entered is " & Format$(authorizeFunds, "$#,##0.00") & Environment.NewLine
                    freightCost = authorizeFunds * 0.2
                    If freightCost < 20 Then
                        freightCost = 20
                    End If
                    zmsg &= "The anticipated Freight (20%) is " & Format$(freightCost, "$#,##0.00") & Environment.NewLine & Environment.NewLine
                    ORDR_TOTAL_AMT += freightCost
                    zmsg &= "Do you want to add the freight the to Authorize amount for a total of: " & Format$(ORDR_TOTAL_AMT, "$#,##0.00")

                    If MessageBox.Show(zmsg, "Authorize", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                        ORDR_TOTAL_AMT = authorizeFunds
                    End If
            End Select

        End If

        If ORDR_TOTAL_AMT <= 0 Then
            MessageBox.Show("You cannot charge $0.00 for sales Order No: " & ORDR_NO, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Not ASCMAIN1.Logical_Lock("ARTCUSTC", CUST_CODE, , , , 1) Then Exit Sub
        If Not ASCMAIN1.Logical_Open("ARTCCPA1", "*", , , , 1) Then Exit Sub
        Get_PARM("SOTPARM1")

        Using frmCCProcessor As New TAC.TAFCARDF(Me)
            frmCCProcessor.test_mode = ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
            frmCCProcessor.CUST_CODE = CUST_CODE
            frmCCProcessor.CCPA_REASON = IIf(chargeAgainstAuth, "C", "O")
            frmCCProcessor.ORDR_NO = ORDR_NO
            frmCCProcessor.TRAN_TYPE = TRAN_TYPE

            ' Set to capture 
            If chargeAgainstAuth Then
                frmCCProcessor.rowARTCCPA1 = rowARTCCPA1_AUTH
            ElseIf rowARTCCPA1_AUTH IsNot Nothing Then
                Dim COLs() As String = {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_LAST4", "CUST_CREDIT_CARD_KEY" _
                 , "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE" _
                 , "CUST_CREDIT_CARD_NAME", "CUST_CREDIT_CARD_ADDR1" _
                 , "CUST_CREDIT_CARD_CITY", "CUST_CREDIT_CARD_STATE", "CUST_CREDIT_CARD_ZIP_CODE", "CUST_CREDIT_CARD_COUNTRY"}
                With frmCCProcessor.rowARTCCPA1
                    For Each COLUMN_NAME As String In COLs
                        .Item(COLUMN_NAME) = rowARTCCPA1_AUTH.Item(COLUMN_NAME)
                    Next
                End With
            End If

            With frmCCProcessor.rowARTCCPA1
                .Item("CUST_CODE") = CUST_CODE
                .Item("CCPA_AMT") = ORDR_TOTAL_AMT
                .Item("CCPA_NOTE") = IIf(chargeAgainstAuth, "Pre-Auth Sale", "Credit Card Order")
            End With

            If chargeAgainstAuth Then
                frmCCProcessor.overrideSaleWithCapture = True
                frmCCProcessor.overrideSaleApprovalCode = rowARTCCPA1_AUTH.Item("RESPONSE_APPROVAL_CODE") & String.Empty
                frmCCProcessor.overrideSaleTransactionID = rowARTCCPA1_AUTH.Item("TRANS_ID") & String.Empty
                frmCCProcessor.overrideSaleCreditCardFullName = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NAME") & String.Empty
            End If

            Try
                frmCCProcessor.ShowDialog()
                Dim row As DataRow = ASCDATA1.GetDataRow("select * from ARTCCPA1 where CCPA_NO = :PARM1", "V", New Object() {frmCCProcessor.CCPA_NO & String.Empty})
                If row IsNot Nothing AndAlso (row.Item("CCPA_STATUS") & String.Empty = "T" OrElse row.Item("CCPA_STATUS") & String.Empty = "S") Then
                    rowSOTORDR1.Item("CCPA_NO") = frmCCProcessor.CCPA_NO & String.Empty
                    rowSOTORDR1.Item("CC_TRANS_ID") = row.Item("TRANS_ID")

                    If TRAN_TYPE = "A" AndAlso row.Item("CCPA_STATUS") & String.Empty = "T" Then
                        ASCDATA1.ExecuteSQL("UPDATE SOTORDR1 SET CCPA_NO = '" & rowSOTORDR1.Item("CCPA_NO") & "', CC_TRANS_ID = '" & rowSOTORDR1.Item("CC_TRANS_ID") & "' WHERE ORDR_NO = '" & ORDR_NO & "'")
                    End If

                    If Not chargeAgainstAuth Then
                        rowSOTORDC1 = dst.Tables("SOTORDC1").NewRow
                        rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                        rowSOTORDC1.Item("TRANS_TYPE") = IIf(TRAN_TYPE = "A", "C", "D")
                        rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                        rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowSOTORDC1.Item("CCPA_NO") = row.Item("CCPA_NO")
                        rowSOTORDC1.Item("CCPA_STATUS") = row.Item("CCPA_STATUS")
                        rowSOTORDC1.Item("AMOUNT") = row.Item("CCPA_AMT")
                        rowSOTORDC1.Item("BALANCE") = row.Item("CCPA_AMT")
                        rowSOTORDC1.Item("ACTIVE_IND") = "1"
                        dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)
                    Else
                        rowSOTORDC1 = dst.Tables("SOTORDC1").Rows.Find(New Object() {ORDR_NO, TRANS_NO})
                        Dim rowSOTORDC2 As DataRow = dst.Tables("SOTORDC2").NewRow
                        rowSOTORDC2.Item("ORDR_NO") = row.Item("ORDR_NO")
                        rowSOTORDC2.Item("TRANS_NO") = row.Item("TRANS_NO")
                        rowSOTORDC2.Item("TRANS_LNO") = Val(dst.Tables("SOTORDC2").Compute("MAX(TRANS_LNO)", "ORDR_NO = '" & row.Item("ORDR_NO") & "' AND TRANS_NO = " & TRANS_NO) & String.Empty) + 1
                        rowSOTORDC2.Item("INV_DATE") = DateTime.Now
                        rowSOTORDC2.Item("INV_NO") = "Approved"
                        rowSOTORDC2.Item("AMOUNT_APPLIED") = row.Item("AMOUNT")
                        dst.Tables("SOTORDC2").Rows.Add(rowSOTORDC2)
                    End If

                ElseIf row IsNot Nothing Then
                    If Not chargeAgainstAuth Then
                        rowSOTORDC1 = dst.Tables("SOTORDC1").NewRow
                        rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                        rowSOTORDC1.Item("TRANS_TYPE") = IIf(TRAN_TYPE = "A", "C", "D")
                        rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                        rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowSOTORDC1.Item("CCPA_NO") = row.Item("CCPA_NO")
                        rowSOTORDC1.Item("CCPA_STATUS") = row.Item("CCPA_STATUS")
                        rowSOTORDC1.Item("AMOUNT") = row.Item("CCPA_AMT")
                        rowSOTORDC1.Item("BALANCE") = 0
                        rowSOTORDC1.Item("ACTIVE_IND") = "0"
                        dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)
                    Else
                        Dim rowSOTORDC2 As DataRow = dst.Tables("SOTORDC2").NewRow
                        rowSOTORDC2.Item("ORDR_NO") = row.Item("ORDR_NO")
                        rowSOTORDC2.Item("TRANS_NO") = row.Item("TRANS_NO")
                        rowSOTORDC2.Item("TRANS_LNO") = Val(dst.Tables("SOTORDC2").Compute("MAX(TRANS_LNO)", "ORDR_NO = '" & row.Item("ORDR_NO") & "' AND TRANS_NO = " & TRANS_NO) & String.Empty) + 1
                        rowSOTORDC2.Item("INV_DATE") = DateTime.Now
                        rowSOTORDC2.Item("INV_NO") = "Declined"
                        rowSOTORDC2.Item("AMOUNT_APPLIED") = 0
                        dst.Tables("SOTORDC2").Rows.Add(rowSOTORDC2)
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Using

        ASCMAIN1.MultiTask_Release(, , 1)
        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub


    Private Function GetFreightCosts() As Decimal

        Dim freightCosts As Decimal = 0

        Try
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", MyBase.Absx1.txtFor("WHSE_CODE").Text)

            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text)
            Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE")
            Dim rowSOTCARR1 As DataRow = LookUp("SOTCARR1", CARRIER_CODE)
            Dim rowSOTCARR3 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTCARR3 WHERE CARRIER_CODE = :PARM1", "V", New Object() {CARRIER_CODE})

            Dim ShippingLabelDirectory As String = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)
            Dim isInternationalShipment As Boolean = False
            If rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty = String.Empty Then
                isInternationalShipment = False
            Else
                isInternationalShipment = Not (rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.ToUpper.StartsWith("US")
            End If

            ' See if the consumer uses their own account to pay for  freight
            If rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" Then
                Return 0
            End If

            Select Case rowSOTCARR1.Item("PROVIDER_TYPE")
                Case WHCSHIP1.ProviderTypeFedex
                    If Not isInternationalShipment Then
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
                        clsShip.RequestedServiceType = nsoftware.InShip.ServiceTypes.stFedExGround
                    Else
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpressInternational)
                        clsShip.RequestedServiceType = nsoftware.InShip.ServiceTypes.stFedExInternationalGround
                    End If
                Case WHCSHIP1.ProviderTypeUPS
                    clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)
                    clsShip.RequestedServiceType = nsoftware.InShip.ServiceTypes.stUPSGround
                Case WHCSHIP1.ProviderTypeUSPS
                    clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.USPS)
                    Return 0
                Case WHCSHIP1.ProviderTypeCanada
                    clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.CanadaPost)
                    Return 0
                Case Else
                    Return 0
            End Select

            With clsShip.Sender
                .FirstName = rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty
                .MiddleInitial = ""
                .LastName = ""
                .Address1 = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
                .Address2 = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
                .City = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
                .State = rowICTWHSE1.Item("WHSE_STATE") & String.Empty
                .ZipCode = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
                .CountryCode = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
                .Company = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                .Phone = rowICTWHSE1.Item("WHSE_PHONE") & String.Empty

                .IsResidental = False
                .IsPOBox = False

                If .Company.Length = 0 Then
                    .Company = (.FirstName & " " & .LastName).ToString.Trim
                End If
            End With

            With clsShip.Recipient
                .FirstName = rowSOTORDR5.Item("CUST_CONTACT") & String.Empty
                .MiddleInitial = ""
                .LastName = "" 'txtFromLastName.Text.Trim
                .Address1 = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                .Address2 = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                .City = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                .State = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                .ZipCode = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                .CountryCode = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
                If .CountryCode.Trim = String.Empty OrElse .CountryCode.Trim.ToUpper.StartsWith("US") Then
                    .CountryCode = "US"
                End If
                .Company = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                .Phone = "1234567890"
            End With

            Dim weight As Decimal = 0
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "", DataViewRowState.CurrentRows)
                If Val(rowSOTORDR2.Item("ITEM_WEIGHT") & String.Empty) = 0 Then
                    ' Assume 4 ounces iof there is no weight.
                    weight += 0.25 + Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                Else
                    weight += Val(rowSOTORDR2.Item("ITEM_WEIGHT") & String.Empty) + Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                End If
            Next

            shipPackageDetailList.Clear()
            Dim shipPackageDetail As New nsoftware.InShip.PackageDetail
            With shipPackageDetail
                .PackagingType = nsoftware.InShip.TPackagingTypes.ptYourPackaging
                .Weight = Convert.ToInt32(weight)
                .Length = 17.5
                .Width = 17.5
                .Height = 13.5
                .Id = "00000001"
            End With

            shipPackageDetailList.Add(shipPackageDetail)

            Try
                If ASCMAIN1.Running_in_VS Then
                    ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "C:\")
                End If
                If ShippingLabelDirectory.Length > 0 Then
                    If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                        My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                    End If
                End If
            Catch ex As Exception
                ShippingLabelDirectory = String.Empty
            End Try

            If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
                ShippingLabelDirectory = ShippingLabelDirectory & "\"
            End If

            ' Credentials
            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

            If isInternationalShipment Then
                If Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty) > 0 Then
                    clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty)
                End If
            Else
                If Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty) > 0 Then
                    clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
                End If
            End If

            If isInternationalShipment Then
                clsShip.CommodityDetailList.Clear()
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "", DataViewRowState.CurrentRows)
                    Dim CommodityDetail As New nsoftware.InShip.CommodityDetail
                    CommodityDetail.Description = rowSOTORDR2.Item("ITEM_DESC") & String.Empty
                    CommodityDetail.NumberOfPieces = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                    CommodityDetail.Quantity = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                    CommodityDetail.QuantityUnit = "EA"
                    CommodityDetail.UnitPrice = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                    CommodityDetail.Weight = rowSOTORDR2.Item("ITEM_WEIGHT")
                    CommodityDetail.Manufacturer = "US"
                    clsShip.CommodityDetailList.Add(CommodityDetail)
                Next
                clsShip.TotalCustomsValue = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & String.Empty)
            End If

            With clsShip
                .EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = "X"
                If IsDate(MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value) AndAlso MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value >= DateTime.Now Then
                    .ShipDate = MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value
                Else
                    .ShipDate = DateTime.Now
                End If

            End With

            clsShip.PackageDetailList = shipPackageDetailList
            clsShip.GetRates()

            freightCosts = 0

            For Each Charge As KeyValuePair(Of Integer, Decimal) In clsShip.ShipmentListCharge
                freightCosts += Val(Charge.Value & String.Empty)
            Next

        Catch ex As Exception
            freightCosts = -1
        Finally

        End Try

        Return freightCosts

    End Function

    Sub Setup_Multiple_Order_Grid(show_header_fields As Boolean)
        For Each COLUMN_NAME As String In New String() _
            {"ORDR_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_CUST_PO", "ORDR_DEPT", "ORDR_ADDR_TYPE_ST", "ORDR_HOLD", "ORDR_STATUS"}
            grdSOTORDRB.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not show_header_fields
            ' "CUST_STORE_NAME", ""
        Next
        For Each COLUMN_NAME As String In New String() _
            {"ORDR_GROUP_NO", "CUST_STORE_LOCATION"} ', "ORDR_ARRIVAL_DATE", "ORDR_LAST_ARRIVAL_DATE", "TERM_CODE", "WHSE_CODE", "SHIP_VIA_CODE", "FRT_TERMS", "REASON_CODE", "ORDR_HOLD_REASON", "ORDR_SHIP_INSTR", "ORDR_INV_COMMENT"}
            grdSOTORDRB.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
        Next
    End Sub

    'Sub Log_Changes( _
    '    row As DataRow, _
    '    TABLE_NAME As String, _
    '    ByRef Check_Changed_Fields As Boolean, _
    '    ByRef REV_LNO As Integer, _
    '    LAST_DATE As Date)

    '    For i As Integer = 0 To row.Table.Columns.Count - 1
    '        Dim COLUMN_NAME As String = dst.Tables(TABLE_NAME).Columns(i).ColumnName

    '        If row.Item(COLUMN_NAME) & "" _
    '        <> row.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
    '            Check_Changed_Fields = True
    '            ASCMAIN1.Progress("-", COLUMN_NAME)
    '            Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
    '            With rowSOTORDXR
    '                .Item("REV_NO") = REV_NO
    '                REV_LNO += 1
    '                .Item("REV_LNO") = REV_LNO
    '                .Item("ORDR_NO") = ORDR_NO
    '                .Item("ORDR_LNO") = 0
    '                .Item("INIT_DATE") = LAST_DATE
    '                .Item("INIT_OPER") = ASCMAIN1.USER_ID
    '                .Item("COLUMN_NAME") = COLUMN_NAME
    '                .Item("OLD_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Original)
    '                .Item("NEW_VALUE") = row.Item(COLUMN_NAME)
    '                .Item("EMODE") = EntryMode
    '            End With
    '            dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
    '            Check_Changed_Fields = True
    '        End If
    '    Next i
    'End Sub

    Sub Toggle_Customer_Style_Fields(show As Boolean)
        With grdSOTORDR2.DisplayLayout.Bands(0)
            .Columns("CUST_UPC").Hidden = Not show
            .Columns("CUST_SKU").Hidden = Not show
            .Columns("CUST_STYLE_CODE").Hidden = Not show
            .Columns("CUST_COLOR_CODE").Hidden = Not show
            .Columns("CUST_SIZE_CODE").Hidden = Not show
            .Columns("STYLE_RETAIL").Hidden = Not show
        End With
    End Sub

    Sub Toggle_Disc_Comm_Fields(show As Boolean)
        With grdSOTORDR2.DisplayLayout.Bands(0)
            .Columns("COMM_RATE").Hidden = Not show
            .Columns("DISC_AMT").Hidden = Not show
            .Columns("DISC_PCT").Hidden = Not show
        End With
    End Sub

    Private Sub grdSOTORDRB_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDRB.AfterRowUpdate
        multistore_changes_were_made_to_qty = True
    End Sub

    Private Sub grdSOTORDRB_BackgroundImageChanged(sender As Object, e As System.EventArgs) Handles grdSOTORDRB.BackgroundImageChanged

    End Sub

    Private Sub grdSOTORDRB_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDRB.BeforeCellUpdate
        If e.Cell.Column.Key.StartsWith("QTY_") Then
            Dim row As DataRow = dst.Tables("SOTORDRB").Rows.Find(e.Cell.Row.Cells("ORDR_NO").Value)
            If row.Item(e.Cell.Column.Key) & "" = "" Then
                e.Cancel = True
            Else

                Dim COL As Integer = Val(Mid(e.Cell.Column.Key, 5))
                Dim rowSOTORDRI As DataRow = dst.Tables("SOTORDRI").Select("COL = " & CStr(COL))(0)
                Dim ORDR_NO As String = row.Item("ORDR_NO")
                Dim STYLE_CODE_ORIG As String = rowSOTORDRI.Item("STYLE_CODE_ORIG")
                Dim rowSOTORDRQ As DataRow = dst.Tables("SOTORDRQ").Select("ORDR_NO = '" & ORDR_NO & "' and STYLE_CODE_ORIG = '" & STYLE_CODE_ORIG & "'")(0)
                If rowSOTORDRQ.Item("ORDR_QTY") < Val(e.NewValue & "") Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    Function Make_ARTCUST2_BTST() As DataRow

        If Absx1.optFor("ORDR_TYPE_CODE").Value & "" = "XFR" Then
        Else
            If MsgBox("Address Record 000000 does Not Exist" & vbCrLf & vbCrLf & "OK to Establish a record using Customer Address?",
                      MsgBoxStyle.YesNo,
                      "Option to Set up Ship-To 0000000") = MsgBoxResult.No Then
                Return Nothing
            End If
        End If

        ASCMAIN1.sql = "Insert into ARTCUST2 (" & vbCrLf _
            & "CUST_CODE,CUST_ADDR_TYPE,CUST_ADDR_CODE," & vbCrLf _
            & "CUST_NAME,CUST_ADDR1,CUST_ADDR2,CUST_ADDR3," & vbCrLf _
            & "CUST_CITY,CUST_STATE,CUST_ZIP_CODE,CUST_COUNTRY," & vbCrLf _
            & "CUST_CONTACT,CUST_PHONE,CUST_EXT,CUST_FAX," & vbCrLf _
            & "INIT_OPER,LAST_OPER,INIT_DATE,LAST_DATE," & vbCrLf _
            & "CUST_ADDR_NAME,CUST_ADDR_STATUS,CUST_EMAIL)" & vbCrLf _
            & "Select CUST_CODE,'MK' CUST_ADDR_TYPE,'000000' CUST_ADDR_CODE," & vbCrLf _
            & "CUST_NAME,CUST_ADDR1,CUST_ADDR2,CUST_ADDR3," & vbCrLf _
            & "CUST_CITY,CUST_STATE,CUST_ZIP_CODE,CUST_COUNTRY," & vbCrLf _
            & "CUST_CONTACT,CUST_PHONE,CUST_EXT,CUST_FAX," & vbCrLf _
            & "'" & ASCMAIN1.USER_ID & "' INIT_OPER, '" & ASCMAIN1.USER_ID & "' LAST_OPER, SYSDATE INIT_DATE, SYSDATE LAST_DATE," & vbCrLf _
            & "CUST_NAME CUST_ADDR_NAME,'A' CUST_ADDR_STATUS,CUST_EMAIL" & vbCrLf _
            & " from ARTCUST1 where CUST_CODE = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", CUST_CODE)

        Return LookUp("ARTCUST2", New String() {CUST_CODE, "MK", "000000"})

    End Function

    Sub Set_XFR_Visibility(TF As Boolean)
        lblSREP_CODE.Visible = Not TF
        lblSREP2_CODE.Visible = Not TF
        lblTERM_CODE.Visible = Not TF
        ' lblREASON_CODE.Visible = Not TF
        lblFRT_TERMS.Visible = Not TF

        txtSREP_CODE.Visible = Not TF
        txtSREP2_CODE.Visible = Not TF
        txtTERM_CODE.Visible = Not TF
        ' txtREASON_CODE.Visible = Not TF
        txtFRT_TERMS.Visible = Not TF

        txtSREP_NAME.Visible = Not TF
        txtSREP2_NAME.Visible = Not TF
        txtTERM_DESC.Visible = Not TF
        ' txtREASON_DESC.Visible = Not TF
        txtFRT_TERMS_DESC.Visible = Not TF

        Absx1.chkFor("CUST_FACTOR_IND").Visible = Not TF

    End Sub

    Sub Allocate_Order(ORDR_NO As String)

        ASCMAIN1.sql = "Update SOTORDR2" _
            & " Set ORDR_QTY_ALLO = 0, ORDR_RELEASE = NULL, ORDR_RELEASE_AVAIL = NULL" _
            & " where ORDR_NO = '" & ORDR_NO & "'"
        ASCDATA1.ExecuteSQL()

        Dim c As ASCBASE1 = Me.clsASCBASE1

        If clsASCBASE1_allo Is Nothing Then
            clsASCBASE1_allo = New ASCBASE1
            Me.clsASCBASE1 = clsASCBASE1_allo
            Setup_DataLayer()

            TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me,
                "",
                False,
                True,
                False,
                "", Now.Date.AddDays(30)) ' using 30 days release date horizon

            With dst
                Dim SOTSUPP1 As String = ASCMAIN1.Temp_Table("Select * from SOTSUPP1")
                ASCMAIN1.sql = "Select * from " & SOTSUPP1
                Create_TDA(.Tables.Add, "SOTSUPP1", "**", 0, False)
                TABLE_NAMEs.Add("SOTSUPP1", SOTSUPP1)

                Dim SOTDEMD1 As String = ASCMAIN1.Temp_Table("Select * from SOTDEMD1")
                ASCMAIN1.sql = "Select * from " & SOTDEMD1
                Create_TDA(.Tables.Add, "SOTDEMD1", "**", 0, False)
                TABLE_NAMEs.Add("SOTDEMD1", SOTDEMD1)
            End With


        Else
            Me.clsASCBASE1 = clsASCBASE1_allo
            Setup_DataLayer()
        End If

        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_CODE,COLOR_CODE")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Allocate(STYLE_CODE, COLOR_CODE)
        Next



        ' NEED TO DO A FEW OTHERS LIKE ICTSTDQ1 IN SOROREL1.BUILD WORKFILE 
        ' - MAYBE SOMEDAY WE EXTERNALIZE THOSE UPDATES AND CALL A COMMON ROUTINE?

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is Select * from " & TABLE_NAMEs("SOTORDR2") & " where ORDR_NO = '" & ORDR_NO & "';" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2 Set" _
            & "      ORDR_QTY_ALLO = R1.ORDR_QTY_ALLO" _
            & "    , ORDR_RELEASE = R1.ORDR_RELEASE" _
            & "    , ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" _
            & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()


        Me.clsASCBASE1 = c
        Setup_DataLayer()

    End Sub

    Private Sub Allocate(STYLE_CODE As String, COLOR_CODE As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Allocating ... (Please Wait)")

        'grdSOTALLO1.Visible = False

        Dim SOTORDR0 As String = TABLE_NAMEs("SOTORDR0")
        Dim SOTORDR1 As String = TABLE_NAMEs("SOTORDR1")
        Dim SOTORDR2 As String = TABLE_NAMEs("SOTORDR2")
        Dim SOTRSRV1 As String = TABLE_NAMEs("SOTRSRV1")
        Dim SOTRSRV2 As String = TABLE_NAMEs("SOTRSRV2")
        Dim ARTCUST1 As String = TABLE_NAMEs("ARTCUST1")

        Dim SOTSUPP1 As String = TABLE_NAMEs("SOTSUPP1")
        Dim SOTDEMD1 As String = TABLE_NAMEs("SOTDEMD1")

        For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR0", "ARTCUST1", "ICTSTDQ1", "SOTORDR2", "SOTRSRV1", "SOTRSRV2"}
            ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAMEs(TABLE_NAME))
        Next

        For Each sql As String In TABLE_NAMEs.Keys
            If sql.StartsWith("sql") Then
                Dim sqlstmt As String = Replace(TABLE_NAMEs(sql), "'STYLE_CODE'", "'" & STYLE_CODE & "'")
                ASCDATA1.ExecuteSQL(sqlstmt)
            End If
        Next

        dst.Tables("SOTSUPP0").Rows.Clear()
        dst.Tables("SOTSUPPI").Rows.Clear()
        dst.Tables("SOTORDR7").Rows.Clear()
        dst.Tables("ICTSTDQ1").Rows.Clear()

        Dim edi850cust As List(Of String) = TAC.SOCMAIN1.Get_EDI_Custs("850")

        TAC.SOCMAIN1.Allocation(Me, False, True, "", "", edi850cust,
                                  SOTSUPP1, SOTDEMD1,
                                  TABLE_NAMEs,
                                  True, True, STYLE_CODE, COLOR_CODE)
        '(optASL.Value = "1")

        ' Truncate SOTORDR1 SOTORDR0 ARTCUST1 ICTSTDQ1 SOTORDR2 SOTRSRV1 SOTRSRV2
        ' Execute all sql's loaded into TABLE_NAMEs dictionary, in the order that they were placed
        ' Clear Rows for SOTSUPP0 SOTSUPPI SOTORDR7 and refill as necessary
    End Sub

    Sub Get_PO_Cost()
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            If Val(rowSOTORDR2.Item("PO_COST") & "") = 0 Then
                rowSOTORDR2.Item("PO_COST") = TAC.SOCMAIN1.Get_PO_Cost(Me, STYLE_CODE, rowICTSTYL1.Item("VEND_CODE"), rowSOTORDR1)
            End If
        Next
    End Sub

    Private Sub chkOnHoldOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkOnHoldOnly.CheckedChanged
        Load_SOTORDRX()
    End Sub

    Private Sub grdSOTORDP1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDP1.AfterRowActivate
        Setup_grdSOTORDP2()

        If EntryMode = "N" Or EntryMode = "E" Then
            If grdSOTORDP1.ActiveRow.Cells("INV_STATUS").Value & "" = "1" Then
                grdSOTORDP1.DisplayLayout.Bands(0).Columns("INV_REF").CellActivation = UltraWinGrid.Activation.NoEdit
                grdSOTORDP1.DisplayLayout.Bands(0).Columns("INV_DATE").CellActivation = UltraWinGrid.Activation.NoEdit
                grdSOTORDP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTORDP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdSOTORDP1.DisplayLayout.Bands(0).Columns("INV_REF").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSOTORDP1.DisplayLayout.Bands(0).Columns("INV_DATE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSOTORDP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTORDP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            End If
        End If
    End Sub

    Sub Setup_grdSOTORDP2()
        If grdSOTORDP1.ActiveRow Is Nothing Then
            grdSOTORDP2.Visible = False
        Else
            grdSOTORDP2.Visible = True
            Dim INV_NO As String = grdSOTORDP1.ActiveRow.Cells("INV_NO").Value & ""
            Dim dvw As DataView = DirectCast(grdSOTORDP2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "INV_NO = '" & INV_NO & "'"
            grdSOTORDP2.Text = "Pro-Forma Invoice Details for Invoice " & INV_NO
        End If
    End Sub

    Sub Generate_Pro_Forma_Invoice()
        Dim INV_NO As String = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        Dim rowSOTORDP1 As DataRow = dst.Tables("SOTORDP1").NewRow
        rowSOTORDP1.Item("ORDR_NO") = ORDR_NO
        rowSOTORDP1.Item("INV_NO") = INV_NO
        rowSOTORDP1.Item("INV_COMMENT") = Absx1.txtFor("ORDR_INV_COMMENT").Text
        If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
            rowSOTORDP1.Item("INV_COMMENT") = Absx1.txtFor("PF_NOTE").Text
            '  rowSOTORDR1.Item("ORDR_FOB") = Absx1.txtFor("PF_FOB").Text
        End If

        dst.Tables("SOTORDP1").Rows.Add(rowSOTORDP1)

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim rowSOTORDP2 As DataRow = dst.Tables("SOTORDP2").NewRow
            rowSOTORDP2.Item("ORDR_NO") = ORDR_NO
            rowSOTORDP2.Item("INV_NO") = INV_NO
            Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
            rowSOTORDP2.Item("ORDR_LNO") = ORDR_LNO
            Dim ORDR_QTY As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
            Dim ORDR_QTY_USED As Int64 = Val(dst.Tables("SOTORDP2").Compute("SUM(ORDR_QTY_SHIP)", "ORDR_LNO = " & CStr(ORDR_LNO)) & "")
            Dim ORDR_QTY_LEFT As Int64 = ORDR_QTY - ORDR_QTY_USED
            If ORDR_QTY_LEFT <= 0 Then
                ORDR_QTY_LEFT = 0
            Else
                rowSOTORDP2.Item("ORDR_QTY_SHIP") = ORDR_QTY_LEFT
                dst.Tables("SOTORDP2").Rows.Add(rowSOTORDP2)
            End If
        Next

        For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDP1.Rows
            If grow.Cells("INV_NO").Value = INV_NO Then
                grdSOTORDP1.ActiveRow = grow
                Setup_grdSOTORDP2()
                Exit For
            End If
        Next
    End Sub

    Private Sub grdSOTORDP1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTORDP1.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("INV_STATUS").Value & "" = "1" Then
                MsgBox("Cannot Delete or Change a Pro-Forma Invoice which has been Issued during a Receipt",
                       MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
                Exit Sub
            End If
        Next
    End Sub

    Private Sub grdSOTORDP1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDP1.BeforeRowUpdate
        If e.Row.Cells("INV_STATUS").Value & "" <> "" Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdSOTORDRB_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDRB.InitializeLayout

    End Sub

    Sub Set_SOTORDRI_to_SOTORDRB()
        If multistore_changes_were_made_to_qty Then
            For Each row As DataRow In dst.Tables("SOTORDRI").Select("")
                Dim COL As Integer = Val(row.Item("COL") & "")
                row.Item("ORDR_QTY_OPEN") = Val(dst.Tables("SOTORDRB").Compute("SUM(QTY_" & Format(COL, "000") & ")", "") & "")
            Next
            multistore_changes_were_made_to_qty = False
        End If
    End Sub

    Private Sub tabMultiOrder_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMultiOrder.SelectedTabChanged
        If tabMultiOrder.SelectedTab.Key = "Styles" Then
            If multistore_changes_were_made_to_qty Then
                Set_SOTORDRI_to_SOTORDRB()
            End If
        End If
    End Sub

    Private Sub grdSOTORDP1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDP1.InitializeRow
        If e.Row.Cells("INV_STATUS").Value & "" = "1" Then
            e.Row.Cells("INV_NO").Appearance.ForeColor = Drawing.Color.Green
            e.Row.Cells("INV_NO").ToolTipText = "Invoice has been Updated - Goods Received"
        End If
    End Sub

    Sub Setup_DISC_AMT(WHSE_CODE As String)
        With dst.Tables("SOTORDR2")
            If WHSE_CODE = "FE" Or WHSE_CODE = "FD" Then
                .Columns("DISC_AMT").Expression = "ISNULL(ORDR_UNIT_PRICE_STD,0)-ISNULL(ORDR_UNIT_PRICE,0)"
                .Columns("DISC_PCT").Expression = "IIF(ISNULL(ORDR_UNIT_PRICE_STD,0)=0,0,100*DISC_AMT/ISNULL(ORDR_UNIT_PRICE_STD,0))"
            Else
                .Columns("DISC_AMT").Expression = "ISNULL(STYLE_PRICE,0)-ISNULL(ORDR_UNIT_PRICE,0)"
                .Columns("DISC_PCT").Expression = "IIF(ISNULL(STYLE_PRICE,0)=0,0,100*DISC_AMT/ISNULL(STYLE_PRICE,0))"
            End If
        End With
    End Sub

    Function Get_ORDR_FOB(ORDR_TYPE_CODE As String, WHSE_CODE As String) As String
        Dim ORDR_FOB As String = ""

        If ORDR_TYPE_CODE = "BTB" Then
            ORDR_FOB = "Port of Origin"
        Else
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            If rowICTWHSE1 IsNot Nothing AndAlso rowICTWHSE1.Item("WHSE_CITY") & "" <> "" Then
                ORDR_FOB = rowICTWHSE1.Item("WHSE_CITY") & "," & rowICTWHSE1.Item("WHSE_STATE")
            Else
                ORDR_FOB = ""
            End If
        End If
        Return ORDR_FOB
    End Function

    Private Sub ProcessCreditCardDeposit(ByVal processType As String, ByVal TRANS_NO As String)

        Dim errorMsg As String = String.Empty

        If Not InquiryMode Then
            MessageBox.Show("The feature is available only in Inquiry Mode.", processType, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim row As DataRow = ASCDATA1.GetDataRow("Select * from SOTORDR1 where ORDR_NO = :PARM1", "V", New Object() {ORDR_NO})
        Dim CCPA_NO As String = String.Empty

        If TRANS_NO.Length > 0 AndAlso dst.Tables("SOTORDC1").Select("TRANS_NO = " & TRANS_NO).Length > 0 Then
            CCPA_NO = dst.Tables("SOTORDC1").Select("TRANS_NO = " & TRANS_NO)(0).Item("CCPA_NO") & String.Empty
        End If

        If row Is Nothing OrElse Not ",O,P,".Contains(row.Item("ORDR_STATUS")) Then
            If CCPA_NO <> row.Item("CCPA_NO") & String.Empty AndAlso processType = "Void Authorization" Then
                ' Permit Voiding an used Authorization. Sometimes the girls double authorize.
            Else
                MessageBox.Show("Sales Order Status changed. Action Aborted!", processType, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        End If

        ' Do not permit any additions if the order is locked by someone else
        Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & String.Empty
        If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , , , 4) Then
            Exit Sub
        End If

        Select Case processType
            Case "CC Authorization", "Void Authorization"

                CCPA_NO = (rowSOTORDR1.Item("CCPA_NO") & String.Empty).ToString.Trim

                If processType = "Void Authorization" Then
                    CCPA_NO = dst.Tables("SOTORDC1").Select("TRANS_NO = " & TRANS_NO)(0).Item("CCPA_NO")
                End If

                Dim OriginalAuthAmount As Decimal = 0
                Dim previouslySettledAmount As Decimal = 0
                Dim tblARTCCPA1 As DataTable = New DataTable
                Dim rowARTCCPA1_AUTH As DataRow = Nothing

                errorMsg = String.Empty
                ' See if there is an existing Authorization on the Sales Order
                If CCPA_NO <> String.Empty Then
                    tblARTCCPA1 = ASCDATA1.GetDataTable("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = :PARM1 OR CCPA_NO_AUTH = :PARM1", "ARTCCPA1", "V", New Object() {CCPA_NO})
                    If clsTACENCRY.UseEncryption Then
                        For Each rowARTCCPA1 As DataRow In tblARTCCPA1.Select("")
                            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE", "CUST_CREDIT_CARD_EXP_DATE"}
                                ' this line was used to temporarily fix the incorrectly encrypted CC values in an ARTCCPA1 record that we think came in over the old API - but not sure how it encrypted anything
                                'rowARTCCPA1.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCCPA1.Item(field) & String.Empty)

                                Select Case field
                                    Case "CUST_CREDIT_CARD_EXP_DATE"
                                        If rowARTCCPA1.Item(field & "_E") & String.Empty = String.Empty Then
                                            Continue For
                                        End If

                                        If rowARTCCPA1.Item(field) & String.Empty <> String.Empty Then
                                            Continue For
                                        End If

                                End Select

                                rowARTCCPA1.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1.Item(field & "_E") & String.Empty)
                                rowARTCCPA1.Item(field & "_E") = DBNull.Value
                            Next
                        Next
                    End If

                    OriginalAuthAmount = Val(tblARTCCPA1.Compute("SUM(CCPA_AMT)", "CCPA_NO = '" & CCPA_NO & "' AND CCPA_STATUS = 'T'") & String.Empty)
                    previouslySettledAmount = Val(tblARTCCPA1.Compute("SUM(CCPA_AMT)", "CCPA_NO_AUTH = '" & CCPA_NO & "' AND CCPA_STATUS = 'S'") & String.Empty)
                    rowARTCCPA1_AUTH = tblARTCCPA1.Rows.Find(CCPA_NO)

                    If rowARTCCPA1_AUTH IsNot Nothing AndAlso rowARTCCPA1_AUTH.Item("CCPA_DATE_VOID") & String.Empty = String.Empty Then
                        errorMsg = "The sales order has an active Authorization for the amount of " & Format(OriginalAuthAmount, "$#,#0.00")
                        errorMsg &= Environment.NewLine
                        errorMsg &= "There are sales against the Authorization totaling " & Format(previouslySettledAmount, "$#,#0.00")
                        errorMsg &= Environment.NewLine
                        errorMsg &= "Leaving a balance of " & Format(OriginalAuthAmount - previouslySettledAmount, "$#,#0.00")
                        errorMsg &= Environment.NewLine

                        If processType = "Void Authorization" Then
                            errorMsg &= "The current Authorization will be released. Do you want to Void the balance?"
                        Else
                            errorMsg &= "The current Authorization will be released. Do you want to continue to try and create a new Authorization?"
                        End If
                    End If
                End If

                If errorMsg = String.Empty Then
                    If processType = "Void Authorization" Then
                        MessageBox.Show("The selected Authorization cannot be voided.", "Void Authorization", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    Else
                        errorMsg = "Do you want to continue to try and create a new Authorization?"
                    End If
                End If

                'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" And Format(Now, "MM/dd/yyyy") = "11/20/2022" Then
                '    ' don't ask questions - we are voiding many orders to prepare for new component
                'Else
                If MessageBox.Show(errorMsg, processType, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then '
                    Exit Select
                End If
                'End If

                'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" And Format(Now, "MM/dd/yyyy") = "11/20/2022" Then
                '    previouslySettledAmount = 0
                'End If

                If previouslySettledAmount < OriginalAuthAmount AndAlso rowARTCCPA1_AUTH.Item("CCPA_DATE_VOID") & String.Empty = String.Empty Then
                    Try
                        Dim CUST_CREDIT_CARD_EXP_DATE As String = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
                        CUST_CREDIT_CARD_EXP_DATE = CUST_CREDIT_CARD_EXP_DATE.PadRight(4, "0")

                        Dim CreditCardProcessor As TAC.TAFCARDF
                        CreditCardProcessor = New TAC.TAFCARDF(Me)

                        CreditCardProcessor.MerchantSetup()
                        With CreditCardProcessor.objCCProcessor
                            .CustomerCreditCard.CardNumber = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NO") & String.Empty
                            .CustomerCreditCard.CardExpMonth = CUST_CREDIT_CARD_EXP_DATE.Substring(0, 2)
                            .CustomerCreditCard.CardExpYear = CUST_CREDIT_CARD_EXP_DATE.Substring(2)
                        End With

                        Dim Creditcard As New TAC.ARCCCARD.CreditCard
                        Creditcard = CreditCardProcessor.CreateCreditCardInfo(rowARTCCPA1_AUTH)
                        With Creditcard
                            .InvoiceNumber = ""
                            .TransArmorToken = ""
                            .RefundAmount = Val(rowARTCCPA1_AUTH.Item("CCPA_AMT") & String.Empty)
                        End With

                        If CreditCardProcessor.objCCProcessor.VoidTransaction(Creditcard) Then
                            dst.Tables("SOTORDC1").Rows.Find(New Object() {ORDR_NO, TRANS_NO}).Item("ACTIVE_IND") = "0"
                            Update_Record_TDA("SOTORDC1")
                            Dim sql As String = "Update ARTCCPA1 SET CCPA_DATE_VOID = SYSDATE, LAST_DATE = SYSDATE"
                            sql &= ", CCPA_REASON_VOID = 'Voided to create a New Authorization', LAST_OPER = '" & ASCMAIN1.USER_ID & "'"
                            sql &= " WHERE CCPA_NO = '" & rowARTCCPA1_AUTH.Item("CCPA_NO") & "'"
                            ASCDATA1.ExecuteSQL(sql)
                            sql = "Update SOTORDR1 SET CC_TRANS_ID = NULL, CCPA_NO = NULL WHERE ORDR_NO = '" & ORDR_NO & "'"
                            ASCDATA1.ExecuteSQL(sql)

                            Dim rowSOTORDC2 As DataRow = dst.Tables("SOTORDC2").NewRow
                            rowSOTORDC2.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDC2.Item("TRANS_NO") = TRANS_NO
                            rowSOTORDC2.Item("TRANS_LNO") = Val(dst.Tables("SOTORDC2").Compute("MAX(TRANS_LNO)", "ORDR_NO = '" & row.Item("ORDR_NO") & "' AND TRANS_NO = " & TRANS_NO) & String.Empty) + 1
                            rowSOTORDC2.Item("INV_DATE") = DateTime.Now
                            rowSOTORDC2.Item("INV_NO") = "Voided"
                            rowSOTORDC2.Item("AMOUNT_APPLIED") = OriginalAuthAmount - previouslySettledAmount
                            dst.Tables("SOTORDC2").Rows.Add(rowSOTORDC2)
                            Update_Record_TDA("SOTORDC2")

                            'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" And Format(Now, "MM/dd/yyyy") = "11/20/2022" Then
                            'Else
                            MessageBox.Show(processType & " successful.", "CC Processor", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            'End If

                        Else
                            MessageBox.Show($"Void Transaction Failed: {CreditCardProcessor.objCCProcessor.LastError}", "CC Processor", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            'Exit Sub
                        End If

                    Catch ex As Exception
                        MessageBox.Show("Error trying to Void previous CC Authorization: " & ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If

                If processType = "CC Authorization" Then
                    ' Try to get an Authorization
                    Credit_Card("A", CCPA_NO, , True)
                    Update_Record_TDA("SOTORDC1")
                    Update_Record_TDA("SOTORDC2")
                End If

            Case "CC Deposit"
                Credit_Card("S", , , True)
                Update_Record_TDA("SOTORDC1")
                Update_Record_TDA("SOTORDC2")

            Case "Add On Account"
                Dim onAccountAmount As Decimal = 0

                Dim inputValue As String = InputBox("Enter the amount of On Account funds you want to apply to this sale. These funds will be deducted from any Credit Card Sales Transactions.", "On Account", "0.00") & String.Empty
                inputValue = inputValue.Trim
                If inputValue.Length = 0 Then inputValue = 0
                onAccountAmount = CDec(inputValue)

                Select Case onAccountAmount
                    Case Is < 0
                        MessageBox.Show("On Account amount must be greater than $0.00", "On Account", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Select
                    Case 0
                        ' Assume the user clicked cancel
                        Exit Select
                    Case Is > 0
                        If MessageBox.Show("Please confirm you want to apply an On Account amount for " & Format(onAccountAmount, "$#,##0.00") & " for this sales order." _
                                            , "On Account", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Select
                        End If

                        Dim rowSOTORDC1 As DataRow = dst.Tables("SOTORDC1").NewRow
                        rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                        rowSOTORDC1.Item("TRANS_TYPE") = "O"
                        rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                        rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        'rowSOTORDC1.Item("CCPA_NO") = ""
                        'rowSOTORDC1.Item("CCPA_STATUS") = ""
                        rowSOTORDC1.Item("AMOUNT") = onAccountAmount
                        rowSOTORDC1.Item("BALANCE") = onAccountAmount
                        rowSOTORDC1.Item("ACTIVE_IND") = "1"
                        dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)
                        Update_Record_TDA("SOTORDC1")
                End Select

            Case "Additional Funds"
                Dim additionalFunds As Decimal = 0

                Dim inputValue As String = InputBox("Enter the amount of Additional funds you want to apply to each invoice for this sale.", "Additional Funds", "0.00") & String.Empty
                inputValue = inputValue.Trim
                If inputValue.Length = 0 Then inputValue = 0
                additionalFunds = CDec(inputValue)

                Select Case additionalFunds
                    Case Is < 0
                        MessageBox.Show("Additional Funds amount must be greater than $0.00", "Additional Funds", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Select
                    Case 0
                        ' Assume the user clicked cancel
                        Exit Select
                    Case Is > 0
                        If MessageBox.Show("Please confirm you want to apply an Additional Funds amount for " & Format(additionalFunds, "$#,##0.00") & " for this sales order." _
                                            , "Additional Funds", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Select
                        End If

                        Dim rowSOTORDC1 As DataRow = dst.Tables("SOTORDC1").NewRow
                        rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                        rowSOTORDC1.Item("TRANS_TYPE") = "A"
                        rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                        rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        'rowSOTORDC1.Item("CCPA_NO") = ""
                        'rowSOTORDC1.Item("CCPA_STATUS") = ""
                        rowSOTORDC1.Item("AMOUNT") = additionalFunds
                        rowSOTORDC1.Item("BALANCE") = 0
                        rowSOTORDC1.Item("ACTIVE_IND") = "1"
                        dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)
                        Update_Record_TDA("SOTORDC1")
                End Select

            Case "De-Activate"
                Try
                    Dim rowSOTORDC1 As DataRow = dst.Tables("SOTORDC1").Rows.Find(New Object() {ORDR_NO, TRANS_NO})
                    If rowSOTORDC1 IsNot Nothing Then
                        If MessageBox.Show("Do you want to De-Activate the selected entry?", "De-Activate", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Select
                        End If
                        rowSOTORDC1.Item("ACTIVE_IND") = "0"
                        Update_Record_TDA("SOTORDC1")
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message, "De-Activate", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "Charge Against Auth"
                If 1 = 1 Then
                    Exit Sub
                End If
                Dim rowSOTORDC1 As DataRow = dst.Tables("SOTORDC1").Rows.Find(New Object() {ORDR_NO, TRANS_NO})
                Credit_Card("C", rowSOTORDC1.Item("CCPA_NO") & String.Empty, TRANS_NO)
                Update_Record_TDA("SOTORDC1")
                Update_Record_TDA("SOTORDC2")

        End Select

        ' Clean up any locks
        ASCMAIN1.MultiTask_Release(, , 4)

    End Sub

    Private Sub grdSOTORDC1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDC1.InitializeRow
        If e.Row.Band.Key <> "SOTORDC1" Then Exit Sub
        If e.Row.Cells("CCPA_NO").Value & String.Empty = String.Empty Then Exit Sub
        If dst.Tables("SOTORDR1").Rows.Count = 0 Then Exit Sub

        If e.Row.Cells("CCPA_NO").Value & String.Empty = dst.Tables("SOTORDR1").Rows(0).Item("CCPA_NO") & String.Empty Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        Else
            e.Row.Appearance.BackColor = Drawing.Color.White
        End If

    End Sub

    Private Sub optCustomerOrders_ValueChanged(sender As Object, e As System.EventArgs) Handles optCustomerOrders.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDRX()
    End Sub

    Sub Create_POs(ORDR_NOs As List(Of String))
        grdPOTORDR1.Parent = splPOs.Panel1
        grdPOTORDR1.Visible = True
        Create_PO(False, ORDR_NOs)
        ' AT THE END OF THIS PROCESS, NEED TO CLEAR_RECORD TO EMPTY POTORDR1/2 ETC
        grdSOTORDRX.Visible = False
        splPOs.Visible = True
        SplitContainer14.Visible = False

        Set_Control_POs(True)
    End Sub

    Sub Set_Control_POs(tf As Boolean)
        With UltraExplorerBar1
            .Groups("Create POs from SOs").Visible = tf
            .Groups("Show Orders").Visible = Not tf
            .Groups("Screen Control").Visible = Not tf
        End With

        UltraGroupBox1.Visible = Not tf

    End Sub

    Private Sub grdSOTORDRX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDRX.InitializeLayout
        For Each colname As String In New String() {"CUST_CITY", "CUST_STATE", "CUST_COUNTRY"}
            grdSOTORDRX.DisplayLayout.Bands(0).Columns(colname).Hidden = Not (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
        Next
    End Sub

    Private Function PARTIALSTYLE(STYLE_CODE As String) As String
        Dim RETVAL As String = ""
        Dim NEW_STYLE As String = ""
        ASCMAIN1.sql = String.Format("SELECT COUNT(*) RECCNT FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
        Dim STYLE_COUNT As Int32 = Val(ASCDATA1.GetDataValue)
        If STYLE_COUNT = 1 Then
            ASCMAIN1.sql = String.Format("SELECT STYLE_CODE FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
            NEW_STYLE = ASCDATA1.GetDataValue
            RETVAL = NEW_STYLE
        End If
        Return RETVAL
    End Function


    Public Overrides Function CustomSummary_End(
        ByVal summarySettings As UltraWinGrid.SummarySettings,
        ByVal rows As UltraWinGrid.RowsCollection,
        ByVal CustomValue As Double,
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdSOTORDR2"
                Dim KEY As String = summarySettings.Key
                If KEY = "GP_PCT" Then
                    TOTALS.Add("GP_AMT", 0)
                    TOTALS.Add("ORDR_AMT_SHIP", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("ORDR_AMT_SHIP") <> 0 Then CustomValue = 100 * TOTALS("GP_AMT") / TOTALS("ORDR_AMT_SHIP")

                ElseIf KEY = "MU_PCT" Then
                    TOTALS.Add("PO_COST_EXT", 0)
                    TOTALS.Add("ORDR_AMT", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("PO_COST_EXT") <> 0 Then CustomValue = 100 * (TOTALS("ORDR_AMT") - TOTALS("PO_COST_EXT")) / TOTALS("PO_COST_EXT")

                ElseIf KEY = "" Then
                    Stop
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Public Overrides Function CustomStringSummary_End(
        ByVal summarySettings As UltraWinGrid.SummarySettings,
        ByVal rows As UltraWinGrid.RowsCollection,
        ByVal CustomValue As String,
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdSOTORDR2"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Totals"
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals(
       ByVal rows As UltraWinGrid.RowsCollection,
       ByRef TOTALS As Dictionary(Of String, Decimal),
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY = "GP_PCT" Then
                    TOTALS("GP_AMT") += Val(grow2.Cells("GP_AMT").Value & "")
                    TOTALS("ORDR_AMT_SHIP") += Val(grow2.Cells("ORDR_AMT_SHIP").Value & "")

                ElseIf KEY = "MU_PCT" Then
                    TOTALS("PO_COST_EXT") += Val(grow2.Cells("PO_COST").Value & "") * Val(grow2.Cells("ORDR_QTY").Value & "")
                    TOTALS("ORDR_AMT") += Val(grow2.Cells("ORDR_AMT").Value & "")
                ElseIf KEY = "" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub

    Private Sub tabSOTORDRX_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSOTORDRX.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        If tabSOTORDRX.SelectedTab.Key = "Pro-Forma Invoices" Then
            If tabSOTORDRX.SelectedTab.Tag = "*" Then
                Load_SOTORDPX()
            End If
        End If

        UltraExplorerBar1.Groups("Show Orders").Visible = (tabSOTORDRX.SelectedTab.Key = "Open Orders")
    End Sub

    Private Sub grdSOTORDP1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTORDP1.DoubleClickRow
        If e.Row.IsDataRow And Not ScreenMode Then
            Dim ORDR_NO As String = e.Row.Cells("ORDR_NO").Value & ""
            Absx1.txtFor("ORDR_NO").Text = ORDR_NO
            Click_Command("View")
        End If
    End Sub

    Private Sub txtORDR_SHIP_INSTR_ValueChanged(sender As Object, e As EventArgs) Handles txtORDR_SHIP_INSTR.ValueChanged
        If (EntryMode = "N" Or EntryMode = "E") And ScreenMode Then
            If txtORDR_SHIP_INSTR.Text = "" Then
                chkORDR_INCL_VAS.ForeColor = Drawing.Color.Empty
            Else
                chkORDR_INCL_VAS.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub cmdAutoPO_Click(sender As Object, e As EventArgs) Handles cmdAutoPO.Click

        ' note this method has been commandeered to handle voiding all auths
        If Not ASCMAIN1.Running_in_VS Or ASCMAIN1.USER_ID <> "wjz" Then Exit Sub

        Stop

        ASCMAIN1.sql = "SELECT SOTORDC1.CCPA_NO, SOTORDR1.ORDR_NO, SOTORDC1.TRANS_NO
FROM SOTORDR1,ARTCCPA1,SOTORDC1
 WHERE SOTORDR1.ORDR_STATUS IN ('O','P') 
   AND SOTORDC1.ORDR_NO = SOTORDR1.ORDR_NO AND SOTORDC1.ACTIVE_IND = '1'
   AND ARTCCPA1.CCPA_NO = SOTORDC1.CCPA_NO AND ARTCCPA1.CCPA_DATE_VOID IS NULL AND ARTCCPA1.CCPA_AMT > 1"
        Dim tblv As DataTable = ASCDATA1.GetDataTable

        For Each rowv As DataRow In tblv.Rows
            ORDR_NO = rowv.Item("ORDR_NO")
            Dim CCPA_NO As String = rowv.Item("CCPA_NO")
            Dim TRANS_NO As String = rowv.Item("TRANS_NO")

            ASCMAIN1.Progress(ORDR_NO)

            rowSOTORDR1 = Fill_Record("SOTORDR1", ORDR_NO)

            dst.Tables("SOTORDC2").Rows.Clear()
            dst.Tables("SOTORDC1").Rows.Clear()

            ASCMAIN1.sql = "Select SOTORDC1.*, ARTCCPA1.CUST_CREDIT_CARD_LAST4, ARTCCPA1.CCPA_DATE_VOID" _
                 & " from SOTORDC1, ARTCCPA1 " _
                 & " where SOTORDC1.ccpa_no = ARTCCPA1.ccpa_no (+)" _
                 & " and SOTORDC1.ORDR_NO = '" & ORDR_NO & "'"
            Fill_Records("SOTORDC1", String.Empty, True, ASCMAIN1.sql)
            Fill_Records("SOTORDC2", ORDR_NO)

            ProcessCreditCardDeposit("Void Authorization", TRANS_NO)


            'Stop
        Next

        Exit Sub

        ASCMAIN1.sql = "SELECT SOTORDR1.ORDR_NO from SOTORDR1 where CUST_CODE = 'LOBLAW' and ORDR_STATUS = 'O'"
        ASCMAIN1.sql &= " and ORDR_NO <> '0000940918'"
        Dim tbl As DataTable = ASCDATA1.GetDataTable

        Dim automatic As Boolean = True

        For Each row As DataRow In tbl.Select("", "ORDR_NO")
            Absx1.txtFor("ORDR_NO").Text = row.Item("ORDR_NO")
            Click_Command("View")

            If ScreenMode Then
                blnAutomatic = True
                Click_Command("Re-Queue for Credit")
                blnAutomatic = False
            Else
                Stop
            End If

            Click_Command("Done")

            If ScreenMode Then
                Stop
            End If
        Next
    End Sub

    Private Sub PF_WEIGHT_ValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub PF_WEIGHT_UOM_ValueChanged(sender As Object, e As EventArgs)

    End Sub
    Sub SET_PROFORMA_CONTROLS()

        If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then

            With UltraExplorerBar1
                .Groups("Pro Forma Invoice").Visible = chkExportInfo.Checked
            End With
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("PF_QTY").Hidden = Not chkExportInfo.Checked
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("PF_DUTY_HTS_CODE").Hidden = Not chkExportInfo.Checked
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("PF_ORDER_NO").Hidden = Not chkExportInfo.Checked

            grdSOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            Absx1.dteFor("PF_INV_DATE").Value = Absx1.dteFor("ORDR_SHIP_DATE").Value

            With grdSOTORDR2.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If gcol.Key = "PF_QTY" Or gcol.Key = "PF_DUTY_HTS_CODE" Or gcol.Key = "PF_ORDER_NO" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                Next
            End With


        End If

    End Sub
    Private Sub chkExportInfo_CheckedChanged(sender As Object, e As EventArgs) Handles chkExportInfo.CheckedChanged
        SET_PROFORMA_CONTROLS()
    End Sub

    Private Sub UltraTextEditor94_ValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub UltraOptionSet3_ValueChanged(sender As Object, e As EventArgs) Handles optPF_OVERSEAS_DOMESTIC.ValueChanged
        If optPF_OVERSEAS_DOMESTIC.Value = "O" Then
            UltraTextEditor93.Visible = True
            UltraLabel88.Visible = True
        Else
            UltraTextEditor93.Visible = False
            UltraLabel88.Visible = False
            Absx1.txtFor("PO_SHIPMENT_NO").Text = ""
        End If
    End Sub

    Private Sub chkShortView_CheckedChanged(sender As Object, e As EventArgs) Handles chkShortView.CheckedChanged

        If COLUMN_NAMEs_All.Count = 0 Then
            For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDR2.DisplayLayout.Bands(0).Columns
                If Not gcol.Hidden Then
                    COLUMN_NAMEs_All.Add(gcol.Key)
                End If
            Next

            COLUMN_NAMEs_Short.Add("ORDR_LNO")
            COLUMN_NAMEs_Short.Add("STYLE_CODE")
            COLUMN_NAMEs_Short.Add("STYLE_DESC")
            COLUMN_NAMEs_Short.Add("COLOR_CODE")
            COLUMN_NAMEs_Short.Add("ORDR_QTY_OPEN")
            COLUMN_NAMEs_Short.Add("ORDR_RELEASE_AVAIL")
            COLUMN_NAMEs_Short.Add("ORDR_UNIT_PRICE")
            COLUMN_NAMEs_Short.Add("ORDR_AMT_OPEN")
            COLUMN_NAMEs_Short.Add("DATE_1")
            COLUMN_NAMEs_Short.Add("QTY_1")
            COLUMN_NAMEs_Short.Add("AMT_1")
            COLUMN_NAMEs_Short.Add("DATE_2")
            COLUMN_NAMEs_Short.Add("QTY_2")
            COLUMN_NAMEs_Short.Add("AMT_2")
            COLUMN_NAMEs_Short.Add("DATE_3")
            COLUMN_NAMEs_Short.Add("QTY_3")
            COLUMN_NAMEs_Short.Add("AMT_3")
            COLUMN_NAMEs_Short.Add("DATE_4")
            COLUMN_NAMEs_Short.Add("QTY_4")
            COLUMN_NAMEs_Short.Add("AMT_4")
        End If

        For Each COLUMN_NAME As String In COLUMN_NAMEs_All
            grdSOTORDR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = chkShortView.Checked And Not COLUMN_NAMEs_Short.Contains(COLUMN_NAME)
        Next
    End Sub
    Sub Import_XFR_File()
        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            'Dim filter As String = "xlsb files (*.xlsb)|*.xlsx|All files (*.*)|*.*"
            Dim filter As String = "All files (*.*)|*.*"
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            Try
                ASCMAIN1.Progress("Now Building Order From Excel", "")
                Me.Cursor = Cursors.WaitCursor

                DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
                Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
                Dim xws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                xws = XWB.Worksheets(1)
                Dim ERROR_CODEs As List(Of String) = New List(Of String)
                Dim ORDR_LNO As Integer = 1
                Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("ORDR_NO")
                Dim CUST_CODE As String = "TRANSFERS"
                Dim CUST_STORE_NO As String = "LUKY21"
                Dim CUST_STORE_NAME As String = "AMAZON.COM SERVICES, INC"
                Dim STYLE_COLORs As New Dictionary(Of String, Decimal)
                Dim STYLE_COLOR As String = ""
                Dim ORDR_CUST_PO As String = ""
                Dim ORDR_SHIP_DATE As Date = Nothing
                Dim ORDR_CANCEL_DATE As Date = Nothing
                Dim PRE_PACK_OVERRIDE As String = ""

                Me.Cursor = Cursors.WaitCursor

                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)

                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
                If rowARTCUST2 IsNot Nothing Then
                    CUST_STORE_NAME = rowARTCUST2.Item("CUST_NAME") & ""
                End If

                If rowARTCUST1 IsNot Nothing Then
                    CUST_CODE = CUST_CODE

                    CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
                    If CUST_BILL_TO_CUST = "" Then
                        CUST_BILL_TO_CUST = CUST_CODE
                    End If

                    Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
                    If rowARTCUST1_BT Is Nothing Then
                        EMsg &= vbCr & "Unable to determine Bill-To Customer"
                    Else
                        If rowARTCUST1_BT.Item("POST_CODE") & "" = "" Then
                            ' EMsg &= vbCr & "No value specified for the Post Code for Bill-To Customer " & CUST_BILL_TO_CUST
                        Else
                            If LookUp("ARTPOST1", rowARTCUST1_BT.Item("POST_CODE")) Is Nothing Then
                                '   EMsg &= vbCr & "Invalide AR Post Code specified for Bill-To Customer " & CUST_BILL_TO_CUST
                            End If
                        End If
                    End If

                End If

                For i As Integer = 2 To xws.UsedRange.Rows.Count Step +1
                    If xws.Cells(i, 1).value <> "" Then
                        Dim SKU As String = ""
                        If xws.Cells(i, 4).value = "" Then
                            SKU = ""
                        Else
                            SKU = xws.Cells(i, 4).value.ToString
                        End If

                        Dim STYLECOLOR As String() = Split(SKU, "-")
                        Dim STYLE_CODE As String = ""
                        Dim STYLE_DESC As String = ""
                        Dim COLOR_CODE As String = ""
                        Dim ORDR_SELLER_FEE As Double = 0
                        Dim ORDR_FULLFILL_FEE As Double = 0
                        Dim ORDR_RETAIL_PRICE As Double = 0
                        Dim ORDR_UNIT_PRICE As Double = 0
                        Dim ORDR_QTY_OPEN As Int32 = 0



                        If i = 2 Then
                            ORDR_SHIP_DATE = xws.Cells(i, 2).value.ToString
                            ORDR_CANCEL_DATE = xws.Cells(i, 3).value.ToString
                            ORDR_CUST_PO = xws.Cells(i, 1).value.ToString
                        End If



                        Dim rowSOTCSTY1 As DataRow = LookUp("SOTCSTY1", New String() {"AMAZONFBA", SKU})
                        If rowSOTCSTY1 Is Nothing Then


                            ERROR_CODEs.Add("SKU is missing from Style Cross Reference File " & SKU & " On Line No " & i)

                            Dim rowERROR_TBL As DataRow = Nothing
                            rowERROR_TBL = dst.Tables("ERROR_TBL").NewRow
                            With rowERROR_TBL
                                .Item("SKU") = SKU
                                .Item("ERROR_DETAIL") = "Ln# " & i
                            End With
                            dst.Tables("ERROR_TBL").Rows.Add(rowERROR_TBL)

                            STYLE_CODE = ""
                            COLOR_CODE = ""
                        Else
                            STYLE_CODE = rowSOTCSTY1.Item("STYLE_CODE")
                            COLOR_CODE = rowSOTCSTY1.Item("COLOR_CODE")
                        End If


                        If STYLE_CODE <> "" And Val(xws.Cells(i, 8).value) <> 0 Then
                            Dim rowICTSTYL1 As DataRow = clsASCBASE1.LookUp("ICTSTYL1", STYLE_CODE)
                            If rowICTSTYL1 Is Nothing Then
                                ERROR_CODEs.Add("Invalid Style Code " & STYLE_CODE & " for " & SKU & " On Line No " & i)
                                STYLE_CODE = ""
                            Else
                                STYLE_DESC = rowICTSTYL1.Item("STYLE_DESC")
                                ORDR_QTY_OPEN = xws.Cells(i, 8).value
                                PRE_PACK_OVERRIDE = xws.Cells(i, 8).value

                                ORDR_RETAIL_PRICE = 0
                                ORDR_SELLER_FEE = 0
                                ORDR_FULLFILL_FEE = 0
                                ORDR_SELLER_FEE = Math.Abs(ORDR_SELLER_FEE)
                                ORDR_FULLFILL_FEE = Math.Abs(ORDR_FULLFILL_FEE)
                                ORDR_UNIT_PRICE = 0
                                ' ORDR_UNIT_PRICE = ORDR_RETAIL_PRICE - ((ORDR_SELLER_FEE + ORDR_FULLFILL_FEE) / ORDR_QTY_OPEN)
                            End If
                        End If


                        'add SOTORDR2
                        If STYLE_CODE <> "" Then
                            ''Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
                            ''Dim DIVISOR As Integer = 0
                            ''If rowICTSTYLS IsNot Nothing And rowICTSTYLS.Item("STYLE_SIZE") & "" = "" Then
                            ''    For II As Integer = 1 To 12
                            ''        DIVISOR = DIVISOR + Val(rowICTSTYLS.Item("QTY_" & Format(II, "00")) & "")
                            ''    Next
                            ''End If
                            ''If DIVISOR <> 0 Then
                            ''    ORDR_QTY_OPEN = ORDR_QTY_OPEN / DIVISOR
                            ''Else
                            ''    Stop
                            ''End If

                            Dim ORDR_LNO_NEW As String = ORDR_LNO
                            STYLE_COLOR = STYLE_CODE & COLOR_CODE
                            If Not STYLE_COLORs.ContainsKey(STYLE_COLOR) Then
                                STYLE_COLORs.Add(STYLE_COLOR, ORDR_LNO)
                            Else
                                ORDR_LNO_NEW = STYLE_COLORs(STYLE_COLOR)
                            End If

                            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New String() {ORDR_NO, ORDR_LNO_NEW})
                            If rowSOTORDR2 IsNot Nothing Then
                                rowSOTORDR2.Item("ORDR_QTY") = rowSOTORDR2.Item("ORDR_QTY") + ORDR_QTY_OPEN
                                rowSOTORDR2.Item("ORDR_QTY_OPEN") = rowSOTORDR2.Item("ORDR_QTY_OPEN") + ORDR_QTY_OPEN
                                rowSOTORDR2.Item("ORDR_QTY_ORIG") = rowSOTORDR2.Item("ORDR_QTY_ORIG") + ORDR_QTY_OPEN
                            Else
                                rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                                With rowSOTORDR2
                                    .Item("ORDR_NO") = ORDR_NO
                                    .Item("ORDR_LNO") = ORDR_LNO
                                    .Item("STYLE_CODE") = STYLE_CODE
                                    .Item("COLOR_CODE") = COLOR_CODE

                                    .Item("STYLE_DESC") = STYLE_DESC
                                    .Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE
                                    .Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE
                                    .Item("ORDR_QTY") = ORDR_QTY_OPEN
                                    .Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                                    .Item("ORDR_QTY_ORIG") = ORDR_QTY_OPEN
                                    .Item("ORDR_QTY_ALLO") = 0
                                    .Item("INNER_PACK_QTY") = 0
                                    .Item("ORDR_EXTD_COST") = 0
                                    .Item("STYLE_UOM") = "EA"
                                    .Item("ORDR_QTY_PICK") = 0
                                    .Item("ORDR_QTY_SHIP") = 0
                                    .Item("ORDR_QTY_CANC") = 0
                                    .Item("ORDR_STATUS") = "O"
                                    .Item("ORDR_QTY_PRE_ALLO") = 0
                                    .Item("QTY_PER_PP") = 0
                                    .Item("CARTON_PACK_QTY") = 0
                                    .Item("STYLE_PRICE") = 0
                                    .Item("ORDR_UNIT_PRICE_CALC") = 0
                                    .Item("ORDR_UNIT_PRICE_MANUAL") = ""
                                    .Item("STYLE_RETAIL") = 0
                                    .Item("PO_COST") = 0
                                    .Item("COMM_RATE") = 0
                                    .Item("ORDR_RETAIL_PRICE") = ORDR_UNIT_PRICE
                                    .Item("ORDR_SELLER_FEE") = ORDR_SELLER_FEE
                                    .Item("ORDR_FULLFILL_FEE") = ORDR_FULLFILL_FEE

                                End With
                                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
                                ORDR_LNO = ORDR_LNO + 1
                            End If
                        End If
                    End If
        Next

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Rows
                    Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE") & ""
                    Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
                    Dim DIVISOR As Integer = 0
                    If rowICTSTYLS IsNot Nothing Then 'And rowICTSTYLS.Item("STYLE_SIZE") & "" = "" Then
                        Dim rowICTSTYL1 As DataRow = clsASCBASE1.LookUp("ICTSTYL1", STYLE_CODE)
                        If rowICTSTYL1 IsNot Nothing And rowICTSTYL1.Item("SIZE_CODE") & "" = "" Then
                            For II As Integer = 1 To 12
                                DIVISOR = DIVISOR + Val(rowICTSTYLS.Item("QTY_" & Format(II, "00")) & "")
                            Next
                        End If
                    End If
                    If DIVISOR <> 0 Then
                        Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                        WRITE_SOTORDR3(ORDR_NO, rowSOTORDR2.Item("ORDR_LNO"), STYLE_CODE, ORDR_QTY_OPEN, DIVISOR)
                        'If PRE_PACK_OVERRIDE = "" Then

                        'Else
                        '    ' WRITE 1 RECORD OUT TO SOTORDR3 
                        '    Dim rowSOTORDR3 As DataRow = dst.Tables("SOTORDR3").NewRow
                        '    With rowSOTORDR3
                        '        .Item("ORDR_NO") = ORDR_NO
                        '        .Item("ORDR_LNO") = rowSOTORDR2.Item("ORDR_LNO") & ""
                        '        .Item("ORDR_SUB_LNO") = 1
                        '        .Item("CUST_STYLE_CODE") = STYLE_CODE
                        '        .Item("CUST_COLOR_CODE") = PRE_PACK_OVERRIDE
                        '        .Item("ORDR_QTY") = ORDR_QTY_OPEN
                        '    End With
                        '    dst.Tables("SOTORDR3").Rows.Add(rowSOTORDR3)
                        'End If

                        ''  ORDR_QTY_OPEN = ORDR_QTY_OPEN / DIVISOR
                        rowSOTORDR2.Item("ORDR_QTY") = ORDR_QTY_OPEN
                        rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                        rowSOTORDR2.Item("ORDR_QTY_ORIG") = ORDR_QTY_OPEN
                    End If

                Next


                Dim CUST_ADDR_TYPE As String = ""
                Dim CUST_ADDR_CODE As String = ""


                For RR As Integer = 1 To 2
                    If RR = 1 Then
                        CUST_ADDR_TYPE = "BT"
                        CUST_ADDR_CODE = "000000"
                    ElseIf RR = 2 Then
                        CUST_ADDR_TYPE = "ST"
                        CUST_ADDR_CODE = "LUKY21"
                    End If
                    ADD_SOTORDR5(ORDR_NO, CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE)
                Next

                Dim ORDR_GROUP_NO As String = ORDR_NO

                Dim rowSOTORDR1 As DataRow = Nothing
                rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
                With rowSOTORDR1
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_DATE") = DATETIME_STAMP.Date
                    .Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE")
                    .Item("CUST_NAME") = rowARTCUST1.Item("CUST_CODE")
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("CUST_STORE_NAME") = CUST_STORE_NAME
                    .Item("ORDR_FOB") = "Monroe Township NJ"
                    .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                    .Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE
                    .Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
                    .Item("POST_CODE") = "TRADE"
                    .Item("TERM_CODE") = "49"
                    .Item("SREP_CODE") = "036"
                    .Item("SREP2_CODE") = "045"
                    .Item("WHSE_CODE") = "NJC"
                    .Item("WHSE_CODE_TO") = "AMAZ02"
                    .Item("SALES_DIVISION_CODE") = "15"
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date
                    .Item("ORDR_SOURCE") = "A"
                    .Item("FRT_TERMS") = "COL"
                    .Item("ORDR_ADDR_TYPE_ST") = "MK"
                    .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
                    .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
                    .Item("ORDR_STATUS") = "O"
                    .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                    .Item("ORDR_HOLD") = "0"
                    .Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_CODE")
                    .Item("CUST_FACTOR_IND") = "0"
                    .Item("CURR_CODE") = "USD"
                    .Item("CURR_EXCH_RATE") = 1
                    .Item("ORDR_ORIG_SHIP_DATE") = ORDR_SHIP_DATE
                    .Item("ORDR_ORIG_CANCEL_DATE") = ORDR_CANCEL_DATE
                    .Item("ORDR_TYPE_CODE") = "XFR"
                    ' B2C inStead of REG

                End With
                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                Dim ORDR_AMT As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
                Dim ORDR_QTY_ORIG As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "") & "")

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("", "")





                If ERROR_CODEs.Count <> 0 Then
                    'Using fr As New ASFMSGBF
                    '    fr.Show_grd(ERROR_TBL, Me, "The following Import Errors have been identified")
                    'End Using

                    If dst.Tables("ERROR_TBL").Rows.Count <> 0 Then
                        Using F As New ASFMSGBF
                            F.Show_grd(dst.Tables("ERROR_TBL"), Me, "The following Import Errors have been identified", "DGJ")
                        End Using
                    End If


                    '             MsgBox("These are the following Errors in the SpreadSheet: " & Join(ERROR_CODEs.ToArray, vbCrLf), MsgBoxStyle.OkOnly, "Cannot Update Spreadsheet")
                    dst.Tables("SOTORDR1").Rows.Clear()
                    dst.Tables("SOTORDR3").Rows.Clear()
                    dst.Tables("SOTORDR2").Rows.Clear()
                    dst.Tables("SOTORDR5").Rows.Clear()
                    dst.Tables("SOTORDR0").Rows.Clear()
                    dst.Tables("ERROR_TBL").Rows.Clear()

                Else
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("ORDR_GROUP_NO")

                    For Each row As DataRow In dst.Tables("SOTORDR1").Select("")
                        row.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                    Next
                    For Each row As DataRow In dst.Tables("SOTORDR0").Select("")
                        row.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                        row.Item("ORDR_NO_MIN") = ORDR_NO
                        row.Item("ORDR_NO_MAX") = ORDR_NO
                    Next

                    BeginTrans()
                    Update_Record_TDA("SOTORDR1")
                    Update_Record_TDA("SOTORDR2")
                    Update_Record_TDA("SOTORDR3")
                    Update_Record_TDA("SOTORDR5")
                    Update_Record_TDA("SOTORDR0")

                    Dependent_Updates(1, ORDR_NO)

                    ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

                    CommitTrans()
                    MsgBox("This Excel File has been successfully Updated to the Sales Order",
                            MsgBoxStyle.OkOnly, "Verification")

                    dst.Tables("SOTORDR1").Rows.Clear()
                    dst.Tables("SOTORDR3").Rows.Clear()
                    dst.Tables("SOTORDR2").Rows.Clear()
                    dst.Tables("SOTORDR5").Rows.Clear()
                    dst.Tables("SOTORDR0").Rows.Clear()
                    dst.Tables("ERROR_TBL").Rows.Clear()
                End If
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "XFR Import, Excel Format Issues", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

    End Sub

    Sub Import_Amazon_File()
        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            'Dim filter As String = "xlsb files (*.xlsb)|*.xlsx|All files (*.*)|*.*"
            Dim filter As String = "All files (*.*)|*.*"
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using
        'Try
        If FILENAME <> "" Then
            Try
                ASCMAIN1.Progress("Now Building Order From Excel", "")
                Me.Cursor = Cursors.WaitCursor

                DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
                Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
                Dim xws As Microsoft.Office.Interop.Excel.Worksheet = XWB.Worksheets("SKU Summary")
                Dim ERROR_CODEs As List(Of String) = New List(Of String)
                Dim ORDR_LNO As Integer = 1
                Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("ORDR_NO")
                Dim CUST_CODE As String = "AMAZONFBA"
                Dim CUST_STORE_NO As String = "AMAFBA"
                Dim CUST_STORE_NAME As String = "AMAZON.COM SERVICES, INC"

                Me.Cursor = Cursors.WaitCursor

                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)

                'Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
                'If rowARTCUST2 Is Nothing Then
                '    CUST_STORE_NAME = rowARTCUST2.Item("CUST_NAME") & ""
                'End If

                If rowARTCUST1 IsNot Nothing Then
                    CUST_CODE = CUST_CODE

                    CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
                    If CUST_BILL_TO_CUST = "" Then
                        CUST_BILL_TO_CUST = CUST_CODE
                    End If

                    Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
                    If rowARTCUST1_BT Is Nothing Then
                        EMsg &= vbCr & "Unable to determine Bill-To Customer"
                    Else
                        If rowARTCUST1_BT.Item("POST_CODE") & "" = "" Then
                            ' EMsg &= vbCr & "No value specified for the Post Code for Bill-To Customer " & CUST_BILL_TO_CUST
                        Else
                            If LookUp("ARTPOST1", rowARTCUST1_BT.Item("POST_CODE")) Is Nothing Then
                                '   EMsg &= vbCr & "Invalide AR Post Code specified for Bill-To Customer " & CUST_BILL_TO_CUST
                            End If
                        End If
                    End If
                    ORDR_CUST_PO = "AFBA" & Format(DATETIME_STAMP, "yyMMdd") & ORDR_NO
                    'Absx1.txtFor("ORDR_CUST_PO").Text = Absx1.txtFor("ORDR_CUST_PO").Text.Trim.Replace("'", "")
                    'ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                    'If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                    '    EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                    'End If

                End If

                Dim ORDR_DATE As String = ""
                Dim ORDR_CANCEL_DATE As String = ""


                For i As Integer = 1 To 1
                    If xws.Cells(i, 1).value IsNot Nothing Then
                        ORDR_CUST_PO = xws.Cells(i, 1).value.ToString & ""
                    End If
                    If xws.Cells(i, 2).value IsNot Nothing Then
                        ORDR_DATE = xws.Cells(i, 2).value.ToString & ""
                    End If
                    If xws.Cells(i, 3).value IsNot Nothing Then
                        ORDR_CANCEL_DATE = xws.Cells(i, 3).value.ToString & ""
                    End If
                Next
                If ORDR_CUST_PO = "" Then
                    ORDR_CUST_PO = "AFBA" & Format(DATETIME_STAMP, "yyMMdd") & ORDR_NO
                End If
                If ORDR_DATE = "" Then
                    ORDR_DATE = DATETIME_STAMP.Date
                End If
                If ORDR_CANCEL_DATE = "" Then
                    ORDR_CANCEL_DATE = DateAdd(DateInterval.Month, 1, DATETIME_STAMP.Date)
                End If



                For i As Integer = 4 To xws.UsedRange.Rows.Count Step +1
                    Dim SKU As String = xws.Cells(i, 3).value.ToString
                    Dim STYLECOLOR As String() = Split(SKU, "-")
                    Dim STYLE_CODE As String = ""
                    Dim STYLE_DESC As String = ""
                    Dim COLOR_CODE As String = ""
                    Dim ORDR_SELLER_FEE As Double = 0
                    Dim ORDR_FULLFILL_FEE As Double = 0
                    Dim ORDR_RETAIL_PRICE As Double = 0
                    Dim ORDR_UNIT_PRICE As Double = 0
                    Dim ORDR_QTY_OPEN As Int32 = 0

                    Dim rowSOTCSTY1 As DataRow = LookUp("SOTCSTY1", New String() {CUST_CODE, SKU})
                    If rowSOTCSTY1 Is Nothing Then
                        If Val(xws.Cells(i, 4).value.ToString) <> 0 Then
                            ERROR_CODEs.Add("SKU is missing from Style Cross Reference File " & SKU & " On Line No " & i)

                            Dim rowERROR_TBL As DataRow = Nothing
                            rowERROR_TBL = dst.Tables("ERROR_TBL").NewRow
                            With rowERROR_TBL
                                .Item("SKU") = SKU
                                .Item("ERROR_DETAIL") = "Ln# " & i
                            End With
                            dst.Tables("ERROR_TBL").Rows.Add(rowERROR_TBL)
                        End If

                        STYLE_CODE = ""
                        COLOR_CODE = ""
                        ORDR_QTY_OPEN = 0
                    Else
                        STYLE_CODE = rowSOTCSTY1.Item("STYLE_CODE")
                        COLOR_CODE = rowSOTCSTY1.Item("COLOR_CODE")
                        ORDR_QTY_OPEN = Val(xws.Cells(i, 4).value.ToString)
                    End If


                    If STYLE_CODE <> "" And ORDR_QTY_OPEN <> 0 Then
                        Dim rowICTSTYL1 As DataRow = clsASCBASE1.LookUp("ICTSTYL1", STYLE_CODE)
                        If rowICTSTYL1 Is Nothing Then
                            ERROR_CODEs.Add("Invalid Style Code " & STYLE_CODE & " for " & SKU & " On Line No " & i)
                            STYLE_CODE = ""
                        Else
                            STYLE_DESC = rowICTSTYL1.Item("STYLE_DESC")

                            ORDR_RETAIL_PRICE = Val(xws.Cells(i, 6).value.ToString)
                            ORDR_SELLER_FEE = xws.Cells(i, 7).value.ToString
                            ORDR_FULLFILL_FEE = xws.Cells(i, 8).value.ToString

                            ORDR_SELLER_FEE = Math.Abs(ORDR_SELLER_FEE)
                            ORDR_FULLFILL_FEE = Math.Abs(ORDR_FULLFILL_FEE)
                            ORDR_UNIT_PRICE = ORDR_RETAIL_PRICE - ((ORDR_SELLER_FEE + ORDR_FULLFILL_FEE) / ORDR_QTY_OPEN)
                        End If
                    End If


                    'add SOTORDR2
                    If STYLE_CODE <> "" And ORDR_QTY_OPEN <> 0 Then
                        Dim rowSOTORDR2 As DataRow = Nothing
                        rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                        With rowSOTORDR2
                            .Item("ORDR_NO") = ORDR_NO
                            .Item("ORDR_LNO") = ORDR_LNO
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE

                            .Item("STYLE_DESC") = STYLE_DESC
                            .Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE
                            .Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE
                            .Item("ORDR_QTY") = ORDR_QTY_OPEN
                            .Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                            .Item("ORDR_QTY_ORIG") = ORDR_QTY_OPEN
                            .Item("ORDR_QTY_ALLO") = 0
                            .Item("INNER_PACK_QTY") = 0
                            .Item("ORDR_EXTD_COST") = 0
                            .Item("STYLE_UOM") = "EA"
                            .Item("ORDR_QTY_PICK") = 0
                            .Item("ORDR_QTY_SHIP") = 0
                            .Item("ORDR_QTY_CANC") = 0
                            .Item("ORDR_STATUS") = "O"
                            .Item("ORDR_QTY_PRE_ALLO") = 0
                            .Item("QTY_PER_PP") = 0
                            .Item("CARTON_PACK_QTY") = 0
                            .Item("STYLE_PRICE") = 0
                            .Item("ORDR_UNIT_PRICE_CALC") = 0
                            .Item("ORDR_UNIT_PRICE_MANUAL") = ""
                            .Item("STYLE_RETAIL") = 0
                            .Item("PO_COST") = 0
                            .Item("COMM_RATE") = 0
                            .Item("ORDR_RETAIL_PRICE") = ORDR_UNIT_PRICE
                            .Item("ORDR_SELLER_FEE") = ORDR_SELLER_FEE
                            .Item("ORDR_FULLFILL_FEE") = ORDR_FULLFILL_FEE

                        End With
                        dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
                        ORDR_LNO = ORDR_LNO + 1

                    End If
                Next

                Dim CUST_ADDR_TYPE As String = ""
                Dim CUST_ADDR_CODE As String = ""

                For RR As Integer = 1 To 2
                    If RR = 1 Then
                        CUST_ADDR_TYPE = "BT"
                        CUST_ADDR_CODE = "000000"
                    ElseIf RR = 2 Then
                        CUST_ADDR_TYPE = "ST"
                        CUST_ADDR_CODE = "AMAFBA"
                        '  ElseIf RR = 3 Then
                        '    CUST_ADDR_TYPE = "BY"
                        '    CUST_ADDR_CODE = "AMAZONFBA"
                        'ElseIf RR = 4 Then
                        '  CUST_ADDR_TYPE = "MK"
                        '  CUST_ADDR_CODE = "AMAFBA"
                    Else
                        ' CUST_ADDR_TYPE = "DC"
                        ' CUST_ADDR_CODE = "000000"
                    End If
                    ADD_SOTORDR5(ORDR_NO, CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE)
                Next

                Dim ORDR_GROUP_NO As String = ORDR_NO


                Dim rowSOTORDR1 As DataRow = Nothing
                rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
                With rowSOTORDR1
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_DATE") = ORDR_DATE
                    .Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE")
                    .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("CUST_STORE_NAME") = CUST_STORE_NAME
                    .Item("ORDR_FOB") = "EDISON NJ"
                    ' SHIP FROM WHSE
                    .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                    .Item("ORDR_SHIP_DATE") = ORDR_DATE
                    .Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
                    .Item("POST_CODE") = "TRADE"
                    .Item("TERM_CODE") = "01"
                    .Item("SREP_CODE") = "045"
                    .Item("WHSE_CODE") = "AMAZ02"
                    .Item("SALES_DIVISION_CODE") = "15"
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date
                    .Item("ORDR_SOURCE") = "K"
                    .Item("FRT_TERMS") = "COL"
                    .Item("ORDR_ADDR_TYPE_ST") = "MK"
                    .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
                    .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
                    .Item("ORDR_STATUS") = "O"
                    .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                    .Item("ORDR_HOLD") = "0"
                    .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
                    .Item("CUST_FACTOR_IND") = "0"
                    .Item("CURR_CODE") = "USD"
                    .Item("CURR_EXCH_RATE") = 1
                    .Item("ORDR_ORIG_SHIP_DATE") = ORDR_DATE
                    .Item("ORDR_ORIG_CANCEL_DATE") = ORDR_CANCEL_DATE
                    .Item("ORDR_TYPE_CODE") = "REG"
                    ' B2C inStead of REG

                End With
                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                Dim ORDR_AMT As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
                Dim ORDR_QTY_ORIG As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "") & "")

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("", "")


                If ERROR_CODEs.Count <> 0 Then
                    'Using fr As New ASFMSGBF
                    '    fr.Show_grd(ERROR_TBL, Me, "The following Import Errors have been identified")
                    'End Using

                    If dst.Tables("ERROR_TBL").Rows.Count <> 0 Then
                        Using F As New ASFMSGBF
                            F.Show_grd(dst.Tables("ERROR_TBL"), Me, "The following Import Errors have been identified", "DGJ")
                        End Using
                    End If

                    '             MsgBox("These are the following Errors in the SpreadSheet: " & Join(ERROR_CODEs.ToArray, vbCrLf), MsgBoxStyle.OkOnly, "Cannot Update Spreadsheet")
                    dst.Tables("SOTORDR1").Rows.Clear()
                    dst.Tables("SOTORDR2").Rows.Clear()
                    dst.Tables("SOTORDR5").Rows.Clear()
                    dst.Tables("SOTORDR0").Rows.Clear()
                    dst.Tables("ERROR_TBL").Rows.Clear()
                Else
                    ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("ORDR_GROUP_NO")

                    For Each row As DataRow In dst.Tables("SOTORDR1").Select("")
                        row.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                    Next

                    For Each row As DataRow In dst.Tables("SOTORDR0").Select("")
                        row.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                        row.Item("ORDR_NO_MIN") = ORDR_NO
                        row.Item("ORDR_NO_MAX") = ORDR_NO
                    Next

                    BeginTrans()
                    Update_Record_TDA("SOTORDR1")
                    Update_Record_TDA("SOTORDR2")
                    Update_Record_TDA("SOTORDR5")
                    Update_Record_TDA("SOTORDR0")

                    Dependent_Updates(1, ORDR_NO)

                    ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

                    CommitTrans()
                    MsgBox("This Excel File has been successfully Updated to the Sales Order",
                          MsgBoxStyle.OkOnly, "Verification")

                    dst.Tables("SOTORDR1").Rows.Clear()
                    dst.Tables("SOTORDR2").Rows.Clear()
                    dst.Tables("SOTORDR5").Rows.Clear()
                    dst.Tables("SOTORDR0").Rows.Clear()
                    dst.Tables("ERROR_TBL").Rows.Clear()

                End If
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Amazon Import, Excel Format Issues", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Sub ADD_SOTORDR5(ORDR_NO As String, CUST_CODE As String, CUST_ADDR_TYPE As String, CUST_ADDR_CODE As String)
        Dim rowSOTORDR5 As DataRow = Nothing
        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_ADDR_CODE})
        Dim CUST_NAME As String = "AMAZON.COM SERVICES, INC"
        Dim CUST_ADDR1 As String = "401 INDEPENDENCE ROAD"
        Dim CUST_CITY As String = "FLORENCE"
        Dim CUST_STATE As String = "NJ"
        Dim CUST_ZIP_CODE As String = "08518-220"
        Dim CUST_COUNTRY As String = "USA"

        If rowARTCUST2 IsNot Nothing Then
            CUST_NAME = rowARTCUST2.Item("CUST_NAME") & ""
            CUST_ADDR1 = rowARTCUST2.Item("CUST_ADDR1") & ""
            CUST_CITY = rowARTCUST2.Item("CUST_CITY") & ""
            CUST_STATE = rowARTCUST2.Item("CUST_STATE") & ""
            CUST_ZIP_CODE = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
            CUST_COUNTRY = rowARTCUST2.Item("CUST_COUNTRY") & ""

        End If

        rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
        With rowSOTORDR5
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
            .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
            .Item("CUST_NAME") = CUST_NAME
            .Item("CUST_ADDR1") = CUST_ADDR1
            .Item("CUST_CITY") = CUST_CITY

            .Item("CUST_STATE") = CUST_STATE
            .Item("CUST_ZIP_CODE") = CUST_ZIP_CODE
            .Item("CUST_COUNTRY") = CUST_COUNTRY
        End With
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
    End Sub
    Sub WRITE_SOTORDR3(ORDR_NO As String, ORDR_LNO As Integer, STYLE_CODE As String, QTY_ORDR As Integer, DIVISOR As Integer)
        Dim rowSOTORDR3 As DataRow = Nothing
        Dim SOTORDR3_ORDR_QTY As Integer = 0
        Dim CUST_COLOR_CODE As String = ""
        Dim ORDR_SUB_LNO As Integer = 1

        Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
        If rowICTSTYLS IsNot Nothing Then
            For II As Integer = 1 To 12
                If Val(rowICTSTYLS.Item("QTY_" & Format(II, "00")) & "") <> 0 Then
                    SOTORDR3_ORDR_QTY = (QTY_ORDR / DIVISOR) * Val(rowICTSTYLS.Item("QTY_" & Format(II, "00")) & "")
                    CUST_COLOR_CODE = rowICTSTYLS.Item("SIZE_" & Format(II, "00")) & ""
                    'WRITE SOTORDR3
                    rowSOTORDR3 = dst.Tables("SOTORDR3").NewRow
                    With rowSOTORDR3
                        .Item("ORDR_NO") = ORDR_NO
                        .Item("ORDR_LNO") = ORDR_LNO
                        .Item("ORDR_SUB_LNO") = ORDR_SUB_LNO
                        .Item("CUST_STYLE_CODE") = STYLE_CODE
                        .Item("CUST_COLOR_CODE") = CUST_COLOR_CODE
                        .Item("ORDR_QTY") = SOTORDR3_ORDR_QTY
                    End With
                    dst.Tables("SOTORDR3").Rows.Add(rowSOTORDR3)
                    ORDR_SUB_LNO = ORDR_SUB_LNO + 1
                End If
            Next
        End If

    End Sub
#Region "Promo System"
    Private Sub btnShowPromo_Click(sender As Object, e As EventArgs) Handles btnShowPromo.Click
        Dim F As New ASFMSGBF
        F.grdGroupBy = True
        F.grdFilter = True
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("P1.PROMO_DESC As Promotion,")
        sql.AppendLine("P1.PROMO_START_DATE As Beginning,")
        sql.AppendLine("P1.PROMO_END_DATE As Ending,")
        sql.AppendLine("P2.STYLE_CODE As Style,")
        sql.AppendLine("S1.STYLE_DESC As Description,")
        sql.AppendLine("MAX(P2.PROMO_UNIT_PRICE) As Price")
        sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2, ICTSTYL1 S1")
        sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
        sql.AppendLine("AND P2.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND (P1.PROMO_START_DATE <= SYSDATE AND P1.PROMO_END_DATE >= SYSDATE)")
        sql.AppendLine("GROUP BY")
        sql.AppendLine("P1.PROMO_DESC,")
        sql.AppendLine("P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE,")
        sql.AppendLine("P2.STYLE_CODE,")
        sql.AppendLine("S1.STYLE_DESC")
        Dim tblICTPROMX As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        If tblICTPROMX.Rows.Count > 0 Then
            F.Show_grd(tblICTPROMX, Me, "Current Active Promotions", "")
            F.Dispose()
            F = Nothing
        End If
    End Sub

    Private Sub ShowPromo(ByVal STYLE_CODE As String)
        If EntryMode = "E" Then
            Dim OnPromo As Boolean = False
            Dim PROMO_START_DATE As DateTime
            Dim PROMO_END_DATE As DateTime
            Dim sql As New System.Text.StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT")
            sql.AppendLine("P1.PROMO_START_DATE,")
            sql.AppendLine("P1.PROMO_END_DATE,")
            sql.AppendLine("MAX(P2.PROMO_UNIT_PRICE) PROMO_UNIT_PRICE")
            sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2")
            sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
            sql.AppendLine("AND P2.STYLE_CODE = :PARM1")
            sql.AppendLine("GROUP BY P1.PROMO_START_DATE,")
            sql.AppendLine("P1.PROMO_END_DATE")
            Dim tblICTPROMX As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
            For Each rowICTPROMX As DataRow In tblICTPROMX.Select("", "PROMO_START_DATE")
                PROMO_START_DATE = CDate(rowICTPROMX.Item("PROMO_START_DATE").ToString & String.Empty)
                PROMO_END_DATE = CDate(rowICTPROMX.Item("PROMO_END_DATE").ToString & String.Empty)
                If PROMO_START_DATE <= Now() And PROMO_END_DATE >= Now() Then
                    OnPromo = True
                End If
            Next
            If OnPromo Then
                lblPromo.Text = String.Format("Style On Promo {0} - {1}", PROMO_START_DATE.ToShortDateString, PROMO_END_DATE.ToShortDateString)
                lblPromo.Visible = True
                btnShowPromo.Visible = True
            Else
                lblPromo.Text = ""
                lblPromo.Visible = False
                btnShowPromo.Visible = False
            End If
        End If
    End Sub

    Public Function ImportDetailsToGrid() As Text.StringBuilder
        Dim RetVal As New Text.StringBuilder With {.Length = 0}
        Dim rowSOTORDR1 As DataRow = Nothing
        Dim ORDR_LNO As Int64 = 0
        Dim ORDR_NO As String = ""
        Dim CUST_CODE As String = ""
        Dim errFound As Boolean = False

        If dst.Tables.Contains("SOTORDR1") Then
            If dst.Tables("SOTORDR1").Rows.Count = 1 Then
                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows(0)
                ORDR_NO = rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty
                CUST_CODE = rowSOTORDR1.Item("CUST_CODE").ToString & String.Empty
            Else
                MsgBox("Error in Form.  Please Let ABS Know", vbCritical, "Hmm")
                Return RetVal
                Exit Function
            End If
        Else
            MsgBox("Error in Form.  Please Let ABS Know", vbCritical, "Hmm")
            Return RetVal
            Exit Function
        End If
        If dst.Tables.Contains("SOTORDR2") Then
            If dst.Tables("SOTORDR2").Rows.Count > 0 Then
                Dim filter As String = ""
                ORDR_LNO = Val(dst.Tables("SOTORDR2").Compute("max(ORDR_LNO)", filter)) + 1
            Else
                ORDR_LNO = 1
            End If
        Else
            MsgBox("Error in Form.  Please Let ABS Know", vbCritical, "Hmm")
            Return RetVal
            Exit Function
        End If
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            'Dim filter As String = "xlsb files (*.xlsb)|*.xlsx|All files (*.*)|*.*"
            Dim filter As String = "All files (*.*)|*.*"
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            ASCMAIN1.Progress("Now Building Order From Excel", "")
            Cursor = Cursors.WaitCursor
            Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
            Dim xws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            xws = XWB.Worksheets(1)
            Try
                Dim COLUMNS As New Dictionary(Of String, Int64)
                Dim COLIST As New List(Of String)
                COLIST.Add(("Style Code").ToUpper)
                COLIST.Add(("Color Code").ToUpper)
                COLIST.Add(("Order Qty").ToUpper)
                COLIST.Add(("Price").ToUpper)
                COLIST.Add(("Cust SKU").ToUpper)
                COLIST.Add(("Cust Style").ToUpper)
                COLIST.Add(("Cust Color").ToUpper)

                'DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                Dim BeginFound As Boolean = False
                Dim EndFound As Boolean = False
                Dim BlankRows As Int64 = 0
                For CurRow As Int64 = 1 To 2000
                    If BeginFound And Not EndFound Then
                        If IsNothing(xws.Cells(CurRow, 1).value) Then
                            EndFound = True
                        Else
                            Dim STYLE_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "STYLE_CODE")
                            Dim COLOR_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "COLOR_CODE")
                            Dim CUST_SKU As String = GetValueFromExcel(xws, COLUMNS, CurRow, "CUST_SKU")
                            Dim CUST_STYLE_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "CUST_STYLE_CODE")
                            Dim CUST_COLOR_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "CUST_COLOR_CODE")
                            Dim QTY_STR As String = GetValueFromExcel(xws, COLUMNS, CurRow, "ORDR_QTY")
                            Dim PRICE_STR As String = GetValueFromExcel(xws, COLUMNS, CurRow, "ORDR_UNIT_PRICE")
                            Dim QTY As Int64 = 0
                            Dim STYLE_ASST_QTY As Int64 = 1
                            Dim eMsg As New Text.StringBuilder With {.Length = 0}
                            If IsNumeric(QTY_STR) Then
                                QTY = Val(QTY_STR)
                                If QTY > 1000 Then
                                    eMsg.AppendLine("- QTY > 1000")
                                End If
                                If QTY < 1 Then
                                    eMsg.AppendLine("- QTY < 1")
                                End If
                            Else
                                eMsg.AppendLine("- Non-Numeric QTY.")
                            End If
                            Dim PRICE As Decimal = 0.00
                            If IsNumeric(PRICE_STR) Then
                                PRICE = Val(PRICE_STR)
                                If PRICE < 0 Then
                                    eMsg.AppendLine("- Price < 0.")
                                End If
                                If PRICE > 1500 Then
                                    eMsg.AppendLine("- Price > 1500.")
                                End If
                            Else
                                eMsg.AppendLine("- Non-Numeric Price.")
                            End If
                            Dim ordrRow As Int64 = dst.Tables("SOTORDR2").Select($"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'").Count
                            If ordrRow > 0 Then
                                eMsg.AppendLine("- Style / Color Already On Order.")
                            End If
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            If Not IsNothing(rowICTSTYL1) Then
                                Dim TMP As String = rowICTSTYL1.Item("STYLE_ASST_QTY").ToString & String.Empty
                                If IsNumeric(TMP) Then
                                    If Val(TMP) > 1 Then
                                        STYLE_ASST_QTY = Val(TMP)
                                        QTY = QTY * STYLE_ASST_QTY
                                        PRICE = PRICE / STYLE_ASST_QTY
                                    End If
                                End If
                                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                                SQLS.AppendLine("SELECT COUNT(*)")
                                SQLS.AppendLine("FROM ICTSTYC1")
                                SQLS.AppendLine($"WHERE STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'")
                                ASCMAIN1.sql = SQLS.ToString()
                                If Val(ASCDATA1.GetDataValue) = 0 Then
                                    eMsg.AppendLine("- Invlid Style / Color.")
                                End If
                            Else
                                eMsg.AppendLine($"- Invalid Style: {STYLE_CODE}")
                            End If

                            If eMsg.Length > 0 Then
                                If Not IsNothing(xws.Range($"A{CurRow}").Comment) Then
                                    xws.Range($"A{CurRow}").Comment.Delete()
                                End If
                                xws.Range($"A{CurRow}").AddComment(eMsg.ToString)
                                errFound = True
                            Else
                                If Not IsNothing(xws.Range($"A{CurRow}").Comment) Then
                                    xws.Range($"A{CurRow}").Comment.Delete()
                                End If
                                Dim rowSOTORDR2 As DataRow = Nothing
                                rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                                Dim ORDR_UNIT_PRICE_STD As Decimal = 0
                                Dim ORDR_PRICE_SOURCE As String = ""
                                Dim ORDR_UNIT_PRICE_CALC As Decimal = TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1,
                                           STYLE_CODE, COLOR_CODE, QTY, ORDR_PRICE_SOURCE)
                                With rowSOTORDR2
                                    .Item("ORDR_NO") = ORDR_NO
                                    .Item("ORDR_LNO") = ORDR_LNO
                                    .Item("STYLE_CODE") = STYLE_CODE
                                    .Item("COLOR_CODE") = COLOR_CODE
                                    .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC").ToString & String.Empty
                                    .Item("ORDR_QTY") = QTY
                                    .Item("ORDR_QTY_OPEN") = QTY
                                    .Item("ORDR_QTY_ORIG") = QTY
                                    .Item("ORDR_QTY_ALLO") = 0
                                    .Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY").ToString & String.Empty
                                    .Item("ORDR_EXTD_COST") = 0
                                    .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM").ToString & String.Empty
                                    .Item("ORDR_QTY_PICK") = 0
                                    .Item("ORDR_QTY_SHIP") = 0
                                    .Item("ORDR_QTY_CANC") = 0
                                    .Item("ORDR_STATUS") = "O"
                                    .Item("ORDR_QTY_PRE_ALLO") = 0
                                    .Item("QTY_PER_PP") = 0
                                    .Item("ORDR_PRICE_SOURCE") = ORDR_PRICE_SOURCE
                                    .Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY").ToString & String.Empty
                                    .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString & String.Empty
                                    '.Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE").ToString & String.Empty
                                    .Item("STYLE_PRICE") = rowICTSTYL1.Item("STYLE_PRICE").ToString & String.Empty
                                    .Item("STYLE_RETAIL") = 0
                                    .Item("PO_COST") = 0
                                    .Item("COMM_RATE") = 0
                                    .Item("ORDR_UNIT_PRICE_CALC") = ORDR_UNIT_PRICE_CALC
                                    If PRICE <> 0 Then
                                        .Item("ORDR_UNIT_PRICE") = PRICE
                                        .Item("ORDR_UNIT_PRICE_CURR") = PRICE
                                        .Item("ORDR_UNIT_PRICE_MANUAL") = "1"
                                    Else
                                        .Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_CALC
                                        .Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_CALC
                                    End If
                                    If CUST_SKU.Length > 0 Then
                                        .Item("CUST_SKU") = CUST_SKU
                                    End If
                                    If CUST_STYLE_CODE.Length > 0 Then
                                        .Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                                    End If
                                    If CUST_COLOR_CODE.Length > 0 Then
                                        .Item("CUST_COLOR_CODE") = CUST_COLOR_CODE
                                    End If
                                End With
                                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
                                ORDR_LNO = ORDR_LNO + 1
                            End If
                        End If
                    Else
                        BlankRows += 1
                        If BlankRows >= 10 Then
                            Exit For
                        End If
                        For R As Int64 = 1 To 7
                            If Not IsNothing(xws.Cells(CurRow, R).value) Then
                                Dim COL As String = xws.Cells(CurRow, R).value.ToString.Trim.ToUpper
                                If COLIST.Contains(COL) Then
                                    COLUMNS.Add(COL, R)
                                    BeginFound = True
                                End If
                            End If
                        Next
                        If BeginFound Then
                            If Not (COLUMNS.ContainsKey(("Style Code").ToUpper) And COLUMNS.ContainsKey(("Color Code").ToUpper)) Then
                                BeginFound = False
                            End If
                        End If
                    End If
                Next
                XWB.Save()
                excel.Visible = True
                xws = Nothing
                XWB = Nothing
                excel = Nothing

                'XWB.Close()
                'excel.Quit()
                'If Not IsNothing(xws) Then
                '    Runtime.InteropServices.Marshal.ReleaseComObject(xws)
                'End If
                'If Not IsNothing(xws) Then
                '    Runtime.InteropServices.Marshal.ReleaseComObject(xws)
                '    xws = Nothing
                'End If
                'If Not IsNothing(XWB) Then
                '    Runtime.InteropServices.Marshal.ReleaseComObject(XWB)
                '    XWB = Nothing
                'End If
                'If Not IsNothing(excel) Then
                '    Runtime.InteropServices.Marshal.ReleaseComObject(excel)
                '    excel = Nothing
                'End If
                GC.Collect()
                GC.WaitForPendingFinalizers()
                'excel.Quit()
                'xws = Nothing
                'XWB = Nothing
                'excel = Nothing
            Catch ex As Exception
                If Not IsNothing(xws) Then
                    Runtime.InteropServices.Marshal.ReleaseComObject(xws)
                End If
                If Not IsNothing(XWB) Then
                    Runtime.InteropServices.Marshal.ReleaseComObject(XWB)
                End If
                If Not IsNothing(excel) Then
                    Runtime.InteropServices.Marshal.ReleaseComObject(excel)
                End If
            End Try

        End If
        If errFound Then
            RetVal.Length = 0
            RetVal.AppendLine("Errors Found In Details.")
            RetVal.AppendLine("Please See Excel Notes For Details.")
        End If
        Return RetVal
    End Function

    Private Shared Function GetValueFromExcel(xws As Microsoft.Office.Interop.Excel.Worksheet, ByVal COLUMNS As Dictionary(Of String, Int64), curRow As Long, ByVal CODE As String) As String
        Dim RetVal As String = ""
        Select Case CODE
            Case "STYLE_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Style Code").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Style Code").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "COLOR_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Color Code").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Color Code").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "ORDR_QTY"
                RetVal = "0"
                If COLUMNS.ContainsKey(("Order Qty").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Order Qty").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "ORDR_UNIT_PRICE"
                RetVal = "0"
                If COLUMNS.ContainsKey(("Price").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Price").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "CUST_SKU"
                RetVal = ""
                If COLUMNS.ContainsKey(("Cust SKU").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Cust SKU").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "CUST_STYLE_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Cust Style").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Cust Style").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "CUST_COLOR_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Cust Color").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Cust Color").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
        End Select

        Return RetVal
    End Function

#End Region
End Class