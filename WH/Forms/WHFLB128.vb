Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing.Printing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFLB128


#Region "Declarations"

    Dim SOTPICKX As String

    Dim SHIP_BOL_NOs As String = ""
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

    Dim SOTORDR0 As String = ""
    Dim SOTPICK1 As String = ""
    Dim WHTILOCS As String = ""

    Dim CUST_CODEs_856 As New List(Of String)
    Dim SO_PARM_UPC_VENDOR_ID As String = ""
    Dim sqlSOTCART1 As String
    Dim printingPickTickets As Boolean = False

    Dim PICK_GROUP_LINES As Decimal = 0
    Dim UCC128TAB As Integer
    Dim NLAB1TAB As Integer

    Dim ORDR_NO_MT As String
    Private sqlDerelease As String = String.Empty
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        'SetUpPortsAndPrinters()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Refresh_SOTPICKX("")

        Get_PARM("SOTPARM1")
        SO_PARM_UPC_VENDOR_ID = ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID")

        AUDIT.Add("ARTCUST2", "E")

        With dst

            ASCMAIN1.sql = "Select * from " & SOTPICKX
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False)
            .Tables("SOTPICKX").Columns("SHIP_BOL_NO").AllowDBNull = True
            'WHY IN THE WORLD WOULD THIS BE NULL?
            With .Tables("SOTPICKX").Columns
                .Add("SELECTED", GetType(System.String))
                .Add("OPT_PICK_TICKET", GetType(System.String))
                .Add("OPT_UCC128", GetType(System.String))
                .Add("OPT_PULL_STORE", GetType(System.String))
                .Add("OPT_PULL_STYLE", GetType(System.String))
                .Add("OPT_MANIFEST", GetType(System.String))
                .Add("CUST_856", GetType(System.String))
            End With
            .Tables("SOTPICKX").Columns("SELECTED").DefaultValue = "0"

            ASCMAIN1.sql = "Select SOTPICK1.PICK_NO" & vbCrLf _
                & ", SOTPICK1.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTSHIP1" & vbCrLf _
                & " where SOTORDR1.CUST_CODE = :PARM1" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = :PARM2" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_BATCH_NO = :PARM3" & vbCrLf _
                & "   and SOTSHIP1.SHIP_ADDR_TYPE = :PARM4" & vbCrLf _
                & "   and (SOTSHIP1.SHIP_ADDR_TYPE = 'MK' or SOTSHIP1.SHIP_ADDR_CODE = :PARM5)"
            Create_TDA(.Tables.Add, "SOTPICKY", "**", 0, False, "VVVVV", 1)

            ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
                & ", SOTORDR0.CUST_CODE" & vbCrLf _
                & ", DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'DC',SOTSHIP1.SHIP_BOL_NO,'MK') SHIP_BOL_NO_X" & vbCrLf _
                & " from SOTSHIP1,SOTORDR0" & vbCrLf _
                & " where SOTSHIP1.ORDR_GROUP_NO = :PARM1" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "V", 1, "SHIP_PICK_PRINTED")
            ' .Tables("SOTSHIP1").Columns.Add("SHIP_BOL_NO_BC")

            ASCMAIN1.sql = "Select * from SOTORDR0 where ROWNUM <1"
            SOTORDR0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")

            ASCMAIN1.sql = "Select SOTORDR0.* from " & SOTORDR0 & " SOTORDR0 "
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTORDR1.*, 'MK' AS MARK_FOR, 'ST' AS SHIP_TO" _
                & " from SOTORDR1," & SOTORDR0 & " SOTORDR0" _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)
            .Tables("SOTORDR1").Columns("ORDR_SHIP_INSTR").MaxLength = 512

            ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTCOLR1.COLOR_CODE_LONG, ICTSIZE1.SIZE_CODE SIZE_DESC, SOTORDR1.CUST_CODE" & vbCrLf _
                & " from SOTORDR2,SOTORDR1," & SOTORDR0 & " SOTORDR0, ICTSIZE1, ICTCOLR1 " & vbCrLf _
                & " where ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and ICTSIZE1.NRF_SIZE_CODE (+) = SOTORDR2.CUST_SIZE_CODE" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR3.*" & vbCrLf _
                & " from SOTORDR3,SOTORDR1," & SOTORDR0 & " SOTORDR0 " & vbCrLf _
                & " where SOTORDR3.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR3", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select SOTORDR9.*" & vbCrLf _
                & " from SOTORDR9,SOTORDR1," & SOTORDR0 & " SOTORDR0 " & vbCrLf _
                & " where SOTORDR9.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR9", "**", 0, False, "", 2)

            'ASCMAIN1.sql = "Select X.*, SOTSVIA2.TRANSIT_BUS_DAYS" & vbCrLf _
            '    & " from (" & vbCrLf _
            '    & " Select SOTORDR5.*, SOTORDR1.WHSE_CODE" & vbCrLf _
            ASCMAIN1.sql = "Select SOTORDR5.*, SOTORDR1.WHSE_CODE" & vbCrLf _
                & " from SOTORDR5,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
                & " where SOTORDR5.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
               & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, True, "", 2)

            Create_TDA(.Tables.Add, "SOTLABL1", "*", 1, False, "", 2)

            Create_TDA(.Tables.Add, "SOTPICK0", "*", 1, False)
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)
            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, False)
            .Tables("ARTCUST2").Columns.Add("CUST_ADDR_DC")
            .Tables("ARTCUST2").Columns.Add("DC_STATE")

            ASCMAIN1.sql = "Select SOTPICK1.SHIP_BOL_NO, SOTPICK1.PICK_NO, 0 LABEL_COUNTER" & vbCrLf _
                & ", SOTPICK0.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DEPT, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO" & vbCrLf _
                & ", STYLE_CODE STYLE_CODE_1, STYLE_CODE STYLE_CODE_2, STYLE_CODE STYLE_CODE_3" & vbCrLf _
                & " from SOTPICK1,SOTPICK0,SOTORDR1,SOTORDR2" & vbCrLf _
                & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK3", "**", 0, False, "", 3)

            ' .Tables("SOTPICKL").Columns.Add("LOCATIONS")

            ASCMAIN1.sql = "Select Distinct EDT850T2.*" & vbCrLf _
                & " from EDT850T2,EDT850T1,SOTORDR1," & SOTORDR0 & " SOTORDR0 " & vbCrLf _
                & " where EDT850T1.EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "   and EDT850T2.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO"
            Create_TDA(.Tables.Add, "EDT850T2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select Distinct EDT850T1.*" & vbCrLf _
                & " from EDT850T1,SOTORDR1," & SOTORDR0 & " SOTORDR0 " & vbCrLf _
                & " where EDT850T1.EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "EDT850T1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTPICK1.*,SOTORDR1.ORDR_DEPT,SOTORDR1.ORDR_CUST_PO from SOTPICK1,SOTORDR1 where ROWNUM < 1"
            '  ASCMAIN1.sql = "Select * from SOTPICK1 where ROWNUM <1"
            SOTPICK1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK1 & " Add Primary Key (PICK_NO)")

            ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR1.CUST_STORE_NO, ARTCUST2.CUST_NAME CUST_STORE_NAME" & vbCrLf _
                & ", TRIM(SUBSTR(LPAD(SOTORDR1.CUST_STORE_NO,6,' '),3,4)) CUST_STORE_NO4, SOTSREP1.SREP_NAME " & vbCrLf _
                & " from " & SOTPICK1 & " SOTPICK1, SOTORDR1, ARTCUST2, SOTSREP1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' and ARTCUST2.CUST_ADDR_CODE (+) = SOTORDR1.CUST_STORE_NO AND SOTSREP1.SREP_CODE (+) = SOTORDR1.SREP_CODE "
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "", 1)
            ' .Tables("SOTPICK1").Columns.Add("PICK_NO_BC", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("CART_SERIAL_NO", GetType(System.Int32))
            .Tables("SOTPICK1").Columns.Add("PICK_TOTAL_QTY", GetType(System.Int32))

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, nvl(SOTORDR2.CUST_SIZE_CODE,'AST') CUST_SIZE_CODE, ICTSTYL1.CARTON_PACK_QTY, SOTPICK1.SHIP_BOL_NO," & vbCrLf _
                & "  ICTSTYC1.STYLE_BIN, ICTSTYC1.STYLE_BIN as LOCATION_CODE, ICTSTYL1.CASE_CUBE, ICTSTYC1.UPC_CODE, ICTSTYL1.CASE_WEIGHT_GRS, nvl(ICTSTYL1.CARTONS_PER_UNIT, 0) CARTONS_PER_UNIT, SOTORDR2.CUST_SKU" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "RGI", ", nvl(ICTSTYL1.STYLE_ASST_QTY,0) STYLE_ASST_QTY" & vbCrLf, "") _
                & IIf(ASCMAIN1.CLIENT = "RGI", ", ICTSTYL1.WHSE_MESSAGE" & vbCrLf, "") _
                & " from SOTPICK2," & SOTPICK1 & " SOTPICK1, SOTORDR2, ICTSTYL1, ICTSTYC1" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTPICK2.PICK_QTY <> 0"  'To avoid picking up records representing a cancellation or backorder generated during Pick Ticket Release
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, True, "", 2)

            If Not .Tables("SOTPICK2").Columns.Contains("LOCATION_ROUTE_SEQ") Then
                .Tables("SOTPICK2").Columns.Add("LOCATION_ROUTE_SEQ", GetType(System.Int32))
            End If

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.STYLE_DESC, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " from SOTPICK2, SOTPICK1, SOTORDR2" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK2.PICK_QTY <> 0" & vbCrLf _
                & "   and SOTPICK2.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICKL", "**", 0, False, "V", 2)
            .Tables("SOTPICKL").Columns.Add("LABEL_QTY", GetType(System.Int32))

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            .Tables("SOTPICK1").Columns.Add("PICK_TOT", GetType(System.Int32), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")

            sqlSOTCART1 = "Select SOTCART1.*,SOTPICK1.SHIP_BOL_NO,SOTPICK1.ORDR_NO" & vbCrLf _
                & ", SUBSTR(SOTCART1.CART_NO,11,9) CART_NO_9" & vbCrLf _
                & ", SUBSTR(SOTCART1.CART_NO,20,1) CART_NO_DIGIT" & vbCrLf _
                & ", SUBSTR(SOTCART1.CART_NO,5,6) CART_NO_PFX" & vbCrLf _
                & ", '(00) 0 0 ' || SUBSTR(SOTCART1.CART_NO,5,6) || ' ' || SUBSTR(SOTCART1.CART_NO,11,9) || SUBSTR(SOTCART1.CART_NO,20,1) CART_NO_FMT" & vbCrLf
            ASCMAIN1.sql = sqlSOTCART1 _
                & " from SOTCART1," & SOTPICK1 & " SOTPICK1" _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO"

            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "", 1)
            .Tables("SOTCART1").Columns.Add("CART_1_OF_9", GetType(System.String))
            .Tables("SOTCART1").Columns.Add("STYLE_CODE")
            .Tables("SOTCART1").Columns.Add("STYLE_DESC")

            .Tables("SOTCART1").Columns.Add("NUM_CARTONS", GetType(System.Decimal))
            .Tables("SOTCART1").Columns.Add("TOTAL_WEIGHT", GetType(System.Decimal))
            .Tables("SOTCART1").Columns.Add("CASE_CUBE_FT", GetType(System.Decimal))

            .Tables("SOTCART1").Columns("NUM_CARTONS").DefaultValue = 0
            .Tables("SOTCART1").Columns("TOTAL_WEIGHT").DefaultValue = 0
            .Tables("SOTCART1").Columns("CASE_CUBE_FT").DefaultValue = 0

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")

            With .Tables("SOTCART1").Columns
                .Add("CART_SEQ_MAX", GetType(System.Int32))
                Create_Relation("SOTORDR1", "SOTCART1", "ORDR_NO")

                .Add("CUST_STORE_NO", GetType(System.String), "PARENT(SOTORDR1_SOTCART1).CUST_STORE_NO")
                .Add("CUST_ZIP_CODE", GetType(System.String))
                .Add("CART_SERIAL_NO", GetType(System.Int32))
                .Add("CUST_CODE", GetType(System.String), "PARENT(SOTORDR1_SOTCART1).CUST_CODE")
                .Add("CUST_SELECTED", GetType(System.String))
                .Add("CUST_DC_NO", GetType(System.String), "PARENT(SOTORDR1_SOTCART1).CUST_DC_NO")
            End With

            ASCMAIN1.sql = "Select SOTCART2.*" _
                & ", DECODE(SOTORDR2.RANGE_STYLE_CODE,'',SOTORDR2.STYLE_CODE, SOTORDR2.RANGE_STYLE_CODE) CART_ITEM" & vbCrLf _
                & ", ICTSTYL1.SUB_UNIT_PACK_QTY" & vbCrLf _
                & " from SOTCART2,SOTCART1,SOTORDR2,ICTSTYL1," & SOTPICK1 & " SOTPICK1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTCART2.STYLE_CODE"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "", 2)

            Create_Relation("SOTORDR2", "SOTCART2", "ORDR_NO,ORDR_LNO")
            dst.Tables("SOTORDR2").Columns.Add("CART_COUNT", GetType(System.Int32), "COUNT(CHILD(SOTORDR2_SOTCART2).CART_LNO)")

            'Changed GTIN_UPC_CODE to use GTIN on 3/21/06 - WR.
            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, SIZE_CODE" & vbCrLf _
                & ", UPC_CODE, GTIN_PACK_CODE, GTIN_CODE, GTIN_DESC" & vbCrLf _
                & " from ICVLUPC1, ICTGTINT" & vbCrLf _
                & " where ICVLUPC1.UPC_CODE = ICTGTINT.GTIN_UPC_CODE"
            Create_TDA(.Tables.Add, "SOTGTINT", "**", 0, False, "", 4)

            ASCMAIN1.sql = "" _
                & "Select RANGE_UPC_CODE UPC_CODE, CUST_CODE, RANGE_SKU CUST_SKU" & vbCrLf _
                & " from ICTRSTY1" & vbCrLf _
                & " where RANGE_UPC_CODE is NOT NULL" & vbCrLf _
                & "   and RANGE_SKU is NOT NULL" & vbCrLf _
                & " union " & vbCrLf _
                & "Select CUST_UPC UPC_CODE, CUST_CODE, CUST_STYLE_CODE CUST_SKU" & vbCrLf _
                & " from SOTCSTY1" & vbCrLf _
                & " where CUST_UPC is NOT NULL" & vbCrLf _
                & "   and CUST_STYLE_CODE is NOT NULL"
            Create_TDA(.Tables.Add, "SOTCSKU1", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select ICTSTYL1.* from ICTSTYL1" & vbCrLf _
                & " where STYLE_CODE in" & vbCrLf _
                & "(Select Distinct SOTORDR2.STYLE_CODE from SOTORDR2,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
                & " where SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO)"
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1" & vbCrLf _
                & " where (STYLE_CODE,COLOR_CODE) in" & vbCrLf _
                & "(Select Distinct SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE from SOTORDR2,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
                & " where SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO)" & vbCrLf _
                & "  and WHTLOCB1.LOCATION_QTY > 0"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "", 5)

            ASCMAIN1.sql = "Select ICTSTYC1.* from ICTSTYC1" & vbCrLf _
                & " where (STYLE_CODE,COLOR_CODE) in" & vbCrLf _
                & "(Select Distinct SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE from SOTORDR2,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
                & " where SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO)"
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False, "", 2)
            .Tables("ICTSTYC1").Columns.Add("LOCATION_CODES")

            For Each TABLE_NAME As String In New String() {"SOTSDIV1", "ICTWHSE1", "SOTSVIA1", "TATTERM1", "SOTUCCL1"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            With .Tables.Add("SOTCART3")
                .Columns.Add("SHIP_BOL_NO")
                .Columns.Add("CART_SERIAL_NO", GetType(Int64))
                .Columns.Add("CART_TOTAL", GetType(Int64))
                .PrimaryKey = New DataColumn() { .Columns("SHIP_BOL_NO"), .Columns("CART_SERIAL_NO")}
            End With

            With .Tables.Add("SOTCART4")
                .Columns.Add("CART_NO")
                .Columns.Add("PICK_NO")
                .Columns.Add("CUST_COUNTRY")
                .Columns.Add("ORDR_CUST_PO")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("SKU")
                .Columns.Add("STYLE_DESC")
                .Columns.Add("CART_TOTAL_UNITS")
                .Columns.Add("UPC_CODE")
                .Columns.Add("CART_WEIGHT")
                .Columns.Add("CASE_CBM")
                .Columns.Add("ORIGIN_COUNTRY")
                .Columns.Add("CART_1_OF_9", GetType(System.String))
                .PrimaryKey = New DataColumn() { .Columns("CART_NO")}
            End With

            With dst.Tables.Add("SOTRANG1")
                .Columns.Add("EDI_SKU")
                .Columns.Add("EDI_UPC")
            End With

            ASCMAIN1.sql = "Select X.WHSE_CODE, ICTWHSE1.WHSE_DESC" & vbCrLf _
                & ", X.ORDR_CNT_PICK, X.ORDR_AMT_PICK" & vbCrLf _
                & ", Y.ORDR_CNT_PICK_UNPRINTED, Y.ORDR_AMT_PICK_UNPRINTED" & vbCrLf _
                & "  from ICTWHSE1" & vbCrLf _
                & ", (Select SOTORDR0.WHSE_CODE, Sum(ORDR_CNT_PICK) ORDR_CNT_PICK, Sum (ORDR_AMT_PICK) ORDR_AMT_PICK " & vbCrLf _
                & "  from SOTORDR0" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "VAN", ", EDT850T1" & vbCrLf, "") _
                & "where SOTORDR0.ORDR_CNT_PICK <> 0" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "VAN", " and EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO and ((SOTORDR0.CUST_CODE = 'KOHLS' and EDT850T1.EDI_DEPT_DESC = 'PACK BY STORE') or (SOTORDR0.CUST_CODE = 'WALMART' ))" & vbCrLf, "") _
                & " group by SOTORDR0.WHSE_CODE) X" & vbCrLf _
                & ", (Select SOTSHIP1.WHSE_CODE, Count (Distinct SOTPICK1.PICK_NO) ORDR_CNT_PICK_UNPRINTED" & vbCrLf _
                & ", Sum (SOTPICK2.PICK_QTY * SOTPICK2.PICK_UNIT_PRICE) ORDR_AMT_PICK_UNPRINTED" & vbCrLf _
                & "from SOTSHIP1,SOTPICK1,SOTPICK2, SOTORDR1 where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & " and SOTSHIP1.SHIP_STATUS IN ('P', 'H')" & vbCrLf _
                & " and SOTSHIP1.SHIP_PICK_PRINTED is Null" & vbCrLf _
                & " and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "RGI", " and NVL(SOTORDR1.ORDR_TYPE_CODE,'') <> 'B2C'" & vbCrLf, "") _
                & IIf(ASCMAIN1.CLIENT = "NYA" And ASCMAIN1.USER_CODES = "CA", "  and SOTSHIP1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")" & vbCrLf, "") _
                & "group by SOTSHIP1.WHSE_CODE" & vbCrLf _
                & ") Y" & vbCrLf _
                & "  where ICTWHSE1.WHSE_CODE = X.WHSE_CODE " & vbCrLf _
                & IIf(ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN", "", " and ICTWHSE1.LP_CODE is Null" & vbCrLf) _
                & IIf(ASCMAIN1.CLIENT = "NYA" And ASCMAIN1.USER_CODES = "CA", "  and X.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")" & vbCrLf, "") _
                & "    and Y.WHSE_CODE (+) = X.WHSE_CODE"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False)
            .Tables("ICTWHSEX").Columns("ORDR_CNT_PICK").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select * from SOTORDXR where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDXR", "**", 0, True, "V")

            sqlDerelease = "Select SOTPICK1.*, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO from SOTPICK1, SOTORDR1 where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTDREL1", sqlDerelease, 0, False, String.Empty, 0)

            ASCMAIN1.sql = "Select SOTNLAB2.* from SOTNLAB2 WHERE SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTNLAB2", "**", 0, False, "V", 2)
            .Tables("SOTNLAB2").Columns.Add("PRETICKET_STR", GetType(System.String), "IIF(PRE_TICKET=1, 'YES', 'NO')")
            .Tables("SOTNLAB2").Columns.Add("CART_TOT", GetType(System.Int32), "MAX(CART_NO)")


            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then

                ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, STYLE_BIN FROM ICTSTYC1 WHERE ROWNUM < 1"
                WHTILOCS = ASCMAIN1.Temp_Table
                ASCDATA1.ExecuteSQL("Alter Table " & WHTILOCS & " Add Primary Key (STYLE_CODE, COLOR_CODE)")

                ASCMAIN1.sql = "Select WHTILOCS.* from " & WHTILOCS & " WHTILOCS "
                Create_TDA(.Tables.Add("WHTILOCS"), WHTILOCS, "**", 0)
                ' Create_TDA(.Tables.Add, "WHTILOCS", "**", 0, False, "V", 2)

                ASCMAIN1.sql = "Select PICK_NO, PICK_LNO from SOTPICK2 WHERE ROWNUM <1"
                Create_TDA(.Tables.Add, "SOTPLOC1", "**", 0, False)
                With .Tables("SOTPLOC1").Columns
                    .Add("LOCATION_CODE", GetType(System.String))
                    .Add("LOCATION_QTY", GetType(System.String))
                    .Add("SEQUENCE", GetType(System.String))
                End With
                '.Tables("SOTPLOC1").Columns("SELECTED").DefaultValue = "0"

                ASCMAIN1.sql = "Select ARTCUSTQ.* from ARTCUSTQ WHERE CUST_CODE = :PARM1 and CUST_ADDR_CODE = :PARM2"
                Create_TDA(.Tables.Add, "ARTCUSTQ", "**", 0, False, "VV", 2)
            End If



            If ASCMAIN1.CLIENT = "VAN" Then

                ASCMAIN1.sql = "Select WHTSCSEQ.* from WHTSCSEQ"
                Create_TDA(.Tables.Add, "WHTSCSEQ", "**", 0, False, , 3)
                Fill_Records("WHTSCSEQ")

            End If
        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICKL.DataSource = dst.Tables("SOTPICKL")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTNLAB2.DataSource = dst.Tables("SOTNLAB2")

        grdSOTPICKX.DisplayLayout.Bands(0).Columns("CCPA_NO_STATUS").Hidden = Not (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI")

        With grdSOTPICKX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            If ASCMAIN1.CLIENT = "RGI" Then
                .Columns("PICK_NO_CONS").Hidden = False
                .Columns("PICK_RELEASED").Hidden = False
            End If
            For Each COLUMN_NAME As String In New String() {"SELECTED", "SHIP_BOL_NO", "ORDR_GROUP_NO", "PICK_BATCH_NO", "CUST_CODE", "CUST_CODE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                '    If New String() {"SELECTED", "OPT_UCC128"}.Contains(GCOL.Key) Then
                '        GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                '    Else
                '        GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                '    End If
                '    If New String() {"OPT_PICK_TICKET", "OPT_PULL_STORE", "OPT_PULL_STYLE", "OPT_MANIFEST"}.Contains(GCOL.Key) Then
                '        GCOL.Hidden = True
                '    End If
                'Else
                If New String() {"SELECTED", "OPT_PICK_TICKET", "OPT_UCC128", "OPT_PULL_STORE", "OPT_PULL_STYLE",
                                  "OPT_MANIFEST", "ORDR_HIGH_PRIORITY"}.Contains(GCOL.Key) Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit

                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = Color.WhiteSmoke
                End If

                If New String() {"OPT_PICK_TICKET", "OPT_UCC128", "OPT_PULL_STORE", "OPT_PULL_STYLE", "ORDR_HIGH_PRIORITY",
                                 "OPT_MANIFEST"}.Contains(GCOL.Key) Then
                    GCOL.Hidden = True
                End If
                'End If
            Next
        End With

        With grdSOTPICKL.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"LABEL_QTY"}.Contains(GCOL.Key) Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = Color.Beige
                End If
            Next
        End With

        Create_Summary(grdSOTPICKX, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTPICKX, New String() {"SELECTED", "ORDR_CNT_PICK", "ORDR_QTY_PICK", "ORDR_AMT_PICK", "ORDR_CNT_CART"})
        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")

        Create_Summary(grdSOTPICKL, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICKL, New String() {"PICK_QTY", "LABEL_QTY"})

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEX, New String() {"ORDR_CNT_PICK", "ORDR_AMT_PICK", "ORDR_CNT_PICK_UNPRINTED", "ORDR_AMT_PICK_UNPRINTED"})

        Show_Filter(grdSOTPICKX, True)
        grdSOTPICKX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.sql = "Select EDTTRPM1.CUST_CODE from EDTTRPM1 " & vbCrLf _
                & " where (EDI_STATUS = 'P' or EDI_STATUS = 'T')" & vbCrLf _
                & "   and EDI_DOC_NO = '856'" & vbCrLf _
                & "   and CUST_CODE is Not Null"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            CUST_CODEs_856.Add(CUST_CODE)
        Next

        Bind_Controls(grpSHIPTO, "SOTORDR5")
        'Bind_Controls(grpSHIPTO, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'ST'", "", DataViewRowState.CurrentRows))

        ASCMAIN1.Add_Value_List(grdSOTPICKX, "ORDR_SOURCE", Nothing, New String() {":", "K:Keyboard", "W:Web", "E:EDI"})

        Dim ZebraPrinters As New List(Of String)
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                If printerName.ToUpper.StartsWith("ZDESIGNER") Or printerName.ToUpper.StartsWith("MONARCH") Or printerName.ToUpper.StartsWith("AVERY") Or printerName.ToUpper.StartsWith("ZEBRA") Then
                    ZebraPrinters.Add(printerName)
                End If
            Next printerName
            If ZebraPrinters.Count >= 1 Then
                cboZebraPrinter.DataSource = ZebraPrinters
            End If
        End If

        For I As Integer = 0 To tabLabels.Tabs.Count - 1
            If tabLabels.Tabs(I).Text = "UCC128" Then
                UCC128TAB = I
            End If
            If tabLabels.Tabs(I).Text.ToUpper = "MANUAL LABELS" Then
                NLAB1TAB = I
            End If
        Next

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                End If

            Case "Refresh"


            Case "Print"
                Dim ORDR_GROUP_NOs As String = Get_Selected_ORDR_GROUP_NOs()

                If dst.Tables("SOTPICKX").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "You Must First Select Something to Print"
                End If

                ORDR_GROUP_NOs = ""

                If ASCMAIN1.CLIENT = "VAN" Then
                    Dim PICK_NO_CONSs As New List(Of String)
                    For Each ROW As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                        Dim SHIP_BOL_NO As String = ROW.Item("SHIP_BOL_NO")
                        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                            Dim PICK_NO_CONS As String = rowSOTPICK1.Item("PICK_NO_CONS") & ""
                            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")

                            Dim rowreal As DataRow = LookUp("SOTPICK1", PICK_NO)
                            PICK_NO_CONS = rowreal.Item("PICK_NO_CONS") & ""
                            If Not PICK_NO_CONSs.Contains(PICK_NO_CONS) Then
                                PICK_NO_CONSs.Add(PICK_NO_CONS)
                                ASCMAIN1.sql = "Select PICK_NO, SHIP_BOL_NO from SOTPICK1 where PICK_NO_CONS = '" & PICK_NO_CONS & "' and PICK_STATUS = 'P' and PICK_NO <> '" & PICK_NO & "'"
                                For Each rowSOTPICK1other As DataRow In ASCDATA1.GetDataTable().Select("")
                                    Dim PICK_NOother As String = rowSOTPICK1other.Item("PICK_NO")
                                    Dim SHIP_BOL_NOother As String = rowSOTPICK1other.Item("SHIP_BOL_NO")
                                    Dim rowother As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NOother)

                                    Dim rowSOTPICKXother() = dst.Tables("SOTPICKX").Select("SHIP_BOL_NO = '" & SHIP_BOL_NOother & "'")
                                    If rowSOTPICKXother.Length <> 1 OrElse rowSOTPICKXother(0).Item("SELECTED") <> "1" Then
                                        EMsg &= vbCr & "You have not selected Shipment " & SHIP_BOL_NOother & ", which includes Pick Ticket " & PICK_NOother & " to be Consolidated with PT " & PICK_NO_CONS
                                        Exit For
                                    End If

                                Next
                                If EMsg <> "" Then Exit For
                            End If

                        Next
                        If EMsg <> "" Then Exit For
                    Next
                End If

            Case "Update"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                If ScreenMode Then
                    Refresh_SOTPICKX("")
                Else
                    Refresh_ICTWHSEX()
                End If

            Case "Print"
                EntryMode = "E"
                printingPickTickets = True
                Print_Pick_Tickets("")
                Print_Documents()


            Case "Cancel", "Done"
                Mode_Settings(False)


        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Load").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                '.Items("Refresh").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode
                '.Items("Re-Print Confirmed").Settings.Enabled = iScreenMode
                '.Items("Re-Print Confirmed").Visible = False

                '.Items("PO Summary").Visible = ScreenMode And ASCMAIN1.CLIENT = "VAN"

            End With
            .Groups("Selection Options").Visible = False 'ScreenMode
            '.Groups("Label Printing").Visible = ScreenMode


        End With

        grdICTWHSEX.Visible = Not ScreenMode
        grdSOTPICKX.Visible = ScreenMode

        'spl.Panel1Collapsed = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            'Setup_tabLabels()
        Else
            Clear_Record()
        End If


        'lblPICK_NO.Visible = False
        'txtPICK_NO.Visible = False
        'UltraLabel6.Visible = False
        'UltraTextEditor3.Visible = False
        cboZebraPrinter.Visible = True



    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDR2", "SOTPICK1", "SOTSHIP1", "SOTPICKX",
             "SOTCART2", "SOTPICK2", "SOTCART1", "SOTCART3", "SOTORDR5",
             "SOTORDR0", "SOTORDR9", "EDT850T2", "EDT850T1", "SOTRANG1",
             "SOTLABL1", "SOTPICK0", "ARTCUST1", "ARTCUST2", "WHTLOCB1", "ICTSTYC1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        WHSE_CODE = ""
        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                ' WHSE 18 FOR CA USERS
                Absx1.txtFor("WHSE_CODE").Text = TAC.TACMAIN1.NyaCanadaWhseArray(0) & String.Empty  '"18"
            Else
                Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""
            End If
        End If
        Refresh_ICTWHSEX()

    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Refresh_SOTPICKX("")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_Pick_Tickets(Optional PICK_NOs As String = "")

        ASCMAIN1.Progress("Now Preparing To Print Documents", "Please Wait ...")

        EnforceConstraints(False)
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTPICK1)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR0", "SOTORDR1", "SOTORDR2", "SOTORDR3", "SOTORDR5", "SOTORDR9",
             "SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTCART1", "SOTCART2",
             "EDT850T1", "EDT850T2", "SOTPICK0", "ARTCUST1", "ARTCUST2"}

            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)


        SHIP_BOL_NOs = ""

        If PICK_NOs <> "" Then
            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where PICK_NO In (" & PICK_NOs & ")"
            ASCDATA1.ExecuteSQL("Insert into " & SOTPICK1 & " " & ASCMAIN1.sql)
            ASCMAIN1.AnalyzeTable(SOTPICK1)

            Dim ORDR_GROUP_NO As String = grdSOTPICKX.ActiveRow.Cells("ORDR_GROUP_NO").Value
            Dim ORDR_GROUP_NOs As String = "'" & ORDR_GROUP_NO & "'"

            Fill_Records("SOTSHIP1", ORDR_GROUP_NO)

            Customer_Data(ORDR_GROUP_NOs)
            Get_SOTORDRx(ORDR_GROUP_NOs)
            Get_SOTPICKx()
            Get_EDI_Data()
        Else
            Dim g As Integer = 0



            If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") And dst.Tables("SOTPICKX").Select("SELECTED = '1' and PICK_NO_CONS <> ''").Count > 0 Then
                Dim PICK_NO_CONSs As New List(Of String)
                For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                    Dim PICK_NO_CONS As String = rowSOTPICKX.Item("PICK_NO_CONS") & ""
                    If Not PICK_NO_CONSs.Contains(PICK_NO_CONS) Then
                        PICK_NO_CONSs.Add(PICK_NO_CONS)
                    End If
                Next

                For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("PICK_NO_CONS <> ''")
                    'Dim danaGROUPno As String = rowSOTPICKX.Item("ORDR_GROUP_NO") & ""
                    If PICK_NO_CONSs.Contains(rowSOTPICKX.Item("PICK_NO_CONS") & "") Then
                        rowSOTPICKX.Item("SELECTED") = "1"
                    End If
                Next
            End If


            Dim ORDR_GROUP_NOs As String = ""

            Dim PICK_BATCH_NOs As String = ""

            For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                Dim ORDR_GROUP_NO As String = rowSOTPICKX.Item("ORDR_GROUP_NO")
                Dim SHIP_BOL_NO As String = rowSOTPICKX.Item("SHIP_BOL_NO") & ""
                Dim SHIP_ADDR_TYPE As String = rowSOTPICKX.Item("SHIP_ADDR_TYPE") & ""
                Dim SHIP_ADDR_CODE As String = rowSOTPICKX.Item("SHIP_ADDR_CODE") & ""
                Dim PICK_BATCH_NO As String = rowSOTPICKX.Item("PICK_BATCH_NO") & ""
                Dim CUST_CODE As String = rowSOTPICKX.Item("CUST_CODE")

                g = g + 1
                ASCMAIN1.Progress("-", CStr(g) & ":" & CUST_CODE & "/" & ORDR_GROUP_NO)

                If InStr(ORDR_GROUP_NOs, ORDR_GROUP_NO) = 0 Then
                    ORDR_GROUP_NOs = ORDR_GROUP_NOs & ",'" & ORDR_GROUP_NO & "'"
                End If
                If InStr(PICK_BATCH_NOs, PICK_BATCH_NO) = 0 Then
                    PICK_BATCH_NOs = PICK_BATCH_NOs & ",'" & PICK_BATCH_NO & "'"
                    Fill_Records("SOTPICK0", PICK_BATCH_NO, False)

                    ASCMAIN1.sql = "Select SOTSHIP1.*, DECODE (SHIP_ADDR_TYPE,'DC',SHIP_BOL_NO,'MK') SHIP_BOL_NO_X" _
                        & " from SOTSHIP1 where PICK_BATCH_NO = '" & PICK_BATCH_NO & "'"
                    Fill_Records("SOTSHIP1", PICK_BATCH_NO, False, ASCMAIN1.sql)
                End If

                If SHIP_BOL_NO <> "" Then
                    Get_SHIP_BOL_NO(SHIP_BOL_NO, SHIP_BOL_NOs)
                Else
                    ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" _
                        & " from SOTSHIP1 " _
                        & " where SOTSHIP1.PICK_BATCH_NO = '" & PICK_BATCH_NO & "'" _
                        & "   and SOTSHIP1.SHIP_ADDR_TYPE = 'MK'" _
                        & "   and SOTSHIP1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                    ASCMAIN1.sql &= "   and SOTSHIP1.SHIP_STATUS IN ('P', 'H')"
                    For Each rowSOTSHIP1 As DataRow In ASCDATA1.GetDataTable.Rows
                        SHIP_BOL_NO = rowSOTSHIP1.Item("SHIP_BOL_NO")
                        Get_SHIP_BOL_NO(SHIP_BOL_NO, SHIP_BOL_NOs)
                    Next
                End If
            Next

            PICK_BATCH_NOs = Mid$(PICK_BATCH_NOs, 2)
            ORDR_GROUP_NOs = Mid$(ORDR_GROUP_NOs, 2)

            Customer_Data(ORDR_GROUP_NOs)
            Get_SOTORDRx(ORDR_GROUP_NOs)
            Get_SOTPICKx()
            Get_EDI_Data()

            ASCMAIN1.Progress("")
        End If
        ' get location for each style, color 
        ' get pick sequence for each location
        ' split pick ticket detals if necessary X
        ' AND WRITE TO ORACLE X
        'add last single stragler to last split X

        If ASCMAIN1.CLIENT = "RGI" Then
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO") & ""
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO ='" & PICK_NO & "'")
                    Dim STYLE_CODE As String = rowSOTPICK2.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowSOTPICK2.Item("COLOR_CODE")
                    Dim PICK_QTY As Int64 = rowSOTPICK2.Item("PICK_QTY")

                    Dim LOCATION_CODE As String = ""
                    Dim LOCATION_ROUTE_SEQ As Int32 = 0
                    TAC.SOCMAIN1.GET_STYLE_COLOR_LOCATIONS(WHSE_CODE, STYLE_CODE, COLOR_CODE, LOCATION_CODE, LOCATION_ROUTE_SEQ, PICK_QTY)

                    rowSOTPICK2.Item("LOCATION_CODE") = LOCATION_CODE
                    rowSOTPICK2.Item("LOCATION_ROUTE_SEQ") = LOCATION_ROUTE_SEQ
                Next
            Next

        End If
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("CART_SERIAL_NO") = 0
        Next
    End Sub

    Sub Get_SHIP_BOL_NO(SHIP_BOL_NO As String, ByRef SHIP_BOL_NOs As String)
        If InStr(SHIP_BOL_NOs, SHIP_BOL_NO) = 0 Then
            SHIP_BOL_NOs = SHIP_BOL_NOs & ",'" & SHIP_BOL_NO & "'"
        End If

        ASCMAIN1.sql = " Select SOTPICK1.*,SOTORDR1.ORDR_DEPT,SOTORDR1.ORDR_CUST_PO from SOTPICK1,SOTORDR1 where SOTORDR1.ORDR_NO (+) = SOTPICK1.ORDR_NO and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCMAIN1.sql &= "   and SOTPICK1.PICK_STATUS = 'P'"
        ASCMAIN1.sql = "Insert into " & SOTPICK1 & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Customer_Data(ORDR_GROUP_NOs As String)
        ASCMAIN1.sql = "Select ARTCUST1.* from ARTCUST1" _
            & " where CUST_CODE in (" _
            & " Select DISTINCT CUST_CODE from SOTORDR0" _
            & " where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & "))"
        Fill_Records("ARTCUST1", "", False, ASCMAIN1.sql)

        If ASCMAIN1.CLIENT = "VAN" Then
            ' WJZ: I BELEIVE THAT THE VAN SECTION SHOULD BE GOOD FOR ANYONE - LEAVING IT FOR VAN ONLY FOR NOW SO THAT WE CAN SEE IF IT WORKS FOR VAN

            ASCMAIN1.sql = "Select ARTCUST2.*" & vbCrLf _
                & ",ARTCUST2_DC.CUST_ADDR_CODE CUST_ADDR_DC, ARTCUST2_DC.CUST_STATE DC_STATE" & vbCrLf _
                & " from ARTCUST2,ARTCUST3,ARTCUST2 ARTCUST2_DC" & vbCrLf _
                & " where ARTCUST2.CUST_CODE in (" & vbCrLf _
                & " Select DISTINCT CUST_CODE from SOTORDR0" & vbCrLf _
                & " where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & "))" & vbCrLf _
                & "   and ARTCUST3.CUST_CODE (+) = ARTCUST2.CUST_CODE   " & vbCrLf _
                & "   and ARTCUST3.CUST_ADDR_TYPE (+) = 'MK'" & vbCrLf _
                & "   and ARTCUST3.CUST_ADDR_CODE (+) = ARTCUST2.CUST_ADDR_CODE" & vbCrLf _
                & "   and ARTCUST3.CUST_ADDR_TYPE2 (+) = 'DC'" & vbCrLf _
                & "   and ARTCUST2_DC.CUST_CODE (+) = ARTCUST3.CUST_CODE   " & vbCrLf _
                & "   and ARTCUST2_DC.CUST_ADDR_TYPE (+) = 'DC'" & vbCrLf _
                & "   and ARTCUST2_DC.CUST_ADDR_CODE (+) = ARTCUST3.CUST_ADDR_CODE2"
            Fill_Records("ARTCUST2", "", False, ASCMAIN1.sql)

        Else

            ASCMAIN1.sql = "Select ARTCUST2.*" _
                & " from ARTCUST2" _
                & " where CUST_CODE in (" _
                & " Select DISTINCT CUST_CODE from SOTORDR0" _
                & " where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & "))"
            Fill_Records("ARTCUST2", "", False, ASCMAIN1.sql)

            'Add Stores DC Info
            For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("CUST_ADDR_TYPE = 'MK'")
                ASCMAIN1.sql = "Select ARTCUST2.* from ARTCUST2,ARTCUST3" _
                    & " where ARTCUST3.CUST_CODE = :PARM1" _
                    & "   and ARTCUST3.CUST_ADDR_TYPE = 'MK'" _
                    & "   and ARTCUST3.CUST_ADDR_CODE = :PARM2" _
                    & "   and ARTCUST3.CUST_ADDR_TYPE2 = 'DC'" _
                    & "   and ARTCUST2.CUST_CODE = ARTCUST3.CUST_CODE" _
                    & "   and ARTCUST2.CUST_ADDR_TYPE = 'DC'" _
                    & "   and ARTCUST2.CUST_ADDR_CODE = ARTCUST3.CUST_ADDR_CODE2"

                Dim rowARTCUST3_DC As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {rowARTCUST2.Item("CUST_CODE"),
                                                                                                        rowARTCUST2.Item("CUST_ADDR_CODE")})
                If rowARTCUST3_DC IsNot Nothing Then
                    rowARTCUST2.Item("CUST_ADDR_DC") = rowARTCUST3_DC.Item("CUST_ADDR_CODE")
                    rowARTCUST2.Item("DC_STATE") = rowARTCUST3_DC.Item("CUST_STATE")
                End If
            Next
        End If

    End Sub

    Sub Get_SOTORDRx(ORDR_GROUP_NOs As String)
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTORDR0)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0 & " Select * from SOTORDR0 where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & ")")
        ASCMAIN1.AnalyzeTable(SOTORDR0)

        Fill_Records("SOTORDR0", "", False)
        Fill_Records("SOTORDR1", "", False)
        Fill_Records("SOTORDR2", "", False)
        Fill_Records("SOTORDR3", "", False)
        Fill_Records("SOTORDR9", "", False)
        Fill_Records("SOTORDR5", "", False)
        'Fill_Records("SOTCONF2", "", False)
        Fill_Records("ICTSTYL1", "", False)
        If ASCMAIN1.CLIENT = "VAN" Then
        Else
            Fill_Records("WHTLOCB1", "", False)
        End If

        Fill_Records("ICTSTYC1", "", False)

        If ASCMAIN1.CLIENT = "VAN" Then
            ' SKIP THIS PART - WAVE WILL TAKE CARE OF IT
        Else
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("")
                Dim LOCATION_CODES As String = ""
                Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE")
                For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'", "LOCATION_QTY DESC")
                    LOCATION_CODES &= "," & rowWHTLOCB1.Item("LOCATION_CODE")
                Next
                rowICTSTYC1.Item("LOCATION_CODES") = Mid(LOCATION_CODES, 2)
            Next
        End If
    End Sub

    Sub Get_SOTPICKx()
        Fill_Records("SOTPICK1", "", False)
        'For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
        '    rowSOTPICK1.Item("PICK_NO_BC") = BC_OCode128(Trim$(rowSOTPICK1.Item("PICK_NO")), 1, 0, 0)
        'Next

        If ASCMAIN1.CLIENT = "VAN" Then
            ' NEED TO CHANGE PICK_NO_CONS TO PICK_NO FOR NON-WALMART IN ORDER TO GET THE WALMART PT TO WORK
            ' IN THE FUTURE, PROB BEST TO HAVE A DIFFERENT CRYSTAL REPORT
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                If rowSOTORDR1.Item("CUST_CODE") = "KOHLS" Then
                    rowSOTPICK1.Item("PICK_NO_CONS") = rowSOTPICK1.Item("PICK_NO")
                End If
                '  rowSOTPICK1.Item("PICK_NO_CONS") = rowSOTPICK1.Item("PICK_NO")
            Next
        End If

        Fill_Records("SOTPICK2", "", False)
        Fill_Records("SOTCART1", "", False)
        Carton_Serialization()
        Fill_Records("SOTCART2", "", False)
        SetCartItemCode()
    End Sub

    Sub Get_EDI_Data()
        Fill_Records("EDT850T2", "", False)
        Fill_Records("EDT850T1", "", False)
    End Sub

    Sub Carton_Serialization()
        ASCMAIN1.Progress("-", "Carton Serialization")

        For Each row As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("SOTCART1"), New String() {"PICK_NO"}).Rows
            Dim PICK_NO As String = row.Item("PICK_NO")
            Dim sqlw As String = "PICK_NO = '{0}'"
            sqlw = String.Format(sqlw, PICK_NO)
            Dim CART_SEQ_MAX As Int32 = Val(dst.Tables("SOTCART1").Select(sqlw).Length)
            Dim CART_SERIAL_NO As Integer = 0
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select(sqlw, "CART_NO")
                CART_SERIAL_NO += 1
                rowSOTCART1.Item("CART_SERIAL_NO") = CART_SERIAL_NO
                rowSOTCART1.Item("CART_SEQ_MAX") = CART_SEQ_MAX
                rowSOTCART1.Item("CART_1_OF_9") = CStr(CART_SERIAL_NO) & " of " & CStr(CART_SEQ_MAX)
            Next
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTCART1"), New String() {"PICK_NO"}).Select("")
            Dim PICK_NO As String = row.Item("PICK_NO")
            Dim sqlw As String = "PICK_NO = '" & PICK_NO & "'"
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            If ASCMAIN1.CLIENT = "VAN" And (rowSOTPICK1.Item("PICK_NO_CONS") & "") <> rowSOTPICK1.Item("PICK_NO") And (rowSOTPICK1.Item("PICK_NO_CONS") & "") <> "" Then
                PICK_NO = rowSOTPICK1.Item("PICK_NO_CONS") & ""
                rowSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            End If
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = Val(dst.Tables("SOTCART1").Compute("Count(CART_NO)", sqlw) & "")
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = Val(dst.Tables("SOTCART1").Compute("Sum(CART_TOTAL_WGT_CALC)", sqlw) & "")
            If ASCMAIN1.CLIENT = "VAN" Then
                rowSOTPICK1.Item("PICK_TOTAL_QTY") = Val(dst.Tables("SOTCART1").Compute("Sum(CART_TOTAL_UNITS)", sqlw) & "") + Val(rowSOTPICK1.Item("PICK_TOTAL_QTY") & "")
            Else
                rowSOTPICK1.Item("PICK_TOTAL_QTY") = Val(dst.Tables("SOTCART1").Compute("Sum(CART_TOTAL_UNITS)", sqlw) & "")
            End If

        Next
    End Sub

    Sub Update_Record()

        BeginTrans()

        CommitTrans("")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICKX, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Select All X")
        Load_Popup_Menu(grdSOTPICK1, "BBB", "Select All", "De-Select All", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTCART1, "BBB", "Select All", "De-Select All", "Print UCC128 Labels")
        Load_Popup_Menu(grdSOTNLAB2, "BBBB", "Delete Labels", "Set Preticket", "Unset Preticket", "Make Assorted")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdSOTPICKX"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All X"), UltraWinToolbars.ButtonTool)
                    If grdSOTPICKX.ActiveCell Is Nothing OrElse
                            (grdSOTPICKX.ActiveCell.Value & "" = "" _
                             Or Not New String() {"ORDR_GROUP_NO", "CUST_CODE", "PICK_BATCH_NO"}.Contains(grdSOTPICKX.ActiveCell.Column.Key)) Then
                        tlb_btn.SharedProps.Visible = False
                        tlb_btn.Tag = ""
                    Else
                        tlb_btn.Tag = grdSOTPICKX.ActiveCell.Column.Key & " = '" & grdSOTPICKX.ActiveCell.Value & "'"
                        tlb_btn.SharedProps.Caption = "Select All " & grdSOTPICKX.ActiveCell.Column.Header.Caption & " = " & grdSOTPICKX.ActiveCell.Value
                        ' tlb_btn.SharedProps.Caption = "Select All " & grdSOTPICKX.ActiveCell.Value
                        tlb_btn.SharedProps.Visible = True
                    End If

                    'tlb_btn = DirectCast(tlb_pop.Tools("De-Release Shipment"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = chkDeRelease.Checked
                    'tlb_btn = DirectCast(tlb_pop.Tools("De-Release Selected Shipments"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = chkDeRelease.Checked

                Case "grdSOTPICK1"
                    'tlb_btn = DirectCast(tlb_pop.Tools("De-Release Pick Ticket"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = chkDeRelease.Checked

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All", "Select All X"

                If grd.Name = "grdSOTPICK1" Or grd.Name = "grdSOTCART1" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        grow.Selected = (e.Tool.Key = "Select All")
                    Next
                Else
                    If e.Tool.Key = "Select All X" Then
                        Dim sqlw As String = IIf(e.Tool.Key = "Select All X", e.Tool.Tag, "")
                        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select(sqlw)
                            rowSOTPICKX.Item("SELECTED") = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                        Next
                    Else
                        For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                            grow.Cells("SELECTED").Value = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                            grow.Update()
                        Next
                    End If
                    Display_Totals()
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Print UCC128 Labels"
                If grdSOTCART1.Selected.Rows.Count = 0 Then
                    If grdSOTCART1.ActiveRow IsNot Nothing Then grdSOTCART1.ActiveRow.Selected = True
                End If
                If grdSOTCART1.Selected.Rows.Count <> 0 Then
                    Print_UCC128_Labels()
                End If

            Case "Carton Summary"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Load_SOTPICK1(SHIP_BOL_NO, True)
                Dim RPT As String = "SORPICKS"
                Print_Report_Begin()
                Generate_Report(RPT, "Shipment Carton Summary")
                Print_Report_End()

            'Manual Labels  "Delete Labels", "Set Preticket", "Unset Preticket", "Make Assorted"
            Case "Delete Labels"
                If grd.Name = "grdSOTNLAB2" Then
                    dst.Tables("SOTNLAB2").Rows.Clear()
                End If
            Case "Set Preticket", "Unset Preticket"
                If grd.Name = "grdSOTNLAB2" Then
                    For Each row As DataRow In dst.Tables("SOTNLAB2").Select("")
                        row.Item("PRE_TICKET") = IIf(e.Tool.Key = "Set Preticket", "1", "0")
                    Next
                End If

            Case "Make Assorted"
                If grd.Name = "grdSOTNLAB2" Then
                    For Each row As DataRow In dst.Tables("SOTNLAB2").Select("")
                        row.Item("STYLE_CODE") = "ASSORTED"
                        row.Item("COLOR_CODE") = "AST"
                        row.Item("SIZE_CODE") = "AST"
                    Next
                End If

                'Case "De-Release Pick Ticket"
                '    Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Text
                '    Dim pickList As List(Of String) = New List(Of String)
                '    pickList.Add(PICK_NO)
                '    DeRelease_PickTickets(pickList)

                'Case "De-Release Selected Shipments"
                '    Dim pickList As List(Of String) = New List(Of String)
                '    'For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                '    '    pickList.Add(grow.Cells("PICK_NO").Value)
                '    'Next
                '    For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                '        Dim SHIP_BOL_NO As String = rowSOTPICKX.Item("SHIP_BOL_NO")
                '        ASCMAIN1.sql = "Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = :PARM1"
                '        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {SHIP_BOL_NO}).Select("")
                '            pickList.Add(row.Item("PICK_NO"))
                '        Next
                '    Next
                '    If pickList.Count > 0 Then DeRelease_PickTickets(pickList)

                'Case "De-Release Shipment"
                '    Dim pickList As List(Of String) = New List(Of String)
                '    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & grd.ActiveRow.Cells("SHIP_BOL_NO").Text & "'")
                '        pickList.Add(rowSOTPICK1.Item("PICK_NO"))
                '    Next
                '    DeRelease_PickTickets(pickList)
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("Load")
        End Select
    End Sub

#End Region

    Sub Print_Record()
        ' NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM   
        'Fill_Records("SOTSVIA1", SHIP_CODE)
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Dim RPT As String = "SORSHIP1" ' unneccesary if Report Name is Like Form Name
        'Generate_Report(RPT, "Shipper Invoice Report", , , , , False)
        'Print_Report_End()
    End Sub

#Region "grdSOTWHSEX"

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("Load")
    End Sub

#End Region

#Region "grdSOTPICKX"

    Private Sub grdSOTPICKX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICKX.AfterRowActivate
        Setup_SOTSHIP1()
    End Sub

    Private Sub grdSOTPICKX_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICKX.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTPICKX_ClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.ClickCellEventArgs) Handles grdSOTPICKX.ClickCell

        'If optBOL.Value = "1" Then
        '    If e.Cell.Column.Key = "QTY_ADDR_LABEL" Then
        '        If Val(e.Cell.Value & "") > 3 Then
        '            e.Cell.Value = 0
        '        Else
        '            e.Cell.Value = Val(e.Cell.Value) + 1
        '        End If
        '    End If
        '    'e.Cell.Row.Cells("SELECTED").Value = "1"
        'End If
        ' If e.Cell.Row.DataChanged Then
        e.Cell.Row.Update()
        ' End If
    End Sub

    Private Sub grdSOTPICKX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTPICKX.DoubleClickRow
        'Absx1.txtFor("ORDR_NO").Text = e.Row.Cells("ORDR_NO").Value
        'Click_Command("Load")
    End Sub

#End Region

    Sub Print_Documents()

        Dim RPT As String = "SORPICK1"


        'Print_Report_Begin()

        '  If options("OPT_PULL_STORE") Then Print_Report("SORPICK4", "Distribution")
        If grdSOTPICKX.ActiveRow.Cells("ORDR_SOURCE").Value = "K" Then
            Print_Manual_Labels()
        Else
            Print_UCC128_Labels_for_Selected_Shipments()
        End If


        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
            rowSOTPICKX.Item("SELECTED") = "0"
        Next

        'Dim PRINTER_ID As String = rowPRINTER_ID.Item(0)
        'Dim rowSOTPRNT1 As DataRow = dst.Tables("SOTPRNT1").Rows.Find(PRINTER_ID)
        'Dim PRINTER_PORT As String = rowSOTPRNT1.Item("PRINTER_PORT") & ""
        ''If PRINTER_PORT = "" Then PRINTER_PORT = "\\192.168.130.201\" & PRINTER_ID

    End Sub

    Sub Print_Report(RPT As String, Optional title As String = "")
        Generate_Report(RPT, title)
    End Sub

    Sub Refresh_SOTPICKX(CUST_CODE_x As String)

        If SOTPICKX <> "" Then
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTPICKX)
        End If

        For i As Integer = 0 To 1 ' 0 : DC, 1 : MK
            If i = 0 Then
                ASCMAIN1.sql = "Select SOTORDR0.ORDR_GROUP_NO, SOTSHIP1.PICK_BATCH_NO" & vbCrLf _
                    & ", SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_SPEC_INST" & vbCrLf _
                    & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                    & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_VIA_CODE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_BOL_NO SHIP_BOL_NO_X" & vbCrLf _
                    & ", ARTCUST1.CUST_NAME, ARTCUST1.CUST_ROUTING_INST" & vbCrLf _
                    & ", SOTSHIP1.SHIP_PICK_PRINTED, SOTORDR0.ORDR_TYPE_CODE, SOTORDR0.ORDR_SOURCE" & vbCrLf
            Else
                ASCMAIN1.sql = "Select DISTINCT SOTORDR0.ORDR_GROUP_NO, SOTSHIP1.PICK_BATCH_NO" & vbCrLf _
                    & ", SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_SPEC_INST" & vbCrLf _
                    & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                    & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_ADDR_TYPE, NULL SHIP_ADDR_CODE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_VIA_CODE" & vbCrLf _
                    & ", 'MK' SHIP_BOL_NO_X" & vbCrLf _
                    & ", ARTCUST1.CUST_NAME, ARTCUST1.CUST_ROUTING_INST" & vbCrLf _
                    & ", SOTSHIP1.SHIP_PICK_PRINTED, SOTORDR0.ORDR_TYPE_CODE, SOTORDR0.ORDR_SOURCE" & vbCrLf
            End If

            If SOTPICKX <> "" Then
                ASCMAIN1.sql &= ", 0 ORDR_QTY_PICK, 0 ORDR_AMT_PICK, 0 ORDR_CNT_PICK, 0 ORDR_CNT_CART, NULL ORDR_HIGH_PRIORITY, NULL ORDR_HIGH_PRIORITY_NOTE, '0' CCPA_NO_STATUS, '' PICK_NO_CONS, '' PICK_RELEASED"
            End If

            ASCMAIN1.sql &= " from SOTSHIP1,SOTORDR0,ARTCUST1" & vbCrLf _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"

            If ASCMAIN1.CLIENT = "RGI" Then
                'no dropship, order type code <> 'B2C'
                ASCMAIN1.sql &= " and SOTORDR0.ORDR_TYPE_CODE <> 'B2C'"
            End If

            If ASCMAIN1.CLIENT = "NYA" Then
                ' WHSE 18 FOR CA USERS
                If ASCMAIN1.USER_CODES = "CA" Then
                    ASCMAIN1.sql &= " and SOTSHIP1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")"
                End If
            End If

            If ASCMAIN1.CLIENT = "VAN" Then
                ' ONLY MULTIPO WALMART FOR NOW
                ' and KOHLS
                ASCMAIN1.sql = ASCMAIN1.sql.Replace(" from SOTSHIP1,SOTORDR0,ARTCUST1",
                                                    " from SOTSHIP1,SOTORDR0,ARTCUST1,EDT850T1")
                'ASCMAIN1.sql = ASCMAIN1.sql.Replace(" where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO",
                '                                    " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and ((SOTORDR0.CUST_CODE = 'KOHLS' and EDT850T1.EDI_DEPT_DESC = 'PACK BY STORE') or (SOTORDR0.CUST_CODE = 'WALMART' ) or (SOTORDR0.CUST_CODE in ('WALCOSTAR','WALELSAV','WALGUAT','WALHOND','WALNICAR'))) and EDT850T1.EDI_DOC_SEQ_NO(+) = SOTORDR0.EDI_DOC_SEQ_NO ")
                ASCMAIN1.sql = ASCMAIN1.sql.Replace(" where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO",
                                                    " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and EDT850T1.EDI_DOC_SEQ_NO(+) = SOTORDR0.EDI_DOC_SEQ_NO ")


            End If

            If SOTPICKX = "" Then
                ASCMAIN1.sql &= " and ROWNUM < 1"
                SOTPICKX = ASCMAIN1.Temp_Table
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add Primary Key (SHIP_BOL_NO)") ' TEST
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_QTY_PICK NUMBER (8,0)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_AMT_PICK NUMBER (13,2)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_CNT_PICK NUMBER (8,0)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_CNT_CART NUMBER (8,0)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_HIGH_PRIORITY VARCHAR2(1)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_HIGH_PRIORITY_NOTE VARCHAR2(30)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add CCPA_NO_STATUS VARCHAR2(1)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add PICK_NO_CONS VARCHAR2(10)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add PICK_RELEASED DATE")
                Exit Sub
            End If

            If CUST_CODE_x = "" Then
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS IN ('P', 'H')"
            Else
                ' FOR REVIVE-FROM-THE-DEAD SPECIALS
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'F' and SOTORDR0.CUST_CODE = '" & CUST_CODE_x & "' "
            End If

            If i = 0 Then
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_ADDR_TYPE = 'DC'"
            Else
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_ADDR_TYPE = 'MK'"
            End If

            ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is NOT NULL "

            If ASCMAIN1.CLIENT = "VAN" Then
                'For now we are only showing those customer that have been put into production.
                'ASCMAIN1.sql &= " and (SOTORDR0.CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE NVL(LABEL_TEMPLATE_CODE,'NULL') <> 'NULL') or (SOTORDR0.CUST_CODE in ('WALCOSTAR','WALELSAV','WALGUAT','WALHOND','WALNICAR')))"
            End If
            ASCDATA1.ExecuteSQL("Insert into " & SOTPICKX & " " & ASCMAIN1.sql)
        Next i

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & " Select SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY) ORDR_QTY_PICK" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY * SOTPICK2.PICK_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
            & ", Max (SOTORDR1.ORDR_HIGH_PRIORITY) ORDR_HIGH_PRIORITY" & vbCrLf _
            & ", Max (SOTORDR1.ORDR_HIGH_PRIORITY_NOTE) ORDR_HIGH_PRIORITY_NOTE " & vbCrLf _
            & ", Max (SOTPICK1.PICK_RELEASED) PICK_RELEASED " & vbCrLf _
            & " from SOTPICK2,SOTPICK1,SOTORDR1," & SOTPICKX & " SOTPICKX" & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTPICKX.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & " group by SOTPICK1.SHIP_BOL_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & SOTPICKX & " Set ORDR_QTY_PICK = R1.ORDR_QTY_PICK, ORDR_AMT_PICK = R1.ORDR_AMT_PICK, ORDR_HIGH_PRIORITY = R1.ORDR_HIGH_PRIORITY, ORDR_HIGH_PRIORITY_NOTE = R1.ORDR_HIGH_PRIORITY_NOTE, PICK_RELEASED = R1.PICK_RELEASED" & vbCrLf _
            & "    where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "   Update " & SOTPICKX & " Set ORDR_CNT_PICK = " & vbCrLf _
            & "    (Select Count (*) from SOTPICK1 where SHIP_BOL_NO = R1.SHIP_BOL_NO and SOTPICK1.PICK_STATUS = 'P')" & vbCrLf _
            & "    where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "   Update " & SOTPICKX & " Set ORDR_CNT_CART = " & vbCrLf _
            & "    (Select Count (*) from SOTCART1,SOTPICK1 where SOTPICK1.PICK_NO = SOTCART1.PICK_NO and SOTPICK1.SHIP_BOL_NO = R1.SHIP_BOL_NO and SOTPICK1.PICK_STATUS = 'P')" & vbCrLf _
            & "    where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        If (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI") Then
            ' INIT_DATE used for nrew sales order release business rules.
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & " Select DISTINCT SHIP_BOL_NO, CCPA_NO_STATUS " & vbCrLf _
                & " from SOTPICK1 " & vbCrLf _
                & " where CCPA_NO_STATUS = '1' AND INIT_DATE >= '03-AUG-2015' AND CCPA_NO_AUTH IS NULL" & vbCrLf _
                & " and SHIP_BOL_NO in (select SHIP_BOL_NO from " & SOTPICKX & "); " & vbCrLf _
                & " Begin for R1 in C1 loop " & vbCrLf _
                & "     Update " & SOTPICKX & " set CCPA_NO_STATUS = R1.CCPA_NO_STATUS WHERE SHIP_BOL_NO = R1.SHIP_BOL_NO; " & vbCrLf _
                & " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If


        If (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI") Then
            ' INIT_DATE used for nrew sales order release business rules.
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & " Select DISTINCT SHIP_BOL_NO, PICK_NO_CONS " & vbCrLf _
                & " from SOTPICK1 " & vbCrLf _
                & " where  NVL(PICK_NO_CONS ,'') IS NOT NULL" & vbCrLf _
                & " and SHIP_BOL_NO in (select SHIP_BOL_NO from " & SOTPICKX & "); " & vbCrLf _
                & " Begin for R1 in C1 loop " & vbCrLf _
                & "     Update " & SOTPICKX & " set PICK_NO_CONS = R1.PICK_NO_CONS WHERE SHIP_BOL_NO = R1.SHIP_BOL_NO; " & vbCrLf _
                & " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = ""

        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        Fill_Records("SOTPICKX")

        If ASCMAIN1.DBS_COMPANY = "VANX" Or ASCMAIN1.DBS_SERVER = "VANX" Then
            For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("")
                rowSOTPICKX.Item("OPT_PICK_TICKET") = "0"
                rowSOTPICKX.Item("OPT_PULL_STYLE") = "0"
                rowSOTPICKX.Item("OPT_MANIFEST") = "0"
                rowSOTPICKX.Item("OPT_UCC128") = "1"
                rowSOTPICKX.Item("CUST_856") = "0"
            Next
        Else
            For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("")
                Dim CUST_CODE As String = rowSOTPICKX.Item("CUST_CODE")
                rowSOTPICKX.Item("OPT_PICK_TICKET") = "1"

                If Val(rowSOTPICKX.Item("ORDR_CNT_PICK") & "") > 1 Then
                    '  rowSOTPICKX.Item("OPT_PULL_STORE") = "1"
                    rowSOTPICKX.Item("OPT_PULL_STYLE") = "1"
                Else
                    '  rowSOTPICKX.Item("OPT_PULL_STORE") = "0"
                    rowSOTPICKX.Item("OPT_PULL_STYLE") = "0"
                End If

                If Val(rowSOTPICKX.Item("ORDR_CNT_CART") & "") > 5 Then ' MAYBE SHOULD BE PARAMETERIZED
                    rowSOTPICKX.Item("OPT_MANIFEST") = "1"
                Else
                    rowSOTPICKX.Item("OPT_MANIFEST") = "0"
                End If

                Dim rowSOTUCCL1 As DataRow = dst.Tables("SOTUCCL1").Rows.Find(CUST_CODE)
                If rowSOTUCCL1 IsNot Nothing AndAlso rowSOTUCCL1.Item("UCC128_PREPRINT") & "" = "1" Then
                    rowSOTPICKX.Item("OPT_UCC128") = "1"
                Else
                    rowSOTPICKX.Item("OPT_UCC128") = "0"
                End If

                If CUST_CODEs_856.Contains(CUST_CODE) Then
                    rowSOTPICKX.Item("CUST_856") = "1"
                Else
                    rowSOTPICKX.Item("CUST_856") = "0"
                End If
            Next
        End If

        grdSOTPICKX.Text = "Shipments for " & WHSE_CODE

        Display_Totals()
        Sort_grdColumns(grdSOTPICKX, "")

        EnforceConstraints(True)

        grdSOTPICKX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Display_Totals()

    End Sub

    Function Get_Selected_ORDR_GROUP_NOs() As String
        Dim ORDR_GROUP_NO As String = ""
        Dim ORDR_GROUP_NOs As String = ""

        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
            ORDR_GROUP_NO = rowSOTPICKX.Item("ORDR_GROUP_NO")
            If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then
                ORDR_GROUP_NOs &= ",'" & ORDR_GROUP_NO & "'"
            End If
        Next
        ORDR_GROUP_NOs = Mid$(ORDR_GROUP_NOs, 2)
        Return ORDR_GROUP_NOs
    End Function

    Function GET_CUSTS_CONSOLIDATED() As String
        Dim CUST_CODE_CONS As String = ""

        Dim CUST_CODE_CONS_CNT As Decimal = 0

        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")

            If Not CUST_CODE_CONS.Contains(rowSOTPICKX.Item("CUST_CODE")) Then
                CUST_CODE_CONS &= ",'" & rowSOTPICKX.Item("CUST_CODE") & "" & "'"
                CUST_CODE_CONS_CNT = +1
            End If
        Next

        Return String.Empty
    End Function

    Private Sub grdSOTPICKX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICKX.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("SELECTED").Value & "" = "1" Then
                e.Row.Cells("SELECTED").Appearance.BackColor = Drawing.Color.Green
            Else
                e.Row.Cells("SELECTED").Appearance.BackColor = Drawing.Color.Empty
            End If
            If e.Row.Cells("ORDR_HIGH_PRIORITY").Value & "" = "1" Then
                e.Row.Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Sub Refresh_ICTWHSEX()
        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
    End Sub

    Sub Setup_SOTSHIP1()
        If grdSOTPICKX.ActiveRow Is Nothing OrElse Not grdSOTPICKX.ActiveRow.IsDataRow Then
            splSOTPICK1.Visible = False
            grdSOTPICK1.Text = ""
        Else
            splSOTPICK1.Visible = True
            Dim SHIP_BOL_NO As String = grdSOTPICKX.ActiveRow.Cells("SHIP_BOL_NO").Value

            Load_SOTPICK1(SHIP_BOL_NO, True)

            If grdSOTPICKX.ActiveRow.Cells("ORDR_SOURCE").Value = "K" Then
                'hide carton tab
                tabLabels.Tabs(UCC128TAB).Visible = False
                tabLabels.Tabs(NLAB1TAB).Visible = True
                tabLabels.Tabs(NLAB1TAB).Selected = True
            Else
                tabLabels.Tabs(UCC128TAB).Visible = True
                tabLabels.Tabs(NLAB1TAB).Visible = False
                tabLabels.Tabs(UCC128TAB).Selected = True
            End If


            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICK1.Text = "Pick Tickets for Shipment " & SHIP_BOL_NO

        End If
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        Setup_SOTPICK1()
    End Sub

    Sub Setup_SOTPICK1()
        If printingPickTickets Then Exit Sub

        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTCART1.Visible = False
            grdSOTCART1.Text = ""
        Else
            grdSOTCART1.Visible = True
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value

            Load_SOTCART1("", PICK_NO)
            grdSOTCART1.Text = "Cartons for Pick Ticket " & PICK_NO

            Sort_grdColumns(grdSOTCART1, "CART_NO")
        End If

        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTPICKL.Visible = False
        Else
            grdSOTPICKL.Visible = True
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            Fill_Records("SOTPICKL", PICK_NO)
            Sort_grdColumns(grdSOTPICKL, "PICK_LNO")
        End If

        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            tabLabels.Visible = False
        Else
            tabLabels.Visible = True
            Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
            ASCMAIN1.sql = "Select SOTORDR5.*, SOTORDR1.WHSE_CODE" & vbCrLf _
                & " from SOTORDR5,SOTORDR1" & vbCrLf _
                & " where SOTORDR5.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                & "   and SOTORDR5.CUST_ADDR_TYPE = 'ST'"
            Fill_Records("SOTORDR5", "", True, ASCMAIN1.sql)

        End If
    End Sub

    Sub Load_SOTPICK1(SHIP_BOL_NO As String, load_cartons As Boolean)

        EnforceConstraints(False)
        Dim PICK_NO_CONS = ""

        ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
            & ", SOTORDR0.CUST_CODE" & vbCrLf _
            & ", DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'DC',SOTSHIP1.SHIP_BOL_NO,'MK') SHIP_BOL_NO_X" & vbCrLf _
            & " from SOTSHIP1,SOTORDR0" & vbCrLf _
            & " where SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
            & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
        Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select T1.PICK_NO, T1.ORDR_NO, T1.PICK_FREIGHT,T1.PICK_PICKER,T1.ORDR_PICK_SEQ,T1.PICK_STATUS,  " & vbCrLf _
            & "T1.PICK_RELEASED,T1.PICK_PRINTED,T1.PICK_PACKED,T1.PICK_SHIPPED,T1.PICK_BATCH_NO,T1.SHIP_BOL_NO,T1.INV_NO, " & vbCrLf _
            & "T1.PICK_CNT_CARTONS,T1.PICK_TOTAL_WGT, t1.INIT_OPER, t1.LAST_OPER, t1.INIT_DATE, t1.LAST_DATE, t1.PICK_PRINTED_OPER, " & vbCrLf _
            & "T1.PICK_NO_REV, t1.CCPA_NO, t1.SHIP_CNTL_NO, t1.CCPA_NO_STATUS, t1.CCPA_NO_AUTH, t1.CONFIG_NO, " & vbCrLf _
            & "'" & PICK_NO_CONS & "'AS PICK_NO_CONS, SOTORDR1.ORDR_DEPT, SOTORDR1.ORDR_CUST_PO " & vbCrLf _
            & ",SOTORDR1.CUST_STORE_NO from SOTPICK1 T1, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = T1.ORDR_NO" & vbCrLf _
            & "   And T1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, nvl(SOTORDR2.CUST_SIZE_CODE,'AST') CUST_SIZE_CODE, nvl(ICTSTYL1.CARTON_PACK_QTY,0) CARTON_PACK_QTY, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & " ,ICTSTYL1.CASE_CUBE, ICTSTYL1.CASE_WEIGHT_GRS, nvl(ICTSTYL1.CARTONS_PER_UNIT, 0) CARTONS_PER_UNIT, SOTORDR2.CUST_SKU" & vbCrLf _
            & IIf(ASCMAIN1.CLIENT = "RGI", ", nvl(ICTSTYL1.STYLE_ASST_QTY,0) STYLE_ASST_QTY" & vbCrLf, "") _
            & " from SOTPICK1,SOTPICK2,SOTORDR2,ICTSTYL1" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)
        ' loop here to update locations
        dst.Tables("SOTCART1").Rows.Clear()
        dst.Tables("SOTCART2").Rows.Clear()

        ASCMAIN1.sql = "Select SOTORDR1.*, 'MK' AS MARK_FOR, 'ST' AS SHIP_TO" & vbCrLf _
           & " from SOTORDR1,SOTPICK1" & vbCrLf _
           & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
           & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTORDR1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
           & " from SOTORDR2,SOTORDR1,SOTPICK1" & vbCrLf _
           & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
           & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
           & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)

        Fill_Records("SOTNLAB2", SHIP_BOL_NO)

        EnforceConstraints(True)

        If load_cartons Then
            Load_SOTCART1(SHIP_BOL_NO, "")
        End If

    End Sub

    Sub Load_SOTCART1(SHIP_BOL_NO As String, PICK_NO As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Shipment Data")

        EnforceConstraints(False)

        ASCMAIN1.sql = sqlSOTCART1 & " from SOTCART1,SOTPICK1" _
            & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO" & vbCrLf _
            & IIf(SHIP_BOL_NO <> "",
                  "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'",
                  "   and SOTCART1.PICK_NO = '" & PICK_NO & "'")
        Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
            & " from SOTCART2,SOTCART1,SOTPICK1" & vbCrLf _
            & " where SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTCART1.PICK_NO" & vbCrLf _
            & IIf(SHIP_BOL_NO <> "",
                  "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'",
                  "   and SOTCART1.PICK_NO = '" & PICK_NO & "'")
        Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)
        SetCartItemCode()

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_UCC128_Labels()
        Try
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTCART1.Selected.Rows
                Dim CART_NO As String = grow.Cells("CART_NO").Value
                Dim cartonLabel As New TAC.CartonLabel(CART_NO)
                If ASCMAIN1.DBS_COMPANY = "VAN" Then
                    Dim LabelTemplateOverride As String = ""
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine("SELECT MIN(SOTORDR1.CUST_CODE) AS CUST_CODE")
                    SQLS.AppendLine("FROM SOTCART1, SOTPICK1, SOTORDR1")
                    SQLS.AppendLine("WHERE SOTCART1.PICK_NO = SOTPICK1.PICK_NO")
                    SQLS.AppendLine("AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO")
                    SQLS.AppendLine(String.Format("AND CART_NO = '{0}'", CART_NO))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim CUST_CODE As String = ASCDATA1.GetDataValue
                    Select Case CUST_CODE
                        'Case Is = "BURLING"
                        '    Dim iResult As MsgBoxResult
                        '    Dim iTitle As String = "Burlington"
                        '    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                        '    iMSG.AppendLine("Do You Want To Print Buk Labels")
                        '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        '    If iResult = MsgBoxResult.Yes Then
                        '        LabelTemplateOverride = "BURLING2"
                        '    End If
                    End Select
                    If LabelTemplateOverride.Length > 0 Then
                        cartonLabel.PrintLabel(, cboZebraPrinter.Text, LabelTemplateOverride)
                    Else
                        cartonLabel.PrintLabel(, cboZebraPrinter.Text)
                    End If
                Else
                    cartonLabel.PrintLabel()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub PrintCartonContentLabels()
        Try
            Dim CART_NO As String = String.Empty
            Dim PICK_NO As String = String.Empty

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                PICK_NO = rowSOTPICK1.Item("PICK_NO") & String.Empty

                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'")
                    CART_NO = rowSOTCART1.Item("CART_NO")
                    Dim cartonLabel As New TAC.CartonLabel(CART_NO, "CONTENT")
                    cartonLabel.PrintLabel()
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub Print_UCC128_Labels_for_Selected_Shipments()
        Try
            If ASCMAIN1.DBS_COMPANY = "VAN" Then
                dst.Tables("SOTCART4").Clear()
                Dim CUST_CODE_LAST As String = ""
                Dim LabelTemplateOverride As String = ""
                Dim SORT_ORDER As String = "SHIP_BOL_NO,PICK_NO"
                'They Always Need To Run Customers Seperatly.
                Dim CUST_CODE_FIRST As String = ""
                If dst.Tables("SOTPICK1").Rows.Count > 0 Then
                    Dim SQ As New System.Text.StringBuilder With {.Length = 0}
                    SQ.AppendLine("SELECT MIN(CUST_CODE) AS CUST_CODE")
                    SQ.AppendLine("FROM SOTORDR1")
                    SQ.AppendLine("WHERE ORDR_NO IN (")
                    SQ.AppendLine("  SELECT ORDR_NO")
                    SQ.AppendLine("  FROM SOTPICK1")
                    SQ.AppendLine(String.Format("  WHERE PICK_NO = '{0}'", dst.Tables("SOTPICK1").Rows.Item(0).Item("PICK_NO")))
                    SQ.AppendLine(")")
                    ASCMAIN1.sql = SQ.ToString()
                    CUST_CODE_FIRST = ASCDATA1.GetDataValue
                    If "WALMART,WALCOSTAR,WALELSAV,WALGUAT,WALHOND,WALNICAR".Contains(CUST_CODE_FIRST) Then
                        SORT_ORDER = "PICK_TOT, CUST_STORE_NO DESC"
                    End If
                    If (CUST_CODE_FIRST = "KOHLS") Then
                        SORT_ORDER = "PICK_TOT DESC, PICK_NO"
                    End If
                End If

                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("", SORT_ORDER)
                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine("SELECT MIN(CUST_CODE) AS CUST_CODE")
                    SQLS.AppendLine("FROM SOTORDR1")
                    SQLS.AppendLine("WHERE ORDR_NO IN (")
                    SQLS.AppendLine("  SELECT ORDR_NO")
                    SQLS.AppendLine("  FROM SOTPICK1")
                    SQLS.AppendLine(String.Format("  WHERE PICK_NO = '{0}'", PICK_NO))
                    SQLS.AppendLine(")")
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim CUST_CODE As String = ASCDATA1.GetDataValue
                    If CUST_CODE <> CUST_CODE_LAST Then
                        CUST_CODE_LAST = CUST_CODE
                        LabelTemplateOverride = ""
                        Select Case CUST_CODE
                            'Case Is = "BURLING"
                            '    Dim iResult As MsgBoxResult
                            '    Dim iTitle As String = "Burlington"
                            '    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                            '    iMSG.AppendLine("Do You Want To Print Buk Labels")
                            '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                            '    If iResult = MsgBoxResult.Yes Then
                            '        LabelTemplateOverride = "BURLING2"
                            '    End If
                            Case Is = "WALMART"
                                Dim row1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(New String() {rowSOTPICK1.Item("ORDR_NO")})
                                If row1.Item("CUST_DC_NO") & "" = "" Then
                                    LabelTemplateOverride = "WAL_NODC"
                                End If
                        End Select
                    End If

                    If "WALCOSTAR,WALELSAV,WALGUAT,WALHOND,WALNICAR".Contains(CUST_CODE) Then
                        Dim rrow As DataRow
                        ASCMAIN1.sql = "SELECT COUNTRY_NAME FROM TATCNTRY, ARTCUST1 where TATCNTRY.COUNTRY_CODE = ARTCUST1.CUST_COUNTRY and CUST_CODE = '" & CUST_CODE & "'"
                        Dim CUST_COUNTRY As String = ASCDATA1.GetDataValue
                        rrow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & rowSOTPICK1.Item("ORDR_NO") & "'")(0)
                        Dim ORDR_CUST_PO As String = rrow.Item("ORDR_CUST_PO")

                        For Each rowCART_NO As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'")
                            Dim CART_NO As String = rowCART_NO.Item("CART_NO")
                            For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'")
                                Dim rowSOTCART4 As DataRow = dst.Tables("SOTCART4").NewRow
                                rowSOTCART4.Item("CART_NO") = CART_NO
                                rowSOTCART4.Item("PICK_NO") = PICK_NO
                                rowSOTCART4.Item("CUST_COUNTRY") = CUST_COUNTRY
                                rowSOTCART4.Item("ORDR_CUST_PO") = ORDR_CUST_PO
                                rowSOTCART4.Item("CART_TOTAL_UNITS") = rowCART_NO.Item("CART_TOTAL_UNITS")
                                rowSOTCART4.Item("CART_1_OF_9") = rowCART_NO.Item("CART_1_OF_9")

                                Dim ORDR_NO = rowSOTCART2.Item("ORDR_NO") & ""
                                Dim ORDR_LNO = rowSOTCART2.Item("ORDR_LNO") & ""
                                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                                Dim STYLE_CODE As String = rowSOTORDR2.Item("RANGE_STYLE_CODE")
                                rowSOTCART4.Item("STYLE_CODE") = STYLE_CODE
                                'If STYLE_CODE = "NB9151A" Then
                                '    STYLE_CODE = "NB1951A"
                                'End If

                                ASCMAIN1.sql = "SELECT * FROM SOTCSTY1 WHERE CUST_CODE = '" & CUST_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "'"
                                Dim rowSOTCSTY1 As DataRow = ASCDATA1.GetDataRow
                                If IsNothing(rowSOTCSTY1) Then
                                    MsgBox("Error Missing Customer Style refference " & CUST_CODE & " / " & STYLE_CODE, MsgBoxStyle.Critical)
                                End If
                                Dim SKU As String = rowSOTCSTY1.Item("CUST_STYLE_CODE") & ""
                                Dim UPC_CODE As String = rowSOTCSTY1.Item("CUST_UPC") & ""
                                rowSOTCART4.Item("SKU") = SKU
                                rowSOTCART4.Item("UPC_CODE") = UPC_CODE

                                ASCMAIN1.sql = "SELECT ICTSTYL1.*, TATCNTRY.COUNTRY_NAME FROM ICTSTYL1, TATCNTRY WHERE TATCNTRY.COUNTRY_CODE = ICTSTYL1.COUNTRY_CODE and STYLE_CODE = '" & STYLE_CODE & "'"
                                Dim rowICTSTYL1 As DataRow = ASCDATA1.GetDataRow
                                ' kilogram conversion
                                rowSOTCART4.Item("CART_WEIGHT") = (Val(rowICTSTYL1.Item("CASE_WEIGHT_GRS") & "") * 0.453592).ToString("###0.00")
                                rowSOTCART4.Item("CASE_CBM") = rowICTSTYL1.Item("CASE_CUBE")
                                rowSOTCART4.Item("ORIGIN_COUNTRY") = rowICTSTYL1.Item("COUNTRY_NAME")
                                rowSOTCART4.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC") & " " & rowICTSTYL1.Item("STYLE_DESC2")

                                dst.Tables("SOTCART4").Rows.Add(rowSOTCART4)
                                Debug.Print(rowCART_NO.Item("CART_1_OF_9"))
                                Exit For
                            Next
                        Next


                    Else
                        ASCMAIN1.sql = "Select CART_NO from SOTCART1 where PICK_NO = '" & PICK_NO & "'"
                        For Each rowCART_NO As DataRow In ASCDATA1.GetDataTable.Select("", "CART_NO")
                            Dim CART_NO As String = rowCART_NO.Item("CART_NO")
                            Dim cartonLabel As New TAC.CartonLabel(CART_NO)
                            If LabelTemplateOverride.Length > 0 Then
                                cartonLabel.PrintLabel(, cboZebraPrinter.Text, LabelTemplateOverride)
                            Else
                                cartonLabel.PrintLabel(, cboZebraPrinter.Text)
                            End If
                        Next
                    End If
                Next
                If dst.Tables("SOTCART4").Select("").Count > 0 Then
                    Print_Report("SORCART4", "Walmart Export Carton Labels")
                End If

            Else
                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("", "SHIP_BOL_NO,PICK_NO")
                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                    ASCMAIN1.sql = "Select CART_NO from SOTCART1 where PICK_NO = '" & PICK_NO & "'"
                    For Each rowCART_NO As DataRow In ASCDATA1.GetDataTable.Select("", "CART_NO")
                        Dim CART_NO As String = rowCART_NO.Item("CART_NO")
                        Dim cartonLabel As New TAC.CartonLabel(CART_NO)
                        cartonLabel.PrintLabel()
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub Print_Address_Labels_from_Details(LABEL_TEMPLATE As String)
        Dim TOTAL_LABELS As Integer = Val(dst.Tables("SOTPICKL").Compute("SUM(LABEL_QTY)", "LABEL_QTY > 0") & "")
        Dim STARTING_LABEL As Integer = 1
        For Each rowSOTPICKL As DataRow In dst.Tables("SOTPICKL").Select("LABEL_QTY > 0")
            Dim labelQty As Int32 = Val(rowSOTPICKL.Item("LABEL_QTY") & "")
            Dim labelComment As String = rowSOTPICKL.Item("STYLE_CODE") & " - " & rowSOTPICKL.Item("STYLE_DESC")
            Print_Address_Labels(labelQty, labelComment, LABEL_TEMPLATE, TOTAL_LABELS, STARTING_LABEL)
            STARTING_LABEL += labelQty
        Next
    End Sub

    Sub Print_Address_Labels(labelQty As Integer, labelComment As String, LABEL_TEMPLATE As String, Optional TOTAL_LABELS As Integer = 0, Optional STARTING_LABEL As Integer = 0)
        Dim pickNo As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)

        Try
            Dim addressLabel As New AddressLabel(pickNo, labelComment, LABEL_TEMPLATE, rowSOTORDR5)
            addressLabel.Set1of9(STARTING_LABEL, TOTAL_LABELS)
            addressLabel.PrintLabel(labelQty)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Sets the Printer Settings
    ''' </summary>
    ''' <remarks></remarks>
    ''' 

    Function Default_Printer()
        Dim settings As New PrinterSettings
        For Each printer As String In PrinterSettings.InstalledPrinters
            settings.PrinterName = printer
            If settings.IsDefaultPrinter Then
                Return printer
            End If
        Next
        Return String.Empty
    End Function

    Private Sub txtPICK_NO_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtPICK_NO.ValueChanged
        If txtPICK_NO.TextLength = 10 Then
            Dim PICK_NO As String = txtPICK_NO.Text
            Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", PICK_NO)
            If rowSOTPICK1 IsNot Nothing Then
                Dim SHIP_BOL_NO As String = rowSOTPICK1.Item("SHIP_BOL_NO")
                Dim row As DataRow = dst.Tables("SOTPICKX").Rows.Find(SHIP_BOL_NO)
                If row IsNot Nothing Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTPICKX.Rows
                        If grow.IsDataRow AndAlso grow.Cells("SHIP_BOL_NO").Value & "" = SHIP_BOL_NO Then
                            grdSOTPICKX.ActiveRow = grow
                            For Each grow2 As UltraWinGrid.UltraGridRow In grdSOTPICK1.Rows
                                If grow2.IsDataRow AndAlso grow2.Cells("PICK_NO").Value & "" = PICK_NO Then
                                    grdSOTPICK1.ActiveRow = grow2
                                    Exit Sub
                                End If
                            Next
                            Exit Sub
                        End If
                    Next
                End If
            Else
                ' DISPLAY MESSAGE?
            End If
            txtPICK_NO.Text = ""
        End If
    End Sub

    Private Sub tabLabels_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabLabels.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

    Private Sub SetCartItemCode()

        Dim CART_NO As String = String.Empty
        Dim PICK_NO As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim ORDR_LNO As Int16 = 0
        Dim rowSOTCART2 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing

        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
            CART_NO = rowSOTCART1.Item("CART_NO")
            PICK_NO = rowSOTCART1.Item("PICK_NO") & String.Empty

            Dim numItems As Int16 = ASCDATA1.SelectDistinct(dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'"), "STYLE_CODE").Rows.Count

            rowSOTCART2 = dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'")(0)
            ORDR_NO = rowSOTCART2.Item("ORDR_NO") & String.Empty
            ORDR_LNO = Val(rowSOTCART2.Item("ORDR_LNO") & String.Empty)

            rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})

            If numItems = 1 Then
                If rowSOTORDR2 IsNot Nothing Then
                    rowSOTCART1.Item("STYLE_CODE") = rowSOTORDR2.Item("STYLE_CODE")
                    rowSOTCART1.Item("STYLE_DESC") = rowSOTORDR2.Item("STYLE_DESC")
                Else
                    rowSOTCART1.Item("STYLE_CODE") = "Could not resolve Style"
                End If
            Else
                rowSOTCART1.Item("STYLE_CODE") = "Mixed"
            End If



            ' DELETE FROM HERE --------------------------------

            If PICK_NO.Length > 0 Then
                If rowSOTORDR2 IsNot Nothing Then
                    If dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "' and STYLE_CODE = '" & rowSOTORDR2.Item("STYLE_CODE") & "'").Length > 0 Then
                        Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "' and STYLE_CODE = '" & rowSOTORDR2.Item("STYLE_CODE") & "'")(0)
                        If Val(rowSOTPICK2.Item("CARTON_PACK_QTY") & String.Empty) > 0 Then
                            rowSOTCART1.Item("NUM_CARTONS") = Val(rowSOTCART1.Item("CART_TOTAL_UNITS") & String.Empty) / Val(rowSOTPICK2.Item("CARTON_PACK_QTY") & String.Empty)
                            rowSOTCART1.Item("TOTAL_WEIGHT") = Val(rowSOTCART1.Item("NUM_CARTONS") & String.Empty) * Val(rowSOTPICK2.Item("CASE_WEIGHT_GRS") & String.Empty)
                            rowSOTCART1.Item("CASE_CUBE_FT") = Val(rowSOTCART1.Item("NUM_CARTONS") & String.Empty) * (Val(rowSOTPICK2.Item("CASE_CUBE") & String.Empty) / 1728)
                        End If
                    End If

                End If
            End If

            ' DELETE TO HERE --------------------------------

            ' ED - THIS SECTION IS HOW I THINK THE ABOVE SECTION SHOULD HAVE BEEN WRITTEN

            If ASCMAIN1.Running_in_VS Then
                If CART_NO.EndsWith("248") Or CART_NO.EndsWith("262") Then
                    'Stop
                End If
            End If


            Dim NUM_CARTONS As Decimal = 0
            Dim TOTAL_WEIGHT As Decimal = 0
            Dim CASE_CUBE_FT As Decimal = 0

            For Each rowSOTCART2 In dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'")

                ORDR_NO = rowSOTCART2.Item("ORDR_NO")
                ORDR_LNO = Val(rowSOTCART2.Item("ORDR_LNO"))
                Dim QTY_PACKED As Int64 = Val(rowSOTCART2.Item("QTY_PACKED"))

                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Select _
                    ("PICK_NO = '" & PICK_NO & "' and ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO))(0)

                'rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                'Dim CARTON_PACK_QTY As Int32 = Val(rowSOTORDR2.Item("CARTON_PACK_QTY") & "")

                Dim CARTON_PACK_QTY As Int32 = Val(rowSOTPICK2.Item("CARTON_PACK_QTY") & "")
                Dim CASE_WEIGHT_GRS As Decimal = Val(rowSOTPICK2.Item("CASE_WEIGHT_GRS") & "")
                Dim CASE_CUBE As Decimal = Val(rowSOTPICK2.Item("CASE_CUBE") & "")
                If CARTON_PACK_QTY = 0 Then CARTON_PACK_QTY = 1
                Dim CARTONS As Decimal = QTY_PACKED / CARTON_PACK_QTY
                NUM_CARTONS += CARTONS
                TOTAL_WEIGHT += CARTONS * System.Math.Round(CASE_WEIGHT_GRS + 0.51, 0)
                CASE_CUBE_FT += CARTONS * (CASE_CUBE * 1.1)
            Next

            rowSOTCART1.Item("NUM_CARTONS") = NUM_CARTONS
            rowSOTCART1.Item("TOTAL_WEIGHT") = TOTAL_WEIGHT
            rowSOTCART1.Item("CASE_CUBE_FT") = CASE_CUBE_FT / 1728
        Next

    End Sub

    Private Sub create_import_file_itemlocs()

        'Dim filename As String = "C:\dmp\ITEMSLOCS.txt"
        ' NEED TO ARCHIVE THE LOCATIONS FILE & MOVE IT OUT OF THE WAY
        Dim filename As String = "S:\WAREHOUSE\LOCATIONS\ITEMSLOCS.txt"
        Dim data As String = ""
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        Dim BIN As String = ""

        Get_PARM("ICTPARMR")


        Using sr As New System.IO.StreamReader(filename)
            data = sr.ReadToEnd
            Dim datarec() As String = Split(data, vbCrLf)

            For i As Integer = 0 To UBound(datarec)
                If datarec(i) <> "" Then
                    Dim datastr() As String = Split(datarec(i), vbTab)
                    STYLE_CODE = Replace(Replace(Replace(datastr(0), Chr(34), ""), "'", ""), " ", "")
                    COLOR_CODE = Replace(Replace(datastr(1), Chr(34), ""), " ", "")
                    BIN = Replace(Replace(datastr(2), Chr(34), ""), " ", "")
                    ASCMAIN1.sql = "UPDATE ICTSTYC1 SET STYLE_BIN = '" & BIN & "' where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
                    ASCDATA1.ExecuteSQL()
                End If
            Next
        End Using

    End Sub


    Private Sub create_import_file_itemlocs_WORK()

        'Dim filename As String = "C:\dmp\ITEMSLOCS.txt"
        ' NEED TO ARCHIVE THE LOCATIONS FILE & MOVE IT OUT OF THE WAY
        Dim filename As String = "S:\WAREHOUSE\LOCATIONS\ITEMSLOCS.txt"
        Dim data As String = ""
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        Dim STYLE_BIN As String = ""

        Dim LOCDATE As String = FormatDATE(Now) 'FormatDATE(DATETIME_STAMP.Date)
        Dim drCDATE As String = Now

        Dim HISTfilename As String = "S:\WAREHOUSE\LOCATIONS\HIST\ITEMSLOCS"

        '  Get_PARM("ICTPARMR")

        If My.Computer.FileSystem.FileExists(filename) Then

            ASCDATA1.ExecuteSQL("Truncate Table " & WHTILOCS)

            dst.Tables("WHTILOCS").Rows.Clear()

            Using sr As New System.IO.StreamReader(filename)
                data = sr.ReadToEnd
                Dim datarec() As String = Split(data, vbCrLf)

                For i As Integer = 0 To UBound(datarec)
                    If datarec(i) <> "" Then
                        Dim datastr() As String = Split(datarec(i), vbTab)

                        STYLE_CODE = Replace(Replace(Replace(datastr(0), Chr(34), ""), ") '", ""), " ", "")
                        COLOR_CODE = Mid(Replace(Replace(datastr(1), Chr(34), ""), " ", ""), 1, 4)
                        STYLE_BIN = Replace(Replace(datastr(2), Chr(34), ""), " ", "")

                        Dim rowWHTLOCS As DataRow = dst.Tables("WHTILOCS").Rows.Find({STYLE_CODE, COLOR_CODE})
                        If rowWHTLOCS Is Nothing Then
                            Dim rowWHTILOCS As DataRow = dst.Tables("WHTILOCS").NewRow
                            With rowWHTILOCS
                                .Item("STYLE_CODE") = STYLE_CODE
                                .Item("COLOR_CODE") = COLOR_CODE
                                .Item("STYLE_BIN") = STYLE_BIN
                            End With
                            If datastr(1) <> "" Then
                                dst.Tables("WHTILOCS").Rows.Add(rowWHTILOCS)
                            Else
                                Dim DANAC As String = ""
                            End If

                        Else
                        End If

                    End If
                Next
            End Using

            Update_Record_TDA("WHTILOCS", "COLOR_CODE IS NOT NULL")

            Dim DANA As String = ""

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & " SELECT * FROM " & WHTILOCS & ";" & vbCrLf _
                & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
                & " update ICTSTYC1 SET STYLE_BIN = R1.STYLE_BIN WHERE STYLE_CODE = TRIM(R1.STYLE_CODE) AND COLOR_CODE = TRIM(R1.COLOR_CODE) ;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()


            My.Computer.FileSystem.MoveFile(filename, HISTfilename & "_" & LOCDATE & ".txt", True)

            EMsg = "Location Update Complete "
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")

        Else
            EMsg = "NO LOCATIONS File found "

            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub


        End If

    End Sub

    Public Shared Function FormatDATE(ByVal PDATE As String)

        Dim datestr() As String = Split(PDATE, "/")

        If Trim(PDATE) <> "" Then
            FormatDATE = Replace(PDATE, "/", "_") ' CDate((Mid(PDATE, 6, 2) & "/" & Mid(PDATE, 9, 2) & "/" & Mid(PDATE, 1, 4)))
            FormatDATE = Mid(datestr(2), 1, 4) & "_" & datestr(1) & "_" & datestr(0)
        Else
            FormatDATE = Nothing
        End If

    End Function

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        create_import_file_itemlocs_WORK()
    End Sub

    Private Sub grdSOTPICKX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTPICKX.InitializeLayout

    End Sub

    Private Sub grdSOTPICK1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTPICK1.DoubleClickRow
        If grdSOTPICKX.ActiveRow.Cells("ORDR_SOURCE").Value = "K" And e.Row.Band.Key = "SOTPICK1_SOTPICK2" Then
            Dim label_cnt As Integer = 0
            Dim pick_qty As Integer = e.Row.Cells("PICK_QTY").Value + 0
            Dim CARTON_QTY As Integer = 0
            Dim CARTON_NO As Integer = 0
            Dim msg = ""
            Dim rowSOTNLAB2 As DataRow
            Dim SHIP_BOL_NO As String = grdSOTPICKX.ActiveRow.Cells("SHIP_BOL_NO").Value

            Dim LAST_CARTON = dst.Tables("SOTNLAB2").Compute("MAX(CART_NO)", "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
            If Not (DBNull.Value.Equals(LAST_CARTON)) Then
                CARTON_NO = Val(LAST_CARTON)
            End If

            Do While CARTON_QTY < 1
                msg = InputBox("Enter Units per box", "Carton Qty")
                'Get User Text using Form
                'Dim frmASFMSGBF As New ASFMSGBF
                'Dim Label As New System.Text.StringBuilder With {.Length = 0}
                'Label.AppendLine("Enter Units per box")
                'Dim Caption As String = "Carton Qty"
                'msg = frmASFMSGBF.Get_txtblock_from_User(Label.ToString, Caption, "", False, 0)

                If Val(msg) > 0 Then
                    CARTON_QTY = Val(msg)
                End If
            Loop

            Do While label_cnt < 1
                Dim q = Int(pick_qty / CARTON_QTY)
                If q <> pick_qty / CARTON_QTY Then
                    q = q + 1
                End If
                'Dim frmASFMSGBF As New ASFMSGBF
                'Dim Label As New System.Text.StringBuilder With {.Length = 0}
                'Label.AppendLine("Enter Labels to Create")
                'Dim Caption As String = "Labels"
                'msg = frmASFMSGBF.Get_txtblock_from_User(Label.ToString, Caption, q.ToString, False, 0)
                msg = InputBox("Enter Labels to create", "How Many Labels", q.ToString)
                If Val(msg) > 0 Then
                    label_cnt = Val(msg)
                End If
            Loop

            For r As Integer = 1 To label_cnt
                rowSOTNLAB2 = dst.Tables("SOTNLAB2").NewRow
                CARTON_NO += 1
                With rowSOTNLAB2
                    .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                    .Item("CART_NO") = CARTON_NO
                    .Item("ORDR_CUST_PO") = grdSOTPICKX.ActiveRow.Cells("ORDR_CUST_PO").Value
                    .Item("STYLE_CODE") = e.Row.Cells("STYLE_CODE").Value
                    .Item("COLOR_CODE") = e.Row.Cells("COLOR_CODE").Value
                    .Item("SIZE_CODE") = e.Row.Cells("CUST_SIZE_CODE").Value
                    .Item("STYLE_DESC") = e.Row.Cells("STYLE_DESC").Value
                    .Item("SHIP_DEPT") = e.Row.ParentRow.Cells("ORDR_DEPT").Value
                    .Item("PRE_TICKET") = "0"
                    .Item("CART_QTY") = CARTON_QTY
                    If r = label_cnt And pick_qty - (CARTON_QTY * (r - 1)) > 0 Then
                        .Item("CART_QTY") = pick_qty - (CARTON_QTY * (r - 1))
                    End If
                    .Item("CART_ID") = e.Row.Cells("CUST_SKU").Value
                End With
                dst.Tables("SOTNLAB2").Rows.Add(rowSOTNLAB2)

            Next

        End If
    End Sub

    Sub Print_Manual_Labels()

        Dim LABEL_CODE As String = "NON_EDI"
        Dim cartonLabel As New TestLabel(LABEL_CODE, "")

        'For Each rowSOTNLAB2 As DataRow In dst.Tables("SOTNLAB2").Select("")
        ' Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("").First

        For Each rowSOTNLAB2 As DataRow In dst.Tables("SOTNLAB2").Select("", "CART_NO")

            Dim labelData As New Dictionary(Of String, DataRow)
            labelData.Add("SOTORDR5", rowSOTORDR5)
            labelData.Add("SOTNLAB2", rowSOTNLAB2)
            labelData.Add("ICTWHSE1", rowICTWHSE1)

            cartonLabel.PrintTestLabel(cboZebraPrinter.Text, labelData)
        Next
        'Next

        BeginTrans()
        Update_Record_TDA("SOTNLAB2")
        CommitTrans()

        'printed = True

    End Sub

End Class

