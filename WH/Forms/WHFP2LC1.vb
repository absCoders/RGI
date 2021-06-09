Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFP2LC1

#Region "Declarations"

    Dim WHTWAVEX As String = ""

    Dim WAVE_NO As String = ""
    Dim rowWHTWAVE1 As DataRow

    Dim SHIP_BOL_NOs As String = ""
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

    Dim CUST_CODE As String
    Dim P2L_LINE_ID As String

    Dim sqlCS As String = ""

    'sqlCS = "Data Source= ABSSVR2019; Initial Catalog=LPPick; Integrated Security=SSPI"
    'sqlCS = "Data Source= SVR-VDI-NJ-PK1; Initial Catalog=LPPick; User Id= abs; Password= v4n$4L3"

    Dim WHTRPLCX As String = ""

    Dim P2L_WAVE_STATUS As String = ""

    Dim MAXZONES As Integer = 0

    Dim AppearanceGreenBack As New Infragistics.Win.Appearance
    Dim AppearanceRedBack As New Infragistics.Win.Appearance
    Dim AppearanceRed As New Infragistics.Win.Appearance
    Dim AppearanceEmpty As New Infragistics.Win.Appearance

    Dim candidate2finalize As Boolean = False

    Dim expressions As New Dictionary(Of String, Dictionary(Of String, String))
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        AppearanceEmpty.ForeColor = Color.Empty
        AppearanceRed.ForeColor = Color.Red
        AppearanceGreenBack.BackColor = Color.LightGreen
        AppearanceRedBack.BackColor = Color.Red

        InquiryMode = (ASCMAIN1.MENU_ITEM_OBJECT = "WHFP2LCI")

        P2L_WAVE_STATUS = "WHTWAVE1.P2L_WAVE_STATUS = 'P'"
        If InquiryMode Then
            P2L_WAVE_STATUS = "WHTWAVE1.P2L_WAVE_STATUS IN ('P','C')"
        End If

        Create_WorkTables()

        Get_PARM("SOTPARM1")
        Get_PARM("WHTPARM1")
        sqlCS = ROWs("WHTPARM1").Item("WH_PARM_P2L_CONN")
        If ASCMAIN1.DBS_SERVER <> "VAN" Then
            If sqlCS.Contains("SVR-VDI-NJ-PK1") Then
                MsgBox("You are NOT logged into the Production Database yet your SQL Connection string appears to be Production", MsgBoxStyle.OkOnly, "Please Call ABS")
                Stop
                ' maybe after refreshing the test database, you forgot to set the CS to a test CS
                ' UPDATE WHTPARM1 SET WH_PARM_P2L_CONN = 'Data Source= ABSSVR2019; Initial Catalog=LPPick; Integrated Security=SSPI'
            End If
        End If


        With dst
            ASCMAIN1.sql = "Select WHTWAVEX.*, WHTWAVE1.CUST_CODE, WHTWAVE1.WAVE_DATE, WHTWAVE1.WHSE_CODE, WHTWAVE1.P2L_LINE_ID" & vbCrLf _
                & $" from {WHTWAVEX} WHTWAVEX, WHTWAVE1 where WHTWAVE1.WAVE_NO = WHTWAVEX.WAVE_NO"
            Create_TDA(.Tables.Add, "WHTWAVEX", "**", 0, False)

            '  Dim SHIP_STATUS_WHERE As String = "SOTSHIP1.SHIP_STATUS = 'P'"
            Dim SHIP_STATUS_WHERE As String = "SOTPICK1.PICK_STATUS = 'P'"
            If ASCMAIN1.Running_in_VS Or InquiryMode Then
                SHIP_STATUS_WHERE = "SOTPICK1.PICK_STATUS IN ('P','F')"
            End If

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            Create_TDA(.Tables.Add, "SOTCART1", "*", 1, False, "", 1)

            ASCMAIN1.sql = "Select SOTCART2.*, WHTSCSEQ.STYLE_SEQ, ICVLUPC1.UPC_CODE UPC_CODE_1" & vbCrLf _
                & " From SOTCART2, WHTSCSEQ, ICVLUPC1" & vbCrLf _
                & " Where SOTCART2.STYLE_CODE = WHTSCSEQ.STYLE_CODE" & vbCrLf _
                & "   and SOTCART2.COLOR_CODE = WHTSCSEQ.COLOR_CODE" & vbCrLf _
                & "   and SOTCART2.STYLE_CODE = ICVLUPC1.STYLE_CODE" & vbCrLf _
                & "   and SOTCART2.COLOR_CODE = ICVLUPC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 1, False, "", 2)

            'ASCMAIN1.sql = "Select WHTWAVE3.*, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
            '    & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
            '    & ", SOTORDR0.ORDR_CNT_PICK PTS, SOTORDR0.ORDR_QTY_PICK UNITS" & vbCrLf _
            '    & " from WHTWAVE3, SOTORDR0, SOTSHIP1" & vbCrLf _
            '    & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
            '    & " and SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
            '    & " and " & SHIP_STATUS_WHERE & "" & vbCrLf _
            '    & " and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
            'Create_TDA(.Tables.Add, "WHTWAVE3", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select  WHTWAVE3.*, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
                & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTPICK1.PICK_STATUS" & vbCrLf _
                & ", COUNT(DISTINCT SOTPICK1.PICK_no) PTS, SUM(SOTPICK2.PICK_QTY) UNITS" & vbCrLf _
                & " from WHTWAVE3, SOTORDR0, SOTSHIP1,SOTPICK1, SOTPICK2" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & " and SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & " and " & SHIP_STATUS_WHERE & "" & vbCrLf _
                & " and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & " and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO " & vbCrLf _
                & " GROUP BY WHTWAVE3.WAVE_NO, WHTWAVE3.SHIP_BOL_NO, WHTWAVE3.P2L_SHIP_STATUS, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
                & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTPICK1.PICK_STATUS"

            Create_TDA(.Tables.Add, "WHTWAVE3", "**", 0, True, "V", 2)




            ' & "   and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
            With .Tables("WHTWAVE3")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns.Add("CTNS", GetType(System.Int32))
                .Columns.Add("CTNS_WIP", GetType(System.Int32))
                .Columns.Add("UNITS_WIP", GetType(System.Int32))
                .Columns.Add("CTNS_PICK", GetType(System.Int32))
                .Columns.Add("UNITS_PICK", GetType(System.Int32))
                .Columns.Add("CTNS_CANC", GetType(System.Int32))
                .Columns.Add("CTNS_PALLETIZED", GetType(System.Int32))
                .Columns.Add("UNITS_CANC", GetType(System.Int32))
                ' cancel carton functionality to come
                For I As Integer = 1 To MAXZONES
                    .Columns.Add("ZONE_" & CStr(Format(I, "00")), GetType(System.Int32))
                Next
            End With
            Dim ZONEV As String = ""
            Dim ZONEDC As String = ""

            For i As Integer = 1 To MAXZONES
                ZONEV = ZONEV & ", SUM(ZONE" & CStr(Format(i, "00")) & ") ZONE_" & CStr(Format(i, "00"))
                ZONEDC = ZONEDC & ",SUM(DECODE(LOCATION_ZONE,'" & Format(i, "00") & "',SOTCART2.QTY_PACKED,0)) ZONE_" & CStr(Format(i, "00")) & vbCrLf
            Next
            ASCMAIN1.sql = "Select  WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & ZONEDC & vbCrLf _
                & ",SUM(SOTCART2.QTY_REL) TOTAL_UNITS" & vbCrLf _
                & " From WHTWAVE3, SOTCART2, SOTCART1, SOTPICK1, WHTSCSEQ, WHTLOCM1" & vbCrLf _
                & " Where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & " And SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & " And WHTSCSEQ.STYLE_CODE = SOTCART2.STYLE_CODE" & vbCrLf _
                & " And WHTSCSEQ.COLOR_CODE = SOTCART2.COLOR_CODE" & vbCrLf _
                & " And WHTSCSEQ.CUST_CODE = :PARM2" & vbCrLf _
                & " And WHTLOCM1.LOCATION_ROUTE_SEQ = WHTSCSEQ.STYLE_SEQ" & vbCrLf _
                & " And WHTLOCM1.LOCATION_CODE Like :PARM3" & vbCrLf _
                & " And WHTLOCM1.WHSE_CODE = :PARM4" & vbCrLf _
                & " And SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & " And SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & " GROUP BY WHTWAVE3.SHIP_BOL_NO"
            Create_TDA(.Tables.Add, "WHTWAVET", "**", 0, False, "VVVV", 1)


            ASCMAIN1.sql = "Select WHTWAVE3.SHIP_BOL_NO, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                & ", SOTCART1.CART_NO, SOTCART1.CART_PACKER, SOTCART1.CART_PACKED, SOTCART1.PICK_NO, SOTCART1.PALLET_NO" & vbCrLf _
                & ", SOTCART1.CART_TOTAL_UNITS, SOTCART1.CART_TOTAL_UNITS_REL" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTORDR1, SOTCART1" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and " & SHIP_STATUS_WHERE & "" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
            Create_TDA(.Tables.Add, "WHTWAVEC", "**", 0, False, "V", 3)
            With .Tables("WHTWAVEC").Columns
                .Add("CART_TOTAL_UNITS_PCK", GetType(System.Int32), "IIF(ISNULL(CART_PACKER,'')<>'',CART_TOTAL_UNITS,0)")
                .Add("CART_TOTAL_UNITS_CXL", GetType(System.Int32), "IIF(ISNULL(CART_PACKER,'')<>'',ISNULL(CART_TOTAL_UNITS_REL,0)-ISNULL(CART_TOTAL_UNITS_PCK,0),0)")
            End With


            ASCMAIN1.sql = "Select WHTWAVE3.SHIP_BOL_NO, SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE" & vbCrLf _
                & ", Sum (SOTCART2.QTY_PACKED) QTY_PACKED" & vbCrLf _
                & ", Sum (SOTCART2.QTY_REL) QTY_REL" & vbCrLf _
                & ", Sum (DECODE(SOTCART1.CART_PACKER,NULL,0,NVL(SOTCART2.QTY_PACKED,0))) QTY_PCK" & vbCrLf _
                & ", Sum (DECODE(SOTCART1.CART_PACKER,NULL,0,NVL(SOTCART2.QTY_REL,0)-NVL(SOTCART2.QTY_PACKED,0))) QTY_CXL" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTCART1, SOTCART2" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and " & SHIP_STATUS_WHERE & "" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & " group by WHTWAVE3.SHIP_BOL_NO,SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTWAVEZ", "**", 0, False, "V", 3)

            Create_Relation("WHTWAVE3", "WHTWAVEZ", "SHIP_BOL_NO")
            With .Tables("WHTWAVEZ").Columns
                .Add("SELECTED", GetType(System.String), "PARENT(WHTWAVE3_WHTWAVEZ).SELECTED")
                .Add("P2L_SHIP_STATUS", GetType(System.String), "PARENT(WHTWAVE3_WHTWAVEZ).P2L_SHIP_STATUS")
                .Add("QTY_2BI", GetType(System.Int32), "IIF(SELECTED='1' and P2L_SHIP_STATUS = 'O',QTY_PACKED,0)")
                .Add("QTY_2BD", GetType(System.Int32), "IIF(SELECTED='0' and P2L_SHIP_STATUS = 'P',QTY_PACKED,0)")
            End With

            ASCMAIN1.sql = "Select SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE" & vbCrLf _
                & ", Sum (QTY_PACKED) QTY_PACKED" & vbCrLf _
                & ", Sum (DECODE(WHTWAVE3.P2L_SHIP_STATUS,'P', QTY_REL,0)) QTY_P2L_P" & vbCrLf _
                & ", Sum (DECODE(WHTWAVE3.P2L_SHIP_STATUS,'O', QTY_REL,0)) QTY_P2L_O" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTCART1, SOTCART2" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and " & SHIP_STATUS_WHERE & "" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & " group by SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTWAVES", "**", 0, False, "V", 2)

            Create_Relation("WHTWAVES", "WHTWAVEZ", "STYLE_CODE,COLOR_CODE")
            With .Tables("WHTWAVES").Columns
                .Add("QTY_2BI", GetType(System.Int32), "SUM(CHILD(WHTWAVES_WHTWAVEZ).QTY_2BI)")
                .Add("QTY_2BD", GetType(System.Int32), "SUM(CHILD(WHTWAVES_WHTWAVEZ).QTY_2BD)")
                .Add("QTY_REL", GetType(System.Int32), "SUM(CHILD(WHTWAVES_WHTWAVEZ).QTY_REL)")
                .Add("QTY_PCK", GetType(System.Int32), "SUM(CHILD(WHTWAVES_WHTWAVEZ).QTY_PCK)")
                .Add("QTY_CXL", GetType(System.Int32), "SUM(CHILD(WHTWAVES_WHTWAVEZ).QTY_CXL)")
                .Add("QTY_ON_HAND", GetType(System.Int32))
                .Add("QTY_ON_HAND_OTHER", GetType(System.Int32))
                .Add("QTY_WO_PICK", GetType(System.Int32))
                .Add("QTY_COMM", GetType(System.Int32))
                .Add("QTY_AVA", GetType(System.Int32), "ISNULL(QTY_ON_HAND,0)+ISNULL(QTY_WO_PICK,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_2BI,0)+ISNULL(QTY_2BD,0)")
                .Add("QTY_WO_OPEN", GetType(System.Int32))
                .Add("QTY_NET", GetType(System.Int32), "ISNULL(QTY_AVA,0)+ISNULL(QTY_WO_OPEN,0)")
            End With

            ASCMAIN1.sql = "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
                & ", Sum (LOCATION_QTY) LOCATION_QTY" & vbCrLf _
                & ", Sum (LOCATION_QTY_WAVE) LOCATION_QTY_WAVE" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTCART1, SOTCART2, WHTLOCB1, WHTWAVE1" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and WHTWAVE1.WAVE_NO = WHTWAVE3.WAVE_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and " & SHIP_STATUS_WHERE & "" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & "   and WHTLOCB1.WHSE_CODE = WHTWAVE1.WHSE_CODE" & vbCrLf _
                & "   and WHTLOCB1.LOCATION_CODE = WHTWAVE1.LOCATION_CODE_DEPOSIT" & vbCrLf _
                & "   and WHTLOCB1.STYLE_CODE = SOTCART2.STYLE_CODE" & vbCrLf _
                & "   and WHTLOCB1.COLOR_CODE = SOTCART2.COLOR_CODE" & vbCrLf _
                & " group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTWAVEQ", "**", 0, False, "V", 3)
            '& "   and WHTLOCB1.BAR_CODE = '0000000000'" & vbCrLf _

            ASCMAIN1.sql = "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
                & ", Sum (LOCATION_QTY) LOCATION_QTY_OTHER" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTCART1, SOTCART2, WHTLOCB1, WHTWAVE1, WHTLOCM1" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and WHTWAVE1.WAVE_NO = WHTWAVE3.WAVE_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and " & SHIP_STATUS_WHERE & "" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & "   and WHTLOCB1.WHSE_CODE = WHTWAVE1.WHSE_CODE" & vbCrLf _
                & "   and NVL(WHTLOCM1.LOCATION_USE,'A') = 'A'" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = WHTWAVE1.WHSE_CODE" & vbCrLf _
                & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & "   and WHTLOCB1.STYLE_CODE = SOTCART2.STYLE_CODE" & vbCrLf _
                & "   and WHTLOCB1.COLOR_CODE = SOTCART2.COLOR_CODE" & vbCrLf _
                & "  group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTWAVEV", "**", 0, False, "V", 3)
            ' NEED TO MANIPULATE LOCATION REVIEW WITH RICK


            ASCMAIN1.sql = "Select SOTCART1.CART_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_DC_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_NO, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & " from SOTCART1, SOTPICK1, SOTORDR1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
                & "   and " & SHIP_STATUS_WHERE & ""
            Create_TDA(.Tables.Add, "SOTCARTA", "**", 0, False, "V", 1)

            ' original SOTCARTB
            'ASCMAIN1.sql = "Select SOTCART2.CART_NO, SOTCART2.CART_LNO, SOTCART2.QTY_PACKED, SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE, WHTLOCM1.LOCATION_CODE" & vbCrLf _
            '    & " from SOTCART2, SOTCART1, SOTPICK1, WHTSCSEQ, WHTLOCM1, SOTORDR1" & vbCrLf _
            '    & " where SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
            '    & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            '    & "   and WHTSCSEQ.STYLE_CODE = SOTCART2.STYLE_CODE" & vbCrLf _
            '    & "   And WHTSCSEQ.COLOR_CODE = SOTCART2.COLOR_CODE" & vbCrLf _
            '    & "   And WHTSCSEQ.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            '    & "   And WHTLOCM1.LOCATION_ROUTE_SEQ = WHTSCSEQ.STYLE_SEQ" & vbCrLf _
            '    & "   And WHTLOCM1.LOCATION_CODE Like :PARM2" & vbCrLf _
            '    & "   and WHTLOCM1.WHSE_CODE = SOTORDR1.WHSE_CODE" & vbCrLf _
            '    & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            '    & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            '    & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
            'Create_TDA(.Tables.Add, "SOTCARTB", "**", 0, False, "VV", 0)


            '   new SOTCARTB
            ASCMAIN1.sql = "Select SOTCART2.CART_NO, SOTCART2.CART_LNO, SOTCART2.QTY_PACKED" & vbCrLf _
                & ", DECODE(WHTRPLCX.STYLE_CODE,NULL,SOTCART2.STYLE_CODE,WHTRPLCX.R_STYLE_CODE) STYLE_CODE" & vbCrLf _
                & ", DECODE(WHTRPLCX.COLOR_CODE,NULL,SOTCART2.COLOR_CODE,WHTRPLCX.R_COLOR_CODE) COLOR_CODE" & vbCrLf _
                & ", WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & $" from SOTCART2, SOTCART1, SOTPICK1, WHTSCSEQ, WHTLOCM1, SOTORDR1, {WHTRPLCX} WHTRPLCX" & vbCrLf _
                & " where SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
                & " and " & SHIP_STATUS_WHERE & "" & vbCrLf _
                & " And WHTSCSEQ.STYLE_CODE = DECODE(WHTRPLCX.R_STYLE_CODE,NULL,SOTCART2.STYLE_CODE,WHTRPLCX.R_STYLE_CODE)" & vbCrLf _
                & " And WHTSCSEQ.COLOR_CODE = DECODE(WHTRPLCX.R_COLOR_CODE,NULL,SOTCART2.COLOR_CODE,WHTRPLCX.R_COLOR_CODE)" & vbCrLf _
                & " And WHTSCSEQ.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & " And WHTRPLCX.STYLE_CODE(+) = SOTCART2.STYLE_CODE" & vbCrLf _
                & " And WHTRPLCX.COLOR_CODE(+) = SOTCART2.COLOR_CODE" & vbCrLf _
                & " And WHTLOCM1.LOCATION_ROUTE_SEQ = WHTSCSEQ.STYLE_SEQ" & vbCrLf _
                & " And WHTLOCM1.LOCATION_CODE Like :PARM2" & vbCrLf _
                & " And WHTLOCM1.WHSE_CODE = SOTORDR1.WHSE_CODE" & vbCrLf _
                & " And SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & " And SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & " And SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
            Create_TDA(.Tables.Add, "SOTCARTB", "**", 0, False, "VV", 0)




            ' maybe this sql needs to avoid looking at waves that have been deposited, 
            '  since the pick would already in the on hand of the Deposit location
            ASCMAIN1.sql = "Select WHTWAVE2.STYLE_CODE, WHTWAVE2.COLOR_CODE" & vbCrLf _
                & ", SUM (DECODE(WHTINST1.WAVE_INST_STATUS,'0',WHTINST2.LOCATION_QTY_WAVE,0)) OPEN" & vbCrLf _
                & ", SUM (DECODE(WHTINST1.WAVE_INST_STATUS,'1',WHTINST2.LOCATION_QTY_PICK,0)) PICK" & vbCrLf _
                & " from WHTINST2,WHTINST1,WHTWAVE2,WHTWAVE1" & vbCrLf _
                & " where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
                & "   and WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                & "   and WHTWAVE2.WAVE_NO = WHTINST1.WAVE_NO AND WHTWAVE2.WAVE_LNO = WHTINST1.WAVE_LNO" & vbCrLf _
                & "   and WHTWAVE1.P2L_LINE_ID = :PARM1 and " & P2L_WAVE_STATUS & "" & vbCrLf _
                & " group by WHTWAVE2.STYLE_CODE, WHTWAVE2.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTINSTX", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select TATEVNT1.* " _
                & " from TATEVNT1 " _
                & " where TABLE_NAME = 'WHTP2LC1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, False, "V", 0)



            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.SHIP_ADDR_CODE, SOTCART1.PKG_CODE" & vbCrLf _
                & ", SUM(WHTPKGM1.INNER_CUBE) CUBE, COUNT (*) CARTONS" & vbCrLf _
                & " From SOTPICK1, SOTCART1, WHTPKGM1, SOTORDR0, SOTSHIP1" & vbCrLf _
                & " Where SOTPICK1.SHIP_BOL_NO In (" & vbCrLf _
                & " Select SHIP_BOL_NO FROM WHTWAVE3 WHERE WAVE_NO = :PARM1" & ")" & vbCrLf _
                & " And SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & " And WHTPKGM1.PKG_CODE (+) = SOTCART1.PKG_CODE" & vbCrLf _
                & " And SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " And SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & " GROUP BY SOTSHIP1.SHIP_BOL_NO, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.SHIP_ADDR_CODE, SOTCART1.PKG_CODE"
            Create_TDA(.Tables.Add, "WHTWAVEY", "**", 0, False, "V", 0)

            Create_TDA(.Tables.Add, "WHTRPLCW", "*", 0, False)

        End With

        grdWHTWAVEX.DataSource = dst.Tables("WHTWAVEX")
        grdWHTWAVE3.DataSource = dst.Tables("WHTWAVE3")
        grdWHTWAVEC.DataSource = dst.Tables("WHTWAVEC")
        grdWHTWAVES.DataSource = dst.Tables("WHTWAVES")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdWHTWAVEY.DataSource = dst.Tables("WHTWAVEY")

        With grdWHTWAVEX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
            'For Each COLUMN_NAME As String In New String() {"SELECTED", "SHIP_BOL_NO", "ORDR_GROUP_NO", "PICK_BATCH_NO", "CUST_CODE", "CUST_CODE", "ORDR_CUST_PO"}
            '    .Columns(COLUMN_NAME).Header.Fixed = True
            'Next

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If GCOL.Key = "SHIP_STYLES" Then
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                ElseIf New String() {"WAVE_NO", "ORDR_GROUP_NO", "WAVE_DATE", "CUST_CODE", "WHSE_CODE", "ORDR_TYPE_CODE", "P2L_LINE_ID"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                Else
                    If GCOL.Key.EndsWith("_2BI") Then
                        GCOL.Header.Appearance.BackColor2 = Color.Violet
                    ElseIf GCOL.Key.EndsWith("_2BP") Then
                        GCOL.Header.Appearance.BackColor2 = Color.Gold
                    Else
                        GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                    End If
                End If
            Next
        End With

        With grdWHTWAVE3.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit


                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = Activation.AllowEdit
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                ElseIf New String() {"SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO", "SHIP_ADDR_CODE", "ORDR_GROUP_NO", "ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf New String() {"CTNS_WIP", "UNITS_WIP"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Gold
                ElseIf New String() {"UNITS_PICK", "CTNS_PICK", "CTNS_PALLETIZED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.SeaGreen
                ElseIf New String() {"UNITS_CANC", "CTNS_CANC"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightPink
                ElseIf GCOL.Key.StartsWith("ZONE") Then
                    GCOL.Width = 70
                    GCOL.Header.Caption = Replace(GCOL.Key, "ZONE_", "Zone")
                    GCOL.Format = "#,##0"
                    GCOL.Header.Appearance.BackColor2 = Color.AliceBlue
                ElseIf GCOL.Key = "PTS" Then
                    GCOL.Format = "##,##0"
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf GCOL.Key = "UNITS" Then
                    GCOL.Format = "####,##0"
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                End If

            Next

            'If InquiryMode Then
            '    .Columns("SELECTED").Hidden = True
            'End If
        End With

        With grdWHTWAVES.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit



                If GCOL.Key = "QTY_2BI" Or GCOL.Key = "QTY_2BD" Then
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                ElseIf New String() {"STYLE_CODE", "COLOR_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf New String() {"QTY_PACKED", "QTY_P2L_P", "QTY_P2L_O", "QTY_REL"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Violet
                ElseIf New String() {"QTY_PCK", "QTY_CXL"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Gold
                Else
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                End If
            Next
        End With

        With grdWHTWAVEC.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit

                If GCOL.Key = "CART_TOTAL_UNITS_PCK" Or GCOL.Key = "CART_TOTAL_UNITS_CXL" Then
                    GCOL.Header.Appearance.BackColor2 = Color.Gold
                End If
            Next
        End With

        With grdWHTWAVEY.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit


                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = Activation.AllowEdit
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                ElseIf New String() {"SHIP_BOL_NO", "ORDR_CUST_PO", "SHIP_ADDR_CODE", "PKG_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf New String() {"CUBE"}.Contains(GCOL.Key) Then
                    GCOL.Width = 80
                    GCOL.Format = "##0.0000"
                    GCOL.Header.Appearance.BackColor2 = Color.SeaGreen
                ElseIf New String() {"CARTONS"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.SeaGreen
                    GCOL.Width = 80
                    GCOL.Format = "###,##0"
                    GCOL.Header.Appearance.BackColor2 = Color.SeaGreen
                End If

            Next

            'If InquiryMode Then
            '    .Columns("SELECTED").Hidden = True
            'End If
        End With




        Create_Summary(grdWHTWAVEX, "WAVE_NO", "Count")
        Create_Summary(grdWHTWAVEX, New String() {"SHIP_CNT", "SHIP_CNT_2BI", "SHIP_CTNS", "SHIP_CTNS_2BI", "SHIP_UNITS", "SHIP_UNITS_2BI", "SHIP_CNT_2BP", "SHIP_CTNS_2BP", "SHIP_UNITS_2BP"})

        Create_Summary(grdWHTWAVE3, "SHIP_BOL_NO", "Count")

        Create_Summary(grdWHTWAVE3, New String() {"SELECTED", "PTS", "UNITS", "CTNS", "CTNS_WIP", "UNITS_WIP", "UNITS_PICK", "CTNS_PICK", "UNITS_CANC", "CTNS_CANC", "CTNS_PALLETIZED"})
        For i As Integer = 1 To MAXZONES
            Create_Summary(grdWHTWAVE3, "ZONE_" & CStr(Format(i, "00")))
        Next

        Create_Summary(grdWHTWAVEC, "CART_NO", "Count")
        Create_Summary(grdWHTWAVEC, New String() {"CART_TOTAL_UNITS_REL", "CART_TOTAL_UNITS", "CART_TOTAL_UNITS_PCK", "CART_TOTAL_UNITS_CXL"})

        Create_Summary(grdWHTWAVES, "STYLE_CODE", "Count")
        Create_Summary(grdWHTWAVES, New String() {"QTY_PACKED", "QTY_P2L_P", "QTY_P2L_O", "QTY_2BI", "QTY_2BD", "QTY_REL", "QTY_PCK", "QTY_CXL", "QTY_ON_HAND", "QTY_ON_HAND_OTHER", "QTY_WO_PICK", "QTY_COMM", "QTY_AVA", "QTY_WO_OPEN", "QTY_NET"})

        Show_Filter(grdWHTWAVEX, True)

        Create_Summary(grdWHTWAVEY, New String() {"CUBE", "CARTONS"})


        'ASCMAIN1.Add_Value_List(grdWHTWAVE3, "ORDR_SOURCE", Nothing, New String() {":", "K:Keyboard", "W:Web", "E:EDI"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("WAVE_NO").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Wave"
                Else
                    WAVE_NO = Absx1.txtFor("WAVE_NO").Text
                    rowWHTWAVE1 = LookUp("WHTWAVE1", WAVE_NO)
                    If rowWHTWAVE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Wave"
                    Else
                        If rowWHTWAVE1.Item("P2L_WAVE_STATUS") <> "P" And Not InquiryMode Then
                            EMsg &= vbCrLf & $"Wave {WAVE_NO} is not Pending P2L Induction"
                        End If

                        If Not InquiryMode Then
                            ASCMAIN1.sql = "Select Count(1) from whtINST1" & vbCrLf _
                           & " Where WAVE_INST_STATUS = '1'" & vbCrLf _
                           & " and WAVE_NO = :PARM1"
                            If Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", WAVE_NO) & "") <> 0 Then
                                MsgBox("This Wave has Picks that have not yet been Deposited.", MsgBoxStyle.OkOnly, "Warning")
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then

                    If Not InquiryMode Then
                        If Not ASCMAIN1.Logical_Lock("WHTWAVE1", WAVE_NO) Then Exit Sub
                    End If

                End If

            Case "Refresh"

            Case "Update"

                ' VERIFY THAT ALL CARTONS IN A TO-BE-DE-INDUCTED SHIPMENT ARE DELETABLE
                Dim CART_NOs_Not_Deletable As String = ""

                ASCMAIN1.sql = $"SELECT R_STYLE_CODE, R_COLOR_CODE FROM {WHTRPLCX} WHTRPLCX" & vbCrLf _
                & " MINUS" & vbCrLf _
                & $" SELECT STYLE_CODE, COLOR_CODE FROM WHTSCSEQ WHERE CUST_CODE = '{rowWHTWAVE1.Item("CUST_CODE")}'"
                For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                    EMsg &= vbCr & $"Missing Location and SEQ for {row.Item("R_STYLE_CODE")} - {row.Item("R_COLOR_CODE")}"
                Next

                If EMsg = "" Then
                    Using sqlConn As New System.Data.SqlClient.SqlConnection(sqlCS)
                        sqlConn.Open()
                        ' Dim CART_NOs As String = ""
                        For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("SELECTED = '0' and P2L_SHIP_STATUS = 'P'")
                            Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
                            Fill_Records("SOTCARTA", SHIP_BOL_NO)
                            For Each rowSOTCARTA As DataRow In dst.Tables("SOTCARTA").Select("", "CART_NO")
                                Dim CART_NO As String = rowSOTCARTA("CART_NO")
                                'CART_NOs &= $",'{CART_NO}"
                                Dim sql As String = $"SELECT [PickOrderStatus], [PickOrderState], [HasBeenInducted] FROM [LPPick].[dbo].[PickOrders] where [PickOrderNumber] = '{CART_NO}'"
                                Dim sqlCmd As New System.Data.SqlClient.SqlCommand(sql, sqlConn)
                                Using dr As System.Data.SqlClient.SqlDataReader = sqlCmd.ExecuteReader()
                                    Do While dr.Read
                                        '[PickOrderStatus], [PickOrderState], [HasBeenInducted]
                                        Dim PickOrderStatus As Integer = Val(dr("PickOrderStatus") & "")
                                        Dim PickOrderState As Integer = Val(dr("PickOrderState") & "")
                                        Dim HasBeenInducted As Integer = Val(dr("HasBeenInducted") & "")
                                        If PickOrderStatus <> 0 Or PickOrderState = 5 Or HasBeenInducted <> 0 Then
                                            CART_NOs_Not_Deletable &= $",'{CART_NO}"
                                        End If
                                    Loop
                                End Using
                                If CART_NOs_Not_Deletable <> "" Then
                                    If EMsg.Length > 100 Then
                                        EMsg &= vbCr & "."
                                    Else
                                        EMsg &= vbCr & $"Cannot  De-Induct Shipment {SHIP_BOL_NO} from P2L - Cartons have been picked"
                                    End If

                                End If
                            Next
                        Next
                        sqlConn.Close()
                    End Using
                End If

            Case "Finalize"
                ASCMAIN1.sql = "Select Count(1) from WHTINST1" & vbCrLf _
                & " Where WAVE_INST_STATUS = '1'" & vbCrLf _
                & "   and WAVE_NO = :PARM1"
                If Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", WAVE_NO) & "") <> 0 Then
                    EMsg &= vbCr & "Cannot Finalize a Wave which has Picks that have not yet been Deposited."
                End If
                Dim CTNS_WIP As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"CART_PACKER Is NULL") & "")
                If CTNS_WIP <> 0 Then
                    EMsg &= vbCr & "Cannot Finalize a Wave which has Cartons with Open Status."
                End If

                If EMsg = "" Then
                    If MsgBox($"Are you sure that you want To Finalize Wave {WAVE_NO}?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
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

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Refresh"
                If ScreenMode And InquiryMode Then
                    Load_Record()
                Else
                    Refresh_WHTWAVEX()
                End If

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Finalize"
                Finalize_Wave()
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Load").Settings.Enabled = not_iScreenMode


                If InquiryMode And Not candidate2finalize Then
                    .Items("Refresh").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                End If

                .Items("Done").Settings.Enabled = iScreenMode

                .Items("FInalize").Settings.Enabled = iScreenMode
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode

                .Items("FInalize").Visible = Not InquiryMode And candidate2finalize
                .Items("Update").Visible = Not InquiryMode And Not candidate2finalize
                .Items("Cancel").Visible = Not InquiryMode
                .Items("Done").Visible = InquiryMode


            End With
        End With

        If ScreenMode Then
            grdWHTWAVEY.Parent = contPackage_Breakdown2
        Else
            grdWHTWAVEY.Parent = SplitContainer1.Panel2
        End If


        UltraTabControl2.Visible = Not ScreenMode
        splMain.Visible = ScreenMode

        lblCUST_CODE.Visible = ScreenMode
        txtCUST_CODE.Visible = ScreenMode
        lblP2L_LINE_ID.Visible = ScreenMode
        txtP2L_LINE_ID.Visible = ScreenMode
        lblWAVE_DATE.Visible = ScreenMode
        dteWAVE_DATE.Visible = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            Dim CTNS_WIP As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"CART_PACKER Is NULL") & "")
            If InquiryMode Or CTNS_WIP = 0 Then
                tabWHTWAVEX.SelectedTab = tabWHTWAVEX.Tabs("Already Inducted")
            End If

            If candidate2finalize Then
                Dim CTNS_CANC As Int32 = Val(dst.Tables("WHTWAVE3").Compute("SUM(CTNS_CANC)", "") & "")
                Dim UNITS_CANC As Int32 = Val(dst.Tables("WHTWAVE3").Compute("SUM(UNITS_CANC)", "") & "")

                MsgBox($"Total Cartons Cancelled = {CTNS_CANC}" & vbCrLf & $"Total Units Cancelled = {UNITS_CANC}", MsgBoxStyle.OkOnly, "Please Note the following:")
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        'For Each TABLE_NAME As String In New String() _
        '    {"SOTORDR1", "SOTORDR2"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next
        EnforceConstraints(True)

        candidate2finalize = False

        WAVE_NO = ""
        Absx1.txtFor("WAVE_NO").Text = ""

        WHSE_CODE = ""
        Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""

        Refresh_WHTWAVEX()
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        WHSE_CODE = rowWHTWAVE1.Item("WHSE_CODE")
        txtWHSE_CODE.Text = WHSE_CODE
        'Refresh_SOTSHIPX()

        rowICTWHSE1 = ASCDATA1.GetDataRow($"Select * from ICTWHSE1 where WHSE_CODE = '{WHSE_CODE}'")


        EnforceConstraints(False)

        CUST_CODE = rowWHTWAVE1.Item("CUST_CODE")
        P2L_LINE_ID = rowWHTWAVE1.Item("P2L_LINE_ID")

        rowWHTWAVE1 = LookUp("WHTWAVE1", WAVE_NO)
        txtCUST_CODE.Text = CUST_CODE
        txtP2L_LINE_ID.Text = P2L_LINE_ID
        dteWAVE_DATE.Value = rowWHTWAVE1.Item("WAVE_DATE")

        Manage_Expressions("WHTWAVE3", True)
        Manage_Expressions("WHTWAVES", True)
        Manage_Expressions("WHTWAVEZ", True)

        Fill_Records("WHTWAVEC", WAVE_NO)
        Fill_Records("WHTWAVE3", WAVE_NO)

        Fill_Records("TATEVNT1", WAVE_NO)
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)





        Fill_Records("WHTWAVEY", WAVE_NO)

        grdWHTWAVEY.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdWHTWAVEY.DisplayLayout.Bands(0).SortedColumns.Add("SHIP_ADDR_CODE", False, True)
        grdWHTWAVEY.Text = $"Package Breakdown in Wave {WAVE_NO}"

        Fill_Records("WHTWAVET", New String() {WAVE_NO, CUST_CODE, P2L_LINE_ID & "%", WHSE_CODE})

        ASCMAIN1.sql = $"Truncate Table {WHTRPLCX}"
        ASCDATA1.ExecuteSQL()

        Dim SQLWHTRPLCX As String = "Select WHTRPLCW.STYLE_CODE, WHTRPLCW.COLOR_CODE, WHTRPLCW.R_STYLE_CODE, WHTRPLCW.R_COLOR_CODE" & vbCrLf _
            & " from WHTRPLCW, WHTSCSEQ" & vbCrLf _
            & " where WHTRPLCW.R_STYLE_CODE = WHTSCSEQ.STYLE_CODE" & vbCrLf _
            & " And WHTRPLCW.R_COLOR_CODE = WHTSCSEQ.COLOR_CODE" & vbCrLf _
            & " And WHTRPLCW.WAVE_NO = '" & WAVE_NO & "'" & vbCrLf _
            & " UNION " & vbCrLf _
            & " Select STYLE_CODE,COLOR_CODE,R_STYLE_CODE,R_COLOR_CODE FROM WHTRPLC1"
        'Dim WHTRPLCX As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = $"Insert into {WHTRPLCX} " & SQLWHTRPLCX
        ASCDATA1.ExecuteSQL()

        Sort_grdColumns(grdWHTWAVE3, "SHIP_BOL_NO")
        For Each row As DataRow In dst.Tables("WHTWAVE3").Select("P2L_SHIP_STATUS in ('P','O')")
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            If row.Item("P2L_SHIP_STATUS") = "P" Then
                row.Item("SELECTED") = "1"
                Dim CTNS_WIP As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NULL") & "")
                row.Item("CTNS_WIP") = CTNS_WIP
                Dim UNITS_WIP As Int32 = Val(dst.Tables("WHTWAVEC").Compute("SUM (CART_TOTAL_UNITS)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NULL") & "")
                row.Item("UNITS_WIP") = UNITS_WIP

                Dim CTNS_PICK As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NOT NULL") & "")
                row.Item("CTNS_PICK") = CTNS_PICK
                Dim UNITS_PICK As Int32 = Val(dst.Tables("WHTWAVEC").Compute("SUM (CART_TOTAL_UNITS)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NOT NULL") & "")
                row.Item("UNITS_PICK") = UNITS_PICK
                Dim CART_TOTAL_UNITS_REL As Int32 = Val(dst.Tables("WHTWAVEC").Compute("SUM (CART_TOTAL_UNITS_REL)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NOT NULL") & "")
                Dim CART_TOTAL_UNITS As Int32 = Val(dst.Tables("WHTWAVEC").Compute("SUM (CART_TOTAL_UNITS)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NOT NULL") & "")
                row.Item("UNITS_CANC") = CART_TOTAL_UNITS_REL - CART_TOTAL_UNITS

                Dim CTNS_CANC As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NOT NULL and CART_TOTAL_UNITS = 0") & "")
                row.Item("CTNS_CANC") = CTNS_CANC

                Dim CTNS_PALLETIZED As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NOT NULL and PALLET_NO IS NOT NULL") & "")
                row.Item("CTNS_PALLETIZED") = CTNS_PALLETIZED

            End If
            Dim CTNS As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}'") & "")
            row.Item("CTNS") = CTNS

        Next

        dst.Tables("WHTWAVE3").AcceptChanges()

        For Each rowWHTWAVET As DataRow In dst.Tables("WHTWAVET").Select("")
            Dim SHIP_BOL_NO As String = rowWHTWAVET.Item("SHIP_BOL_NO")
            Dim rowWHTWAVE3 As DataRow = dst.Tables("WHTWAVE3").Rows.Find(New String() {WAVE_NO, SHIP_BOL_NO})
            If Not rowWHTWAVE3 Is Nothing Then
                For I As Integer = 1 To MAXZONES
                    rowWHTWAVE3.Item("ZONE_" & CStr(Format(I, "00"))) = Val(rowWHTWAVET.Item("ZONE_" & CStr(Format(I, "00"))) & "")
                    ' rowWHTWAVE3.Item("ZONE_" & CStr(Format(I, "00"))) = Val(rowWHTWAVE7.Item("Zone_" & CStr(Format(I, "00"))) & "")
                Next
            End If
        Next

        Dim ZONESCOUNT As Integer = 0
        ASCMAIN1.sql = "Select max(LOCATION_ZONE) from WHTLOCM1 WHERE LOCATION_CODE LIKE '" & P2L_LINE_ID & "%'"
        If Val(ASCDATA1.GetDataValue) <> 0 Then
            ZONESCOUNT = Val(ASCDATA1.GetDataValue)
        End If

        For I As Integer = 1 To MAXZONES
            If I > ZONESCOUNT Then
                grdWHTWAVE3.DisplayLayout.Bands(0).Columns("Zone_" & CStr(Format(I, "00"))).Hidden = True
            Else
                grdWHTWAVE3.DisplayLayout.Bands(0).Columns("Zone_" & CStr(Format(I, "00"))).Hidden = False
            End If
        Next

        Fill_Records("WHTWAVES", WAVE_NO)
        Sort_grdColumns(grdWHTWAVES, "STYLE_CODE, COLOR_CODE")

        Manage_Expressions("WHTWAVE3", False)
        Manage_Expressions("WHTWAVES", False)
        Manage_Expressions("WHTWAVEZ", False)

        Fill_Records("WHTWAVEZ", WAVE_NO)

        Fill_Records("WHTWAVEQ", WAVE_NO)
        For Each rowWHTWAVEQ As DataRow In dst.Tables("WHTWAVEQ").Select("")
            Dim STYLE_CODE As String = rowWHTWAVEQ.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTWAVEQ.Item("COLOR_CODE")
            Dim LOCATION_QTY As Int32 = Val(rowWHTWAVEQ.Item("LOCATION_QTY") & "")
            Dim LOCATION_QTY_WAVE As Int32 = Val(rowWHTWAVEQ.Item("LOCATION_QTY_WAVE") & "")
            Dim rowWHTWAVES As DataRow = dst.Tables("WHTWAVES").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            rowWHTWAVES.Item("QTY_ON_HAND") = LOCATION_QTY
            rowWHTWAVES.Item("QTY_COMM") = LOCATION_QTY_WAVE
        Next

        'Fill_Records("WHTWAVEV", WAVE_NO)
        For Each rowWHTWAVEV As DataRow In dst.Tables("WHTWAVEV").Select("")
            Dim STYLE_CODE As String = rowWHTWAVEV.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTWAVEV.Item("COLOR_CODE")
            Dim LOCATION_QTY_OTHER As Int32 = Val(rowWHTWAVEV.Item("LOCATION_QTY_OTHER") & "")
            Dim rowWHTWAVES As DataRow = dst.Tables("WHTWAVES").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            rowWHTWAVES.Item("QTY_ON_HAND_OTHER") = LOCATION_QTY_OTHER
        Next

        Fill_Records("WHTINSTX", P2L_LINE_ID)
        For Each rowWHTINSTX As DataRow In dst.Tables("WHTINSTX").Select("")
            Dim STYLE_CODE As String = rowWHTINSTX.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTINSTX.Item("COLOR_CODE")
            Dim OPEN As Int32 = Val(rowWHTINSTX.Item("OPEN") & "")
            Dim PICK As Int32 = Val(rowWHTINSTX.Item("PICK") & "")
            Dim rowWHTWAVES As DataRow = dst.Tables("WHTWAVES").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            'rowWHTWAVES.Item("QTY_ON_HAND") = Val(rowWHTWAVES.Item("QTY_ON_HAND") & "") + PICK
            If rowWHTWAVES IsNot Nothing Then
                rowWHTWAVES.Item("QTY_WO_PICK") = PICK
                rowWHTWAVES.Item("QTY_WO_OPEN") = OPEN
            End If
        Next

        EnforceConstraints(True)

        tabWHTWAVEX.SelectedTab = tabWHTWAVEX.Tabs("To Be Inducted")
        Setup_tabWHTWAVEX()


        If Not InquiryMode Then

            Dim CTNS_WIP As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"CART_PACKER IS NULL") & "")
            Dim SHPS_WIP As Int32 = Val(dst.Tables("WHTWAVE3").Compute("COUNT(SHIP_BOL_NO)", $"ISNULL(P2L_SHIP_STATUS,'?') <> 'P'") & "")

            If CTNS_WIP = 0 And SHPS_WIP = 0 Then
                If MsgBox("This Wave appears to be completely picked." _
                          & vbCrLf & vbCrLf & "Are you looking to Finalize this Wave?",
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Option to Finalize") = MsgBoxResult.Yes Then
                    candidate2finalize = True
                End If

            End If
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        dst.Tables("WHTWAVE3").AcceptChanges()

        For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("SELECTED = '1' and P2L_SHIP_STATUS = 'O'")
            Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
            rowWHTWAVE3.Item("P2L_SHIP_STATUS") = "P"
            Create_P2L_xml(rowWHTWAVE3)

            TAC.TACMAIN1.Record_Event("WHTP2LC1", WAVE_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "IND", "Induction", SHIP_BOL_NO)
        Next

        For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("SELECTED = '0' and P2L_SHIP_STATUS = 'P'")
            Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
            rowWHTWAVE3.Item("P2L_SHIP_STATUS") = "O"
            Create_P2L_Delete_xml(rowWHTWAVE3)

            TAC.TACMAIN1.Record_Event("WHTP2LC1", WAVE_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "REV", "Reverse Induction", SHIP_BOL_NO)
        Next

        Update_Record_TDA("WHTWAVE3")

        CommitTrans("")
    End Sub

    Sub Finalize_Wave()

        BeginTrans()

        ASCMAIN1.sql = $"Update WHTWAVE3 Set P2L_SHIP_STATUS = 'C' where WAVE_NO = '{WAVE_NO}'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Update WHTWAVE1 Set P2L_WAVE_STATUS = 'C', WAVE_STATUS = 'F' where WAVE_NO = '{WAVE_NO}'"
        ASCDATA1.ExecuteSQL()

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        rowWHTMOVE1.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
        rowWHTMOVE1.Item("WHSE_TRAN_TYPE") = "F"
        rowWHTMOVE1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
        rowWHTMOVE1.Item("WHSE_CODE") = WHSE_CODE
        rowWHTMOVE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTMOVE1.Item("INIT_DATE") = DATETIME_STAMP
        rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowWHTMOVE1.Item("LAST_DATE") = DATETIME_STAMP
        rowWHTMOVE1.Item("STATUS") = "U"
        dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)

        'We need to move only picked items to the ship location - Carton and Carton Details
        'We need to adjust for wave_substitutions in kohls
        'Bar_code is default between both locations

        Dim BAR_CODE As String = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")
        Dim WHSE_TRAN_LNO As Integer
        'WHTWAVES

        For Each rowWHTWAVES As DataRow In dst.Tables("WHTWAVES").Select("")
            Dim LOCATION_QTY_PICK As Int64 = Val(rowWHTWAVES.Item("QTY_PACKED") & "")
            If LOCATION_QTY_PICK <> 0 Then

                Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
                With rowWHTMOVE2
                    .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                    WHSE_TRAN_LNO += 1
                    .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                    .Item("LOCATION_CODE_FROM") = rowWHTWAVE1.Item("LOCATION_CODE_DEPOSIT")
                    .Item("LOCATION_CODE_TO") = rowICTWHSE1.Item("WHSE_LOC_SHP")
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("WHSE_TRAN_QTY") = LOCATION_QTY_PICK
                    ' .Item("WHSE_TRAN_QTY_ORIG") = LOCATION_QTY_PICK
                    .Item("STYLE_CODE") = rowWHTWAVES.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = rowWHTWAVES.Item("COLOR_CODE")
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("STATUS") = "U"
                    .Item("LOAD_NO_FROM") = rowWHTWAVE1.Item("LOAD_NO_DEPOSIT")
                    .Item("LOAD_NO_TO") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                End With
                dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
            End If
        Next

        Update_Record_TDA("WHTMOVE1")
        Update_Record_TDA("WHTMOVE2")

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                                New Object() {WHSE_TRAN_NO, 0, 1},
                                New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})

        'Inventory Adjustments for substitutions  - use WHTRPLCX
        ASCMAIN1.sql = $"Select * from {WHTRPLCX} WHTRPLCX"
        Fill_Records("WHTRPLCW", "", True, ASCMAIN1.sql)

        If dst.Tables("WHTRPLCW").Select("").Length <> 0 Then
            Dim ORDR_CUST_POs As String = ""
            For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("")
                ORDR_CUST_POs &= ";" & rowWHTWAVE3.Item("ORDR_CUST_PO")
            Next
            ORDR_CUST_POs = Mid(ORDR_CUST_POs, 2)

            Dim ADJ_NO As String = Add_ICTIADJ1("Subs " & Mid(CUST_CODE & ":" & ORDR_CUST_POs, 1, 200), ROWs("WHTPARM1").Item("WH_PARM_CS_PICK"))

            Dim ADJ_LNO As Int64 = 0
            'WHTWAVES is summarized by Style & color
            For Each rowWHTRPLC1 As DataRow In dst.Tables("WHTRPLC1").Select("")
                Dim STYLE_CODE As String = rowWHTRPLC1.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowWHTRPLC1.Item("COLOR_CODE")

                For Each rowWHTWAVES As DataRow In dst.Tables("WHTWAVES").Select($"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'")
                    Dim LOCATION_QTY_PICK As Int64 = Val(rowWHTWAVES.Item("QTY_PACKED") & "")
                    If LOCATION_QTY_PICK <> 0 Then
                        Dim ADJ_QTY As Int64 = Val(rowWHTWAVES.Item("QTY_PACKED") & "")
                        Dim ADJ_REF As String = "SUB FOR"
                        For r As Integer = 0 To 1
                            Add_ICTIADJ2(ADJ_NO, ADJ_LNO, STYLE_CODE, COLOR_CODE, ADJ_QTY, ADJ_REF)

                            If r = 0 Then
                                ADJ_QTY = -1 * ADJ_QTY
                                STYLE_CODE = rowWHTRPLC1.Item("R_STYLE_CODE")
                                COLOR_CODE = rowWHTRPLC1.Item("R_COLOR_CODE")
                                ADJ_REF = "SUB WITH"
                            End If
                        Next
                    End If
                Next
            Next

            Update_Record_TDA("ICTIADJ1")
            Update_Record_TDA("ICTIADJ2")
            ICCMAIN1.Shuttle_ADJ_to_ICTTRAN1_SQL(ADJ_NO)

            ASCDATA1.ExecuteSP("ICPIADJI", "VN", New Object() {ADJ_NO, 1}, New String() {"ADJ_NO_in", "S"})
            ASCDATA1.ExecuteSP("ICPIADJG", "V", New Object() {ADJ_NO}, New String() {"ADJ_NO_in"})

            ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                            New Object() {"A", ADJ_NO, ASCMAIN1.SESSION_NO},
                            New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
        End If

        CommitTrans("Update Complete")
        TAC.TACMAIN1.Record_Event("WHTP2LC1", WAVE_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "FIN", "Finalized", "")

        'CommitTrans($"Wave {WAVE_NO} has been Finalized")
    End Sub
    Function Add_ICTIADJ1(ADJ_NOTE As String, REASON_CODE As String) As String

        Dim ADJ_NO As String = ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
        Else
            ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        End If

        Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow
        With rowICTIADJ1
            .Item("ADJ_NO") = ADJ_NO
            .Item("ADJ_DATE") = DATETIME_STAMP.Date
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("REASON_CODE") = REASON_CODE
            .Item("ADJ_NOTE") = ADJ_NOTE
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("REGISTER_IND") = "0"
            .Item("ADJ_SOURCE") = "W"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("TOTAL_COSTS") = 0
            .Item("ADJ_REF") = WAVE_NO
        End With
        dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

        Return ADJ_NO
    End Function

    Sub Add_ICTIADJ2(ADJ_NO As String, ByRef ADJ_LNO As Integer, STYLE_CODE As String, COLOR_CODE As String, ADJ_QTY As Int64, ADJ_REF As String)
        Dim rowICTIADJ2 As DataRow = dst.Tables("ICTIADJ2").NewRow
        With rowICTIADJ2
            .Item("ADJ_NO") = ADJ_NO
            ADJ_LNO += 1
            .Item("ADJ_LNO") = ADJ_LNO
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("ADJ_QTY") = ADJ_QTY
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            .Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
            .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
            .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("LOCATION_CODE") = rowICTWHSE1.Item("WHSE_LOC_SHP")
            .Item("BAR_CODE") = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")
            .Item("ADJ_REF") = ADJ_REF
        End With
        dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTWAVE3, "SSBB", "Show Filter", "Show GroupBox", "Select All", "De-Select All", "HeatMap")
        Load_Popup_Menu(grdWHTWAVES, "SSB", "Show Filter", "Show GroupBox", "Style Status Inquiry")
        Load_Popup_Menu(grdWHTWAVEC, "SSB", "Show Filter", "Show GroupBox", "Cancel Carton", "Carton Contents")
        Load_Popup_Menu(grdWHTWAVEY, "SS", "Show Filter", "Show GroupBox")
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
                Case "grdWHTWAVE3"
                    'If InquiryMode Then
                    '    tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                    '    tlb_btn.SharedProps.Visible = False
                    '    tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                    '    tlb_btn.SharedProps.Visible = False
                    'End If
                Case "grdWHTWAVEC"
                    If InquiryMode Or candidate2finalize Then
                        tlb_btn = DirectCast(tlb_pop.Tools("Cancel Carton"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn = DirectCast(tlb_pop.Tools("Cancel Carton"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = (tabWHTWAVEX.SelectedTab.Key = "Already Inducted")
                    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                If grd.Name = "grdWHTWAVE3" Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Executing: " & e.Tool.Key)
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        Dim SHIP_BOL_NO As String = grow.Cells("SHIP_BOL_NO").Value
                        ASCMAIN1.Progress("-", SHIP_BOL_NO)
                        grow.Cells("SELECTED").Value = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                        grow.Update()
                    Next
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                End If

            Case "Carton Contents"
                Dim CART_NO As String = grdWHTWAVEC.ActiveRow.Cells("CART_NO").Value
                If grd.Name = "grdWHTWAVEC" Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Executing: " & e.Tool.Key)

                    ASCMAIN1.Progress("-", CART_NO)

                    Dim rowSOTCART1 As DataRow = Fill_Record("SOTCART1", CART_NO)
                    If rowSOTCART1 Is Nothing Then
                        MsgBox("Cannot Locate Carton Record", MsgBoxStyle.OkOnly, $"Cannot Cancel Carton {CART_NO}")
                        Exit Sub
                    End If
                    If rowSOTCART1.Item("CART_PACKER") & "" = "" Then
                        MsgBox($"Carton {CART_NO} is still open, not available Audit", MsgBoxStyle.OkOnly, $"Carton Open in P2L")
                        Exit Sub
                    End If

                    Fill_Records("SOTCART2", CART_NO)
                    Print_Report_Begin()

                    CR_params.Add("SUBT", "Carton Contents")
                    Generate_Report("WHRP2LC1")
                    Print_Report_End()
                End If

            Case "Cancel Carton"
                Dim CART_NO As String = grdWHTWAVEC.ActiveRow.Cells("CART_NO").Value

                If MsgBox($"Do you really want to Cancel (ie, zero out picks for) Carton {CART_NO}", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                If grd.Name = "grdWHTWAVEC" Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Executing: " & e.Tool.Key)

                    ASCMAIN1.Progress("-", CART_NO)


                    Dim rowSOTCART1 As DataRow = Fill_Record("SOTCART1", CART_NO)
                    If rowSOTCART1 Is Nothing Then
                        MsgBox("Cannot Locate Carton Record", MsgBoxStyle.OkOnly, $"Cannot Cancel Carton {CART_NO}")
                        Exit Sub
                    End If
                    If rowSOTCART1.Item("CART_PACKER") & "" <> "" Then
                        MsgBox($"Carton {CART_NO} is not available for Cancellation", MsgBoxStyle.OkOnly, $"Cannot Cancel Carton {CART_NO}")
                        Exit Sub
                    End If

                    rowSOTCART1.Item("CART_PACKER") = "P2L"
                    rowSOTCART1.Item("CART_PACKED") = Now + ASCMAIN1.NowTSD
                    grdWHTWAVEC.ActiveRow.Cells("CART_PACKER").Value = rowSOTCART1.Item("CART_PACKER")
                    grdWHTWAVEC.ActiveRow.Cells("CART_PACKED").Value = rowSOTCART1.Item("CART_PACKED")

                    Dim PICK_NO As String = rowSOTCART1.Item("PICK_NO")

                    Dim CART_TOTAL_UNITS As Int64 = 0

                    Fill_Records("SOTCART2", CART_NO)
                    Fill_Records("SOTPICK2", PICK_NO)

                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("")
                        rowSOTCART2.Item("QTY_PACKED") = 0
                    Next

                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("")

                        Dim STYLE_CODE As String = rowSOTCART2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTCART2.Item("COLOR_CODE")
                        Dim CART_LNO As Int32 = Val(rowSOTCART2.Item("CART_LNO") & "")

                        Dim PICK_LNO As String = Val(rowSOTCART2.Item("ORDR_LNO") & "")
                        Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Rows.Find(New Object() {PICK_NO, PICK_LNO})
                        ' NOTE THAT THE LINE ABOVE ASSUMES THAT PICK_LNO = ORDR_LNO

                        rowSOTPICK2.Item("PICK_QTY_CANC") = rowSOTPICK2.Item("PICK_QTY_CONF")
                        rowSOTPICK2.Item("PICK_QTY_CONF") = 0
                    Next

                    rowSOTCART1.Item("CART_TOTAL_UNITS") = 0

                    grdWHTWAVEC.ActiveRow.Cells("CART_TOTAL_UNITS_REL").Value = rowSOTCART1.Item("CART_TOTAL_UNITS_REL")
                    grdWHTWAVEC.ActiveRow.Cells("CART_TOTAL_UNITS").Value = rowSOTCART1.Item("CART_TOTAL_UNITS")

                    Try
                        BeginTrans()

                        Dim SHIP_BOL_NO As String = grdWHTWAVE3.ActiveRow.Cells("SHIP_BOL_NO").Value
                        Dim rowWHTWAVE3 As DataRow = dst.Tables("WHTWAVE3").Rows.Find(New String() {WAVE_NO, SHIP_BOL_NO})
                        Create_P2L_Delete_xml(rowWHTWAVE3, CART_NO)

                        Update_Record_TDA("SOTCART1")
                        Update_Record_TDA("SOTCART2")

                        Update_Record_TDA("SOTPICK2")

                        TAC.TACMAIN1.Record_Event("WHTP2LC1", WAVE_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "CXL", "Canceled Carton", CART_NO)

                        CommitTrans()

                        grdWHTWAVEC.ActiveRow.Update()

                    Catch ex As Exception
                        Rollback()
                        grdWHTWAVEC.ActiveRow.CancelUpdate()
                    End Try

                    MsgBox($"Carton {CART_NO} Cancelled", MsgBoxStyle.OkOnly, "Success")
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")

            Case "HeatMap"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Value
                Dim SHIP_ADDR_CODE As String = grd.ActiveRow.Cells("SHIP_ADDR_CODE").Value

                Dim P2L_LINE_ID As String = rowWHTWAVE1.Item("P2L_LINE_ID")

                Fill_Records("SOTCARTB", New String() {SHIP_BOL_NO, P2L_LINE_ID & "%"})

                Dim QTYs As New Dictionary(Of String, Integer)
                ASCMAIN1.sql = $"Select LOCATION_CODE from WHTLOCM1 where LOCATION_CODE like '{P2L_LINE_ID}%'"
                For Each row As DataRow In ASCDATA1.GetDataTable().Select("", "LOCATION_CODE")
                    Dim LOCATION_CODE As String = row.Item("LOCATION_CODE")
                    QTYs.Add(LOCATION_CODE, 0)
                Next

                For Each row As DataRow In dst.Tables("SOTCARTB").Select("")
                    Dim LOCATION_CODE As String = row.Item("LOCATION_CODE")
                    Dim QTY_PACKED As String = Val(row.Item("QTY_PACKED") & "")
                    QTYs(LOCATION_CODE) += QTY_PACKED
                Next

                Dim FILENAME As String = ASCMAIN1.Folders("Temp") & SHIP_BOL_NO & ".csv"
                Using sw As New System.IO.StreamWriter(FILENAME)
                    Dim x As String = "LOCATION_CODE,LINE,BAY,LEVEL,LANE,QTY"
                    sw.WriteLine(x)

                    'F1-01-A-1,F1,1,A,1,27
                    For Each LOCATION_CODE As String In QTYs.Keys
                        Dim QTY As Integer = QTYs(LOCATION_CODE)
                        Dim LINE As String = Split(LOCATION_CODE, "-")(0)
                        Dim BAY As String = Split(LOCATION_CODE, "-")(1)
                        Dim LEVEL As String = Split(LOCATION_CODE, "-")(2)
                        Dim LANE As String = Split(LOCATION_CODE, "-")(3)
                        x = $"{LOCATION_CODE},{LINE},{BAY},{LEVEL},{LANE},{QTY}"
                        sw.WriteLine(x)
                    Next
                End Using
                'http://127.0.0.1:5500/?SHIP_BOL_NO=101&ORDR_CUST_PO=102A&SHIP_ADDR_CODE=6060D
                'System.Diagnostics.Process.Start("https://www.absolution.com/webnavigator/index.html")
                ' System.Diagnostics.Process.Start($"https://www.absolution.com/webnavigator/?SHIP_BOL_NO={SHIP_BOL_NO}&ORDR_CUST_PO={ORDR_CUST_PO}&SHIP_ADDR_CODE={SHIP_ADDR_CODE}")
                System.Diagnostics.Process.Start($"https://api.vandale.com/mystaticfiles/index.html?SHIP_BOL_NO={SHIP_BOL_NO}&ORDR_CUST_PO={ORDR_CUST_PO}&SHIP_ADDR_CODE={SHIP_ADDR_CODE}")
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

    Sub Refresh_WHTWAVEX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Waves")

        Create_WorkTables()
        Fill_Records("WHTWAVEX")
        Sort_grdColumns(grdWHTWAVEX, "WAVE_NO".ToLower)

        'Fill_Records("WHTWAVEY")
        'Sort_grdColumns(grdWHTWAVEY, "WAVE_NO".ToLower)


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Create_WorkTables()

        Dim sqlWHTWAVEX As String = "Select WHTWAVE1.WAVE_NO from WHTWAVE1 where " & P2L_WAVE_STATUS

        If WHTWAVEX = "" Then
            WHTWAVEX = ASCMAIN1.Temp_Table(sqlWHTWAVEX)
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CNT NUMBER (3,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CTNS NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_UNITS NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CNT_2BI NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CTNS_2BI NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_UNITS_2BI NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CNT_2BP NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CTNS_2BP NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_UNITS_2BP NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_STYLES NUMBER (7,0)")

            ASCMAIN1.sql = "Select WHTWAVE2.STYLE_CODE, WHTWAVE2.COLOR_CODE, STYLE_CODE_SUB R_STYLE_CODE, COLOR_CODE_SUB R_COLOR_CODE" & vbCrLf _
            & " from WHTWAVE2, WHTSCSEQ WHERE ROWNUM <2" & vbCrLf
            WHTRPLCX = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select max(LOCATION_ZONE) from WHTLOCM1"
            If Val(ASCDATA1.GetDataValue) <> 0 Then
                MAXZONES = Val(ASCDATA1.GetDataValue)
            End If

        Else
            ASCMAIN1.sql = $"Truncate Table {WHTWAVEX}"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Insert into {WHTWAVEX} (WAVE_NO) " & sqlWHTWAVEX
            ASCDATA1.ExecuteSQL()

            Dim sqlC As String = " where Current of C1"
            ASCMAIN1.sql = "" _
                & $"Begin" & vbCrLf _
                & $" Declare Cursor C1 is Select * from {WHTWAVEX} for Update;" & vbCrLf _
                & $" Begin" & vbCrLf _
                & $"  For R1 in C1 Loop" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CNT       = (Select Count(*) from WHTWAVE3 where WAVE_NO = R1.WAVE_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CNT_2BI   = (Select Count(*) from WHTWAVE3 where WAVE_NO = R1.WAVE_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CNT_2BP   = (Select Count(*) from WHTWAVE3 where WAVE_NO = R1.WAVE_NO and WHTWAVE3.P2L_SHIP_STATUS = 'P') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS      = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS_2BI  = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS_2BP  = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and WHTWAVE3.P2L_SHIP_STATUS = 'P' and SOTCART1.CART_PACKER IS NULL) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS     = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS_2BI = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS_2BP = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO and WHTWAVE3.P2L_SHIP_STATUS = 'P' and SOTCART1.CART_PACKER IS NULL) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_STYLES    = (Select Count (Distinct SOTCART2.STYLE_CODE) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO) {sqlC};" & vbCrLf _
                & $"  End Loop;" & vbCrLf _
                & $" End;" & vbCrLf _
                & $"End;"
            ASCDATA1.ExecuteSQL()
        End If

    End Sub

    Private Sub grdWHTWAVEX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdWHTWAVEX.DoubleClickRow
        If e.Row.IsFilterRow Or Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Dim WAVE_NO As String = e.Row.Cells("WAVE_NO").Value
        Absx1.txtFor("WAVE_NO").Text = WAVE_NO
        Click_Command("Load")
    End Sub

    Private Sub tabWHTWAVEX_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabWHTWAVEX.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabWHTWAVEX()
    End Sub

    Sub Setup_tabWHTWAVEX()

        Dim dvw As DataView = DirectCast(grdWHTWAVE3.DataSource, DataTable).DefaultView

        If tabWHTWAVEX.SelectedTab.Key = "To Be Inducted" Then
            splWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            grdWHTWAVE3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            dvw.RowFilter = "P2L_SHIP_STATUS = 'O'"
            grdWHTWAVE3.Text = "Shipments to be Inducted"

        ElseIf tabWHTWAVEX.SelectedTab.Key = "Already Inducted" Then
            splWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            grdWHTWAVE3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            dvw.RowFilter = "P2L_SHIP_STATUS = 'P'"
            If InquiryMode Then
                dvw.RowFilter = "P2L_SHIP_STATUS = 'P' OR P2L_SHIP_STATUS = 'C'"
            End If
            grdWHTWAVE3.Text = "Shipments already Inducted"
        End If

        Setup_WHTWAVEC()

    End Sub

    Private Sub grdWHTWAVES_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTWAVES.InitializeRow

        Dim QTY_AVA As Int32 = Val(e.Row.Cells("QTY_AVA").Value & "")
        If QTY_AVA < 0 Then
            e.Row.Cells("QTY_AVA").Appearance = AppearanceRed
        Else
            e.Row.Cells("QTY_AVA").Appearance = AppearanceEmpty
        End If
        If e.Row.Cells("QTY_AVA").Value & "" <> "" Then
            e.Row.Cells("QTY_AVA").ToolTipText = "On Hand + Picked - Qty Sel + Qty Del"
        End If

        Dim QTY_NET As Int32 = Val(e.Row.Cells("QTY_NET").Value & "")
        If QTY_NET < 0 Then
            e.Row.Cells("QTY_NET").Appearance = AppearanceRed
        Else
            e.Row.Cells("QTY_NET").Appearance = AppearanceEmpty
        End If
    End Sub

    Private Sub Create_P2L_xml(rowWHTWAVE3 As DataRow)

        Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
        Dim QTY_PACKED_WHTWAVEZ As Int32 = Val(dst.Tables("WHTWAVEZ").Compute("SUM(QTY_PACKED)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}'") & "")

        Dim xmlString As New System.Text.StringBuilder

        Dim P2L_LINE_ID As String = rowWHTWAVE1.Item("P2L_LINE_ID")

        Fill_Records("SOTCARTA", SHIP_BOL_NO)
        Fill_Records("SOTCARTB", New String() {SHIP_BOL_NO, P2L_LINE_ID & "%"})

        Dim QTY_PACKED_SOTCART2 As Int32 = Val(dst.Tables("SOTCARTB").Compute("SUM(QTY_PACKED)", "") & "")
        If QTY_PACKED_SOTCART2 <> QTY_PACKED_WHTWAVEZ Then
            Throw New Exception("Qty about to send to P2L does not agree with Shipment Qty Released")
        End If

        xmlString.AppendLine("<LPXML>")
        For Each rowSOTCARTA As DataRow In dst.Tables("SOTCARTA").Select("", "CART_NO")

            Dim CART_NO As String = rowSOTCARTA("CART_NO")
            xmlString.AppendLine($"<PickOrder PickOrderNumber='{CART_NO}'>")

            Dim ORDR_CUST_PO As String = rowSOTCARTA("ORDR_CUST_PO")
            Dim CUST_DC_NO As String = rowSOTCARTA("CUST_DC_NO")
            Dim CUST_STORE_NO As String = rowSOTCARTA("CUST_STORE_NO")
            Dim ORDR_NO As String = rowSOTCARTA("ORDR_NO")
            Dim PICK_NO As String = rowSOTCARTA("PICK_NO")
            Dim ORDR_GROUP_NO As String = rowSOTCARTA("ORDR_GROUP_NO")
            xmlString.AppendLine($"<PickOrderXtra ORDR_CUST_PO='{ORDR_CUST_PO}' CUST_DC_NO='{CUST_DC_NO}' CUST_STORE_NO='{CUST_STORE_NO}' ORDR_NO='{ORDR_NO}' PICK_NO='{PICK_NO}' ORDR_GROUP_NO ='{ORDR_GROUP_NO}' SHIP_BOL_NO='{SHIP_BOL_NO}' WAVE_NO='{WAVE_NO}' />")

            For Each rowSOTCARTB As DataRow In dst.Tables("SOTCARTB").Select($"CART_NO = '{CART_NO}' and QTY_PACKED <> 0", "LOCATION_CODE")
                Dim LOCATION_CODE As String = rowSOTCARTB("LOCATION_CODE")
                Dim STYLE_CODE As String = rowSOTCARTB("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTCARTB("COLOR_CODE")
                Dim CART_LNO As String = rowSOTCARTB("CART_LNO")
                Dim QTY_PACKED As Int32 = Val(rowSOTCARTB("QTY_PACKED") & "")
                xmlString.AppendLine($"<PickLine LocationName='{LOCATION_CODE}' PickOrderQty='{CStr(QTY_PACKED)}'>")
                xmlString.AppendLine($"<PickLineXtra STYLE_CODE='{STYLE_CODE}' COLOR_CODE='{COLOR_CODE}' CART_LNO='{CART_LNO}'/>")
                xmlString.AppendLine("</PickLine>")
            Next

            xmlString.AppendLine("</PickOrder>")
        Next
        xmlString.AppendLine("</LPXML>")

        Dim doc As New System.Xml.XmlDocument()
        doc.LoadXml(xmlString.ToString)
        doc.Save($"{ASCMAIN1.Folders("Work")}{SHIP_BOL_NO}.xml")

        'INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES(xmlString.ToString)
        'INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES('<LPXML>…</LPXML>')

        Using sqlConn As New System.Data.SqlClient.SqlConnection(sqlCS)

            sqlConn.Open()

            'Dim sqlP As New System.Data.SqlClient.SqlParameter("@parm", SqlDbType.Xml)
            'sqlP.Value = ""
            ' Dim sql As String = "Insert into xxx values (@parm1)"
            Dim sql As String = $"INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES('{doc.InnerXml}')"
            Using sqlCmd As New System.Data.SqlClient.SqlCommand(sql, sqlConn)
                sqlCmd.ExecuteNonQuery()
            End Using
            sqlConn.Close()
        End Using

    End Sub

    Private Sub Create_P2L_Delete_xml(rowWHTWAVE3 As DataRow, Optional CART_NO_to_cancel As String = "")

        Dim xmlString As New System.Text.StringBuilder
        'Dim P2L_LINE_ID As String = rowWHTWAVE1.Item("P2L_LINE_ID")
        Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")

        xmlString.AppendLine("<LPXML>")

        If CART_NO_to_cancel <> "" Then
            xmlString.AppendLine($"<PickOrder PickOrderNumber='{CART_NO_to_cancel}' TransactionCode='Delete'>")
            xmlString.AppendLine("</PickOrder>")
        Else

            Fill_Records("SOTCARTA", SHIP_BOL_NO)

            For Each rowSOTCARTA As DataRow In dst.Tables("SOTCARTA").Select("", "CART_NO")
                Dim CART_NO As String = rowSOTCARTA("CART_NO")
                xmlString.AppendLine($"<PickOrder PickOrderNumber='{CART_NO}' TransactionCode='Delete'>")
                '<PickOrder PickOrderNumber = "00001945460097597523" TransactionCode="Delete"></PickOrder>
                xmlString.AppendLine("</PickOrder>")
            Next
        End If

        xmlString.AppendLine("</LPXML>")

        Dim doc As New System.Xml.XmlDocument()
        doc.LoadXml(xmlString.ToString)
        doc.Save($"{ASCMAIN1.Folders("Work")}{SHIP_BOL_NO & "D" & CART_NO_to_cancel}.xml")

        Using sqlConn As New System.Data.SqlClient.SqlConnection(sqlCS)
            sqlConn.Open()
            'Dim sqlP As New System.Data.SqlClient.SqlParameter("@parm", SqlDbType.Xml)
            'sqlP.Value = ""
            ' Dim sql As String = "Insert into xxx values (@parm1)"
            Dim sql As String = $"INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES('{doc.InnerXml}')"
            Using sqlCmd As New System.Data.SqlClient.SqlCommand(sql, sqlConn)
                sqlCmd.ExecuteNonQuery()
            End Using
            sqlConn.Close()
        End Using

    End Sub

    Private Sub grdWHTWAVE3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTWAVE3.AfterRowActivate
        Setup_WHTWAVEC()
    End Sub

    Sub Setup_WHTWAVEC()
        If grdWHTWAVE3.ActiveRow Is Nothing OrElse Not grdWHTWAVE3.ActiveRow.IsDataRow Then
            grdWHTWAVEC.Visible = False
        Else
            grdWHTWAVEC.Visible = True
            Dim SHIP_BOL_NO As String = grdWHTWAVE3.ActiveRow.Cells("SHIP_BOL_NO").Value
            Dim SHIP_ADDR_CODE As String = grdWHTWAVE3.ActiveRow.Cells("SHIP_ADDR_CODE").Value
            grdWHTWAVEC.Text = $"Cartons in Shipment {SHIP_BOL_NO} - DC {SHIP_ADDR_CODE}"

            Dim dvw As DataView = DirectCast(grdWHTWAVEC.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"SHIP_BOL_NO = '{SHIP_BOL_NO}'"
            Sort_grdColumns(grdWHTWAVEC, "CART_NO")
        End If
    End Sub

    Private Sub grdWHTWAVE3_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTWAVE3.InitializeRow
        If e.Row.Cells("P2L_SHIP_STATUS").Value & "" = "O" Then ' tabWHTWAVEX.SelectedTab.Key = "To Be Inducted" Then
            e.Row.Cells("SELECTED").ToolTipText = "Check this box to Induct this Shipment into P2L"
            If e.Row.Cells("SELECTED").Value & "" = "1" Then
                e.Row.Cells("SELECTED").Appearance = AppearanceGreenBack
            Else
                e.Row.Cells("SELECTED").Appearance = AppearanceEmpty
            End If
        Else
            e.Row.Cells("SELECTED").ToolTipText = "Uncheck this box to Delete this Shipment from P2L"
            If e.Row.Cells("SELECTED").Value & "" <> "1" Then
                e.Row.Cells("SELECTED").Appearance = AppearanceRedBack
            Else
                e.Row.Cells("SELECTED").Appearance = AppearanceEmpty
            End If
        End If
    End Sub

    Sub Manage_Expressions(TABLE_NAME As String, remove_expressions As Boolean)

        Dim table_expressions As New Dictionary(Of String, String)
        If Not expressions.ContainsKey(TABLE_NAME) Then
            expressions.Add(TABLE_NAME, table_expressions)
        Else
            table_expressions = expressions(TABLE_NAME)
        End If

        If remove_expressions Then
            ' Remove Expressions
            table_expressions.Clear()
            For Each dcol As DataColumn In dst.Tables(TABLE_NAME).Columns
                If dcol.Expression <> "" Then
                    table_expressions.Add(dcol.ColumnName, dcol.Expression)
                    dcol.Expression = ""
                End If
            Next
        Else
            ' Restore Expressions
            For Each COLUMN_NAME As String In table_expressions.Keys
                dst.Tables(TABLE_NAME).Columns(COLUMN_NAME).Expression = table_expressions(COLUMN_NAME)
            Next
        End If

    End Sub

    Private Sub grdWHTWAVEX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTWAVEX.AfterRowActivate
        If grdWHTWAVEX.ActiveRow.Cells("WAVE_NO").Value & "" <> "" Then
            Fill_Records("WHTWAVEY", grdWHTWAVEX.ActiveRow.Cells("WAVE_NO").Value)
            grdWHTWAVEY.DisplayLayout.Bands(0).SortedColumns.Clear()
            grdWHTWAVEY.DisplayLayout.Bands(0).SortedColumns.Add("SHIP_ADDR_CODE", False, True)
            grdWHTWAVEY.Text = $"Package Breakdown in Wave {grdWHTWAVEX.ActiveRow.Cells("WAVE_NO").Value}"
        End If


    End Sub

    Private Sub grdWHTWAVEX_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdWHTWAVEX.InitializeLayout

    End Sub
End Class