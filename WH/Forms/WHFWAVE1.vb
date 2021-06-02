Imports System.Xml
Imports Infragistics.Win.UltraWinGrid

Public Class WHFWAVE1
    'in MID NOV we will design the process that takes invty from shipping and puts it back into location

    Dim sqlSOTSHIPX As String
    Dim SOTSHIPX As String
    Dim SOTSHIPX_1 As String

    Dim SOTSHIPC As String
    Dim WAVE_LNO_ctr As Int64

    Dim sqlWHTWAVEX As String
    Dim WHTINSTX As String
    Dim WHTINSTY As String
    Dim WHTINSTZ As String

    Dim WHTLOCBW As String
    Dim sqlSHIP_BOL_NOs As String
    Dim sqlSHIP_BOL_NOs_2 As String
    'Dim TempTable_BOL_NOs As String
    Dim sql_Wave_Filter As String

    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

    Dim WAVE_NO As String = ""
    Dim rowWHTWAVE1 As DataRow
    Dim rowSOTSHIP1 As DataRow

    Dim WAVE_INST_NO_ctr As Integer = 0

    Dim CUST_CODE As String

    Dim COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE
    Dim rowICTSTYL1 As DataRow

    Dim SHIP_BOL_NO As String = ""
    Dim preview As Boolean = False

    Dim WAVE_INST_NO_void As New List(Of String)
    Dim WAVE_TYPE As String = ""
    Dim C As WHC.WHCRF000
    Dim LOCATION_CODE_preferred As String

    Dim WHTINST1_expressions As New Dictionary(Of String, String)
    Dim WHTINST2_expressions As New Dictionary(Of String, String)
    Dim WHTWAVE2_expressions As New Dictionary(Of String, String)

    Dim grdWHTLOCB1_app0 As New Infragistics.Win.Appearance
    Dim grdWHTLOCB1_app1 As New Infragistics.Win.Appearance

    Dim sqlPPK_CODEs As String = ""
    Dim loading_lead_screen As Boolean = False
    'The Walmart list below will also need to be updated in WHRWAVE1 Report
    Dim WalmartCodes As String = "'WALMART','WALMARTCOM','WALCOSTAR','WALELSAV','WALGUAT','WALHOND','WALNICAR'"
    Dim REC_LOCATIONS As String = "'00-REC-A','00-REC-B','00-REC-C'"
    Dim P2L_ALLOW As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")
        Get_PARM("WHTPARM1")
        Get_PARM("SOTPARM1")

        If MENU_ITEM_OBJECT = "WHFWAVEI" Then
            InquiryMode = True
            tabMain.Tabs("Shipments").Visible = False
        End If

        Create_Temp_Table("")

        With dst

            ASCMAIN1.sql = "Select * from " & SOTSHIPX & " where NVL(SHIP_WAVE_STATUS,'0') <> '1'"
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "", 1)
            With .Tables("SOTSHIPX")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns.Add("WAVE_QTY", GetType(System.Int64))
            End With

            With .Tables.Add("SOTSHIPX_PRE")
                .Columns.Add("SHIP_BOL_NO")
                .Columns.Add("WAVE_QTY", GetType(System.Int64))
                .PrimaryKey = New DataColumn() { .Columns("SHIP_BOL_NO")}
            End With

            'ASCMAIN1.sql = "Select * from " & TempTable_BOL_NOs
            'Create_TDA(.Tables.Add, TempTable_BOL_NOs, "**", 0, True, "", 1)

            ASCMAIN1.sql = "Select SOTPICK1.PICK_NO, SOTORDR1.CUST_STORE_NO, SOTPICK1.ORDR_NO" & vbCrLf _
                & ", SOTPICK1.PICK_STATUS, SOTPICK1.PICK_RELEASED, SOTPICK1.PICK_FREIGHT" & vbCrLf _
                & ", SOTPICK1.PICK_PICKER, SOTPICK1.PICK_NO_REV" & vbCrLf _
                & ", SOTPICK1.PICK_PRINTED, SOTPICK1.PICK_PACKED, SOTPICK1.PICK_SHIPPED" & vbCrLf _
                & ", SOTPICK1.PICK_BATCH_NO, SOTPICK1.SHIP_BOL_NO, SOTPICK1.INV_NO" & vbCrLf _
                & ", SOTPICK1.PICK_CNT_CARTONS, SOTPICK1.PICK_TOTAL_WGT" & vbCrLf _
                & ", SOTPICK1.INIT_OPER, SOTPICK1.LAST_OPER, SOTPICK1.INIT_DATE, SOTPICK1.LAST_DATE" & vbCrLf _
                & ", SOTPICK0.PICK_FORCED" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTPICK0 " & vbCrLf _
                & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTPICK0.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTPICK1.SHIP_BOL_NO, EDT850T2.EDI_COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf _
                & " from SOTPICK1,SOTPICK2,SOTORDR2,ICTCOLR1,EDT850T2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from SOTCART1,SOTPICK1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
                & ", SOTCART1.PICK_NO, SOTORDR1.CUST_STORE_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
                & " from SOTCART2,SOTCART1,SOTPICK1,SOTORDR2,SOTORDR1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, True, "V", 2)

            'ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_LOCATOR = '1' And WHSE_CODE = '" & IIf(ASCMAIN1.USER_SECURITY_CODEs.Contains("NJC"), "NJC", "NJE") & "'"
            ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_LOCATOR = '1' And WHSE_CODE = '" & ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & "'"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add("SOTSHIPC"), SOTSHIPC, "*")

            ASCMAIN1.sql = "Select WHTWAVE1.WAVE_NO, WHTWAVE1.WAVE_DATE, WHTWAVE1.WAVE_TYPE, WHTWAVE1.WAVE_STATUS" & vbCrLf _
                & ", WHTINSTX.SC_COUNT, WHTINSTX.WAVE_QTY_OPEN, WHTINSTX.WAVE_QTY_PICK" & vbCrLf _
                & ", WHTINSTY.WAVE_COUNT, WHTINSTY.OPEN_COUNT, WHTINSTY.PICK_COUNT" & vbCrLf _
                & ", WHTINSTZ.WAVE_QTY_REL, WHTINSTZ.WAVE_QTY_ADJ, WHTINSTZ.WAVE_QTY_SUB" & vbCrLf _
                & ", WHTINSTZ.WAVE_QTY_CANC, WHTINSTZ.WAVE_QTY_CONC, WHTINSTZ.WAVE_QTY_BACK" & vbCrLf _
                & ", SOTSHIPX.*" & vbCrLf _
                & " from " & SOTSHIPX & " SOTSHIPX, WHTWAVE1, WHTWAVE3" & vbCrLf _
                & ", " & WHTINSTX & " WHTINSTX" & vbCrLf _
                & ", " & WHTINSTY & " WHTINSTY" & vbCrLf _
                & ", " & WHTINSTZ & " WHTINSTZ" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = WHTWAVE1.WAVE_NO" & vbCrLf _
                & "   and WHTWAVE3.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   and WHTINSTX.WAVE_NO (+) = WHTWAVE3.WAVE_NO" & vbCrLf _
                & "   and WHTINSTY.WAVE_NO (+) = WHTWAVE3.WAVE_NO" & vbCrLf _
                & "   and WHTINSTZ.WAVE_NO (+) = WHTWAVE3.WAVE_NO"
            sqlWHTWAVEX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "WHTWAVEX", "**", 0, False, "", 0)
            .Tables("WHTWAVEX").Columns.Add("WAVE_QTY_LEFT", GetType(System.Int64), "ISNULL(WAVE_QTY_REL,0)-ISNULL(WAVE_QTY_OPEN,0)-ISNULL(WAVE_QTY_PICK,0)-ISNULL(WAVE_QTY_ADJ,0)-ISNULL(WAVE_QTY_CANC,0)-ISNULL(WAVE_QTY_CONC,0)-ISNULL(WAVE_QTY_BACK,0)") ' removed -ISNULL(WAVE_QTY_SUB,0)
            .Tables("WHTWAVEX").Columns.Add("WAVE_INST_STATUS_SUMMARY", GetType(System.String), "IIF(WAVE_STATUS='V','VOID',IIF(WAVE_STATUS='F','FINAL',IIF(OPEN_COUNT=0,'COMP',IIF(PICK_COUNT=0,'WAVE','PART'))))")
            .Tables("WHTWAVEX").Columns("SHIP_BOL_NO").AllowDBNull = True

            ASCMAIN1.sql = "Select WHTINST1.* from WHTINST1 where WAVE_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTINST1", "**", 0, True, "V", 1)
            With .Tables("WHTINST1")
                .Columns.Add("SUPP_INSTR")
                .Columns.Add("SELECTED")
                .Columns.Add("LOCATION_ROUTE_SEQ")
                .Columns("SELECTED").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "Select WHTINST2.*" & vbCrLf _
                & " from WHTINST2,WHTINST1" & vbCrLf _
                & " where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
                & "  and WHTINST1.WAVE_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTINST2", "**", 0, True, "V", 4)

            Create_Relation("WHTINST1", "WHTINST2", "WAVE_INST_NO")

            With .Tables("WHTINST2")
                .Columns.Add("CASE_WAVED", GetType(System.Int32), "IIF(ISNULL(LOCATION_QTY_WAVE,0)=0,0,IIF(ISNULL(LOCATION_QTY_WAVE,0)>0,1,-1))")
                .Columns.Add("CASE_PICKED", GetType(System.Int32), "IIF(ISNULL(LOCATION_QTY_PICK,0)=0,0,IIF(ISNULL(LOCATION_QTY_PICK,0)>0,1,-1))")
                .Columns.Add("WAVE_PICK_TYPE", GetType(System.String), "PARENT(WHTINST1_WHTINST2).WAVE_PICK_TYPE")
                .Columns.Add("LOCATION_CODE", GetType(System.String), "PARENT(WHTINST1_WHTINST2).LOCATION_CODE")
                .Columns.Add("WAVE_INST_STATUS", GetType(System.String), "PARENT(WHTINST1_WHTINST2).WAVE_INST_STATUS")
                .Columns.Add("PPK_COUNTED", GetType(System.String))
            End With

            With .Tables("WHTINST1")
                .Columns.Add("UNITS_WAVE", GetType(System.Int32), "SUM(CHILD(WHTINST1_WHTINST2).LOCATION_QTY_WAVE)")
                .Columns.Add("CASES_WAVE", GetType(System.Int32), "SUM(CHILD(WHTINST1_WHTINST2).CASE_WAVED)")
                ' .Columns.Add("CASES_WAVE", GetType(System.Int32), "IIF(UNITS_WAVE<0,-1,1)*COUNT(CHILD(WHTINST1_WHTINST2).WAVE_INST_NO)")
                .Columns.Add("UNITS_PICK", GetType(System.Int32), "SUM(CHILD(WHTINST1_WHTINST2).LOCATION_QTY_PICK)")
                .Columns.Add("CASES_PICK", GetType(System.Int32), "SUM(CHILD(WHTINST1_WHTINST2).CASE_PICKED)")
            End With

            Create_TDA(.Tables.Add, "WHTWAVE1", "*")
            Create_TDA(.Tables.Add, "WHTWAVE2", "*", 1)
            With .Tables("WHTWAVE2")
                .Columns.Add("WAVE_QTY_LOCS", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_SHIP", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_OPEN", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_DIFF", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_LEFT", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_RACS", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_OTS", GetType(System.Int64))
                .Columns.Add("P2L_QTY_OH", GetType(System.Int64))
                .Columns.Add("P2L_QTY_COMMITED", GetType(System.Int64))
                .Columns.Add("P2L_QTY_RESERVE", GetType(System.Int64))
                .Columns.Add("P2L_WO_OPEN", GetType(System.Int64))
                .Columns.Add("P2L_WO_PICK", GetType(System.Int64))
                .Columns.Add("P2L_QTY_NOT_INDUCTED", GetType(System.Int64))
                .Columns.Add("P2L_QTY_AVAILABLE", GetType(System.Int64), "(ISNULL(P2L_QTY_OH,0)+ISNULL(P2L_WO_OPEN,0)+ISNULL(P2L_WO_PICK,0))-(ISNULL(P2L_QTY_COMMITED,0) + ISNULL(P2L_QTY_NOT_INDUCTED,0))")
            End With


            With .Tables.Add("WHTWAVE2_SUB")
                .Columns.Add("WAVE_NO")
                .Columns.Add("WAVE_LNO_LINK", GetType(System.Int64))
                .Columns.Add("WAVE_LNO", GetType(System.Int64))
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("PICK_QTY", GetType(System.Int64))
                .Columns.Add("WAVE_QTY", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_PICK", GetType(System.Int64))
                .Columns.Add("STYLE_CODE_SUB")
                .Columns.Add("COLOR_CODE_SUB")
                .Columns.Add("WAVE_QTY_SUB", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_ADJ", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_NOTE")
                .Columns.Add("WAVE_QTY_OPEN", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_DIFF", GetType(System.Int64))
                .Columns.Add("WAVE_QTY_LEFT", GetType(System.Int64), "ISNULL(WAVE_QTY_SUB,0)-ISNULL(WAVE_QTY_OPEN,0)-ISNULL(WAVE_QTY_PICK,0)-ISNULL(WAVE_QTY_ADJ,0)")
                .PrimaryKey = New DataColumn() { .Columns("WAVE_NO"), .Columns("WAVE_LNO_LINK"), .Columns("WAVE_LNO")}
            End With

            Create_Relation("WHTWAVE2", "WHTWAVE2_SUB", "WAVE_NO,WAVE_LNO", "WAVE_NO,WAVE_LNO_LINK")
            Create_Relation("WHTWAVE2", "SOTCART2", "STYLE_CODE,COLOR_CODE")

            With .Tables("WHTWAVE2")
                .Columns.Add("WAVE_QTY_SUB2", GetType(System.Int64), "SUM(CHILD(WHTWAVE2_WHTWAVE2_SUB).WAVE_QTY_SUB)")
                .Columns.Add("WAVE_QTY_CONF", GetType(System.Int64), "ISNULL(PICK_QTY,0)-ISNULL(WAVE_QTY_CANC,0)")
                .Columns.Add("WAVE_QTY_PACK", GetType(System.Int64), "SUM(CHILD(WHTWAVE2_SOTCART2).QTY_PACKED)")
                .Columns("WAVE_QTY_LEFT").Expression = "ISNULL(PICK_QTY,0)-ISNULL(WAVE_QTY_OPEN,0)-ISNULL(WAVE_QTY_PICK,0)-ISNULL(WAVE_QTY_ADJ,0)-ISNULL(WAVE_QTY_SUB2,0)-ISNULL(WAVE_QTY_CANC,0)-ISNULL(WAVE_QTY_CONC,0)-ISNULL(WAVE_QTY_BACK,0)+(ISNULL(P2L_QTY_RESERVE,0)-ISNULL(P2L_QTY_AVAILABLE,0))"
                .Columns.Add("SUB_WAVE", GetType(System.Int64), "SUM(CHILD(WHTWAVE2_WHTWAVE2_SUB).WAVE_QTY)")
                .Columns.Add("SUB_PICK", GetType(System.Int64), "SUM(CHILD(WHTWAVE2_WHTWAVE2_SUB).WAVE_QTY_PICK)")
                .Columns.Add("SUB_OPEN", GetType(System.Int64), "SUM(CHILD(WHTWAVE2_WHTWAVE2_SUB).WAVE_QTY_OPEN)")
            End With

            ASCMAIN1.sql = "Select WHTWAVE3.WAVE_NO,SOTSHIPX.*, WHTWAVE3.P2L_SHIP_STATUS" & vbCrLf _
                & " from " & SOTSHIPX & " SOTSHIPX, WHTWAVE1, WHTWAVE3" & vbCrLf _
                & " where WHTWAVE1.WAVE_NO = WHTWAVE3.WAVE_NO" & vbCrLf _
                & "   and WHTWAVE3.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   and WHTWAVE3.WAVE_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTWAVE3", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select WHTLOCB1.WHSE_CODE, WHTLOCB1.LOCATION_CODE, WHTLOCB1.BAR_CODE" & vbCrLf _
                & ", WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_QTY" & vbCrLf _
                & ", WHTLOCB1.INIT_DATE, WHTLOCB1.INIT_OPER, WHTLOCB1.LAST_DATE, WHTLOCB1.LAST_OPER, WHTLOCB1.LOCATION_QTY_WAVE" & vbCrLf _
                & ", WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
                & " from WHTLOCB1,WHTBARC1,WHTBARC0,WHTLOCM1" & vbCrLf _
                & " where WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO" & vbCrLf _
                & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
                & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & "   and (NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1' or (WHTLOCM1.LOCATION_CODE in (" & REC_LOCATIONS & ")))" & vbCrLf _
                & "   and (NVL(WHTLOCB1.LOCATION_QTY,0)) > 0" & vbCrLf _
                & "   and WHTLOCB1.WHSE_CODE = :PARM1 and WHTLOCB1.STYLE_CODE = :PARM2 and WHTLOCB1.COLOR_CODE = :PARM3"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "VVV", 0)
            'Add putaway and REC from ICTWHSE1

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", sum (P2L_QTY_OH) P2L_QTY_OH, sum(P2L_QTY_COMMITED) P2L_QTY_COMMITED" & vbCrLf _
                & ", sum (P2L_QTY_RESERVE) P2L_QTY_RESERVE, sum(P2L_WO_OPEN) P2L_WO_OPEN" & vbCrLf _
                & ", sum (P2L_WO_PICK) P2L_WO_PICK, sum(P2L_QTY_NOT_INDUCTED) P2L_QTY_NOT_INDUCTED " & vbCrLf _
                & "From" & vbCrLf _
                & " ( Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
                & ", sum(LOCATION_QTY) P2L_QTY_OH, sum(LOCATION_QTY_WAVE) P2L_QTY_COMMITED" & vbCrLf _
                & ", 0 P2L_QTY_RESERVE, 0 P2L_WO_OPEN, 0 P2L_WO_PICK, 0 P2L_QTY_NOT_INDUCTED" & vbCrLf _
                & " From WHTLOCB1, WHTP2LM1" & vbCrLf _
                & " Where WHTLOCB1.WHSE_CODE = WHTP2LM1.WHSE_CODE" & vbCrLf _
                & "     and WHTLOCB1.LOCATION_CODE = WHTP2LM1.DEPOSIT_LOCATION" & vbCrLf _
                & "     and WHTP2LM1.P2L_STATUS = 'A'" & vbCrLf _
                & "     and WHTP2LM1.CUST_CODE = :PARM1" & vbCrLf _
                & " Group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
                & " Union" & vbCrLf _
                & " Select WHTWAVE2.STYLE_CODE, WHTWAVE2.COLOR_CODE, 0 P2L_QTY_OH, 0 P2L_QTY_COMMITED, 0 P2L_QTY_RESERVE" & vbCrLf _
                & ", Sum (DECODE(WHTINST1.WAVE_INST_STATUS,'0',WHTINST2.LOCATION_QTY_WAVE,0)) P2L_WO_OPEN" & vbCrLf _
                & ", Sum (DECODE(WHTINST1.WAVE_INST_STATUS,'1',WHTINST2.LOCATION_QTY_PICK,0)) P2L_WO_PICK" & vbCrLf _
                & ", 0 P2L_QTY_NOT_INDUCTED" & vbCrLf _
                & " From WHTINST2,WHTINST1,WHTWAVE2,WHTWAVE1" & vbCrLf _
                & " Where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
                & "  and WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                & "  and WHTWAVE2.WAVE_NO = WHTINST1.WAVE_NO and WHTWAVE2.WAVE_LNO = WHTINST1.WAVE_LNO" & vbCrLf _
                & "  and WHTWAVE1.WAVE_STATUS = 'O' and WHTWAVE1.WAVE_TYPE = 'L' and WHTWAVE1.P2L_WAVE_STATUS = 'P'" & vbCrLf _
                & "  and WHTWAVE1.CUST_CODE = :PARM1" & vbCrLf _
                & " Group By WHTWAVE2.STYLE_CODE, WHTWAVE2.COLOR_CODE" & vbCrLf _
                & " Union" & vbCrLf _
                & " Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, 0 P2L_QTY_OH, 0 P2L_QTY_COMMITED, 0 P2L_QTY_RESERVE" & vbCrLf _
                & ",  0 P2L_WO_OPEN, 0 P2L_WO_PICK, Sum(SOTPICK2.PICK_QTY) P2L_QTY_NOT_INDUCTED" & vbCrLf _
                & " From WHTWAVE1, WHTWAVE3, SOTSHIP1, SOTPICK1, SOTPICK2, SOTORDR2" & vbCrLf _
                & " Where WHTWAVE1.WAVE_NO =  WHTWAVE3.WAVE_NO" & vbCrLf _
                & "  And WHTWAVE3.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "  And SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "  And SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "  And SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "  And SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "  And WHTWAVE1.WAVE_STATUS = 'O' and WHTWAVE1.WAVE_TYPE = 'L' and WHTWAVE1.P2L_WAVE_STATUS = 'P'" & vbCrLf _
                & "  and WHTWAVE3.P2L_SHIP_STATUS = 'O'" & vbCrLf _
                & "  and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                & "  and WHTWAVE1.CUST_CODE = :PARM1" & vbCrLf _
                & " Group By SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & " Union" & vbCrLf _
                & " Select WHTSCSEQ.STYLE_CODE, WHTSCSEQ.COLOR_CODE, 0 P2L_QTY_OH, 0 P2L_QTY_COMMITED" & vbCrLf _
                & " , (WHTSCSEQ.PO_QTY_PER_CTN * P2L_MIN_CTNS_PER_LOC) P2L_QTY_RESERVE" & vbCrLf _
                & ", 0 P2L_WO_OPEN, 0 P2L_WO_PICK, 0 P2L_QTY_NOT_INDUCTED" & vbCrLf _
                & " From WHTSCSEQ, WHTP2LM1" & vbCrLf _
                & " Where WHTSCSEQ.CUST_CODE =  WHTP2LM1.CUST_CODE" & vbCrLf _
                & " and WHTP2LM1.P2L_STATUS = 'A'" & vbCrLf _
                & " and WHTP2LM1.CUST_CODE = :PARM1 )" & vbCrLf _
                & " Group By STYLE_CODE, COLOR_CODE" & vbCrLf
            Create_TDA(.Tables.Add, "WHTWAVEP2L", "**", 0, False, "V", 0)

            With .Tables("WHTLOCB1")
                .Columns.Add("CASES", GetType(System.Int32))
                .Columns.Add("CASE_BREAKDOWN")
                .Columns("WHSE_CODE").AllowDBNull = True
                .Columns("LOCATION_CODE").AllowDBNull = True
                .Columns("BAR_CODE").AllowDBNull = True
                .Columns("STYLE_CODE").AllowDBNull = True
                .Columns("COLOR_CODE").AllowDBNull = True
                .Columns.Add("LOCATION_QTY_AVAIL", GetType(System.Int64), "ISNULL(LOCATION_QTY,0)-ISNULL(LOCATION_QTY_WAVE,0)")
            End With

            ASCMAIN1.sql = "Select POTORDR1.WHSE_CODE, POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
              & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
              & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
              & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO, POTSHIP3.PO_QTY_SHP" & vbCrLf _
              & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
              & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
              & ", POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE" & vbCrLf _
              & " From POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2 " & vbCrLf _
              & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 0)
            With .Tables("POTORDRX")
                .Columns("PO_SHIPMENT_NO").AllowDBNull = True
                .Columns("PO_SHIPMENT_LNO").AllowDBNull = True
                .Columns("PO_REFERENCE").AllowDBNull = True
            End With

            With .Tables.Add("WHTWAVES")
                .Columns.Add("WAVE_PICK_TYPE")
                .Columns.Add("WAVE_STAT_TYPE")
                .Columns.Add("WAVE", GetType(System.Int64))
                .Columns.Add("PICK", GetType(System.Int64))
            End With

            With .Tables.Add("WHTWAVED")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("WAVE_LNO", GetType(System.Int32))
                .Columns.Add("WAVE_QTY", GetType(System.Int64))
                .Columns.Add("WAVE_SUB")
                .Columns.Add("WAVE_LNO_LINK", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE"), .Columns("WAVE_LNO")}
            End With

            Create_TDA(.Tables.Add("WHTLOCBW"), WHTLOCBW, "*", 0, , , 5)


            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "WHTBARC0", "*")

            ASCMAIN1.sql = "Select WHTLOCM1.* from WHTLOCM1 where WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "WHTLOCM1", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            Create_TDA(.Tables.Add, "ICTIADJ1", "*")
            Create_TDA(.Tables.Add, "ICTIADJ2", "*")

            Create_Relation("ICTIADJ1", "ICTIADJ2", "ADJ_NO")

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTSHIP1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", , , "V")

            Create_TDA(.Tables.Add, "SOTSHIP1", "*", , , , , "SHIP_DATE_ROUTED,SHIP_DATE_PLANNED,SHIP_DATE_PACKED,SHIP_APPT_NO,SHIP_NOTES,SHIP_NOTES_3PL")

            ASCMAIN1.sql = "Select INIT_DATE, USER_ID INIT_OPER, COLUMN_NAME, OLD_VALUE, NEW_VALUE" _
                & " from ASTAUDT1 where TABLE_NAME = 'SOTSHIP1' and KEY_VALUE = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIPA", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "" _
                & "Select PPK_CODE, PPK, SCS, COUNT (*) CASES, SUM (PPK) QTY, MIN (BAR_CODE) MINBAR, MAX (BAR_CODE) MAXBAR " & vbCrLf _
                & ", PO_DATE_RECEIVED, TRAN_NO, PO_SHIPMENT_NO, PO_SHIPMENT_LNO" & vbCrLf _
                & " from (" & vbCrLf _
                & "SELECT WHTLOCB1.BAR_CODE, SUM (WHTLOCB1.LOCATION_QTY) PPK" & vbCrLf _
                & ", COUNT (DISTINCT WHTLOCB1.STYLE_CODE || WHTLOCB1.COLOR_CODE) SCS" & vbCrLf _
                & ", WHTBARC1.PPK_CODE, WHTBARC1.PO_DATE_RECEIVED, WHTBARC1.TRAN_NO, WHTBARC1.PO_SHIPMENT_NO, WHTBARC1.PO_SHIPMENT_LNO" & vbCrLf _
                & " from WHTLOCB1, WHTBARC1, " & SOTSHIPC & " SOTSHIPC" & vbCrLf _
                & " where WHTLOCB1.WHSE_CODE = :PARM1 " & vbCrLf _
                & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
                & "   and WHTBARC1.PPK_CODE IS NOT NULL" & vbCrLf _
                & "   and WHTLOCB1.STYLE_CODE = SOTSHIPC.STYLE_CODE and WHTLOCB1.COLOR_CODE = SOTSHIPC.COLOR_CODE" & vbCrLf _
                & " group by WHTLOCB1.BAR_CODE" & vbCrLf _
                & ", WHTBARC1.PPK_CODE, WHTBARC1.PO_DATE_RECEIVED, WHTBARC1.TRAN_NO, WHTBARC1.PO_SHIPMENT_NO, WHTBARC1.PO_SHIPMENT_LNO" & vbCrLf _
                & ") GROUP BY PPK, SCS, PPK_CODE, PO_DATE_RECEIVED, TRAN_NO, PO_SHIPMENT_NO, PO_SHIPMENT_LNO"
            Create_TDA(.Tables.Add, "WHTWAVEP", "**", 0, False, "V", 0)
            .Tables("WHTWAVEP").Columns.Add("SEL")
            .Tables("WHTWAVEP").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select * from WHTPPKM2 where PPK_CODE = :PARM1"
            Create_TDA(.Tables.Add, "WHTPPKM2", "**", 0, False, "V", 3)

            ASCMAIN1.sql = "Select * from WHTSCSEQ where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "WHTSCSEQ", "**", 0, False, "V", 3)
        End With

        Fill_Records("ICTWHSE1")

        For Each WAVE_PICK_TYPE As String In New String() {"L", "C", "U"}
            For Each WAVE_STAT_TYPE As String In New String() {"I", "C", "U"}
                dst.Tables("WHTWAVES").Rows.Add(New String() {WAVE_PICK_TYPE, WAVE_STAT_TYPE})
            Next
        Next

        grdWHTPPKM2.DataSource = dst.Tables("WHTPPKM2")
        grdWHTWAVEP.DataSource = dst.Tables("WHTWAVEP")
        grdWHTINSTS.DataSource = dst.Tables("WHTINST2")
        grdICTIADJ1.DataSource = dst.Tables("ICTIADJ1")
        grdSOTSHIPA.DataSource = dst.Tables("SOTSHIPA")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdWHTWAVE2.DataSource = dst.Tables("WHTWAVE2")
        grdWHTWAVE3.DataSource = dst.Tables("WHTWAVE3")
        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")

        grdWHTWAVEX.DataSource = dst.Tables("WHTWAVEX")

        grdWHTLOCB1.DataSource = dst.Tables("WHTLOCB1")
        grdWHTINST1.DataSource = dst.Tables("WHTINST1")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdWHTWAVES.DataSource = dst.Tables("WHTWAVES")
        grdWHTWAVES.DisplayLayout.Bands(0).Columns("WAVE_PICK_TYPE").SortComparer = New srtComparerWHTWAVES

        Create_Summary(grdWHTWAVEP, "PPK_CODE", "Count")
        Create_Summary(grdWHTWAVEP, New String() {"SEL", "CASES", "QTY"})

        Create_Summary(grdWHTPPKM2, "STYLE_CODE", "Count")
        Create_Summary(grdWHTPPKM2, New String() {"PPK_QTY"})

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTSHIPX, New String() {"PICK_NO_COUNT", "SELECTED", "PICK_QTY_PICK", "WAVE_QTY"})

        Create_Summary(grdSOTCART2, "CUST_STORE_NO", "Count")
        Create_Summary(grdSOTCART2, New String() {"QTY_PACKED", "QTY_PACKED_ORIG"})

        Create_Summary(grdWHTWAVE2, "STYLE_CODE", "Count")
        Create_Summary(grdWHTWAVE2, New String() {"PICK_QTY", "WAVE_QTY_LEFT", "WAVE_QTY_LOCS", "WAVE_QTY_ADJ", "WAVE_QTY_SHIP", "WAVE_QTY_DIFF", "WAVE_QTY_SUB2", "WAVE_QTY_CANC", "WAVE_QTY_CONC", "WAVE_QTY_BACK", "WAVE_QTY_CONF", "WAVE_QTY_PACK"})
        Create_Summary(grdWHTWAVE2, "WAVE_QTY", "Custom")
        Create_Summary(grdWHTWAVE2, "WAVE_QTY_PICK", "Custom")
        Create_Summary(grdWHTWAVE2, "WAVE_QTY_OPEN", "Custom")

        Create_Summary(grdWHTLOCB1, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCB1, New String() {"LOCATION_QTY", "LOCATION_QTY_AVAIL", "CASES"})

        Create_Summary(grdWHTINST1, New String() {"SELECTED", "CASES_WAVE", "CASES_PICK", "UNITS_WAVE", "UNITS_PICK"})

        Create_Summary(grdWHTWAVEX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdWHTWAVEX, New String() {"SC_COUNT", "WAVE_COUNT", "OPEN_COUNT", "PICK_COUNT"})

        Create_Summary(grdWHTINSTS, "BAR_CODE", "Count")
        Create_Summary(grdWHTINSTS, New String() {"CASE_WAVED", "CASE_PICKED"})

        Create_Summary(grdPOTORDRX, New String() {"PO_QTY_SHP"})

        grdWHTWAVES.DisplayLayout.Bands(0).Override.GroupByRowDescriptionMask = "[caption] : [value]"

        With grdWHTWAVES.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With

        grdWHTWAVEP.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        With grdWHTWAVEP.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                End If
            Next
        End With

        grdSOTSHIPX.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdSOTSHIPX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            For Each COLUMN_NAME As String In New String() {"SELECTED", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_GROUP_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_ADDR_TYPE", "SHIP_ADDR_CODE", "FRT_TERMS", "SHIP_VIA_CODE", "WHSE_CODE", "SREP_CODE", "ORDR_DEPT", "SHIP_REF"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_DATE_RECEIVED", "SHIP_DATE_PLANNED", "SHIP_DATE_ROUTED", "SHIP_DATE_PACKED", "SHIP_APPT_NO"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_NOTES", "SHIP_NOTES_3PL"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
        End With

        'grdWHTWAVE2.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Hide
        'grdWHTWAVE2.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide

        grdWHTWAVE2.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdWHTWAVE2.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True

            .Columns("WAVE_QTY_BACK").Hidden = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If gcol.Key = "WAVE_QTY_CANC" Or gcol.Key = "WAVE_QTY_CONC" Or gcol.Key = "WAVE_QTY_BACK" Or gcol.Key = "WAVE_QTY_ADJ" Or gcol.Key = "WAVE_QTY_NOTE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                If gcol.Key = "CARTON_NO" Then
                    gcol.CellAppearance.TextHAlign = HAlign.Right
                End If
            Next

            'For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE"}
            '    .Columns(COLUMN_NAME).Header.Fixed = True
            'Next
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE", "PICK_QTY", "WAVE_QTY"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
            For Each COLUMN_NAME As String In New String() {"WAVE_QTY_PICK", "WAVE_QTY_ADJ", "WAVE_QTY_NOTE", "WAVE_QTY_OPEN", "WAVE_QTY_DIFF", "WAVE_QTY_LEFT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"WAVE_QTY_SHIP", "WAVE_QTY_LOCS"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next
            For Each COLUMN_NAME As String In New String() {"WAVE_QTY_CANC", "WAVE_QTY_CONC", "WAVE_QTY_BACK", "WAVE_QTY_SUB2", "WAVE_QTY_CONF", "WAVE_QTY_PACK"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
        End With

        With grdWHTWAVE2.DisplayLayout.Bands(1)
            .ColHeadersVisible = False
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If gcol.Key = "WAVE_QTY_SUB" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        grdSOTCART2.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdSOTCART2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                If gcol.Key = "QTY_PACKED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "QTY_PACKED"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"QTY_PACKED", "QTY_PACKED_ORIG"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
            For Each COLUMN_NAME As String In New String() {"CART_NO", "UPC_CODE", "SKU_NO", "SIZE_DESC", "STYLE_PREPACK"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
        End With

        grdWHTINST1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdWHTINST1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdWHTINST1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            For Each COLUMN_NAME As String In New String() {"SELECTED", "WAVE_PICK_TYPE", "LOCATION_CODE", "WAVE_INST_STATUS", "CASES_WAVE", "UNITS_WAVE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdWHTINST1.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next

            'For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_GROUP_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            '    .Columns(COLUMN_NAME).Header.Fixed = True
            'Next
        End With

        grdWHTWAVE3.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdWHTWAVE3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
            Next

            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_GROUP_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_GROUP_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_ADDR_TYPE", "SHIP_ADDR_CODE", "FRT_TERMS", "WHSE_CODE", "SREP_CODE", "ORDR_DEPT", "SHIP_REF"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_NOTES", "SHIP_DATE_RECEIVED", "SHIP_DATE_PLANNED", "SHIP_DATE_ROUTED", "SHIP_NOTES_3PL", "SHIP_VIA_CODE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next
        End With

        With grdWHTWAVEX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next

            ' "SHIP_BOL_NO",
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_GROUP_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_ADDR_TYPE", "SHIP_ADDR_CODE", "FRT_TERMS", "WHSE_CODE", "SREP_CODE", "ORDR_DEPT", "SHIP_REF", "SHIP_VIA_CODE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"SC_COUNT", "WAVE_COUNT", "OPEN_COUNT", "PICK_COUNT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Tan
            Next

            For Each COLUMN_NAME As String In New String() {"SHIP_NOTES", "SHIP_DATE_RECEIVED", "SHIP_DATE_PLANNED", "SHIP_DATE_ROUTED", "SHIP_NOTES_3PL"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next

            For Each COLUMN_NAME As String In New String() {"WAVE_QTY_REL", "WAVE_QTY_ADJ", "WAVE_QTY_SUB", "WAVE_QTY_CANC", "WAVE_QTY_CONC", "WAVE_QTY_BACK", "WAVE_QTY_LEFT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next

            'For Each COLUMN_NAME As String In New String() {"SHIP_STATUS", "SHIP_DATE_SHIPPED", "SHIPPED_ACTUAL"}
            '    .Columns(COLUMN_NAME).Hidden = False
            'Next
        End With

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            grdWHTWAVE3.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
            grdWHTWAVEX.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
        End If

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS", Nothing, New String() {":", "P:Pick", "F:Shipped", "C:Cancelled", "D:Deleted"})
        ASCMAIN1.Add_Value_List(grdWHTINST1, "WAVE_PICK_TYPE", Nothing, New String() {":", "L:Pallet", "C:Case", "U:Piece"})
        ASCMAIN1.Add_Value_List(grdWHTINST1, "WAVE_INST_STATUS") ' , Nothing, New String() {":", "0:Waved", "1:Picked"})

        ASCMAIN1.Add_Value_List(grdWHTINSTS, "WAVE_PICK_TYPE", Nothing, New String() {":", "L:Pallet", "C:Case", "U:Piece"})
        ASCMAIN1.Add_Value_List(grdWHTINSTS, "WAVE_INST_STATUS") ' , Nothing, New String() {":", "0:Waved", "1:Picked"})


        ASCMAIN1.Add_Value_List(grdWHTWAVES, "WAVE_PICK_TYPE", Nothing, New String() {":", "L:Pallet", "C:Case", "U:Piece"})
        ASCMAIN1.Add_Value_List(grdWHTWAVES, "WAVE_STAT_TYPE", Nothing, New String() {":", "I:Picks", "C:Cases", "U:Units"})

        ASCMAIN1.Add_Value_List(grdWHTWAVEX, "WAVE_INST_STATUS_SUMMARY", Nothing, New String() {":", "PART:Partial", "COMP:Complete", "WAVE:Waved", "FINAL:Finalized", "VOID:Voided"})
        ASCMAIN1.Add_Value_List(grdWHTWAVEX, "SHIP_STATUS")

        ASCMAIN1.Add_Value_List(grdWHTWAVE3, "SHIP_STATUS")

        Show_Filter(grdWHTWAVEX, True)

        Toggle_grdWHTINST1()

        calFrom.Value = Now.Date.AddDays(-60)
        calTo.Value = Now.Date

        ASCMAIN1.Add_Value_List(cmbType, "SHIP_EVENT")
        Bind_Controls(grpEditShipment, "SOTSHIP1")
        Toggle_EditShipment()

        tabSOTSHIPX.Tabs("Adjustments").Visible = InquiryMode

        Dim COLUMN_NAMEs As New List(Of String)
        COLUMN_NAMEs.Add(":")
        For Each COLUMN_NAME As String In New String() {"SHIP_DATE_ROUTED", "SHIP_DATE_PLANNED", "SHIP_DATE_PACKED",
                                                        "SHIP_APPT_NO", "SHIP_NOTES", "SHIP_NOTES_3PL"}
            COLUMN_NAMEs.Add(COLUMN_NAME & ":" & grdSOTSHIPX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption)
        Next
        ASCMAIN1.Add_Value_List(grdSOTSHIPA, "COLUMN_NAME", Nothing, COLUMN_NAMEs.ToArray)

        grdWHTLOCB1_app1.ForeColor = Drawing.Color.Magenta
        grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOCATION_CODE").Header.ToolTipText = "Magenta = 'Location Locked'"

        btnP2L.Visible = ASCMAIN1.Running_in_VS


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Deposit"
                If Absx1.txtFor("LOCATION_CODE_DEPOSIT").Text = "" Then
                    EMsg &= vbCr & IIf(WAVE_TYPE = "L", "Pick Line must be selected", "Wave Deposit Location is Mandatory")
                Else
                    Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("LOCATION_CODE_DEPOSIT").Text})
                    If rowWHTLOCM1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Wave Deposit Location"
                    End If
                End If
                If dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '1'").Length = 0 Then
                    EMsg &= vbCr & "There are no Picks that have not been Deposited."
                End If

                'If Not ASCMAIN1.Running_in_VS Then
                '    EMsg &= vbCr & "Only for ABS at this time"
                'End If

            Case "Preview"
                If dst.Tables("SOTSHIPX").Select("SELECTED='1'").Length = 0 Then
                    EMsg &= vbCr & "No Shipments Selected for Preview"
                End If

            Case "New Work Order"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 IsNot Nothing Then
                        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Warehouse " & Absx1.txtFor("WHSE_CODE").Text
                    End If
                End If

            Case "New", "New P2L Order"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 IsNot Nothing Then
                        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Warehouse " & Absx1.txtFor("WHSE_CODE").Text
                    End If
                End If

                If dst.Tables("SOTSHIPX").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "You Must Select at least 1 Shipment to Generate a Wave"
                Else
                    Dim CUST_CODEs As New List(Of String)
                    Dim P2L_Flags As New List(Of String)
                    P2L_ALLOW = ""
                    For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SELECTED = '1'")
                        CUST_CODE = rowSOTSHIPX.Item("CUST_CODE")
                        If Not CUST_CODEs.Contains(CUST_CODE) Then CUST_CODEs.Add(CUST_CODE)
                        P2L_ALLOW = rowSOTSHIPX.Item("P2L_ALLOW")
                        If Not P2L_Flags.Contains(P2L_ALLOW) Then
                            P2L_Flags.Add(P2L_ALLOW)
                        End If
                    Next

                    If CUST_CODEs.Count > 1 Then EMsg &= vbCr & "All Shipments in a Wave must belong to a Single Customer"
                    If P2L_Flags.Count > 1 Then EMsg &= vbCr & "Mixed P2L shipments not allowed"
                End If

                If EMsg = "" And eItemKey = "New P2L Order" Then
                    If P2L_ALLOW <> "Y" Then
                        EMsg &= vbCr & "Shipment is not Eligeble for Pick to Light"
                    Else
                        sqlSHIP_BOL_NOs = ""
                        For Each row As DataRow In dst.Tables("SOTSHIPX").Select("ISNULL(SELECTED,'0')='1'")
                            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
                            sqlSHIP_BOL_NOs &= ",'" & SHIP_BOL_NO & "'"
                        Next
                        Fill_Records("WHTSCSEQ", CUST_CODE)

                        ASCMAIN1.sql = "Select Distinct SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                            & " from SOTORDR1,SOTORDR2,SOTPICK2,SOTPICK1" & vbCrLf _
                            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                            & "   and SOTPICK1.SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NOs, 2) & ")" & vbCrLf
                        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                            Dim Cnt As Int32 = dst.Tables("WHTSCSEQ").Compute("Count(STYLE_CODE)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                            If Cnt = 0 Then
                                EMsg &= vbCr & "Shipment has invalid styles for Pick to Light"
                                Exit For
                            End If
                        Next
                        ASCMAIN1.sql = "Select * From WHTP2lM1 WHERE CUST_CODE = :PARM1 AND WHSE_CODE = :PARM2 AND P2L_STATUS = 'A'"
                        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", New Object() {CUST_CODE, WHSE_CODE}).Select("")
                            Absx1.txtFor("LOCATION_CODE_DEPOSIT").Text = row.Item("DEPOSIT_LOCATION")
                            Absx1.txtFor("P2L_LINE_ID").Text = row.Item("P2L_LINE_ID")
                        Next

                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("WHTWAVE1", WHSE_CODE) Then Exit Sub
                End If
                If EMsg = "" Then
                    ASCDATA1.DeleteRows("SOTSHIPX", "ISNULL(SELECTED,'0')<>'1'")

                    For Each row As DataRow In dst.Tables("SOTSHIPX").Select("")
                        Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
                        If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                        If rowSOTSHIP1.Item("SHIP_WAVE_STATUS") & "" = "1" Then
                            EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " has been Waved"
                            ASCMAIN1.MultiTask_Release()
                        End If
                    Next
                End If

            Case "Edit", "View"

                ' ASCMAIN1.Format_Field(Absx1.txtFor("WAVE_NO").Text, "WAVE_NO")
                '    Validate_Code("WAVE_NO")

                WHSE_CODE = ""
                WAVE_NO = ""
                P2L_ALLOW = "N"

                If Absx1.txtFor("WAVE_NO").Text = "" Then
                    EMsg &= vbCr & "No Wave No Specified"
                Else
                    WAVE_NO = Absx1.txtFor("WAVE_NO").Text
                    rowWHTWAVE1 = LookUp("WHTWAVE1", WAVE_NO)
                    If rowWHTWAVE1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Wave No " & WAVE_NO
                    Else
                        CUST_CODE = rowWHTWAVE1.Item("CUST_CODE") & ""
                        WHSE_CODE = rowWHTWAVE1.Item("WHSE_CODE") & ""

                        If rowWHTWAVE1.Item("WAVE_TYPE") = "L" Then P2L_ALLOW = "Y"

                        Absx1.txtFor("WHSE_CODE").Text = rowWHTWAVE1.Item("WHSE_CODE") & ""
                        If rowWHTWAVE1.Item("WAVE_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowWHTWAVE1.Item("WAVE_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "Wave No " & WAVE_NO & " has been Cancelled"
                                Case "D"
                                    EMsg &= vbCr & "Wave No " & WAVE_NO & " has been Deleted"
                                Case "F"

                                    ASCMAIN1.sql = "Select * from SOTSHIP1" _
                                     & " Where WAVE_NO = '" & WAVE_NO & "'"
                                    For Each rowSOTSHIP1 As DataRow In ASCDATA1.GetDataTable.Rows
                                        If rowSOTSHIP1.Item("SHIP_STATUS") <> "P" Then
                                            EMsg &= vbCr & "Wave Contains BOL's No Longer in Pick, Cannot Edit"
                                            Exit For
                                        End If
                                    Next


                                Case Else
                                    EMsg &= vbCr & "Wave No " & WAVE_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("WHTWAVE1", WAVE_NO) Then Exit Sub


                    ASCMAIN1.sql = "Select * from SOTSHIP1" _
                    & " Where WAVE_NO = '" & WAVE_NO & "'"
                    For Each rowSOTSHIP1 As DataRow In ASCDATA1.GetDataTable.Rows
                        If Not ASCMAIN1.Logical_Lock("SOTSHIP1", rowSOTSHIP1.Item("SHIP_BOL_NO")) Then Exit Sub
                    Next
                End If

            Case "Update"
                If Absx1.dteFor("WAVE_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Wave Date is Mandatory"
                End If

                If Absx1.txtFor("LOCATION_CODE_DEPOSIT").Text = "" Then
                    EMsg &= vbCr & "Wave Deposit Location is Mandatory"
                Else
                    Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("LOCATION_CODE_DEPOSIT").Text})
                    If rowWHTLOCM1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Wave Deposit Location"
                    End If
                End If

                Dim qty_was_waved As Boolean = False
                For Each rowWHTWAVE2 As DataRow In dst.Tables("WHTWAVE2").Select("")
                    Dim STYLE_CODE As String = rowWHTWAVE2.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowWHTWAVE2.Item("COLOR_CODE")
                    If Val(rowWHTWAVE2.Item("WAVE_QTY_CANC") & "") <> 0 _
                    Or Val(rowWHTWAVE2.Item("WAVE_QTY_CONC") & "") <> 0 _
                    Or Val(rowWHTWAVE2.Item("WAVE_QTY_BACK") & "") <> 0 _
                    Or Val(rowWHTWAVE2.Item("WAVE_QTY_ADJ") & "") <> 0 Then
                        If rowWHTWAVE2.Item("WAVE_QTY_NOTE") & "" = "" Then
                            EMsg &= vbCr & "Note Required for Cancel, Conceal, Back-Order or Pick Adjustment: " & STYLE_CODE & "-" & COLOR_CODE
                        End If
                    End If
                    If Val(rowWHTWAVE2.Item("WAVE_QTY") & "") <> 0 Then
                        qty_was_waved = True
                    End If
                    Dim subs() As DataRow = rowWHTWAVE2.GetChildRows("WHTWAVE2_WHTWAVE2_SUB")
                    If subs.Length <> 0 Then
                        For Each subSC As DataRow In subs
                            If subSC.Item("STYLE_CODE") = subSC("STYLE_CODE_SUB") & "" _
                            And subSC.Item("COLOR_CODE") = subSC("COLOR_CODE_SUB") & "" Then
                                EMsg &= vbCr & "Cannot Sub a Style to itself (see " & subSC.Item("STYLE_CODE") & "-" & subSC.Item("COLOR_CODE") & ")"
                            End If
                            If Val(subSC.Item("WAVE_QTY_SUB") & "") = 0 Then
                                If Val(rowWHTWAVE2.Item("WAVE_QTY_SUB2") & "") = 0 Then
                                    'EMsg &= vbCr & "Sub Qty Not Specified (see subs for " & subSC.Item("STYLE_CODE") & "-" & subSC.Item("COLOR_CODE") & ")"
                                End If
                            End If
                            If WAVE_TYPE = "L" Then
                                Dim isP2L As Int64 = Val(dst.Tables("WHTSCSEQ").Compute("COUNT(STYLE_CODE)", "STYLE_CODE='" & subSC("STYLE_CODE_SUB") & "' and COLOR_CODE='" & subSC("COLOR_CODE_SUB") & "'") & "")
                                If isP2L = 0 Then
                                    EMsg &= vbCr & "Sub Style W/O Match No or Loc " & subSC.Item("STYLE_CODE") & "-" & subSC.Item("COLOR_CODE") & "."
                                End If
                            End If
                        Next
                    End If
                Next

                If chkFinalize.Checked Then
                    If dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '1'").Length <> 0 Then
                        EMsg &= vbCr & "Cannot Finalize a Wave which has Picks that have not yet been Deposited."
                    End If
                    If WAVE_TYPE = "W" Or WAVE_TYPE = "L" Then
                        ' NO NEED TO BALANCE WORK ORDER WAVES WITH QTY PACKED
                    Else
                        If dst.Tables("WHTWAVE2").Select("ISNULL(WAVE_QTY_CONF,0) <> ISNULL(WAVE_QTY_PACK,0)").Length <> 0 Then
                            ' EMsg &= vbCr & "Cannot Finalize a Wave which is Out of Balance between Qty Confirmed and Qty Packed"
                            If MsgBox("Qty Confirmed does NOT equal Qty Packed." & vbCrLf & vbCrLf & "OK to Continue?", MsgBoxStyle.OkCancel, "Verificaiton") = MsgBoxResult.Cancel Then
                                Exit Sub
                            End If
                        End If
                    End If
                    If WAVE_TYPE = "L" Then
                        If dst.Tables("WHTWAVE3").Select("P2L_SHIP_STATUS <> 'C'").Length <> 0 Then
                            EMsg &= vbCr & "Cannot Finalize a Wave which has Shipments that have not yet Fully Picked in P2L."
                        End If

                    End If

                End If
                'If Not qty_was_waved Then
                '    EMsg &= vbCr & "Nothing Waved - no point in Updating"
                'End If


            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

                If Pick_Test() Then
                    EMsg &= vbCr & "Wave has been (at least partially) Picked - Cannot be Deleted"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Wave as Deleted" _
                                  & vbCrLf & "  and Restore Customer Shipments to Released Not Waved Status",
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Void Wave"

                If Not Pick_Test() Then
                    EMsg &= vbCr & "Wave has never been Picked - Use Delete function"
                Else
                    If WAVE_INST_NO_void.Count <> 0 Then
                        EMsg &= vbCr & "You have Voided Picks pending update." & vbCr & "  Update the Wave first, and then call it back up to Void"
                    Else
                        Dim PICK As Integer = dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '1'").Length
                        If PICK <> 0 Then
                            EMsg &= vbCr & "Some Styles have been Picked, but not Deposited." & vbCr & "  Complete all Deposits, and then call up the Wave again to Void"
                        End If
                    End If
                    If EMsg = "" Then
                        If MsgBox("Do you want to Void this Wave" _
                                  & vbCrLf & "  and Restore Customer Shipments to Released Not Waved Status",
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
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
            Case "Deposit"
                Dim WAVE_NO_to_deposit = Me.WAVE_NO

                Perform_Deposit()
                Mode_Settings(False)

                Absx1.txtFor("WAVE_NO").Text = WAVE_NO_to_deposit
                Click_Command("View")

            Case "Refresh"
                SHIP_BOL_NO = ""
                sql_Wave_Filter = " and SOTSHIP1.SHIP_STATUS in ('P')"
                Load_SOTSHIPX()

            Case "Preview"
                Wave_Preview()

            Case "New Work Order"
                WAVE_TYPE = "W"
                CUST_CODE = ""
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "New"
                WAVE_TYPE = "S"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "New P2L Order"
                WAVE_TYPE = "L"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"

                Dim WAVE_NO_TO_UPDATE As String = WAVE_NO

                Update_Record()
                Record_Event("UPDATE WAVE", IIf(EntryMode = "N", "New ", "Existing ") & "Wave Updated")
                Mode_Settings(False)

                ASCMAIN1.sql = "Select * from (" & vbCrLf _
                    & "Select WHTINST2.BAR_CODE, COUNT (*) INSTS" & vbCrLf _
                    & ", Sum (DECODE(WHTINST1.WAVE_INST_STATUS,'0',1,0)) OPEN" & vbCrLf _
                    & ", Sum (DECODE(WHTINST1.WAVE_INST_STATUS,'1',1,0)) PCKD" & vbCrLf _
                    & ", Sum (DECODE(WHTINST1.WAVE_PICK_TYPE,'L',1,0)) L" & vbCrLf _
                    & ", Sum (DECODE(WHTINST1.WAVE_PICK_TYPE,'C',1,0)) C" & vbCrLf _
                    & ", Sum (DECODE(WHTINST1.WAVE_PICK_TYPE,'U',1,0)) U" & vbCrLf _
                    & ", Min (WHTINST1.WAVE_NO) W1, MAX (WHTINST1.WAVE_NO) W2" & vbCrLf _
                    & ", Min (WHTINST1.WAVE_INST_NO) I1, MAX (WHTINST1.WAVE_INST_NO) I2" & vbCrLf _
                    & " from whtinst2,WHTINST1" & vbCrLf _
                    & " where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
                    & "   and (WHTINST1.WAVE_INST_STATUS = '0' OR ((WHTINST1.WAVE_INST_STATUS = '1' or WHTINST1.WAVE_INST_STATUS = '2') and NVL(WHTINST2.LOCATION_QTY_PICK,0) <> 0))" & vbCrLf _
                    & "   and WHTINST2.BAR_CODE IN (" & vbCrLf _
                    & "Select WHTINST2.BAR_CODE FROM WHTINST2,WHTINST1" & vbCrLf _
                    & " where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO " & vbCrLf _
                    & "   and WHTINST1.WAVE_NO = '" & WAVE_NO_TO_UPDATE & "'" & vbCrLf _
                    & "   and (WHTINST1.WAVE_INST_STATUS = '0' OR ((WHTINST1.WAVE_INST_STATUS = '1' or WHTINST1.WAVE_INST_STATUS = '2') and NVL(WHTINST2.LOCATION_QTY_PICK,0) <> 0))" & vbCrLf _
                    & " group by WHTINST2.BAR_CODE HAVING COUNT (DISTINCT WHTINST1.WAVE_INST_NO) > 1)" & vbCrLf _
                    & " group by WHTINST2.bar_code" & vbCrLf _
                    & ") where L<>0 OR C<> 0"

                Dim TBL As DataTable = ASCDATA1.GetDataTable
                If TBL.Rows.Count <> 0 Then
                    Using F As New ASFMSGBF
                        F.Show_grd(TBL, Me, "Cases with Multiple Instructions, Wave " & WAVE_NO_TO_UPDATE & " - Please alert ABS")
                    End Using
                End If

                ASCMAIN1.sql = "" _
                   & "Select WHTLOCB1.* from WHTLOCB1" & vbCrLf _
                   & " where WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                   & " and (LOCATION_CODE, BAR_CODE) in " & vbCrLf _
                   & "(Select Distinct WHTINST1.LOCATION_CODE, WHTINST2.BAR_CODE" & vbCrLf _
                   & " from WHTINST1,WHTINST2" & vbCrLf _
                   & " where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
                   & "   and WHTINST1.WAVE_NO = '" & WAVE_NO_TO_UPDATE & "')" & vbCrLf _
                   & " and (LOCATION_QTY < 0 or (NVL(LOCATION_QTY,0) - NVL(LOCATION_QTY_WAVE,0)) < 0)"

                Dim TBL2 As DataTable = ASCDATA1.GetDataTable
                If TBL2.Rows.Count <> 0 Then
                    Using F As New ASFMSGBF
                        F.Show_grd(TBL2, Me, "Cases with Negative Availability, Wave " & WAVE_NO_TO_UPDATE & " - Please alert ABS")
                    End Using
                End If


            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Void Wave"
                Void_Wave()
                Record_Event("VOID WAVE", IIf(rowWHTWAVE1.Item("WAVE_STATUS") = "F", "Finalized ", "") & "Wave Voided")
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then

            With UltraExplorerBar1
                .Groups("Wave Stats").Visible = ScreenMode
                .Groups("Find Waves").Visible = Not ScreenMode And InquiryMode
                .Groups("Wave Status Filter").Visible = Not ScreenMode And InquiryMode

                If ScreenMode Or InquiryMode Then
                    .Groups("Edit Shipment").Visible = False
                End If

                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If EntryMode = "V" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode = "V" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                        .Items("Delete").Settings.Enabled = not_iScreenMode
                        .Items("Void Wave").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                        .Items("Delete").Settings.Enabled = iScreenMode
                        .Items("Void Wave").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Refresh").Visible = Not ScreenMode

                    .Items("New Work Order").Visible = (EntryMode <> "V") And Not ScreenMode And Not InquiryMode
                    .Items("New P2L Order").Visible = (EntryMode <> "V") And Not ScreenMode And Not InquiryMode
                    .Items("New").Visible = (EntryMode <> "V") And Not ScreenMode And Not InquiryMode
                    .Items("Edit").Visible = Not InquiryMode
                    .Items("View").Visible = (EntryMode <> "N" And EntryMode <> "E")
                    .Items("Done").Visible = (EntryMode = "V")
                    .Items("Print").Visible = ScreenMode
                    .Items("Update").Visible = (EntryMode <> "V") And Not InquiryMode
                    .Items("Cancel").Visible = (EntryMode <> "V") And Not InquiryMode
                    .Items("Delete").Visible = (EntryMode = "E") And Not InquiryMode
                    .Items("Void Wave").Visible = (EntryMode = "E")
                    .Items("Preview").Visible = Not ScreenMode
                    .Items("Deposit").Visible = Not InquiryMode And ScreenMode And (EntryMode = "E")
                End With

            End With
        End If

        chkEmptyWave.Visible = Not ScreenMode

        tabMain.Visible = Not ScreenMode
        splSOTSHIPX.Visible = ScreenMode
        lblPreview.Visible = preview

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        lblWAVE_DATE.Visible = ScreenMode
        dteWAVE_DATE.Visible = ScreenMode
        lblLOCATION_CODE_DEPOSIT.Visible = ScreenMode
        txtLOCATION_CODE_DEPOSIT.Visible = ScreenMode
        lblStatus.Visible = ScreenMode
        lblWAVE_TYPE.Visible = ScreenMode
        optWAVE_TYPE.Visible = ScreenMode

        chkFinalize.Visible = ScreenMode And Not InquiryMode And (EntryMode = "N" Or EntryMode = "E")

        cmdWave.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        chkWaveFromReceiving.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        lblPreferredLocation.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        txtPreferredLocation.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        chkForcePalletPick.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        chkNoUnitPick.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        optPPK.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        cmbLOCATION_USE.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
        lblLOCATION_USE.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")

        If ScreenMode Then

            optPPK.Value = "A"
            grdPOTORDRX.Parent = tabWHTMOVE2.Tabs("PO Shipments").TabPage
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = True
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = True
            grdWHTLOCB1.Parent = tabWHTMOVE2.Tabs("Locations").TabPage
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = True
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = True

            If cmbLOCATION_USE.Value = "" Then cmbLOCATION_USE.Value = "A"

            If WAVE_TYPE = "W" Then
                grdWHTWAVE2.DisplayLayout.MaxBandDepth = 1
                grdWHTWAVE2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
                splSOTSHIPX.Panel1Collapsed = True
                '  grdWHTWAVE2.Text = "Styles"
                SplitContainer1.Panel2Collapsed = False
                tabSOTSHIPX.Tabs("Pick Tickets").Visible = False
                tabSOTSHIPX.Tabs("Pre-Packs").Visible = (EntryMode = "N" Or EntryMode = "E")
                grdWHTWAVE2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdWHTWAVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                With grdWHTWAVE2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("STYLE_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
                    .Columns("STYLE_CODE").ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                    .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("COLOR_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
                    .Columns("COLOR_CODE").ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                    .Columns("PICK_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit

                    .Columns("WAVE_QTY_SUB2").Hidden = True
                    .Columns("WAVE_QTY_CANC").Hidden = True
                    .Columns("WAVE_QTY_CONC").Hidden = True
                    '.Columns("WAVE_QTY_BACK").Hidden = True
                    .Columns("WAVE_QTY_CONF").Hidden = True
                    .Columns("WAVE_QTY_PACK").Hidden = True
                    .Columns("WAVE_QTY_SHIP").Hidden = True
                End With
                tabWHTMOVE2.Tabs("PO Shipments").Visible = False
                tabWHTMOVE2.Tabs("Cartons").Visible = False
            Else
                grdWHTWAVE2.DisplayLayout.MaxBandDepth = 100
                grdWHTWAVE2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
                splSOTSHIPX.Panel1Collapsed = False
                ' grdWHTWAVE2.Text = "Styles"
                SplitContainer1.Panel2Collapsed = True
                tabSOTSHIPX.Tabs("Pick Tickets").Visible = True
                tabSOTSHIPX.Tabs("Pre-Packs").Visible = (EntryMode = "N" Or EntryMode = "E") And WAVE_TYPE = "S"
                grdWHTWAVE2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdWHTWAVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                With grdWHTWAVE2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("STYLE_CODE").Style = UltraWinGrid.ColumnStyle.Default
                    .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("COLOR_CODE").Style = UltraWinGrid.ColumnStyle.Default
                    .Columns("PICK_QTY").CellActivation = UltraWinGrid.Activation.NoEdit

                    .Columns("WAVE_QTY_SUB2").Hidden = False
                    .Columns("WAVE_QTY_CANC").Hidden = False
                    .Columns("WAVE_QTY_CONC").Hidden = False
                    '.Columns("WAVE_QTY_BACK").Hidden = False
                    .Columns("WAVE_QTY_NOTE").Hidden = False
                    .Columns("WAVE_QTY_CONF").Hidden = False
                    .Columns("WAVE_QTY_PACK").Hidden = False
                    .Columns("WAVE_QTY_SHIP").Hidden = False
                End With
                tabWHTMOVE2.Tabs("PO Shipments").Visible = True
                tabWHTMOVE2.Tabs("Cartons").Visible = True
            End If

            If EntryMode = "N" Then
                With grdWHTWAVE2.DisplayLayout.Bands(0)
                    .Columns("WAVE_QTY_PICK").Hidden = True
                    .Columns("WAVE_QTY_ADJ").Hidden = True
                    .Columns("WAVE_QTY_NOTE").Hidden = True
                    .Columns("WAVE_QTY_CONF").Hidden = True
                    .Columns("WAVE_QTY_PACK").Hidden = True
                    .Columns("WAVE_QTY_CANC").Hidden = True
                    .Columns("WAVE_QTY_CONC").Hidden = True
                End With
            End If

            'cmdWave.Visible = (EntryMode = "E")
            'chkWaveFromReceiving.Visible = (EntryMode = "E")

            chkFinalize.Visible = (EntryMode = "E")
            If EntryMode = "E" Then
                ASCMAIN1.sql = "Select WAVE_NO" _
                    & ", Sum (DECODE(WAVE_INST_STATUS,'0',1,0)) WAVED" & vbCrLf _
                    & ", Sum (DECODE(WAVE_INST_STATUS,'1',1,0)) PICKED" & vbCrLf _
                    & ", Sum (DECODE(WAVE_INST_STATUS,'2',1,0)) DEPOSITED" & vbCrLf _
                    & " from WHTINST1 where WAVE_NO = '" & WAVE_NO & "'" & vbCrLf _
                    & " group by WAVE_NO"
                Dim row As DataRow = ASCDATA1.GetDataRow

                If dst.Tables("WHTWAVE2").Select("WAVE_QTY_LEFT > 0").Length <> 0 Then
                    ' If dst.Tables("WHTWAVE2").Select("WAVE_QTY_LEFT <> 0").Length <> 0 Then
                    chkFinalize.Enabled = False
                Else
                    chkFinalize.Enabled = (dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '1'").Length = 0) And (dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '0'").Length = 0) And (dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '2'").Length <> 0)
                    'If WAVE_TYPE = "W" Then
                    '    chkFinalize.Enabled = (Val(row.Item("WAVED") & "") = 0 And Val(row.Item("PICKED") & "") <> 0)
                    'Else
                    '    chkFinalize.Enabled = (Val(row.Item("WAVED") & "") = 0 And Val(row.Item("PICKED") & "") = 0 And Val(row.Item("DEPOSITED") & "") <> 0)
                    'End If
                End If
            End If

            grdWHTWAVE3.Parent = splSOTSHIPX.Panel1
            grdWHTWAVE3.DisplayLayout.GroupByBox.Hidden = True

            With grdWHTINST1.DisplayLayout.Bands(0)
                .Columns("SELECTED").Hidden = Not EntryMode = "V"
            End With

            If EntryMode = "N" Or EntryMode = "E" Then
                If rowWHTWAVE1.Item("WAVE_STATUS") & "" = "F" Then
                    grdWHTINST1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdWHTWAVE2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdWHTWAVE2.DisplayLayout.Bands(1).Override.AllowDelete = DefaultableBoolean.False

                    SplitContainer2.Panel1.Enabled = False

                    UltraExplorerBar1.Groups("Screen Control").Items("Update").Visible = False
                    UltraExplorerBar1.Groups("Screen Control").Items("Delete").Visible = False
                    UltraExplorerBar1.Groups("Screen Control").Items("Deposit").Visible = False

                Else
                    Set_Read_Only_for_ctl(Absx1.txtFor("LOCATION_CODE_DEPOSIT"), False)
                    Set_Read_Only_for_ctl(Absx1.dteFor("WAVE_DATE"), False)
                    Set_Read_Only_for_ctl(Absx1.txtFor("P2L_LINE_ID"), True)

                    If WAVE_TYPE = "L" Then
                        lblLOCATION_CODE_DEPOSIT.Text = "P2L Line"
                        Absx1.txtFor("LOCATION_CODE_DEPOSIT").Visible = False
                    Else
                        lblLOCATION_CODE_DEPOSIT.Text = "Deposit Location"
                        Absx1.txtFor("P2L_LINE_ID").Visible = False
                    End If

                    grdWHTINST1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                    grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdWHTWAVE2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdWHTWAVE2.DisplayLayout.Bands(1).Override.AllowDelete = DefaultableBoolean.True
                End If
            Else
                grdWHTINST1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdWHTWAVE2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdWHTWAVE2.DisplayLayout.Bands(1).Override.AllowDelete = DefaultableBoolean.False

            End If
        Else
            Clear_Record()

            splPicks.Panel2Collapsed = True
            cmdClosePicks.Visible = False
            btnAutomate.Visible = False

            tabSOTSHIPX.Tabs("Pre-Packs").Visible = False

            grdPOTORDRX.Parent = tabShipments.Tabs("PO Shipments").TabPage
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = False
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False

            grdWHTLOCB1.Parent = tabShipments.Tabs("Locations").TabPage
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = False
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False

            grdWHTWAVE3.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdWHTWAVE3.Parent = tabMain.Tabs("Shipments").TabPage
            grdWHTWAVE3.DisplayLayout.GroupByBox.Hidden = False
            Show_Filter(grdSOTSHIPX, True)

            Setup_for_Edit()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"WHTWAVEX", "WHTLOCB1", "WHTWAVE1", "WHTWAVE2", "WHTWAVE2_SUB", "SOTSHIPC", "WHTLOCBW",
                 "WHTWAVE3", "WHTINST1", "WHTINST2", "POTORDRX", "SOTCART1", "SOTCART2", "SOTPICK1", "SOTPICK2", "WHTLOCM1",
                 "WHTMOVE1", "WHTMOVE2", "ICTIADJ1", "ICTIADJ2", "WHTWAVEP", "WHTPPKM2", "WHTSCSEQ"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            If dst.Tables("ICTWHSE1").Rows.Count = 1 Then
                Absx1.txtFor("WHSE_CODE").Text = dst.Tables("ICTWHSE1").Rows(0).Item("WHSE_CODE")
            End If
        End If

        SplitContainer2.Panel1.Enabled = True
        chkWaveFromReceiving.Checked = False
        chkFinalize.Checked = False
        sql_Wave_Filter = " and SOTSHIP1.SHIP_STATUS in ('P')"
        grdSOTSHIPX.Text = "Shipments Released Not Waved"

        'ASCMAIN1.sql = "Truncate Table " & TempTable_BOL_NOs
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        WAVE_TYPE = ""
        sqlPPK_CODEs = ""
        chkNoPPK.Checked = False

        If Not preview Then
            If Not chkNoAutoRefresh.Checked Then
                Load_SOTSHIPX()
            End If
        End If
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            If preview Then
                WAVE_NO = "0000000000"
            Else
                WAVE_NO = ASCMAIN1.Next_Control_No("WHTWAVE1.WAVE_NO")
            End If

            rowWHTWAVE1 = dst.Tables("WHTWAVE1").NewRow
            rowWHTWAVE1.Item("WAVE_NO") = WAVE_NO
            rowWHTWAVE1.Item("WAVE_DATE") = DATETIME_STAMP.Date

            rowWHTWAVE1.Item("CUST_CODE") = CUST_CODE
            rowWHTWAVE1.Item("WHSE_CODE") = WHSE_CODE
            rowWHTWAVE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTWAVE1.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTWAVE1.Item("WAVE_STATUS") = "O"

            rowWHTWAVE1.Item("WAVE_TYPE") = WAVE_TYPE

            If WAVE_TYPE = "L" Then
                rowWHTWAVE1.Item("P2L_WAVE_STATUS") = "P"
                rowWHTWAVE1.Item("P2L_LINE_ID") = Absx1.txtFor("P2L_LINE_ID").Text
                rowWHTWAVE1.Item("LOCATION_CODE_DEPOSIT") = Absx1.txtFor("LOCATION_CODE_DEPOSIT").Text
            End If

            dst.Tables("WHTWAVE1").Rows.Add(rowWHTWAVE1)

            If WAVE_TYPE = "W" Then
                sqlSHIP_BOL_NOs = ""
                sqlSHIP_BOL_NOs_2 = ""
                dst.Tables("WHTWAVE2").Rows.Clear()
                dst.Tables("WHTWAVE3").Rows.Clear()
                dst.Tables("POTORDRX").Rows.Clear()
                dst.Tables("WHTWAVE3").Rows.Clear()

            Else
                Get_SHIP_BOL_NOs()

                ASCMAIN1.sql = "Select '" & WAVE_NO & "' WAVE_NO,SOTSHIPX.*" & vbCrLf _
                    & ", " & IIf(WAVE_TYPE = "L", "'O'", "''") & " P2L_SHIP_STATUS" & vbCrLf _
                    & " from " & SOTSHIPX & " SOTSHIPX" & vbCrLf _
                    & " where SOTSHIPX.SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NOs, 2) & ")" & vbCrLf
                ' & IIf(sqlSHIP_BOL_NOs_2 <> "", " Or SOTSHIPX.SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NOs_2, 2) & ")", "")
                Fill_Records("WHTWAVE3", "", True, ASCMAIN1.sql)

                ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                    & ", Sum (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
                    & " from SOTORDR1,SOTORDR2,SOTPICK2,SOTPICK1" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                    & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                    & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                    & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                    & "   and SOTPICK1.SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NOs, 2) & ")" & vbCrLf _
                    & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
                '& IIf(sqlSHIP_BOL_NOs_2 <> "", " Or SOTPICK1.SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NOs_2, 2) & ")", "") & vbCrLf _
                ASCMAIN1.sql = "Select '" & WAVE_NO & "' WAVE_NO, ROWNUM WAVE_LNO" & vbCrLf _
                    & ", STYLE_CODE, COLOR_CODE, PICK_QTY from (Select * from (" & ASCMAIN1.sql & ")" & vbCrLf _
                    & " order by STYLE_CODE, COLOR_CODE)"
                Fill_Records("WHTWAVE2", "", True, ASCMAIN1.sql)
            End If

            WAVE_LNO_ctr = Val(dst.Tables("WHTWAVE2").Compute("MAX(WAVE_LNO)", "") & "")

            rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)


            If WAVE_TYPE = "L" Or (WAVE_TYPE = "W" And ASCMAIN1.Running_in_VS) Then ' temporary "w" for first wave - set CUST_CODE
                Dim STYLE_CODE As String
                Dim COLOR_CODE As String
                Fill_Records("WHTSCSEQ", CUST_CODE)
                For Each rowWHTSCSEQ As DataRow In dst.Tables("WHTSCSEQ").Select("PO_QTY_PER_CTN is not null")
                    STYLE_CODE = rowWHTSCSEQ.Item("STYLE_CODE")
                    COLOR_CODE = rowWHTSCSEQ.Item("COLOR_CODE")
                    If dst.Tables("WHTWAVE2").Compute("count(STYLE_CODE)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") = 0 Then
                        Dim rowWHTWAVE2 As DataRow = dst.Tables("WHTWAVE2").NewRow
                        WAVE_LNO_ctr += 1
                        rowWHTWAVE2.Item("WAVE_NO") = WAVE_NO
                        rowWHTWAVE2.Item("WAVE_LNO") = WAVE_LNO_ctr
                        rowWHTWAVE2.Item("STYLE_CODE") = STYLE_CODE
                        rowWHTWAVE2.Item("COLOR_CODE") = COLOR_CODE
                        rowWHTWAVE2.Item("PICK_QTY") = 0
                        dst.Tables("WHTWAVE2").Rows.Add(rowWHTWAVE2)
                    End If
                Next
                dst.Tables("WHTWAVE2").AcceptChanges()
            End If


            If WAVE_TYPE = "W" Or WAVE_TYPE = "L" Then
            Else
                Create_Wave_Instructions(True)
            End If

        Else
            WAVE_NO = Absx1.txtFor("WAVE_NO").Text
            rowWHTWAVE1 = Fill_Record("WHTWAVE1", WAVE_NO)
            WAVE_TYPE = rowWHTWAVE1.Item("WAVE_TYPE")

            WHSE_CODE = rowWHTWAVE1.Item("WHSE_CODE")
            rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
            CUST_CODE = rowWHTWAVE1.Item("CUST_CODE") & ""

            If WAVE_TYPE = "W" Then
            Else
                Load_SOTSHIPX_data(" and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from WHTWAVE3 where WAVE_NO = '" & WAVE_NO & "')")
                Fill_Records("WHTWAVE3", WAVE_NO)
                Get_SHIP_BOL_NOs()
            End If

            Fill_Records("WHTWAVE2", WAVE_NO)

            WAVE_LNO_ctr = Val(dst.Tables("WHTWAVE2").Compute("MAX(WAVE_LNO)", "") & "")
            dst.Tables("WHTWAVE2_SUB").Rows.Clear()

            For Each row As DataRow In dst.Tables("WHTWAVE2").Select("STYLE_CODE_SUB is Not Null")
                Dim rowSUB As DataRow = dst.Tables("WHTWAVE2_SUB").NewRow
                For Each COLUMN_NAME As String In New String() _
                    {"WAVE_NO", "WAVE_LNO", "STYLE_CODE", "COLOR_CODE", "WAVE_QTY", "WAVE_QTY_PICK",
                     "STYLE_CODE_SUB", "COLOR_CODE_SUB", "WAVE_QTY_SUB", "WAVE_QTY_ADJ", "WAVE_QTY_NOTE"}
                    rowSUB.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                Dim sqlw As String = "STYLE_CODE = '" & row.Item("STYLE_CODE") & "' and COLOR_CODE = '" & row.Item("COLOR_CODE") & "' and STYLE_CODE_SUB is Null"
                Dim row2 As DataRow = dst.Tables("WHTWAVE2").Select(sqlw)(0)
                Dim WAVE_LNO_LINK As Int64 = Val(row2.Item("WAVE_LNO") & "")
                rowSUB.Item("WAVE_LNO_LINK") = WAVE_LNO_LINK
                dst.Tables("WHTWAVE2_SUB").Rows.Add(rowSUB)
            Next
            ASCDATA1.DeleteRows("WHTWAVE2", "STYLE_CODE_SUB is Not Null")

            Manage_Expressions("Remove")
            Fill_Records("WHTINST1", WAVE_NO)
            Fill_Records("WHTINST2", WAVE_NO)
            Manage_Expressions("Restore")
            dst.AcceptChanges()
        End If

        WAVE_INST_NO_void.Clear()

        EnforceConstraints(True)

        grdWHTINST1.Text = "Wave Pick Instructions"
        Sort_grdColumns(grdWHTINST1, "WAVE_INST_NO")

        dst.Tables("POTORDRX").Rows.Clear()
        If WAVE_TYPE = "L" Or (WAVE_TYPE = "W" And ASCMAIN1.Running_in_VS) Then ' temporary for first wave
            Fill_Records("WHTWAVEP2L", CUST_CODE)
        End If

        For Each rowWHTWAVE2 As DataRow In dst.Tables("WHTWAVE2").Select("")
            Dim STYLE_CODE As String = rowWHTWAVE2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTWAVE2.Item("COLOR_CODE")

            Debug.Print("A" & ":" & STYLE_CODE & ":" & COLOR_CODE & ":" & Now)

            ASCMAIN1.sql = "Select * from (" & vbCrLf _
                & "Select POTSHIP1.WHSE_CODE, POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
                & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO, POTSHIP3.PO_QTY_SHP" & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
                & ", POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE" & vbCrLf _
                & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2 " & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO " & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
                & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                & ")"
            ' WHERE WHSE_CODE = 'NJE'
            Fill_Records("POTORDRX", "", False, ASCMAIN1.sql)

            Dim WHSE_QTY_SHIP As Int64 = Val(dst.Tables("POTORDRX").Compute("SUM(PO_QTY_SHP)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & "")
            If WHSE_QTY_SHIP <> 0 Then
                rowWHTWAVE2.Item("WAVE_QTY_SHIP") = WHSE_QTY_SHIP
            End If

            'Doug request 8/17/20 email not to remove previous waves on qty oh, same below
            Dim rowWHSE As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            ASCMAIN1.sql = "Select Sum (NVL(WHTLOCB1.LOCATION_QTY,0))" & vbCrLf _
                & " from WHTLOCB1,WHTLOCM1" & vbCrLf _
                & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                & "   and WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and WHTLOCB1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & "   and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
                & "   and (NVL(WHTLOCB1.LOCATION_QTY,0)) > 0" & vbCrLf _
                & "   and (NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1' or WHTLOCB1.LOCATION_CODE in (" & REC_LOCATIONS & ",'" & rowWHSE.Item("WHSE_LOC_LNF") & "')) "
            Dim WAVE_QTY_LOCS As Int64 = Val(ASCDATA1.GetDataValue() & "")
            If WAVE_QTY_LOCS <> 0 Then rowWHTWAVE2.Item("WAVE_QTY_LOCS") = WAVE_QTY_LOCS

            If IsWalmart(CUST_CODE) Then
                ASCMAIN1.sql = "Select Sum (NVL(WHTLOCB1.LOCATION_QTY,0))" & vbCrLf _
                    & " from WHTLOCB1,WHTLOCM1" & vbCrLf _
                    & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                    & "   and WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & "   and WHTLOCB1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                    & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                    & "   and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
                    & "   and (NVL(WHTLOCB1.LOCATION_QTY,0)) > 0" & vbCrLf _
                    & "   and (NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1' or WHTLOCB1.LOCATION_CODE in (" & REC_LOCATIONS & ",'" & rowWHSE.Item("WHSE_LOC_LNF") & "')) " & vbCrLf _
                    & "   and WHTLOCM1.LOCATION_CODE between '05' and '40'"
                Dim WAVE_QTY_RACS As Int64 = Val(ASCDATA1.GetDataValue() & "")
                If WAVE_QTY_RACS <> 0 Then rowWHTWAVE2.Item("WAVE_QTY_RACS") = WAVE_QTY_RACS

                ASCMAIN1.sql = "Select (WHSE_QTY_ON_HAND - WHSE_QTY_PICK) from ICTSTAT2 WHERE WHSE_CODE = :PARM1 AND STYLE_CODE = :PARM2 AND COLOR_CODE = :PARM3"
                Dim WAVE_QTY_OTS As Int64 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "VVV", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE}) & "")
                If WAVE_QTY_OTS <> 0 Then rowWHTWAVE2.Item("WAVE_QTY_OTS") = WAVE_QTY_OTS

            End If
            If WAVE_TYPE = "L" Or (WAVE_TYPE = "W" And ASCMAIN1.Running_in_VS) Then ' temporary "W" for first work order
                For Each rowWHTWAVEP2L As DataRow In dst.Tables("WHTWAVEP2L").Select("STYLE_CODE='" & STYLE_CODE & "' and COLOR_CODE='" & COLOR_CODE & "'")
                    rowWHTWAVE2.Item("P2L_QTY_OH") = rowWHTWAVEP2L.Item("P2L_QTY_OH")
                    rowWHTWAVE2.Item("P2L_QTY_COMMITED") = rowWHTWAVEP2L.Item("P2L_QTY_COMMITED")
                    rowWHTWAVE2.Item("P2L_QTY_RESERVE") = rowWHTWAVEP2L.Item("P2L_QTY_RESERVE")
                    rowWHTWAVE2.Item("P2L_WO_OPEN") = rowWHTWAVEP2L.Item("P2L_WO_OPEN") ' should subtract this wave's contribution on Edit
                    rowWHTWAVE2.Item("P2L_WO_PICK") = rowWHTWAVEP2L.Item("P2L_WO_PICK") ' should subtract this wave's contribution on Edit
                    rowWHTWAVE2.Item("P2L_QTY_NOT_INDUCTED") = rowWHTWAVEP2L.Item("P2L_QTY_NOT_INDUCTED")
                Next
            End If
        Next

        Debug.Print("B" & ":" & Now)

        ASCMAIN1.sql = "Select WHTINST1.WAVE_NO, WHTINST1.WAVE_SUB, WHTINST2.WAVE_LNO" & vbCrLf _
            & ", WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE, WHTINST1.WAVE_INST_STATUS" & vbCrLf _
            & "   , Sum (WHTINST2.LOCATION_QTY_WAVE) WAVE_QTY_WAVE" & vbCrLf _
            & "   , Sum (WHTINST2.LOCATION_QTY_PICK) WAVE_QTY_PICK" & vbCrLf _
            & "   from WHTINST2,WHTINST1" & vbCrLf _
            & "   where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
            & "     and WHTINST1.WAVE_NO = '" & WAVE_NO & "'" & vbCrLf _
            & "   group by WHTINST1.WAVE_NO, WHTINST1.WAVE_SUB, WHTINST2.WAVE_LNO" & vbCrLf _
            & ", WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE, WHTINST1.WAVE_INST_STATUS"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim WAVE_INST_STATUS As String = row.Item("WAVE_INST_STATUS")
            Dim WAVE_LNO As Int64 = Val(row.Item("WAVE_LNO") & "")
            Dim WAVE_SUB As String = row.Item("WAVE_SUB")


            Debug.Print("C" & ":" & STYLE_CODE, COLOR_CODE, Now)

            Dim rowWHTWAVE2 As DataRow = Nothing

            If WAVE_SUB = "1" AndAlso dst.Tables("WHTWAVE2_SUB").Select("WAVE_NO = '" & WAVE_NO & "' and WAVE_LNO = " & CStr(WAVE_LNO)).Length > 0 Then ' THIS IS A SUB INSTRUCTION, UPDATE WHTWAVE2_SUB
                'Dim rowWHTWAVE2_orig() As DataRow = dst.Tables("WHTWAVE2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                'Dim WAVE_LNO_LINK As Integer = Val(rowWHTWAVE2_orig(0).Item("WAVE_LNO") & "") ' there should be only 1 row that matches this
                'rowWHTWAVE2 = dst.Tables("WHTWAVE2_SUB").Rows.Find(New String() {WAVE_NO, WAVE_LNO_LINK, WAVE_LNO})
                rowWHTWAVE2 = dst.Tables("WHTWAVE2_SUB").Select("WAVE_NO = '" & WAVE_NO & "' and WAVE_LNO = " & CStr(WAVE_LNO))(0)
            Else
                rowWHTWAVE2 = dst.Tables("WHTWAVE2").Rows.Find(New String() {WAVE_NO, WAVE_LNO})
                ' rowWHTWAVE2 = dst.Tables("WHTWAVE2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0)
            End If
            If rowWHTWAVE2 IsNot Nothing Then
                If WAVE_INST_STATUS = "0" Then
                    rowWHTWAVE2.Item("WAVE_QTY_OPEN") = Val(rowWHTWAVE2.Item("WAVE_QTY_OPEN") & "") + Val(row.Item("WAVE_QTY_WAVE") & "")
                ElseIf WAVE_INST_STATUS = "V" Then
                    ' SKIP - WAVE REVERSED
                Else
                    rowWHTWAVE2.Item("WAVE_QTY_DIFF") = Val(rowWHTWAVE2.Item("WAVE_QTY_DIFF") & "") + (Val(row.Item("WAVE_QTY_PICK") & "") - Val(row.Item("WAVE_QTY_WAVE") & ""))
                End If
            End If

        Next

        Sort_grdColumns(grdWHTWAVE2, "STYLE_CODE,COLOR_CODE")


        Debug.Print("D" & ":" & Now)

        Manage_Expressions("Remove")
        For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("")
            SHIP_BOL_NO = rowWHTWAVE3.Item("SHIP_BOL_NO")

            Debug.Print("E" & ":" & SHIP_BOL_NO & ":" & Now)
            Fill_Records("SOTCART1", SHIP_BOL_NO, False)
            Fill_Records("SOTCART2", SHIP_BOL_NO, False)
            Fill_Records("SOTPICK1", SHIP_BOL_NO, False)
            Fill_Records("SOTPICK2", SHIP_BOL_NO, False)
        Next

        Manage_Expressions("Restore")

        If WAVE_TYPE = "W" Then
            Fill_Records("WHTSCSEQ", , True, "SELECT * FROM WHTSCSEQ")
        Else
            If IsWalmart(CUST_CODE) Then
                Fill_Records("WHTSCSEQ", , True, "SELECT * FROM WHTSCSEQ where CUST_CODE in (" & WalmartCodes & ")")
            Else
                Fill_Records("WHTSCSEQ", CUST_CODE, True)
            End If
        End If

        If EntryMode = "N" Then
            lblStatus.Text = "New Wave"
        Else
            Select Case rowWHTWAVE1.Item("WAVE_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "V"
                    lblStatus.Text = "Voided"
                Case "D"
                    lblStatus.Text = "Deleted"
                Case "F"
                    lblStatus.Text = "Finalized"
                Case Else
                    lblStatus.Text = "Status Unknown"
            End Select
        End If


        Debug.Print("F" & ":" & Now)

        If WAVE_TYPE = "W" Then
            tabSOTSHIPX.Visible = True
        Else
            Setup_grdSOTSHIPX()
        End If

        If InquiryMode Then
            ASCMAIN1.sql = "Select * from ICTIADJ1 where ADJ_SOURCE = 'W' and ADJ_REF = '" & WAVE_NO & "'"
            Fill_Records("ICTIADJ1", "", True, ASCMAIN1.sql)
            ASCMAIN1.sql = "Select * from ICTIADJ2 where ADJ_NO in (Select ADJ_NO from ICTIADJ1 where ADJ_SOURCE = 'W' and ADJ_REF = '" & WAVE_NO & "')"
            Fill_Records("ICTIADJ2", "", True, ASCMAIN1.sql)
        End If


        Display_Totals()

        grdWHTWAVE2.Rows.ExpandAll(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, WAVE_NO)
        For Each TABLE_NAME As String In New String() _
            {"WHTWAVE1", "WHTWAVE2", "WHTWAVE3"}
            Delete_Records_1(TABLE_NAME)
        Next

        Dim SQLD As String = " where WAVE_INST_NO in (Select WAVE_INST_NO from WHTINST1 where WAVE_NO = '" & WAVE_NO & "' and WAVE_INST_STATUS = '0')"
        ASCDATA1.ExecuteSQL("Delete from WHTINST2" & SQLD)
        ASCDATA1.ExecuteSQL("Delete from WHTINST1" & SQLD)
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where WAVE_NO = '" & WAVE_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Void_Wave()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()

        Dependent_Updates(-1, WAVE_NO)

        ASCDATA1.ExecuteSQL("Update WHTWAVE1 set WAVE_STATUS = 'V' where WAVE_NO = '" & WAVE_NO & "'")
        ASCDATA1.ExecuteSQL("Update WHTINST1 set WAVE_INST_STATUS = 'C' where WAVE_NO = '" & WAVE_NO & "' and WAVE_INST_STATUS = '0'")

        CommitTrans("Void Wave Complete")
    End Sub

    Sub Dependent_Updates(S As Integer, WAVE_NO As String)

        Dim SS As String = "+1 *"
        If S = -1 Then SS = "-1 *"

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select WHTINST2.*,WHTINST1.WAVE_PICK_TYPE, WHTINST1.LOCATION_CODE" & vbCrLf _
            & "   from WHTINST1,WHTINST2" & vbCrLf _
            & "   where WHTINST1.WAVE_NO = '" & WAVE_NO & "'" & vbCrLf _
            & "     and WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
            & "     and WHTINST1.WAVE_INST_STATUS = '0';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "    Update WHTLOCB1 Set LOCATION_QTY_WAVE = NVL(LOCATION_QTY_WAVE,0) " & SS & " NVL(R1.LOCATION_QTY_WAVE,0)" & vbCrLf _
            & "     where WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "       and LOCATION_CODE = R1.LOCATION_CODE" & vbCrLf _
            & "       and BAR_CODE = R1.BAR_CODE" & vbCrLf _
            & "       and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Dim SHIP_WAVE_STATUS As String = ""
        Dim sqlWAVE As String = ""
        If S = 1 Then
            SHIP_WAVE_STATUS = "1"
            sqlWAVE = ", WAVE_NO = '" & WAVE_NO & "'"
        Else
            SHIP_WAVE_STATUS = "0"
            sqlWAVE = ", WAVE_NO = NULL"
        End If
        ASCMAIN1.sql = "Update SOTSHIP1" & vbCrLf _
            & " Set SHIP_WAVE_STATUS = '" & SHIP_WAVE_STATUS & "'" & sqlWAVE & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from WHTWAVE3 where WAVE_NO = '" & WAVE_NO & "')"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        If EntryMode = "N" Then
            Dim LOAD_NO_DEPOSIT As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
            Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
            With rowWHTBARC0
                .Item("LOAD_NO") = LOAD_NO_DEPOSIT
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LOAD_STATUS") = "A"
                .Item("LOAD_COMMENT") = "LOAD FOR DEPOSIT"
                .Item("LOCATION_CODE") = rowWHTWAVE1.Item("LOCATION_CODE_DEPOSIT")
                .Item("TRAN_TYPE") = "L"
                .Item("TRAN_NO") = WAVE_NO
                .Item("LOAD_DATE") = DATETIME_STAMP.Date
            End With
            dst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
            Update_Record_TDA("WHTBARC0")
            rowWHTWAVE1.Item("LOAD_NO_DEPOSIT") = LOAD_NO_DEPOSIT

        Else
            Delete_Records()
        End If

        EnforceConstraints(False)
        For Each rowSUB As DataRow In dst.Tables("WHTWAVE2_SUB").Select("")
            Dim row As DataRow = dst.Tables("WHTWAVE2").NewRow
            For Each COLUMN_NAME As String In New String() _
                {"WAVE_NO", "WAVE_LNO", "STYLE_CODE", "COLOR_CODE", "WAVE_QTY", "WAVE_QTY_PICK",
                 "STYLE_CODE_SUB", "COLOR_CODE_SUB", "WAVE_QTY_SUB", "WAVE_QTY_ADJ", "WAVE_QTY_NOTE"}
                row.Item(COLUMN_NAME) = rowSUB.Item(COLUMN_NAME)
            Next

            dst.Tables("WHTWAVE2").Rows.Add(row)
        Next

        Debug.Print("A" & ":" & Now)


        INIT_LAST("WHTWAVE1", False, , True)
        Dim sqldelete As String = "WAVE_NO = '" & WAVE_NO & "'"
        Update_Record_TDA("WHTWAVE1", sqldelete)
        Update_Record_TDA("WHTWAVE2", sqldelete)
        Update_Record_TDA("WHTWAVE3", sqldelete)

        'If WAVE_TYPE = "L" And EntryMode = "N" Then
        '    ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " & vbCrLf _
        '                    & " SELECT SOTPICK2.*, SOTCART1.CART_NO" & vbCrLf _
        '                    & " FROM  SOTPICK1,SOTPICK2,SOTCART1,WHTWAVE3" & vbCrLf _
        '                    & " WHERE SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
        '                    & " AND WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
        '                    & " AND SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
        '                    & " AND SOTPICK2.PICK_NO = SOTPICK1.PICK_NO;" & vbCrLf _
        '                    & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
        '                    & " UPDATE SOTCART2 SET QTY_REL = R1.PICK_QTY WHERE CART_NO = R1.CART_NO AND ORDR_NO = R1.ORDR_NO AND ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
        '                    & " END LOOP;END; END;"
        '    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", WAVE_NO)

        '    ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " & vbCrLf _
        '                    & " SELECT SOTPICK1.PICK_NO, sum (PICK_QTY) PICK_QTY, sum (PICK_QTY_CONF) PICK_QTY_CONF " & vbCrLf _
        '                    & " FROM  SOTPICK1,SOTPICK2" & vbCrLf _
        '                    & " WHERE SOTPICK1.SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM WHTWAVE3 WHERE WAVE_NO = :PARM1)" & vbCrLf _
        '                    & " AND SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
        '                    & " GROUP BY SOTPICK1.PICK_NO;" & vbCrLf _
        '                    & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
        '                    & " UPDATE SOTCART1 SET CART_TOTAL_UNITS_REL = R1.PICK_QTY WHERE PICK_NO = R1.PICK_NO;" & vbCrLf _
        '                    & " END LOOP; END; END;"
        '    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", WAVE_NO)
        'End If

        Debug.Print("B" & ":" & Now)


        If WAVE_TYPE = "S" Then
            Dim only_in_pick As Boolean = True
            For Each row As DataRow In dst.Tables("WHTWAVE3").Select("")
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
                Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                If rowSOTSHIP1.Item("SHIP_STATUS") <> "P" Then
                    only_in_pick = False
                    Exit For
                End If
            Next
            If only_in_pick Then
                Update_Record_TDA("SOTCART2")
                ' verify that naseema is doing carton adjustments
            End If
        End If

        Debug.Print("C" & ":" & Now)

        If dst.Tables("WHTINST2").Rows.Count > 0 Then
            dst.Tables("WHTINST2").AcceptChanges()
            For Each row As DataRow In dst.Tables("WHTINST2").Select("")
                row.SetAdded()
            Next

            ASCMAIN1.sql = "delete WHTINST2 where WAVE_INST_NO in (Select WAVE_INST_NO from WHTINST1 where " & sqldelete & ")"
            ASCDATA1.ExecuteSQL()

            Create_BAs("WHTINST2", True)
            Update_BAs("WHTINST2", True)
        End If
        'Update_Record_TDA("WHTINST2", "WAVE_INST_NO in (Select WAVE_INST_NO from WHTINST1 where " & sqldelete & ")")
        Update_Record_TDA("WHTINST1", sqldelete)
        Debug.Print("D" & ":" & Now)

        Dependent_Updates(1, WAVE_NO)
        Debug.Print("E" & ":" & Now)

        If WAVE_INST_NO_void.Count <> 0 Then
            For Each WAVE_INST_NO As String In WAVE_INST_NO_void
                ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                                   New Object() {"V", WAVE_INST_NO, ASCMAIN1.SESSION_NO},
                                   New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
                'ASCMAIN1.sql = "Update WHTINST2 Set LOCATION_QTY_WAVE = -1 * LOCATION_QTY_WAVE, LOCATION_QTY_PICK = -1 * LOCATION_QTY_PICK where WAVE_INST_NO = '" & WAVE_INST_NO & "'"
                'ASCDATA1.ExecuteSQL()
            Next
        End If
        WAVE_INST_NO_void.Clear()

        If chkFinalize.Checked Then
            rowWHTWAVE1.Item("WAVE_STATUS") = "F"
            Update_Record_TDA("WHTWAVE1", sqldelete) ' 2nd Update to WHTWAVE1

            If WAVE_TYPE = "W" Or WAVE_TYPE = "L" Then
                ' deposits stay in stage
                ' Pick To light treated like:
                'wave finalization
                ' - calc qtys to move to shipping (P2L wave)
                ' - use cartons Not wave instructions
                '00000000 bar code
                'a) TAG LINES AS 0 CARTON LOCATIONS
                'b) fix the data
            Else

                Dim ORDR_CUST_POs As String = ""
                For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("")
                    ORDR_CUST_POs &= ";" & rowWHTWAVE3.Item("ORDR_CUST_PO")
                Next
                ORDR_CUST_POs = Mid(ORDR_CUST_POs, 2)

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

                ' deposited -> shipping
                ' Move everything that was Picked and Deposited into Shipping
                ' Important that everything was Deposited before doing this

                Dim WHSE_TRAN_LNO As Integer
                For Each rowWHTINST1 As DataRow In dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '2'")
                    For Each rowWHTINST2 As DataRow In rowWHTINST1.GetChildRows("WHTINST1_WHTINST2")
                        Dim LOCATION_QTY_PICK As Int64 = Val(rowWHTINST2.Item("LOCATION_QTY_PICK") & "")
                        If LOCATION_QTY_PICK <> 0 Then

                            Dim BAR_CODE As String = IIf(rowWHTINST1.Item("BAR_CODE_OTHER") & "" <> "", rowWHTINST1.Item("BAR_CODE_OTHER"), rowWHTINST2.Item("BAR_CODE") & "")
                            Dim rowWHTBARC1 As DataRow = Fill_Record("WHTBARC1", BAR_CODE, , False)
                            If IsNothing(rowWHTBARC1) Then
                                MsgBox("Cannot find Carton LPN " & BAR_CODE & ", Wave No " & WAVE_NO & ". Inventory Out of Balance warning", MsgBoxStyle.Exclamation, "Error, Contact ABS(Rick)")
                            End If

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
                                .Item("STYLE_CODE") = rowWHTINST2.Item("STYLE_CODE")
                                .Item("COLOR_CODE") = rowWHTINST2.Item("COLOR_CODE")
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
                Next

                Update_Record_TDA("WHTMOVE1")
                Update_Record_TDA("WHTMOVE2")

                ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                                   New Object() {WHSE_TRAN_NO, 0, 1},
                                   New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})



                ' Concealed Shortages discovered in cartons after pick

                If dst.Tables("WHTWAVE2").Select("WAVE_QTY_ADJ <> 0").Length <> 0 Then
                    Dim ADJ_NO As String = Add_ICTIADJ1(Mid("CS " & CUST_CODE & ":" & ORDR_CUST_POs, 1, 200), ROWs("WHTPARM1").Item("WH_PARM_CS_PICK"))

                    'Dim ADJ_NO As String = ""
                    'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    '    ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
                    'Else
                    '    ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
                    'End If

                    Dim ADJ_LNO As Int64 = 0
                    Dim TOTAL_COSTS As Decimal = 0

                    'Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow
                    'With rowICTIADJ1
                    '    .Item("ADJ_NO") = ADJ_NO
                    '    .Item("ADJ_DATE") = DATETIME_STAMP.Date
                    '    .Item("WHSE_CODE") = WHSE_CODE
                    '    .Item("REASON_CODE") = ROWs("WHTPARM1").Item("WH_PARM_CS_PICK")
                    '    .Item("ADJ_NOTE") = Mid("CS " & CUST_CODE & ":" & ORDR_CUST_POs, 1, 200)
                    '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    '    .Item("INIT_DATE") = DATETIME_STAMP
                    '    .Item("REGISTER_IND") = "0"
                    '    .Item("ADJ_SOURCE") = "W"
                    '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    '    .Item("TOTAL_COSTS") = TOTAL_COSTS
                    '    .Item("ADJ_REF") = WAVE_NO
                    'End With
                    'dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

                    For Each rowWHTWAVE2 As DataRow In dst.Tables("WHTWAVE2").Select("WAVE_QTY_ADJ <> 0")
                        Dim STYLE_CODE As String = rowWHTWAVE2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowWHTWAVE2.Item("COLOR_CODE")
                        Dim ADJ_QTY As Int64 = Val(rowWHTWAVE2.Item("WAVE_QTY_ADJ") & "")

                        Add_ICTIADJ2(ADJ_NO, ADJ_LNO, STYLE_CODE, COLOR_CODE, ADJ_QTY, "CS PICK")

                        'Dim rowICTIADJ2 As DataRow = dst.Tables("ICTIADJ2").NewRow
                        'With rowICTIADJ2
                        '    .Item("ADJ_NO") = ADJ_NO
                        '    ADJ_LNO += 1
                        '    .Item("ADJ_LNO") = ADJ_LNO
                        '    .Item("STYLE_CODE") = STYLE_CODE
                        '    .Item("COLOR_CODE") = COLOR_CODE
                        '    .Item("ADJ_QTY") = ADJ_QTY
                        '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        '    .Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
                        '    .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                        '    .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                        '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                        '    .Item("LOCATION_CODE") = rowICTWHSE1.Item("WHSE_LOC_SHP")
                        '    .Item("BAR_CODE") = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")
                        '    .Item("ADJ_REF") = "CS PICK"
                        'End With
                        'dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
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



                ' Concealed Shortages passed onto customer

                If dst.Tables("WHTWAVE2").Select("WAVE_QTY_CONC <> 0").Length <> 0 Then
                    Dim ADJ_NO As String = Add_ICTIADJ1(Mid("CS " & CUST_CODE & ":" & ORDR_CUST_POs, 1, 200), ROWs("WHTPARM1").Item("WH_PARM_CS_SHIP"))

                    'Dim ADJ_NO As String = ""
                    'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    '    ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
                    'Else
                    '    ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
                    'End If

                    Dim ADJ_LNO As Int64 = 0
                    Dim TOTAL_COSTS As Decimal = 0

                    'Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow
                    'With rowICTIADJ1
                    '    .Item("ADJ_NO") = ADJ_NO
                    '    .Item("ADJ_DATE") = DATETIME_STAMP.Date
                    '    .Item("WHSE_CODE") = WHSE_CODE
                    '    .Item("REASON_CODE") = ROWs("WHTPARM1").Item("WH_PARM_CS_SHIP")
                    '    .Item("ADJ_NOTE") = Mid("CS " & CUST_CODE & ":" & ORDR_CUST_POs, 1, 200)
                    '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    '    .Item("INIT_DATE") = DATETIME_STAMP
                    '    .Item("REGISTER_IND") = "0"
                    '    .Item("ADJ_SOURCE") = "W"
                    '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    '    .Item("TOTAL_COSTS") = TOTAL_COSTS
                    '    .Item("ADJ_REF") = WAVE_NO
                    'End With
                    'dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

                    For Each rowWHTWAVE2 As DataRow In dst.Tables("WHTWAVE2").Select("WAVE_QTY_CONC <> 0")
                        Dim STYLE_CODE As String = rowWHTWAVE2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowWHTWAVE2.Item("COLOR_CODE")
                        Dim ADJ_QTY As Int64 = Val(rowWHTWAVE2.Item("WAVE_QTY_CONC") & "")

                        Add_ICTIADJ2(ADJ_NO, ADJ_LNO, STYLE_CODE, COLOR_CODE, ADJ_QTY, "CS SHIP")

                        'Dim rowICTIADJ2 As DataRow = dst.Tables("ICTIADJ2").NewRow
                        'With rowICTIADJ2
                        '    .Item("ADJ_NO") = ADJ_NO
                        '    ADJ_LNO += 1
                        '    .Item("ADJ_LNO") = ADJ_LNO
                        '    .Item("STYLE_CODE") = STYLE_CODE
                        '    .Item("COLOR_CODE") = COLOR_CODE
                        '    .Item("ADJ_QTY") = ADJ_QTY
                        '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        '    .Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
                        '    .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                        '    .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                        '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                        '    .Item("LOCATION_CODE") = rowICTWHSE1.Item("WHSE_LOC_SHP")
                        '    .Item("BAR_CODE") = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")
                        '    .Item("ADJ_REF") = "CS SHIP"
                        'End With
                        'dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
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


                ' Substitutions - these rows should all be the sub rows

                If dst.Tables("WHTWAVE2").Select("WAVE_QTY_SUB <> 0").Length <> 0 Then
                    Dim ADJ_NO As String = Add_ICTIADJ1("Subs " & Mid(CUST_CODE & ":" & ORDR_CUST_POs, 1, 200), ROWs("WHTPARM1").Item("WH_PARM_CS_PICK"))

                    'Dim ADJ_NO As String = ""
                    'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    '    ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
                    'Else
                    '    ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
                    'End If
                    Dim ADJ_LNO As Int64 = 0
                    Dim TOTAL_COSTS As Decimal = 0

                    'Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow
                    'With rowICTIADJ1
                    '    .Item("ADJ_NO") = ADJ_NO
                    '    .Item("ADJ_DATE") = DATETIME_STAMP.Date
                    '    .Item("WHSE_CODE") = WHSE_CODE
                    '    .Item("REASON_CODE") = ROWs("WHTPARM1").Item("WH_PARM_CS_PICK")
                    '    .Item("ADJ_NOTE") = "Subs " & Mid(CUST_CODE & ":" & ORDR_CUST_POs, 1, 200)
                    '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    '    .Item("INIT_DATE") = DATETIME_STAMP
                    '    .Item("REGISTER_IND") = "0"
                    '    .Item("ADJ_SOURCE") = "W"
                    '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    '    .Item("TOTAL_COSTS") = TOTAL_COSTS
                    '    .Item("ADJ_REF") = WAVE_NO
                    'End With
                    'dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

                    For Each rowWHTWAVE2 As DataRow In dst.Tables("WHTWAVE2").Select("WAVE_QTY_SUB <> 0")
                        Dim STYLE_CODE As String = rowWHTWAVE2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowWHTWAVE2.Item("COLOR_CODE")
                        Dim ADJ_QTY As Int64 = Val(rowWHTWAVE2.Item("WAVE_QTY_SUB") & "")
                        Dim ADJ_REF As String = "SUB FOR"
                        For r As Integer = 0 To 1
                            Add_ICTIADJ2(ADJ_NO, ADJ_LNO, STYLE_CODE, COLOR_CODE, ADJ_QTY, ADJ_REF)

                            'Dim rowICTIADJ2 As DataRow = dst.Tables("ICTIADJ2").NewRow
                            'With rowICTIADJ2
                            '    .Item("ADJ_NO") = ADJ_NO
                            '    ADJ_LNO += 1
                            '    .Item("ADJ_LNO") = ADJ_LNO
                            '    .Item("STYLE_CODE") = STYLE_CODE
                            '    .Item("COLOR_CODE") = COLOR_CODE
                            '    .Item("ADJ_QTY") = ADJ_QTY
                            '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            '    .Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
                            '    .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                            '    .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                            '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                            '    .Item("LOCATION_CODE") = rowICTWHSE1.Item("WHSE_LOC_SHP")
                            '    .Item("BAR_CODE") = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")
                            '    .Item("ADJ_REF") = ADJ_REF
                            'End With
                            'dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)

                            If r = 0 Then
                                ADJ_QTY = -1 * ADJ_QTY
                                STYLE_CODE = rowWHTWAVE2.Item("STYLE_CODE_SUB")
                                COLOR_CODE = rowWHTWAVE2.Item("COLOR_CODE_SUB")
                                ADJ_REF = "SUB WITH"
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

            End If ' WAVE_TYPE = 'S' on the Else
        End If ' chkFinalized.Checked

        CommitTrans("Update Complete")

        Debug.Print("F" & ":" & Now)


        For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("")
            Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
            Dim row As DataRow = dst.Tables("SOTSHIPX_PRE").Rows.Find(SHIP_BOL_NO)
            If row IsNot Nothing Then
                row.Delete()
            End If
            dst.Tables("SOTSHIPX_PRE").AcceptChanges()
        Next
        Debug.Print("G" & ":" & Now)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Overrides Function Audit_Context() As Audit_Entity

        Dim E As New Audit_Entity
        E.TABLE_NAME = "WHTWAVE1"
        E.KEY_VALUE = WAVE_NO
        E.KEY_DESC = "Wave No"
        Return E
    End Function


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

    Sub Print_Record()

        Fill_Records("WHTLOCM1", New String() {WHSE_CODE}, False)

        Dim LOCATION_CODE_DEPOSIT As String = Absx1.txtFor("LOCATION_CODE_DEPOSIT").Text
        If LOCATION_CODE_DEPOSIT <> "" Then
            If dst.Tables("WHTLOCM1").Rows.Find(New String() {WHSE_CODE, LOCATION_CODE_DEPOSIT}) Is Nothing Then
                MsgBox("Invalid Deposit Location Specified", MsgBoxStyle.OkOnly, "Cannot Print Wave Report")
            End If
        Else
            MsgBox("No Deposit Location Specified", MsgBoxStyle.OkOnly, "Cannot Print Wave Report")
            Exit Sub
        End If

        Display_Totals()
        Print_Report_Begin()

        If EntryMode <> "N" Then
            Dim Suppress_Completed_Instructions As Boolean = False
            Dim iResponse As MsgBoxResult = MsgBox("Suppress Completed Instructions?", MsgBoxStyle.YesNo, "Pay Attention!")
            If iResponse = MsgBoxResult.No Then
                Suppress_Completed_Instructions = False
            Else
                Suppress_Completed_Instructions = True
            End If
            Dim rowWHTLOCM1 As DataRow

            For Each rowWHTINST1 As DataRow In dst.Tables("WHTINST1").Select
                rowWHTLOCM1 = dst.Tables("WHTLOCM1").Rows.Find(New String() {WHSE_CODE, rowWHTINST1.Item("LOCATION_CODE")})
                rowWHTINST1.Item("LOCATION_ROUTE_SEQ") = rowWHTLOCM1.Item("LOCATION_ROUTE_SEQ")
                If rowWHTINST1.Item("WAVE_INST_STATUS") = "0" Then
                    rowWHTINST1.Item("SUPP_INSTR") = "0"
                Else
                    rowWHTINST1.Item("SUPP_INSTR") = IIf(Suppress_Completed_Instructions, "1", "0")
                End If
            Next
        End If

        Dim Pallet_Pick_Msg As String = ""
        Dim Case_Pick_Msg As String = ""
        Dim Units_Pick_Msg As String = ""
        For Each rowWHTWAVES As DataRow In dst.Tables("WHTWAVES").Select("WAVE_STAT_TYPE = 'I'")
            Select Case rowWHTWAVES.Item("WAVE_PICK_TYPE")
                Case "L"
                    Pallet_Pick_Msg = "Pallets Waved/Picked: " & Val(rowWHTWAVES.Item("WAVE") & "") & "/" & Val(rowWHTWAVES.Item("PICK") & "")
                Case "C"
                    Case_Pick_Msg = "Case Waved/Picked: " & Val(rowWHTWAVES.Item("WAVE") & "") & "/" & Val(rowWHTWAVES.Item("PICK") & "")
                Case "U"
                    Units_Pick_Msg = "Pieces Waved/Picked: " & Val(rowWHTWAVES.Item("WAVE") & "") & "/" & Val(rowWHTWAVES.Item("PICK") & "")
            End Select
        Next


        CR_params.Add("SUBT", "")
        CR_params.Add("WAVE_PICK_MSG", Pallet_Pick_Msg & "           " & Case_Pick_Msg & "           " & Units_Pick_Msg)
        Generate_Report("WHRWAVE1")
        Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where = "WHSE_LOCATOR = '1'"
            Case "LOCATION_CODE_DEPOSIT"
                sql_where = "LOCATION_USE = 'D' and WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
            Case "P2L_LINE_ID"
                sql_where = "P2L_STATUS = 'A' and WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "' and CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
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

            Case "View"
                Absx1.txtFor("WAVE_NO").Text = key
                Click_Command("View")

        End Select

        Return return_key
    End Function


#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPX, "SSSBBBBBS", "Show Filter", "Show GroupBox", "Show Pins", "Customer Order Inquiry", "De-Select All", "Select All for Customer", "Select Selected", "Select All", "Show Billed Not Waved")
        Load_Popup_Menu(grdWHTWAVE2, "BBBS", "Style Status Inquiry", "Location Inquiry", "Add Sub", "Show Details")
        Load_Popup_Menu(grdWHTLOCB1, "BS", "Location Inquiry", "Summary by Location")
        Load_Popup_Menu(grdWHTINST1, "BSBBBBBB", "Location Inquiry", "Show Pick Details", "Void Pick", "De-Select All", "Select All", "Pick Selected", "Pick", "Show Instruction Events")
        Load_Popup_Menu(grdWHTWAVEX, "SBB", "Show Filter", "Customer Order Inquiry", "Wave Inquiry")
        Load_Popup_Menu(grdWHTWAVE3, "B", "Customer Order Inquiry")
        Load_Popup_Menu(grdWHTWAVEP, "BB", "De-Select All", "Select All")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case e.SourceControl.Name
            Case "grdSOTSHIPX"
                If ScreenMode Then e.Cancel = True
                Exit Sub
        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdSOTSHIPX"
                tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow
            Case "grdWHTWAVE2"
                tlb_sbt = DirectCast(tlb_pop.Tools("Show Details"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = Not SplitContainer1.Panel2Collapsed
                tlb_btn = DirectCast(tlb_pop.Tools("Add Sub"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" Or EntryMode = "E") And (WAVE_TYPE = "S" Or (WAVE_TYPE = "L" And EntryMode = "N"))
            Case "grdWHTINST1"
                tlb_btn = DirectCast(tlb_pop.Tools("Void Pick"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E")
                tlb_btn = DirectCast(tlb_pop.Tools("Pick"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (Not InquiryMode And EntryMode = "V") And Not cmdClosePicks.Visible
                tlb_btn = DirectCast(tlb_pop.Tools("Pick Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (Not InquiryMode And EntryMode = "V") And Not cmdClosePicks.Visible
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (Not InquiryMode And EntryMode = "V") And Not cmdClosePicks.Visible
                tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (Not InquiryMode And EntryMode = "V") And Not cmdClosePicks.Visible
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Show Instruction Events"
                ASCMAIN1.CodeSelector.Get_SQL("WAVE_INST_NO")

                ASCMAIN1.CodeSelector.SQL = "Select INIT_DATE, INIT_OPER, EVENT_DESC from WHTINSTE Where WAVE_INST_NO = '" & grd.ActiveRow.Cells("WAVE_INST_NO").Value & "' Order by INIT_DATE desc"
                ASCMAIN1.CodeSelector.MultipleSelections = False
                Using F As New ASFCODE1
                    F.ShowDialog()
                End Using

            Case "De-Select All"
                Dim tname As String = ""
                Select Case grd.Name
                    Case "grdSOTSHIPX"
                        tname = "SOTSHIPX"
                    Case "grdWHTINST1"
                        tname = "WHTINST1"
                    Case "grdWHTWAVEP"
                        tname = "WHTWAVEP"
                End Select
                If tname <> "" Then
                    If tname = "WHTWAVEP" Then
                        For Each row As DataRow In dst.Tables(tname).Select("SEL = '1'")
                            row.Item("SEL") = "0"
                        Next
                    Else
                        For Each row As DataRow In dst.Tables(tname).Select("SELECTED = '1'")
                            row.Item("SELECTED") = "0"
                        Next
                    End If

                End If


            Case "Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grd.Name = "grdWHTWAVEP" Then
                        If Not grow.IsFilteredOut And grow.IsDataRow Then
                            grow.Cells("SEL").Value = "1"
                        End If
                    Else
                        If Not grow.IsFilteredOut And grow.IsDataRow Then
                            grow.Cells("SELECTED").Value = "1"
                        End If
                    End If
                Next
                grd.UpdateData()

            Case "Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If Not grow.IsFilteredOut And grow.IsDataRow Then
                        grow.Cells("SELECTED").Value = "1"
                    End If
                Next
                grd.UpdateData()

            Case "Show Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                SplitContainer1.Panel2Collapsed = Not tlb_sbt.Checked

            Case "Summary by Location"
                ' tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Toggle_grdWHTLOCB1()
            Case "Show Billed Not Waved"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Dim sql As String = ""
                If tlb_sbt.Checked Then
                    sql_Wave_Filter = " and SOTSHIP1.SHIP_STATUS in ('F')" _
                    & " And SHIP_BOL_NO Not in (Select Distinct SHIP_BOL_NO from WHTWAVE3,WHTWAVE1 where WHTWAVE1.WAVE_NO = WHTWAVE3.WAVE_NO AND WHTWAVE1.WAVE_STATUS <> 'V')"

                    grdSOTSHIPX.Text = "Shipments Billed Not Waved"
                Else
                    sql_Wave_Filter = " and SOTSHIP1.SHIP_STATUS in ('P')"
                    grdSOTSHIPX.Text = "Shipments Released Not Waved"
                End If

                Load_SOTSHIPX()

            Case "Show Pick Details"
                Toggle_grdWHTINST1()

            Case "Pick Selected"
                ' Cannot do Piece Pick Automatically because of the entry of the new LPN - 
                'For Each row As DataRow In dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '0' and WAVE_PICK_TYPE <> 'U'")
                For Each row As DataRow In dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '0' and SELECTED = '1'", "WAVE_PICK_TYPE")
                    Dim WAVE_PICK_TYPE As String = row.Item("WAVE_PICK_TYPE")
                    Dim WAVE_INST_NO As String = row.Item("WAVE_INST_NO")
                    Start_App(WAVE_PICK_TYPE, WAVE_INST_NO)
                    If cmdClosePicks.Visible Then

                        Automate_Pick()

                        splPicks.Panel2Collapsed = True
                        cmdClosePicks.Visible = False
                        btnAutomate.Visible = False
                        C = Nothing
                    End If
                Next
                Close_Picks()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All for Customer"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If Not grow.IsFilteredOut And grow.Cells("CUST_CODE").Value = CUST_CODE Then
                        grow.Cells("SELECTED").Value = "1"
                    End If
                Next
                grd.UpdateData()

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")

            Case "Wave Inquiry"
                Dim WAVE_NO As String = grd.ActiveRow.Cells("WAVE_NO").Value
                Context_Launch("View", WAVE_NO, e.Tool.Key, "WHFWAVEI")

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Location Inquiry"
                Dim KEY As String = ""
                If grd.Name = "grdWHTWAVE2" Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                    KEY = "S:" & STYLE_CODE
                ElseIf grd.Name = "grdWHTLOCB1" Then
                    Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Value
                    KEY = "L:" & LOCATION_CODE
                ElseIf grd.Name = "grdWHTINST1" Then
                    If grd.ActiveRow.Band.Index = 0 Then
                        Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Value
                        KEY = "L:" & LOCATION_CODE
                    Else
                        Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                        KEY = "S:" & STYLE_CODE
                    End If
                End If

                Context_Launch("Select", KEY, e.Tool.Key, "WHFLOCS1")

            Case "Add Sub"

                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Index = 0 Then

                    'Dim grow As UltraWinGrid.UltraGridRow = grd.DisplayLayout.Bands(1).AddNew

                    Dim rowWHTWAVE2_SUB As DataRow = dst.Tables("WHTWAVE2_SUB").NewRow
                    rowWHTWAVE2_SUB.Item("WAVE_NO") = grd.ActiveRow.Cells("WAVE_NO").Value
                    rowWHTWAVE2_SUB.Item("WAVE_LNO_LINK") = grd.ActiveRow.Cells("WAVE_LNO").Value
                    WAVE_LNO_ctr += 1
                    rowWHTWAVE2_SUB.Item("WAVE_LNO") = WAVE_LNO_ctr
                    rowWHTWAVE2_SUB.Item("STYLE_CODE") = grd.ActiveRow.Cells("STYLE_CODE").Value
                    rowWHTWAVE2_SUB.Item("COLOR_CODE") = grd.ActiveRow.Cells("COLOR_CODE").Value
                    rowWHTWAVE2_SUB.Item("STYLE_CODE_SUB") = grd.ActiveRow.Cells("STYLE_CODE").Value
                    rowWHTWAVE2_SUB.Item("COLOR_CODE_SUB") = grd.ActiveRow.Cells("COLOR_CODE").Value
                    dst.Tables("WHTWAVE2_SUB").Rows.Add(rowWHTWAVE2_SUB)

                    grd.ActiveRow.ExpandAll()
                    grd.ActiveRow.ChildBands(0).Rows(0).Activate()
                End If

            Case "Void Pick"

                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow OrElse grd.ActiveRow.Band.Index <> 0 Then

                Else

                    If grd.ActiveRow.Cells("WAVE_INST_STATUS").Value & "" <> "1" Then
                        MsgBox("You May Reverse an Instruction Only if it has been Picked" _
                               & vbCrLf & "  (and not already reversed)",
                               MsgBoxStyle.OkOnly, "Cannot Reverse an Instruction that is Not Picked")
                    Else
                        Dim WAVE_INST_NO As String = grd.ActiveRow.Cells("WAVE_INST_NO").Value
                        If MsgBox("OK to Void Pick Instruction " & WAVE_INST_NO, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            Dim row1_original As DataRow = dst.Tables("WHTINST1").Rows.Find(WAVE_INST_NO)
                            row1_original.Item("WAVE_INST_STATUS") = "V"
                            Dim row1_reversed As DataRow = dst.Tables("WHTINST1").NewRow
                            row1_reversed.ItemArray = row1_original.ItemArray
                            Dim WAVE_INST_NO_reversing As String = ASCMAIN1.Next_Control_No("WHTINST1.WAVE_INST_NO")
                            row1_original.Item("WAVE_INST_NO_REVERSED_BY") = WAVE_INST_NO_reversing
                            WAVE_INST_NO_void.Add(WAVE_INST_NO_reversing)
                            'row1_reversed.Item("LAST_OPER") = WAVE_INST_NO
                            row1_reversed.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            row1_reversed.Item("INIT_DATE") = DATETIME_STAMP
                            row1_reversed.Item("WAVE_INST_NO") = WAVE_INST_NO_reversing
                            row1_reversed.Item("WAVE_INST_NO_REVERSING") = WAVE_INST_NO
                            'row1_reversed.Item("LOCATION_CODE") = row1_original.Item("LOCATION_CODE_OTHER")
                            'row1_reversed.Item("LOAD_NO") = row1_original.Item("LOAD_NO_OTHER")
                            'row1_reversed.Item("LOCATION_CODE_OTHER") = row1_original.Item("LOCATION_CODE")
                            'row1_reversed.Item("LOAD_NO_OTHER") = row1_original.Item("LOAD_NO")
                            dst.Tables("WHTINST1").Rows.Add(row1_reversed)

                            For Each row2_original As DataRow In dst.Tables("WHTINST2").Select("WAVE_INST_NO = '" & WAVE_INST_NO & "'")
                                Dim row2_reversed As DataRow = dst.Tables("WHTINST2").NewRow
                                row2_reversed.ItemArray = row2_original.ItemArray
                                row2_reversed.Item("WAVE_INST_NO") = WAVE_INST_NO_reversing
                                row2_reversed.Item("LOCATION_QTY_WAVE") = -1 * Val(row2_reversed.Item("LOCATION_QTY_WAVE") & "")
                                row2_reversed.Item("LOCATION_QTY_PICK") = -1 * Val(row2_reversed.Item("LOCATION_QTY_PICK") & "")
                                dst.Tables("WHTINST2").Rows.Add(row2_reversed)
                                Dim WAVE_LNO As Int32 = Val(row2_original.Item("WAVE_LNO") & "")
                                Dim rowWHTWAVE2 As DataRow = dst.Tables("WHTWAVE2").Rows.Find(New Object() {WAVE_NO, WAVE_LNO})
                                If rowWHTWAVE2 Is Nothing Then
                                    Dim WAVE_LNO_LINK As Int64 = 0
                                    Dim rowWHTWAVE2_SUBS() As DataRow = dst.Tables("WHTWAVE2_SUB").Select("WAVE_NO = '" & WAVE_NO & "' and WAVE_LNO = " & CStr(WAVE_LNO))
                                    If rowWHTWAVE2_SUBS.Length = 1 Then
                                        WAVE_LNO_LINK = rowWHTWAVE2_SUBS(0).Item("WAVE_LNO_LINK")
                                        rowWHTWAVE2 = dst.Tables("WHTWAVE2").Rows.Find(New Object() {WAVE_NO, WAVE_LNO_LINK})
                                    End If

                                End If
                                rowWHTWAVE2.Item("WAVE_QTY_PICK") = Val(rowWHTWAVE2.Item("WAVE_QTY_PICK") & "") - Val(row2_original.Item("LOCATION_QTY_PICK") & "")
                                rowWHTWAVE2.Item("WAVE_QTY") = Val(rowWHTWAVE2.Item("WAVE_QTY") & "") - Val(row2_original.Item("LOCATION_QTY_WAVE") & "")

                            Next
                        End If
                    End If
                End If

                grdWHTINST1.ActiveRow.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

            Case "Pick"
                Start_App(grdWHTINST1.ActiveRow.Cells("WAVE_PICK_TYPE").Value, grdWHTINST1.ActiveRow.Cells("WAVE_INST_NO").Value)

        End Select
    End Sub

#End Region

    Sub Start_App(WAVE_PICK_TYPE As String, Optional WAVE_INST_NO As String = "")

        ASCMAIN1.sql = "Select APP_ID from WHTGUNA1 where PICK_TYPE = '" & WAVE_PICK_TYPE & "' and USE_CLASS = '1'"
        Dim row As DataRow = ASCDATA1.GetDataRow
        If row Is Nothing Then
            MsgBox("Problem Identifying Pick Application", MsgBoxStyle.OkOnly, "Cannot Pick")
            Exit Sub
        End If

        Dim APP_ID As String = row.Item("APP_ID")

        Dim rowWHTGUNA1 As DataRow = LookUp("WHTGUNA1", APP_ID)
        Dim APP_DESC As String = rowWHTGUNA1.Item("APP_DESC")
        Dim PROCEDURE_NAME As String = rowWHTGUNA1.Item("PROCEDURE_NAME") & ""
        Dim GUN_PARAM As String = rowWHTGUNA1.Item("PICK_TYPE") & ""

        splPicks.Panel2Collapsed = False
        cmdClosePicks.Visible = True
        btnAutomate.Visible = True
        'If WAVE_PICK_TYPE = "U" Then
        '    btnAutomate.Enabled = False
        'Else
        btnAutomate.Enabled = True
        'End If

        txtPickChat.Text = ""

        Dim GUN_LOC As String = "99-G00-A"

        'Dim C As WHC.WHCRF000
        C = WHC.WHCFACT1.CreateWhcClass(PROCEDURE_NAME, New WHC.GunEnvironment With
            {.DBS_COMPANY = ASCMAIN1.DBS_COMPANY, .DBS_SERVER = ASCMAIN1.DBS_SERVER, .DBS_PASSWORD = ASCMAIN1.DBS_PASSWORD,
             .THREAD_NO = 0, .APP_ID = APP_ID, .APP_DESC = APP_DESC,
             .USER_ID = ASCMAIN1.USER_ID, .GUN_LOC = GUN_LOC, .PICK_TYPE = GUN_PARAM, .WHSE_CODE = WHSE_CODE})

        AddHandler C.RespondToScan, AddressOf Display_Text
        txt2.Focus()

        Display_Text(0, C.Hello)
        If WAVE_INST_NO <> "" Then
            txt2.Text = "I" & WAVE_INST_NO
        Else
            txt2.Text = WAVE_NO
        End If

    End Sub

    Sub Display_Text(THREAD_NO As Integer, TXT As String)
        txtPickChat.Text &= vbCrLf & "Thread " & THREAD_NO & ":" & TXT
        txtPickChat.SelectionStart = txtPickChat.Text.Length - 1
        txtPickChat.ScrollToCaret()

        txt2.Text = C.RESPONSE_anticipated_next
        txt2.SelectionStart = txt2.Text.Length
        txt2.ScrollToCaret()
    End Sub

    Sub txt2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt2.KeyDown

        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Try
                Dim txt2 As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
                'Dim THREAD_NO As Integer = txt2.Tag
                C.GetResponseToScan(txt2.Text)
                'txt2.Text = ""
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub cmdClosePicks_Click(sender As System.Object, e As System.EventArgs) Handles cmdClosePicks.Click
        Close_Picks()
    End Sub

    Sub Close_Picks()
        splPicks.Panel2Collapsed = True
        cmdClosePicks.Visible = False
        btnAutomate.Visible = False

        C = Nothing

        Dim WAVE_NO As String = Me.WAVE_NO
        Click_Command("Done")
        Absx1.txtFor("WAVE_NO").Text = WAVE_NO
        Click_Command("View")
    End Sub

    Private Sub btnAutomate_Click(sender As System.Object, e As System.EventArgs) Handles btnAutomate.Click
        Automate_Pick()
        If MsgBox("Automated Pick is Complete." & vbCrLf & vbCrLf & "Refresh Screen?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
            Close_Picks()
        End If
    End Sub

    Sub Automate_Pick()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now performing Automated Pick")

        Do While txt2.Text <> ""
            C.GetResponseToScan(txt2.Text)
        Loop

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'validate whse before loading
                    If Not ScreenMode Then Load_SOTSHIPX()

                End If

            Case "WAVE_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    '  spl.Focus()
                    Click_Command("View", e)
                End If


            Case "PICK_NO", "SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTSHIPX("S")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If Not ScreenMode Then Load_SOTSHIPX()

            Case "WAVE_NO"
                If Not ScreenMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                'If Not ScreenMode Then Load_SOTSHIPX()
        End Select
    End Sub

#End Region

    Sub Load_SOTSHIPX(Optional WAVE_TYPE As String = "")
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Data")

        loading_lead_screen = True

        EnforceConstraints(False)

        ASCMAIN1.sql = " and SOTSHIP1.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        If InquiryMode Then
            Dim sqlw As String = ""
            Dim C As String = ""
            If Absx1.txtFor("PICK_NO").Text <> "" Then
                sqlw &= " and SOTSHIP1.SHIP_BOL_NO in" & vbCrLf _
                    & " (Select SHIP_BOL_NO from SOTPICK1" & vbCrLf _
                    & " where SOTPICK1.PICK_NO = '" & Replace(Replace(Absx1.txtFor("PICK_NO").Text, "'", ""), ";", "") & "')"
                C &= "; Pick Ticket " & Absx1.txtFor("PICK_NO").Text
            End If

            If Absx1.txtFor("SHIP_BOL_NO").Text <> "" Then
                sqlw &= " and SOTSHIP1.SHIP_BOL_NO = '" & Replace(Replace(Absx1.txtFor("SHIP_BOL_NO").Text, "'", ""), ";", "") & "'"
                C &= "; Shipment " & Absx1.txtFor("SHIP_BOL_NO").Text
            End If

            If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                sqlw &= " and SOTSHIP1.SHIP_BOL_NO in" & vbCrLf _
                    & " (Select SHIP_BOL_NO from SOTSHIP1,SOTORDR0" & vbCrLf _
                    & " where SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                    & "   and SOTORDR0.ORDR_CUST_PO = '" & Replace(Replace(Absx1.txtFor("ORDR_CUST_PO").Text, "'", ""), ";", "") & "')"
                C &= "; Customer PO " & Absx1.txtFor("ORDR_CUST_PO").Text
            End If

            If Absx1.txtFor("CUST_CODE").Text <> "" Then
                sqlw &= " and SOTSHIP1.SHIP_BOL_NO in" & vbCrLf _
                    & " (Select SHIP_BOL_NO from SOTSHIP1,SOTORDR0" & vbCrLf _
                    & " where SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                    & "   and SOTORDR0.CUST_CODE = '" & Replace(Replace(Absx1.txtFor("CUST_CODE").Text, "'", ""), ";", "") & "')"
                C &= "; Customer " & Absx1.txtFor("CUST_CODE").Text
            End If

            ASCMAIN1.sql &= sqlw

            If optFilter.Value = "B^W" Then
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS in ('F')" _
                & " and SHIP_DATE_SHIPPED >= '" & Format(DateAdd("d", -14, Now), "dd-MMM-yy") & "'" _
                & " And SHIP_BOL_NO Not in (Select Distinct SHIP_BOL_NO from WHTWAVE3,WHTWAVE1 where WHTWAVE1.WAVE_NO = WHTWAVE3.WAVE_NO and WHTWAVE1.WAVE_STATUS <> 'V')"

                grdWHTWAVEX.Text = "Shipments Billed Not Waved"
            Else
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_BOL_NO in" & vbCrLf _
                    & " (Select SHIP_BOL_NO from WHTWAVE3,WHTWAVE1" & vbCrLf _
                    & " where WHTWAVE3.WAVE_NO = WHTWAVE1.WAVE_NO" & vbCrLf _
                    & "   and WHTWAVE1.WAVE_DATE between '" & Format(calFrom.Value, "dd-MMM-yyyy") & "' and '" & Format(calTo.Value, "dd-MMM-yyyy") & "')"

                grdWHTWAVEX.Text = "Waves Generated between " _
                    & Format(calFrom.Value, "MM/dd/yy") & " and " _
                    & Format(calTo.Value, "MM/dd/yy") _
                    & C
            End If


        Else
            ASCMAIN1.sql &= sql_Wave_Filter

            grdWHTWAVEX.Text = "Waves connected to Shipments In Pick"
        End If

        Create_Temp_Table(ASCMAIN1.sql)

        If InquiryMode Then
            grdWHTWAVE3.Text = "Shipments in Wave"
        Else
            grdWHTWAVE3.Text = "Shipments Released not Waved"
        End If

        Fill_Records("SOTSHIPX")

        For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("")
            Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("SHIP_BOL_NO")
            Dim row As DataRow = dst.Tables("SOTSHIPX_PRE").Rows.Find(SHIP_BOL_NO)
            If row IsNot Nothing Then
                rowSOTSHIPX.Item("WAVE_QTY") = row.Item("WAVE_QTY")
            End If
            dst.Tables("SOTSHIPX_PRE").AcceptChanges()
        Next

        Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)



        ASCMAIN1.sql = sqlWHTWAVEX
        ' NOTE - THIS NEXT BLOCK IS REPEATED BELOW FOR WORK ORDER WAVES
        If InquiryMode Then
            If optFilter.Value = "B^W" Then
                ASCMAIN1.sql = " Select '0000000000' WAVE_NO, Null WAVE_DATE, '' WAVE_TYPE, '' WAVE_STATUS" & vbCrLf _
                & " , 0 SC_COUNT, 0 WAVE_QTY_OPEN, 0 WAVE_QTY_PICK" & vbCrLf _
                & " , 0 WAVE_COUNT, 0 OPEN_COUNT, 0 PICK_COUNT" & vbCrLf _
                & " , 0 WAVE_QTY_REL, 0 WAVE_QTY_ADJ, 0 WAVE_QTY_SUB" & vbCrLf _
                & " , 0 WAVE_QTY_CANC, 0 WAVE_QTY_CONC, 0 WAVE_QTY_BACK" & vbCrLf _
                & " , SOTSHIPX.*" & vbCrLf _
                & "  from " & SOTSHIPX & " SOTSHIPX"
            Else
                ASCMAIN1.sql &= vbCrLf _
                & "   and WHTWAVE1.WAVE_DATE between '" & Format(calFrom.Value, "dd-MMM-yyyy") & "' and '" & Format(calTo.Value, "dd-MMM-yyyy") & "'"

                If optFilter.Value <> "N" Then
                    Select Case optFilter.Value
                        Case "W^B"
                            ASCMAIN1.sql &= vbCrLf _
                            & "   and WHTWAVE1.WAVE_STATUS = 'O' and SOTSHIPX.SHIP_STATUS = 'P'"
                        Case "B^F"
                            ASCMAIN1.sql &= vbCrLf _
                          & "   and WHTWAVE1.WAVE_STATUS = 'O' and SOTSHIPX.SHIP_STATUS = 'F'"
                        Case "W"
                            ASCMAIN1.sql &= vbCrLf _
                            & "   and WHTWAVE1.WAVE_TYPE = 'W'"
                    End Select
                End If
            End If
        Else
            ASCMAIN1.sql &= vbCrLf _
                & "   and WHTWAVE1.WAVE_STATUS = 'O'"
        End If
        Fill_Records("WHTWAVEX", , , ASCMAIN1.sql)


        If WAVE_TYPE = "S" Then
            ' NO NEED TO INCLUDE WORK ORDER WAVES
        Else
            If optFilter.Value <> "B^W" Then
                ASCMAIN1.sql = "Select WHTWAVE1.WAVE_NO, WHTWAVE1.WAVE_DATE, WHTWAVE1.WAVE_TYPE, WHTWAVE1.WAVE_STATUS" & vbCrLf _
                & ", WHTINSTX.SC_COUNT, WHTINSTX.WAVE_QTY_OPEN, WHTINSTX.WAVE_QTY_PICK" & vbCrLf _
                & ", WHTINSTY.WAVE_COUNT, WHTINSTY.OPEN_COUNT, WHTINSTY.PICK_COUNT" & vbCrLf _
                & ", WHTINSTZ.WAVE_QTY_REL, WHTINSTZ.WAVE_QTY_ADJ, WHTINSTZ.WAVE_QTY_SUB" & vbCrLf _
                & ", WHTINSTZ.WAVE_QTY_CANC, WHTINSTZ.WAVE_QTY_CONC, WHTINSTZ.WAVE_QTY_BACK" & vbCrLf _
                & ", SOTSHIPX.*" & vbCrLf _
                & " from " & SOTSHIPX & " SOTSHIPX, WHTWAVE1" & vbCrLf _
                & ", " & WHTINSTX & " WHTINSTX" & vbCrLf _
                & ", " & WHTINSTY & " WHTINSTY" & vbCrLf _
                & ", " & WHTINSTZ & " WHTINSTZ" & vbCrLf _
                & " where WHTINSTX.WAVE_NO (+) = WHTWAVE1.WAVE_NO" & vbCrLf _
                & "   and WHTINSTY.WAVE_NO (+) = WHTWAVE1.WAVE_NO" & vbCrLf _
                & "   and WHTINSTZ.WAVE_NO (+) = WHTWAVE1.WAVE_NO" & vbCrLf _
                & "   and SOTSHIPX.SHIP_BOL_NO (+) = WHTWAVE1.WAVE_TYPE" & vbCrLf _
                & "   and WHTWAVE1.WAVE_TYPE = 'W'"

                ' NOTE - THIS NEXT BLOCK IS REPEATED ABOVE FOR SALES ORDER WAVES
                If InquiryMode Then
                    ASCMAIN1.sql &= vbCrLf _
                        & "   and WHTWAVE1.WAVE_DATE between '" & Format(calFrom.Value, "dd-MMM-yyyy") & "' and '" & Format(calTo.Value, "dd-MMM-yyyy") & "'"

                    If optFilter.Value <> "N" Then
                        If optFilter.Value = "W^B" Then
                            ASCMAIN1.sql &= vbCrLf _
                                & "   and WHTWAVE1.WAVE_STATUS = 'O' and SOTSHIPX.SHIP_STATUS = 'P'"
                        End If
                        If optFilter.Value = "B^F" Then
                            ASCMAIN1.sql &= vbCrLf _
                                & "   and WHTWAVE1.WAVE_STATUS = 'O' and SOTSHIPX.SHIP_STATUS = 'F'"
                        End If
                    End If

                Else
                    ASCMAIN1.sql &= vbCrLf _
                        & "   and WHTWAVE1.WAVE_STATUS = 'O'"
                End If
                ''  Dim CC As Integer = dst.Tables("WHTWAVEX").Rows.Count
                Fill_Records("WHTWAVEX", , False, ASCMAIN1.sql)
            End If
        End If


        Dim WAVE_COUNT As Integer = ASCDATA1.SelectDistinct("WHTWAVEX", New String() {"WAVE_NO"}).Rows.Count
        grdWHTWAVEX.Text &= " - Wave Count: " & CStr(WAVE_COUNT) & " Waves"

        Sort_grdColumns(grdWHTWAVEX, "WAVE_NO".ToLower)

        EnforceConstraints(True)

        loading_lead_screen = False

        Setup_grdTATEVNT1()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdWHTWAVE3_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTWAVE3.AfterRowActivate
        If EntryMode <> "" Then Setup_grdSOTSHIPX()
    End Sub

    Private Sub grdWHTWAVE3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTWAVE3.InitializeRow

        If e.Row.IsDataRow Then
            If e.Row.Cells("SHIP_STATUS").Value & "" = "D" Then
                e.Row.Cells("SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Deleted"
            ElseIf e.Row.Cells("SHIP_STATUS").Value & "" = "F" Then
                e.Row.Cells("SHIP_STATUS").Appearance.BackColor = Drawing.Color.LightGreen
                e.Row.ToolTipText = "Shipped"
            Else
                If e.Row.Cells("SHIP_WAVE_STATUS").Value & "" = "1" Then
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Blue
                    e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Waved"
                Else
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Empty
                    e.Row.Cells("SHIP_BOL_NO").ToolTipText = ""
                End If
            End If
        End If
    End Sub

    Sub Setup_Edit_Wave()
        If grdWHTWAVEX.ActiveRow Is Nothing OrElse Not grdWHTWAVEX.ActiveRow.IsDataRow Then
            UltraExplorerBar1.Groups("Edit Work Order").Visible = False
            UltraExplorerBar1.Groups("Edit Shipment").Visible = False
        Else
            If grdWHTWAVEX.ActiveRow.Cells("WAVE_TYPE").Value & "" = "W" Then
                Dim Wave_No_Edit As String = grdWHTWAVEX.ActiveRow.Cells("WAVE_NO").Value
                Dim WHTWAVE1_Edit As DataRow = clsASCBASE1.LookUp("WHTWAVE1", Wave_No_Edit)

                UltraExplorerBar1.Groups("Edit Shipment").Visible = False
                UltraExplorerBar1.Groups("Edit Work Order").Visible = Not (ScreenMode Or InquiryMode Or tabMain.ActiveTab.Text <> "Waves")
                UltraExplorerBar1.Groups("Edit Work Order").Text = "Edit Work Order " & Wave_No_Edit

                dtPACKED.Value = WHTWAVE1_Edit.Item("SHIP_DATE_PACKED") & ""
                dtPLANNED.Value = WHTWAVE1_Edit.Item("SHIP_DATE_PLANNED") & ""
                dtROUTED.Value = WHTWAVE1_Edit.Item("SHIP_DATE_ROUTED") & ""
                txtWHSE_NOTES.Text = WHTWAVE1_Edit.Item("SHIP_NOTES_3PL") & ""
                txtSHIP_NOTES.Text = WHTWAVE1_Edit.Item("SHIP_NOTES") & ""
                txtAPPT_NO.Text = WHTWAVE1_Edit.Item("SHIP_APPT_NO") & ""
            Else
                SHIP_BOL_NO = grdWHTWAVEX.ActiveRow.Cells("SHIP_BOL_NO").Value
                rowSOTSHIP1 = Fill_Record("SOTSHIP1", SHIP_BOL_NO)

                UltraExplorerBar1.Groups("Edit Work Order").Visible = False
                UltraExplorerBar1.Groups("Edit Shipment").Visible = Not (ScreenMode Or InquiryMode)
                UltraExplorerBar1.Groups("Edit Shipment").Text = "Edit Shipment " & SHIP_BOL_NO
            End If

        End If
    End Sub

    Sub Setup_grdTATEVNT1()

        If loading_lead_screen Then Exit Sub

        If grdSOTSHIPX.ActiveRow Is Nothing OrElse Not grdSOTSHIPX.ActiveRow.IsDataRow Then
            tabShipments.Visible = False
            UltraExplorerBar1.Groups("Edit Shipment").Visible = False

        Else
            tabShipments.Visible = True
            EnforceConstraints(False)

            chkEditShipment.Checked = False
            SHIP_BOL_NO = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value
            rowSOTSHIP1 = Fill_Record("SOTSHIP1", SHIP_BOL_NO)

            UltraExplorerBar1.Groups("Edit Shipment").Visible = Not (ScreenMode Or InquiryMode)
            UltraExplorerBar1.Groups("Edit Shipment").Text = "Edit Shipment " & SHIP_BOL_NO

            Dim CUST_CODE As String = grdSOTSHIPX.ActiveRow.Cells("CUST_CODE").Value
            Dim ORDR_CUST_PO As String = grdSOTSHIPX.ActiveRow.Cells("ORDR_CUST_PO").Value

            Fill_Records("TATEVNT1", New String() {SHIP_BOL_NO})
            'tabShipments.Tabs("Events").Text = "Events for Shipment " & SHIP_BOL_NO & "; " & CUST_CODE & " PO " & ORDR_CUST_PO
            grdTATEVNT1.Text = "Events for Shipment " & SHIP_BOL_NO & "; " & CUST_CODE & " PO " & ORDR_CUST_PO
            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

            Fill_Records("SOTSHIPA", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTSHIPA, "INIT_DATE")
            grdSOTSHIPA.Text = "Audit Trail for Shipment " & SHIP_BOL_NO & "; " & CUST_CODE & " PO " & ORDR_CUST_PO

            ASCDATA1.ExecuteSQL("Delete from " & SOTSHIPC)
            ASCMAIN1.sql = "Select Distinct SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                 & " from SOTORDR2,SOTPICK2,SOTPICK1" & vbCrLf _
                 & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                 & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                 & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                 & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPC & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from (" & vbCrLf _
                & "Select POTSHIP1.WHSE_CODE, POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
                & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO, POTSHIP3.PO_QTY_SHP" & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
                & ", POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE" & vbCrLf _
                & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, " & SOTSHIPC & " SOTSHIPC " & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO " & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
                & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTORDR2.STYLE_CODE = SOTSHIPC.STYLE_CODE" & vbCrLf _
                & "   and POTORDR2.COLOR_CODE = SOTSHIPC.COLOR_CODE" & vbCrLf _
                & ")"

            Fill_Records("POTORDRX", "", , ASCMAIN1.sql)
            Sort_grdColumns(grdPOTORDRX, "PO_SHIP_ETA")
            grdPOTORDRX.Text = "Inbound Shipments for Style-Colors in Shipment " & SHIP_BOL_NO & "; " & CUST_CODE & " PO " & ORDR_CUST_PO


            'If CUST_CODE = "WALMART" Then
            If IsWalmart(CUST_CODE) Then
                dst.Tables("WHTLOCB1").Rows.Clear()
                '  Fill_Records("WHTLOCB1", New String() {SHIP_BOL_NO})

                grdWHTLOCB1.Text = "Locations View skipped for Shipment " & SHIP_BOL_NO & "; " & CUST_CODE & " PO " & ORDR_CUST_PO

            Else
                '  Fill_Records("WHTLOCB1", New String() {SHIP_BOL_NO})
                Toggle_grdWHTLOCB1()
                grdWHTLOCB1.Text = "Locations for Style-Colors in Shipment " & SHIP_BOL_NO & "; " & CUST_CODE & " PO " & ORDR_CUST_PO

            End If

            cmbType.Value = DBNull.Value
            txtNote.Text = ""

            EnforceConstraints(True)
        End If
    End Sub

    Sub Setup_grdSOTSHIPX()
        If grdWHTWAVE3.ActiveRow Is Nothing OrElse Not grdWHTWAVE3.ActiveRow.IsDataRow Then
            tabSOTSHIPX.Visible = False
        Else
            tabSOTSHIPX.Visible = True
            EnforceConstraints(False)
            Dim SHIP_BOL_NO As String = grdWHTWAVE3.ActiveRow.Cells("SHIP_BOL_NO").Value

            'Fill_Records("SOTPICK1", SHIP_BOL_NO)
            'Fill_Records("SOTPICK2", SHIP_BOL_NO)
            Dim dvw As DataView = DirectCast(grdSOTPICK1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"

            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICK1.Text = "Pick Tickets for Shipment " & SHIP_BOL_NO

            EnforceConstraints(True)
        End If
    End Sub

    Sub Create_Temp_Table(SQLW As String)

        Dim sqlWHTINSTX As String = "Select WHTINST1.WAVE_NO" & vbCrLf _
            & ", COUNT(DISTINCT WHTINST2.STYLE_CODE || WHTINST2.COLOR_CODE) SC_COUNT" & vbCrLf _
            & ", SUM(DECODE(WHTINST1.WAVE_INST_STATUS,'0',LOCATION_QTY_WAVE,0)) WAVE_QTY_OPEN" & vbCrLf _
            & ", SUM(DECODE(WHTINST1.WAVE_INST_STATUS,'1',LOCATION_QTY_PICK,'2',LOCATION_QTY_PICK,0)) WAVE_QTY_PICK" & vbCrLf _
            & " from WHTINST1,WHTINST2 where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
            & " group by WHTINST1.WAVE_NO"

        Dim sqlWHTINSTY As String = "Select WHTINST1.WAVE_NO" & vbCrLf _
            & ", COUNT (*) WAVE_COUNT" _
            & ", SUM (DECODE(WHTINST1.WAVE_INST_STATUS,'0',1,0)) OPEN_COUNT" _
            & ", SUM (DECODE(WHTINST1.WAVE_INST_STATUS,'1',1,0)) PICK_COUNT" _
            & " from WHTINST1 " _
            & " group by WHTINST1.WAVE_NO"

        Dim sqlWHTINSTZ As String = "Select WHTWAVE2.WAVE_NO" & vbCrLf _
            & ", Sum (PICK_QTY) WAVE_QTY_REL" & vbCrLf _
            & ", Sum (WAVE_QTY_ADJ) WAVE_QTY_ADJ" & vbCrLf _
            & ", Sum (WAVE_QTY_SUB) WAVE_QTY_SUB" & vbCrLf _
            & ", Sum (WAVE_QTY_CANC) WAVE_QTY_CANC" & vbCrLf _
            & ", Sum (WAVE_QTY_CONC) WAVE_QTY_CONC" & vbCrLf _
            & ", Sum (WAVE_QTY_BACK) WAVE_QTY_BACK" & vbCrLf _
            & " from WHTWAVE2" & vbCrLf _
            & " group by WHTWAVE2.WAVE_NO"

        If SOTSHIPX = "" Then

            ASCMAIN1.sql = "Select WHTLOCB1.*, WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, '0' PPK, WHTBARC1.PPK_CODE" & vbCrLf _
                & ", WHTLOCM1.LOCATION_LOCKED, WHTLOCM1.LOCATION_NOT_WAVED, WHTLOCM1.LOCATION_USE" & vbCrLf _
                & " from WHTLOCB1,WHTBARC1,WHTBARC0,WHTLOCM1" _
                & " where ROWNUM < 0"
            WHTLOCBW = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & WHTLOCBW & " Add Primary Key (WHSE_CODE,LOCATION_CODE,BAR_CODE,STYLE_CODE,COLOR_CODE)")

            sqlSOTSHIPX = "Select SOTSHIP1.SHIP_BOL_NO,SOTSHIP1.SHIP_VIA_CODE,SOTSHIP1.SHIP_REF" & vbCrLf _
                & ",SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE,SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & ",SOTSHIP1.FRT_TERMS,SOTSHIP1.WHSE_CODE,SOTSHIP1.SREP_CODE,SOTSHIP1.ORDR_DEPT" & vbCrLf _
                & ",SOTSHIP1.SHIP_DATE_RECEIVED,SOTSHIP1.SHIP_NOTES,SOTSHIP1.SHIP_SPEC_INST" & vbCrLf _
                & ",SOTSHIP1.SHIP_DATE_PLANNED,SOTSHIP1.SHIP_DATE_ROUTED,SOTSHIP1.SHIP_NOTES_3PL" & vbCrLf _
                & ",SOTSHIP1.SHIP_WAVE_STATUS,SOTSHIP1.SHIP_DATE_PACKED,SOTSHIP1.SHIP_APPT_NO" & vbCrLf _
                & ",SOTSHIP1.SHIP_STATUS,SOTSHIP1.SHIP_DATE_SHIPPED,SOTSHIP1.SHIPPED_ACTUAL" & vbCrLf _
                & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO,SOTORDR0.ORDR_ORIG_SHIP_DATE,SOTORDR0.ORDR_ORIG_CANCEL_DATE" & vbCrLf _
                & ",ARTCUST1.CUST_NAME" & vbCrLf _
                & " from SOTSHIP1,SOTORDR0,ARTCUST1" & vbCrLf _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            SOTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_NO_MIN VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_NO_MAX VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_NO_MIN VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_NO_MAX VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_NO_COUNT NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_QTY NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_QTY_PICK NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add P2L_ALLOW VARCHAR2(1)")

            ASCMAIN1.AnalyzeTable(SOTSHIPX)

            ASCMAIN1.sql = "Select SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "  , SUM (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
                & "  , SUM (DECODE(SOTPICK1.PICK_STATUS,'P',SOTPICK2.PICK_QTY,0)) PICK_QTY_PICK" & vbCrLf _
                & "   from SOTPICK2,SOTPICK1," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
                & "   where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "     and SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   group by SOTSHIPX.SHIP_BOL_NO" & vbCrLf
            SOTSHIPX_1 = ASCMAIN1.Temp_Table

            WHTINSTX = ASCMAIN1.Temp_Table("Select * from (" & sqlWHTINSTX & ") where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & WHTINSTX & " Add Primary Key (WAVE_NO)")
            For Each COLUMN_NAME As String In New String() {"SC_COUNT", "WAVE_QTY_OPEN", "WAVE_QTY_PICK"}
                ASCDATA1.ExecuteSQL("Alter Table " & WHTINSTX & " Modify " & COLUMN_NAME & " NUMBER (8,0)")
            Next

            WHTINSTY = ASCMAIN1.Temp_Table("Select * from (" & sqlWHTINSTY & ") where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & WHTINSTY & " Add Primary Key (WAVE_NO)")
            For Each COLUMN_NAME As String In New String() {"WAVE_COUNT", "OPEN_COUNT", "PICK_COUNT"}
                ASCDATA1.ExecuteSQL("Alter Table " & WHTINSTY & " Modify " & COLUMN_NAME & " NUMBER (8,0)")
            Next
            'For Each COLUMN_NAME As String In New String() {"SC_COUNT"}
            '    ASCDATA1.ExecuteSQL("Alter Table " & WHTINSTY & " Add " & COLUMN_NAME & " NUMBER (8,0)")
            'Next

            WHTINSTZ = ASCMAIN1.Temp_Table("Select * from (" & sqlWHTINSTZ & ") where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & WHTINSTZ & " Add Primary Key (WAVE_NO)")
            For Each COLUMN_NAME As String In New String() {"WAVE_QTY_REL", "WAVE_QTY_ADJ", "WAVE_QTY_SUB", "WAVE_QTY_CANC", "WAVE_QTY_CONC", "WAVE_QTY_BACK"}
                ASCDATA1.ExecuteSQL("Alter Table " & WHTINSTZ & " Modify " & COLUMN_NAME & " NUMBER (8,0)")
            Next

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE from SOTORDR2 where ROWNUM < 1"
            SOTSHIPC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPC & " Add Primary Key (STYLE_CODE, COLOR_CODE)")


            'ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO from SOTSHIP1 Where rownum < 1"
            'TempTable_BOL_NOs = ASCMAIN1.Temp_Table
        Else
            Load_SOTSHIPX_data(SQLW)

            'the code that was here was moved to sub Load_Dependent_Data
            'so it maybe called from inside Load_SOTSHIPX_data, which is called from other places


            ' THESE 3 SQLS NEED TO BE EVENTUALLY WHERE-CLAUSED TO A LIST OF WAVES THAT WE ARE INTERESTED IN, INSTEAD OF THE ENTIRE TABLES

            ASCDATA1.ExecuteSQL("Delete from " & WHTINSTX)
            ASCDATA1.ExecuteSQL("Insert into " & WHTINSTX & " " & sqlWHTINSTX)

            ASCDATA1.ExecuteSQL("Delete from " & WHTINSTY)
            ASCDATA1.ExecuteSQL("Insert into " & WHTINSTY & " " & sqlWHTINSTY)

            ASCDATA1.ExecuteSQL("Delete from " & WHTINSTZ)
            ASCDATA1.ExecuteSQL("Insert into " & WHTINSTZ & " " & sqlWHTINSTZ)
        End If
    End Sub

    Sub Load_SOTSHIPX_data(sqlw As String)
        ASCMAIN1.sql = Replace(sqlSOTSHIPX, " from ", ", NULL PICK_NO_MIN, NULL PICK_NO_MAX, NULL ORDR_NO_MIN, NULL ORDR_NO_MAX, 0 PICK_NO_COUNT, 0 PICK_QTY, 0 PICK_QTY_PICK, 'N' P2L_ALLOW from ") & sqlw
        ASCDATA1.ExecuteSQL("Delete from " & SOTSHIPX)
        ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " " & ASCMAIN1.sql)
        Load_Dependent_Data()
    End Sub
    Sub Load_Dependent_Data()

        ASCMAIN1.AnalyzeTable(SOTSHIPX)
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX_1)

        ASCMAIN1.sql = "Select SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
            & "  , SUM (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
            & "  , SUM (DECODE(SOTPICK1.PICK_STATUS,'P',SOTPICK2.PICK_QTY,0)) PICK_QTY_PICK" & vbCrLf _
            & "   from SOTPICK2,SOTPICK1," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
            & "   where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "     and SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
            & "   group by SOTSHIPX.SHIP_BOL_NO" & vbCrLf
        SOTSHIPX_1 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select * from " & SOTSHIPX_1 & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & SOTSHIPX & " Set " & vbCrLf _
            & "     PICK_QTY = R1.PICK_QTY" & vbCrLf _
            & "   , PICK_QTY_PICK = R1.PICK_QTY_PICK" & vbCrLf _
            & "   where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ' TRUNCATE SOTSHIPX_COPY
        ' INSERT INTO SOTSHIPX_COPY SELECT * FROM SOTSHIPX
        ' TRUNCATE TABLE SOTSHIPX
        ' INSERT INTO SOTSHIPX SELECT ...
        ASCMAIN1.sql = "INSERT INTO " & SOTSHIPX _
& "select " _
& "SOTSHIPX_1.SHIP_BOL_NO," _
& "SOTSHIPX_1.SHIP_VIA_CODE," _
& "SOTSHIPX_1.SHIP_REF," _
& "SOTSHIPX_1.SHIP_ADDR_TYPE," _
& "SOTSHIPX_1.SHIP_ADDR_CODE," _
& "SOTSHIPX_1.ORDR_GROUP_NO," _
& "SOTSHIPX_1.FRT_TERMS," _
& "SOTSHIPX_1.WHSE_CODE," _
& "SOTSHIPX_1.SREP_CODE," _
& "SOTSHIPX_1.ORDR_DEPT," _
& "SOTSHIPX_1.SHIP_DATE_RECEIVED," _
& "SOTSHIPX_1.SHIP_NOTES," _
& "SOTSHIPX_1.SHIP_SPEC_INST," _
& "SOTSHIPX_1.SHIP_DATE_PLANNED," _
& "SOTSHIPX_1.SHIP_DATE_ROUTED," _
& "SOTSHIPX_1.SHIP_NOTES_3PL," _
& "SOTSHIPX_1.SHIP_WAVE_STATUS," _
& "SOTSHIPX_1.SHIP_DATE_PACKED," _
& "SOTSHIPX_1.SHIP_APPT_NO," _
& "SOTSHIPX_1.SHIP_STATUS," _
& "SOTSHIPX_1.SHIP_DATE_SHIPPED," _
& "SOTSHIPX_1.SHIPPED_ACTUAL," _
& "SOTSHIPX_1.CUST_CODE," _
& "SOTSHIPX_1.ORDR_SHIP_DATE," _
& "SOTSHIPX_1.ORDR_CANCEL_DATE," _
& "SOTSHIPX_1.ORDR_CUST_PO," _
& "SOTSHIPX_1.ORDR_ORIG_SHIP_DATE," _
& "SOTSHIPX_1.ORDR_ORIG_CANCEL_DATE," _
& "SOTSHIPX_1.CUST_NAME," _
& "SOTSHIPX_1.PICK_NO_MIN," _
& "SOTSHIPX_1.PICK_NO_MAX," _
& "SOTSHIPX_1.ORDR_NO_MIN," _
& "SOTSHIPX_1.ORDR_NO_MAX," _
& "SOTSHIPX_1.PICK_NO_COUNT," _
& "SOTSHIPX.PICK_QTY," _
& "SOTSHIPX.PICK_QTY_PICK" _
& "SOTSHIPX.P2L_ALLOW" _
& "from " & SOTSHIPX_1 & " SOTSHIPX_1,  " & "SOTSHIPX_COPY" & "  SOTSHIPX" _
& "WHERE SOTSHIPX.SHIP_BOL_NO(+) = SOTSHIPX_1.SHIP_BOL_NO"
        ' ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & "  , MIN (SOTPICK1.PICK_NO) PICK_NO_MIN, MAX (SOTPICK1.PICK_NO) PICK_NO_MAX" & vbCrLf _
            & "  , MIN (SOTPICK1.ORDR_NO) ORDR_NO_MIN, MAX (SOTPICK1.ORDR_NO) ORDR_NO_MAX" & vbCrLf _
            & "  , COUNT (SOTPICK1.PICK_NO) PICK_NO_COUNT" & vbCrLf _
            & "  , SUM (SOTPICK1.PICK_FREIGHT) PICK_FREIGHT" & vbCrLf _
            & "   from SOTPICK1" & vbCrLf _
            & "   where SOTPICK1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")" & vbCrLf _
            & "   group by SOTPICK1.SHIP_BOL_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & SOTSHIPX & " Set " & vbCrLf _
            & "     PICK_NO_MIN = R1.PICK_NO_MIN" & vbCrLf _
            & "   , PICK_NO_MAX = R1.PICK_NO_MAX" & vbCrLf _
            & "   , ORDR_NO_MIN = R1.ORDR_NO_MIN" & vbCrLf _
            & "   , ORDR_NO_MAX = R1.ORDR_NO_MAX" & vbCrLf _
            & "   , PICK_NO_COUNT = R1.PICK_NO_COUNT" & vbCrLf _
            & "   where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & " Select SOTSHIPX.SHIP_BOL_NO, case when (SOTORDR0.CUST_CODE = 'WALMART' and EDI_PROMOTION = 'POS REPLEN')" & vbCrLf _
            & "                or (SOTORDR0.CUST_CODE = 'KOHLS' and EDI_DEPT_DESC = 'BULK') then 'Y' else 'N' end P2L_ALLOW" & vbCrLf _
            & " From SOTORDR0, " & SOTSHIPX & " SOTSHIPX, EDT850T1, WHTP2LM1" & vbCrLf _
            & " where SOTORDR0.CUST_CODE = WHTP2LM1.CUST_CODE" & vbCrLf _
            & " and SOTSHIPX.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
            & " and SOTORDR0.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & SOTSHIPX & " Set " & vbCrLf _
            & "     P2L_ALLOW = R1.P2L_ALLOW" & vbCrLf _
            & "   where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub grdWHTWAVEX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTWAVEX.AfterRowActivate
        If Not ScreenMode Then
            Setup_Edit_Wave()
        End If
    End Sub

    Private Sub grdWHTWAVEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTWAVEX.DoubleClickRow
        If Not ScreenMode Then
            If e.Row.IsDataRow Then
                Absx1.txtFor("WAVE_NO").Text = e.Row.Cells("WAVE_NO").Value
                Click_Command("View")
            End If
        End If
    End Sub

    Sub Display_Totals()
        For Each row As DataRow In dst.Tables("WHTWAVES").Select("")
            Dim WAVE_PICK_TYPE As String = row.Item("WAVE_PICK_TYPE")
            Dim WAVE_STAT_TYPE As String = row.Item("WAVE_STAT_TYPE")
            Dim S As String = ""
            If WAVE_STAT_TYPE = "I" Then
                S = "COUNT(WAVE_INST_NO)"
            ElseIf WAVE_STAT_TYPE = "C" Then
                S = "SUM(CASES_WAVE)"
            ElseIf WAVE_STAT_TYPE = "U" Then
                S = "SUM(UNITS_WAVE)"
            End If
            Dim WAVE As Int32 = Val(dst.Tables("WHTINST1").Compute(S, "WAVE_PICK_TYPE = '" & WAVE_PICK_TYPE & "'") & "")
            If WAVE_STAT_TYPE = "C" Or WAVE_STAT_TYPE = "U" Then S = Replace(S, "WAVE", "PICK")
            Dim PICK As Int32 = Val(dst.Tables("WHTINST1").Compute(S, "WAVE_PICK_TYPE = '" & WAVE_PICK_TYPE & "' and WAVE_INST_STATUS = '1'") & "")
            row.Item("WAVE") = WAVE
            row.Item("PICK") = PICK
        Next

        grdWHTWAVES.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdWHTWAVES.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        grdWHTWAVES.DisplayLayout.Bands(0).Columns("WAVE_PICK_TYPE").HiddenWhenGroupBy = DefaultableBoolean.True
        grdWHTWAVES.DisplayLayout.Bands(0).SortedColumns.Add("WAVE_PICK_TYPE", False, True)
        grdWHTWAVES.Rows.ExpandAll(True)

        grdWHTWAVES.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)


    End Sub

    Private Sub grdWHTWAVE2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTWAVE2.AfterCellUpdate
        'If e.Cell.Column.Key = "WAVE_QTY_CANC" Then
        '    If Val(e.Cell.Value & "") <> 0 Then
        '        e.Cell.Row.Cells("WAVE_QTY_CONC").Value = DBNull.Value
        '    End If
        'End If
        'If e.Cell.Column.Key = "WAVE_QTY_CONC" Then
        '    If Val(e.Cell.Value & "") <> 0 Then
        '        e.Cell.Row.Cells("WAVE_QTY_CANC").Value = DBNull.Value
        '    End If
        'End If
    End Sub

    Private Sub grdWHTWAVE2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTWAVE2.AfterRowActivate
        Setup_grdWHTWAVE2()

    End Sub

    Sub Setup_grdWHTWAVE2()
        If grdWHTWAVE2.ActiveRow Is Nothing OrElse Not grdWHTWAVE2.ActiveRow.IsDataRow OrElse grdWHTWAVE2.ActiveRow.IsAddRow Then
            tabWHTMOVE2.Visible = False

        Else
            tabWHTMOVE2.Visible = True
            Dim STYLE_CODE As String = grdWHTWAVE2.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdWHTWAVE2.ActiveRow.Cells("COLOR_CODE").Value

            If grdWHTWAVE2.ActiveRow.Band.Index = 1 Then
                STYLE_CODE = grdWHTWAVE2.ActiveRow.Cells("STYLE_CODE_SUB").Value
                COLOR_CODE = grdWHTWAVE2.ActiveRow.Cells("COLOR_CODE_SUB").Value
            End If

            Toggle_grdWHTLOCB1()
            grdWHTLOCB1.Text = "Locations for Style-Color " & STYLE_CODE & "-" & COLOR_CODE

            Setup_POTORDRX(STYLE_CODE, COLOR_CODE)

            If grdWHTWAVE2.ActiveRow.Band.Index = 1 Then
                If Val(grdWHTWAVE2.ActiveRow.Cells("WAVE_QTY").Value & "") = 0 Then
                    grdWHTWAVE2.DisplayLayout.Bands(1).Columns("STYLE_CODE_SUB").CellActivation = UltraWinGrid.Activation.AllowEdit
                    grdWHTWAVE2.DisplayLayout.Bands(1).Columns("COLOR_CODE_SUB").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    grdWHTWAVE2.DisplayLayout.Bands(1).Columns("STYLE_CODE_SUB").CellActivation = UltraWinGrid.Activation.NoEdit
                    grdWHTWAVE2.DisplayLayout.Bands(1).Columns("COLOR_CODE_SUB").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End If

            Setup_SOTCART2(STYLE_CODE, COLOR_CODE)

            Setup_WHTINSTS(STYLE_CODE, COLOR_CODE)
        End If

    End Sub

    Sub Toggle_grdWHTINST1()

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Pick Details"), UltraWinToolbars.StateButtonTool)
        With grdWHTINST1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"LOCATION_CODE_OTHER", "LOAD_NO_OTHER", "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER", "WAVE_LNO", "WAVE_SUB", "LOAD_NO"} ', "CASES_PICK", "UNITS_PICK"}
                .Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
            Next
        End With
    End Sub

    Sub Toggle_grdWHTLOCB1()
        'Exit Sub

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Summary by Location"), UltraWinToolbars.StateButtonTool)
        Dim SummaryByLocation As Boolean = tlb_sbt.Checked

        If ScreenMode Or EntryMode <> "" Then

            If grdWHTWAVE2.ActiveRow Is Nothing OrElse Not grdWHTWAVE2.ActiveRow.IsDataRow Then Exit Sub

            Dim STYLE_CODE As String = grdWHTWAVE2.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdWHTWAVE2.ActiveRow.Cells("COLOR_CODE").Value

            If grdWHTWAVE2.ActiveRow.Band.Index = 1 Then
                STYLE_CODE = grdWHTWAVE2.ActiveRow.Cells("STYLE_CODE_SUB").Value
                COLOR_CODE = grdWHTWAVE2.ActiveRow.Cells("COLOR_CODE_SUB").Value
            End If


            If Not SummaryByLocation Then
                Fill_Records("WHTLOCB1", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            Else
                Fill_Records("WHTLOCB1", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
                Dim LOADs As New Dictionary(Of String, Dictionary(Of Int32, Int32))

                For Each row As DataRow In dst.Tables("WHTLOCB1").Select("")
                    Dim LOAD_NO As String = row.Item("LOAD_NO")
                    If Not LOADs.ContainsKey(LOAD_NO) Then
                        LOADs.Add(LOAD_NO, New Dictionary(Of Int32, Int32))
                    End If
                    Dim LOCATION_QTY As Int32 = Val(row.Item("LOCATION_QTY") & "")
                    If LOADs(LOAD_NO).ContainsKey(LOCATION_QTY) Then
                        LOADs(LOAD_NO)(LOCATION_QTY) += 1
                    Else
                        LOADs(LOAD_NO).Add(LOCATION_QTY, 1)
                    End If
                Next

                ASCMAIN1.sql = "Select WHTLOCB1.WHSE_CODE, WHTLOCB1.LOCATION_CODE, NULL BAR_CODE" & vbCrLf _
                    & ", WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, SUM (WHTLOCB1.LOCATION_QTY) LOCATION_QTY" & vbCrLf _
                    & ", NULL, NULL, NULL, NULL, SUM (WHTLOCB1.LOCATION_QTY_WAVE) LOCATION_QTY_WAVE" & vbCrLf _
                    & ", WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, WHTLOCM1.LOCATION_LOCKED, Count(*) CASES" & vbCrLf _
                    & " from WHTLOCB1,WHTBARC0,WHTBARC1,WHTLOCM1,ICTWHSE1" & vbCrLf _
                    & " where WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO" & vbCrLf _
                    & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
                    & "   and WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & "   and WHTLOCB1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                    & "   and (NVL(WHTLOCB1.LOCATION_QTY,0)) > 0" & vbCrLf _
                    & "   and ICTWHSE1.WHSE_CODE = WHTLOCB1.WHSE_CODE" & vbCrLf _
                    & "   and WHTLOCM1.WHSE_CODE = WHTLOCB1.WHSE_CODE" & vbCrLf _
                    & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                    & "   and (NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1' or WHTLOCB1.LOCATION_CODE in (" & REC_LOCATIONS & "))" & vbCrLf _
                    & " group by WHTLOCB1.WHSE_CODE, WHTLOCB1.LOCATION_CODE, WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
                    & ", WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
                    & ", WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE"
                Fill_Records("WHTLOCB1", "", True, ASCMAIN1.sql)

                For Each row As DataRow In dst.Tables("WHTLOCB1").Select("")
                    Dim CASE_BREAKDOWN As String = ""
                    Dim LOAD_NO As String = row.Item("LOAD_NO")
                    For Each CASE_QTY As Int32 In LOADs(LOAD_NO).Keys
                        CASE_BREAKDOWN &= "," & CStr(LOADs(LOAD_NO)(CASE_QTY)) & "x" & CStr(CASE_QTY)
                    Next
                    row.Item("CASE_BREAKDOWN") = Mid(CASE_BREAKDOWN, 2)
                Next
            End If
        Else

            Dim WHSE_CODE As String = grdSOTSHIPX.ActiveRow.Cells("WHSE_CODE").Value

            Dim rowWHSE As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

            'ASCMAIN1.sql = "Select WHTLOCB1.WHSE_CODE, WHTLOCB1.LOCATION_CODE, WHTLOCB1.BAR_CODE" & vbCrLf _
            '   & ", WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_QTY" & vbCrLf _
            '   & ", WHTLOCB1.INIT_DATE, WHTLOCB1.INIT_OPER, WHTLOCB1.LAST_DATE, WHTLOCB1.LAST_OPER, WHTLOCB1.LOCATION_QTY_WAVE" & vbCrLf _
            '   & ", WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
            '   & " from WHTLOCB1,WHTBARC1,WHTBARC0,WHTLOCM1," & SOTSHIPC & " SOTSHIPC" & vbCrLf _
            '   & " where WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO" & vbCrLf _
            '   & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
            '   & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
            '   & "   and (NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1' or " & vbCrLf _
            '   & "   (WHTLOCM1.LOCATION_CODE = '" & rowICTWHSE1.Item("WHSE_LOC_REC") & "' or WHTLOCM1.LOCATION_CODE = '" & rowICTWHSE1.Item("WHSE_LOC_PAW") & "'))" & vbCrLf _
            '   & "   and (NVL(WHTLOCB1.LOCATION_QTY,0)) > 0" & vbCrLf _
            '   & "   and WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            '   & "   and WHTLOCB1.STYLE_CODE = SOTSHIPC.STYLE_CODE and WHTLOCB1.COLOR_CODE = SOTSHIPC.COLOR_CODE"

            ASCMAIN1.sql = "Select WHTLOCB1.WHSE_CODE, WHTLOCB1.LOCATION_CODE, WHTLOCB1.BAR_CODE" & vbCrLf _
               & ", WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_QTY" & vbCrLf _
               & ", WHTLOCB1.INIT_DATE, WHTLOCB1.INIT_OPER, WHTLOCB1.LAST_DATE, WHTLOCB1.LAST_OPER, WHTLOCB1.LOCATION_QTY_WAVE" & vbCrLf _
               & ", WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
               & " from WHTLOCB1,WHTBARC1,WHTBARC0,WHTLOCM1,ICTWHSE1," & SOTSHIPC & " SOTSHIPC" & vbCrLf _
               & " where WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO" & vbCrLf _
               & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
               & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
               & "   and (NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1' or WHTLOCB1.LOCATION_CODE in (" & REC_LOCATIONS & ",'" & rowWHSE.Item("WHSE_LOC_LNF") & "')) " & vbCrLf _
               & "   and (NVL(WHTLOCB1.LOCATION_QTY,0)) > 0" & vbCrLf _
               & "   and WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
               & "   and ICTWHSE1.WHSE_CODE = WHTLOCB1.WHSE_CODE" & vbCrLf _
               & "   and WHTLOCB1.STYLE_CODE = SOTSHIPC.STYLE_CODE and WHTLOCB1.COLOR_CODE = SOTSHIPC.COLOR_CODE"

            Fill_Records("WHTLOCB1", "", True, ASCMAIN1.sql)

        End If

        Sort_grdColumns(grdWHTLOCB1, "LOCATION_CODE")

        With grdWHTLOCB1.DisplayLayout.Bands(0)
            .Columns("CASES").Hidden = Not SummaryByLocation
            .Columns("BAR_CODE").Hidden = SummaryByLocation
            .Columns("CASE_BREAKDOWN").Hidden = Not SummaryByLocation
        End With
    End Sub

    Sub Setup_POTORDRX(STYLE_CODE As String, COLOR_CODE As String)
        Dim dvw As DataView = DirectCast(grdPOTORDRX.DataSource, DataTable).DefaultView
        dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
        Sort_grdColumns(grdPOTORDRX, "PO_SHIP_ETA")
        grdPOTORDRX.Text = "Inbound Shipments for " & STYLE_CODE & "-" & COLOR_CODE
    End Sub

    Sub Setup_WHTINSTS(STYLE_CODE As String, COLOR_CODE As String)
        Dim dvw As DataView = DirectCast(grdWHTINSTS.DataSource, DataTable).DefaultView
        dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
        Sort_grdColumns(grdWHTINSTS, "WAVE_INST_NO")
        grdWHTINSTS.Text = "Pick Details for " & STYLE_CODE & "-" & COLOR_CODE
    End Sub

    Sub Setup_SOTCART2(STYLE_CODE As String, COLOR_CODE As String)
        Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
        Sort_grdColumns(grdSOTCART2, "CUST_STORE_NO")
        grdSOTCART2.Text = "Carton Details for " & STYLE_CODE & "-" & COLOR_CODE
    End Sub

    Sub Create_WHTLOCBX(from_Load_Record As Boolean)
0:      '**************
        ' Note: WHTLOCM1.LOCATION_USE should always bereferenced as nvl(WHTLOCM1.LOCATION_USE,'A')
        '**************

        Setup_SOTSHIPC()

        If from_Load_Record And EntryMode = "N" Then
            'Note WHTWAVEP is dependant on Setup_SOTSHIPC()
            Fill_Records("WHTWAVEP", WHSE_CODE)
            If dst.Tables("WHTWAVEP").Rows.Count > 0 Then Exit Sub
        End If

        ASCDATA1.ExecuteSQL("Delete from " & WHTLOCBW)
        '& "   and NVL(WHTLOCB1.LOCATION_QTY,0) - NVL(WHTLOCB1.LOCATION_QTY_WAVE,0) > 0" & vbCrLf _

        ' THE ONLY DIFFERENCE BETWEEN THE 2 SQLS UNIONED BELOW SHOULD BE
        ' 1) and WHTLOCB1.BAR_CODE in (Select Distinct WHTLOCB1.BAR_CODE  vs    
        '    and WHTLOCB1.LOCATION_CODE in (Select Distinct WHTLOCB1.LOCATION_CODE

        'ASCMAIN1.sql = "Insert into " & WHTLOCBW & vbCrLf _
        '    & "Select WHTLOCB1.*, WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, '0' PPK" & vbCrLf _
        '    & ", WHTLOCM1.LOCATION_LOCKED, WHTLOCM1.LOCATION_NOT_WAVED, WHTLOCM1.LOCATION_USE" & vbCrLf _
        '    & " from WHTLOCB1,WHTBARC1,WHTBARC0,WHTLOCM1" _
        '    & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
        '    & "   and WHTLOCM1.WHSE_CODE = WHTLOCB1.WHSE_CODE" & vbCrLf _
        '    & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
        '    & "   and (WHTLOCM1.LOCATION_CODE = '" & rowICTWHSE1.Item("WHSE_LOC_REC") & "' or NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1')" & vbCrLf _
        '    & "   and WHTLOCB1.BAR_CODE in" & vbCrLf _
        '    & "    (Select Distinct WHTLOCB1.BAR_CODE from WHTLOCB1" & vbCrLf _
        '    & "      where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
        '    & "        and NVL(WHTLOCB1.LOCATION_QTY,0) > 0" & vbCrLf _
        '    & "        and (WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from " & SOTSHIPC & "))" & vbCrLf _
        '    & "   and NVL(WHTLOCB1.LOCATION_QTY,0) > 0" & vbCrLf _
        '    & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
        '    & "   and WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO" & vbCrLf _
        '    & " union " & vbCrLf _
        '    & "Select WHTLOCB1.*, WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, '0' PPK" & vbCrLf _
        '    & ", WHTLOCM1.LOCATION_LOCKED, WHTLOCM1.LOCATION_NOT_WAVED, WHTLOCM1.LOCATION_USE" & vbCrLf _
        '    & " from WHTLOCB1,WHTBARC1,WHTBARC0,WHTLOCM1" _
        '    & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
        '    & "   and WHTLOCM1.WHSE_CODE = WHTLOCB1.WHSE_CODE" & vbCrLf _
        '    & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
        '    & "   and (WHTLOCM1.LOCATION_CODE = '" & rowICTWHSE1.Item("WHSE_LOC_REC") & "' or NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1')" & vbCrLf _
        '    & "   and WHTLOCB1.LOCATION_CODE in" & vbCrLf _
        '    & "    (Select Distinct WHTLOCB1.LOCATION_CODE from WHTLOCB1" & vbCrLf _
        '    & "      where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
        '    & "        and NVL(WHTLOCB1.LOCATION_QTY,0) > 0" & vbCrLf _
        '    & "        and (WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from " & SOTSHIPC & "))" & vbCrLf _
        '    & "   and NVL(WHTLOCB1.LOCATION_QTY,0) > 0" & vbCrLf _
        '    & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
        '    & "   and WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO"
        'ASCDATA1.ExecuteSQL()

        ' we need all styles in a location to verify load or pallet pick
        ASCMAIN1.sql = "Insert into " & WHTLOCBW & vbCrLf _
            & "Select WHTLOCB1.*, WHTBARC1.LOAD_NO, WHTBARC0.LOAD_DATE, '0' PPK, WHTBARC1.PPK_CODE" & vbCrLf _
            & ", WHTLOCM1.LOCATION_LOCKED, WHTLOCM1.LOCATION_NOT_WAVED, WHTLOCM1.LOCATION_USE" & vbCrLf _
            & " from WHTLOCB1,WHTBARC1,WHTBARC0,WHTLOCM1" _
            & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "   and WHTLOCM1.WHSE_CODE = WHTLOCB1.WHSE_CODE" & vbCrLf _
            & "   and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
            & "   and NVL(WHTLOCM1.LOCATION_LOCKED,'0') <> '1'" & vbCrLf _
            & "   and (" & vbCrLf _
            & "        NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1'" & vbCrLf _
            & IIf(chkWaveFromReceiving.Checked, " or WHTLOCM1.LOCATION_CODE in (" & REC_LOCATIONS & ")", "") _
            & IIf(LOCATION_CODE_preferred <> "", " or WHTLOCM1.LOCATION_CODE = '" & LOCATION_CODE_preferred & "'", "") _
            & "       )" & vbCrLf _
            & IIf(cmbLOCATION_USE.Value <> "", " and nvl(WHTLOCM1.LOCATION_USE,'A') = '" & cmbLOCATION_USE.Value & "'", "") _
            & IIf(sqlPPK_CODEs <> "", " and WHTBARC1.PPK_CODE in (" & sqlPPK_CODEs & ")", "") _
            & IIf(chkNoPPK.Checked, " and WHTBARC1.PPK_CODE is Null", "") _
            & "   and (" & vbCrLf _
            & "       WHTLOCB1.BAR_CODE in" & vbCrLf _
            & "    (Select Distinct WHTLOCB1.BAR_CODE from WHTLOCB1" & vbCrLf _
            & "      where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "        and NVL(WHTLOCB1.LOCATION_QTY,0) > 0" & vbCrLf _
            & "        and (WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from " & SOTSHIPC & "))" & vbCrLf _
            & "   or " & vbCrLf _
            & "       WHTLOCB1.LOCATION_CODE in" & vbCrLf _
            & "    (Select Distinct WHTLOCB1.LOCATION_CODE from WHTLOCB1" & vbCrLf _
            & "      where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "        and NVL(WHTLOCB1.LOCATION_QTY,0) > 0" & vbCrLf _
            & "        and (WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from " & SOTSHIPC & "))" & vbCrLf _
            & "   ) " & vbCrLf _
            & "   and NVL(WHTLOCB1.LOCATION_QTY,0) > 0" & vbCrLf _
            & "   and WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" & vbCrLf _
            & "   and WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO"

        If from_Load_Record And (EntryMode = "N" And chkEmptyWave.Checked) Then
            ' DO NOT GATHER INVTY AVAILABILITY WHEN TRYING TO CREATE AN EMPTY WAVE
        Else
            ASCDATA1.ExecuteSQL()
        End If

        'ASCMAIN1.sql = "Delete from " & WHTLOCBW & " where LOCATION_CODE <> '" & rowICTWHSE1.Item("WHSE_LOC_REC") & "' and NVL(LOCATION_NOT_WAVED,'0') = '1'"



        ' THE BELOW SQL WAS LIFTED FROM DEPENDENT UPDATES
        ASCMAIN1.sql = "" _
               & "Begin" & vbCrLf _
               & " Declare Cursor C1 is" & vbCrLf _
               & "  Select WHTINST2.*,WHTINST1.WAVE_PICK_TYPE, WHTINST1.LOCATION_CODE" & vbCrLf _
               & "   from WHTINST1,WHTINST2" & vbCrLf _
               & "   where WHTINST1.WAVE_NO = '" & WAVE_NO & "'" & vbCrLf _
               & "     and WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
               & "     and WHTINST1.WAVE_INST_STATUS = '0';" & vbCrLf _
               & " Begin" & vbCrLf _
               & "  For R1 in C1 Loop" & vbCrLf _
               & "    Update " & WHTLOCBW & " WHTLOCB1 Set LOCATION_QTY_WAVE = NVL(LOCATION_QTY_WAVE,0) - NVL(R1.LOCATION_QTY_WAVE,0)" & vbCrLf _
               & "     where WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
               & "       and LOCATION_CODE = R1.LOCATION_CODE" & vbCrLf _
               & "       and BAR_CODE = R1.BAR_CODE" & vbCrLf _
               & "       and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
               & "  End Loop;" & vbCrLf _
               & " End;" & vbCrLf _
               & "End;"
        ASCDATA1.ExecuteSQL()

        'if we delete the above we lose ability to see if units are waved for that load
        'ASCMAIN1.sql = "Delete from " & WHTLOCBW & " where NVL(LOCATION_QTY,0) - NVL(LOCATION_QTY_WAVE,0) <= 0"
        'ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Delete from " & WHTLOCBW & " where LOCATION_LOCKED = '1'"
        'ASCDATA1.ExecuteSQL()

        For Each rowWHTINST1 As DataRow In dst.Tables("WHTINST1").Select("WAVE_INST_STATUS = '0'")
            For Each rowWHTINST2 As DataRow In rowWHTINST1.GetChildRows("WHTINST1_WHTINST2")
                Dim BAR_CODE As String = rowWHTINST2.Item("BAR_CODE")
                Dim LOCATION_CODE As String = rowWHTINST1.Item("LOCATION_CODE")
                Dim STYLE_CODE As String = rowWHTINST2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowWHTINST2.Item("COLOR_CODE")
                Dim LOCATION_QTY_WAVE As Int64 = Val(rowWHTINST2.Item("LOCATION_QTY_WAVE"))
                ASCMAIN1.sql = "Update " & WHTLOCBW & " WHTLOCB1" & vbCrLf _
                    & " Set LOCATION_QTY_WAVE = NVL(LOCATION_QTY_WAVE,0) + " & CStr(LOCATION_QTY_WAVE) & vbCrLf _
                    & "     where WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                    & "       and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                    & "       and BAR_CODE = '" & BAR_CODE & "'" & vbCrLf _
                    & "       and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf

                ASCDATA1.ExecuteSQL()
            Next
        Next

        ' NEED TO WEAN OFF USE OF PPK AND USE PPK_CODE INSTEAD

        ASCMAIN1.sql = "Update " & WHTLOCBW & " WHTLOCB1 Set PPK = '1'" & vbCrLf _
            & " where BAR_CODE in" & vbCrLf _
            & " (Select BAR_CODE from " & vbCrLf _
            & " (Select Distinct BAR_CODE, STYLE_CODE, COLOR_CODE from " & WHTLOCBW & ")" & vbCrLf _
            & " group by BAR_CODE having Count (*) > 1)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select WHTLOCBW.*" & vbCrLf _
            & " from " & WHTLOCBW & " WHTLOCBW, WHTLOCM1" & vbCrLf _
            & " where NVL(WHTLOCBW.LOCATION_QTY,0) - NVL(WHTLOCBW.LOCATION_QTY_WAVE,0) > 0" & vbCrLf _
            & "   and WHTLOCM1.LOCATION_CODE = WHTLOCBW.LOCATION_CODE" & vbCrLf _
            & "   and WHTLOCM1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "   and NVL(WHTLOCM1.LOCATION_LOCKED,'0') <> '1'" & vbCrLf _
            & "   and (NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1'" _
            & IIf(chkWaveFromReceiving.Checked, " or WHTLOCM1.LOCATION_CODE in (" & REC_LOCATIONS & ")", "") _
            & IIf(LOCATION_CODE_preferred <> "", " or WHTLOCM1.LOCATION_CODE = '" & LOCATION_CODE_preferred & "'", "") _
            & IIf(cmbLOCATION_USE.Value <> "", " and nvl(WHTLOCM1.LOCATION_USE,'A') = '" & cmbLOCATION_USE.Value & "'", "") _
            & ")"

        Fill_Records("WHTLOCBW", "", True, ASCMAIN1.sql)
    End Sub

    Sub Create_Wave_Instructions(from_Load_Record As Boolean)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Waving ...")

        Dim CartonUsed As Int32

        LOCATION_CODE_preferred = Replace(Replace(Absx1.txtFor("LOCATION_CODE").Text, ";", ""), "'", "")
        If LOCATION_CODE_preferred <> "" Then
            Dim rowWHTLOCM1_preferred As DataRow = LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE_preferred})
            If rowWHTLOCM1_preferred Is Nothing Then
                MsgBox("Invalid Preferred Location (" & LOCATION_CODE_preferred & ")", MsgBoxStyle.OkOnly, "Cannot Wave")
                LOCATION_CODE_preferred = ""
                Exit Sub
            End If
        End If

        Create_WHTLOCBX(from_Load_Record) ' Create Supply Records based on Styles Listed in Demand table WHTWAVE2

        For Each TABLE_NAME As String In New String() {"WHTWAVE2", "WHTWAVE2_SUB"}

            Dim WAVE_SUB As String = IIf(TABLE_NAME = "WHTWAVE2", "0", "1")

            For Each rowWHTWAVE2 As DataRow In dst.Tables(TABLE_NAME).Select("WAVE_QTY_LEFT > 0", "STYLE_CODE,COLOR_CODE")

                Dim STYLE_CODE As String = rowWHTWAVE2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowWHTWAVE2.Item("COLOR_CODE")
                If TABLE_NAME = "WHTWAVE2_SUB" Then
                    STYLE_CODE = rowWHTWAVE2.Item("STYLE_CODE_SUB")
                    COLOR_CODE = rowWHTWAVE2.Item("COLOR_CODE_SUB")
                End If

                ' If STYLE_CODE = "K4999" And COLOR_CODE = "AST" Then Stop

                ASCMAIN1.Progress("-", STYLE_CODE & "-" & COLOR_CODE)

                Dim PICK_QTY As Int64 = Val(rowWHTWAVE2.Item("PICK_QTY") & "")
                Dim WAVE_QTY As Int64 = Val(rowWHTWAVE2.Item("WAVE_QTY") & "")
                Dim WAVE_QTY_LEFT As Int64 = Val(rowWHTWAVE2.Item("WAVE_QTY_LEFT") & "")
                Dim WAVE_LNO As Int64 = Val(rowWHTWAVE2.Item("WAVE_LNO") & "")

                Dim QTY_TO_PICK_remaining As Int64 = WAVE_QTY_LEFT ' PICK_QTY - WAVE_QTY
                Dim QTY_TO_PICK_to_start As Int64 = QTY_TO_PICK_remaining

                If QTY_TO_PICK_remaining > 0 Then

                    ASCMAIN1.sql = "Select WHTLOCBW.LOAD_NO, WHTLOCBW.LOCATION_CODE" & vbCrLf _
                        & ", WHTLOCBW.STYLE_CODE, WHTLOCBW.COLOR_CODE" & vbCrLf _
                        & ", TO_CHAR(WHTLOCBW.LOAD_DATE,'YYYYQ') YYYYQ" & vbCrLf _
                        & ", COUNT (*) CASES" & vbCrLf _
                        & ", MIN (PPK_CODE) PPK_CODE, MAX (PPK_CODE) PPK_CODE2" & vbCrLf _
                        & ", SUM (NVL(WHTLOCBW.LOCATION_QTY,0) - NVL(WHTLOCBW.LOCATION_QTY_WAVE,0)) QTY" & vbCrLf _
                        & ", SUM (NVL(WHTLOCBW.LOCATION_QTY_WAVE,0)) QTY_WAVE" & vbCrLf _
                        & " from " & WHTLOCBW & " WHTLOCBW, WHTLOCM1" & vbCrLf _
                        & " where NVL(WHTLOCBW.LOCATION_QTY,0) - NVL(WHTLOCBW.LOCATION_QTY_WAVE,0) > 0" & vbCrLf _
                        & "   and WHTLOCBW.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                        & "   and WHTLOCBW.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                        & "   and NVL(WHTLOCBW.LOCATION_QTY,0) - NVL(WHTLOCBW.LOCATION_QTY_WAVE,0) > 0" & vbCrLf _
                        & "   and WHTLOCM1.LOCATION_CODE = WHTLOCBW.LOCATION_CODE" & vbCrLf _
                        & "   and WHTLOCM1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                        & IIf(optPPK.Value = "A", "", "   and WHTLOCBW.PPK = " & IIf(optPPK.Value = "N", "'0'" & vbCrLf, "1")) _
                        & IIf(LOCATION_CODE_preferred = "", "", "   and WHTLOCBW.LOCATION_CODE = '" & LOCATION_CODE_preferred & "'" & vbCrLf) _
                        & IIf(cmbLOCATION_USE.Value = "", "", " and nvl(WHTLOCM1.LOCATION_USE,'A') = '" & cmbLOCATION_USE.Value & "'" & vbCrLf) _
                        & "   and NVL(WHTLOCM1.LOCATION_LOCKED,'0') <> '1'" & vbCrLf _
                        & "   and (" _
                        & " NVL(WHTLOCM1.LOCATION_NOT_WAVED,'0') <> '1'" _
                        & IIf(chkWaveFromReceiving.Checked, " or WHTLOCM1.LOCATION_CODE in (" & REC_LOCATIONS & ")", "") _
                        & IIf(LOCATION_CODE_preferred <> "", " or WHTLOCM1.LOCATION_CODE = '" & LOCATION_CODE_preferred & "'", "") _
                        & ")" & vbCrLf _
                        & " group by WHTLOCBW.LOAD_NO, WHTLOCBW.LOCATION_CODE, WHTLOCBW.STYLE_CODE, WHTLOCBW.COLOR_CODE, TO_CHAR(WHTLOCBW.LOAD_DATE,'YYYYQ')"

                    For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYQ")

                        ' EACH OF THESE ROWS IS AN ENTIRE LOAD
                        ' BUT THERE MAY BE OTHER STYLES ON THE LOAD OR LOCATION - BEWARE

                        Dim QTY As Int64 = Val(row.Item("QTY") & "")
                        Dim QTY_WAVE As Int64 = Val(row.Item("QTY_WAVE") & "")
                        Dim CASES As Int64 = Val(row.Item("CASES") & "")
                        Dim LOAD_NO As String = row.Item("LOAD_NO")
                        Dim LOCATION_CODE As String = row.Item("LOCATION_CODE")
                        Dim PPK_CODE As String = row.Item("PPK_CODE") & ""
                        Dim sqlPPK As String = ""
                        If PPK_CODE <> "" Then
                            If PPK_CODE <> row.Item("PPK_CODE2") Then
                                PPK_CODE = "" ' PPK_CODE is not reliable - there are more than 1 in this batch
                            Else
                                sqlPPK = " and NVL(PPK_CODE,'?') <> '" & PPK_CODE & "'"
                            End If
                        End If

                        If LOCATION_CODE = rowICTWHSE1.Item("WHSE_LOC_REC") Then
                            ASCMAIN1.sql = "Select Count (*) from " & WHTLOCBW & vbCrLf _
                                & " where LOCATION_QTY > 0" & vbCrLf _
                                & "   and LOAD_NO = '" & LOAD_NO & "'" & vbCrLf _
                                & "   and (STYLE_CODE <> '" & STYLE_CODE & "' or COLOR_CODE <> '" & COLOR_CODE & "' or LOCATION_QTY_WAVE <> 0)" _
                                & sqlPPK
                        Else
                            ASCMAIN1.sql = "Select Count (*) from " & WHTLOCBW & vbCrLf _
                                & " where LOCATION_QTY > 0" & vbCrLf _
                                & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                                & "   and (STYLE_CODE <> '" & STYLE_CODE & "' or COLOR_CODE <> '" & COLOR_CODE & "' or LOCATION_QTY_WAVE <> 0 or LOAD_NO <> '" & LOAD_NO & "')" _
                                & sqlPPK
                        End If

                        ' NEED TO REVISIT THIS FOR PPKS

                        Dim other_styles As Boolean = (Val(ASCDATA1.GetDataValue & "") <> 0)

                        ' can do a pallet pick if:
                        ' 1) the total qty on the pallet is <= qty needed
                        ' 2) there are no other styles on the pallet
                        ' 3) there are no wave commitments reaching into this pallet

                        If chkForcePalletPick.Checked Then other_styles = False

                        Dim rowWHTINST1 As DataRow = Nothing
                        If QTY <= QTY_TO_PICK_remaining And Not other_styles And QTY_WAVE = 0 Then
                            rowWHTINST1 = Create_WHTINST1_row(row, "L", CASES, QTY, WAVE_LNO, WAVE_SUB)
                            If optPPK.Value = "P" Then
                                QTY = ppkLoad(rowWHTINST1.Item("WAVE_INST_NO"), LOAD_NO, STYLE_CODE, COLOR_CODE, QTY_TO_PICK_remaining)
                            End If
                            QTY_TO_PICK_remaining -= QTY

                        Else
                            Dim sqlw_exact As String = "" _
                                & "     STYLE_CODE = '" & STYLE_CODE & "'" _
                                & " and COLOR_CODE = '" & COLOR_CODE & "'" _
                                & IIf(LOCATION_CODE_preferred = "", "", "   and LOCATION_CODE = '" & LOCATION_CODE_preferred & "'" & vbCrLf) _
                                & " and ISNULL(LOCATION_QTY,0) > 0" _
                                & " and ISNULL(LOCATION_QTY_WAVE,0) = 0"

                            Dim sqlw As String = "" _
                                & "     LOAD_NO = '" & LOAD_NO & "'" _
                                & " and STYLE_CODE = '" & STYLE_CODE & "'" _
                                & " and COLOR_CODE = '" & COLOR_CODE & "'" _
                                & " and ISNULL(LOCATION_QTY,0) > 0" _
                                & " and ISNULL(LOCATION_QTY,0) - ISNULL(LOCATION_QTY_WAVE,0) > 0"

                            'If STYLE_CODE = "28643KM" And COLOR_CODE = "001" Then Stop
                            For Each rowWHTLOCBW As DataRow In dst.Tables("WHTLOCBW").Select(sqlw, "LOCATION_CODE, LOAD_NO, LOCATION_QTY DESC, LOCATION_QTY_WAVE")
                                ' EACH OF THESE ROWS IS A CASE

                                ' can do a case pick if:
                                ' 1) the total qty in the case is <= qty needed
                                ' 3) there are no wave commitments reaching into this case

                                'ppkload assigns loads from oracle for speed, need to spik bar-codes already assigned there.
                                If optPPK.Value = "P" Then
                                    CartonUsed = dst.Tables("WHTINST2").Compute("COUNT(BAR_CODE)", "BAR_CODE = '" & rowWHTLOCBW.Item("BAR_CODE") & "'")
                                    If CartonUsed > 0 Then
                                        Continue For
                                    End If
                                End If

                                Dim rowWHTLOCBW_exact() As DataRow = dst.Tables("WHTLOCBW").Select(sqlw_exact & " and LOCATION_QTY = " & CStr(QTY_TO_PICK_remaining) & " and ISNULL(LOCATION_QTY_WAVE,0) = 0", "LOCATION_CODE, LOAD_NO, LOCATION_QTY DESC, LOCATION_QTY_WAVE")
                                If rowWHTLOCBW_exact.Length <> 0 Then
                                    QTY = Val(rowWHTLOCBW_exact(0).Item("LOCATION_QTY") & "") - Val(rowWHTLOCBW_exact(0).Item("LOCATION_QTY_WAVE") & "")
                                    QTY_WAVE = Val(rowWHTLOCBW_exact(0).Item("LOCATION_QTY_WAVE") & "")  ' S/B 0 FOR CASE PICK

                                    If QTY_WAVE = 0 And rowWHTINST1 Is Nothing OrElse
                                        (rowWHTINST1.Item("LOCATION_CODE") <> rowWHTLOCBW_exact(0).Item("LOCATION_CODE") Or
                                         rowWHTINST1.Item("LOAD_NO") <> rowWHTLOCBW_exact(0).Item("LOAD_NO")) Then
                                        rowWHTINST1 = Create_WHTINST1_row(rowWHTLOCBW_exact(0), "C", 1, QTY, WAVE_LNO, WAVE_SUB)
                                        If optPPK.Value = "P" Then
                                            QTY = ppkLoad(rowWHTINST1.Item("WAVE_INST_NO"), LOAD_NO, STYLE_CODE, COLOR_CODE, QTY_TO_PICK_remaining)
                                        End If
                                    Else
                                        Create_WHTINST2_row(rowWHTINST1.Item("WAVE_INST_NO"), rowWHTLOCBW_exact(0), QTY, WAVE_LNO)
                                    End If
                                    QTY_TO_PICK_remaining -= QTY
                                    Exit For
                                End If

                                QTY = Val(rowWHTLOCBW.Item("LOCATION_QTY") & "") - Val(rowWHTLOCBW.Item("LOCATION_QTY_WAVE") & "")
                                QTY_WAVE = Val(rowWHTLOCBW.Item("LOCATION_QTY_WAVE") & "")

                                If QTY_WAVE = 0 And QTY <= QTY_TO_PICK_remaining And QTY_WAVE = 0 Then
                                    If rowWHTINST1 Is Nothing OrElse
                                        (rowWHTINST1.Item("LOCATION_CODE") <> rowWHTLOCBW.Item("LOCATION_CODE") Or
                                         rowWHTINST1.Item("LOAD_NO") <> rowWHTLOCBW.Item("LOAD_NO")) Then
                                        rowWHTINST1 = Create_WHTINST1_row(rowWHTLOCBW, "C", 1, QTY, WAVE_LNO, WAVE_SUB)
                                        If optPPK.Value = "P" Then
                                            QTY = ppkLoad(rowWHTINST1.Item("WAVE_INST_NO"), LOAD_NO, STYLE_CODE, COLOR_CODE, QTY_TO_PICK_remaining)
                                        End If
                                    Else
                                        Create_WHTINST2_row(rowWHTINST1.Item("WAVE_INST_NO"), rowWHTLOCBW, QTY, WAVE_LNO)
                                    End If

                                    QTY_TO_PICK_remaining -= QTY

                                Else
                                    Dim QTY_TO_PICK As Decimal = QTY
                                    If QTY_TO_PICK_remaining < QTY Then
                                        QTY_TO_PICK = QTY_TO_PICK_remaining
                                    End If
                                    If chkNoUnitPick.Checked = False Then
                                        'Skip Unit Picking requested 2/24/2021
                                        Create_WHTINST1_row(rowWHTLOCBW, "U", 0, QTY_TO_PICK, WAVE_LNO, WAVE_SUB)
                                    End If
                                    'QTY_TO_PICK_remaining = 0
                                    QTY_TO_PICK_remaining -= QTY_TO_PICK
                                End If

                                ASCMAIN1.Progress("-", Val(PICK_QTY) & "/" & Val(QTY_TO_PICK_remaining))
                                If QTY_TO_PICK_remaining <= 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        If QTY_TO_PICK_remaining <= 0 Then
                            Exit For
                        End If
                    Next
                End If

                WAVE_QTY += QTY_TO_PICK_to_start - QTY_TO_PICK_remaining
                rowWHTWAVE2.Item("WAVE_QTY") = WAVE_QTY

                Dim WAVE_QTY_OPEN As Int64 = Val(rowWHTWAVE2.Item("WAVE_QTY_OPEN") & "")
                WAVE_QTY_OPEN += QTY_TO_PICK_to_start - QTY_TO_PICK_remaining
                rowWHTWAVE2.Item("WAVE_QTY_OPEN") = WAVE_QTY_OPEN

                Update_Record_TDA("WHTLOCBW")
            Next
        Next

        chkForcePalletPick.Checked = False
        chkNoUnitPick.Checked = False
        optPPK.Value = "A"

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Create_WHTINST1_row(row As DataRow, WAVE_PICK_TYPE As String, CASES As Int64, QTY As Int64, WAVE_LNO As Int64, WAVE_SUB As String) As DataRow

        Dim WAVE_INST_NO As String = ""
        If preview Then
            WAVE_INST_NO_ctr += 1
            WAVE_INST_NO = Format(WAVE_INST_NO_ctr, "0000000000")
        Else
            WAVE_INST_NO = ASCMAIN1.Next_Control_No("WHTINST1.WAVE_INST_NO")
        End If

        Dim LOAD_NO As String = row.Item("LOAD_NO")
        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
        Dim COLOR_CODE As String = row.Item("COLOR_CODE")

        Dim rowWHTINST1 As DataRow = dst.Tables("WHTINST1").NewRow
        With rowWHTINST1
            .Item("WAVE_INST_NO") = WAVE_INST_NO
            .Item("WAVE_NO") = WAVE_NO
            .Item("LOCATION_CODE") = row.Item("LOCATION_CODE")
            .Item("LOAD_NO") = LOAD_NO
            .Item("WAVE_PICK_TYPE") = WAVE_PICK_TYPE
            .Item("WAVE_INST_STATUS") = "0"
            .Item("WAVE_LNO") = WAVE_LNO
            .Item("WAVE_SUB") = WAVE_SUB
        End With
        dst.Tables("WHTINST1").Rows.Add(rowWHTINST1)

        If optPPK.Value = "P" Then
            ' don't create details here use special function
        Else
            If WAVE_PICK_TYPE = "L" Then

                Dim sqlw As String = "" _
                    & "     LOAD_NO = '" & LOAD_NO & "'" _
                    & " and STYLE_CODE = '" & STYLE_CODE & "'" _
                    & " and COLOR_CODE = '" & COLOR_CODE & "'" _
                    & " and ISNULL(LOCATION_QTY,0) - ISNULL(LOCATION_QTY_WAVE,0) >= 0"

                For Each rowWHTLOCBW As DataRow In dst.Tables("WHTLOCBW").Select(sqlw) ' ASCDATA1.GetDataTable.Select("")
                    Dim QTY_AVAIL As Int64 = Val(rowWHTLOCBW.Item("LOCATION_QTY") & "") - Val(rowWHTLOCBW.Item("LOCATION_QTY_WAVE") & "")
                    Create_WHTINST2_row(WAVE_INST_NO, rowWHTLOCBW, QTY_AVAIL, WAVE_LNO)
                Next
            Else
                ' row is WHTLOCBW
                Create_WHTINST2_row(WAVE_INST_NO, row, QTY, WAVE_LNO)
            End If
        End If

        Return rowWHTINST1

    End Function

    Function Create_WHTINST2_row(WAVE_INST_NO As String, row As DataRow, QTY As Int64, WAVE_LNO As Int64, Optional recurse_PPK As Boolean = True)
        If QTY < 0 Then Return Nothing
        If QTY = 0 Then Return Nothing

        Dim BAR_CODE As String = row.Item("BAR_CODE")
        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
        Dim COLOR_CODE As String = row.Item("COLOR_CODE")

        If ASCMAIN1.Running_in_VS And BAR_CODE = "51450819" Then Stop

        Dim rowWHTINST2 As DataRow = dst.Tables("WHTINST2").NewRow
        With rowWHTINST2
            .Item("WAVE_INST_NO") = WAVE_INST_NO
            .Item("BAR_CODE") = BAR_CODE
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("LOCATION_QTY_WAVE") = QTY
            .Item("WAVE_LNO") = WAVE_LNO
            .Item("BAR_CODE_ORIG") = BAR_CODE
        End With
        dst.Tables("WHTINST2").Rows.Add(rowWHTINST2)
        row.Item("LOCATION_QTY_WAVE") = Val(row.Item("LOCATION_QTY_WAVE") & "") + QTY

        'Dim rowWHTWAVE2 As DataRow = dst.Tables("")
        'rowWHTWAVE2.Item("WAVE_QTY_OPEN") = Val(rowWHTWAVE2.Item("WAVE_QTY_OPEN")) + QTY

        'ASCMAIN1.sql = "Select * from WHTLOCB1 where LOCATION_CODE = '" & row.Item("LOCATION_CODE") & "' and BAR_CODE = '" & BAR_CODE & "'"
        'Dim ROWCHECK As DataRow = ASCDATA1.GetDataRow
        'If Val(ROWCHECK.Item("LOCATION_QTY_WAVE") & "") <> 0 Then Stop

        ' If WAVE_PICK_TYPE = "L" OR  If WAVE_PICK_TYPE = "C" Then FOR EACH PPK LOADED INTO WHTINST2
        '  - also write the other PPK components to WHTINST2

        If Not recurse_PPK Then ' WE ARE IN A PPK RECURSION
            'This code is not in use since new code handles pre-pack in loop w/o recursion 
            'it also fixes problem of creating instructions that don't point to a wave.
            'Dim rowWHTWAVE2s() As DataRow = dst.Tables("WHTWAVE2").Select("WAVE_NO = '" & WAVE_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
            'Dim rowWHTWAVE2 As DataRow = Nothing
            'If rowWHTWAVE2s.Length = 0 Then
            '    'Stop ' WE ARE RECURSING TO A STYLE THAT IS NOT ON THE ORDER
            '    MsgBox("Style: " & STYLE_CODE & " Clr: " & COLOR_CODE & ", is not part of the shipment", vbOKOnly, "PrePack Problem")
            '    WAVE_LNO = Val(dst.Tables("WHTWAVE2").Compute("MAX(WAVE_LNO)", "") & "") + 1
            '    rowWHTWAVE2 = dst.Tables("WHTWAVE2").Rows.Add(New String() {WAVE_NO, WAVE_LNO, STYLE_CODE, COLOR_CODE})
            '    rowWHTWAVE2.Item("WAVE_QTY") = Val(rowWHTWAVE2.Item("WAVE_QTY") & "") + QTY
            '    rowWHTWAVE2.Item("WAVE_QTY_OPEN") = Val(rowWHTWAVE2.Item("WAVE_QTY_OPEN") & "") + QTY
            'Else
            '    rowWHTWAVE2 = rowWHTWAVE2s(0)
            '    rowWHTWAVE2.Item("WAVE_QTY") = Val(rowWHTWAVE2.Item("WAVE_QTY") & "") + QTY
            '    rowWHTWAVE2.Item("WAVE_QTY_OPEN") = Val(rowWHTWAVE2.Item("WAVE_QTY_OPEN") & "") + QTY
            'End If
        Else
            If row.Item("PPK") & "" = "1" Then
                Dim WAVE_PICK_TYPE As String = rowWHTINST2.GetParentRow("WHTINST1_WHTINST2").Item("WAVE_PICK_TYPE") & ""
                If WAVE_PICK_TYPE = "L" Or WAVE_PICK_TYPE = "C" Then
                    Dim sqlPPK As String = "" _
                        & "     BAR_CODE = '" & BAR_CODE & "'" _
                        & " and (STYLE_CODE > '" & STYLE_CODE & "'" _
                        & "  or (STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE > '" & COLOR_CODE & "'))"

                    For Each rowPPK As DataRow In dst.Tables("WHTLOCBW").Select(sqlPPK)
                        Dim QTY_AVAIL As Int64 = Val(rowPPK.Item("LOCATION_QTY") & "") - Val(rowPPK.Item("LOCATION_QTY_WAVE") & "")

                        ' find the WAVE_LNO for the SC (if it is a sub, then we have a design issue in that we won't know which demand record, ie which WAVE_LNO, to relate to)
                        Dim sqlw As String = "STYLE_CODE = '" & rowPPK.Item("STYLE_CODE") & "' and COLOR_CODE = '" & rowPPK.Item("COLOR_CODE") & "' and STYLE_CODE_SUB is Null"
                        Dim rowsPPK() As DataRow = dst.Tables("WHTWAVE2").Select(sqlw) ' THIS TABLE HAS NO SUBS IN IT ANYWAY
                        Dim rowWHTWAVE2 As DataRow = Nothing
                        Dim WAVE_LNO_PPK As Int64 = 0
                        If rowsPPK.Length <> 1 Then
                            If rowsPPK.Length = 0 Then
                                MsgBox("Style: " & STYLE_CODE & " Clr: " & COLOR_CODE & ", is not part of the shipment", vbOKOnly, "PrePack Problem")
                                WAVE_LNO_PPK = Val(dst.Tables("WHTWAVE2").Compute("MAX(WAVE_LNO)", "") & "") + 1
                                If WAVE_LNO_PPK < Val(dst.Tables("WHTWAVE2_SUB").Compute("MAX(WAVE_LNO)", "") & "") + 1 Then
                                    WAVE_LNO_PPK = Val(dst.Tables("WHTWAVE2_SUB").Compute("MAX(WAVE_LNO)", "") & "") + 1
                                End If
                                rowWHTWAVE2 = dst.Tables("WHTWAVE2").Rows.Add(New String() {WAVE_NO, WAVE_LNO_PPK, rowPPK.Item("STYLE_CODE"), rowPPK.Item("COLOR_CODE")})
                            Else
                                MsgBox("Issue with PPK and Wave Demand Table - please call ABS")
                                Stop
                            End If
                        Else
                            WAVE_LNO_PPK = rowsPPK(0).Item("WAVE_LNO")
                            rowWHTWAVE2 = rowsPPK(0)
                        End If
                        'Create_WHTINST2_row(WAVE_INST_NO, rowPPK, QTY_AVAIL, WAVE_LNO_PPK, False)
                        'Dim stopWatch As Stopwatch = stopWatch.StartNew()
                        rowWHTINST2 = dst.Tables("WHTINST2").NewRow
                        With rowWHTINST2
                            .Item("WAVE_INST_NO") = WAVE_INST_NO
                            .Item("BAR_CODE") = BAR_CODE
                            .Item("STYLE_CODE") = rowPPK.Item("STYLE_CODE")
                            .Item("COLOR_CODE") = rowPPK.Item("COLOR_CODE")
                            .Item("LOCATION_QTY_WAVE") = QTY_AVAIL
                            .Item("WAVE_LNO") = WAVE_LNO_PPK
                            .Item("BAR_CODE_ORIG") = BAR_CODE
                        End With
                        dst.Tables("WHTINST2").Rows.Add(rowWHTINST2)
                        rowPPK.Item("LOCATION_QTY_WAVE") = Val(row.Item("LOCATION_QTY_WAVE") & "") + QTY_AVAIL
                        rowWHTWAVE2.Item("WAVE_QTY") = Val(rowWHTWAVE2.Item("WAVE_QTY") & "") + QTY_AVAIL
                        rowWHTWAVE2.Item("WAVE_QTY_OPEN") = Val(rowWHTWAVE2.Item("WAVE_QTY_OPEN") & "") + QTY_AVAIL
                        'stopWatch.Stop()
                        'If (dst.Tables("WHTINST2").Rows.Count Mod 1000) = 3 Then
                        '    Debug.WriteLine(dst.Tables("WHTINST2").Rows.Count & " Inserts " & stopWatch.ElapsedMilliseconds)
                        'End If
                    Next
                End If
            End If
        End If

        Return rowWHTINST2

    End Function

    Private Function ppkLoad(WAVE_INST_NO As String, LOAD_NO As String, STYLE_CODE As String, COLOR_CODE As String, QTY_TO_PICK_remaining As Int64) As Integer
        Manage_Expressions("Remove")
        'load all pre-pack styles fast into instructions

        ASCMAIN1.sql = " select '" & WAVE_INST_NO & "' WAVE_INST_NO, BAR_CODE, STYLE_CODE, COLOR_CODE, LOCATION_QTY LOCATION_QTY_WAVE, 0 LOCATION_QTY_PICK, 0 WAVE_LNO, BAR_CODE BAR_CODE_ORIG" & vbCrLf _
            & " from " & WHTLOCBW & " WHTLOCBW " & vbCrLf _
            & " where BAR_CODE in (  " & vbCrLf _
            & "select BAR_CODE from (  " & vbCrLf _
            & "select WHTLOCBW.*, sum(LOCATION_QTY) over (order by BAR_CODE rows between unbounded preceding and current row) as total  from " & WHTLOCBW & " WHTLOCBW  " & vbCrLf _
            & " where LOAD_NO = '" & LOAD_NO & "'" & vbCrLf _
            & " and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & " and COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & " and nvl(LOCATION_QTY_WAVE,0) = 0" & vbCrLf _
            & ") t  " & vbCrLf _
            & "where t.total <= " & QTY_TO_PICK_remaining & ")"
        Fill_Records("WHTINST2", "", False, ASCMAIN1.sql)

        Dim qty As Int64 = dst.Tables("WHTINST2").Compute("sum(LOCATION_QTY_WAVE)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and PPK_COUNTED is null")
        'QTY_TO_PICK_remaining = QTY_TO_PICK_remaining - (p_qty - qty)

        'using the view to get distinct rows within the filter
        Dim view As DataView = New DataView(dst.Tables("WHTINST2"))
        view.RowFilter = "WAVE_INST_NO = '" & WAVE_INST_NO & "'"
        Dim tblStyleColor As DataTable = view.ToTable(True, "STYLE_CODE", "COLOR_CODE")

        Dim rowWAVE As DataRow = Nothing
        For Each row As DataRow In tblStyleColor.Select("")
            If rowWAVE Is Nothing OrElse rowWAVE.Item("STYLE_CODE") <> row.Item("STYLE_CODE") OrElse rowWAVE.Item("COLOR_CODE") <> row.Item("COLOR_CODE") Then
                Dim rowsWave() As DataRow = dst.Tables("WHTWAVE2").Select("STYLE_CODE = '" & row.Item("STYLE_CODE") & "' and COLOR_CODE = '" & row.Item("COLOR_CODE") & "' and STYLE_CODE_SUB is Null")
                If rowsWave.Length > 0 Then
                    rowWAVE = rowsWave(0)
                Else
                    MsgBox("Style: " & STYLE_CODE & " Clr: " & COLOR_CODE & ", is not part of the shipment", vbOKOnly, "PrePack Problem")
                    Dim WAVE_LNO As Int64 = Val(dst.Tables("WHTWAVE2").Compute("MAX(WAVE_LNO)", "") & "") + 1
                    If WAVE_LNO < Val(dst.Tables("WHTWAVE2_SUB").Compute("MAX(WAVE_LNO)", "") & "") + 1 Then
                        WAVE_LNO = Val(dst.Tables("WHTWAVE2_SUB").Compute("MAX(WAVE_LNO)", "") & "") + 1
                    End If
                    rowWAVE = dst.Tables("WHTWAVE2").Rows.Add(New String() {WAVE_NO, WAVE_LNO, row.Item("STYLE_CODE"), row.Item("COLOR_CODE")})
                End If
            End If
            Dim wQTY As Int64 = 0
            For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("STYLE_CODE = '" & row.Item("STYLE_CODE") & "' and COLOR_CODE = '" & row.Item("COLOR_CODE") & "'  and PPK_COUNTED is null")
                rowWHTINST2.Item("WAVE_LNO") = rowWAVE.Item("WAVE_LNO")
                rowWHTINST2.Item("PPK_COUNTED") = "1"
                wQTY = wQTY + Val(rowWHTINST2.Item("LOCATION_QTY_WAVE") & "")
                For Each rowPPK As DataRow In dst.Tables("WHTLOCBW").Select("BAR_CODE = '" & rowWHTINST2.Item("BAR_CODE") & "' and STYLE_CODE = '" & rowWHTINST2.Item("STYLE_CODE") & "' and COLOR_CODE = '" & rowWHTINST2.Item("COLOR_CODE") & "'")
                    rowPPK.Item("LOCATION_QTY_WAVE") = Val(rowPPK.Item("LOCATION_QTY_WAVE") & "") + Val(rowWHTINST2.Item("LOCATION_QTY_WAVE") & "")
                Next
            Next
            If Not (rowWAVE.Item("STYLE_CODE") = STYLE_CODE And rowWAVE.Item("COLOR_CODE") = COLOR_CODE) Then
                'the original wave record is updated outside of this call
                rowWAVE.Item("WAVE_QTY") = Val(rowWAVE.Item("WAVE_QTY") & "") + wQTY
                rowWAVE.Item("WAVE_QTY_OPEN") = Val(rowWAVE.Item("WAVE_QTY_OPEN") & "") + wQTY
            End If
        Next
        Manage_Expressions("Restore")
        Return qty
    End Function

    Private Sub grdWHTWAVES_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTWAVES.InitializeRow
        With e.Row.Cells("WAVE_PICK_TYPE")
            If .Value = "L" Then
                e.Row.Cells("WAVE_STAT_TYPE").Appearance.ForeColor = Drawing.Color.Blue
            ElseIf .Value = "C" Then
                e.Row.Cells("WAVE_STAT_TYPE").Appearance.ForeColor = Drawing.Color.Green
            ElseIf .Value = "U" Then
                e.Row.Cells("WAVE_STAT_TYPE").Appearance.ForeColor = Drawing.Color.DarkOrange
            End If
        End With

        With e.Row.Cells("PICK")
            If Val(e.Row.Cells("WAVE").Value & "") <> Val(.Value & "") And Val(.Value & "") <> 0 Then
                .Appearance.ForeColor = Drawing.Color.Red
            Else
                .Appearance.ForeColor = Drawing.Color.Empty
            End If
        End With

    End Sub

    Private Sub grdWHTWAVE2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTWAVE2.InitializeRow

        If e.Row.Band.Index = 0 Then
            With e.Row.Cells("WAVE_QTY_LEFT")
                If Val(.Value & "") <> 0 Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText = "Warning - Re-Waving Required"
                Else
                    .Appearance.ForeColor = Drawing.Color.Empty
                    .ToolTipText = ""
                End If
            End With

            With e.Row.Cells("WAVE_QTY_CONF")
                If Val(.Value & "") > Val(e.Row.Cells("PICK_QTY").Value & "") Then
                    .Appearance.ForeColor = Drawing.Color.Green
                    .ToolTipText = "Warning - Qty Conf exceeds the original Qty Released to Pick"
                ElseIf Val(.Value & "") < Val(e.Row.Cells("PICK_QTY").Value & "") Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText = "Warning - Qty Conf is less than the original Qty Released to Pick"
                Else
                    .Appearance.ForeColor = Drawing.Color.Empty
                    .ToolTipText = ""
                End If
            End With

            With e.Row.Cells("WAVE_QTY_PACK")
                If Val(.Value & "") > Val(e.Row.Cells("WAVE_QTY_CONF").Value & "") Then
                    .ToolTipText = "Warning - Qty Packed exceeds the Qty to be Confirmed"
                    .Appearance.BackColor = Drawing.Color.Green
                    .Appearance.ForeColor = Drawing.Color.White
                ElseIf Val(.Value & "") < Val(e.Row.Cells("WAVE_QTY_CONF").Value & "") Then
                    .ToolTipText = "Warning - Qty Packed is less than the Qty to be Confirmed"
                    .Appearance.BackColor = Drawing.Color.Red
                    .Appearance.ForeColor = Drawing.Color.White
                Else
                    .ToolTipText = ""
                    .Appearance.BackColor = Drawing.Color.Empty
                    .Appearance.ForeColor = Drawing.Color.Empty
                End If
            End With
        End If


        If e.Row.Band.Index = 1 Then
            If e.Row.Cells("STYLE_CODE_SUB").Value & e.Row.Cells("COLOR_CODE_SUB").Value & "" = e.Row.Cells("STYLE_CODE").Value & e.Row.Cells("COLOR_CODE").Value Then
                e.Row.Cells("STYLE_CODE_SUB").Appearance.BackColor = Drawing.Color.Yellow
                e.Row.Cells("COLOR_CODE_SUB").Appearance.BackColor = Drawing.Color.Yellow
            Else
                e.Row.Cells("STYLE_CODE_SUB").Appearance.BackColor = Drawing.Color.Empty
                e.Row.Cells("COLOR_CODE_SUB").Appearance.BackColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Private Sub grdWHTINST1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWHTINST1.AfterRowsDeleted
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Deleting Pick Instructions")

        For Each rowWHTWAVED As DataRow In dst.Tables("WHTWAVED").Select("")
            Dim STYLE_CODE As String = rowWHTWAVED.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTWAVED.Item("COLOR_CODE")
            Dim WAVE_LNO As Int64 = Val(rowWHTWAVED.Item("WAVE_LNO") & "")
            Dim WAVE_LNO_LINK As Int64 = Val(rowWHTWAVED.Item("WAVE_LNO_LINK") & "")
            Dim WAVE_SUB As String = rowWHTWAVED.Item("WAVE_SUB")

            Dim row As DataRow = Nothing
            If WAVE_SUB = "1" AndAlso dst.Tables("WHTWAVE2_SUB").Select("WAVE_NO = '" & WAVE_NO & "' and WAVE_LNO = " & CStr(WAVE_LNO)).Length > 0 Then
                row = dst.Tables("WHTWAVE2_SUB").Rows.Find(New String() {WAVE_NO, WAVE_LNO_LINK, WAVE_LNO})
            Else
                row = dst.Tables("WHTWAVE2").Rows.Find(New String() {WAVE_NO, WAVE_LNO})
                ' rowWHTWAVE2 = dst.Tables("WHTWAVE2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0)
            End If
            If row IsNot Nothing Then
                row.Item("WAVE_QTY") = Val(row.Item("WAVE_QTY") & "") - Val(rowWHTWAVED.Item("WAVE_QTY") & "")
                row.Item("WAVE_QTY_OPEN") = Val(row.Item("WAVE_QTY_OPEN") & "") - Val(rowWHTWAVED.Item("WAVE_QTY") & "")
            End If
        Next
        Display_Totals()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdWHTINST1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdWHTINST1.BeforeRowsDeleted
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Delete Pick Instructions")

        dst.Tables("WHTWAVED").Rows.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Band.Index = 1 Then
                If grow.ParentRow.Cells("WAVE_PICK_TYPE").Value = "L" Then
                    MsgBox("Cannot Delete a Single Case Instruction from a Pallet Pick", MsgBoxStyle.OkOnly, "Cannot Delete")
                    e.Cancel = True
                ElseIf grow.ParentRow.ChildBands(0).Rows.Count = 1 Then
                    MsgBox("Cannot Delete the Only Case Instruction from a Pick - Delete Entire Pick", MsgBoxStyle.OkOnly, "Cannot Delete")
                    e.Cancel = True
                ElseIf grow.ParentRow.Selected Then
                    MsgBox("Cannot Delete a Case Instruction if you are Also Deleting the Entire Pick", MsgBoxStyle.OkOnly, "Cannot Delete")
                    e.Cancel = True
                End If
            Else
                If grow.Cells("WAVE_INST_STATUS").Value & "" <> "0" Then
                    MsgBox("Cannot Delete a Case Instruction which has already been picked", MsgBoxStyle.OkOnly, "Cannot Delete")
                    e.Cancel = True
                End If
            End If
            If Not e.Cancel Then
                If grow.Band.Index = 1 Then
                    Load_Deleted_Wave_Instruction(grow)
                Else
                    grow.Expanded = True
                    For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                        Load_Deleted_Wave_Instruction(grow2)
                    Next
                End If
            End If
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Deleted_Wave_Instruction(grow As UltraWinGrid.UltraGridRow)
        Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value
        Dim WAVE_LNO As Int64 = Val(grow.Cells("WAVE_LNO").Value & "")
        Dim WAVE_SUB As String = grow.ParentRow.Cells("WAVE_SUB").Value

        Dim WAVE_LNO_LINK As Int64 = 0
        If WAVE_SUB = "1" AndAlso dst.Tables("WHTWAVE2_SUB").Select("WAVE_NO = '" & WAVE_NO & "' and WAVE_LNO = " & CStr(WAVE_LNO)).Length > 0 Then
            Dim rowWHTWAVE2_SUB As DataRow = dst.Tables("WHTWAVE2_SUB").Select("WAVE_LNO = " & CStr(WAVE_LNO))(0)
            WAVE_LNO_LINK = rowWHTWAVE2_SUB.Item("WAVE_LNO_LINK")
        End If

        Dim rowWHTWAVED As DataRow = dst.Tables("WHTWAVED").Rows.Find(New String() {STYLE_CODE, COLOR_CODE, WAVE_LNO})
        If rowWHTWAVED Is Nothing Then
            rowWHTWAVED = dst.Tables("WHTWAVED").Rows.Add({STYLE_CODE, COLOR_CODE, WAVE_LNO, 0, WAVE_SUB, WAVE_LNO_LINK})
        End If
        rowWHTWAVED.Item("WAVE_QTY") = Val(rowWHTWAVED.Item("WAVE_QTY") & "") + Val(grow.Cells("LOCATION_QTY_WAVE").Value & "")
    End Sub

    Private Sub grdWHTINST1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTINST1.InitializeRow
        If e.Row.Band.Index = 0 Then
            With e.Row.Cells("WAVE_PICK_TYPE")
                If .Value = "L" Then
                    .Appearance.ForeColor = Drawing.Color.Blue
                ElseIf .Value = "C" Then
                    .Appearance.ForeColor = Drawing.Color.Green
                ElseIf .Value = "U" Then
                    .Appearance.ForeColor = Drawing.Color.DarkOrange
                End If
            End With

            With e.Row.Cells("WAVE_INST_STATUS")
                If .Value = "0" Then ' Waved
                    .Appearance.ForeColor = Drawing.Color.Empty
                ElseIf .Value = "2" Then ' Deposit
                    .Appearance.ForeColor = Drawing.Color.Purple
                ElseIf .Value = "C" Then ' Cancelled
                    .Appearance.ForeColor = Drawing.Color.Orange
                ElseIf .Value = "1" Then ' Picked
                    If Val(e.Row.Cells("UNITS_WAVE").Value & "") = Val(e.Row.Cells("UNITS_PICK").Value & "") Then
                        .Appearance.ForeColor = Drawing.Color.Green
                    Else
                        .Appearance.ForeColor = Drawing.Color.Red
                    End If
                Else ' Void
                    .Appearance.ForeColor = Drawing.Color.Red
                End If
            End With

        End If
    End Sub

    Function Pick_Test() As Boolean
        ASCMAIN1.sql = "Select Count (*) from WHTWAVE2 where WAVE_QTY_PICK <> 0 and WAVE_NO = '" & WAVE_NO & "'"
        Dim wave_was_picked As Boolean = (Val(ASCDATA1.GetDataValue) <> 0)
        If Not wave_was_picked Then
            ASCMAIN1.sql = "Select Count (*) from WHTINST1 where NVL(WAVE_INST_STATUS,'0') <> '0' and WAVE_NO = '" & WAVE_NO & "'"
            wave_was_picked = (Val(ASCDATA1.GetDataValue) <> 0)
        End If
        Return wave_was_picked
    End Function

    Private Sub cmdWave_Click(sender As System.Object, e As System.EventArgs) Handles cmdWave.Click

        Dim sqlsame As String = "STYLE_CODE = STYLE_CODE_SUB and COLOR_CODE = COLOR_CODE_SUB"
        Dim rowsame() As DataRow = dst.Tables("WHTWAVE2_SUB").Select(sqlsame)
        If rowsame.Length <> 0 Then
            MsgBox("Sub Style same as Original Style on Some Demand Records" _
                   & vbCrLf & " (See " & rowsame(0).Item("STYLE_CODE") & "-" & rowsame(0).Item("COLOR_CODE") & ")",
                   MsgBoxStyle.OkOnly, "Cannot Create More Wave Instructions")
            Exit Sub
        End If

        If dst.Tables("WHTWAVE2").Select("WAVE_QTY_LEFT > 0").Length = 0 And
            dst.Tables("WHTWAVE2_SUB").Select("WAVE_QTY_LEFT > 0").Length = 0 Then
            MsgBox("Nothing left to Wave", MsgBoxStyle.OkOnly, "Cannot Create More Wave Instructions")
            Exit Sub
        End If

        If chkForcePalletPick.Checked Then
            If txtPreferredLocation.Text = "" Then
                MsgBox("You may force a Pallet Pick ONLY when selecting a Single Location", MsgBoxStyle.OkOnly, "Cannot Create More Wave Instructions")
                Exit Sub
            End If
        End If

        sqlPPK_CODEs = ""
        If Not chkNoPPK.Checked Then
            Dim PPK_CODEs As New List(Of String)
            For Each row As DataRow In dst.Tables("WHTWAVEP").Select("SEL = '1'")
                Dim PPK_CODE As String = row.Item("PPK_CODE")
                If Not PPK_CODEs.Contains(PPK_CODE) Then
                    PPK_CODEs.Add(PPK_CODE)
                End If
            Next
            If PPK_CODEs.Count > 0 Then
                sqlPPK_CODEs = "'" & Join(PPK_CODEs.ToArray, "','") & "'"
            End If
        End If

        Create_Wave_Instructions(False)

        Absx1.txtFor("LOCATION_CODE").Text = ""

        grdWHTWAVE2.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        ' chkPrePackPallets.Checked = False
    End Sub

    Sub Get_SHIP_BOL_NOs()
        If preview Then
            sqlSHIP_BOL_NOs = ",'" & SHIP_BOL_NO & "'"
            Exit Sub
        End If

        Dim TABLE_NAME As String = ""
        If EntryMode = "N" Then
            TABLE_NAME = "SOTSHIPX"
        Else
            TABLE_NAME = "WHTWAVE3"
        End If
        sqlSHIP_BOL_NOs = ""
        sqlSHIP_BOL_NOs_2 = ""
        Dim ctr As Integer = 0
        For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
            ctr += 1


            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            sqlSHIP_BOL_NOs &= ",'" & SHIP_BOL_NO & "'"
            'If ctr > 1000 Then
            '    sqlSHIP_BOL_NOs_2 &= ",'" & SHIP_BOL_NO & "'"
            'Else
            '    sqlSHIP_BOL_NOs &= ",'" & SHIP_BOL_NO & "'"
            'End If

            'Dim rowSOTBOLNO As DataRow = dst.Tables(TempTable_BOL_NOs).NewRow
            'With rowSOTBOLNO
            '    .Item("SHIP_BOL_NO") = row.Item("SHIP_BOL_NO")
            'End With
            'dst.Tables(TempTable_BOL_NOs).Rows.Add(rowSOTBOLNO)

        Next
        'Update_Record_TDA(TempTable_BOL_NOs)
    End Sub

    Private Sub grdWHTWAVE2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTWAVE2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdWHTWAVE2, sql_where)
                    Dim COLOR_CODE As String = .Cells("COLOR_CODE").Value & ""
                    Dim STYLE_CODE = Select_Style(COLOR_CODE)

                    If STYLE_CODE <> "" Then

                        If Validate_Style(STYLE_CODE) <> "" Then
                            .Cells("STYLE_CODE").Value = STYLE_CODE
                            .Cells("COLOR_CODE").Value = COLOR_CODE
                            .Update()
                        End If

                        ''  Dim z As String = .Cells("STYLE_CODE_SUB").Value
                        'Dim STYLE_CODE As String = Validate_Style(STYLE_CODE_SUB)
                        'If STYLE_CODE = "" Then
                        '    'STYLE_CODE = z
                        '    'Validate_Style(z)
                        'Else
                        '    'If .Cells("STYLE_CODE_SUB").Value = "" Then
                        '    '    .Cells("STYLE_CODE_SUB").Value = .Cells("STYLE_CODE").Value
                        '    'End If
                        '    .Cells("STYLE_CODE_SUB").Value = STYLE_CODE
                        '    .Cells("COLOR_CODE_SUB").Value = COLOR_CODE_SUB
                        '    ' .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        '    .Update()
                        'End If
                    End If

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                    grdClickCellButton(grdWHTWAVE2, sql_where, , , "COLOR_CODE")

                Case "STYLE_CODE_SUB"

                    Dim STYLE_CODE_SUB As String
                    Dim COLOR_CODE_SUB As String
                    If Val(.Cells("WAVE_QTY").Value & "") = 0 And
                       Val(.Cells("WAVE_QTY_PICK").Value & "") = 0 Then

                        COLOR_CODE_SUB = .Cells("COLOR_CODE_SUB").Value & ""
                        STYLE_CODE_SUB = Select_Style(COLOR_CODE_SUB)

                        If STYLE_CODE_SUB <> "" Then

                            If Validate_Style(STYLE_CODE_SUB) <> "" Then
                                .Cells("STYLE_CODE_SUB").Value = STYLE_CODE_SUB
                                .Cells("COLOR_CODE_SUB").Value = COLOR_CODE_SUB
                                .Update()
                            End If

                            ''  Dim z As String = .Cells("STYLE_CODE_SUB").Value
                            'Dim STYLE_CODE As String = Validate_Style(STYLE_CODE_SUB)
                            'If STYLE_CODE = "" Then
                            '    'STYLE_CODE = z
                            '    'Validate_Style(z)
                            'Else
                            '    'If .Cells("STYLE_CODE_SUB").Value = "" Then
                            '    '    .Cells("STYLE_CODE_SUB").Value = .Cells("STYLE_CODE").Value
                            '    'End If
                            '    .Cells("STYLE_CODE_SUB").Value = STYLE_CODE
                            '    .Cells("COLOR_CODE_SUB").Value = COLOR_CODE_SUB
                            '    ' .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                            '    .Update()
                            'End If
                        End If
                    Else
                        ' CANNOT SUB STYLE IF PICKED, SHIPPED OR CANCELLED
                    End If

                    grdWHTWAVE2.UpdateData()

                Case "COLOR_CODE_SUB"
                    Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                    grdClickCellButton(grdWHTWAVE2, sql_where, , , "COLOR_CODE")
            End Select
        End With
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
            STYLE_CODE = Validate_Style(STYLE_CODE)
        End If

        If COLOR_CODE <> "" Then
            If STYLE_CODE <> "" Then
                Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 Is Nothing Then
                    COLOR_CODE = ""
                    'MsgBox("Color Code '" & COLOR_CODE & "' is not Associated with Style " & STYLE_CODE)
                    'STYLE_CODE = ""
                End If
            End If
        End If

        If COLOR_CODE = "" Then
            If COLOR_CODEs.Count = 1 Then
                COLOR_CODE = COLOR_CODEs(0)
            Else
                Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"

                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("COLOR_CODE", , sql_where)
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""

                    'ASCMAIN1.CodeSelector.SQL = "Select * from (" & ASCMAIN1.CodeSelector.SQL & ")" _
                    '    & " where COLOR_CODE in " _
                    '    & " (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "')"

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

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim E As String = ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

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

        If E <> "" Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If E = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

    Private Sub grdWHTWAVE2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdWHTWAVE2.BeforeRowsDeleted
        If e.Rows.Count > 1 Then
            MsgBox("You cannot delete multiple rows at 1 time")
            e.Cancel = True
        Else
            If WAVE_TYPE = "W" Then
                If Val(e.Rows(0).Cells("WAVE_QTY").Value & "") <> 0 Or Val(e.Rows(0).Cells("WAVE_QTY_PICK").Value & "") <> 0 Then
                    MsgBox("Cannot Delete a Style with Wave or Pick Activity", MsgBoxStyle.OkOnly, "Cannot Delete Style")
                    e.Cancel = True
                    Exit Sub
                End If
            Else
                If Val(e.Rows(0).Cells("WAVE_QTY").Value & "") <> 0 Then
                    If MsgBox("Would you like to Zero the Sub Qty?", MsgBoxStyle.YesNo,
                              "Cannot Delete a Sub after a Wave Instruction has been Generated") = MsgBoxResult.Yes Then
                        e.Rows(0).Cells("WAVE_QTY_SUB").Value = 0
                        e.Rows(0).Update()
                    End If
                    e.Cancel = True
                    e.Rows(0).Selected = False
                Else
                    'e.Rows(0).Cells("WAVE_SUB").Value = 0
                    'e.Rows(0).Cells("WAVE_QTY_SUB").Value = 0
                    'e.Rows(0).Update()
                End If
            End If
        End If

    End Sub

    Private Sub grdWHTWAVE2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTWAVE2.BeforeRowUpdate
        If e.Row.Band.Index = 0 Then
            If Not e.Cancel Then
                Validate_Columns("WAVE_QTY_CANC", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
                If e.Cancel Then MsgBox("Invalid Qty to Cancel")
            End If
            If Not e.Cancel Then
                Validate_Columns("WAVE_QTY_CONC", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
                If e.Cancel Then MsgBox("Invalid Qty to Conceal")
            End If
            If Not e.Cancel Then
                Validate_Columns("WAVE_QTY_BACK", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
                If e.Cancel Then MsgBox("Invalid Qty to Conceal")
            End If
            If WAVE_TYPE = "W" Then
                Validate_Columns("STYLE_CODE", e.Cancel)
                If e.Cancel Then MsgBox("Invalid Style")
                If Not e.Cancel Then
                    Validate_Columns("COLOR_CODE", e.Cancel)
                    If e.Cancel Then MsgBox("Invalid Color")
                Else

                End If
            End If
        Else
            Validate_Columns("STYLE_CODE_SUB", e.Cancel)
            If e.Cancel Then MsgBox("Invalid Style")
            If Not e.Cancel Then
                Validate_Columns("COLOR_CODE_SUB", e.Cancel)
                If e.Cancel Then MsgBox("Invalid Color")
            Else

            End If
            If P2L_ALLOW = "Y" Then
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE_SUB").Value
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE_SUB").Value
                Dim isP2L As Int64 = Val(dst.Tables("WHTSCSEQ").Compute("COUNT(STYLE_CODE)", "STYLE_CODE='" & STYLE_CODE & "' and COLOR_CODE='" & COLOR_CODE & "'") & "")
                If isP2L = 0 Then
                    MsgBox("Not a Pick To Light Style")
                    e.Cancel = True
                End If
            End If
            'If Not e.Cancel Then
            '    Validate_Columns("WAVE_QTY_SUB", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
            '    If e.Cancel Then MsgBox("Invalid Qty to Sub")
            'End If
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("WAVE_NO").Value = WAVE_NO
            WAVE_LNO_ctr += 1
            e.Row.Cells("WAVE_LNO").Value = WAVE_LNO_ctr

            Toggle_grdWHTLOCB1()

            Dim WAVE_QTY_LOCS As Int64 = Val(dst.Tables("WHTLOCB1").Compute("SUM(LOCATION_QTY)", "") & "")
            e.Row.Cells("WAVE_QTY_LOCS").Value = WAVE_QTY_LOCS

            If IsWalmart(CUST_CODE) Then
                Dim WAVE_QTY_RACS As Int64 = Val(dst.Tables("WHTLOCB1").Compute("SUM(LOCATION_QTY)", "LOCATION_CODE between '05' and '40'") & "")
                e.Row.Cells("WAVE_QTY_RACS").Value = WAVE_QTY_RACS

                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value
                ASCMAIN1.sql = "Select (WHSE_QTY_ON_HAND - WHSE_QTY_PICK) from ICTSTAT2 WHERE WHSE_CODE = :PARM1 AND STYLE_CODE = :PARM2 AND COLOR_CODE = :PARM3"
                Dim WAVE_QTY_OTS As Int64 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "VVV", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE}) & "")
                e.Row.Cells("WAVE_QTY_OTS").Value = WAVE_QTY_OTS
            End If

        End If

    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)

        With grdWHTWAVE2.ActiveRow

            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    Cancel = (STYLE_CODE = "")

                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If

                Case "STYLE_CODE_SUB"
                    Dim STYLE_CODE_SUB As String = Validate_Style(.Cells("STYLE_CODE_SUB").Value & "")
                    Cancel = (STYLE_CODE_SUB = "")

                Case "COLOR_CODE_SUB"
                    If .Cells("COLOR_CODE_SUB").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE_SUB").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If

                Case "WAVE_QTY_SUB"
                    If Trim(.Cells("STYLE_CODE_SUB").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    'If Trim(.Cells("WAVE_QTY_SUB").Value & "") = "" Then
                    '    MsgBox("Sub Qty Not Specified", vbOKOnly, "Cannot Update Record")
                    '    Cancel = True
                    '    grdWHTWAVE2.ActiveCell = grdWHTWAVE2.ActiveRow.Cells("WAVE_QTY_SUB")
                    '    Exit Sub
                    'End If
                    If Val(.Cells("WAVE_QTY_SUB").Value & "") < 0 Then
                        MsgBox("Sub Qty May Not be Negative", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                    End If

                Case "WAVE_QTY_CANC", "WAVE_QTY_CONC", "WAVE_QTY_BACK"
                    Dim LBL As String = grdWHTWAVE2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
                    'If Trim(.Cells(COLUMN_NAME).Value & "") = "" Then
                    '    MsgBox(LBL & " Qty Not Specified", vbOKOnly, "Cannot Update Record")
                    '    Cancel = True
                    '    grdWHTWAVE2.ActiveCell = grdWHTWAVE2.ActiveRow.Cells(COLUMN_NAME)
                    '    Exit Sub
                    'End If
                    If Val(.Cells(COLUMN_NAME).Value & "") < 0 Then
                        MsgBox(LBL & " Qty May Not be Negative", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                    End If

            End Select
        End With
    End Sub

    Private Sub grdWHTWAVEX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTWAVEX.InitializeRow

        If e.Row.IsDataRow Then

            If e.Row.Cells("SHIP_STATUS").Value & "" = "D" Then
                e.Row.Cells("SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Deleted"
            ElseIf e.Row.Cells("SHIP_STATUS").Value & "" = "F" Then
                e.Row.Cells("SHIP_STATUS").Appearance.BackColor = Drawing.Color.LightGreen
                e.Row.ToolTipText = "Shipped"
            Else
                If e.Row.Cells("SHIP_WAVE_STATUS").Value & "" = "1" Then
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Blue
                    e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Waved"
                Else
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Empty
                    e.Row.Cells("SHIP_BOL_NO").ToolTipText = ""
                End If
            End If

            Dim WAVE_QTY_LEFT As Int64 = 0
            With e.Row.Cells("WAVE_QTY_LEFT")
                WAVE_QTY_LEFT = Val(.Value & "")
                If WAVE_QTY_LEFT > 0 Then
                    .Appearance.ForeColor = Drawing.Color.Red
                End If
            End With

            With e.Row.Cells("WAVE_INST_STATUS_SUMMARY")
                If .Value & "" = "WAVE" Then
                    .ToolTipText = ""
                    e.Row.Appearance.BackColor = Drawing.Color.Empty
                ElseIf .Value & "" = "FINAL" Then
                    .ToolTipText = "Wave is Finalized"
                    .Appearance.ForeColor = Drawing.Color.Blue
                    .Appearance.BackColor = Drawing.Color.Empty
                ElseIf .Value & "" = "VOID" Then
                    .ToolTipText = "Wave is Voided"
                    .Appearance.ForeColor = Drawing.Color.Red
                    .Appearance.BackColor = Drawing.Color.Empty
                ElseIf .Value & "" = "COMP" Then
                    .ToolTipText = "All Wave Picks are Complete"
                    .Appearance.BackColor = Drawing.Color.LightGreen
                Else
                    .ToolTipText = "Some Wave Picks are Complete"
                    .Appearance.BackColor = Drawing.Color.Yellow
                End If

                If WAVE_QTY_LEFT > 0 Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText &= vbCrLf & "Qty To Wave > 0"
                End If
            End With

            'If Val(e.Row.Cells("WAVE_COUNT").Value & "") = Val(e.Row.Cells("OPEN_COUNT").Value & "") And _
            '    Val(e.Row.Cells("PICK_COUNT").Value & "") = 0 Then
            '    e.Row.ToolTipText = ""
            '    e.Row.Appearance.BackColor = Drawing.Color.Empty

            'ElseIf Val(e.Row.Cells("WAVE_COUNT").Value & "") = Val(e.Row.Cells("PICK_COUNT").Value & "") And _
            '    Val(e.Row.Cells("OPEN_COUNT").Value & "") = 0 Then
            '    e.Row.ToolTipText = "All Wave Picks are Complete"
            '    e.Row.Appearance.BackColor = Drawing.Color.LightGreen
            'Else
            '    e.Row.ToolTipText = "Some Wave Picks are Complete"
            '    e.Row.Appearance.BackColor = Drawing.Color.Yellow
            'End If
        End If

    End Sub

    Private Sub grdSOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSHIPX.AfterRowActivate
        If Not ScreenMode Then
            Setup_grdTATEVNT1()
        End If
    End Sub

    Private Sub grdSOTSHIPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIPX.InitializeRow
        If e.Row.IsDataRow Then
            With e.Row.Cells("WAVE_QTY")
                If .Value & "" = "" Then
                    .Appearance.BackColor = Drawing.Color.Empty
                    .ToolTipText = ""
                Else
                    If Val(.Value & "") = Val(e.Row.Cells("PICK_QTY").Value & "") Then
                        .Appearance.BackColor = Drawing.Color.LightGreen
                        .ToolTipText = "All of the Released Qty can be Waved"
                    ElseIf Val(.Value & "") = 0 Then
                        .Appearance.BackColor = Drawing.Color.Orange
                        .ToolTipText = "None of the Released Qty can be Waved"
                    Else
                        .Appearance.BackColor = Drawing.Color.Yellow
                        .ToolTipText = "Some of the Released Qty can be Waved"
                    End If
                End If
            End With


            If e.Row.Cells("ORDR_SHIP_DATE").Value & "" <> "" And e.Row.Cells("ORDR_ORIG_SHIP_DATE").Value & "" <> "" Then
                If Format(e.Row.Cells("ORDR_SHIP_DATE").Value, "yyyyMMdd") <> Format(e.Row.Cells("ORDR_ORIG_SHIP_DATE").Value, "yyyyMMdd") Then
                    e.Row.Cells("ORDR_SHIP_DATE").ToolTipText = "Ship Date has been Changed"
                    e.Row.Cells("ORDR_SHIP_DATE").Appearance.ForeColor = Drawing.Color.Red
                Else
                    e.Row.Cells("ORDR_SHIP_DATE").ToolTipText = ""
                    e.Row.Cells("ORDR_SHIP_DATE").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If

            If e.Row.Cells("ORDR_CANCEL_DATE").Value & "" <> "" And e.Row.Cells("ORDR_ORIG_CANCEL_DATE").Value & "" <> "" Then
                If Format(e.Row.Cells("ORDR_CANCEL_DATE").Value, "yyyyMMdd") <> Format(e.Row.Cells("ORDR_ORIG_CANCEL_DATE").Value, "yyyyMMdd") Then
                    e.Row.Cells("ORDR_CANCEL_DATE").ToolTipText = "Cancel Date has been Changed"
                    e.Row.Cells("ORDR_CANCEL_DATE").Appearance.ForeColor = Drawing.Color.Red
                Else
                    e.Row.Cells("ORDR_CANCEL_DATE").ToolTipText = "Ship Date has been Changed"
                    e.Row.Cells("ORDR_CANCEL_DATE").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If

        End If
    End Sub

    Sub Wave_Preview()

        preview = True

        grdSOTSHIPX.Selected.Rows.Clear()
        Dim tbl As DataTable = dst.Tables("SOTSHIPX").Copy
        For Each row As DataRow In tbl.Select("SELECTED = '1'")
            SHIP_BOL_NO = row.Item("SHIP_BOL_NO")
            WAVE_INST_NO_ctr = 0

            ASCMAIN1.Progress("-", SHIP_BOL_NO)

            For Each GROW As UltraWinGrid.UltraGridRow In grdSOTSHIPX.Rows
                If GROW.Cells("SHIP_BOL_NO").Value = SHIP_BOL_NO Then
                    grdSOTSHIPX.ActiveRow = GROW
                End If
            Next
            Application.DoEvents()

            CUST_CODE = row.Item("CUST_CODE")
            WHSE_CODE = row.Item("WHSE_CODE")
            rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
            For Each row2 As DataRow In dst.Tables("SOTSHIPX").Select("")
                If row2.Item("SHIP_BOL_NO") = SHIP_BOL_NO Then
                    row2.Item("SELECTED") = "1"
                Else
                    row2.Item("SELECTED") = "0"
                End If
            Next

            If dst.Tables("SOTSHIPX").Select("SELECTED='1'").Length = 1 Then
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

                Dim WAVE_QTY As Int64 = Val(dst.Tables("WHTWAVE2").Compute("SUM(WAVE_QTY)", "") & "")

                Dim rowSOTSHIPX_PRE As DataRow = dst.Tables("SOTSHIPX_PRE").Rows.Find(SHIP_BOL_NO)
                If rowSOTSHIPX_PRE Is Nothing Then
                    rowSOTSHIPX_PRE = dst.Tables("SOTSHIPX_PRE").Rows.Add(SHIP_BOL_NO)
                End If
                rowSOTSHIPX_PRE.Item("WAVE_QTY") = WAVE_QTY

                'row.Item("WAVE_QTY") = WAVE_QTY
                dst.Tables("SOTSHIPX").Select("SELECTED='1'")(0).Item("WAVE_QTY") = WAVE_QTY

                Mode_Settings(False)
            End If
        Next

        For Each row2 As DataRow In dst.Tables("SOTSHIPX").Select("")
            row2.Item("SELECTED") = "0"
        Next

        preview = False
        lblPreview.Visible = False

        dst.Tables("ASTSQLX1").Rows.Clear()
        ASCMAIN1.Progress("", "")

        MsgBox("Wave Preview is Complete", MsgBoxStyle.OkOnly, "Verificaiton")

    End Sub

    Private Sub btnFetch_Click(sender As System.Object, e As System.EventArgs) Handles btnFetch.Click
        Load_SOTSHIPX("S")
    End Sub

    Private Sub grdWHTWAVE2_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTWAVE2.InitializeLayout

    End Sub

    Private Sub grdWHTWAVE2_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTWAVE2.BeforeRowActivate

        With grdWHTWAVE2.DisplayLayout.Bands(0)
            If e.Row.IsAddRow Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit

                For Each COLUMN_NAME As String In New String() {"WAVE_QTY_ADJ", "WAVE_QTY_NOTE"}
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Next

            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

                For Each COLUMN_NAME As String In New String() {"WAVE_QTY_CANC", "WAVE_QTY_CONC", "WAVE_QTY_BACK", "WAVE_QTY_ADJ", "WAVE_QTY_NOTE", "WAVE_QTY_BACK"}
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                Next
            End If
        End With

    End Sub

    Private Sub grdWHTWAVE2_BeforeRowFilterChanged(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowFilterChangedEventArgs) Handles grdWHTWAVE2.BeforeRowFilterChanged

    End Sub

    Private Sub grdWHTWAVE2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWHTWAVE2.BeforeExitEditMode
        If grdWHTWAVE2.ActiveCell IsNot Nothing Then
            With grdWHTWAVE2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        'If .EditorResolved.Value & "" <> "" AndAlso .EditorResolved.Value <> CStr(.EditorResolved.Value & "").ToUpper Then
                        If .EditorResolved.IsValid AndAlso .EditorResolved.Value & "" <> "" Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value, .Column.Key)
                        End If
                        'End If
                End Select
            End With
        End If
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If cmbType.Value & "" = "" Then
            MsgBox("No Event Type Selected", MsgBoxStyle.OkOnly, "Cannot Add Event")
            Exit Sub
        End If
        If txtNote.Text = "" Then
            MsgBox("No Event Description Provided", MsgBoxStyle.OkOnly, "Cannot Add Event")
            Exit Sub
        End If

        Record_Event(cmbType.Value, txtNote.Text)


        cmbType.Value = DBNull.Value
        txtNote.Text = ""
    End Sub

    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String)
        Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
        With rowTATEVNT1
            .Item("TABLE_NAME") = "SOTSHIP1"
            .Item("TABLE_KEY") = SHIP_BOL_NO
            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("EVENT_TYPE") = EVENT_TYPE
            .Item("EVENT_DESC") = EVENT_DESC
            .Item("EVENT_KEY") = ""
            .Item("FORM_NAME") = Me.Name
        End With
        dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
        Update_Record_TDA("TATEVNT1")
    End Sub

    Private Sub btnShipmentUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnShipmentUpdate.Click

        Dim SBN As String = grpEditShipment.Tag & ""
        If grdSOTSHIPX.ActiveRow Is Nothing OrElse rowSOTSHIP1 Is Nothing OrElse SBN <> SHIP_BOL_NO Or SBN <> rowSOTSHIP1.Item("SHIP_BOL_NO") Then
            MsgBox("Shipment " & SBN & " is not the Active Shipment any longer", MsgBoxStyle.OkOnly, "Cannot Update")
            chkEditShipment.Checked = False
            Exit Sub
        End If

        Dim rowSOTSHIPX As DataRow = dst.Tables("SOTSHIPX").Rows.Find(SHIP_BOL_NO)

        If Me.BindingContext.Contains(dst.Tables("SOTSHIP1")) Then
            ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
            Dim X As CurrencyManager = Me.BindingContext(dst.Tables("SOTSHIP1"))
            X.EndCurrentEdit()
        End If

        dst.Tables("ASTAUDT1").Rows.Clear()

        For Each COLUMN_NAME As String In New String() {"SHIP_DATE_ROUTED", "SHIP_DATE_PLANNED", "SHIP_DATE_PACKED",
                                                        "SHIP_APPT_NO", "SHIP_NOTES", "SHIP_NOTES_3PL"}
            rowSOTSHIPX.Item(COLUMN_NAME) = rowSOTSHIP1.Item(COLUMN_NAME)
            If rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Current) & "" <> rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                'If rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Original) & "" = "" Then
                '    Record_Event("DATA_CHANGE", grdSOTSHIPX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption & " Changed to " & rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Current))
                '    ' Record_Event(COLUMN_NAME, "Changed to " & rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Current))
                'Else
                '    Record_Event("DATA_CHANGE", grdSOTSHIPX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption & " Changed from " & rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Original) & " to " & rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Current))
                '    ' Record_Event(COLUMN_NAME, "Changed from " & rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Original) & " to " & rowSOTSHIP1.Item(COLUMN_NAME, DataRowVersion.Current))
                'End If

                If rowSOTSHIPX.Item(COLUMN_NAME) & "" <> rowSOTSHIPX.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                    Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                    With rowASTAUDT1
                        .Item("TABLE_NAME") = "SOTSHIP1"
                        .Item("KEY_VALUE") = SHIP_BOL_NO
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("USER_ID") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                        .Item("OLD_VALUE") = rowSOTSHIPX.Item(COLUMN_NAME, DataRowVersion.Original) & ""
                        .Item("NEW_VALUE") = rowSOTSHIPX.Item(COLUMN_NAME) & ""
                        .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                        .Item("SELECTION_NO") = Me.SELECTION_NO
                        .Item("XNO") = Me.XNO
                    End With
                    dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                End If

            End If
        Next

        Update_Record_TDA("ASTAUDT1")
        Update_Record_TDA("SOTSHIP1")

        Fill_Records("SOTSHIPA", SHIP_BOL_NO)
        Sort_grdColumns(grdSOTSHIPA, "INIT_DATE")

        chkEditShipment.Checked = False
    End Sub

    Private Sub btnShipmentCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnShipmentCancel.Click
        chkEditShipment.Checked = False
    End Sub

    Private Sub chkEditShipment_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkEditShipment.CheckedChanged
        If chkEditShipment.Checked Then
            If Not grdSOTSHIPX.Visible Or grdSOTSHIPX.ActiveRow Is Nothing OrElse Not grdSOTSHIPX.ActiveRow.IsDataRow Or grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value & "" <> SHIP_BOL_NO Then
                chkEditShipment.Checked = False
                Exit Sub
            End If
            If ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then
                rowSOTSHIP1 = Fill_Record("SOTSHIP1", SHIP_BOL_NO)

                If rowSOTSHIP1.Item("SHIP_STATUS") & "" <> "P" Then
                    MsgBox("Shipment " & SHIP_BOL_NO & " is No Longer In Pick", MsgBoxStyle.OkOnly, "Cannot Edit Shipment " & SHIP_BOL_NO)
                    chkEditShipment.Checked = False
                    Exit Sub
                End If

                grpEditShipment.Tag = SHIP_BOL_NO
                Toggle_EditShipment()
            Else
                chkEditShipment.Checked = False
            End If
        Else
            Toggle_EditShipment()
        End If
    End Sub

    Sub Toggle_EditShipment()
        btnShipmentCancel.Visible = chkEditShipment.Checked
        btnShipmentUpdate.Visible = chkEditShipment.Checked
        Set_Read_Only(grpEditShipment, Not chkEditShipment.Checked)
        Set_Read_Only_for_ctl(chkEditShipment, False)
        If Not chkEditShipment.Checked Then
            dst.Tables("SOTSHIP1").RejectChanges()
            If Me.BindingContext.Contains(dst.Tables("SOTSHIP1")) Then
                ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
                Dim X As CurrencyManager = Me.BindingContext(dst.Tables("SOTSHIP1"))
                X.CancelCurrentEdit()
            End If
            ASCMAIN1.MultiTask_Release()
            grpEditShipment.Tag = ""
        End If
    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_for_Edit()
    End Sub

    Sub Setup_for_Edit()
        If tabMain.SelectedTab.Key = "Shipments" Then
            Setup_grdTATEVNT1()
            UltraExplorerBar1.Groups("Edit Work Order").Visible = False
        Else
            UltraExplorerBar1.Groups("Edit Shipment").Visible = False
            Setup_Edit_Wave()
        End If
    End Sub

    Private Sub grdWHTLOCB1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTLOCB1.InitializeRow


        'If e.Row.IsDataRow Then
        '    If e.Row.Cells("LOCATION_LOCKED").Value & "" = "1" Then
        '        e.Row.Cells("LOCATION_CODE").Appearance.ForeColor = Drawing.Color.Magenta
        '        e.Row.Cells("LOCATION_CODE").ToolTipText = "Location Locked"
        '    Else
        '        e.Row.Cells("LOCATION_CODE").Appearance.ForeColor = Drawing.Color.Empty
        '        e.Row.Cells("LOCATION_CODE").ToolTipText = ""
        '    End If
        'End If


        If e.Row.IsDataRow Then
            If e.Row.Cells("LOCATION_LOCKED").Value & "" = "1" Then
                e.Row.Cells("LOCATION_CODE").Appearance = grdWHTLOCB1_app1

            Else
                '  e.Row.Cells("LOCATION_CODE").Appearance = grdWHTLOCB1_app0

            End If
        End If

    End Sub

    Sub Manage_Expressions(action As String)

        If action = "Remove" Then

            ' Remove Expressions
            WHTINST1_expressions.Clear()
            For Each dcol As DataColumn In dst.Tables("WHTINST1").Columns
                If dcol.Expression <> "" Then
                    WHTINST1_expressions.Add(dcol.ColumnName, dcol.Expression)
                    dcol.Expression = ""
                End If
            Next

            WHTINST2_expressions.Clear()
            For Each dcol As DataColumn In dst.Tables("WHTINST2").Columns
                If dcol.Expression <> "" Then
                    WHTINST2_expressions.Add(dcol.ColumnName, dcol.Expression)
                    dcol.Expression = ""
                End If
            Next

            WHTWAVE2_expressions.Clear()
            For Each dcol As DataColumn In dst.Tables("WHTWAVE2").Columns
                If dcol.Expression <> "" Then
                    WHTWAVE2_expressions.Add(dcol.ColumnName, dcol.Expression)
                    dcol.Expression = ""
                End If
            Next
        Else

            ' Restore Expressions

            For Each COLUMN_NAME As String In WHTINST1_expressions.Keys
                dst.Tables("WHTINST1").Columns(COLUMN_NAME).Expression = WHTINST1_expressions(COLUMN_NAME)
            Next

            For Each COLUMN_NAME As String In WHTINST2_expressions.Keys
                dst.Tables("WHTINST2").Columns(COLUMN_NAME).Expression = WHTINST2_expressions(COLUMN_NAME)
            Next

            For Each COLUMN_NAME As String In WHTWAVE2_expressions.Keys
                dst.Tables("WHTWAVE2").Columns(COLUMN_NAME).Expression = WHTWAVE2_expressions(COLUMN_NAME)
            Next
        End If

    End Sub

    Sub Perform_Deposit()

        Dim G As New GunEnvironment
        G.WHSE_CODE = rowWHTWAVE1.Item("WHSE_CODE")
        G.USER_ID = ASCMAIN1.USER_ID

        Dim LOCATION_CODE_DEPOSIT As String = rowWHTWAVE1.Item("LOCATION_CODE_DEPOSIT")
        Dim LOAD_NO_DEPOSIT As String = rowWHTWAVE1.Item("LOAD_NO_DEPOSIT")

        For Each TABLE_NAME As String In New String() {"WHTBARC1", "WHTINST1", "WHTINST2", "WHTMOVE1", "WHTMOVE2"}
            If Not dst.Tables.Contains(TABLE_NAME) Then Create_TDA(dst.Tables.Add, TABLE_NAME, "*")
        Next

        BeginTrans()

        ASCMAIN1.sql = "Select Distinct LOCATION_CODE_OTHER from WHTINST1" & vbCrLf _
            & " where WAVE_INST_STATUS = '1'" & vbCrLf _
            & " and WAVE_NO = '" & WAVE_NO & "'"

        For Each rowGUN_LOC As DataRow In ASCDATA1.GetDataTable.Select("")
            G.GUN_LOC = rowGUN_LOC.Item("LOCATION_CODE_OTHER")

            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"WHTBARC1", "WHTINST1", "WHTINST2", "WHTMOVE1", "WHTMOVE2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)

            ' this is a clone of the routine found in WHCRF005

            dst.Tables("WHTBARC1").Rows.Clear()

            Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

            Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
            With rowWHTMOVE1
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("WHSE_TRAN_TYPE") = "D"
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("WHSE_CODE") = G.WHSE_CODE
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = G.USER_ID
                .Item("STATUS") = "U"
            End With
            dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
            Update_Record_TDA("WHTMOVE1")

            Dim WHSE_TRAN_LNO_ctr As Integer = 0

            Manage_Expressions("Remove")
            ASCMAIN1.sql = "SELECT * FROM WHTINST1" & vbCrLf _
                & " where WAVE_INST_STATUS = '1'" _
                & " and WAVE_NO = '" & WAVE_NO & "'" _
                & " and LOCATION_CODE_OTHER = '" & G.GUN_LOC & "'"
            Fill_Records("WHTINST1", "", True, ASCMAIN1.sql)
            Manage_Expressions("Restore")

            For Each rowWHTINST1 As DataRow In dst.Tables("WHTINST1").Select("", "WAVE_INST_NO, WAVE_NO")

                Dim WAVE_INST_NO As String = rowWHTINST1.Item("WAVE_INST_NO") & ""
                Dim LOAD_NO As String = rowWHTINST1.Item("LOAD_NO") & ""
                Dim LOCATION_CODE As String = rowWHTINST1.Item("LOCATION_CODE") & ""

                Manage_Expressions("Remove")
                ASCMAIN1.sql = "SELECT * FROM WHTINST2 " & vbCrLf _
                        & " where WAVE_INST_NO = '" & WAVE_INST_NO & "'" _
                        & " and LOCATION_QTY_PICK > 0"
                Fill_Records("WHTINST2", "", True, ASCMAIN1.sql)
                Manage_Expressions("Restore")

                For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("", "BAR_CODE, STYLE_CODE, COLOR_CODE")

                    ' If rowWHTINST1.Item("BAR_CODE_OTHER") & "" <> "" Then Stop
                    Dim BAR_CODE As String = IIf(rowWHTINST1.Item("BAR_CODE_OTHER") & "" <> "", rowWHTINST1.Item("BAR_CODE_OTHER"), rowWHTINST2.Item("BAR_CODE") & "")
                    Dim rowWHTBARC1 As DataRow = Fill_Record("WHTBARC1", BAR_CODE, , False)

                    Dim STYLE_CODE As String = rowWHTINST2.Item("STYLE_CODE") & ""
                    Dim COLOR_CODE As String = rowWHTINST2.Item("COLOR_CODE") & ""
                    Dim LOCATION_QTY_PICK As Integer = Val(rowWHTINST2.Item("LOCATION_QTY_PICK") & "")

                    Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
                    With rowWHTMOVE2
                        .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                        WHSE_TRAN_LNO_ctr += 1
                        .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                        .Item("LOCATION_CODE_FROM") = G.GUN_LOC
                        .Item("LOCATION_CODE_TO") = LOCATION_CODE_DEPOSIT
                        .Item("BAR_CODE") = BAR_CODE
                        .Item("WHSE_TRAN_QTY") = LOCATION_QTY_PICK
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE
                        .Item("INIT_OPER") = G.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("STATUS") = "U"
                        .Item("LOAD_NO_FROM") = LOAD_NO
                        .Item("LOAD_NO_TO") = LOAD_NO_DEPOSIT
                    End With
                    dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
                    'If Not IsNothing(rowWHTBARC1) Then
                    rowWHTBARC1.Item("LOAD_NO") = LOAD_NO_DEPOSIT
                    '    rowWHTBARC1.Item("LOCATION_CODE") = LOCATION_CODE_DEPOSIT
                    'End If
                Next

                rowWHTINST1.Item("WAVE_INST_STATUS") = "2"
            Next

            Update_Record_TDA("WHTBARC1")
            Update_Record_TDA("WHTINST1")
            Update_Record_TDA("WHTMOVE2")

            ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                           New Object() {WHSE_TRAN_NO, 0, 1},
                           New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})
        Next


        'temp fix to correct offsetting LPN's in stage area
        ASCMAIN1.sql = " BEGIN " & vbCrLf _
        & " DECLARE " & vbCrLf _
        & " CURSOR C2 IS " & vbCrLf _
        & " Select X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE," & vbCrLf _
        & " X.BAR_CODE as POS_ID,  Y.BAR_CODE as NEG_ID," & vbCrLf _
        & " X.LOCATION_QTY POS_QTY, Y.LOCATION_QTY as NEG_QTY, " & vbCrLf _
        & " X.LOAD_NO as POS_LOAD, Y.LOAD_NO as NEG_LOAD" & vbCrLf _
        & " from " & vbCrLf _
        & " (Select LOCATION_CODE, WHTLOCB1.BAR_CODE, STYLE_CODE, " & vbCrLf _
        & " COLOR_CODE, LOCATION_QTY, LOAD_NO " & vbCrLf _
        & " From WHTLOCB1, WHTBARC1 " & vbCrLf _
        & " Where LOCATION_CODE Like '00-STG-%'" & vbCrLf _
        & " and location_qty > 0" & vbCrLf _
        & " And WHTLOCB1.BAR_CODE = WHTBARC1.BAR_CODE) X," & vbCrLf _
        & " (Select LOCATION_CODE, WHTLOCB1.BAR_CODE, STYLE_CODE, " & vbCrLf _
        & " COLOR_CODE, LOCATION_QTY , LOAD_NO" & vbCrLf _
        & " From WHTLOCB1, WHTBARC1 " & vbCrLf _
        & " Where LOCATION_CODE Like '00-STG-%'" & vbCrLf _
        & " and location_qty < 0" & vbCrLf _
        & " And WHTLOCB1.BAR_CODE = WHTBARC1.BAR_CODE) Y" & vbCrLf _
        & " Where X.LOCATION_CODE = Y.LOCATION_CODE" & vbCrLf _
        & " And X.STYLE_CODE = Y.STYLE_CODE" & vbCrLf _
        & " And X.COLOR_CODE = Y.COLOR_CODE" & vbCrLf _
        & " And X.LOCATION_QTY = ABS(Y.LOCATION_QTY); " & vbCrLf _
        & " BEGIN " & vbCrLf _
        & " FOR R2 IN C2 LOOP " & vbCrLf _
        & " INSERT INTO WHTLOCB1_CATCH SELECT SYSDATE, WHTLOCB1.* FROM WHTLOCB1 where BAR_CODE = R2.POS_ID and LOCATION_CODE = R2.LOCATION_CODE;" & vbCrLf _
        & " INSERT INTO WHTLOCB1_CATCH SELECT SYSDATE, WHTLOCB1.* FROM WHTLOCB1 where BAR_CODE = R2.NEG_ID and LOCATION_CODE = R2.LOCATION_CODE;" & vbCrLf _
        & " UPDATE WHTLOCB1 SET LOCATION_QTY = 0  WHERE BAR_CODE = R2.POS_ID and LOCATION_CODE = R2.LOCATION_CODE; " & vbCrLf _
        & " UPDATE WHTLOCB1 SET LOCATION_QTY = 0  WHERE BAR_CODE = R2.NEG_ID and LOCATION_CODE = R2.LOCATION_CODE;" & vbCrLf _
        & " END LOOP; " & vbCrLf _
        & " END; END; "
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        'Temporarily disable the update to see what happens, the update can be run manually if needed.

        CommitTrans("Deposit Complete")

    End Sub

    Private Sub grdICTIADJ1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTIADJ1.InitializeRow
        If e.Row.Band.Index = 0 Then
            If e.Row.Cells("REVERSED_BY_ADJ_NO").Value & "" <> "" Or e.Row.Cells("REVERSES_ADJ_NO").Value & "" <> "" Then
                e.Row.Appearance.ForeColor = Drawing.Color.Red
                e.Row.ToolTipText = "Reversed"
            End If
        End If
    End Sub

    Private Sub grdWHTINSTS_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTINSTS.InitializeRow
        With e.Row.Cells("WAVE_PICK_TYPE")
            If .Value & "" = "L" Then
                .Appearance.ForeColor = Drawing.Color.Blue
            ElseIf .Value & "" = "C" Then
                .Appearance.ForeColor = Drawing.Color.Green
            ElseIf .Value & "" = "U" Then
                .Appearance.ForeColor = Drawing.Color.DarkOrange
            End If
        End With

        With e.Row.Cells("WAVE_INST_STATUS")
            If .Value & "" = "0" Then
                .Appearance.ForeColor = Drawing.Color.Empty
            ElseIf .Value & "" = "1" Then
                If Val(e.Row.Cells("LOCATION_QTY_WAVE").Value & "") = Val(e.Row.Cells("LOCATION_QTY_PICK").Value & "") Then
                    .Appearance.ForeColor = Drawing.Color.Green
                Else
                    .Appearance.ForeColor = Drawing.Color.Red
                End If
            Else
                .Appearance.ForeColor = Drawing.Color.Red
            End If
        End With
    End Sub

    Private Sub optFilter_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optFilter.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Load_SOTSHIPX()
    End Sub

    Private Sub grdWHTINST1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTINST1.InitializeLayout

    End Sub

    Private Sub grdSOTSHIPX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTSHIPX.InitializeLayout

    End Sub


    Private Sub cmdUpdate_Edit_Wave_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate_Edit_Wave.Click

        Dim Wave_No_Edit As String = grdWHTWAVEX.ActiveRow.Cells("WAVE_NO").Value
        Dim WHTWAVE1_Edit As DataRow = Fill_Record("WHTWAVE1", Wave_No_Edit)
        dst.Tables("ASTAUDT1").Rows.Clear()

        For Each COLUMN_NAME As String In New String() _
                {"SHIP_APPT_NO_WAVE", "SHIP_NOTES_WAVE", "SHIP_NOTES_3PL_WAVE"}
            Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
            With rowASTAUDT1
                .Item("TABLE_NAME") = "WHTWAVE1"
                .Item("KEY_VALUE") = Wave_No_Edit
                .Item("COLUMN_NAME") = Replace(COLUMN_NAME, "_WAVE", "")
                .Item("USER_ID") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("OLD_VALUE") = WHTWAVE1_Edit.Item(Replace(COLUMN_NAME, "_WAVE", "")) & ""
                .Item("NEW_VALUE") = Absx1.txtFor(COLUMN_NAME).Text & ""
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("SELECTION_NO") = Me.SELECTION_NO
                .Item("XNO") = Me.XNO
            End With
            dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
        Next
        For Each COLUMN_NAME As String In New String() _
        {"SHIP_DATE_ROUTED_WAVE", "SHIP_DATE_PLANNED_WAVE", "SHIP_DATE_PACKED_WAVE"}
            Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
            With rowASTAUDT1
                .Item("TABLE_NAME") = "WHTWAVE1"
                .Item("KEY_VALUE") = Wave_No_Edit
                .Item("COLUMN_NAME") = Replace(COLUMN_NAME, "_WAVE", "")
                .Item("USER_ID") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("OLD_VALUE") = WHTWAVE1_Edit.Item(Replace(COLUMN_NAME, "_WAVE", "")) & ""
                .Item("NEW_VALUE") = Absx1.dteFor(COLUMN_NAME).Value & ""
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("SELECTION_NO") = Me.SELECTION_NO
                .Item("XNO") = Me.XNO
            End With
            dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
        Next
        Update_Record_TDA("ASTAUDT1")


        ASCMAIN1.sql = "Update WHTWAVE1 Set SHIP_NOTES_3PL = '" & Absx1.txtFor("SHIP_NOTES_3PL_WAVE").Text & "" & "'," & vbCrLf _
            & " SHIP_NOTES = '" & Absx1.txtFor("SHIP_NOTES_WAVE").Text & "" & "'," & vbCrLf _
            & " SHIP_APPT_NO = '" & Absx1.txtFor("SHIP_APPT_NO_WAVE").Text & "" & "'," & vbCrLf _
            & " SHIP_DATE_PACKED  = '" & IIf(Absx1.dteFor("SHIP_DATE_PACKED_WAVE").Value & "" = "", "", Format(Absx1.dteFor("SHIP_DATE_PACKED_WAVE").Value, "dd-MMM-yy")) & "'," & vbCrLf _
            & " SHIP_DATE_PLANNED = '" & IIf(Absx1.dteFor("SHIP_DATE_PLANNED_WAVE").Value & "" = "", "", Format(Absx1.dteFor("SHIP_DATE_PLANNED_WAVE").Value, "dd-MMM-yy")) & "'," & vbCrLf _
            & " SHIP_DATE_ROUTED  = '" & IIf(Absx1.dteFor("SHIP_DATE_ROUTED_WAVE").Value & "" = "", "", Format(Absx1.dteFor("SHIP_DATE_ROUTED_WAVE").Value, "dd-MMM-yy")) & "'" & vbCrLf _
            & " Where WAVE_NO = '" & Wave_No_Edit & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        MsgBox("Update Complete", MsgBoxStyle.OkOnly, "Success")
        Clear_Record()
        Setup_for_Edit()

    End Sub

    Private Sub grdWHTWAVEX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTWAVEX.InitializeLayout

    End Sub

    Public Overrides Function CustomSummary_End(
    ByVal summarySettings As UltraWinGrid.SummarySettings,
    ByVal rows As UltraWinGrid.RowsCollection,
    ByVal CustomValue As Double,
    ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdWHTWAVE2"
                Dim KEY As String = summarySettings.Key
                If KEY = "WAVE_QTY" Then
                    TOTALS.Add("WAVE_QTY", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("WAVE_QTY") <> 0 Then CustomValue = TOTALS("WAVE_QTY")
                ElseIf KEY = "WAVE_QTY_PICK" Then
                    TOTALS.Add("WAVE_QTY_PICK", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("WAVE_QTY_PICK") <> 0 Then CustomValue = TOTALS("WAVE_QTY_PICK")
                ElseIf KEY = "WAVE_QTY_OPEN" Then
                    TOTALS.Add("WAVE_QTY_OPEN", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("WAVE_QTY_OPEN") <> 0 Then CustomValue = TOTALS("WAVE_QTY_OPEN")
                End If

                'Case "grdSOTINVHX"
                '    Dim KEY As String = summarySettings.Key
                '    If KEY = "GPP" Then
                '        TOTALS.Add("ORDR_AMT_SHIP", 0)
                '        TOTALS.Add("GPA", 0)
                '        CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                '        If TOTALS("ORDR_AMT_SHIP") <> 0 Then CustomValue = 100 * TOTALS("GPA") / TOTALS("ORDR_AMT_SHIP")
                '    Else
                '        Stop
                '    End If
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
            Case "grdWHTWAVE2"
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
                If KEY = "WAVE_QTY" Then
                    TOTALS("WAVE_QTY") += Val(grow2.Cells("SUB_WAVE").Value & "") + Val(grow2.Cells("WAVE_QTY").Value & "")
                ElseIf KEY = "WAVE_QTY_PICK" Then
                    TOTALS("WAVE_QTY_PICK") += Val(grow2.Cells("SUB_PICK").Value & "") + Val(grow2.Cells("WAVE_QTY_PICK").Value & "")
                ElseIf KEY = "WAVE_QTY_OPEN" Then
                    TOTALS("WAVE_QTY_OPEN") += Val(grow2.Cells("SUB_OPEN").Value & "") + Val(grow2.Cells("WAVE_QTY_OPEN").Value & "")
                ElseIf KEY = "TRADE_CLASS_CODE" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub

    Private Sub cmdGetPrePacks_Click(sender As Object, e As EventArgs) Handles cmdGetPrePacks.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Pre-Packs")

        Setup_SOTSHIPC()

        Fill_Records("WHTWAVEP", WHSE_CODE)
        Sort_grdColumns(grdWHTWAVEP, "PPK_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Setup_SOTSHIPC()
        ASCDATA1.ExecuteSQL("Delete from " & SOTSHIPC)
        dst.Tables("SOTSHIPC").Rows.Clear()
        For Each TABLE_NAME As String In New String() {"WHTWAVE2", "WHTWAVE2_SUB"}
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("WAVE_QTY_LEFT > 0")
                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")

                If TABLE_NAME = "WHTWAVE2_SUB" Then
                    STYLE_CODE = row.Item("STYLE_CODE_SUB")
                    COLOR_CODE = row.Item("COLOR_CODE_SUB")
                End If
                If dst.Tables("SOTSHIPC").Rows.Find(New String() {STYLE_CODE, COLOR_CODE}) Is Nothing Then
                    dst.Tables("SOTSHIPC").Rows.Add(New String() {STYLE_CODE, COLOR_CODE})
                End If
            Next
        Next
        Update_Record_TDA("SOTSHIPC")
    End Sub

    Private Sub grdWHTWAVEP_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTWAVEP.AfterRowActivate
        If grdWHTWAVEP.ActiveRow Is Nothing Then
            grdWHTPPKM2.Visible = False
        Else
            Dim PPK_CODE As String = grdWHTWAVEP.ActiveRow.Cells("PPK_CODE").Value & ""
            Fill_Records("WHTPPKM2", PPK_CODE)
            grdWHTPPKM2.Text = "Prepack " & PPK_CODE
            grdWHTPPKM2.Visible = True
            Sort_grdColumns(grdWHTPPKM2, "STYLE_CODE, COLOR_CODE")
        End If
    End Sub

    Private Function IsWalmart(ByRef CUST_CODE As String) As Boolean
        If String.IsNullOrEmpty(CUST_CODE) Then
            Return False
        End If
        Return WalmartCodes.Contains(CUST_CODE)
    End Function

    Private Sub btnP2L_Click(sender As Object, e As EventArgs) Handles btnP2L.Click
        'This is a general test area, P2L was tested succesfully
        ' now using this sub to experiment with RFID decoding the UPC code
        Dim RFID As String = "30340BDFC80A725AE4C8E787"
        'Convert.ToString(Convert.ToInt64(hexstring, 16), 2)

        Dim Binarystring = Convert.ToString(Convert.ToInt64(RFID.Substring(0, 12), 16), 2).PadLeft(48, "0") & Convert.ToString(Convert.ToInt64(RFID.Substring(12), 16), 2).PadLeft(48, "0")
        Debug.Print(Binarystring)
        Dim Prefix As String = Convert.ToInt32(Binarystring.substring(14, 24), 2).ToString.PadLeft(7, "0")
        Dim ControlNo As String = Convert.ToInt32(Binarystring.substring(38, 20), 2).ToString.PadLeft(6, "0")
        Dim upc As String = TAC.SOCMAIN1.UPC(Me, ControlNo.Substring(1, 5), Prefix.Substring(1, 6))

        ' no need to continue P2L test
        Stop
        Exit Sub

        'select the following range of ordr_group_no's 303118 - 303159
        Dim CARTON_NO As String
        Dim PickOrderSql As String


        '        StringBuilder sb = New StringBuilder();
        'sb.Append("Hello ");
        'sb.AppendLine("World!");
        'sb.AppendLine("Hello C#");

        ASCMAIN1.sql = "select * from sotordr0 " & vbCrLf _
                    & " Where sotordr0.ordr_group_no between '0000303118' and '0000303159'"
        Dim tblOrderGroup As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        PickOrderSql = "select * from sotpick1, sotordr1, sotcart1 " & vbCrLf _
                    & " where sotcart1.pick_no = sotpick1.pick_no" & vbCrLf _
                    & " and sotpick1.ordr_no = sotordr1.ordr_no" & vbCrLf _
                    & " and sotordr1.ordr_group_no between '0000303118' and '0000303159'"
        Dim tblPickOrder As DataTable = ASCDATA1.GetDataTable(PickOrderSql)

        ASCMAIN1.sql = "select CART_NO, WHTLOCM1.LOCATION_CODE, SOTCART2.QTY_PACKED, SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE from sotcart2, WHTSCSEQ, WHTLOCM1" & vbCrLf _
                    & " where sotcart2.style_code = whtscseq.style_code" & vbCrLf _
                    & " and sotcart2.color_code = whtscseq.color_code" & vbCrLf _
                    & " and whtscseq.style_seq = whtlocm1.location_route_seq" & vbCrLf _
                    & " and whtscseq.cust_code = 'WALMART'" & vbCrLf _
                    & " and whtlocm1.whse_code = 'NJC'" & vbCrLf _
                    & " and cart_no in (" & PickOrderSql.Replace("*", "cart_no") & ")"
        Dim tblPickLine As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)


        For Each rowGroup As DataRow In tblOrderGroup.Select("", "ORDR_GROUP_NO")
            Dim Group As String = rowGroup("ORDR_GROUP_NO")
            Dim xmlString As New Text.StringBuilder
            xmlString.AppendLine("<LPXML>")
            For Each rowCarton As DataRow In tblPickOrder.Select("ORDR_GROUP_NO = '" & Group & "'", "CART_NO")
                CARTON_NO = rowCarton("CART_NO")
                Dim carton As String = "<PickOrder PickOrderNumber='" & CARTON_NO & "'>"
                xmlString.AppendLine(carton)
                'Extra info if needed by P2L
                '<PickOrderXtra ShipMethod="UPS"/>
                xmlString.AppendLine(String.Format("<PickOrderXtra CUST_PO='{0}' CUST_DC='{1}' CUST_STORE='{2}'/>", rowCarton("ORDR_CUST_PO"), rowCarton("CUST_DC_NO"), rowCarton("CUST_STORE_NO")))
                'this is the detail section
                For Each rowItem As DataRow In tblPickLine.Select("CART_NO = '" & CARTON_NO & "'", "LOCATION_CODE")
                    xmlString.AppendLine(String.Format("<PickLine LocationName='{0}' PickOrderQty='{1}'>", rowItem("LOCATION_CODE"), rowItem("QTY_PACKED")))
                    xmlString.AppendLine(String.Format("<PickLineXtra STYLE_CODE='{0}' COLOR_CODE='{1}'/>", rowItem("STYLE_CODE"), rowItem("COLOR_CODE")))
                    xmlString.AppendLine(String.Format("</PickLine>"))
                Next
                'done with the detail
                xmlString.AppendLine("</PickOrder>")
            Next
            xmlString.AppendLine("</LPXML>")
            Dim fileName = "P2L_" & Group & ".xml"
            Dim doc = New XmlDocument()
            doc.LoadXml(xmlString.ToString)
            doc.Save(fileName)
        Next


        'INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES(xmlString.ToString)

    End Sub

    Private Sub grdWHTWAVE2_CellChange(sender As Object, e As CellEventArgs) Handles grdWHTWAVE2.CellChange

    End Sub
End Class

Public Class GunEnvironment
    Public DBS_SERVER As String
    Public DBS_COMPANY As String
    Public DBS_PASSWORD As String

    Public THREAD_NO As Integer
    Public APP_ID As String
    Public APP_DESC As String
    Public USER_ID As String
    Public GUN_LOC As String
    Public PICK_TYPE As String
    Public WHSE_CODE As String
End Class

Public Class srtComparerWHTWAVES
    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        Dim xCell As UltraWinGrid.UltraGridCell = DirectCast(x, UltraWinGrid.UltraGridCell)
        Dim yCell As UltraWinGrid.UltraGridCell = DirectCast(y, UltraWinGrid.UltraGridCell)

        Dim xv As String = xCell.Value & ""
        Dim yv As String = yCell.Value & ""

        Dim COLUMN_NAME As String = xCell.Column.Key

        If xv = yv Then
            Return 0
        Else
            If COLUMN_NAME = "WAVE_PICK_TYPE" Then
                If xv = "L" Then
                    Return -1
                ElseIf yv = "L" Then
                    Return 1
                Else
                    Return IIf(xv < yv, -1, 1)
                End If
            End If
        End If

    End Function

End Class