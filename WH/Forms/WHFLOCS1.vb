Public Class WHFLOCS1

    ' GCV - PRINT 
    ' GCV - CONTROLS ON MOVE TO (SHOULD ALREADY BE DONE - NEED TO TEST); PRESERVE LOAD / LOCATION INTEGRITY, MUST MOVE BARCODES, MUST MAKE CHANGES WHICH DO NOT BREAK NYA
    Dim sqlWHTLOCB2where As String
    Dim rowICTWHSE1 As DataRow
    Dim WHSE_CODE As String
    Dim LOCATION_CODE As String
    Dim STYLE_CODE As String
    Dim sqlWHTINSTX As String
    Dim SOURCE_GRID As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst

            ASCMAIN1.sql = " Select  WHTLOCM1.*, X.CYCLE_NO, CYCLE_STATUS, CYCLE_RESOLUTION, INIT_OPER, LAST_OPER, " _
            & " INIT_DATE, LAST_DATE as LAST_CYCLE_DATE, CYCLE_TYPE, CASES_BOOK, CASES_PHYS " _
            & " from WHTLOCM1," _
            & " (Select * from WHTCYCL1 Where (LOCATION_CODE, CYCLE_NO) in (Select LOCATION_CODE, Max(CYCLE_NO) CYCLE_NO " _
            & "  from WHTCYCL1 group by LOCATION_CODE)) X " _
            & " Where WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE(+)"

           
            Create_TDA(.Tables.Add, "WHTLOCMM", "**", 0, False, "", 2)
            With .Tables("WHTLOCMM")
                .Columns.Add("CASES", GetType(System.Int64))
                .Columns.Add("BAR_CODE")
                .Columns.Add("QTY_LOCS", GetType(System.Int64))
                .Columns.Add("QTY_WAVE", GetType(System.Int64))
                .Columns.Add("LAST_DATE", GetType(System.DateTime))
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
            End With

            With .Tables.Add("WHTWMAPA")
                .Columns.Add("AISLE")
                For S As Integer = 1 To 99
                    For L As Integer = 1 To 4
                        .Columns.Add(Format(S, "000") & Mid("ABCD", L, 1))
                    Next
                Next
                .PrimaryKey = New DataColumn() {.Columns("AISLE")}
            End With

          

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")

            ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE, LOCATION_QTY INST, LOCATION_QTY LOCB, LOCATION_QTY DIFF from WHTLOCB1"
            Create_TDA(.Tables.Add, "WHTLOCBR", "**", 0, False, "", 5)


            ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE from WHTLOCM1"
            Create_TDA(.Tables.Add, "WHTLOCM1", "**", 0, False, "", 2)

            'ASCMAIN1.sql = "Select X.WHSE_CODE, X.LOCATION_CODE, X.TRANS, X.INIT_DATE" _
            '    & " from WHTLOCM1,(" _
            '    & "Select WHSE_CODE, LOCATION_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
            '    & " from WHTLOCB2 where INIT_DATE > SYSDATE - 24 group by WHSE_CODE, LOCATION_CODE" _
            '    & ") X where WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE" _
            '    & " And WHTLOCM1.WHSE_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "WHTLOCMA", "**", 0, False, "V", 2)
            ASCMAIN1.sql = "Select X.WHSE_CODE, X.LOCATION_CODE, X.TRANS, X.INIT_DATE" _
                & " from WHTLOCM1,(" _
                & "Select WHSE_CODE, LOCATION_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
                & " from WHTLOCB2 where INIT_DATE > SYSDATE - 24 and rownum <1 group by WHSE_CODE, LOCATION_CODE" _
                & ") X where WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCMA", "**", 0, False, "", 2)
            .Tables("WHTLOCMA").Columns("TRANS").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC from ICTSTYL1"
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select X.STYLE_CODE, ICTSTYL1.STYLE_DESC, X.TRANS, X.INIT_DATE" _
                & " from ICTSTYL1,(" _
                & "Select WHSE_CODE, STYLE_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
                & " from WHTLOCB2 where INIT_DATE > SYSDATE - 24 group by WHSE_CODE, STYLE_CODE" _
                & ") X where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" _
                & " And X.WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYLA", "**", 0, False, "V", 1)
            .Tables("ICTSTYLA").Columns("TRANS").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select WHTLOCB1.*,WHTBARC1.LOAD_NO,WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
                & " from WHTLOCB1,WHTBARC1,WHTLOCM1" & vbCrLf _
                & " where WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
                & "   and WHTLOCM1.LOCATION_CODE (+) = WHTLOCB1.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "", 5)
            .Tables("WHTLOCB1").Columns.Add("PERSIST")
            .Tables("WHTLOCB1").Columns.Add("LOCATION_QTY_AVAIL", GetType(System.Int64), "ISNULL(LOCATION_QTY,0)-ISNULL(LOCATION_QTY_WAVE,0)")

            ASCMAIN1.sql = "Select WHTLOCB2.* from WHTLOCB2"
            Create_TDA(.Tables.Add, "WHTLOCB2", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1"
            Create_TDA(.Tables.Add, "WHTCONTC", "**", 0, False, "", 5)

            ASCMAIN1.sql = "Select Distinct BAR_CODE,'0' SEL from WHTLOCB1 where rownum <1"
            Create_TDA(.Tables.Add, "WHTCYCLX", "**", 0, False, "", 1)

            '   With .Tables.Add("WHTCYCLX")
            ' .Columns.Add("BAR_CODE")
            ' .Columns.Add("SEL")
            ' .PrimaryKey = New DataColumn() {.Columns("BAR_CODE")}
            'End With

            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_CODE , ICTSTYL1.STYLE_DESC, ICTSTYC1.UPC_CODE, WHTLOCB1.LOCATION_QTY  " & vbCrLf _
                & " from WHTLOCB1, WHTLOCM1, ICTSTYL1, ICTSTYC1 " & vbCrLf _
                & " where WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE " & vbCrLf _
                & " and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE " & vbCrLf _
                & " and NVL(LOCATION_USE,'A') = 'A' " & vbCrLf _
                & " and WHTLOCB1.WHSE_CODE = 'MS' " & vbCrLf _
                & " and WHTLOCB1.LOCATION_QTY > 0 " & vbCrLf _
                & " and ICTSTYL1.STYLE_CODE = WHTLOCB1.STYLE_CODE " & vbCrLf _
                & " and ICTSTYC1.STYLE_CODE = WHTLOCB1.STYLE_CODE " & vbCrLf _
                & " and ICTSTYC1.COLOR_CODE = WHTLOCB1.COLOR_CODE " & vbCrLf _
                & " and (WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE) in ( " & vbCrLf _
                & " select distinct STYLE_CODE, COLOR_CODE " & vbCrLf _
                & " from WHTLOCB1, WHTLOCM1 " & vbCrLf _
                & " where WHTLOCB1.LOCATION_QTY > 0 " & vbCrLf _
                & " and WHTLOCB1.LOCATION_CODE like '%-C' " & vbCrLf _
                & " and WHTLOCB1.WHSE_CODE = 'MS' " & vbCrLf _
                & " and NVL(LOCATION_USE,'A') = 'A' " & vbCrLf _
                & " and WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE " & vbCrLf _
                & " and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE " & vbCrLf _
                & " minus " & vbCrLf _
                & " select distinct STYLE_CODE, COLOR_CODE " & vbCrLf _
                & " from WHTLOCB1, WHTLOCM1 " & vbCrLf _
                & " where WHTLOCB1.LOCATION_QTY > 0 " & vbCrLf _
                & " and WHTLOCB1.LOCATION_CODE not like '%-C' " & vbCrLf _
                & " and WHTLOCB1.WHSE_CODE = 'MS' " & vbCrLf _
                & " and NVL(LOCATION_USE,'A') = 'A' " & vbCrLf _
                & " and WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE " & vbCrLf _
                & " and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE) "
                Create_TDA(.Tables.Add, "WHTRPRT1", "**", 0, False, "", 3)
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select WHTMOVE2.*" & vbCrLf _
                & ", WHTMOVE1.WHSE_TRAN_TYPE, WHTMOVE1.WHSE_CODE, WHTMOVE1.SESSION_NO, DTL.PICK_NO, DTL.ORDR_NO, DTL.CUST_NAME" & vbCrLf _
                & " from WHTMOVE1, WHTMOVE2, ( " & vbCrLf _
                & " select SOTPICK5.WHSE_TRAN_NO, SOTPICK5.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR5.CUST_NAME " & vbCrLf _
                & " from SOTPICK5, SOTPICK1, SOTORDR5 " & vbCrLf _
                & " where SOTPICK5.PICK_NO =  SOTPICK1.PICK_NO " & vbCrLf _
                & " and SOTORDR5.ORDR_NO = SOTPICK1.ORDR_NO and SOTORDR5.CUST_ADDR_TYPE = 'ST') DTL" & vbCrLf _
                & " where WHTMOVE1.WHSE_TRAN_NO = WHTMOVE2.WHSE_TRAN_NO" & vbCrLf _
                & "  and DTL.WHSE_TRAN_NO(+) = WHTMOVE1.WHSE_TRAN_NO " & vbCrLf _
                & "   and WHTMOVE2.WHSE_TRAN_NO = :PARM1"
            Else
                ASCMAIN1.sql = "Select WHTMOVE2.*" & vbCrLf _
                & ", WHTMOVE1.WHSE_TRAN_TYPE, WHTMOVE1.WHSE_CODE, WHTMOVE1.SESSION_NO, '' PICK_NO, '' ORDR_NO, '' CUST_NAME" & vbCrLf _
                & " from WHTMOVE1, WHTMOVE2" & vbCrLf _
                & " where WHTMOVE1.WHSE_TRAN_NO = WHTMOVE2.WHSE_TRAN_NO" & vbCrLf _
                & "   and WHTMOVE2.WHSE_TRAN_NO = :PARM1"
            End If
            Create_TDA(.Tables.Add, "WHTMOVEX", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select " & vbCrLf _
                & If(ASCMAIN1.CLIENT = "RGI", " SUBSTR(X.LOCATION_CODE,1,3) AA", " SUBSTR(X.LOCATION_CODE,1,2) AA") & vbCrLf _
                & If(ASCMAIN1.CLIENT = "RGI", " , SUBSTR(X.LOCATION_CODE,5,2) BBB", ", SUBSTR(X.LOCATION_CODE,4,3) BBB") & vbCrLf _
                & ", SUBSTR(X.LOCATION_CODE,8,1) L " & vbCrLf _
                & ", X.WHSE_CODE, X.LOCATION_CODE, WHTLOCM1.LOCATION_DESC" & vbCrLf _
                & ", WHTLOCM1.LOCATION_SINGLE_LOAD, WHTLOCM1.LOCATION_LOCKED, WHTLOCM1.LOCATION_NOT_WAVED, WHTLOCM1.LOCATION_USE" & vbCrLf _
                & ", X.STYLE_CODE, ICTSTYL1.STYLE_DESC, X.COLOR_CODE, X.CASES, X.UNITS" & vbCrLf _
                & " from (Select" & vbCrLf _
                & "WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", COUNT (DISTINCT BAR_CODE) CASES, SUM (LOCATION_QTY) UNITS" & vbCrLf _
                & " from WHTLOCB1 where LOCATION_QTY <> 0" & vbCrLf _
                & " group by WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE) X, ICTSTYL1, WHTLOCM1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "WHTLOCBJ", "**", 0, False, "V", 0)
            With .Tables("WHTLOCBJ")
                .Columns("CASES").DataType = GetType(System.Int64)
                .Columns("UNITS").DataType = GetType(System.Int64)
                .Columns("STYLE_CODE").AllowDBNull = True
                .Columns("COLOR_CODE").AllowDBNull = True
            End With

            'With .Tables.Add("WHTMOVE3")
            '    .Columns.Add("WHSE_TRAN_NO")
            '    .Columns.Add("STYLE_CODE")
            '    .Columns.Add("COLOR_CODE")
            '    .Columns.Add("CASE_QTY", GetType(System.Int64))
            '    .Columns.Add("STYLE_DESC")
            '    .PrimaryKey = New DataColumn() {.Columns("WHSE_TRAN_NO"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            'End With

            'With .Tables.Add("WHTMOVEC")
            '    .Columns.Add("WHSE_TRAN_NO")
            '    .Columns.Add("BAR_CODE")
            '    .PrimaryKey = New DataColumn() {.Columns("WHSE_TRAN_NO"), .Columns("BAR_CODE")}
            'End With


            ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_LOCATOR = '1' And WHSE_CODE = '" & ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & "'"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select LOCATION_QTY CASE_PACK_NO, LOCATION_QTY CASES" _
                & " from WHTLOCB1 where ROWNUM < 1"
            Create_TDA(.Tables.Add, "WHTLOCBC", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select LOCATION_QTY CASE_PACK_NO, STYLE_CODE, COLOR_CODE, LOCATION_QTY QTY" _
                & " from WHTLOCB1 where ROWNUM < 1"
            Create_TDA(.Tables.Add, "WHTLOCBD", "**", 0, False, "", 3)

            Create_Relation("WHTLOCBC", "WHTLOCBD", "CASE_PACK_NO")

            With .Tables("WHTLOCBC").Columns
                .Add("STYLES", GetType(System.Int32), "COUNT(CHILD.STYLE_CODE)")
                .Add("CASE_PACK_QTY", GetType(System.Int32), "SUM(CHILD.QTY)")
                .Add("UNITS", GetType(System.Int32), "CASE_PACK_QTY * CASES")
                .Add("STYLE_CODE", GetType(System.String), "MIN(CHILD.STYLE_CODE)")
                .Add("COLOR_CODE", GetType(System.String), "MIN(CHILD.COLOR_CODE)")
            End With

            'WHTLOCBM used for looking at details for a style/clr when LOCB1 is summarized and we need to move cartons
            ASCMAIN1.sql = "Select WHTLOCB1.*,WHTBARC1.LOAD_NO,WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
                & " from WHTLOCB1,WHTBARC1,WHTLOCM1" & vbCrLf _
                & " where WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
                & "   and WHTLOCM1.LOCATION_CODE (+) = WHTLOCB1.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCBM", "**", 0, False, "", 5)
            .Tables("WHTLOCBM").Columns.Add("PERSIST")
            .Tables("WHTLOCBM").Columns.Add("LOCATION_QTY_AVAIL", GetType(System.Int64), "ISNULL(LOCATION_QTY,0)-ISNULL(LOCATION_QTY_WAVE,0)")

            ASCMAIN1.sql = "Select WHTINST2.*" & vbCrLf _
                & ",WHTINST1.WAVE_NO,WHTINST1.WAVE_PICK_TYPE,WHTINST1.LOCATION_CODE,WHTINST1.LOAD_NO,WHTINST1.WAVE_INST_STATUS" & vbCrLf _
                & " from WHTINST1,WHTINST2, WHTWAVE1" & vbCrLf _
                & " where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
                & " And WHTINST1.WAVE_NO = WHTWAVE1.WAVE_NO"
            sqlWHTINSTX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "WHTINSTX", "**", 0, False, "", 4)

            ASCMAIN1.sql = "Select * from WHTCYCL1"
            Create_TDA(.Tables.Add, "WHTCYCL1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from WHTCYCL2" _
            & " where CYCLE_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTCYCL2", "**", 0, False, "V", 2)

            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "select WHSE_CODE, STYLE_CODE, COLOR_CODE, QTY, TRANS, FIRST_DT, ( trunc(sysdate) - trunc(FIRST_DT)) FIRST_DAYS, " & vbCrLf _
                & " LAST_DT, ( trunc(sysdate) - trunc(LAST_DT) ) LAST_DAYS " & vbCrLf _
                & " from ( " & vbCrLf _
                & " select WHTLOCB2.WHSE_CODE, WHTLOCB2.STYLE_CODE, WHTLOCB2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND QTY, " & vbCrLf _
                & " COUNT(1) TRANS, min(WHTLOCB2.init_date) FIRST_DT, max(WHTLOCB2.init_date) LAST_DT " & vbCrLf _
                & " from WHTLOCB2, ICTSTYL1, ICTSTAT2 " & vbCrLf _
                & " where ICTSTYL1.CUST_CODE <> 'WALMART' " & vbCrLf _
                & " and ICTSTYL1.STYLE_STATUS = 'A' " & vbCrLf _
                & " and ICTSTAT2.STYLE_CODE =  ICTSTYL1.Style_code" & vbCrLf _
                & " and ICTSTAT2.WHSE_CODE = WHTLOCB2.WHSE_CODE" & vbCrLf _
                & " and ICTSTAT2.STYLE_CODE = WHTLOCB2.STYLE_CODE" & vbCrLf _
                & " and ICTSTAT2.COLOR_CODE = WHTLOCB2.COLOR_CODE" & vbCrLf _
                & " AND ICTSTAT2.WHSE_QTY_ON_HAND <> 0" & vbCrLf _
                & " group by WHTLOCB2.WHSE_CODE, WHTLOCB2.STYLE_CODE, WHTLOCB2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND)" & vbCrLf _
                & " where ROWNUM < 1" & vbCrLf
                Create_TDA(.Tables.Add, "WHTCYCLS", "**", 0, False, "", 3)
                With .Tables("WHTCYCLS").Columns
                    .Add("SCORE", GetType(System.Decimal), "(qty * .5) +  trans + FIRST_DAYS + LAST_DAYS")
                End With

                ASCMAIN1.sql = "Select * from WHTCYCL3" & vbCrLf _
                & " where WHSE_CODE = :PARM1" & vbCrLf _
                & " and STYLE_CODE = :PARM2" & vbCrLf _
                & " and COLOR_CODE = :PARM3" & vbCrLf
                Create_TDA(.Tables.Add, "WHTCYCL3", "**", 0, False, "VVV", 2)
            End If


            ASCMAIN1.sql = "Select POTORDR1.INIT_DATE, POTORDR1.WHSE_CODE, POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & ", POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
                & ", POTORDR1.FACTORY_CODE, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
                & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO" & vbCrLf _
                & ", POTSHIP2.PO_DATE_RECEIVED" & vbCrLf _
                & ", POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_REC" & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
                & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
                & " From POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2 " & vbCrLf _
                & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 0)
            With .Tables("POTORDRX")
                .Columns("PO_SHIPMENT_NO").AllowDBNull = True
                .Columns("PO_SHIPMENT_LNO").AllowDBNull = True
                .Columns("PO_REFERENCE").AllowDBNull = True
                '  .Columns("PO_SHIPMENT_NO").AllowDBNull = True
            End With

            ASCMAIN1.sql = "Select * from WHTLOCM1 where WHSE_CODE = :PARM1 and LOCATION_LOCKED = '1' AND LOCATION_USE IS NULL"
            Create_TDA(.Tables.Add, "WHTLOCML", "**", 0, False, "V", 2)

        End With

        Fill_Records("ICTWHSE1")
        Fill_Records("WHTLOCMM")

        grdWHTWMAPA.DataSource = dst.Tables("WHTWMAPA")
        grdWHTLOCBR.DataSource = dst.Tables("WHTLOCBR")
        grdWHTLOCB1.DataSource = dst.Tables("WHTLOCB1")
        grdWHTLOCB2.DataSource = dst.Tables("WHTLOCB2")
        grdWHTMOVEX.DataSource = dst.Tables("WHTMOVEX")
        grdWHTLOCM1.DataSource = dst.Tables("WHTLOCM1")
        grdWHTLOCMA.DataSource = dst.Tables("WHTLOCMA")
        grdICTSTYL1.DataSource = dst.Tables("ICTSTYL1")
        grdICTSTYLA.DataSource = dst.Tables("ICTSTYLA")
        grdWHTLOCMM.DataSource = dst.Tables("WHTLOCMM")
        grdWHTINSTX.DataSource = dst.Tables("WHTINSTX")
        grdWHTLOCML.DataSource = dst.Tables("WHTLOCML")
        grdWHTCYCL1.DataSource = dst.Tables("WHTCYCL1")
        grdWHTCYCL2.DataSource = dst.Tables("WHTCYCL2")
        grdWHTLOCBC.DataSource = dst.Tables("WHTLOCBC")
        grdWHTLOCBJ.DataSource = dst.Tables("WHTLOCBJ")
        grdWHTCONTC.DataSource = dst.Tables("WHTCONTC")
        grdWHTLOCBM.DataSource = dst.Tables("WHTLOCBM")

        If Not IsNothing(dst.Tables("WHTCYCLS")) Then
            grdWHTCYCLS.DataSource = dst.Tables("WHTCYCLS")
            grdWHTCYCL3.DataSource = dst.Tables("WHTCYCL3")
        End If

        Create_Summary(grdWHTLOCB1, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCB1, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCB1, New String() {"LOCATION_QTY", "LOCATION_QTY_WAVE", "LOCATION_QTY_AVAIL"})

        Create_Summary(grdWHTLOCBM, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCBM, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCBM, New String() {"LOCATION_QTY", "LOCATION_QTY_WAVE", "LOCATION_QTY_AVAIL"})

        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_TYPE", "Count")
        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_QTY")

        Create_Summary(grdWHTLOCBJ, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCBJ, "CASES")
        Create_Summary(grdWHTLOCBJ, "UNITS")

        Create_Summary(grdWHTMOVEX, "WHSE_TRAN_LNO", "Count")
        Create_Summary(grdWHTMOVEX, "WHSE_TRAN_QTY")

        Create_Summary(grdWHTLOCMM, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCMM, New String() {"SELECTED", "CASES", "QTY_LOCS", "QTY_WAVE"})

        Create_Summary(grdWHTLOCBC, "CASE_PACK_NO", "Count")
        Create_Summary(grdWHTLOCBC, New String() {"CASES", "UNITS"})

        Create_Summary(grdWHTLOCBR, "BAR_CODE", "Count")
        Create_Summary(grdWHTLOCBR, New String() {"INST", "LOCB", "DIFF"})

        grdWHTLOCMM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For Each gcol As UltraWinGrid.UltraGridColumn In grdWHTLOCMM.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SELECTED" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
            If gcol.Key = "LOCATION_CODE" Then
                ' gcol.Format = "##-AAA-A"
                ' gcol.MaskInput = "##-AAA-A"
                ' gcol.CellDisplayStyle = UltraWinGrid.CellDisplayStyle.FormattedText

                'gcol.MaskInput = "AA-AAA-A"
                'gcol.MaskDataMode = UltraWinMaskedEdit.MaskMode.Raw
                'gcol.MaskClipMode = UltraWinMaskedEdit.MaskMode.IncludeBoth
                'gcol.MaskDisplayMode = UltraWinMaskedEdit.MaskMode.IncludeBoth

            End If
        Next

        With grdWHTLOCB1.DisplayLayout.Bands(0)
            .Columns("LOCATION_CODE").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("STYLE_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("COLOR_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("BAR_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOAD_NO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOCATION_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LOCATION_QTY_WAVE").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LOCATION_QTY_AVAIL").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LAST_OPER").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("LAST_DATE").Header.Appearance.BackColor2 = Drawing.Color.Orange
        End With

        With grdWHTLOCBM.DisplayLayout.Bands(0)
            .Columns("LOCATION_CODE").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("STYLE_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("COLOR_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("BAR_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOAD_NO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOCATION_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LOCATION_QTY_WAVE").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LOCATION_QTY_AVAIL").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LAST_OPER").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("LAST_DATE").Header.Appearance.BackColor2 = Drawing.Color.Orange
        End With

        With grdWHTLOCB2.DisplayLayout.Bands(0)
            .Columns("WHSE_TRAN_NO").Header.Appearance.BackColor2 = Drawing.Color.Tan
            .Columns("WHSE_TRAN_LNO").Header.Appearance.BackColor2 = Drawing.Color.Tan
            .Columns("BAR_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("BAR_CODE_OTHER").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOAD_NO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOAD_NO_OTHER").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOCATION_CODE_OTHER").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("WHSE_TRAN_TYPE").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("WHSE_TRAN_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("INIT_OPER").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("INIT_DATE").Header.Appearance.BackColor2 = Drawing.Color.Orange
        End With

        With grdWHTMOVEX.DisplayLayout.Bands(0)
            .Columns("WHSE_TRAN_NO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("WHSE_TRAN_LNO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOCATION_CODE_FROM").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("LOCATION_CODE_TO").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("WHSE_TRAN_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("BAR_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("STYLE_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("COLOR_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("INIT_OPER").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("INIT_DATE").Header.Appearance.BackColor2 = Drawing.Color.Orange
        End With

        If ASCMAIN1.CLIENT = "RGI" Then
            For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In New Infragistics.Win.UltraWinGrid.UltraGrid() {grdWHTLOCB1, grdWHTLOCB2, grdWHTMOVEX}
                With grd.DisplayLayout.Bands(0)
                    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                        If New String() {"INIT_DATE", "LAST_DATE"}.Contains(gcol.Key) Then
                            gcol.Format = "MM/dd/yy HH:mm"
                        End If
                    Next
                End With
            Next
        End If

        grpStyle.Top = grpLocation.Top
        grpStyle.Left = grpLocation.Left
        grpDetail.Top = grpLocation.Top

        Integrity_Check(False)

        ASCMAIN1.Add_Value_List(grdWHTINSTX, "WAVE_PICK_TYPE")
        ASCMAIN1.Add_Value_List(grdWHTLOCMM, "LOCATION_USE")
        ASCMAIN1.Add_Value_List(grdWHTLOCBJ, "LOCATION_USE")
        ASCMAIN1.Add_Value_List(grdWHTLOCB2, "WHSE_TRAN_TYPE")
        ASCMAIN1.Add_Value_List(grdWHTINSTX, "WAVE_INST_STATUS")

        Show_Filter(grdWHTLOCMM, True)
        grdWHTLOCMM.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdWHTCYCL1, "CYCLE_STATUS", , New String() {":", "G:GOOD", "D:DISCREPANCY"})
        ASCMAIN1.Add_Value_List(grdWHTCYCL1, "CYCLE_RESOLUTION", , New String() {":", "U:POSTED", "V:VOID", "G:LOCATION OK", "P:PENDING"})
        ASCMAIN1.Add_Value_List(grdWHTCYCL1, "CYCLE_TYPE", , New String() {":", "C:COUNT", "V:VERIFY"})


        ASCMAIN1.Add_Value_List(grdWHTLOCMM, "CYCLE_STATUS", , New String() {":", "G:GOOD", "D:DISCREPANCY"})
        ASCMAIN1.Add_Value_List(grdWHTLOCMM, "CYCLE_RESOLUTION", , New String() {":", "U:POSTED", "V:VOID", "G:LOCATION OK", "P:PENDING"})
        ASCMAIN1.Add_Value_List(grdWHTLOCMM, "CYCLE_TYPE", , New String() {":", "C:COUNT", "V:VERIFY"})

        Toggle_Activity(False)

        txtStyle_Search.Text = "[2-9]AS| AS |ASST"

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            tabMain.Tabs("Warehouse Map").Visible = True
        Else
            splWHTLOCB1.Panel2Collapsed = True
            tabMain.Tabs("Warehouse Map").Visible = False
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"
                Validate_Code("WHSE_CODE")
                If cdr IsNot Nothing Then
                    rowICTWHSE1 = cdr

                    WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

                    If rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
                        EMsg &= vbCr & "Warehouse " & WHSE_CODE & " is NOT set up for Location"
                    End If
                End If

                If optViewBy.Value <> "S" Then

                    If txtBAR_CODE.Text <> "" Then
                        If Absx1.txtFor("WHSE_CODE").Text = "" Then
                            EMsg &= vbCr & "You Must First Specify a Whse Before Trying to Locate by Barcode"
                        Else

                            ASCMAIN1.sql = "Select LOCATION_CODE, LOCATION_QTY from WHTLOCB1" _
                                & " where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" _
                                & "   and BAR_CODE = '" & Absx1.txtFor("BAR_CODE").Text & "'"
                            For Each rowLOCATION_CODE As DataRow In ASCDATA1.GetDataTable.Select("", "LOCATION_QTY DESC")
                                LOCATION_CODE = rowLOCATION_CODE.Item("LOCATION_CODE")
                                Exit For
                            Next

                            If LOCATION_CODE <> "" Then
                                Absx1.txtFor("LOCATION_CODE").Text = LOCATION_CODE
                            Else
                                EMsg &= vbCr & "Invalid Barcode"
                                txtBAR_CODE.Focus()
                                txtBAR_CODE.SelectAll()
                            End If
                        End If
                    Else
                        LOCATION_CODE = Absx1.txtFor("LOCATION_CODE").Text
                        If LOCATION_CODE = "" Then
                            EMsg &= vbCr & "You Must First Specify a Location"
                        Else
                            Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE})
                            If rowWHTLOCM1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Location (" & LOCATION_CODE & ")"
                            Else
                                If "00-REC-A,00-REC-B".Contains(LOCATION_CODE) Or ("G,D,S").Contains(rowWHTLOCM1("LOCATION_USE") & "") Then
                                    chkBAR_CODE.Checked = False
                                    chkLOAD_NO.Checked = False
                                End If
                            End If
                        End If
                    End If


                End If

                If optViewBy.Value = "S" Then
                    STYLE_CODE = Absx1.txtFor("STYLE_CODE").Text
                    If STYLE_CODE = "" Then
                        EMsg &= vbCr & "You Must First Specify a Style"
                    Else
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", New String() {STYLE_CODE})
                        If rowICTSTYL1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Style (" & STYLE_CODE & ")"
                        End If
                    End If
                End If

                'If InquireBy = "RANGE" Then
                '    If Absx1.txtFor("LOCATION_CODE").Text = "" And Absx1.txtFor("LOCATION_CODE2").Text = "" Then
                '        EMsg = EMsg & vbCr & "You Must First Specify Both Locations"
                '    End If
                '    If Mid(txtCode(0).Text, 1, 2) > Mid(txtCode(2).Text, 1, 2) Then
                '        EMsg = EMsg & vbCr & "First Rack Can Not Be Greater Than Second Rack"
                '    End If
                '    If Mid(txtCode(0).Text, 3, 2) > Mid(txtCode(2).Text, 3, 2) Then
                '        EMsg = EMsg & vbCr & "First Column Can Not Be Greater Than Second Column"
                '    End If
                '    If Mid(txtCode(0).Text, 5, 1) > Mid(txtCode(2).Text, 5, 1) Then
                '        EMsg = EMsg & vbCr & "First Level Can Not Be Greater Than Second Level"
                '    End If
                'End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Select"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Print"
                Print_Record()

            Case "Done"
                Mode_Settings(False)

            Case "Integrity Check"
                If ASCMAIN1.CLIENT = "VAN" Then
                    Integrity_Check(True, True)
                Else
                    Integrity_Check(True)
                End If


        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Select").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode And ASCMAIN1.CLIENT = "RGI"
                    .Items("Integrity Check").Settings.Enabled = not_iScreenMode
                End With
                .Groups("Locked Locations").Visible = False
                With .Groups("Special")
                    .Visible = False '(ASCMAIN1.CLIENT = "RGI")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpDetail, False)

        splWHTLOCBX.Visible = tf
        ' splVisited.Visible = Not tf
        tabMain.Visible = Not tf

        If Not IsNothing(dst.Tables("WHTCYCLS")) Then
            tabMain.Tabs.Item("Counts by Style").Visible = False

            'For Each tab As UltraWinTabControl.UltraTab In tabMain.Tabs().t
            '    If tab.Key = "Count by Style" Then

            '    End If
            'Next
        Else

        End If

        grpDetail.Visible = False

        If ScreenMode Then

            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOCATION_CODE").Hidden = Not (optViewBy.Value = "S")
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = (optViewBy.Value = "S")

            Dim WHSE_CTN_CTL As String = rowICTWHSE1.Item("WHSE_CTN_CTL") & ""
            Dim hide_BAR_CODE As Boolean = Not (WHSE_CTN_CTL = "L" Or WHSE_CTN_CTL = "C")
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("BAR_CODE").Hidden = hide_BAR_CODE
            grdWHTLOCB2.DisplayLayout.Bands(0).Columns("BAR_CODE").Hidden = hide_BAR_CODE
            grdWHTMOVEX.DisplayLayout.Bands(0).Columns("BAR_CODE").Hidden = hide_BAR_CODE
            If ASCMAIN1.CLIENT = "RGI" Then
                grdWHTMOVEX.DisplayLayout.Bands(0).Columns("PICK_NO").Hidden = False
                grdWHTMOVEX.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = False
                grdWHTMOVEX.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = False
            Else
                grdWHTMOVEX.DisplayLayout.Bands(0).Columns("PICK_NO").Hidden = True
                grdWHTMOVEX.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = True
                grdWHTMOVEX.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
            End If

            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOCATION_QTY_WAVE").Hidden = Not (WHSE_CTN_CTL = "C")
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOCATION_QTY_AVAIL").Hidden = Not (WHSE_CTN_CTL = "C")

            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOAD_NO").Hidden = hide_BAR_CODE
            grdWHTLOCB2.DisplayLayout.Bands(0).Columns("LOAD_NO").Hidden = hide_BAR_CODE

            With grdWHTLOCB2.DisplayLayout.Bands(0)
                .Columns("LOAD_NO_OTHER").Hidden = hide_BAR_CODE
                .Columns("BAR_CODE_OTHER").Hidden = hide_BAR_CODE
            End With
            'grdWHTMOVEX.DisplayLayout.Bands(0).Columns("LOAD_NO").Hidden = hide_BAR_CODE

            With grdWHTINSTX.DisplayLayout.Bands(0)
                .Columns("LOCATION_CODE").Hidden = (optViewBy.Value = "L")
                .Columns("STYLE_CODE").Hidden = (optViewBy.Value = "S")
            End With

            grpDetail.Visible = (WHSE_CTN_CTL = "C")
            If WHSE_CTN_CTL = "C" Then Toggle_grpDetail()

            If optViewBy.Value = "L" Then
                Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {WHSE_CODE, Absx1.txtFor("LOCATION_CODE").Text})
                lblLOCATION_SINGLE_LOAD.Visible = (rowWHTLOCM1.Item("LOCATION_SINGLE_LOAD") & "" = "1")
                lblLOCATION_LOCKED.Visible = (rowWHTLOCM1.Item("LOCATION_LOCKED") & "" = "1")
                lblLOCATION_NOT_WAVED.Visible = (rowWHTLOCM1.Item("LOCATION_NOT_WAVED") & "" = "1")
            End If


            If optViewBy.Value = "L" And Absx1.txtFor("BAR_CODE").Text <> "" Then
                For Each grow As UltraWinGrid.UltraGridRow In grdWHTLOCB1.Rows
                    If grow.IsDataRow Then
                        If grow.Cells("BAR_CODE").Value = Absx1.txtFor("BAR_CODE").Text Then
                            grdWHTLOCB1.ActiveRow = grow
                            grdWHTLOCB1.DisplayLayout.RowScrollRegions(0).ScrollRowIntoView(grow)
                        End If
                    End If
                Next
            End If

        Else
            Clear_Record()
            Setup_ViewBy()
            chkRESOLUTION.Checked = False
            lblLOCATION_SINGLE_LOAD.Visible = False
            lblLOCATION_LOCKED.Visible = False
            lblLOCATION_NOT_WAVED.Visible = False
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() _
            {"WHTLOCB1", "WHTLOCB2", "WHTMOVEX", "WHTMOVE1", "WHTMOVE2", "WHTLOCMM", "WHTLOCBJ", "WHTCONTC", "WHTLOCBM"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Absx1.txtFor("STYLE_CODE").Clear()
        Absx1.txtFor("LOCATION_CODE").Clear()
        Absx1.txtFor("LOCATION_CODE2").Clear()
        Absx1.txtFor("BAR_CODE").Clear()

        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            If dst.Tables("ICTWHSE1").Rows.Count = 1 Then
                Absx1.txtFor("WHSE_CODE").Text = dst.Tables("ICTWHSE1").Rows(0).Item("WHSE_CODE")
            End If
        End If

        If tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Location Master" Then
            Setup_Location_Grids()
        End If

        grdWHTCYCL1.Parent = splCycle_View.Panel1
        grdWHTCYCL2.Parent = splCycle_View.Panel2
        tabCycle.Tabs("Cycle Count").Visible = False
        tabCycle.Tabs("Barcodes").Visible = False

        Fill_Records("ICTSTYLA", Absx1.txtFor("WHSE_CODE").Text)
        Sort_grdColumns(grdICTSTYLA, "INIT_DATE".ToLower)


        ASCMAIN1.sql = "Select X.WHSE_CODE, X.LOCATION_CODE, X.TRANS, X.INIT_DATE" _
                & " from WHTLOCM1,(" _
                & "Select WHSE_CODE, LOCATION_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
                & " from WHTLOCB2 where INIT_DATE > SYSDATE - 24 group by WHSE_CODE, LOCATION_CODE" _
                & ") X where WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE" _
                & " And WHTLOCM1.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        Fill_Records("WHTLOCMA", , , ASCMAIN1.sql)
        'Fill_Records("WHTLOCMA", Absx1.txtFor("WHSE_CODE").Text)
        Sort_grdColumns(grdWHTLOCMA, "INIT_DATE".ToLower)
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        sqlWHTLOCB2where = ""

        If optViewBy.Value = "S" Then
            chkMain.Text = "Location"
            Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            If dst.Tables("ICTSTYL1").Rows.Find(New String() {STYLE_CODE}) Is Nothing Then
                dst.Tables("ICTSTYL1").Rows.Add(New String() {STYLE_CODE, rowICTSTYL1.Item("STYLE_DESC") & ""})
            End If
            Sort_grdColumns(grdICTSTYL1, "STYLE_CODE")

            sqlWHTLOCB2where = "" _
                & " where WHTLOCB2.STYLE_CODE = '" & STYLE_CODE & "'" _
                & " and WHTLOCB2.WHSE_CODE = '" & WHSE_CODE & "'"

            grdWHTLOCB1.Text = "Locations containing Style " & STYLE_CODE

        Else
            chkMain.Text = "Style"


            Dim LOCATION_CODE As String = Absx1.txtFor("LOCATION_CODE").Text
            If dst.Tables("WHTLOCM1").Rows.Find(New String() {WHSE_CODE, LOCATION_CODE}) Is Nothing Then
                dst.Tables("WHTLOCM1").Rows.Add(New String() {WHSE_CODE, LOCATION_CODE})
            End If
            Sort_grdColumns(grdWHTLOCM1, "WHSE_CODE, LOCATION_CODE")

            sqlWHTLOCB2where = "" _
                & " where WHTLOCB2.LOCATION_CODE = '" & LOCATION_CODE & "'" _
                & " and WHTLOCB2.WHSE_CODE = '" & WHSE_CODE & "'"


            grdWHTLOCB1.Text = "Styles in Location " & LOCATION_CODE


            ASCMAIN1.sql = "Select * from WHTCYCL1 Where LOCATION_CODE = '" & LOCATION_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Fill_Records("WHTCYCL1", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdWHTCYCL1, "CYCLE_NO".ToLower)

            grdWHTCYCL1.Parent = splCycle_Edit.Panel1
            grdWHTCYCL2.Parent = splWHTCYCL2.Panel1
            tabCycle.Tabs("Cycle Count").Visible = True
            tabCycle.Tabs("BarCodes").Visible = (rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C")
        End If

        'ASCMAIN1.sql = sqlWHTINSTX & " And WHTWAVE1.WHSE_CODE = '" & WHSE_CODE & "'" &
        'IIf(optViewBy.Value = "L",
        '        " and WHTINST1.LOCATION_CODE = '" & LOCATION_CODE & "'",
        '        " and WHTINST2.STYLE_CODE = '" & STYLE_CODE & "'")
        'If Not (optViewBy.Value = "L" And "00-REC-A,00-REC-B".Contains(LOCATION_CODE)) Then
        '    Fill_Records("WHTINSTX", "", True, ASCMAIN1.sql)
        'End If
        Load_WHTLOCB1()

        Setup_grdWHTLOCB1()
        Setup_grdWHTLOCB2()
        Setup_grdWHTMOVEX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Print_Record()

        Fill_Records("WHTRPRT1")
        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = "WHRLOCS1"
        Generate_Report(RPT, "Inventory Found Only at 'C' Level", "", "")
        Print_Report_End()

    End Sub


    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where = "WHSE_LOCATOR = '1'"
            Case "BAR_CODE"
                sql_where = "WHSE_CODE = '" & WHSE_CODE & "'"
        End Select

    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Select"
                If key.StartsWith("S:") Then
                    optViewBy.Value = "S"
                    Absx1.txtFor("STYLE_CODE").Text = Split(key, ":")(1)
                    Click_Command("Select")
                ElseIf key.StartsWith("L:") Then
                    optViewBy.Value = "L"
                    Absx1.txtFor("LOCATION_CODE").Text = Split(key, ":")(1)
                    Click_Command("Select")
                End If

        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdWHTLOCB1, "SSSSBBBBBBBBBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Show 0 Qty", "Move to ...", "Move to Bin", "Move to LNF", "Location Inquiry", "Void/Replace Case ID", "Lock Location", "Un-Lock Location", "Re-Cartonize", "Re-Configure", "Create Cases", "Adjust", "Change Style", "Back to Stock", "Combine Cases", "Transfer Cases")
        Load_Popup_Menu(grdWHTLOCBM, "SSSSBBBBBBBBBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Show 0-Qty", "Move to ...", "Move to Bin", "Move to LNF", "Location Inquiry", "Void/Replace Case ID", "Lock Location", "Un-Lock Location", "Re-Cartonize", "Re-Configure", "Create Cases", "Adjust", "Change Style", "Back to Stock", "Combine Cases", "Transfer Cases")
        Load_Popup_Menu(grdWHTLOCB2, "BB", "Reverse Entire Move (All Lines shown Below)", "Reverse This Move Line Only")
        Load_Popup_Menu(grdWHTLOCMA, "B", "Load All Locations with Non-0 Status")
        Load_Popup_Menu(grdICTSTYLA, "B", "Load All Styles with Non-0 Status")
        Load_Popup_Menu(grdWHTLOCMM, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Print Placards", "Select Selected", "De-Select All", "Get Activity")
        Load_Popup_Menu(grdWHTLOCBR, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdWHTLOCBJ, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdWHTINSTX, "B", "Wave Inquiry")
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
            'e.Cancel = True
        Else
            Dim allow_reverse As Boolean = (ASCMAIN1.USER_SECURITY_CODEs.Contains("WH"))
            If grd.Name = "grdWHTLOCB2" Then
                allow_reverse = allow_reverse AndAlso grd.ActiveRow IsNot Nothing And grd.ActiveRow.Cells("WHSE_TRAN_TYPE").Value = "M"
            End If

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                'allow_reverse = False
            End If

            SOURCE_GRID = e.SourceControl.Name

            Select Case e.SourceControl.Name
                Case "grdWHTLOCB1"
                    Dim WHSE_CTN_CTL As String = rowICTWHSE1.Item("WHSE_CTN_CTL") & ""
                    tlb_pop.Tools("Move to ...").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL <> "C" Or (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))
                    tlb_pop.Tools("Move to Bin").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL <> "C")
                    tlb_pop.Tools("Move to LNF").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))
                    tlb_pop.Tools("Re-Cartonize").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))
                    tlb_pop.Tools("Re-Configure").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))
                    tlb_pop.Tools("Create Cases").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True)) ' And optViewBy.Value = "S"
                    tlb_pop.Tools("Adjust").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))
                    tlb_pop.Tools("Back to Stock").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))
                    tlb_pop.Tools("Change Style").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True)) And optViewBy.Value = "S"
                    tlb_pop.Tools("Combine Cases").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))   '  And optViewBy.Value = "L" ' GONNA CHECK SAME STYLES BELOW
                    tlb_pop.Tools("Transfer Cases").SharedProps.Visible = allow_reverse And (WHSE_CTN_CTL = "C" And (chkBAR_CODE.Checked And chkLOAD_NO.Checked And chkCOLOR_CODE.Checked = True))   '  And optViewBy.Value = "L" ' GONNA CHECK SAME STYLES BELOW

                    tlb_pop.Tools("Lock Location").SharedProps.Visible = (optViewBy.Value = "L" And Not lblLOCATION_LOCKED.Visible)
                    tlb_pop.Tools("Un-Lock Location").SharedProps.Visible = (optViewBy.Value = "L" And lblLOCATION_LOCKED.Visible)

                    tlb_pop.Tools("Void/Replace Case ID").SharedProps.Visible = chkBAR_CODE.Checked And (WHSE_CTN_CTL = "C")

                Case "grdWHTLOCBM"
                    Dim WHSE_CTN_CTL As String = rowICTWHSE1.Item("WHSE_CTN_CTL") & ""
                    tlb_pop.Tools("Move to ...").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Move to Bin").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Move to LNF").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Re-Cartonize").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Re-Configure").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Create Cases").SharedProps.Visible = allow_reverse  ' And optViewBy.Value = "S"
                    tlb_pop.Tools("Adjust").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Back to Stock").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Change Style").SharedProps.Visible = False
                    tlb_pop.Tools("Combine Cases").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Transfer Cases").SharedProps.Visible = allow_reverse

                    tlb_pop.Tools("Lock Location").SharedProps.Visible = False
                    tlb_pop.Tools("Un-Lock Location").SharedProps.Visible = False

                    tlb_pop.Tools("Void/Replace Case ID").SharedProps.Visible = True


                Case "grdWHTLOCB2"
                    tlb_pop.Tools("Reverse Entire Move (All Lines shown Below)").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Reverse This Move Line Only").SharedProps.Visible = allow_reverse
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Print Placards"
                If dst.Tables("WHTLOCMM").Select("SELECTED = '1'").Length = 0 Then
                    MsgBox("You must first select locations before using this command", MsgBoxStyle.OkOnly, "No Locations Selected")
                Else
                    Print_Report_Begin()
                    CR_params.Add("SUBT", "")
                    Generate_Report("WHRLOCMM", "Location Placards")
                    Print_Report_End()
                End If

            Case "De-Select All"
                For Each row As DataRow In dst.Tables("WHTLOCMM").Select("SELECTED = '1'")
                    row.Item("SELECTED") = "0"
                Next

            Case "Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If Not grow.IsFilteredOut And grow.IsDataRow Then
                        grow.Cells("SELECTED").Value = "1"
                    End If
                Next
                grd.UpdateData()

            Case "Show 0 Qty"
                Setup_grdWHTLOCB1()

            Case "Show 0-Qty"
                Load_WHTLOCBM()

            Case "Show Case Packs"
                Toggle_Case_Packs()

            Case "Move to ...", "Move to Bin", "Move to LNF", "Re-Cartonize", "Re-Configure", "Adjust", "Back to Stock", "Create Cases", "Transfer Cases"
                If e.Tool.Key = "Transfer Cases" Then
                    'all rows in a location go.
                    grdWHTLOCB1.Selected.Rows.AddRange(grdWHTLOCB1.Rows.All)
                End If
                If e.Tool.Key = "Create Cases" Then
                    ' NO SELECTIONS REQUIRED
                Else
                    If grdWHTLOCB1.Selected.Rows.Count = 0 Then
                        grdWHTLOCB1.ActiveRow.Selected = True
                    Else
                        If Not grdWHTLOCB1.ActiveRow.Selected Then
                            Exit Sub
                        End If
                    End If
                End If

                If e.Tool.Key = "Re-Configure" Then

                    Dim STYLE_CODE As String = ""
                    Dim COLOR_CODE As String = ""

                    For Each row As UltraWinGrid.UltraGridRow In grdWHTLOCB1.Selected.Rows
                        If STYLE_CODE = "" Then
                            STYLE_CODE = row.Cells("STYLE_CODE").Value & ""
                            COLOR_CODE = row.Cells("COLOR_CODE").Value & ""
                            ' LOCATION_CODE = row.Cells("LOCATION_CODE").Value & ""
                        Else
                            If STYLE_CODE <> row.Cells("STYLE_CODE").Value & "" _
                            Or COLOR_CODE <> row.Cells("COLOR_CODE").Value & "" Then
                                MsgBox("All selected Cases must be of the same Style & Color", MsgBoxStyle.OkOnly, "Cannot Re-Configure")
                                Exit Sub
                            End If
                        End If

                        Dim LOCATION_CODE As String = row.Cells("LOCATION_CODE").Value & ""
                        Dim BAR_CODE As String = row.Cells("BAR_CODE").Value & ""

                        ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                            & " where WHSE_CODE= '" & WHSE_CODE & "'" & vbCrLf _
                            & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                            & "   and BAR_CODE = '" & BAR_CODE & "'"
                        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                        If rows.Length > 1 Then
                            MsgBox("All selected Cases must be of the same Composition (1 Style/Color)" & vbCrLf & " - see " & BAR_CODE & " in " & LOCATION_CODE & " with " & CStr(rows.Length), MsgBoxStyle.OkOnly, "Cannot Re-Configure")
                            Exit Sub
                        End If
                    Next
                End If

                'If e.Tool.Key = "Back to Stock" Then
                '    If grdWHTLOCB1.Selected.Rows.Count <> 1 Then
                '        MsgBox("Back to Stock must be done 1 Case at a time", MsgBoxStyle.OkOnly, "Cannot BTS")
                '        Exit Sub
                '    End If
                'End If
                Dim LOCATION_CODE_TO As String = ""
                If e.Tool.Key = "Move to Bin" Then
                    LOCATION_CODE_TO = ""
                ElseIf e.Tool.Key = "Move to LNF" Then
                    LOCATION_CODE_TO = rowICTWHSE1.Item("WHSE_LOC_LNF") & ""
                End If
                Move_To(LOCATION_CODE_TO,
                        IIf(e.Tool.Key = "Move to LNF", "LNF",
                        IIf(e.Tool.Key = "Re-Cartonize", "CTN",
                        IIf(e.Tool.Key = "Re-Configure", "CFG",
                        IIf(e.Tool.Key = "Create Cases", "CCS",
                        IIf(e.Tool.Key = "Adjust", "ADJ",
                        IIf(e.Tool.Key = "Back to Stock", "BTS",
                        IIf(e.Tool.Key = "Transfer Cases", "TRN",
                            ""))))))))

            Case "Change Style", "Combine Cases"
                'If ASCMAIN1.USER_ID = "gcv" Then
                '    Stop
                'GCV_Fix()
                '    Stop
                'End If

                If grdWHTLOCB1.Selected.Rows.Count = 0 Then
                    If grdWHTLOCB1.ActiveRow IsNot Nothing AndAlso grdWHTLOCB1.ActiveRow.IsDataRow Then
                        grdWHTLOCB1.ActiveRow.Selected = True
                    End If
                End If
                If grdWHTLOCB1.Selected.Rows.Count = 0 Then Exit Sub
                If e.Tool.Key = "Combine Cases" Then
                    If grdWHTLOCB1.Selected.Rows.Count <> 2 Then
                        MsgBox("You Must Select 2 Cases to Combine", MsgBoxStyle.OkOnly, "Cannot Combine")
                        Exit Sub
                    End If
                End If

                Dim LOCATION_CODE_TO As String = ""
                Dim STYLE_CODE_orig As String = ""
                Dim COLOR_CODE_orig As String = ""
                For Each row As UltraWinGrid.UltraGridRow In grdWHTLOCB1.Selected.Rows
                    If STYLE_CODE_orig = "" Then
                        STYLE_CODE_orig = row.Cells("STYLE_CODE").Value & ""
                        COLOR_CODE_orig = row.Cells("COLOR_CODE").Value & ""
                        LOCATION_CODE_TO = row.Cells("LOCATION_CODE").Value & ""
                    Else
                        If STYLE_CODE_orig <> row.Cells("STYLE_CODE").Value & "" _
                        Or COLOR_CODE_orig <> row.Cells("COLOR_CODE").Value & "" Then
                            MsgBox("All selected Cases must be of the same Style & Color", MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                            Exit Sub
                        Else
                            If e.Tool.Key = "Change Style" Then
                                Dim BAR_CODE As String = row.Cells("BAR_CODE").Value
                                Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", BAR_CODE)
                                If rowWHTBARC1.Item("PPK_CODE") & "" <> "" Then
                                    MsgBox("This option is not available to Pre-Packs", MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Next

                If e.Tool.Key = "Change Style" Then
                    Move_To(LOCATION_CODE_TO, "CHG")
                Else
                    Move_To(LOCATION_CODE_TO, "CMB")
                End If


            Case "Load All Styles with Non-0 Status"

                ASCMAIN1.sql = "Select X.STYLE_CODE, ICTSTYL1.STYLE_DESC, X.TRANS, X.INIT_DATE" _
                    & " from ICTSTYL1,(" _
                    & "Select STYLE_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
                    & " from WHTLOCB2 where STYLE_CODE in (Select Distinct STYLE_CODE from WHTLOCB1 where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "' and LOCATION_QTY <> 0) group by STYLE_CODE" _
                    & ") X where ICTSTYL1.STYLE_CODE = X.STYLE_CODE"
                Fill_Records("ICTSTYLA", "", True, ASCMAIN1.sql)

            Case "Load All Locations with Non-0 Status"

                ASCMAIN1.sql = "Select X.WHSE_CODE, X.LOCATION_CODE, X.TRANS, X.INIT_DATE" _
                     & " from WHTLOCM1,(" _
                     & "Select WHSE_CODE, LOCATION_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
                     & " from WHTLOCB2 where LOCATION_CODE in (Select Distinct LOCATION_CODE from WHTLOCB1 where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "' and LOCATION_QTY <> 0) group by WHSE_CODE, LOCATION_CODE" _
                     & ") X where WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE"
                Fill_Records("WHTLOCMA", "", True, ASCMAIN1.sql)

            Case "Get Activity"
                GetActivity()

            Case "Lock Location", "Un-Lock Location"
                If Not ASCMAIN1.Logical_Lock("WHTLOCM1", WHSE_CODE & ":" & LOCATION_CODE) Then Exit Sub
                Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE})
                Dim LOCATION_LOCKED As String = rowWHTLOCM1.Item("LOCATION_LOCKED") & ""
                If LOCATION_LOCKED = "1" Then
                    If e.Tool.Key = "Lock Location" Then
                        MsgBox("Location " & LOCATION_CODE & " is Already Locked", MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                    Else
                        If MsgBox("OK to " & e.Tool.Key & " " & LOCATION_CODE, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            ASCMAIN1.Record_Event("WHTLOCM1", WHSE_CODE, LOCATION_CODE, Now, ASCMAIN1.USER_ID, "LOCUNL", "UNL LOCINQ", "")

                            ASCMAIN1.sql = "Update WHTLOCM1 Set LOCATION_LOCKED = NULL" _
                                & " where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {WHSE_CODE, LOCATION_CODE})
                            lblLOCATION_LOCKED.Visible = False
                            MsgBox("Location " & LOCATION_CODE & " Successfully Un-Locked")
                        End If
                    End If
                Else
                    If e.Tool.Key = "Un-Lock Location" Then
                        MsgBox("Location " & LOCATION_CODE & " is Already Un-Locked", MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                    Else
                        If MsgBox("OK to " & e.Tool.Key & " " & LOCATION_CODE, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            ASCMAIN1.Record_Event("WHTLOCM1", WHSE_CODE, LOCATION_CODE, Now, ASCMAIN1.USER_ID, "LOCLCK", "LCK LOCINQ", "")

                            ASCMAIN1.sql = "Update WHTLOCM1 Set LOCATION_LOCKED = '1'" _
                                & " where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {WHSE_CODE, LOCATION_CODE})
                            lblLOCATION_LOCKED.Visible = True
                            MsgBox("Location " & LOCATION_CODE & " Successfully Locked")
                        End If
                    End If
                End If
                ASCMAIN1.MultiTask_Release()



        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        If grd.Name = "grdWHTLOCB1" Then

            Select Case e.Tool.Key
                Case "Location Inquiry"
                    Dim KEY As String = ""
                    If optViewBy.Value = "L" Then
                        Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                        KEY = "S:" & STYLE_CODE
                    Else
                        Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Value
                        KEY = "L:" & LOCATION_CODE
                    End If

                    Context_Launch("Select", KEY, e.Tool.Key, "WHFLOCS1")

                Case "Void/Replace Case ID"
                    Dim BAR_CODE As String = grd.ActiveRow.Cells("BAR_CODE").Value & ""
                    Void_and_Replace(BAR_CODE)

            End Select
        End If

        If grd.Name = "grdWHTLOCB2" Then

            Select Case e.Tool.Key
                Case "Reverse Entire Move (All Lines shown Below)"
                    Dim WHSE_TRAN_NO As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""
                    Reverse_Transaction(WHSE_TRAN_NO)

                Case "Reverse This Move Line Only"
                    Dim WHSE_TRAN_NO As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""
                    Dim WHSE_TRAN_LNO As Int32 = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_LNO").Value & ""
                    Reverse_Transaction(WHSE_TRAN_NO, WHSE_TRAN_LNO)
            End Select
        End If

        If grd.Name = "grdWHTINSTX" Then
            Select Case e.Tool.Key
                Case "Wave Inquiry"
                    Dim WAVE_NO As String = grd.ActiveRow.Cells("WAVE_NO").Value
                    Context_Launch("View", WAVE_NO, e.Tool.Key, "WHFWAVEI")
            End Select
        End If

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LOCATION_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select", e)
                End If
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select", e)
                End If
            Case "BAR_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "LOCATION_CODE"
                Click_Command("Select")
            Case "STYLE_CODE"
                Click_Command("Select")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                If rowICTWHSE1 Is Nothing Then
                    UltraExplorerBar1.Groups("Locked Locations").Visible = False
                Else
                    Fill_Records("WHTLOCML", Absx1.txtFor("WHSE_CODE").Text)
                    UltraExplorerBar1.Groups("Locked Locations").Visible = (dst.Tables("WHTLOCML").Rows.Count <> 0)
                End If

        End Select
    End Sub
#End Region

#Region "grdWHTLOCB1"
    Private Sub grdWHTLOCB1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTLOCB1.AfterRowActivate
        Setup_grdWHTLOCB2()
        Setup_grdWHTCONTC()
        Load_Waves()
        Load_WHTLOCBM()
    End Sub
#End Region

    Sub Setup_grdWHTLOCB1()


        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show 0 Qty"), UltraWinToolbars.StateButtonTool)
        Dim dvw As DataView = DirectCast(grdWHTLOCB1.DataSource, DataTable).DefaultView
        If Not tlb_sbt.Checked Then
            dvw.RowFilter = "LOCATION_QTY <> 0 OR PERSIST = '1'"
        Else
            For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("PERSIST = '1'")
                rowWHTLOCB1.Item("PERSIST") = "0"
            Next
            dvw.RowFilter = ""
        End If
    End Sub

    Sub Setup_grdWHTCONTC()
        ASCMAIN1.Progress("Loading Carton Contents", "")
        If grdWHTLOCB1.ActiveRow Is Nothing OrElse Not grdWHTLOCB1.ActiveRow.IsDataRow Then
            tabCycle.Tabs("Carton Contents").Visible = False
        Else
            Dim LOCATION_CODE As String = grdWHTLOCB1.ActiveRow.Cells("LOCATION_CODE").Value & ""
            Dim BAR_CODE As String = grdWHTLOCB1.ActiveRow.Cells("BAR_CODE").Value & ""

            ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1, WHTLOCM1 Where WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" _
                & " and WHTLOCB1.BAR_CODE = '" & BAR_CODE & "' And LOCATION_QTY <> 0 " _
                & " and  Nvl(LOCATION_USE,'X') Not in ('G','D','S')" _
                & " and WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE" _
                & " And WHTLOCM1.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
            Fill_Records("WHTCONTC", "", True, ASCMAIN1.sql)

            Sort_grdColumns(grdWHTCONTC, "STYLE_CODE")

            tabCycle.Tabs("Carton Contents").Visible = IIf(dst.Tables("WHTCONTC").Rows.Count = 1 Or dst.Tables("WHTCONTC").Rows.Count = 0, False, True)
        End If
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Setup_grdWHTLOCB2()
        ASCMAIN1.Progress("Loading Transaction History", "")
        If grdWHTLOCB1.ActiveRow Is Nothing OrElse Not grdWHTLOCB1.ActiveRow.IsDataRow Then
            grdWHTLOCB2.Visible = False
        Else
            grdWHTLOCB2.Visible = True
            Dim LOCATION_CODE As String = grdWHTLOCB1.ActiveRow.Cells("LOCATION_CODE").Value & ""
            Dim STYLE_CODE As String = grdWHTLOCB1.ActiveRow.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = grdWHTLOCB1.ActiveRow.Cells("COLOR_CODE").Value & ""
            Dim BAR_CODE As String = grdWHTLOCB1.ActiveRow.Cells("BAR_CODE").Value & ""
            Dim LOAD_NO As String = grdWHTLOCB1.ActiveRow.Cells("LOAD_NO").Value & ""
            Dim DVW As DataView = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView

            Dim SQL As String = ""
            Dim CAPTION As String = ""
            If chkMain.Checked Then
                SQL &= " and LOCATION_CODE = '" & LOCATION_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "'"
                If optViewBy.Value = "S" Then
                    CAPTION = "Style " & STYLE_CODE
                    CAPTION &= ", Location " & LOCATION_CODE
                Else
                    CAPTION = "Location " & LOCATION_CODE
                    CAPTION &= ", Style " & STYLE_CODE
                End If
            End If
            If chkCOLOR_CODE.Checked Then
                SQL &= " and COLOR_CODE = '" & COLOR_CODE & "'"
                CAPTION &= ", Color " & COLOR_CODE
            End If
            If chkBAR_CODE.Checked Then
                SQL &= " and BAR_CODE = '" & BAR_CODE & "'"
                CAPTION &= ", Case ID " & BAR_CODE
            End If
            If chkLOAD_NO.Checked Then
                'SQL &= " and LOAD_NO = '" & LOAD_NO & "'"
                'CAPTION &= ", Load " & LOAD_NO
            End If


            grdWHTLOCB2.Text = "Audit Trail for " & CAPTION
            'DVW.RowFilter = Mid(SQL, 5)
            ASCMAIN1.sql = "Select * from WHTLOCB2" & sqlWHTLOCB2where & SQL
            Fill_Records("WHTLOCB2", "", True, ASCMAIN1.sql)

            Sort_grdColumns(grdWHTLOCB2, "INIT_DATE")
        End If
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Setup_grdWHTMOVEX()
        If grdWHTLOCB2.ActiveRow Is Nothing Then
            grdWHTMOVEX.Visible = False
        Else
            grdWHTMOVEX.Visible = True
            Dim WHSE_TRAN_NO As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""
            Dim WHSE_TRAN_TYPE As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_TYPE").Value & ""
            If WHSE_TRAN_TYPE <> "M" Then
                grdWHTMOVEX.Visible = False
            Else
                Fill_Records("WHTMOVEX", WHSE_TRAN_NO)
                Sort_grdColumns(grdWHTMOVEX, "WHSE_TRAN_LNO")
                grdWHTMOVEX.Text = "Transaction Details for " & WHSE_TRAN_NO
            End If
        End If
    End Sub

    Sub Integrity_Check(ack_if_ok As Boolean, Optional compare_snapshots As Boolean = False)

        If Absx1.txtFor("WHSE_CODE").Text = "" Then Exit Sub

        If Not compare_snapshots Then
            Exit Sub
        End If

        'If (1 = 1 Or Absx1.txtFor("WHSE_CODE").Text = "") Then
        '    Exit Sub
        'End If

        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, LOC, STY, VAR, LOC - (STY + VAR) OUT FROM (" & vbCrLf _
            & " Select STYLE_CODE, COLOR_CODE, SUM (LOC) LOC, SUM (STY) STY, SUM (VAR) VAR FROM (" & vbCrLf _
            & " Select STYLE_CODE, COLOR_CODE, SUM (LOCATION_QTY) LOC, 0 STY, 0 VAR" & vbCrLf _
            & " from WHTLOCB1" & vbCrLf _
            & " group by STYLE_CODE, COLOR_CODE" & vbCrLf _
            & " Union" & vbCrLf _
            & " Select STYLE_CODE, COLOR_CODE, 0 LOC, SUM (WHSE_QTY_ON_HAND) STY, 0 VAR" & vbCrLf _
            & " from ICTSTAT2 WHERE WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" & vbCrLf _
            & " group by STYLE_CODE, COLOR_CODE" & vbCrLf _
            & " ) group by STYLE_CODE, COLOR_CODE" & vbCrLf _
            & " ) where LOC - (STY + VAR) <> 0"

        If compare_snapshots Then
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("Now taking Snapshot and comparing to previous snapshot")


            Dim OOBALXNO As String = ASCMAIN1.Next_Control_No("OOBALXNO")
            ASCDATA1.ExecuteSQL("Insert into FIND_OOBAL Select '" & OOBALXNO & "' XNO, SYSDATE SNAPSHOT_DATE, X.* FROM (" & ASCMAIN1.sql & ") X")

            ASCMAIN1.sql = "" _
                & "Select STYLE_CODE, COLOR_CODE, OUT FROM FIND_OOBAL WHERE XNO = '" & OOBALXNO & "'" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, OUT FROM FIND_OOBAL WHERE XNO = '" & Format(Val(OOBALXNO) - 1, "0000000000") & "'"
            Dim tbldiff As DataTable = ASCDATA1.GetDataTable

            If tbldiff.Rows.Count = 0 Then
                ASCDATA1.ExecuteSQL("Delete from FIND_OOBAL where XNO = '" & Format(Val(OOBALXNO) - 1, "0000000000") & "'")
                MsgBox("All OK", MsgBoxStyle.OkOnly, "Verification")
            Else
                Using F As New ASFMSGBF
                    F.Show_grd(tbldiff, Me, "Styles Out of Balance (Locator vs Perpetual) - snapshot difference")
                End Using
            End If

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

            Exit Sub


        End If

        Dim tbl As DataTable = ASCDATA1.GetDataTable
        If tbl.Rows.Count = 0 Then
            If ack_if_ok Then
                MsgBox("All OK", MsgBoxStyle.OkOnly, "Verification")
            End If
        Else
            Using F As New ASFMSGBF
                F.Show_grd(tbl, Me, "Styles Out of Balance (Locator vs Perpetual)")
            End Using
        End If


    End Sub

    Sub Setup_ViewBy()
        grpStyle.Visible = (optViewBy.Value = "S")
        grpLocation.Visible = Not (optViewBy.Value = "S")
        lblLOCATION_CODE2.Visible = (optViewBy.Value = "R")
        txtLOCATION_CODE2.Visible = (optViewBy.Value = "R")
    End Sub

    Private Sub optViewBy_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optViewBy.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_ViewBy()
    End Sub

    Private Sub grdICTSTYL1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTYL1.DoubleClickRow
        optViewBy.Value = "S"
        Absx1.txtFor("STYLE_CODE").Text = grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Text
        Click_Command("Select")
    End Sub

    Private Sub grdWHTLOCBJ_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTLOCBJ.DoubleClickRow
        optViewBy.Value = "S"
        Absx1.txtFor("STYLE_CODE").Text = grdWHTLOCBJ.ActiveRow.Cells("STYLE_CODE").Text
        Click_Command("Select")
    End Sub

    Private Sub grdWHTLOCM1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTLOCM1.DoubleClickRow
        If grdWHTLOCM1.ActiveRow IsNot Nothing AndAlso grdWHTLOCM1.ActiveRow.IsDataRow Then
            optViewBy.Value = "L"
            Absx1.txtFor("WHSE_CODE").Text = grdWHTLOCM1.ActiveRow.Cells("WHSE_CODE").Text
            Absx1.txtFor("LOCATION_CODE").Text = grdWHTLOCM1.ActiveRow.Cells("LOCATION_CODE").Text
            Click_Command("Select")
        End If
    End Sub

    Private Sub grdICTSTYLA_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTYLA.DoubleClickRow
        optViewBy.Value = "S"
        Absx1.txtFor("STYLE_CODE").Text = grdICTSTYLA.ActiveRow.Cells("STYLE_CODE").Text
        Click_Command("Select")
    End Sub

    Private Sub grdWHTLOCMA_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTLOCMA.DoubleClickRow
        optViewBy.Value = "L"
        Absx1.txtFor("WHSE_CODE").Text = grdWHTLOCMA.ActiveRow.Cells("WHSE_CODE").Text
        Absx1.txtFor("LOCATION_CODE").Text = grdWHTLOCMA.ActiveRow.Cells("LOCATION_CODE").Text
        Click_Command("Select")
    End Sub

    Private Sub grdWHTLOCB2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTLOCB2.AfterRowActivate
        Setup_grdWHTMOVEX()
    End Sub

    Private Sub grdWHTMOVEX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTMOVEX.InitializeRow
        If e.Row.Cells("WHSE_TRAN_LNO").Value = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_LNO").Value Then
            e.Row.Appearance.ForeColor = Drawing.Color.Blue
        End If
        If e.Row.Cells("STATUS").Value & "" = "R" Then
            e.Row.Cells("WHSE_TRAN_LNO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("WHSE_TRAN_LNO").ToolTipText = "This Line was Reversed"
        End If
    End Sub

    Sub Reverse_Transaction(WHSE_TRAN_NO As String, Optional WHSE_TRAN_LNO As Int32 = 0)
        If Not ASCMAIN1.Logical_Lock("WHTMOVE1", WHSE_TRAN_NO, , , , 1) Then Exit Sub

        Dim rowWHTMOVE1 As DataRow = Fill_Record("WHTMOVE1", WHSE_TRAN_NO)
        If rowWHTMOVE1.Item("STATUS") & "" = "R" Then
            MsgBox("Move Transaction " & WHSE_TRAN_NO & " has already been reversed",
                   MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        If MsgBox("Are you sure that you want to Reverse Move Transaction " _
                  & WHSE_TRAN_NO & IIf(WHSE_TRAN_LNO = 0, "", ", Line " & CStr(WHSE_TRAN_LNO)),
                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        BeginTrans()
        rowWHTMOVE1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        Update_Record_TDA("WHTMOVE1")
        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {WHSE_TRAN_NO, WHSE_TRAN_LNO, -1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})
        CommitTrans()

        Dim sqlw As String = " and WHSE_TRAN_TYPE = 'M' and WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'" _
                             & IIf(WHSE_TRAN_LNO = 0, "", " and WHSE_TRAN_LNO = " & CStr(WHSE_TRAN_LNO))
        For Each rowWHTLOCB2 As DataRow In dst.Tables("WHTLOCB2").Select(Mid(sqlw, 5))
            rowWHTLOCB2.Delete()
        Next

        ASCMAIN1.sql = "Select * from WHTLOCB2" & sqlWHTLOCB2where & sqlw _
                             & IIf(WHSE_TRAN_LNO = 0, "", " and WHSE_TRAN_LNO = " & CStr(WHSE_TRAN_LNO))

        Fill_Records("WHTLOCB2", "", False, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from WHTLOCB1" _
            & Replace(sqlWHTLOCB2where, "WHTLOCB2", "WHTLOCB1") _
            & " and (LOCATION_CODE,STYLE_CODE,COLOR_CODE) in" _
            & " (Select LOCATION_CODE,STYLE_CODE,COLOR_CODE from WHTLOCB2 " _
            & sqlWHTLOCB2where & sqlw _
            & IIf(WHSE_TRAN_LNO = 0, "", " and WHSE_TRAN_LNO = " & CStr(WHSE_TRAN_LNO)) & ")"

        If 1 <> 1 Then ' THE SECTION BELOW ERRORS OUT B/C WHTLOCB1 DOES NOT HAVE A KEY ANY MORE -ALSO WOULD NEED TO MAKE SURE CASE AND LOAD ARE CHECKED
            For Each row As DataRow In ASCDATA1.GetDataTable.Select
                Dim rowWHTLOCB1 As DataRow = dst.Tables("WHTLOCB1").Rows.Find(New String() {
                                                            row.Item("WHSE_CODE"),
                                                            row.Item("LOCATION_CODE"),
                                                            row.Item("BAR_CODE"),
                                                            row.Item("STYLE_CODE"),
                                                            row.Item("COLOR_CODE")})
                rowWHTLOCB1.Item("LOCATION_QTY") = row.Item("LOCATION_QTY")
                rowWHTLOCB1.Item("LAST_DATE") = row.Item("LAST_DATE")
                rowWHTLOCB1.Item("LAST_OPER") = row.Item("LAST_OPER")
                rowWHTLOCB1.Item("PERSIST") = "1"
            Next
        End If

        ' Setup_grdWHTLOCB1()
        Setup_grdWHTLOCB2()
        Setup_grdWHTMOVEX()

        ASCMAIN1.MultiTask_Release(, , 1)

        MsgBox("Move Transaction " & WHSE_TRAN_NO _
               & IIf(WHSE_TRAN_LNO = 0, "", ", Line " & CStr(WHSE_TRAN_LNO)) _
               & " has been Successfully Reversed", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub GCV_Fix()
        dst.Tables("WHTMOVE1").Rows.Clear()
        dst.Tables("WHTMOVE2").Rows.Clear()

        ASCMAIN1.sql = " Select X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE," & vbCrLf _
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
        & " And X.LOCATION_QTY = ABS(Y.LOCATION_QTY) "
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim Style_Code_Fx As String = row.Item("STYLE_CODE")
            Dim Color_Code_Fx As String = row.Item("COLOR_CODE")
            Dim Qty_fx As Integer = row.Item("POS_QTY")

            Dim WHSE_TRAN_NO As String = String.Empty
            If WHSE_TRAN_NO.Length = 0 Then
                WHSE_TRAN_NO = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

                Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
                rowWHTMOVE1.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                rowWHTMOVE1.Item("WHSE_TRAN_TYPE") = "M"
                rowWHTMOVE1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowWHTMOVE1.Item("WHSE_CODE") = WHSE_CODE
                rowWHTMOVE1.Item("STATUS") = "U"
                rowWHTMOVE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowWHTMOVE1.Item("INIT_DATE") = DATETIME_STAMP
                rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowWHTMOVE1.Item("LAST_DATE") = DATETIME_STAMP
                dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
            End If

            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
            With rowWHTMOVE2
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("WHSE_TRAN_LNO") = 1
                .Item("LOCATION_CODE_FROM") = row.Item("LOCATION_CODE")
                .Item("LOCATION_CODE_TO") = row.Item("LOCATION_CODE")
                If rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C" Then
                    .Item("BAR_CODE") = row.Item("POS_ID")
                    .Item("LOAD_NO_FROM") = row.Item("POS_LOAD")
                    .Item("LOAD_NO_TO") = ""
                    If row.Item("NEG_ID") & "" <> "" Then
                        .Item("BAR_CODE_OTHER") = row.Item("NEG_ID")
                    End If
                Else
                    .Item("BAR_CODE") = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")
                    .Item("LOAD_NO_FROM") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                    .Item("LOAD_NO_TO") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                End If
                .Item("WHSE_TRAN_QTY") = Qty_fx
                ' .Item("WHSE_TRAN_QTY_ORIG") = Qty_fx
                .Item("STYLE_CODE") = Style_Code_Fx
                .Item("COLOR_CODE") = Color_Code_Fx
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"
                ' .Item("ERROR_CODES") = String.Empty
                'LookUp("ICTSTYL1", Style_Code_Fx)
                'rowWHTMOVE2.Item("STYLE_DESC") = cdr.Item("STYLE_DESC") & String.Empty
            End With
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

        Next




        Dim r As DataRow = dst.Tables("WHTMOVE1")(0)
        Update_Record_TDA("WHTMOVE1")
        Update_Record_TDA("WHTMOVE2")
        For Each rowWHTMOVE1 As DataRow In dst.Tables("WHTMOVE1").Rows
            ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {rowWHTMOVE1.Item("WHSE_TRAN_NO"), 0, 1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})
        Next

    End Sub

    Sub Move_To(Optional LOCATION_CODE_TO As String = "", Optional movement_type As String = "") ' Optional confirm_only As Boolean = False) ' Sub Move_To(Optional move_to_bin As Boolean = False)

        Dim confirm_only As Boolean = (movement_type = "LNF") Or (movement_type = "CMB")
        Dim dupBarcode As String = ""

        Using ff As New TAC.TAFLOCM1()
            Dim grd As Infragistics.Win.UltraWinGrid.UltraGrid
            Dim tblName As String
            If SOURCE_GRID = "grdWHTLOCB1" Then
                grd = grdWHTLOCB1
                tblName = "WHTLOCB1"
            Else
                grd = grdWHTLOCBM
                tblName = "WHTLOCBM"
            End If

            'Dim LOCATION_QTY_CMB As Int64 = 0
            Dim BAR_CODE_CMB As String = ""
            If movement_type = "CMB" Then
                BAR_CODE_CMB = grd.Selected.Rows(grd.Selected.Rows.Count - 1).Cells("BAR_CODE").Value
            End If

            If movement_type = "CCS" Then
                grd.Selected.Rows.Clear()
            End If

            ff.confirm_only = confirm_only
            ff.movement_type = movement_type
            ff.rowICTWHSE1 = rowICTWHSE1
            ff.WHSE_CODE = WHSE_CODE



            If grd.Selected.Rows.Count = 0 And movement_type <> "CCS" Then
                MsgBox("Nothing to Move", MsgBoxStyle.OkOnly, "Please Select Lines to move")
                Exit Sub
            End If

            Dim BCs As New List(Of String)
            Dim LOCs As New List(Of String)

            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows

                If movement_type = "CMB" Then
                    ff.BAR_CODE_CMB = BAR_CODE_CMB
                End If

                Dim QTY As Int64 = 0
                If confirm_only Or rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C" Then
                    QTY = Val(row.Cells("LOCATION_QTY").Value & "")
                End If

                If rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C" Then

                    If movement_type = "CFG" Then
                        BCs.Add(row.Cells("BAR_CODE").Value)

                        If Not LOCs.Contains(row.Cells("LOCATION_CODE").Value) Then
                            LOCs.Add(row.Cells("LOCATION_CODE").Value)
                        End If
                    End If

                    'Fix duplicate when barcode selected more than once it duplicates entries
                    If movement_type = "" And dupBarcode.Contains(row.Cells("BAR_CODE").Value) Then
                        'skip dup rec
                    Else
                        dupBarcode += "," & row.Cells("BAR_CODE").Value

                        Dim sqlw As String = "WHSE_CODE = '" & WHSE_CODE & "' and BAR_CODE = '" & row.Cells("BAR_CODE").Value & "' and LOCATION_QTY <> 0 and BAR_CODE <> '" & rowICTWHSE1.Item("WHSE_DEF_BAR_CODE") & "'"
                        If movement_type = "TRN" Then
                            sqlw = "WHSE_CODE = '" & WHSE_CODE & "' and BAR_CODE = '" & row.Cells("BAR_CODE").Value & "'" _
                            & " And LOCATION_QTY <> 0 And BAR_CODE <> '" & rowICTWHSE1.Item("WHSE_DEF_BAR_CODE") & "'" _
                            & " And STYLE_CODE = '" & row.Cells("STYLE_CODE").Value & "'" _
                            & " And COLOR_CODE = '" & row.Cells("COLOR_CODE").Value & "'"
                        End If

                        If movement_type = "CHG" Then
                            sqlw = "WHSE_CODE = '" & WHSE_CODE & "' and LOCATION_CODE = '" & row.Cells("LOCATION_CODE").Value & "' and BAR_CODE = '" & row.Cells("BAR_CODE").Value & "' and STYLE_CODE = '" & row.Cells("STYLE_CODE").Value & "' and COLOR_CODE = '" & row.Cells("COLOR_CODE").Value & "' and LOCATION_QTY > 0"
                        ElseIf movement_type = "BTS" Then
                            sqlw = "WHSE_CODE = '" & WHSE_CODE & "' and LOCATION_CODE = '" & row.Cells("LOCATION_CODE").Value & "' and BAR_CODE = '" & row.Cells("BAR_CODE").Value & "' and STYLE_CODE = '" & row.Cells("STYLE_CODE").Value & "' and COLOR_CODE = '" & row.Cells("COLOR_CODE").Value & "' and LOCATION_QTY > 0"
                        ElseIf movement_type = "CFG" Then ' CTN TOO???
                            sqlw &= " and LOCATION_CODE = '" & row.Cells("LOCATION_CODE").Value & "'"

                        ElseIf movement_type = "CMB" Then
                            sqlw = "WHSE_CODE = '" & WHSE_CODE & "' and LOCATION_CODE = '" & row.Cells("LOCATION_CODE").Value & "' and BAR_CODE = '" & row.Cells("BAR_CODE").Value & "' and STYLE_CODE = '" & row.Cells("STYLE_CODE").Value & "' and COLOR_CODE = '" & row.Cells("COLOR_CODE").Value & "'"
                        End If
                        For Each rowWHTLOCB1 As DataRow In dst.Tables(tblName).Select(sqlw)
                            If Val(rowWHTLOCB1.Item("LOCATION_QTY_WAVE") & "") <> 0 Then
                                MsgBox("Cannot Change or Move a Case which has been committed to a Wave", MsgBoxStyle.OkOnly, "Cannot Move")
                                Exit Sub
                            End If

                            Dim LOCATION_QTY As Int64 = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "")
                            If movement_type = "CMB" Then
                                If BAR_CODE_CMB = rowWHTLOCB1.Item("BAR_CODE") Then
                                    Exit For
                                End If
                            End If

                            ff.AddItemToMove(rowWHTLOCB1.Item("WHSE_CODE"),
                                 rowWHTLOCB1.Item("LOCATION_CODE"),
                                 rowWHTLOCB1.Item("STYLE_CODE"),
                                 rowWHTLOCB1.Item("COLOR_CODE"),
                                 rowWHTLOCB1.Item("BAR_CODE"),
                                 rowWHTLOCB1.Item("LOAD_NO") & "",
                                 LOCATION_QTY,
                                 LOCATION_CODE_TO, BAR_CODE_CMB)
                        Next
                    End If
                Else
                    ff.AddItemToMove(row.Cells("WHSE_CODE").Value,
                         row.Cells("LOCATION_CODE").Value,
                         row.Cells("STYLE_CODE").Value,
                         row.Cells("COLOR_CODE").Value,
                         row.Cells("BAR_CODE").Value,
                         row.Cells("LOAD_NO").Value & "",
                         QTY,
                         LOCATION_CODE_TO)
                End If
            Next

            If movement_type = "CFG" Then
                If LOCs.Count > 1 Then
                    MsgBox("Cannot Select LPNs from Different Locations when Re-Configuring", MsgBoxStyle.OkOnly, "Cannot Re-Configure")
                    Exit Sub
                End If
                Dim QTY_BASE As Integer = -1
                ASCMAIN1.sql = "Select BAR_CODE, SUM (LOCATION_QTY) QTY, COUNT (*) SCS" & vbCrLf _
                    & " from WHTLOCB1" & vbCrLf _
                    & " where WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                    & "   and LOCATION_CODE = '" & LOCs(0) & "'" & vbCrLf _
                    & "   and BAR_CODE in ('" & Join(BCs.ToArray, "','") & "')" & vbCrLf _
                    & " group by BAR_CODE"
                For Each rowcheck As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim BAR_CODE As String = rowcheck.Item("BAR_CODE")
                    If Val(rowcheck.Item("SCS") & "") <> 1 Then
                        MsgBox("Cannot Select LPNs with Mixed Content (See " & BAR_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Re-Configure")
                        Exit Sub
                    End If
                    If Val(rowcheck.Item("QTY") & "") <= 0 Then
                        MsgBox("Cannot Select LPNs with Qty <= 0 (See " & BAR_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Re-Configure")
                        Exit Sub
                    End If
                    If QTY_BASE = -1 Then
                        QTY_BASE = Val(rowcheck.Item("QTY") & "")
                    Else
                        If Val(rowcheck.Item("QTY") & "") <> QTY_BASE Then
                            MsgBox("Cannot Select LPNs with Mixed Content (See " & BAR_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Re-Configure")
                            Exit Sub
                        End If
                    End If

                Next
            End If

            'row.Cells("LOCATION_QTY").Value
            ff.ShowDialog()

            Dim WHSE_TRAN_NO As String = ff.WHSE_TRAN_NO
            If WHSE_TRAN_NO <> "" Then

                ASCMAIN1.sql = "Select * from WHTLOCB2" _
                    & sqlWHTLOCB2where _
                    & " and WHSE_TRAN_TYPE = 'M' and WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'"

                Fill_Records("WHTLOCB2", "", False, ASCMAIN1.sql)
                If SOURCE_GRID = "grdWHTLOCB1" Then
                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTLOCB1.Rows ' grdWHTLOCB1.Selected.Rows

                        Dim rowWHTLOCB1 As DataRow = LookUp("WHTLOCB1", New String() {
                                                            row.Cells("WHSE_CODE").Value,
                                                            row.Cells("LOCATION_CODE").Value,
                                                            row.Cells("BAR_CODE").Value,
                                                            row.Cells("STYLE_CODE").Value,
                                                            row.Cells("COLOR_CODE").Value})
                        row.Cells("LOCATION_QTY").Value = rowWHTLOCB1.Item("LOCATION_QTY")
                        row.Cells("LAST_DATE").Value = rowWHTLOCB1.Item("LAST_DATE")
                        row.Cells("LAST_OPER").Value = rowWHTLOCB1.Item("LAST_OPER")
                        row.Cells("PERSIST").Value = "1"
                        row.Update()
                    Next
                End If
                ASCMAIN1.Progress("Refreshing Data", "")
                Load_WHTLOCB1()
            End If

            ' Setup_grdWHTLOCB1()

            Setup_grdWHTLOCB2()
            Setup_grdWHTMOVEX()
            ASCMAIN1.Progress("", "")
        End Using
    End Sub

    Sub Toggle_grpDetail()

        If SELECTION_NO = 0 Then Exit Sub
        grdWHTLOCB1.DisplayLayout.Bands(0).Columns("BAR_CODE").Hidden = Not chkBAR_CODE.Checked
        grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOAD_NO").Hidden = Not chkLOAD_NO.Checked
        grdWHTLOCB1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = Not chkCOLOR_CODE.Checked
        If optViewBy.Value = "S" Then
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOCATION_CODE").Hidden = Not chkMain.Checked
        Else
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = Not chkMain.Checked
        End If

        If ScreenMode Then
            ASCMAIN1.Progress("Now Loading Display")
            Load_WHTLOCB1()
            ASCMAIN1.Progress("")
        End If
    End Sub

    Private Sub chkBAR_CODE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBAR_CODE.CheckedChanged
        Toggle_grpDetail()
    End Sub

    Private Sub chkLOAD_NO_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkLOAD_NO.CheckedChanged
        Toggle_grpDetail()
    End Sub

    Private Sub chkMain_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMain.CheckedChanged
        Toggle_grpDetail()
    End Sub

    Private Sub chkCOLOR_CODE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkCOLOR_CODE.CheckedChanged
        Toggle_grpDetail()
    End Sub

    Sub Load_Waves()
        If SELECTION_NO = 0 Then Exit Sub
        If tabCycle.ActiveTab.Text = "Wave Details" Then
            If grdWHTLOCB1.ActiveRow IsNot Nothing AndAlso grdWHTLOCB1.ActiveRow.IsDataRow Then
                ASCMAIN1.sql = sqlWHTINSTX & " And WHTWAVE1.WHSE_CODE = '" & WHSE_CODE & "'" &
                IIf(optViewBy.Value = "L",
                    " and WHTINST1.LOCATION_CODE = '" & LOCATION_CODE & "' and WHTINST2.STYLE_CODE = '" & grdWHTLOCB1.ActiveRow.Cells("STYLE_CODE").Text & "'",
                    " and WHTINST2.STYLE_CODE = '" & STYLE_CODE & "'")
                Fill_Records("WHTINSTX", "", True, ASCMAIN1.sql)
            End If
        End If
        ' Stop
    End Sub

    Sub Load_WHTLOCBM()
        If SELECTION_NO = 0 Or tabCycle.ActiveTab.Text <> "Barcodes" Then Exit Sub
        If grdWHTLOCB1.ActiveRow IsNot Nothing AndAlso grdWHTLOCB1.ActiveRow.IsDataRow Then
            ASCMAIN1.Progress("Loading Barcode Details", "")

            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show 0-Qty"), UltraWinToolbars.StateButtonTool)


            ASCMAIN1.sql = "Select WHTLOCB1.*,WHTBARC1.LOAD_NO" & vbCrLf _
                & ", WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
                & " from WHTLOCB1,WHTLOCM1,WHTBARC1" & vbCrLf _
                & " Where WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
                & " and WHTLOCM1.LOCATION_CODE (+) = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & " and WHTLOCM1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE" & vbCrLf _
                & " and WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                & IIf(Not tlb_sbt.Checked, " and WHTLOCB1.LOCATION_QTY <> 0", "") & vbCrLf _
                & " and WHTLOCB1.BAR_CODE in ( SELECT BAR_CODE FROM WHTLOCB1" & vbCrLf _
                & " Where WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                & " and WHTLOCB1.STYLE_CODE = '" & grdWHTLOCB1.ActiveRow.Cells("STYLE_CODE").Text & "'" & vbCrLf _
                & " and WHTLOCB1.COLOR_CODE = '" & grdWHTLOCB1.ActiveRow.Cells("COLOR_CODE").Text & "'" & vbCrLf _
                & " and WHTLOCB1.LOCATION_QTY <> 0 )"
            Fill_Records("WHTLOCBM", "", True, ASCMAIN1.sql)

            ASCMAIN1.Progress("", "")
        End If
        ' Stop
    End Sub

    Sub Load_WHTLOCB1()

        Dim use_aggregation As Boolean = rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C" _
                And (Not chkMain.Checked Or Not chkCOLOR_CODE.Checked Or Not chkBAR_CODE.Checked Or Not chkLOAD_NO.Checked)

        If use_aggregation Then

            ' WE ARE DOING AGGREGATION

            ASCMAIN1.sql = "Select WHTLOCB1.WHSE_CODE" & vbCrLf _
                & IIf(chkMain.Checked Or optViewBy.Value = "L", ",WHTLOCB1.LOCATION_CODE", ",'X' LOCATION_CODE") & vbCrLf _
                & IIf(chkBAR_CODE.Checked, ",WHTLOCB1.BAR_CODE", ",'X' BAR_CODE") & vbCrLf _
                & IIf(chkMain.Checked Or optViewBy.Value = "S", ",WHTLOCB1.STYLE_CODE", ",'X' STYLE_CODE") & vbCrLf _
                & IIf(chkCOLOR_CODE.Checked, ",WHTLOCB1.COLOR_CODE", ",'X' COLOR_CODE") & vbCrLf _
                & ", SUM (LOCATION_QTY) LOCATION_QTY" & vbCrLf _
                & ", NULL INIT_DATE, NULL INIT_OPER, NULL LAST_DATE, NULL LAST_OPER" & vbCrLf _
                & ", SUM (LOCATION_QTY_WAVE) LOCATION_QTY_WAVE" & vbCrLf _
                & IIf(chkLOAD_NO.Checked, ",WHTBARC1.LOAD_NO", ",'X' LOAD_NO") & vbCrLf _
                & IIf(chkMain.Checked Or optViewBy.Value = "L", ",WHTLOCM1.LOCATION_LOCKED", ",'X' LOCATION_LOCKED") & vbCrLf _
                & " from WHTLOCB1,WHTLOCM1,WHTBARC1" & Replace(sqlWHTLOCB2where, "WHTLOCB2.", "WHTLOCB1.") & vbCrLf _
                & " and WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
                & " and WHTLOCM1.LOCATION_CODE (+) = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & " and WHTLOCM1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE" & vbCrLf _
                & IIf(LOCATION_CODE = "00-REC-A" Or LOCATION_CODE = "00-REC-B", " And LOCATION_QTY <> 0", "") & vbCrLf _
                & " group by WHTLOCB1.WHSE_CODE" & vbCrLf _
                & IIf(chkMain.Checked Or optViewBy.Value = "L", ",WHTLOCB1.LOCATION_CODE", "") & vbCrLf _
                & IIf(chkBAR_CODE.Checked, ",WHTLOCB1.BAR_CODE", "") & vbCrLf _
                & IIf(chkMain.Checked Or optViewBy.Value = "S", ",WHTLOCB1.STYLE_CODE", "") & vbCrLf _
                & IIf(chkCOLOR_CODE.Checked, ",WHTLOCB1.COLOR_CODE", "") & vbCrLf _
                & IIf(chkLOAD_NO.Checked, ",WHTBARC1.LOAD_NO", "") & vbCrLf _
                & IIf(chkMain.Checked Or optViewBy.Value = "L", ",WHTLOCM1.LOCATION_LOCKED", "") & vbCrLf

        Else
            ' WE ARE GETTIMG EVERYTHING
            ASCMAIN1.sql = "Select WHTLOCB1.*,WHTBARC1.LOAD_NO" & vbCrLf _
                & ", WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
                & " from WHTLOCB1,WHTLOCM1,WHTBARC1" & Replace(sqlWHTLOCB2where, "WHTLOCB2.", "WHTLOCB1.") & vbCrLf _
                & " and WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
                & " and WHTLOCM1.LOCATION_CODE (+) = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & " and WHTLOCM1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE" & vbCrLf _
                & IIf(LOCATION_CODE = "00-REC-A" Or LOCATION_CODE = "00-REC-B", " And LOCATION_QTY <> 0", "")


            'ASCMAIN1.sql = "Select WHTLOCB1.*,WHTBARC1.LOAD_NO" & vbCrLf _
            '& " from WHTLOCB1,WHTBARC1 Where WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
            '& "  And WHTBARC1.BAR_CODE in (Select * from (" & vbCrLf _
            '& " 	Select BAR_CODE from WHTLOCB1_TST " & vbCrLf _
            '& " 	Union" & vbCrLf _
            '& " 	Select BAR_CODE_OTHER from WHTLOCB1_TST " & vbCrLf _
            '& " 	) Where BAR_CODE <> '00000000') And LOCATION_QTY <> 0"

        End If

        With grdWHTLOCB1.DisplayLayout.Bands(0)
            .Columns("INIT_DATE").Hidden = use_aggregation
            .Columns("INIT_OPER").Hidden = use_aggregation
            .Columns("LAST_DATE").Hidden = use_aggregation
            .Columns("LAST_OPER").Hidden = use_aggregation
        End With

        With dst.Tables("WHTLOCB1")
            If Not chkMain.Checked Or Not chkCOLOR_CODE.Checked Or Not chkBAR_CODE.Checked Then
                .PrimaryKey = Nothing
            Else
                If .PrimaryKey Is Nothing Then
                    .PrimaryKey = New DataColumn() {.Columns("WHSE_CODE"), .Columns("LOCATION_CODE"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE"), .Columns("BAR_CODE")}
                End If
            End If
        End With

        Fill_Records("WHTLOCB1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdWHTLOCB1, "LOCATION_CODE,STYLE_CODE,COLOR_CODE")
        Setup_grdWHTLOCB1()


        'ASCMAIN1.sql = "Select * from WHTLOCB2" & sqlWHTLOCB2where
        'If LOCATION_CODE = "00-REC-A" Then
        '    ASCMAIN1.sql &= " and BAR_CODE in (" _
        '    & " Select BAR_CODE" & vbCrLf _
        '    & " from WHTLOCB1" & Replace(sqlWHTLOCB2where, "WHTLOCB2.", "WHTLOCB1.") & vbCrLf _
        '    & " And LOCATION_QTY <> 0)"
        'End If

        'Fill_Records("WHTLOCB2", "", True, ASCMAIN1.sql)

        dst.Tables("WHTLOCBD").Rows.Clear()
        dst.Tables("WHTLOCBC").Rows.Clear()
        Dim CASE_PACK_NO As Integer = 0

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else
            Exit Sub
        End If


        ASCMAIN1.sql = "Select BAR_CODE, LOCATION_CODE," & vbCrLf _
            & " ltrim(sys_connect_by_path(SC ,','),',') STYLE_COLOR_QTY_DNA" & vbCrLf _
            & " from" & vbCrLf _
            & " (select WHTLOCB1.BAR_CODE, WHTLOCB1.LOCATION_CODE, WHTLOCB1.STYLE_CODE || WHTLOCB1.COLOR_CODE || TO_CHAR(WHTLOCB1.LOCATION_QTY,'9999999') SC," & vbCrLf _
            & "    row_number() over(partition by WHTLOCB1.BAR_CODE, WHTLOCB1.LOCATION_CODE  order by WHTLOCB1.STYLE_CODE || WHTLOCB1.COLOR_CODE || TO_CHAR(WHTLOCB1.LOCATION_QTY,'9999999') ) rn," & vbCrLf _
            & "       row_number() over(partition by WHTLOCB1.BAR_CODE, WHTLOCB1.LOCATION_CODE order by WHTLOCB1.STYLE_CODE || WHTLOCB1.COLOR_CODE || TO_CHAR(WHTLOCB1.LOCATION_QTY,'9999999') desc)" & vbCrLf _
            & " rn_desc" & vbCrLf _
            & " from WHTLOCB1 where BAR_CODE in" & vbCrLf _
            & "    (Select Distinct WHTLOCB1.BAR_CODE from WHTLOCB1,WHTBARC1" & Replace(sqlWHTLOCB2where, "WHTLOCB2.", "WHTLOCB1.") & vbCrLf _
            & "      and WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
            & "      and WHTLOCB1.BAR_CODE <> '" & rowICTWHSE1.Item("WHSE_DEF_BAR_CODE") & "'" & vbCrLf _
            & "      and WHTLOCB1.LOCATION_QTY <> 0)" & vbCrLf _
            & "      and WHTLOCB1.LOCATION_CODE NOT in (Select LOCATION_CODE from WHTLOCM1 where LOCATION_USE in ('G','D','S') or LOCATION_CODE in ('00-REC-A','00-REC-B'))" & vbCrLf _
            & "     and WHTLOCB1.LOCATION_QTY <> 0" & vbCrLf _
            & IIf(optViewBy.Value = "L",
                  "     and WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf,
                  "") _
            & ")" & vbCrLf _
            & "    Where rn_desc = 1" & vbCrLf _
            & "  start with rn = 1" & vbCrLf _
            & "  connect by prior BAR_CODE = BAR_CODE" & vbCrLf _
            & "  and prior rn = rn-1"
        ASCMAIN1.sql = "Select STYLE_COLOR_QTY_DNA, COUNT (*) CASES from (" & ASCMAIN1.sql & ") group by STYLE_COLOR_QTY_DNA"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_COLOR_QTY_DNA")
            Dim rowWHTLOCBC As DataRow = dst.Tables("WHTLOCBC").NewRow
            CASE_PACK_NO += 1
            rowWHTLOCBC.Item("CASE_PACK_NO") = CASE_PACK_NO
            rowWHTLOCBC.Item("CASES") = row.Item("CASES")
            dst.Tables("WHTLOCBC").Rows.Add(rowWHTLOCBC)

            Dim STYLE_COLOR_QTY_DNA As String = row.Item("STYLE_COLOR_QTY_DNA")
            For Each SCQ As String In Split(STYLE_COLOR_QTY_DNA, ",")
                Dim QTY As Int64 = Val(Mid(SCQ, Len(SCQ) - 7, 8))
                Dim COLOR_CODE As String = Mid(SCQ, Len(SCQ) - 8 - 3 + 1, 3)
                Dim STYLE_CODE As String = Mid(SCQ, 1, Len(SCQ) - 8 - 3)
                Dim rowWHTLOCBD As DataRow = dst.Tables("WHTLOCBD").Rows.Find(New Object() {CASE_PACK_NO, STYLE_CODE, COLOR_CODE})
                If rowWHTLOCBD IsNot Nothing Then
                    rowWHTLOCBD.Item("QTY") = Val(rowWHTLOCBD.Item("QTY") & "") + QTY
                Else
                    dst.Tables("WHTLOCBD").Rows.Add(New Object() {CASE_PACK_NO, STYLE_CODE, COLOR_CODE, QTY})
                End If
            Next
        Next
    End Sub

    Private Sub grdWHTLOCMM_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTLOCMM.AfterRowActivate
        If grdWHTLOCMM.ActiveRow Is Nothing Then Exit Sub
        Fill_WHTCYCL1()
        Fill_WHTCYCL2()
    End Sub

    Private Sub grdWHTLOCMM_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTLOCMM.DoubleClickRow
        If grdWHTLOCMM.ActiveRow IsNot Nothing AndAlso grdWHTLOCMM.ActiveRow.IsDataRow Then
            optViewBy.Value = "L"
            Absx1.txtFor("WHSE_CODE").Text = grdWHTLOCMM.ActiveRow.Cells("WHSE_CODE").Text
            Absx1.txtFor("LOCATION_CODE").Text = grdWHTLOCMM.ActiveRow.Cells("LOCATION_CODE").Text
            Click_Command("Select")
        End If
    End Sub

    Sub Toggle_Case_Packs()

    End Sub

    Private Sub grdWHTLOCB1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTLOCB1.InitializeRow

        If e.Row.IsDataRow Then
            If Absx1.txtFor("BAR_CODE").Text <> "" Then
                If e.Row.Cells("BAR_CODE").Value & "" = Absx1.txtFor("BAR_CODE").Text Then
                    e.Row.Appearance.BackColor = Drawing.Color.Yellow
                End If
            End If

            If e.Row.Cells("LOCATION_LOCKED").Value & "" = "1" Then
                e.Row.Cells("LOCATION_CODE").Appearance.ForeColor = Drawing.Color.Magenta
                e.Row.Cells("LOCATION_CODE").ToolTipText = "Location Locked"
            Else
                e.Row.Cells("LOCATION_CODE").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("LOCATION_CODE").ToolTipText = ""
            End If
        End If
    End Sub

    Private Sub grdWHTINSTX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTINSTX.InitializeRow
        If e.Row.Band.Index = 0 Then
            With e.Row.Cells("WAVE_PICK_TYPE")
                If .Value = "L" Then
                    .Appearance.ForeColor = Drawing.Color.Blue
                ElseIf .Value = "C" Then
                    .Appearance.ForeColor = Drawing.Color.Green
                ElseIf .Value = "U" Then
                    .Appearance.ForeColor = Drawing.Color.Red
                End If
            End With
        End If
    End Sub

    Sub Toggle_Activity(tf As Boolean)
        With grdWHTLOCMM.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"CASES", "BAR_CODE", "QTY_LOCS", "QTY_WAVE", "LAST_DATE"}
                .Columns(COLUMN_NAME).Hidden = Not tf
            Next
        End With
    End Sub

    Sub GetActivity()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Scanning Locator Status & Activity")

        Toggle_Activity(True)

        Dim LOCATION_CODEs As New List(Of String)
        ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE" & vbCrLf _
            & ", COUNT (*) CASES, MIN (BAR_CODE) BAR_CODE" & vbCrLf _
            & ", SUM (QTY_LOCS) QTY_LOCS" & vbCrLf _
            & ", SUM (QTY_WAVE) QTY_WAVE" & vbCrLf _
            & ", MAX (LAST_DATE) LAST_DATE" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select WHSE_CODE, LOCATION_CODE, BAR_CODE" & vbCrLf _
            & ", SUM (LOCATION_QTY) QTY_LOCS" & vbCrLf _
            & ", SUM (LOCATION_QTY_WAVE) QTY_WAVE" & vbCrLf _
            & ", MAX (LAST_DATE) LAST_DATE" & vbCrLf _
            & " from WHTLOCB1" & vbCrLf _
            & " where (LOCATION_QTY > 0 OR LOCATION_QTY_WAVE > 0)" & vbCrLf _
            & "   and WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" & vbCrLf _
            & " group by WHSE_CODE, LOCATION_CODE, BAR_CODE" & vbCrLf _
            & ") group by WHSE_CODE, LOCATION_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowWHTLOCM1 As DataRow = dst.Tables("WHTLOCMM").Rows.Find _
                                         (New String() {row.Item("WHSE_CODE"), row.Item("LOCATION_CODE")})
            If rowWHTLOCM1 Is Nothing Then
                LOCATION_CODEs.Add(row.Item("LOCATION_CODE"))
            Else
                With rowWHTLOCM1
                    .Item("CASES") = row.Item("CASES")
                    .Item("BAR_CODE") = row.Item("BAR_CODE")
                    .Item("QTY_LOCS") = row.Item("QTY_LOCS")
                    .Item("QTY_WAVE") = row.Item("QTY_WAVE")
                    .Item("LAST_DATE") = row.Item("LAST_DATE")
                End With
            End If
        Next
        grdWHTLOCMM.Text &= " - Details as of " & Now.ToString

        If LOCATION_CODEs.Count <> 0 Then
            MsgBox("The following Locations are Orphans - please call ABS" _
                   & vbCrLf & Join(LOCATION_CODEs.ToArray, ","), _
                   MsgBoxStyle.OkOnly, "Found Orphan Locations")
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Void_and_Replace(BAR_CODE As String)
        'BeginTrans()

        ' PROMPT FOR NEW BARCODE
        ' VERIFY THAT IT DOES NOT EXISTS
        ' VERIFY THAT BC X -> BC Y
        ' UPDATE WRITE BARCODE TO WHTBARC1, AND FIX OTHER TABLES
        If Not ASCMAIN1.Logical_Lock("WHTBARC1", BAR_CODE) Then
            MsgBox("Could not lock access to Case Id " & BAR_CODE)
            Exit Sub
        End If

        Dim BAR_CODE_NEW As String = ASCMAIN1.Get_txt_from_User("Enter New Case ID", "Void/Replace Case ID", , 8)
        If BAR_CODE_NEW = "" Then Exit Sub

        ASCMAIN1.sql = "Select BAR_CODE from WHTBARC1" & vbCrLf _
                            & " where BAR_CODE = '" & BAR_CODE_NEW & "'"

        Dim row() As DataRow = ASCDATA1.GetDataTable.Select("")
        If row.Length = 0 Then
            Dim style = MsgBoxStyle.OkCancel Or MsgBoxStyle.DefaultButton2 Or _
            MsgBoxStyle.Exclamation
            Dim response = MsgBox("Replace Case ID " & BAR_CODE & " with new ID " & BAR_CODE_NEW, style, "Verify Update")
            If response = MsgBoxResult.Ok Then
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & "insert into whtbarc9 " _
                    & "   (BAR_CODE_NEW, BAR_CODE, TRAN_TYPE, TRAN_NO, PO_DATE_RECEIVED, INIT_DATE, " _
                    & "   LAST_DATE, INIT_OPER, LAST_OPER, QTY_TO_PRINT, QTY_PRINTED, " _
                    & "   PRINT_COUNT, STATUS_CODE, PO_ORDER_NO, PO_SHIPMENT_NO, " _
                    & "   PO_SHIPMENT_LNO, BC_COMMENT, CARTON_NO, LOAD_NO) " _
                    & "select '" & BAR_CODE_NEW & "' BAR_CODE_NEW, " _
                     & "  BAR_CODE, TRAN_TYPE, TRAN_NO, PO_DATE_RECEIVED, INIT_DATE, " _
                    & "   LAST_DATE, INIT_OPER, LAST_OPER, QTY_TO_PRINT, QTY_PRINTED, " _
                    & "   PRINT_COUNT, STATUS_CODE, PO_ORDER_NO, PO_SHIPMENT_NO, " _
                    & "   PO_SHIPMENT_LNO, BC_COMMENT, CARTON_NO, LOAD_NO" _
                    & " from whtbarc1 b where bar_code = '" & BAR_CODE & "';" & vbCrLf _
                    & "update whtbarc1 set bar_code = '" & BAR_CODE_NEW & "' where bar_code = '" & BAR_CODE & "';" & vbCrLf _
                    & "update whtinst2 set bar_code = '" & BAR_CODE_NEW & "' where bar_code = '" & BAR_CODE & "';" & vbCrLf _
                    & "update whtlocb1 set bar_code = '" & BAR_CODE_NEW & "' where bar_code = '" & BAR_CODE & "';" & vbCrLf _
                    & "update whtlocb2 set bar_code = '" & BAR_CODE_NEW & "' where bar_code = '" & BAR_CODE & "';" & vbCrLf _
                    & "update whtmove2 set bar_code = '" & BAR_CODE_NEW & "' where bar_code = '" & BAR_CODE & "';" & vbCrLf _
                    & "end;" & vbCrLf

                ASCDATA1.ExecuteSQL()

                MsgBox("Case ID " & BAR_CODE & " has been Replaced by Case ID " & BAR_CODE_NEW)
                Toggle_grpDetail()
                ASCMAIN1.sql = "Select * from WHTLOCB2" & sqlWHTLOCB2where
                Fill_Records("WHTLOCB2", "", True, ASCMAIN1.sql)
            Else
                MsgBox("No update done!", MsgBoxStyle.Exclamation)
            End If
        Else
            MsgBox("Case ID " & BAR_CODE_NEW & "is in use, try a different ID ", MsgBoxStyle.Critical)

        End If

        'CommitTrans()

    End Sub

    Private Sub grdWHTLOCML_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTLOCML.DoubleClickRow
        If grdWHTLOCML.ActiveRow IsNot Nothing AndAlso grdWHTLOCML.ActiveRow.IsDataRow Then
            optViewBy.Value = "L"
            Absx1.txtFor("WHSE_CODE").Text = grdWHTLOCML.ActiveRow.Cells("WHSE_CODE").Text
            Absx1.txtFor("LOCATION_CODE").Text = grdWHTLOCML.ActiveRow.Cells("LOCATION_CODE").Text
            Click_Command("Select")
        End If
    End Sub
    Sub Setup_Reconciliation()

        Dim sql As String = "Select WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", SUM (NVL(INST,0)) INST, SUM (NVL(LOCB,0)) LOCB" & vbCrLf _
            & ", SUM (NVL(INST,0)-NVL(LOCB,0)) DIFF" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select WHTWAVE1.WHSE_CODE, WHTINST1.LOCATION_CODE, WHTINST2.BAR_CODE, WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE" & vbCrLf _
            & ", SUM (LOCATION_QTY_WAVE) INST, 0 LOCB" & vbCrLf _
            & " from WHTINST1,WHTINST2,WHTWAVE1" & vbCrLf _
            & " where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
            & "   and WHTINST1.WAVE_INST_STATUS = '0'" & vbCrLf _
            & "   and WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
            & "   and WHTWAVE1.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" & vbCrLf _
            & " group by WHTWAVE1.WHSE_CODE, WHTINST1.LOCATION_CODE, WHTINST2.BAR_CODE, WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE, 0 INST" & vbCrLf _
            & ", SUM (LOCATION_QTY_WAVE) LOCB" & vbCrLf _
            & " from WHTLOCB1 WHERE LOCATION_QTY_WAVE <> 0" & vbCrLf _
            & "   And WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" & vbCrLf _
            & " group by WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ") group by WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE"

        Fill_Records("WHTLOCBR", "", True, sql)

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS" & vbCrLf _
            & sql & ";" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE WHTLOCB1 SET LOCATION_QTY_WAVE = R1.INST" & vbCrLf _
            & "WHERE WHSE_CODE = '" & WHSE_CODE & "' AND LOCATION_CODE = R1.LOCATION_CODE AND BAR_CODE = R1.BAR_CODE AND STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
            & "AND COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "END LOOP; END; END;"
        If ASCMAIN1.Running_in_VS Then
            Stop
            'ASCDATA1.ExecuteSQL()
            ' LINE ABOVE SHOULD FIX THINGS
        End If

    End Sub

    Sub Setup_WhseMap()
        ASCMAIN1.sql = ""

    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Special").Visible = False

        If tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Reconcile Qtys" Then
            Setup_Reconciliation()
        ElseIf tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Reconcile Qtys" Then
            Setup_WhseMap()
        ElseIf tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Location Master" Then
            Setup_Location_Grids()
            chkRESOLUTION.Checked = False
        ElseIf tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Styles && Locations" Then
            Setup_Styles_and_Locations()
        End If
    End Sub

    Sub Setup_Location_Grids()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Filling Location Data", "")
        Fill_Records("WHTLOCMM")

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.WaitCursor

    End Sub

    Sub Setup_Styles_and_Locations()
        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            MsgBox("Please Select a warehouse and re-select tab", MsgBoxStyle.Exclamation, "No Warehouse")
            Exit Sub
        End If

        UltraExplorerBar1.Groups("Special").Visible = True

        If chkViewAsst.Checked = True Then
            ASCMAIN1.sql = "Select " & vbCrLf _
                & If(ASCMAIN1.CLIENT = "RGI", " SUBSTR(X.LOCATION_CODE,1,3) AA", " SUBSTR(X.LOCATION_CODE,1,2) AA") & vbCrLf _
                & If(ASCMAIN1.CLIENT = "RGI", " , SUBSTR(X.LOCATION_CODE,5,2) BBB", ", SUBSTR(X.LOCATION_CODE,4,3) BBB") & vbCrLf _
                & ", SUBSTR(X.LOCATION_CODE,8,1) L " & vbCrLf _
                & ", X.WHSE_CODE, X.LOCATION_CODE, WHTLOCM1.LOCATION_DESC" & vbCrLf _
                & ", WHTLOCM1.LOCATION_SINGLE_LOAD, WHTLOCM1.LOCATION_LOCKED, WHTLOCM1.LOCATION_NOT_WAVED, WHTLOCM1.LOCATION_USE" & vbCrLf _
                & ", X.STYLE_CODE, ICTSTYL1.STYLE_DESC, X.COLOR_CODE, X.CASES, X.UNITS" & vbCrLf _
                & " from (Select" & vbCrLf _
                & "WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", COUNT (DISTINCT BAR_CODE) CASES, SUM (LOCATION_QTY) UNITS" & vbCrLf _
                & " from WHTLOCB1 where LOCATION_QTY <> 0" & vbCrLf _
                & " group by WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE) X, ICTSTYL1, WHTLOCM1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = :PARM1 " & vbCrLf _
                & "   and REGEXP_LIKE (UPPER(ICTSTYL1.STYLE_DESC), '" & txtStyle_Search.Text & "') "
            Fill_Records("WHTLOCBJ", Absx1.txtFor("WHSE_CODE").Text, True, ASCMAIN1.sql)
        Else
            Fill_Records("WHTLOCBJ", Absx1.txtFor("WHSE_CODE").Text)

            ASCMAIN1.sql = "Select " & vbCrLf _
               & If(ASCMAIN1.CLIENT = "RGI", " SUBSTR(X.LOCATION_CODE,1,3) AA", " SUBSTR(X.LOCATION_CODE,1,2) AA") & vbCrLf _
               & If(ASCMAIN1.CLIENT = "RGI", " , SUBSTR(X.LOCATION_CODE,5,2) BBB", ", SUBSTR(X.LOCATION_CODE,4,3) BBB") & vbCrLf _
               & ", SUBSTR(X.LOCATION_CODE,8,1) L " & vbCrLf _
               & ", X.WHSE_CODE, X.LOCATION_CODE, WHTLOCM1.LOCATION_DESC" & vbCrLf _
               & ", WHTLOCM1.LOCATION_SINGLE_LOAD, WHTLOCM1.LOCATION_LOCKED, WHTLOCM1.LOCATION_NOT_WAVED, WHTLOCM1.LOCATION_USE" & vbCrLf _
               & " from (Select WHSE_CODE, LOCATION_CODE from WHTLOCM1 minus Select Distinct WHSE_CODE, LOCATION_CODE from WHTLOCB1 where LOCATION_QTY <> 0) X, WHTLOCM1" & vbCrLf _
               & " where WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE" & vbCrLf _
               & " and WHTLOCM1.WHSE_CODE = :PARM1"
            Fill_Records("WHTLOCBJ", Absx1.txtFor("WHSE_CODE").Text, False, ASCMAIN1.sql)
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Filling Location Data", "")
        grdWHTLOCBJ.Text = "Styles & Locations in Warehouse " & Absx1.txtFor("WHSE_CODE").Text


       

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            With grdWHTLOCBJ.DisplayLayout.Bands(0)
                .Columns("AA").Hidden = True
                .Columns("BBB").Hidden = True
                .Columns("L").Hidden = True
                .Columns("LOCATION_SINGLE_LOAD").Hidden = True
                .Columns("LOCATION_LOCKED").Hidden = True
                .Columns("LOCATION_NOT_WAVED").Hidden = True
                .Columns("LOCATION_USE").Hidden = True
            End With

        End If
        Sort_grdColumns(grdWHTLOCBJ, "STYLE_CODE")
        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.WaitCursor
    End Sub

    Private Sub grdWHTCYCL1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTCYCL1.AfterRowActivate
        If grdWHTCYCL1.ActiveRow Is Nothing Then Exit Sub
        Fill_WHTCYCL2()
        If grdWHTCYCL1.ActiveRow.Cells("CYCLE_type").Value = "C" Then
            optRESOLUTION.Visible = False
            btnRESOLUTION.Text = "Ack"
        Else
            optRESOLUTION.Visible = True
            btnRESOLUTION.Text = "Update"
        End If
    End Sub

    Sub Fill_WHTCYCL2()
        If grdWHTCYCL1.ActiveRow Is Nothing Then
            dst.Tables("WHTCYCL2").Rows.Clear()
            grdWHTCYCL2.Text = "Cycle Count Details"
        Else
            Fill_Records("WHTCYCL2", New Object() {grdWHTCYCL1.ActiveRow.Cells("CYCLE_NO").Value}, True)
            Sort_grdColumns(grdWHTCYCL2, "BAR_CODE")
            grdWHTCYCL2.Text = "Cycle Count Details for Cycle No : " & grdWHTCYCL1.ActiveRow.Cells("CYCLE_NO").Value
        End If

    End Sub

    Sub Fill_WHTCYCL1()
        ASCMAIN1.sql = "Select * from WHTCYCL1 Where LOCATION_CODE = '" & grdWHTLOCMM.ActiveRow.Cells("LOCATION_CODE").Value & "'"
        Fill_Records("WHTCYCL1", , , ASCMAIN1.sql)
        Sort_grdColumns(grdWHTCYCL1, "INIT_DATE".ToLower)

        grdWHTCYCL1.Text = "Cycle Count For Location: " & grdWHTLOCMM.ActiveRow.Cells("LOCATION_CODE").Value
    End Sub
    Private Sub grdWHTCYCL1_DoubleClick(sender As Object, e As System.EventArgs) Handles grdWHTCYCL1.DoubleClick
        If grdWHTLOCMM.ActiveRow IsNot Nothing AndAlso grdWHTLOCMM.ActiveRow.IsDataRow Then
            optViewBy.Value = "L"
            Absx1.txtFor("WHSE_CODE").Text = grdWHTCYCL1.ActiveRow.Cells("WHSE_CODE").Text
            Absx1.txtFor("LOCATION_CODE").Text = grdWHTCYCL1.ActiveRow.Cells("LOCATION_CODE").Text
            Click_Command("Select")
        End If
    End Sub

    Private Sub grdWHTCYCL1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTCYCL1.InitializeRow
        If e.Row.Cells("CYCLE_STATUS").Value & "" = "D" Then
            e.Row.Cells("CYCLE_STATUS").Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("CYCLE_STATUS").Value & "" = "G" Then
            e.Row.Cells("CYCLE_STATUS").Appearance.ForeColor = Drawing.Color.Green
        End If
        'If e.Row.Cells("CYCLE_type").Value = "C" Then
        '    optRESOLUTION.Visible = False
        '    btnRESOLUTION.Text = "Ack"
        'Else
        '    optRESOLUTION.Visible = True
        '    btnRESOLUTION.Text = "Update"
        'End If
    End Sub


    Private Sub btnRESOLUTION_Click(sender As System.Object, e As System.EventArgs) Handles btnRESOLUTION.Click

        Dim WHSE_TRAN_NO As String = ""
        Dim WHSE_TRAN_LNO As Integer
        Dim BAR_CODE As String
        Dim LOCATION_CODE_ORIG As String


        Dim WHSE_CODE As String = grdWHTCYCL1.ActiveRow.Cells("WHSE_CODE").Value & ""
        Dim LOCATION_CODE As String = grdWHTCYCL1.ActiveRow.Cells("LOCATION_CODE").Value & ""
        Dim EMsg As String = ""
        Dim CYCLE_RESOLUTION As String = ""
        CYCLE_RESOLUTION = ""
        If grdWHTCYCL1.ActiveRow.Cells("CYCLE_TYPE").Value = "C" Then
            CYCLE_RESOLUTION = "U"
        ElseIf grdWHTCYCL1.ActiveRow.Cells("CYCLE_TYPE").Value = "V" Then
            If optRESOLUTION.Value & "" = "" Then
                MsgBox("Please Select a Resolution", MsgBoxStyle.OkOnly, "Cannot Proceed")
                Exit Sub
            End If
            If optRESOLUTION.Value = "V" Then
                CYCLE_RESOLUTION = "V"
            Else

                ' 1st check check to make sure that Location in Cycle is stil locked
                Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE})
                If rowWHTLOCM1.Item("LOCATION_LOCKED") & "" = "" Then
                    EMsg &= vbCr & "Location " & LOCATION_CODE & " is not locked, Unlocked since Cycle scan, Please re-do Cycle Count"
                End If

                ASCMAIN1.sql = "Select Distinct BAR_CODE,'0' SEL from WHTLOCB1" & vbCrLf _
                    & " where WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                    & " and WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                    & " and WHTLOCB1.LOCATION_QTY > 0"
                Fill_Records("WHTCYCLX", "", True, ASCMAIN1.sql)

                For Each rowWHTCYCL2 As DataRow In dst.Tables("WHTCYCL2").Select("")
                    ' 2nd check If CYCLE_NEW = "1" then  to make sure that Location is the same as when cycle scan performed
                    BAR_CODE = rowWHTCYCL2.Item("BAR_CODE")
                    If rowWHTCYCL2.Item("CYCLE_NEW") & "" = "1" And rowWHTCYCL2.Item("BAR_CODE_INVALID") & "" <> "1" Then

                        ASCMAIN1.sql = "Select Distinct LOCATION_CODE from WHTLOCB1" & vbCrLf _
                              & " where BAR_CODE = '" & BAR_CODE & "'" _
                              & " and WHSE_CODE = '" & WHSE_CODE & "'" _
                              & " and LOCATION_QTY > 0"
                        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                        If rows.Length = 0 Then
                            EMsg &= vbCr & " Case ID " & BAR_CODE & " is not found in Warehouse "
                        ElseIf rows.Length > 1 Then
                            EMsg &= vbCr & " Case ID  " & BAR_CODE & " found in Multiple Locations with Qty - Call ABS"
                        ElseIf rowWHTCYCL2.Item("LOCATION_CODE_ORIG") <> rows(0).Item("LOCATION_CODE") Then
                            EMsg &= vbCr & "Locations have changed since Cycle scan, Please re-do Cycle Count"
                        End If
                    Else
                        ' 3nd check make sure that Locations in Scan are accounted for"

                        Dim rowWHTCYCLX As DataRow = dst.Tables("WHTCYCLX").Rows.Find _
                                 (New String() {BAR_CODE})
                        If rowWHTCYCLX Is Nothing Then
                            If rowWHTCYCL2.Item("BAR_CODE_INVALID") & "" <> "1" Then
                                EMsg &= vbCr & "Location Cases in Location have changed since Cycle Count, Please re-do Cycle Count"
                            End If
                        Else
                            With rowWHTCYCLX
                                .Item("SEL") = "1"
                            End With
                        End If

                    End If
                Next
                If dst.Tables("WHTCYCLX").Select("SEL <> '1'").Length <> 0 Then
                    EMsg &= vbCr & "Location Cases have changed since Cycle scan, Please re-do Cycle Count"
                End If

                If EMsg <> "" Then
                    MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Exit Sub
                Else

                    ' Update (Posted)
                    dst.Tables("WHTMOVE1").Rows.Clear()
                    dst.Tables("WHTMOVE2").Rows.Clear()
                    WHSE_TRAN_NO = ""
                    WHSE_TRAN_LNO = 0

                    For Each rowWHTCYCL2 As DataRow In dst.Tables("WHTCYCL2").Select("CYCLE_NEW = '1' or isnull(CYCLE_SCAN,'')<>'1'")
                        BAR_CODE = rowWHTCYCL2.Item("BAR_CODE")
                        LOCATION_CODE_ORIG = rowWHTCYCL2.Item("LOCATION_CODE_ORIG") & ""

                        If rowWHTCYCL2.Item("BAR_CODE_INVALID") & "" <> "1" Then

                            Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", BAR_CODE)
                            CYCLE_RESOLUTION = "U"

                            ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1 " _
                                  & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                                  & " and  WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE_ORIG & "'" _
                                  & " and  WHTLOCB1.BAR_CODE = '" & BAR_CODE & "'" _
                                  & " and  WHTLOCB1.LOCATION_QTY > 0 "
                            For Each rowWHTLOCB1 As DataRow In ASCDATA1.GetDataTable.Select("")

                                If WHSE_TRAN_NO = "" Then
                                    WHSE_TRAN_NO = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
                                End If

                                Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
                                With rowWHTMOVE2


                                    .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                                    WHSE_TRAN_LNO += 1
                                    .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                                    If rowWHTCYCL2.Item("CYCLE_NEW") & "" = "1" Then
                                        .Item("LOCATION_CODE_FROM") = LOCATION_CODE_ORIG
                                        .Item("LOCATION_CODE_TO") = LOCATION_CODE
                                        Dim LOAD_NO As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
                                        ' REM NEW LOAD NO
                                        .Item("LOAD_NO_FROM") = LOAD_NO
                                        .Item("LOAD_NO_TO") = ""
                                    ElseIf rowWHTCYCL2.Item("CYCLE_SCAN") & "" = "" Then
                                        .Item("LOCATION_CODE_FROM") = LOCATION_CODE
                                        .Item("LOCATION_CODE_TO") = rowICTWHSE1.Item("WHSE_LOC_LNF") & ""
                                        .Item("LOAD_NO_FROM") = rowWHTBARC1.Item("LOAD_NO")
                                        .Item("LOAD_NO_TO") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                                    End If
                                    .Item("BAR_CODE") = BAR_CODE

                                    .Item("WHSE_TRAN_QTY") = rowWHTLOCB1.Item("LOCATION_QTY")
                                    .Item("STYLE_CODE") = rowWHTLOCB1.Item("STYLE_CODE")
                                    .Item("COLOR_CODE") = rowWHTLOCB1.Item("COLOR_CODE")
                                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                    .Item("INIT_DATE") = DATETIME_STAMP
                                    .Item("STATUS") = "U"

                                    ' .Item("ERROR_CODES") = String.Empty
                                End With
                                dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

                            Next
                        End If
                    Next
                    If CYCLE_RESOLUTION = "U" Then
                        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
                        With rowWHTMOVE1
                            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                            .Item("WHSE_TRAN_TYPE") = "M"
                            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                            .Item("WHSE_CODE") = WHSE_CODE
                            .Item("STATUS") = "U"
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_DATE") = DATETIME_STAMP
                        End With
                        dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
                    End If
                End If

            End If
        End If

        If CYCLE_RESOLUTION <> "" Then
            BeginTrans()

            If CYCLE_RESOLUTION = "U" And grdWHTCYCL1.ActiveRow.Cells("CYCLE_TYPE").Value = "V" Then
                Update_Record_TDA("WHTMOVE1")
                Update_Record_TDA("WHTMOVE2")
                ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {WHSE_TRAN_NO, 0, 1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})
            End If

            ASCMAIN1.sql = "Update WHTLOCM1 Set LOCATION_LOCKED  = NULL" _
            & " where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", _
                                                    New String() {WHSE_CODE, _
                                                    LOCATION_CODE})

            ASCMAIN1.sql = "Update WHTCYCL1 Set CYCLE_RESOLUTION = '" & CYCLE_RESOLUTION & "'," _
            & " LAST_OPER = '" & ASCMAIN1.USER_ID & "', LAST_DATE = sysdate " _
            & " Where CYCLE_NO = '" & grdWHTCYCL1.ActiveRow.Cells("CYCLE_NO").Value & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            CommitTrans()
        End If

        MsgBox("Resolution Recorded", vbOKOnly, "Success")
        splWHTCYCL2.Panel2Collapsed = True
    End Sub

    Private Sub grdWHTCYCL2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTCYCL2.AfterRowActivate
        If grdWHTCYCL1.ActiveRow Is Nothing Then Exit Sub
        splWHTCYCL2.Panel2Collapsed = IIf(grdWHTCYCL1.ActiveRow.Cells("CYCLE_RESOLUTION").Value = "P", False, True)

    End Sub


    Private Sub grdWHTLOCMM_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTLOCMM.InitializeRow
        If e.Row.Cells("CYCLE_STATUS").Value & "" = "D" Then
            e.Row.Cells("CYCLE_STATUS").Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("CYCLE_STATUS").Value & "" = "G" Then
            e.Row.Cells("CYCLE_STATUS").Appearance.ForeColor = Drawing.Color.Green
        End If
    End Sub

    Private Sub grdWHTLOCMM_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTLOCMM.InitializeLayout

    End Sub

    Private Sub chkRESOLUTION_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkRESOLUTION.CheckedChanged
        Dim dvw As DataView = DirectCast(grdWHTLOCMM.DataSource, DataTable).DefaultView
        Dim Row_filter As String = IIf(chkRESOLUTION.Checked, "CYCLE_RESOLUTION = 'P'", "")
        dvw.RowFilter = Row_filter

    End Sub

    Private Sub tabCycle_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabCycle.SelectedTabChanged
        Load_Waves()
        Load_WHTLOCBM()
        ' Stop
    End Sub


    Private Sub grdWHTLOCB1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTLOCB1.InitializeLayout

    End Sub

    Private Sub chkViewAsst_CheckedValueChanged(sender As Object, e As EventArgs) Handles chkViewAsst.CheckedValueChanged
        Setup_Styles_and_Locations()
    End Sub
End Class