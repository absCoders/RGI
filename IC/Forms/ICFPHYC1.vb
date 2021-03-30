Imports System.Drawing
Imports System.Math

Public Class ICFPHYC1
    Dim rowICTPHYC1 As DataRow
    Dim location_support As Boolean = False
    Dim rowICTWHSE1 As DataRow
    Dim WHSE_CODE As String
    Dim TICKET_NO As String
    Dim grdAttention_app As New Infragistics.Win.Appearance
    Dim grdWARNING_app As New Infragistics.Win.Appearance
    Dim SelUser As String

    ' NOTE THAT IF WE DO NOT INITIALIZE COUNTS TABLES AT MONTH END, THAT THIS SCREEN WILL SHOW COUNTS (WHICH IS USEFUL) AFTER THE PI HAS BEEN POSTED
    '  HOWEVER, THE VARIANCE WILL WORK ONLY FOR LCOATABLE WHSES SINCE WE COMPARE TO WHTLOCB0 (SNAPSHOT BY LOCATION), 
    '  AND WE DIDN'T EVEN THINK TO INITIALIZE THAT TABLE.  THE BOOK INVENTORY WILL SHOW WITH BAD DATA FOR NON-LOCATABLE WHSES, SINCE IT IS LOOKING AT ICTSTAT1.
    '  BUT THIS MIGHT BE EASILY FIXED BY USING THE YP OF THE LAST PI UPDATE

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFPHYCI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        Dim RGI_CODE As String = ""

        If ASCMAIN1.CLIENT = "RGI" Then
            RGI_CODE = " and NVL(ICTPHYC2.STATUS,'A') = 'A' "
        End If

        With dst
            'ASCMAIN1.sql = "Select ICTPHYC1.*, X.STYLE_CODE, X.SC, X.TOTAL_COUNT, UNIT_VARIANCE, ABSOLUTE_VARIANCE" _
            '& " from ICTPHYC1, (Select WHSE_CODE, TICKET_NO, Min (STYLE_CODE) STYLE_CODE, Min (STYLE_CODE || '-' || COLOR_CODE) SC" _
            '& ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) TOTAL_COUNT" _
            '& " from ICTPHYC2 where WHSE_CODE = :PARM1 " & RGI_CODE & " group by WHSE_CODE, TICKET_NO) X" _
            '& " where ICTPHYC1.WHSE_CODE = :PARM1" _
            '& "   and X.WHSE_CODE (+) = ICTPHYC1.WHSE_CODE" _
            '& "   and X.TICKET_NO (+) = ICTPHYC1.TICKET_NO"

            ASCMAIN1.sql = "Select ICTPHYC1.*, X.STYLE_CODE, X.SC, X.TOTAL_COUNT, UNIT_VARIANCE, ABSOLUTE_VARIANCE, nvl(LAST_ACTIVITY,'01-JAN-1999') LAST_ACTIVITY  " & vbCrLf _
                & " from ICTPHYC1, (Select WHSE_CODE, TICKET_NO, Min (STYLE_CODE) STYLE_CODE, Min (STYLE_CODE || '-' || COLOR_CODE) SC" & vbCrLf _
                & " , Sum (TOTAL_COUNT) TOTAL_COUNT, SUM(VARIANCE) UNIT_VARIANCE, SUM(ABS(VARIANCE)) ABSOLUTE_VARIANCE " & vbCrLf _
                & " from( " & vbCrLf _
                & " select  P.WHSE_CODE, P.TICKET_NO, P.LOCATION_CODE, P.STYLE_CODE, P.COLOR_CODE, P.TOTAL_COUNT, nvl(LOCATION_QTY, 0) - nvl(BOOK_INVTY_ADJ, 0) LOCATION_QTY,  (P.TOTAL_COUNT - nvl(LOCATION_QTY, 0) + nvl(BOOK_INVTY_ADJ, 0)) VARIANCE" & vbCrLf _
                & " from WHTLOCB0, " & vbCrLf _
                & " (select ICTPHYC1.WHSE_CODE, ICTPHYC1.TICKET_NO, LOCATION_CODE, STYLE_CODE, COLOR_CODE, sum(NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) TOTAL_COUNT " & vbCrLf _
                & " from ICTPHYC1, ICTPHYC2 " & vbCrLf _
                & " where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE " & vbCrLf _
                & " And ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO and ICTPHYC2.WHSE_CODE = :PARM1 " & RGI_CODE & vbCrLf _
                & "  group by ICTPHYC1.WHSE_CODE,ICTPHYC1.TICKET_NO, LOCATION_CODE, STYLE_CODE, COLOR_CODE) P " & vbCrLf _
                & " where WHTLOCB0.WHSE_CODE(+) = P.WHSE_CODE " & vbCrLf _
                & " and WHTLOCB0.LOCATION_CODE(+) = P.LOCATION_CODE " & vbCrLf _
                & " and WHTLOCB0.STYLE_CODE(+) = P.STYLE_CODE " & vbCrLf _
                & " and WHTLOCB0.COLOR_CODE(+) = P.COLOR_CODE ) " & vbCrLf _
                & " group by WHSE_CODE, TICKET_NO) X, " & vbCrLf _
                & " (select WHSE_CODE, LOCATION_CODE, MAX(INIT_DATE) LAST_ACTIVITY " & vbCrLf _
                & " from WHTLOCB2 " & vbCrLf _
                & "where WHSE_CODE = :PARM1 " & vbCrLf _
                & " group by whse_code, location_code) A " & vbCrLf _
                & " where ICTPHYC1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and X.WHSE_CODE (+) = ICTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and X.TICKET_NO (+) = ICTPHYC1.TICKET_NO" & vbCrLf _
                & "   and A.WHSE_CODE (+) = ICTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and A.LOCATION_CODE (+) = ICTPHYC1.LOCATION_CODE"
            Create_TDA(.Tables.Add, "ICTPHYCX", "**", 0, False, "V")
            .Tables("ICTPHYCX").Columns("TOTAL_COUNT").DataType = GetType(System.Int64)
            .Tables("ICTPHYCX").Columns("UNIT_VARIANCE").DataType = GetType(System.Int64)
            .Tables("ICTPHYCX").Columns("ABSOLUTE_VARIANCE").DataType = GetType(System.Int64)
            .Tables("ICTPHYCX").Columns.Add("SELECTED")
            .Tables("ICTPHYCX").Columns("SELECTED").DefaultValue = "0"

            '& ", NVL(ICTSTYC1.STYLE_COST_FIFO,ICTSTYL1.STYLE_COST) STYLE_COST" & vbCrLf _
            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
                                & ", ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
                                & ", X.BOOK, X.PHYS, (nvl( X.PHYS,0) - nvl(X.BOOK,0)) VARIANCE " & vbCrLf _
                                & ", (nvl( X.PHYS,0) - nvl(X.BOOK,0))*nvl(ICTSTYC1.STYLE_COST_FIFO,0) VARIANCE_COST " & vbCrLf _
                                & ", sum((nvl( X.PHYS,0) - nvl(X.BOOK,0))*nvl(ICTSTYC1.STYLE_COST_FIFO,0)) over (PARTITION BY X.STYLE_CODE) STYLE_COST_VAR " & vbCrLf _
                                & ",  COUNTED_LOCS,  BOOKED_LOCS  " & vbCrLf _
                                & " from ICTSTYL1, ICTSTYC1, (" & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, Sum (BOOK) BOOK, Sum (PHYS) PHYS " & vbCrLf _
                                & ", listagg (COUNTED_LOC, ',') within group ( order by COUNTED_LOC) COUNTED_LOCS  " & vbCrLf _
                                & ", listagg (BOOKED_LOC, ',') within group ( order by BOOKED_LOC) BOOKED_LOCS  " & vbCrLf _
                                & "from (" & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, 0 BOOK, Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS" & vbCrLf _
                                & ", listagg (LOCATION_CODE, ',') within group ( order by LOCATION_CODE) COUNTED_LOC " & vbCrLf _
                                & ", '' BOOKED_LOC  " & vbCrLf _
                                & " from ICTPHYC1, ICTPHYC2 where ICTPHYC2.WHSE_CODE = :PARM1  and NVL(ICTPHYC2.STATUS,'A') = 'A' " & vbCrLf _
                                & " and ICTPHYC1.TICKET_NO =  ICTPHYC2.TICKET_NO " & vbCrLf _
                                & " group by STYLE_CODE, COLOR_CODE" _
                                & " union " & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, Sum (LOCATION_QTY) - sum(BOOK_INVTY_ADJ) BOOK, 0 PHYS" & vbCrLf _
                                & ", '' COUNTED_LOC " & vbCrLf _
                                & ", listagg (LOCATION_CODE, ',') within group ( order by LOCATION_CODE) BOOKED_LOC  " & vbCrLf _
                                & " from WHTLOCB0 where WHSE_CODE = :PARM2 group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                                & " union " & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, WHSE_QTY_BEG BOOK, 0 PHYS, '' BOOKED_LOC , '' COUNTED_LOC" & vbCrLf _
                                & " from ICTSTAT1 where WHSE_CODE = :PARM3 and OPS_YYYYPP = :PARM4" & vbCrLf _
                                & ") group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                                & ") X " & vbCrLf _
                                & "where X.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                                & "   and ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                                & "   and ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE"
            Else
                ASCMAIN1.sql = "Select X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
                                & ", ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
                                & ", X.BOOK, X.PHYS from ICTSTYL1, ICTSTYC1, (" & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, Sum (BOOK) BOOK, Sum (PHYS) PHYS from (" & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, 0 BOOK, Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS" & vbCrLf _
                                & " from ICTPHYC2 where WHSE_CODE = :PARM1  group by STYLE_CODE, COLOR_CODE" _
                                & " union " & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, Sum (LOCATION_QTY) - sum(BOOK_INVTY_ADJ) BOOK, 0 PHYS" & vbCrLf _
                                & " from WHTLOCB0 where WHSE_CODE = :PARM2 group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                                & " union " & vbCrLf _
                                & "Select STYLE_CODE, COLOR_CODE, WHSE_QTY_BEG BOOK, 0 PHYS" & vbCrLf _
                                & " from ICTSTAT1 where WHSE_CODE = :PARM3 and OPS_YYYYPP = :PARM4" & vbCrLf _
                                & ") group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                                & ") X where X.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                                & "   and ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                                & "   and ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE"
            End If
            Create_TDA(.Tables.Add, "ICTPHYCV", "**", 0, False, "VVVV")
            .Tables("ICTPHYCV").Columns("BOOK").DataType = GetType(System.Int64)
            .Tables("ICTPHYCV").Columns("PHYS").DataType = GetType(System.Int64)
            If Not .Tables("ICTPHYCV").Columns.Contains("VARIANCE") Then
                .Tables("ICTPHYCV").Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS,0) - ISNULL(BOOK,0)")
                .Tables("ICTPHYCV").Columns.Add("VARIANCE_COST", GetType(System.Int64), "ISNULL(STYLE_COST,0) * (ISNULL(PHYS,0) - ISNULL(BOOK,0))")
                .Tables("ICTPHYCV").Columns.Add("STYLE_COST_VAR", GetType(System.Int64))
                .Tables("ICTPHYCV").Columns.Add("COUNTED_LOCS", GetType(System.String))
                .Tables("ICTPHYCV").Columns.Add("BOOKED_LOCS", GetType(System.String))
            Else
                .Tables("ICTPHYCV").Columns("VARIANCE").DataType = GetType(System.Int64)
                .Tables("ICTPHYCV").Columns("VARIANCE_COST").DataType = GetType(System.Int64)
                .Tables("ICTPHYCV").Columns("STYLE_COST_VAR").DataType = GetType(System.Int64)
            End If



            Create_TDA(.Tables.Add, "ICTWHSE1", "*")

            Create_TDA(.Tables.Add, "ICTPHYC1", "*")

            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select ICTPHYC2.*,  CASE WHEN NVL(STATUS,'A') = 'A' THEN NVL(V.V_TTL,0) ELSE 0 END V_TTL, S.STYLE_DESC" & vbCrLf _
                    & " from ICTPHYC2 , ICTSTYL1 S, " & vbCrLf _
                    & " ( Select TICKET_NO, STYLE_CODE, COLOR_CODE, SUM(NVL(COUNT_CTNS * CARTON_PACK_QTY, 0) + NVL(COUNT_LOOSE,0))  V_TTL " & vbCrLf _
                    & " from ICTPHYC2  " & vbCrLf _
                    & " where STATUS = 'V' " & vbCrLf _
                    & " group by TICKET_NO, STYLE_CODE, COLOR_CODE) v " & vbCrLf _
                    & " where ICTPHYC2.TICKET_NO = V.TICKET_NO(+)" & vbCrLf _
                    & " and ICTPHYC2.STYLE_CODE = V.STYLE_CODE(+)" & vbCrLf _
                    & " and ICTPHYC2.COLOR_CODE = V.COLOR_CODE(+)" & vbCrLf _
                    & " and ICTPHYC2.STYLE_CODE = S.STYLE_CODE"
                Create_TDA(.Tables.Add, "ICTPHYC2", "**", 2)
                .Tables("ICTPHYC2").Columns.Add("TOTAL_COUNT", GetType(System.Int64), "ISNULL(COUNT_CTNS,0) * ISNULL(CARTON_PACK_QTY,0) + ISNULL(COUNT_LOOSE,0)")
                .Tables("ICTPHYC2").Columns("V_TTL").DataType = GetType(System.Int64)
            Else
                ASCMAIN1.sql = "Select ICTPHYC2.*, ICTSTYL1.STYLE_DESC" _
               & " from ICTPHYC2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = ICTPHYC2.STYLE_CODE" '& RGI_CODE
                Create_TDA(.Tables.Add, "ICTPHYC2", "**", 2)
                .Tables("ICTPHYC2").Columns.Add("TOTAL_COUNT", GetType(System.Int64), "ISNULL(COUNT_CTNS,0) * ISNULL(CARTON_PACK_QTY,0) + ISNULL(COUNT_LOOSE,0)")
            End If

            ASCMAIN1.sql = "Select ICTPHYC2.*" _
                & ", ICTPHYC1.LOCATION_CODE, ICTPHYC1.COUNT_BY, ICTPHYC1.INIT_OPER, ICTPHYC1.INIT_DATE" _
                & " from ICTPHYC2,ICTPHYC1" _
                & " where ICTPHYC2.WHSE_CODE = :PARM1 and ICTPHYC2.STYLE_CODE = :PARM2 and ICTPHYC2.COLOR_CODE = :PARM3" _
                & "   and ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE" _
                & "   and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" & RGI_CODE
            Create_TDA(.Tables.Add, "ICTPHYCI", "**", 0, False, "VVV", 3)
            .Tables("ICTPHYCI").Columns.Add("TOTAL_COUNT", GetType(System.Int64), "ISNULL(COUNT_CTNS,0) * ISNULL(CARTON_PACK_QTY,0) + ISNULL(COUNT_LOOSE,0)")

            ASCMAIN1.sql = "Select WHSE_CODE ,LOCATION_CODE ,BAR_CODE ,STYLE_CODE ,COLOR_CODE ,(WHTLOCB0.LOCATION_QTY - WHTLOCB0.BOOK_INVTY_ADJ) LOCATION_QTY ,INIT_DATE ,INIT_OPER ,LAST_DATE ,LAST_OPER ,LOCATION_QTY_WAVE" _
                & " from WHTLOCB0 where WHSE_CODE = :PARM1 and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3"
            Create_TDA(.Tables.Add, "WHTLOCB0", "**", 0, False, "VVV", 5)

            ASCMAIN1.sql = "Select * from WHTLOCM1"
            Create_TDA(.Tables.Add, "WHTLOCM1", "**", 0, False)

            'ASCMAIN1.sql = "Select * from ICTCLAS1"
            'Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            ASCMAIN1.sql = "Select WHTLOCM1.WHSE_CODE, WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & ", A.TICKETS, A.PHYS_INIT, A.PHYS_LAST, A.VOIDED, A.EMPTY" & vbCrLf _
                & ", B.PHYS_UNITS, B.PHYS_VALUE, B.PHYS_STYLE_COLORS, B.PHYS_SCMIN, B.PHYS_SCMAX" & vbCrLf _
                & ", C.BOOK_UNITS, C.BOOK_VALUE, C.BOOK_STYLE_COLORS, C.BOOK_SCMIN, C.BOOK_SCMAX" & vbCrLf _
                & " from WHTLOCM1," & vbCrLf _
                & "  (Select ICTPHYC1.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", COUNT (*) TICKETS, MIN (INIT_DATE) PHYS_INIT, MAX (INIT_DATE) PHYS_LAST" & vbCrLf _
                & ", MAX (CASE WHEN ICTPHYC1.TICKET_STATUS = 'V' THEN ICTPHYC1.TICKET_NO ELSE NULL END) VOIDED" & vbCrLf _
                & ", MAX (CASE WHEN ICTPHYC1.TICKET_STATUS = 'E' THEN ICTPHYC1.TICKET_NO ELSE NULL END) EMPTY" & vbCrLf _
                & " from ICTPHYC1" & vbCrLf _
                & " where ICTPHYC1.WHSE_CODE = :PARM1" _
                & " group by ICTPHYC1.WHSE_CODE, ICTPHYC1.LOCATION_CODE) A" & vbCrLf _
                & ", (Select ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", SUM (ICTPHYC2.COUNT_CTNS * ICTPHYC2.CARTON_PACK_QTY + ICTPHYC2.COUNT_LOOSE) PHYS_UNITS" & vbCrLf _
                & ", SUM ((ICTPHYC2.COUNT_CTNS * ICTPHYC2.CARTON_PACK_QTY + ICTPHYC2.COUNT_LOOSE) * ICTSTYC1.STYLE_COST_FIFO) PHYS_VALUE" & vbCrLf _
                & ", COUNT (DISTINCT ICTPHYC2.STYLE_CODE || ICTPHYC2.COLOR_CODE) PHYS_STYLE_COLORS" _
                & ", MIN (ICTPHYC2.STYLE_CODE || '-' || ICTPHYC2.COLOR_CODE) PHYS_SCMIN" & vbCrLf _
                & ", MAX (ICTPHYC2.STYLE_CODE || '-' || ICTPHYC2.COLOR_CODE) PHYS_SCMAX" & vbCrLf _
                & " from ICTPHYC1,ICTPHYC2,ICTSTYC1" & vbCrLf _
                & " where ICTPHYC1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and ICTPHYC2.WHSE_CODE = ICTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and ICTPHYC2.TICKET_NO = ICTPHYC1.TICKET_NO" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = ICTPHYC2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = ICTPHYC2.COLOR_CODE" & RGI_CODE & vbCrLf _
                & " group by ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE) B" & vbCrLf _
                & ", (Select WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", SUM (WHTLOCB0.LOCATION_QTY - WHTLOCB0.BOOK_INVTY_ADJ) BOOK_UNITS" & vbCrLf _
                & ", SUM ((WHTLOCB0.LOCATION_QTY - WHTLOCB0.BOOK_INVTY_ADJ) * ICTSTYC1.STYLE_COST_FIFO) BOOK_VALUE" & vbCrLf _
                & ", COUNT (DISTINCT WHTLOCB0.STYLE_CODE || WHTLOCB0.COLOR_CODE) BOOK_STYLE_COLORS" _
                & ", MIN (WHTLOCB0.STYLE_CODE || '-' || WHTLOCB0.COLOR_CODE) BOOK_SCMIN" & vbCrLf _
                & ", MAX (WHTLOCB0.STYLE_CODE || '-' || WHTLOCB0.COLOR_CODE) BOOK_SCMAX" & vbCrLf _
                & " from WHTLOCB0,ICTSTYC1" & vbCrLf _
                & " where WHTLOCB0.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = WHTLOCB0.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = WHTLOCB0.COLOR_CODE" & vbCrLf _
                & " group by WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE) C" & vbCrLf _
                & " where A.WHSE_CODE (+) = WHTLOCM1.WHSE_CODE" & vbCrLf _
                & "   and A.LOCATION_CODE (+) = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & "   and B.WHSE_CODE (+) = WHTLOCM1.WHSE_CODE" & vbCrLf _
                & "   and B.LOCATION_CODE (+) = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & "   and C.WHSE_CODE (+) = WHTLOCM1.WHSE_CODE" & vbCrLf _
                & "   and C.LOCATION_CODE (+) = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTPHYCL", "**", 0, False, "V", 2)
            With .Tables("ICTPHYCL")
                .Columns("TICKETS").DataType = GetType(System.Int64)
                .Columns("PHYS_STYLE_COLORS").DataType = GetType(System.Int64)
                .Columns("BOOK_STYLE_COLORS").DataType = GetType(System.Int64)
                .Columns("PHYS_UNITS").DataType = GetType(System.Int64)
                .Columns("BOOK_UNITS").DataType = GetType(System.Int64)
                .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS_UNITS,0) - ISNULL(BOOK_UNITS,0)")
                .Columns.Add("VARIANCE_COST", GetType(System.Decimal), "ISNULL(PHYS_VALUE,0) - ISNULL(BOOK_VALUE,0)")
            End With

            ASCMAIN1.sql = "" _
                & "Select X.WHSE_CODE, X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
                & ", ICTSTYC1.STYLE_COST_FIFO, ICTSTYL1.STYLE_DESC, MAX(X.TICKET_NO) TICKET_NO, MAX(ICTPHYC1.INIT_OPER) INIT_OPER" & vbCrLf _
                & ", Sum (X.PHYS_UNITS) PHYS_UNITS, Sum (X.BOOK_UNITS) BOOK_UNITS" & vbCrLf _
                & " from ICTSTYC1, ICTSTYL1, ICTPHYC1, (" & vbCrLf _
                & "Select ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE" & vbCrLf _
                & ", SUM (ICTPHYC2.COUNT_CTNS * ICTPHYC2.CARTON_PACK_QTY + ICTPHYC2.COUNT_LOOSE) PHYS_UNITS, 0 BOOK_UNITS" & vbCrLf _
                & ", MAX (ICTPHYC2.TICKET_NO) TICKET_NO" & vbCrLf _
                & " from ICTPHYC1,ICTPHYC2" & vbCrLf _
                & " where ICTPHYC1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and ICTPHYC2.WHSE_CODE = ICTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and ICTPHYC2.TICKET_NO = ICTPHYC1.TICKET_NO" & RGI_CODE & vbCrLf _
                & " group by ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE" & vbCrLf _
                & ", 0 PHYS_UNITS, SUM (WHTLOCB0.LOCATION_QTY - WHTLOCB0.BOOK_INVTY_ADJ) BOOK_UNITS" & vbCrLf _
                & ", NULL TICKET_NO" & vbCrLf _
                & " from WHTLOCB0" & vbCrLf _
                & " where WHTLOCB0.WHSE_CODE = :PARM1" & vbCrLf _
                & " group by WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE" & vbCrLf _
                & ") X" & vbCrLf _
                & " where ICTSTYC1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTPHYC1.WHSE_CODE (+) = X.WHSE_CODE" & vbCrLf _
                & "   and ICTPHYC1.TICKET_NO (+) = X.TICKET_NO" & vbCrLf _
                & " group by X.WHSE_CODE, X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
                & ", ICTSTYC1.STYLE_COST_FIFO, ICTSTYL1.STYLE_DESC"

            Create_TDA(.Tables.Add, "ICTPHYCR", "**", 0, False, "V", 4)
            With .Tables("ICTPHYCR")
                .Columns("PHYS_UNITS").DataType = GetType(System.Int64)
                .Columns("BOOK_UNITS").DataType = GetType(System.Int64)
                .Columns.Add("PHYS_VALUE", GetType(System.Decimal), "PHYS_UNITS * STYLE_COST_FIFO")
                .Columns.Add("BOOK_VALUE", GetType(System.Decimal), "BOOK_UNITS * STYLE_COST_FIFO")
                .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS_UNITS,0) - ISNULL(BOOK_UNITS,0)")
                .Columns.Add("VARIANCE_COST", GetType(System.Decimal), "ISNULL(PHYS_VALUE,0) - ISNULL(BOOK_VALUE,0)")
            End With

            ASCMAIN1.sql = "select X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE, ICTSTYV1.PO_COST, SUM(BOOK) BOOK, SUM(PHYS) PHYS " & vbCrLf _
                & " from (" & vbCrLf _
                & " Select LOCATION_CODE, STYLE_CODE, COLOR_CODE,  (WHTLOCB0.LOCATION_QTY - WHTLOCB0.BOOK_INVTY_ADJ) BOOK, 0 PHYS " & vbCrLf _
                & " from WHTLOCB0 " & vbCrLf _
                & " where WHSE_CODE = :PARM1" & vbCrLf _
                & " and LOCATION_CODE = :PARM2 " & vbCrLf _
                & " and nvl(LOCATION_QTY,0) <> 0 " & vbCrLf _
                & " union " & vbCrLf _
                & " select ICTPHYC1.LOCATION_CODE, ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, 0 BOOK, sum(NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS " & vbCrLf _
                & " from ICTPHYC1, ICTPHYC2 " & vbCrLf _
                & " where ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO " & RGI_CODE & vbCrLf _
                & " and ICTPHYC1.WHSE_CODE = :PARM1 " & vbCrLf _
                & " and ICTPHYC1.LOCATION_CODE = :PARM2 " & vbCrLf _
                & " Group by ICTPHYC1.LOCATION_CODE, ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE " & vbCrLf _
                & " ) X,  ICTSTYV1 " & vbCrLf _
                & " Where  X.STYLE_CODE = ICTSTYV1.STYLE_CODE " & vbCrLf _
                & " group by X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE, ICTSTYV1.PO_COST "
            Create_TDA(.Tables.Add, "WHTLOCBV", "**", 0, False, "VV", 3)
            With .Tables("WHTLOCBV")
                .Columns("PHYS").DataType = GetType(System.Int64)
                .Columns("BOOK").DataType = GetType(System.Int64)
                .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS,0) - ISNULL(BOOK,0)")
                .Columns.Add("AMT_VARIANCE", GetType(System.Double), "iif(VARIANCE < 0,-1,1) * (ISNULL(PHYS,0) - ISNULL(BOOK,0)) * ISNULL(PO_COST,0)")
            End With

        End With

        Fill_Records("WHTLOCM1")
        '  Fill_Records("ICTCLAS1")

        ' cbe.DataSource = ASCDATA1.GetDataTable("Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        'ASCMAIN1.USER_SECURITY_CODEs.Contains("XX")

        Dim rows() As DataRow = ASCDATA1.GetDataTable("SELECT *  FROM WHTLPRT1").Select("")
        For Each row As DataRow In rows
            cbxLabelPrinter.Items.Add(row.Item("LABEL_PRINTER_ID"))
        Next
        If cbxLabelPrinter.Items.Count > 0 Then
            cbxLabelPrinter.SelectedIndex = 0
        End If

        grdWHTLOCB0.DataSource = dst.Tables("WHTLOCB0")
        grdICTPHYC2.DataSource = dst.Tables("ICTPHYC2")
        grdICTPHYCI.DataSource = dst.Tables("ICTPHYCI")
        grdICTPHYCX.DataSource = dst.Tables("ICTPHYCX")
        grdICTPHYCV.DataSource = dst.Tables("ICTPHYCV")
        grdICTPHYCL.DataSource = dst.Tables("ICTPHYCL")
        grdICTPHYCR.DataSource = dst.Tables("ICTPHYCR")
        grdWHTLOCBV.DataSource = dst.Tables("WHTLOCBV")

        Create_Summary(grdICTPHYCX, "TICKET_NO", "Count")
        Create_Summary(grdICTPHYCX, "VERIFIED_OPER", "Count")
        Create_Summary(grdICTPHYCX, "TOTAL_COUNT")
        Create_Summary(grdICTPHYCX, "UNIT_VARIANCE")
        Create_Summary(grdICTPHYCX, "ABSOLUTE_VARIANCE")

        Create_Summary(grdICTPHYCL, "LOCATION_CODE", "Count")
        Create_Summary(grdICTPHYCL, New String() {"PHYS_UNITS", "BOOK_UNITS", "PHYS_VALUE", "BOOK_VALUE", "VARIANCE", "VARIANCE_COST"})

        Create_Summary(grdICTPHYCR, "LOCATION_CODE", "Count")
        Create_Summary(grdICTPHYCR, New String() {"PHYS_UNITS", "BOOK_UNITS", "PHYS_VALUE", "BOOK_VALUE", "VARIANCE", "VARIANCE_COST"})

        Create_Summary(grdICTPHYCV, "STYLE_CODE", "Count")
        Create_Summary(grdICTPHYCV, New String() {"BOOK", "PHYS", "VARIANCE", "VARIANCE_COST"})

        Create_Summary(grdICTPHYC2, "TICKET_LNO", "Count")
        Create_Summary(grdICTPHYC2, "COUNT_CTNS")
        Create_Summary(grdICTPHYC2, "TOTAL_COUNT")

        Create_Summary(grdICTPHYCI, "TICKET_NO", "Count")
        Create_Summary(grdICTPHYCI, "COUNT_CTNS")
        Create_Summary(grdICTPHYCI, "TOTAL_COUNT")

        Create_Summary(grdWHTLOCB0, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCB0, "LOCATION_QTY")

        Create_Summary(grdWHTLOCBV, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCBV, "BOOK")
        Create_Summary(grdWHTLOCBV, "PHYS")
        Create_Summary(grdWHTLOCBV, "VARIANCE")
        Create_Summary(grdWHTLOCBV, "AMT_VARIANCE")


        With grdICTPHYC2.DisplayLayout.Bands("ICTPHYC2")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If gcol.Key = "STYLE_CODE" Or gcol.Key = "COUNT_CTNS" Or gcol.Key = "COUNT_LOOSE" Or gcol.Key = "CARTON_PACK_QTY" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "COUNT_CTNS" Or gcol.Key = "COUNT_CTNS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    'ElseIf gcol.Key = "PHYS" Or gcol.Key = "BOOK" Then
                    '    gcol.Header.Appearance.BackColor = Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If

            Next
            .Columns("TICKET_LNO").Header.Fixed = True
            .Columns("STYLE_CODE").Header.Fixed = True
        End With

        With grdICTPHYCV.DisplayLayout.Bands("ICTPHYCV")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "VARIANCE" Or gcol.Key = "VARIANCE_COST" Or gcol.Key = "STYLE_COST_VAR" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                ElseIf gcol.Key = "PHYS" Or gcol.Key = "BOOK" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
            .Columns("STYLE_CODE").Header.Fixed = True
            If Not ASCMAIN1.CLIENT = "RGI" Then
                .Columns("STYLE_COST_VAR").Hidden = True
                .Columns("COUNTED_LOCS").Hidden = True
                .Columns("BOOKED_LOCS").Hidden = True
            End If
        End With

        With grdICTPHYCX.DisplayLayout.Bands("ICTPHYCX")
            .Columns("TICKET_NO").Header.Fixed = True
            .Columns("SELECTED").Header.Fixed = True
            '.Columns("SELECTED"). = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If gcol.Key = "SELECTED" Or gcol.Key = "TICKET_NO" Or gcol.Key = "COUNTED_BY" Or gcol.Key = "LOCATION_CODE" Or gcol.Key = "VERIFIED_DATE" Or gcol.Key = "VERIFIED_OPER" Then
                    gcol.Header.Appearance.BackColor2 = Color.Pink
                ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "SC" Or gcol.Key = "TOTAL_COUNT" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "UNIT_VARIANCE" Or gcol.Key = "ABSOLUTE_VARIANCE" Or gcol.Key = "LAST_ACTIVITY" Then
                    gcol.Header.Appearance.BackColor2 = Color.Red
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
            .Columns("SC").Hidden = (ASCMAIN1.CLIENT = "NYA")
            .Columns("STYLE_CODE").Hidden = Not (ASCMAIN1.CLIENT = "NYA")
        End With
        grdAttention_app.BackColor = Drawing.Color.Yellow
        grdWARNING_app.ForeColor = Drawing.Color.Red

        With grdICTPHYCI.DisplayLayout.Bands("ICTPHYCI")

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.Turquoise
            Next
        End With


        With grdWHTLOCB0.DisplayLayout.Bands("WHTLOCB0")

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.Yellow
            Next
        End With

        'With grdICTPHYCX.DisplayLayout.Bands("ICTPHYCX")
        '    .Columns("TICKET_NO").Header.Fixed = True
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        gcol.Header.Appearance.BackColor = Color.White
        '        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        If gcol.Key = "TICKET_NO" Or gcol.Key = "COUNTED_BY" Or gcol.Key = "LOCATION_CODE" Then
        '            gcol.Header.Appearance.BackColor = Color.Pink
        '        ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "TOTAL_COUNT" Then
        '            gcol.Header.Appearance.BackColor = Color.LightGreen
        '        Else
        '            gcol.Header.Appearance.BackColor = Color.LightGray
        '        End If
        '    Next
        'End With

        With grdICTPHYCL.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key.StartsWith("PHYS") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key.StartsWith("BOOK") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key.StartsWith("V") Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("LOCATION_CODE").Header.Fixed = True
        End With


        With grdICTPHYCR.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key.StartsWith("PHYS") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key.StartsWith("BOOK") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key.StartsWith("V") Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("LOCATION_CODE").Header.Fixed = True
        End With

        'ASCMAIN1.Add_Value_List(grdICTPHYCX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
        'ASCMAIN1.Add_Value_List(grdICTPHYCV, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grpHeader.Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("WHSE_CODE")

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify a Valid Warehouse"
                Else
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    ElseIf rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        'ElseIf rowICTWHSE1.Item("LP_CODE") & "" <> "" And Not (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
                        '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Counts Entry Allowed"
                    ElseIf rowICTWHSE1.Item("WHSE_PHYS_STATUS") & "" <> "C" Then
                        EMsg &= vbCr & "Warehouse has not been Initialized for Physical Counts Entry"
                    End If
                End If

                If Absx1.txtFor("TICKET_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Ticket"
                Else
                    Dim rowICTPHYC1 As DataRow = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("TICKET_NO").Text})
                    If rowICTPHYC1 IsNot Nothing Then
                        Click_Command("View")
                        Exit Sub
                    Else
                        If Not ASCMAIN1.Logical_Lock("ICTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & Absx1.txtFor("TICKET_NO").Text) Then
                            Exit Sub
                        End If
                    End If
                End If



            Case "Edit"
                If optMode.Value = "U" Or optMode.Value = "V" Then
                    SelUser = grdICTPHYCX.ActiveRow.Cells("INIT_OPER").Value
                    'If Not ASCMAIN1.Logical_Lock("ICTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & SelUser) Then
                    '    Exit Sub
                    'End If

                Else
                    Dim rowICTPHYC1 As DataRow = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("TICKET_NO").Text})
                    If rowICTPHYC1 Is Nothing Then
                        EMsg &= "Ticket is not on File"
                    Else

                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                        If IsNothing(rowICTWHSE1) Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                        ElseIf rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Active"
                            'ElseIf rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                            '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Counts Entry Allowed"
                        ElseIf rowICTWHSE1.Item("WHSE_PHYS_STATUS") & "" <> "C" And optMode.Value <> "U" Then
                            EMsg &= vbCr & "Warehouse has not been Initialized for Physical Counts Entry"
                        End If

                        If Not ASCMAIN1.Logical_Lock("ICTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & Absx1.txtFor("TICKET_NO").Text) Then
                            Exit Sub
                        End If
                    End If
                End If
            Case "View"
                If Absx1.txtFor("TICKET_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTPHYC1 = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("TICKET_NO").Text})
                    If rowICTPHYC1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("TICKET_NO").Text & " on File"
                    End If
                End If

            Case "Update"
                If Not (optMode.Value = "U" Or optMode.Value = "V") Then
                    If location_support Then
                        If Absx1.txtFor("LOCATION_CODE").Text = "" Then
                            EMsg &= vbCr & "You Must Specify a Location"
                        Else
                            Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("LOCATION_CODE").Text})
                            If rowWHTLOCM1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Value Specified for Location"
                            End If
                        End If

                        If Absx1.txtFor("COUNT_BY").Text = "" Then
                            EMsg &= vbCr & "You Must enter either Notes or Initials of the person who did the count"
                        End If
                    End If

                    If grdICTPHYC2.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Details Entered"
                    Else
                        'For Each rowICTPHYC2 As DataRow In dst.Tables("ICTPHYC2").Select("", "", DataViewRowState.CurrentRows)
                        '    If rowICTPHYC2.Item("COST_CATGY_CODE") & "" = "" Then
                        '        EMsg &= vbCr & "Unable to determine Cost Category for " & rowICTPHYC2.Item("STYLE_CODE") & ""
                        '    End If
                        '    If rowICTPHYC2.Item("PROD_CODE") & "" = "" Then
                        '        EMsg &= vbCr & "Unable to determine Product Code for " & rowICTPHYC2.Item("STYLE_CODE") & ""
                        '    End If
                        'Next
                    End If
                End If
            Case "Delete"
                If MessageBox.Show("Are you sure you want to Delete this Entry?", "Confirm Deletion", _
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                If optMode.Value = "U" Or optMode.Value = "V" Then
                    Verify_Counts()
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)


            Case "By Ticket"
                Print_Counts("T")
            Case "By Location"
                Print_Counts("L")
            Case "By Style"
                Print_Counts("S")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If


                    If ScreenMode And (EntryMode <> "N" And EntryMode <> "E") Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Delete").Visible = (ScreenMode And EntryMode = "E" And optMode.Value <> "U")
                End With

                '  .Groups("Variances").Visible = ScreenMode And (EntryMode = "V")
                .Groups("Count Reports").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode And (optMode.Value <> "U" And optMode.Value <> "V")
        SplitContainer3.Visible = ScreenMode And (optMode.Value = "U" Or optMode.Value = "V")
        grpHeader.Visible = ScreenMode

        tab0.Visible = Not ScreenMode
        UltraExplorerBar1.Groups("Mode").Visible = Not ScreenMode

        If ScreenMode Then
            Absx1.txtFor("LOCATION_CODE").Visible = location_support
            Absx1.txtFor("LOCATION_DESC").Visible = location_support
            lblLOCATION_CODE.Visible = location_support

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            ' Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Or EntryMode = "E" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPHYC2, grdICTPHYCX}
                    If optMode.Value = "U" Or optMode.Value = "V" Then
                        With grd.DisplayLayout.Override
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.False
                            If grd.Name = "grdICTPHYCX" Then
                                .AllowUpdate = DefaultableBoolean.True
                            End If
                        End With
                    Else
                        With grd.DisplayLayout.Override
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            .AllowDelete = IIf(grd.Name = "grdICTPHYC2", DefaultableBoolean.True, DefaultableBoolean.False)
                            .AllowUpdate = DefaultableBoolean.True
                        End With
                    End If
                Next
                With grdICTPHYC2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("COUNT_CTNS").CellAppearance.BackColor = Color.LightYellow
                    .Columns("COUNT_LOOSE").CellAppearance.BackColor = Color.LightYellow
                End With
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPHYC2, grdICTPHYCI, grdICTPHYCX}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                'With grdICTPHYC2.DisplayLayout.Bands(0)
                '    .Columns("STYLE_CODE").CellAppearance.BackColor = Color.Empty
                '    .Columns("COUNT_CTNS").CellAppearance.BackColor = Color.Empty
                'End With
            End If

            If grdICTPHYC2.ActiveRow Is Nothing Then
                Setup_ICTPHYC2("", "")
            End If
            Setup_WHTLOCB0(False)

            Absx1.txtFor("LOCATION_CODE").Focus()
            Setup_tab0()

            If optMode.Value = "U" Or optMode.Value = "V" Then
                grdICTPHYC2.Parent = SplitContainer4.Panel1
                grdICTPHYCX.Parent = SplitContainer3.Panel1
                'grdWHTLOCB0.Parent = SplitContainer4.Panel2
            Else
                grdICTPHYC2.Parent = splICTPHYC2.Panel1
                'grdWHTLOCB0.Parent = splItemDetails.Panel2
            End If
            Sort_grdColumns(grdICTPHYC2, "STYLE_CODE, COLOR_CODE")
            Sort_grdColumns(grdWHTLOCBV, "STYLE_CODE, COLOR_CODE")
        Else
            Clear_Record()
            tab0.SelectedTab = tab0.Tabs("Tickets")
            grdICTPHYCX.Parent = tab0.Tabs("Tickets").TabPage
        End If

        With grdICTPHYCX.DisplayLayout.Bands("ICTPHYCX")
            .Columns("SELECTED").Hidden = Not ScreenMode
        End With

        btnEmpty.Visible = ASCMAIN1.Running_in_VS

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTPHYC1", "ICTPHYC2", "ICTPHYCI"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
        Absx1.txtFor("TICKET_NO").Text = ""

        If WHSE_CODE = "" Then
            Absx1.txtFor("WHSE_CODE").Focus()
        Else
            Absx1.txtFor("TICKET_NO").Focus()
        End If

        Refresh_Documents()
        Setup_tab0()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        TICKET_NO = Absx1.txtFor("TICKET_NO").Text

        If EntryMode = "N" Then
            rowICTPHYC1 = dst.Tables("ICTPHYC1").NewRow
            rowICTPHYC1.Item("WHSE_CODE") = WHSE_CODE
            rowICTPHYC1.Item("TICKET_NO") = TICKET_NO ' ASCMAIN1.Next_Control_No("ICTPHYC1.TICKET_NO")

            rowICTPHYC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTPHYC1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTPHYC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTPHYC1.Item("LAST_DATE") = DATETIME_STAMP
            dst.Tables("ICTPHYC1").Rows.Add(rowICTPHYC1)
        Else
            Fill_Record("ICTPHYC1", New String() {WHSE_CODE, TICKET_NO})
            dst.AcceptChanges()
        End If

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        With grdICTPHYC2.DisplayLayout.Bands(0)
            If ASCMAIN1.CLIENT = "RGI" Then
                .Columns("BAR_CODE").Hidden = True
            Else
                .Columns("BAR_CODE").Hidden = Not location_support
            End If
        End With

        Fill_Records("ICTPHYC2", New String() {WHSE_CODE, TICKET_NO})
        Dim dvwC2 As DataView = DirectCast(grdICTPHYC2.DataSource, DataTable).DefaultView
        dvwC2.RowFilter = "STATUS IS NULL"
        If EntryMode = "E" And (optMode.Value = "U" Or optMode.Value = "V") Then
            Dim dvw As DataView = DirectCast(grdICTPHYCX.DataSource, DataTable).DefaultView
            dvw.RowFilter = "INIT_OPER = '" & SelUser & "' " & IIf(optMode.Value = "U", "and (VERIFIED_OPER IS NULL or LAST_ACTIVITY > LAST_DATE)", "and VERIFIED_OPER IS NOT NULL")
            grdICTPHYCX.Text = "Verify Counts for User " & SelUser
            Set_ICTPHYCX()
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("ICTPHYC1")
        Update_Record_TDA("ICTPHYC2")
        CommitTrans("Update Complete")
    End Sub

    Sub Verify_Counts()
        BeginTrans()

        ASCMAIN1.sql = "Update ICTPHYC1 " & vbCrLf _
            & " Set ICTPHYC1.VERIFIED_OPER = '" & ASCMAIN1.USER_ID & "', ICTPHYC1.VERIFIED_DATE = SYSDATE " & vbCrLf _
            & " Where ICTPHYC1.WHSE_CODE = '" & WHSE_CODE & "' and ICTPHYC1.TICKET_NO = :PARM1"

        For Each row As DataRow In dst.Tables("ICTPHYCX").Select("SELECTED = '1'")
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", row.Item("TICKET_NO"))
        Next

        CommitTrans("Update Complete, Records Verified")
    End Sub

    Sub printRecount(Location As String)

        BeginTrans()
        ASCMAIN1.sql = "Insert into ICTPHYC1_RECNT Values(:PARM1)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", Location)
        CommitTrans()

        Using ipp As New nsoftware.IPWorks.Ipport
            ipp.RuntimeLicense = "31504E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004B4857525953375A4A5A375A0000"
            ipp.Connect("192.168.110.223", "4444")
            'ipp.Connect("192.168.120.25", "4444")
            Dim data As String '= "upc123" ' & vbCrLf a new line is needed to send the data across
            Try
                data = cbxLabelPrinter.SelectedItem 'Printer
                data &= "|" & Location
                ipp.SendLine(data)

            Catch ex As Exception

            End Try

            ipp.Disconnect()
        End Using

    End Sub

    Sub Delete_Record()
        BeginTrans()

        Delete_Records("ICTPHYC1")
        Delete_Records("ICTPHYC2")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where WHSE_CODE = '" & WHSE_CODE & "' and TICKET_NO = '" & TICKET_NO & "'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTPHYCX, "SSB", "Show Filter", "Show GroupBox", "Recount Location")
        Load_Popup_Menu(grdICTPHYC2, "BS", "Style Status Inquiry", "Show Old Counts")
        Load_Popup_Menu(grdICTPHYCV, "SSB", "Show Filter", "Show GroupBox", "Style Status Inquiry")
        Load_Popup_Menu(grdICTPHYCL, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTPHYCR, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdWHTLOCB0, "BB", "Location Inquiry", "Show 0's")
        Load_Popup_Menu(grdWHTLOCBV, "SB", "Show Filter", "Style Status Inquiry")
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

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdICTPHYCX"
                tlb_btn = DirectCast(tlb_pop.Tools("Recount Location"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E")

        End Select
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdWHTLOCB0"
                    '  tlb_sbt = DirectCast(tlb.Tools("Show 0s"), UltraWinToolbars.StateButtonTool)

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)
            Case "Show 0's"
                '  tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Setup_WHTLOCB0(True)

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Location Inquiry"
                'Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                'If rowICTSTYL1 IsNot Nothing Then
                '    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                'End If

            Case "Show Old Counts"
                Set_ICTPHYC2_Filter()

            Case "Recount Location"
                Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Value
                ASCMAIN1.sql = "Select count(1) from ICTPHYC1_RECNT where LOCATION_CODE = :PARM1"
                Dim recount = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", LOCATION_CODE)
                If recount > 0 Then
                    If MessageBox.Show("Location " & LOCATION_CODE & " has been re-counted " & recount & " times, recount again?", "Re-count Location",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If
                printRecount(LOCATION_CODE)
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Refresh_Documents()
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "TICKET_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("WHSE_CODE").Text <> "" Then
                        Dim row As DataRow = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("WHSE_CODE").Text})
                        If row IsNot Nothing Then
                            Click_Command("View", e)
                        Else
                            Click_Command("New", e)
                        End If
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Refresh_Documents()
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "TICKET_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                Refresh_Documents()
            Case "COUNT_BY"
                grdICTPHYC2.Focus()
                If grdICTPHYC2.ActiveRow Is Nothing Then
                    If grdICTPHYC2.Rows.Count = 0 Then
                        grdICTPHYC2.DisplayLayout.Bands(0).AddNew()
                    End If
                End If
                If grdICTPHYC2.ActiveRow IsNot Nothing Then
                    grdICTPHYC2.ActiveCell = grdICTPHYC2.ActiveRow.Cells("STYLE_CODE")
                End If

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTPHYC2"

    Private Sub grdICTPHYC2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPHYC2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

                grdCodeDesc(grdICTPHYC2, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE As String = e.Cell.Value

                    'Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = cdr.Item("CARTON_PACK_QTY")

                    'Dim COST_CATGY_CODE As String = cdr.Item("COST_CATGY_CODE") & ""
                    'Dim PROD_CODE As String = cdr.Item("PROD_CODE") & ""
                    'Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")
                    'e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE
                    'e.Cell.Row.Cells("PROD_CODE").Value = PROD_CODE
                    'e.Cell.Row.Cells("STYLE_COST").Value = STYLE_COST


                    Dim COLOR_CODE As String = ""

                    ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
                    Dim rowICTSTYC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rowICTSTYC1s.Length = 1 Then
                        COLOR_CODE = rowICTSTYC1s(0).Item("COLOR_CODE")
                        e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODE
                    End If

                    Setup_ICTPHYC2(STYLE_CODE, COLOR_CODE)
                Else
                    grdICTPHYC2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "COLOR_CODE"
                'grdCodeDesc(grdICTPHYC2, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
                '' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE COLOR_DESC
                'If cdr IsNot Nothing Then
                '    e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")
                'End If
                Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = e.Cell.Value
                Setup_ICTPHYC2(STYLE_CODE, COLOR_CODE)

            Case "COUNT_CTNS"

        End Select
    End Sub

    Private Sub grdICTPHYC2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPHYC2.AfterExitEditMode
        'Select Case grdICTPHYC2.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdICTPHYC2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPHYC2.AfterRowActivate
        With grdICTPHYC2.DisplayLayout.Bands(0)
            If grdICTPHYC2.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTPHYC2.ActiveCell = grdICTPHYC2.ActiveRow.Cells("STYLE_CODE")
                grdICTPHYC2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        Dim STYLE_CODE As String = grdICTPHYC2.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = grdICTPHYC2.ActiveRow.Cells("COLOR_CODE").Value & ""
        Setup_ICTPHYC2(STYLE_CODE, COLOR_CODE)
        'If EntryMode = "V" Then
        '    Show_Variances()
        'End If
    End Sub

    Private Sub grdICTPHYC2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPHYC2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTPHYC2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTPHYC2.AfterRowUpdate
        DisplayTotals()
    End Sub

    Private Sub grdICTPHYC2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTPHYC2.BeforeExitEditMode
        If grdICTPHYC2.ActiveCell Is Nothing Then Exit Sub
        With grdICTPHYC2.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTSTYL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If
            End Select
        End With
    End Sub

    Private Sub grdICTPHYC2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTPHYC2.BeforeRowUpdate
        With grdICTPHYC2
            If e.Row.Cells("STYLE_CODE").Text = "" Then
                '                MsgBox("Missing Value for Style Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

            End If

            If Val(e.Row.Cells("COUNT_CTNS").Value & "") = 0 And Val(e.Row.Cells("COUNT_LOOSE").Value & "") = 0 Then
                'MsgBox("Invalid Value entered for Count", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("TICKET_NO").Text = "" Then
                    .ActiveRow.Cells("WHSE_CODE").Value = WHSE_CODE
                    .ActiveRow.Cells("TICKET_NO").Value = Absx1.CtlFor("TICKET_NO").Text
                    .ActiveRow.Cells("TICKET_LNO").Value = Val(dst.Tables("ICTPHYC2").Compute("Max(TICKET_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdICTPHYC2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPHYC2.ClickCellButton

        If grdICTPHYC2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                'Case "LOCATION_CODE"
                '    sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdICTPHYC2, sql_where, False)

    End Sub

    Private Sub grdICTPHYC2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTPHYC2.Error
        grdICTPHYC2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        'Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTPHYC2").Compute("SUM(LINE_COSTS)", "") & "")
        'Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

#Region "grdICTPHYCX"


    Private Sub grdICTPHYCX_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTPHYCX.BeforeExitEditMode
        If grdICTPHYCX.ActiveCell Is Nothing Then Exit Sub
        If Not grdICTPHYCX.ActiveRow.IsDataRow Then Exit Sub
        With grdICTPHYCX.ActiveCell
            Select Case .Column.Key
                Case "SELECTED"
                    If .Text = "0" And grdICTPHYCX.ActiveRow.Cells("VERIFIED_OPER").Value & "" <> "" Then
                        e.Cancel = True
                    End If
                Case Else
                    e.Cancel = True
            End Select
        End With
    End Sub


    Private Sub grdICTPHYCX_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdICTPHYCX.BeforeRowUpdate
        'Null()
    End Sub

    Private Sub grdICTPHYCX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTPHYCX.AfterRowActivate
        If Not grdICTPHYCX.ActiveRow.IsDataRow Then Exit Sub
        Set_ICTPHYCX()


    End Sub

    Sub Set_ICTPHYCX()
        If EntryMode = "V" Or EntryMode = "" Then Exit Sub
        If grdICTPHYCX.ActiveRow IsNot Nothing AndAlso grdICTPHYCX.ActiveRow.IsDataRow Then
            If Not ASCMAIN1.Logical_Lock("ICTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & grdICTPHYCX.ActiveRow.Cells("LOCATION_CODE").Text) Then
                Exit Sub
            End If
            Dim rowTICKET_NO As String = grdICTPHYCX.ActiveRow.Cells("TICKET_NO").Text
            Fill_Records("ICTPHYC2", New String() {WHSE_CODE, rowTICKET_NO})
            Set_ICTPHYC2_Filter()
            grdICTPHYC2.Text = "Details for Ticket " & rowTICKET_NO & ", Location " & grdICTPHYCX.ActiveRow.Cells("LOCATION_CODE").Text
            Fill_Records("WHTLOCBV", New String() {WHSE_CODE, grdICTPHYCX.ActiveRow.Cells("LOCATION_CODE").Text})
            grdWHTLOCBV.Text = "Variances for Location " & grdICTPHYCX.ActiveRow.Cells("LOCATION_CODE").Text
        End If
    End Sub

    Sub Set_ICTPHYC2_Filter()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Old Counts"), UltraWinToolbars.StateButtonTool)
        Dim dvw As DataView = DirectCast(grdICTPHYC2.DataSource, DataTable).DefaultView
        If Not tlb_sbt.Checked Then
            dvw.RowFilter = "STATUS IS NULL"
        Else
            dvw.RowFilter = ""
        End If
    End Sub
    Private Sub grdICTPHYCX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTPHYCX.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("VERIFIED_OPER").Value & "" = "" Then
                e.Row.Cells("VERIFIED_OPER").Appearance = grdAttention_app
            End If
            If Not (IsDBNull(e.Row.Cells("LAST_DATE").Value)) AndAlso e.Row.Cells("LAST_ACTIVITY").Value > e.Row.Cells("LAST_DATE").Value Then
                e.Row.Cells("LOCATION_CODE").Appearance = grdWARNING_app
                e.Row.Cells("LOCATION_CODE").ToolTipText = "Location had activity after the count."
                e.Row.Cells("LAST_ACTIVITY").Appearance = grdWARNING_app
                e.Row.Cells("LAST_ACTIVITY").ToolTipText = "Location had activity after the count."
            End If
        End If
    End Sub
    Private Sub grdICTPHYCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPHYCX.DoubleClickRow
        If EntryMode = "E" Then Exit Sub

        If e.Row.IsDataRow Then
            Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Text
            If optMode.Value = "U" Or optMode.Value = "V" Then
                Click_Command("Edit")
            Else
                Click_Command("View")
            End If

        End If
    End Sub

#End Region
    Private Sub grdICTPHYCV_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTPHYCV.AfterRowActivate
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        If grdICTPHYCV.ActiveRow IsNot Nothing AndAlso grdICTPHYCV.ActiveRow.IsDataRow Then
            STYLE_CODE = grdICTPHYCV.ActiveRow.Cells("STYLE_CODE").Value
            COLOR_CODE = grdICTPHYCV.ActiveRow.Cells("COLOR_CODE").Value
        End If
        Setup_ICTPHYC2(STYLE_CODE, COLOR_CODE)
    End Sub

    Private Sub grdICTPHYCV_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPHYCV.DoubleClickRow
        'If e.Row.IsDataRow Then
        '    Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Text
        '    Click_Command("View")
        'End If
    End Sub
    Private Sub optMode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMode.ValueChanged
        Show_Tickets()
    End Sub

    Private Sub Show_Tickets()
        If SELECTION_NO = 0 Then Exit Sub
        If EntryMode = "E" Then Exit Sub
        Dim dvw As DataView = DirectCast(grdICTPHYCX.DataSource, DataTable).DefaultView
        If optMode.Value = "A" Then
            dvw.RowFilter = ""
            grdICTPHYCX.Text = "All Physical Counts for Warehouse " & WHSE_CODE
        ElseIf optMode.Value = "U" Then
            dvw.RowFilter = "VERIFIED_OPER IS NULL or LAST_ACTIVITY > LAST_DATE"
            grdICTPHYCX.Text = "Unverified Physical Counts for Warehouse " & WHSE_CODE
        ElseIf optMode.Value = "V" Then
            dvw.RowFilter = "VERIFIED_OPER IS NOT NULL"
            grdICTPHYCX.Text = "Verified Physical Counts for Warehouse " & WHSE_CODE
        Else ' view Dirty tickets
            dvw.RowFilter = "LAST_ACTIVITY > LAST_DATE"
            grdICTPHYCX.Text = "Active since Physical Counts for Warehouse " & WHSE_CODE
        End If
        UltraLabel10.Visible = Not (optMode.Value = "U" Or optMode.Value = "V")
        Absx1.txtFor("TICKET_NO").Visible = Not (optMode.Value = "U" Or optMode.Value = "V")
    End Sub


    Private Sub optVariances_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optVariances.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Variances()
    End Sub

    Sub Show_Variances()
        Dim dvw As DataView = DirectCast(grdICTPHYCV.DataSource, DataTable).DefaultView
        If optVariances.Value = "A" Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "VARIANCE <> 0"
        End If
    End Sub

    Sub Show_Locations()
        'Dim dvw As DataView = DirectCast(grdICTPHYCV.DataSource, DataTable).DefaultView
        'If optVariances.Value = "A" Then
        '    dvw.RowFilter = ""
        'Else
        '    dvw.RowFilter = "VARIANCE <> 0"
        'End If
    End Sub

    Sub Show_Location_Style_Colors()
        'Dim dvw As DataView = DirectCast(grdICTPHYCV.DataSource, DataTable).DefaultView
        'If optVariances.Value = "A" Then
        '    dvw.RowFilter = ""
        'Else
        '    dvw.RowFilter = "VARIANCE <> 0"
        'End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If Not IsNothing(rowICTWHSE1) Then
            location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        End If
        Fill_Records("ICTPHYCX", WHSE_CODE)
        Show_Tickets()
        Sort_grdColumns(grdICTPHYCX, "TICKET_NO")
        tab0.SelectedTab = tab0.Tabs("Tickets")
    End Sub

    Sub Setup_ICTPHYC2(STYLE_CODE As String, COLOR_CODE As String)
        If STYLE_CODE = "" Then
            splItemDetails.Visible = False
        Else
            splItemDetails.Visible = True
            Fill_Records("ICTPHYCI", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            grdICTPHYCI.Text = "Tickets with Style " & STYLE_CODE & "-" & COLOR_CODE
            If location_support Then
                Fill_Records("WHTLOCB0", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
                grdWHTLOCB0.Text = "Book Inventory by Location for Style " & STYLE_CODE & "-" & COLOR_CODE
                Setup_WHTLOCB0(False)
            End If
        End If

    End Sub

    Sub Setup_WHTLOCB0(Show_0s As Boolean)
        Dim dvw As DataView = DirectCast(grdWHTLOCB0.DataSource, DataTable).DefaultView
        If Show_0s Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "LOCATION_QTY <> 0"
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()

    End Sub

    Sub Setup_tab0()
        If tab0.SelectedTab.Key = "Variances" Then
            If Load_Variances() Then
                splItemDetails.Parent = splICTPHYCV.Panel2
                Show_Variances()
            Else
                tab0.SelectedTab = tab0.Tabs("Tickets")
            End If
        ElseIf tab0.SelectedTab.Key = "Locations" Then
            If Load_Locations() Then
                Show_Locations()
            Else
                tab0.SelectedTab = tab0.Tabs("Tickets")
            End If
        ElseIf tab0.SelectedTab.Key = "Location/Style/Color" Then
            If Load_Location_Style_Colors() Then
                Show_Location_Style_colors()
            Else
                tab0.SelectedTab = tab0.Tabs("Tickets")
            End If
        Else
            splItemDetails.Parent = splICTPHYC2.Panel2
        End If

        UltraExplorerBar1.Groups("Variances").Visible = (tab0.SelectedTab.Key = "Variances") And Not ScreenMode
        UltraExplorerBar1.Groups("Mode").Visible = (tab0.SelectedTab.Key = "Tickets") And Not ScreenMode
    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        Load_Variances()
    End Sub

    Function Load_Variances() As Boolean

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            MsgBox("You Must Pick a Valid Warehouse Code", MsgBoxStyle.OkOnly, "Cannot Show Variance")
            Return False
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Compiling Variances")

        If rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
            Fill_Records("ICTPHYCV", New String() {WHSE_CODE, "", WHSE_CODE, ASCMAIN1.CYP})
        Else
            Fill_Records("ICTPHYCV", New String() {WHSE_CODE, WHSE_CODE, "", ""})
        End If
        If ASCMAIN1.CLIENT = "RGI" Then
            Sort_grdColumns(grdICTPHYCV, "STYLE_COST_VAR")
        Else
            Sort_grdColumns(grdICTPHYCV, "STYLE_CODE")
        End If

        If grdICTPHYCV.ActiveRow Is Nothing Then Setup_ICTPHYC2("", "")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return True

    End Function
    Function Load_Locations() As Boolean

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            MsgBox("You Must Pick a Valid Warehouse Code", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        ElseIf rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
            MsgBox("This option is valid only for Warehouses currently in a Physical Inventory", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        End If


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Compiling Location Control Totals")

        'If rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
        '    Fill_Records("ICTPHYCL", New String() {WHSE_CODE, "", WHSE_CODE, ASCMAIN1.CYP})
        'Else
        Fill_Records("ICTPHYCL", New String() {WHSE_CODE})
        'End If
        Sort_grdColumns(grdICTPHYCL, "LOCATION_CODE")
        ' If grdICTPHYCV.ActiveRow Is Nothing Then Setup_ICTPHYC2("")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return True

    End Function

    Function Load_Location_Style_Colors() As Boolean

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            MsgBox("You Must Pick a Valid Warehouse Code", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        ElseIf rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
            MsgBox("This option is valid only for Warehouses currently in a Physical Inventory", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        End If


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Compiling Location Control Totals")

        Fill_Records("ICTPHYCR", New String() {WHSE_CODE})
        Sort_grdColumns(grdICTPHYCR, "LOCATION_CODE,STYLE_CODE,COLOR_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return True

    End Function

    Private Sub grdICTPHYCI_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPHYCI.DoubleClickRow
        If Not ScreenMode Then
            '  Absx1.txtFor("WHSE_CODE").Text = ""
            Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Value & ""
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTPHYCI_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPHYCI.InitializeLayout

    End Sub

    Private Sub grdICTPHYC2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTPHYC2.InitializeRow
        Dim COUNT_LOOSE As Int64 = Val(e.Row.Cells("COUNT_LOOSE").Value & "")
        Dim CARTON_PACK_QTY As Int64 = Val(e.Row.Cells("CARTON_PACK_QTY").Value & "")
        If COUNT_LOOSE >= CARTON_PACK_QTY And CARTON_PACK_QTY <> 0 And CARTON_PACK_QTY <> 1 Then
            e.Row.Cells("COUNT_LOOSE").Appearance = grdWARNING_app
            e.Row.Cells("COUNT_LOOSE").ToolTipText = "Loose Count is greater than or equal to Carton Pack Qty"
        End If
    End Sub

    Sub Print_Counts(BY As String)

        Dim RPT As String = ""
        Dim RPT_TITLE As String = ""
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        If WHSE_CODE = "" Then Exit Sub
        Dim rowICTWHSE1 As DataRow = Fill_Record("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then Exit Sub

        ASCMAIN1.sql = "Select * from ICTPHYC1 where WHSE_CODE = '" & WHSE_CODE & "'"
        Fill_Records("ICTPHYC1", "", True, ASCMAIN1.sql)

        Dim RGI_CODE As String = ""
        If ASCMAIN1.CLIENT = "RGI" Then
            RGI_CODE = " and NVL(ICTPHYC2.STATUS,'A') = 'A' "
        End If

        ASCMAIN1.sql = "Select ICTPHYC2.*, ICTSTYL1.STYLE_DESC" _
            & " from ICTPHYC2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = ICTPHYC2.STYLE_CODE" & RGI_CODE _
            & " and ICTPHYC2.WHSE_CODE = '" & WHSE_CODE & "'"
        Fill_Records("ICTPHYC2", "", True, ASCMAIN1.sql)
        Dim dvwC2 As DataView = DirectCast(grdICTPHYC2.DataSource, DataTable).DefaultView
        dvwC2.RowFilter = "STATUS IS NULL"

        Select Case BY
            Case "T"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Ticket"
            Case "L"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Location"
            Case "S"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Style"
        End Select

        'Synch_TABLE_NAME("ICTSTYL1")
        Print_Report_Begin()
        CR_params.Add("SORT_BY", BY)
        Generate_Report(RPT, RPT_TITLE, "")
        Print_Report_End()

    End Sub

    Private Sub grdICTPHYCV_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPHYCV.InitializeLayout

    End Sub

    Private Sub btnEmpty_Click(sender As Object, e As EventArgs) Handles btnEmpty.Click
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Value
        If String.IsNullOrEmpty(WHSE_CODE) Then
            MsgBox("Enter Whse Code", vbOKOnly, "Whse code missing")
            Exit Sub
        End If
        Dim EmptyLocs = ""
        If 1 <> 1 Then
            'Enter acomplete level to falg
            Dim EmptyIsle = InputBox("Enter a Isle for Empties", "Empty locations")
            If String.IsNullOrEmpty(EmptyIsle) Then
                Exit Sub
            End If
            Dim EmptyLvl = InputBox("Enter a level for Empties", "Empty locations")
            If String.IsNullOrEmpty(EmptyLvl) Then
                Exit Sub
            End If

            ASCMAIN1.sql = "select * from whtlocm1, ICTPHYC1" & vbCrLf _
                    & " where whtlocm1.WHSE_CODE = ICTPHYC1.WHSE_CODE(+)" & vbCrLf _
                    & " and whtlocm1.LOCATION_CODE = ICTPHYC1.LOCATION_CODE(+)" & vbCrLf _
                    & " and WHTLOCM1.WHSE_CODE = 'MS'" & vbCrLf _
                    & " and WHTLOCM1.LOCATION_CODE like '" & EmptyIsle & "-%-" & EmptyLvl & "'"
            'Dim tblEmpty As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        Else
            'enter a range
            Dim EmptyStart = InputBox("Enter a Start Loc for Empties", "Empty locations")
            If String.IsNullOrEmpty(EmptyStart) Then
                Exit Sub
            End If
            Dim EmptyEnd = InputBox("Enter an End Loc for Empties", "Empty locations")
            If String.IsNullOrEmpty(EmptyEnd) Then
                Exit Sub
            End If

            ASCMAIN1.sql = "select * from whtlocm1, ICTPHYC1" & vbCrLf _
                    & " where whtlocm1.WHSE_CODE = ICTPHYC1.WHSE_CODE(+)" & vbCrLf _
                    & " and whtlocm1.LOCATION_CODE = ICTPHYC1.LOCATION_CODE(+)" & vbCrLf _
                    & " and WHTLOCM1.WHSE_CODE = 'MS'" & vbCrLf _
                    & " and WHTLOCM1.LOCATION_CODE between '" & EmptyStart & "' and '" & EmptyEnd & "'" & vbCrLf _
                    & " and WHTLOCM1.LOCATION_CODE like '%" & EmptyStart.ToString.Substring(7, 1) & "'"

            'Dim tblEmpty As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        End If
        Dim tblEmpty As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each rowEmpty As DataRow In tblEmpty.Select("", "LOCATION_CODE")
            If Not IsDBNull(rowEmpty("TICKET_NO")) Then
                If rowEmpty("TICKET_STATUS") <> "E" Then
                    MsgBox("Cannot Flag locations as Empty, counts found", vbCritical, "Update Failed")
                    Exit Sub
                End If
            End If
            EmptyLocs = EmptyLocs & "," & rowEmpty("LOCATION_CODE")
        Next

        If MsgBox(EmptyLocs, vbOKCancel, "Locations") = vbCancel Then
            Exit Sub
        End If

        BeginTrans()
        For Each rowEmpty As DataRow In tblEmpty.Select("")
            Dim LOCATION_CODE = rowEmpty("LOCATION_CODE")
            If LOCATION_CODE.ToString.Contains("-00-") Then
                Continue For
            End If
            Dim TICKET_NO1 As String = ASCMAIN1.Next_Control_No("ICTPHYC1.TICKET_NO")
            rowICTPHYC1 = dst.Tables("ICTPHYC1").NewRow
            With rowICTPHYC1
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("TICKET_NO") = TICKET_NO1
                .Item("COUNT_BY") = ASCMAIN1.USER_ID
                .Item("LOCATION_CODE") = LOCATION_CODE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = Now
                .Item("LAST_DATE") = Now
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("TICKET_STATUS") = "E"
            End With
            dst.Tables("ICTPHYC1").Rows.Add(rowICTPHYC1)


            ASCMAIN1.sql = "Update WHTLOCM1 " & vbCrLf _
                & " set LOCATION_LOCKED = '1' " & vbCrLf _
                & " where WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                & " and LOCATION_CODE = '" & LOCATION_CODE & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        Next
        Update_Record_TDA("ICTPHYC1")

        CommitTrans()
    End Sub
End Class