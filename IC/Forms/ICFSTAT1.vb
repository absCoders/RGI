Imports System.Text
Imports Infragistics.Win.UltraWinGrid

Public Class ICFSTAT1

    ' cmdUPC_Click is still not converted
    ' where clause in style selector
    ' compare multiple style

    Dim multi_style As Boolean = False
    Dim STYLEs As New List(Of String)
    Dim STYLE_CODE As String
    Dim STYLE_CODE_allocated As String
    Dim rowICTSTYL1 As DataRow
    Dim RANGE_STYLE_CODE As String
    Dim rowICTRSTY1 As DataRow
    Dim ICTSTYL1_Recent As String = ""

    Dim COLOR_CODE As String
    Dim AutoAllocate As Boolean
    Dim SOTDEMD1 As String
    Dim SOTSUPP1 As String
    Dim ShowZeroStatus As Boolean = False
    Dim edi850cust As List(Of String)
    Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

    Dim STYLE_CLASS_CODE As String
    Dim CARTON_PACK_QTY As Int32
    Dim STYLE_PRICE As Decimal

    Dim ICTCOSTA As String
    Dim ICTCOSTL As String
    Dim ICTCOSTU As String
    Dim ICTCOSTG As String

    Dim sqlICTQUOTX As String
    Dim rowICTQUOT1 As DataRow
    Dim QuoteEntryMode As String
    Dim QUOTE_NO As String

    Dim sqlICTCOSTL As String
    Private clsSOTORDR1 As TAC.SOCORDR1
    Private salesOrderLineDetails As New List(Of TAC.SOCORDR1.LineDetail)
    Dim iosf As String
    Dim sqlICTDUTY4 As String = ""
    Dim sqlSOTSUPPX As String = ""
    Dim TTM As New UltraWinToolTip.UltraToolTipManager

    Dim blnATONCE As Boolean = False
    Dim SO_PARM_SHIP_WINDOW_DAYS As Integer = 0
    Dim SO_PARM_ARRIVAL_BUFFER_DAYS As Integer = 0
    Dim SO_PARM_RELEASE_AT_ONCE As String = ""
    Dim blnAtOnceChanged As Boolean = False

    Dim sqlFUT_AVAIL_by_Style As String
    Dim sqlFUT_AVAIL_by_Style_Color As String
    Dim sqlFUT_AVAIL_by_Style_Color_Whse As String
    Dim sqlECOM As String = ""

#Region "ABS Standard Routines"

    Private Sub ICFSTAT1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            If (e.KeyCode = Keys.NumPad1 Or e.KeyCode = Keys.D1) And e.Alt Then
                Call Click_Command("Done", e)
            End If
        End If
    End Sub ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        btnTest.Visible = ASCMAIN1.Running_in_VS Or (ASCMAIN1.USER_ID = "mariog" And ASCMAIN1.DBS_COMPANY = "RGI")

        Get_PARM("ICTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("POTPARM1")

        AUDIT.Add("ICTATOP1", "*")
        AUDIT.Add("ICTATOP2", "*")

        If Not ASCMAIN1.Running_in_VS Then btnTest.Visible = False

        TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me,
                "",
                False,
                True,
                False,
                "", Now.Date.AddDays(30)) ' using 30 days release date horizon

        With dst

            Dim FYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60) ' ASCDATA1.GetDataValue("Select Min (OPS_YYYYPP) from ICTTRAN1")
            Dim LYP As String = ASCMAIN1.CYP 'ASCDATA1.GetDataValue("Select Max (OPS_YYYYPP) from ICTTRAN1")
            If LYP < ASCMAIN1.CYP Then LYP = ASCMAIN1.CYP
            ASCMAIN1.sql = "Select * from GLTPARM2 " & vbCrLf _
                & " where OPS_YYYYPP >= '" & FYP & "'" & vbCrLf _
                & "   and OPS_YYYYPP <= '" & LYP & "'"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTCOSTA.*" & vbCrLf _
                & " from ICTCOSTA where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(dst.Tables.Add, "ICTCOSTA", "**", 0, False, "VV")
            With dst.Tables("ICTCOSTA")
                .Columns.Add("USED_COST", GetType(System.Decimal), "IIF(LOT_QTY_USED=0,0,LOT_AMT_USED/LOT_QTY_USED)")
                .Columns.Add("YP_OPEN")
            End With

            '                & " where ICTCOSTL.OPS_YYYYPP_FIFO = :PARM1 and ICTCOSTL.STYLE_CODE = :PARM2 and ICTCOSTL.COLOR_CODE = :PARM3" & vbCrLf _

            ASCMAIN1.sql = "Select ICTCOSTL.*,ICTIREC1.VEND_CODE,ICTIREC2.PO_ORDER_NO,POTSHIP1.COST_COMPLETE,ICTIREC2.PO_COST,ICTIREC2.INV_COST" & vbCrLf _
                & ", POTSHIP3.PO_COST_FREIGHT_IN, POTSHIP3.PO_COST_DUTY, POTSHIP3.PO_COST_CUSTOMS, POTSHIP3.PO_COST_TRUCKING, POTSHIP3.DUTY_RATE_CODE, POTSHIP3.DUTY_RATE" & vbCrLf _
                & " from ICTCOSTL,ICTIREC1,ICTIREC2,POTSHIP1,POTSHIP3" & vbCrLf _
                & " where ICTCOSTL.STYLE_CODE = :PARM1 and ICTCOSTL.COLOR_CODE = :PARM2" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO (+) = ICTIREC2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO (+) = ICTIREC2.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTSHIP3.PO_ORDER_NO (+) = ICTIREC2.PO_ORDER_NO" & vbCrLf _
                & "   and POTSHIP3.PO_ORDER_LNO (+) = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTIREC2.RECEIPT_NO (+) = DECODE(ICTCOSTL.TRAN_TYPE,'R',ICTCOSTL.TRAN_NO,'?')" & vbCrLf _
                & "   and ICTIREC2.RECEIPT_LNO (+) = DECODE(ICTCOSTL.TRAN_TYPE,'R',ICTCOSTL.TRAN_LNO,0)" & vbCrLf _
                & "   and ICTIREC1.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = ICTIREC2.PO_SHIPMENT_NO" & vbCrLf
            sqlICTCOSTL = ASCMAIN1.sql

            Create_TDA(dst.Tables.Add, "ICTCOSTL", "**", 0, False, "VV")


            ASCMAIN1.sql = "Select ICTCOST1.*, '' IMPORT_CODE" _
                & " from ICTCOST1" _
                & " where ICTCOST1.STYLE_CODE = :PARM1" _
                & "   and ICTCOST1.COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTCOST1", "**", 0, True, "VV", 0)


            ASCMAIN1.sql = "SELECT 'O' ORDR_TYPE, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.CUST_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY ORDR" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN OPEN" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_ALLO ALLO" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP SHIP" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_CANC CANC" & vbCrLf _
                & ", 0 ORDERS" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_OPEN" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_CANC" & vbCrLf _
                & ", SOTORDR1.CUST_NAME" & vbCrLf _
                & ", SOTORDR1.ORDR_DATE_RECD, SOTORDR1.INIT_DATE" & vbCrLf _
                & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & " From SOTORDR2, SOTORDR1, ICTATOP1" & vbCrLf _
                & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 2)
            .Tables("SOTORDRX").Columns("CUST_NAME").MaxLength = 100
            With .Tables("SOTORDRX").Columns
                .Add("STYLE_COST", GetType(System.Decimal))
                .Add("COST_SOURCE")
                .Add("SUB")
                .Add("INV_DATE", GetType(System.DateTime))
                .Add("ORDR_UNIT_PRICE_CALC", GetType(System.Decimal), "IIF(ISNULL(ORDR,0)=0,0,ISNULL(ORDR_AMT,0)/ISNULL(ORDR,0))")
                .Add("GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ORDR_UNIT_PRICE_CALC,0)=0,0,100*(ISNULL(ORDR_UNIT_PRICE_CALC,0)-ISNULL(STYLE_COST,0))/ISNULL(ORDR_UNIT_PRICE_CALC,0))")
            End With
            With .Tables("SOTORDRX").Columns
                .Add("PO_SEQ_MAX_WAIT", GetType(System.Int32))
                .Add("RECD_SEQ", GetType(System.Int32))
                .Add("SHIP_SEQ", GetType(System.Int32))
                .Add("SHIP_DATE_PLUS", GetType(System.DateTime))
                .Add("QTY_ALLO_0", GetType(System.Int32))
                .Add("QTY_ALLO_1", GetType(System.Int32))
                .Add("QTY_ALLO_2", GetType(System.Int32))
                .Add("QTY_ALLO_3", GetType(System.Int32))
                .Add("QTY_ALLO_4", GetType(System.Int32))
                .Add("QTY_ALLO_5", GetType(System.Int32))
                .Add("QTY_ALLO_6", GetType(System.Int32))
                .Add("QTY_ALLO_7", GetType(System.Int32))
                .Add("QTY_ALLO_8", GetType(System.Int32))
                .Add("QTY_ALLO_9", GetType(System.Int32))
                .Add("ERROR")
            End With

            ASCMAIN1.sql = "SELECT SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTORDR2.ORDR_NO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.WHSE_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE_SUB, SOTORDR2.STYLE_DESC, SOTORDR2.ORDR_QTY ORDR" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN OPEN" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_ALLO ALLO" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP SHIP" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_CANC CANC" & vbCrLf _
                & ", 0 ORDERS" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT " & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_OPEN" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_CANC" & vbCrLf _
                & " from SOTORDR2,SOTORDR1,ICTCOLR1 " & vbCrLf _
                & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTORDRY", "**", 0, False, "", 0)
            With .Tables("SOTORDRY").Columns
                .Add("STYLE_COST", GetType(System.Decimal))
                .Add("COST_SOURCE")
                .Add("ORDR_UNIT_PRICE_CALC", GetType(System.Decimal), "IIF(ISNULL(ORDR,0)=0,0,ISNULL(ORDR_AMT,0)/ISNULL(ORDR,0))")
                .Add("GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ORDR_UNIT_PRICE_CALC,0)=0,0,(ISNULL(ORDR_UNIT_PRICE_CALC,0)-ISNULL(STYLE_COST,0))/ISNULL(ORDR_UNIT_PRICE_CALC,0))")
            End With
            .Tables("SOTORDRY").Columns("ORDR_NO").AllowDBNull = True

            'sqlFUT_AVAIL_by_Style = "Select STYLE_CODE, SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_AVAIL from ICTSTAT2 group by STYLE_CODE"
            'sqlFUT_AVAIL_by_Style_Color = "Select STYLE_CODE, COLOR_CODE, SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_AVAIL from ICTSTAT2 group by STYLE_CODE, COLOR_CODE"
            'sqlFUT_AVAIL_by_Style_Color_Whse = "Select STYLE_CODE, COLOR_CODE, WHSE_CODE, SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_AVAIL from ICTSTAT2 group by STYLE_CODE, COLOR_CODE, WHSE_CODE"
            sqlFUT_AVAIL_by_Style = ""

            If ASCMAIN1.CLIENT = "RGI" Then
                sqlECOM = "" _
                    & ", (" & vbCrLf _
                    & "Select STYLE_CODE" & vbCrLf _
                    & ", listagg (E1.ECOM_CODE, ',') within group ( order by E1.ECOM_CODE) ECOM_PARTNERS" & vbCrLf _
                    & "from ECTESTY1 Y1, ECTECOM1 E1 " & vbCrLf _
                    & "                    where Y1.ECOM_CODE = E1.ECOM_CODE" & vbCrLf _
                    & "                      and (NVL(Y1.SHIP_ECOM,'0') = '1' or NVL(Y1.SHIP_DROP,'0') = '1')" & vbCrLf _
                    & "                    group by STYLE_CODE" & vbCrLf _
                    & ") E" & vbCrLf
            End If

            ASCMAIN1.sql = "Select ICTSTYL1.*,ICTBODY2.MASTER_BODY_CODE" & vbCrLf _
                & " from ICTSTYL1,ICTBODY2" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = :PARM1" & vbCrLf _
                & "   and ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf
            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select ICTSTYL1.*,ICTBODY2.MASTER_BODY_CODE, E.ECOM_PARTNERS" & vbCrLf _
                    & " from ICTSTYL1,ICTBODY2" & vbCrLf _
                    & sqlECOM _
                    & " where ICTSTYL1.STYLE_CODE = :PARM1" & vbCrLf _
                    & "   and E.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                    & "   and ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf
            End If

            For Each TABLE_NAME As String In New String() {"ICTSTYL1", "ICTSTYL1_RECENT", "ICTSTYL1_VIEW"}
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "V", IIf(TABLE_NAME = "ICTSTYL1_RECENT" Or TABLE_NAME = "ICTSTYL1", 1, 0))

                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_DATE", GetType(System.DateTime))
                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_NO")
                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_CUST_CODE")
                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_CUST_PO")

                .Tables(TABLE_NAME).Columns.Add("QTY_ONHD", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_ONPO", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_TRAN", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_OPEN", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_PICK", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_COMM", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_PROD", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_NETA", GetType(System.Int64), "ISNULL(QTY_ONHD,0) + ISNULL(QTY_ONPO,0) + ISNULL(QTY_TRAN,0) - ISNULL(QTY_OPEN,0) - ISNULL(QTY_PICK,0) - ISNULL(QTY_COMM,0) + ISNULL(QTY_PROD,0)")
                .Tables(TABLE_NAME).Columns.Add("COLOR_CODE")
                .Tables(TABLE_NAME).Columns.Add("COLOR_DESC")
                .Tables(TABLE_NAME).Columns.Add("WHSE_CODE")
                .Tables(TABLE_NAME).Columns.Add("WHSE_DESC")
                .Tables(TABLE_NAME).Columns.Add("IMAGE", GetType(System.Byte()))
                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    .Tables(TABLE_NAME).Columns.Add("OPNQTY", GetType(System.Int64))
                    .Tables(TABLE_NAME).Columns.Add("OPNAMT", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("OPNCST", GetType(System.Decimal), "IIF(OPNQTY=0,0,OPNAMT/OPNQTY)")
                    .Tables(TABLE_NAME).Columns.Add("SHPQTY", GetType(System.Int64))
                    .Tables(TABLE_NAME).Columns.Add("SHPAMT", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("SHPCST", GetType(System.Decimal), "IIF(SHPQTY=0,0,SHPAMT/SHPQTY)")
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_LDP", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_LDP_CODE")
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_ELC", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_CUM", GetType(System.Decimal))
                End If
                .Tables(TABLE_NAME).Columns.Add("STYLE_COST_EXT", GetType(System.Decimal), "ISNULL(STYLE_COST,0)*ISNULL(QTY_NETA,0)")
                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_FIRST", GetType(System.Decimal))
                End If
            Next

            'ASCMAIN1.sql = "Select SOTORDR7.* from SOTORDR7 " & vbCrLf _
            '    & " where SOTORDR7.ORDR_GROUP_NO = :PARM1 " & vbCrLf _
            '    & "   and SOTORDR7.STYLE_CODE = :PARM2" & vbCrLf _
            '    & "   and SOTORDR7.COLOR_CODE = :PARM3" & vbCrLf _
            '    & "   and SOTORDR7.PICK_BATCH_NO is Null" & vbCrLf
            'Create_TDA(.Tables.Add, "SOTORDR7", "**", 0, False, "VVV")

            ASCMAIN1.sql = "Select ICTIADJ1.OPS_YYYYPP, 'A' TRAN_TYPE, ICTIADJ1.REGISTER_XNO TRAN_NO" & vbCrLf _
                & ", ICTIADJ2.ADJ_LNO TRAN_LNO, ICTIADJ2.ADJ_LNO TRAN_SLNO, ICTIADJ2.ADJ_LNO TRAN_XLNO " & vbCrLf _
                & ", ICTIADJ1.REGISTER_XNO TRAN_SOURCE_DOCUMENT, ICTIADJ1.ADJ_DATE TRAN_DATE" & vbCrLf _
                & ", ICTIADJ1.WHSE_CODE TRAN_WHSE_CODE, ICTIADJ1.REGISTER_XNO TRAN_CUST_CODE" & vbCrLf _
                & ", ICTIADJ1.REGISTER_XNO TRAN_VEND_CODE, ICTIADJ1.WHSE_CODE TRAN_WHSE_CODE_TO, ICTIADJ1.REASON_CODE TRAN_ADJ_REASON_CODE" & vbCrLf _
                & ", ICTIADJ2.ADJ_QTY TRAN_QTY, ICTIADJ1.REGISTER_XNO TRAN_REF, 'X' TRAN_STATUS_UPD" & vbCrLf _
                & ", ICTIADJ1.INIT_DATE, ICTIADJ1.INIT_OPER, ICTIADJ1.REGISTER_XNO TRAN_NO_ORIG, 'X' TRAN_TYPE_ORIG" & vbCrLf _
                & ", 'X' TRAN_ORIGINATE" & vbCrLf _
                & " from ICTIADJ1,ICTIADJ2 " & vbCrLf _
                & " where ICTIADJ2.ADJ_NO = ICTIADJ1.ADJ_NO" & vbCrLf _
                & "   and ICTIADJ2.STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTTRANX", "**", 0, False, "V", 0)
            .Tables("ICTTRANX").Columns.Add("TRAN_QTY_X", GetType(System.Int32), "TRAN_QTY * IIF(TRAN_TYPE='S' OR TRAN_TYPE = 'T',-1,1)")
            .Tables("ICTTRANX").Columns.Add("RUNNING_BALANCE", GetType(System.Int64))
            .Tables("ICTTRANX").Columns.Add("TRAN_CCVAT", GetType(System.String), "IIF(TRAN_TYPE='S' OR TRAN_TYPE='C',TRAN_CUST_CODE,IIF(TRAN_TYPE='R',TRAN_VEND_CODE,IIF(TRAN_TYPE='A',TRAN_ADJ_REASON_CODE,IIF(TRAN_TYPE='T',TRAN_WHSE_CODE_TO,''))))")
            .Tables("ICTTRANX").Columns("TRAN_SOURCE_DOCUMENT").MaxLength = -1
            .Tables("ICTTRANX").Columns("TRAN_REF").MaxLength = -1

            For Each TABLE_NAME As String In New String() {"ICTSTATA", "ICTSTATW"}
                With .Tables.Add(TABLE_NAME)
                    For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE", "WHSE_CODE", "STYLE_DESC", "COLOR_DESC", "WHSE_DESC"}
                        If TABLE_NAME = "ICTSTATA" And (COLUMN_NAME = "WHSE_CODE" Or COLUMN_NAME = "WHSE_DESC") Then
                        ElseIf TABLE_NAME = "ICTSTATW" And (COLUMN_NAME = "STYLE_DESC" Or COLUMN_NAME = "COLOR_DESC") Then
                        Else
                            .Columns.Add(COLUMN_NAME)
                        End If
                    Next
                    For Each COLUMN_NAME As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "PHY",
                                                                    "ON_HAND", "ON_ORDER", "TRAN", "OPEN", "PICK", "ALLO", "COMM", "PROD"}
                        .Columns.Add(COLUMN_NAME, GetType(System.Int64))
                    Next
                    If TABLE_NAME = "ICTSTATA" Then
                        .Columns.Add("UPC_CODE")
                        .Columns.Add("STYLE_COLOR_STATUS")
                        .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
                        .Columns.Add("THEME_DESC")
                    Else
                        .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE"), .Columns("WHSE_CODE")}
                    End If
                    .Columns.Add("OTS_INV", GetType(System.Int64), "ISNULL(ON_HAND,0) - ISNULL(PICK,0)")
                    .Columns.Add("OTS_WIP", GetType(System.Int64), "ISNULL(OTS_INV,0) + ISNULL(TRAN,0) + ISNULL(ON_ORDER,0)")
                    .Columns.Add("NET_POS", GetType(System.Int64), "ISNULL(OTS_WIP,0) - ISNULL(OPEN,0) - ISNULL(COMM,0) - ISNULL(PROD,0)")
                End With
            Next
            Create_Relation("ICTSTATA", "ICTSTATW", "STYLE_CODE,COLOR_CODE")

            ASCMAIN1.sql = "Select OPS_YYYYPP, STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", WHSE_QTY_BEG BEG, WHSE_QTY_SHP SHP, WHSE_QTY_RTN RTN, WHSE_QTY_REC REC" & vbCrLf _
                & ", WHSE_QTY_ADJ ADJ, WHSE_QTY_XFR XFR, WHSE_QTY_PHY PHY, 0 ON_HAND" & vbCrLf _
                & " from ICTSTAT1 where STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTATB", "**", 0, False, "V", 3)
            .Tables("ICTSTATB").Columns.Add("LEGEND")

            With .Tables.Add("ICTSTAT1_IMAGES")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("LOGO", GetType(System.Byte()))
                '.Columns.Add("IMAGE", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE")}
            End With

            ASCMAIN1.sql = "Select ICTSTYV1.*, APTVEND1.VEND_NAME" _
              & " from APTVEND1,ICTSTYV1" _
              & " where APTVEND1.VEND_CODE (+) = ICTSTYV1.VEND_CODE" _
              & "  and ICTSTYV1.STYLE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTSTYV1", "**", 0, True, "V")


            sqlICTQUOTX = "Select ICTQUOT1.*, X.STYLE_CODE_PLM, X.STYLES" _
              & " from ICTQUOT1, (Select QUOTE_NO, MIN (STYLE_CODE_PLM) STYLE_CODE_PLM, Count (*) STYLES from ICTQUOT2 group by QUOTE_NO) X" _
              & " where X.QUOTE_NO = ICTQUOT1.QUOTE_NO and ICTQUOT1.QUOTE_TYPE = 'S'"
            ASCMAIN1.sql = sqlICTQUOTX
            Create_TDA(.Tables.Add, "ICTQUOTX", "**", 0, False, "")


            Create_TDA(.Tables.Add, "ICTQUOT1", "*")
            With .Tables("ICTQUOT1")
                .Columns.Add("LOGO", GetType(System.Byte()))
                '   .PrimaryKey = New DataColumn() {.Columns("QUOTE_NO")}
            End With

            '' Create_TDA(.Tables.Add, "ICTQUOT2", "*")
            'ASCMAIN1.sql = "Select * from ICTQUOT2 where ROWNUM < 1"
            'Dim ICTQUOT2 As String = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTQUOT2 & " Add Primary Key (QUOTE_NO,STYLE_CODE_PLM)")
            'ASCMAIN1.sql = "Select ICTQUOT2.*, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.IMAGE_NAME" _
            '    & " from " & ICTQUOT2 & " ICTQUOT2, ICTSTYL1 where ICTSTYL1.STYLE_CODE = ICTQUOT2.STYLE_CODE_PLM"
            'Create_TDA(.Tables.Add("ICTQUOT2"), ICTQUOT2, "**", 0, True, "", 2)

            ASCMAIN1.sql = "Select ICTQUOT2.*, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.SEASON_CODE" _
               & " from ICTQUOT2, ICTSTYL1 where ICTSTYL1.STYLE_CODE = ICTQUOT2.STYLE_CODE_PLM" _
               & " and ICTQUOT2.QUOTE_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTQUOT2", "**", 0, True, "V", 2)
            With .Tables("ICTQUOT2")
                '    .Columns.Add("IMAGE", GetType(System.Byte()))
                '    .Columns("SEQ").DataType = GetType(System.Int32)
                'End With

                'With .Tables("ICTQUOT2")
                .Columns.Add("IMAGE_NAME")
                .Columns.Add("IMAGE", GetType(System.Byte()))
                .Columns("SEQ").DataType = GetType(System.Int32)
                .Columns.Add("WHSE_01")
                .Columns.Add("DATE_01", GetType(System.DateTime))
                .Columns.Add("QTY_01", GetType(System.Int64))
                .Columns.Add("WHSE_02")
                .Columns.Add("DATE_02", GetType(System.DateTime))
                .Columns.Add("QTY_02", GetType(System.Int64))
                .Columns.Add("WHSE_03")
                .Columns.Add("DATE_03", GetType(System.DateTime))
                .Columns.Add("QTY_03", GetType(System.Int64))
                .Columns.Add("WHSE_04")
                .Columns.Add("DATE_04", GetType(System.DateTime))
                .Columns.Add("QTY_04", GetType(System.Int64))
                .Columns.Add("STYLE_SPEC_01")
                .Columns.Add("STYLE_TYPE_DTL_01")
                .Columns.Add("STYLE_PRICE_01", GetType(System.Int64))
                .Columns.Add("STYLE_SPEC_02")
                .Columns.Add("STYLE_TYPE_DTL_02")
                .Columns.Add("STYLE_PRICE_02", GetType(System.Int64))
                .Columns.Add("STYLE_SPEC_03")
                .Columns.Add("STYLE_TYPE_DTL_03")
                .Columns.Add("STYLE_PRICE_03", GetType(System.Int64))
                .Columns.Add("STYLE_SPEC_04")
                .Columns.Add("STYLE_TYPE_DTL_04")
                .Columns.Add("STYLE_PRICE_04", GetType(System.Int64))
            End With

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
                & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & " From POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, ICTATOP2 " & vbCrLf _
                & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 0)
            With .Tables("POTORDRX")
                .Columns("PO_SHIPMENT_NO").AllowDBNull = True
                .Columns("PO_SHIPMENT_LNO").AllowDBNull = True
                .Columns("PO_REFERENCE").AllowDBNull = True
                .Columns.Add("PO_ARRIVAL_DATE_PLUS", GetType(System.DateTime))
                '  .Columns("PO_SHIPMENT_NO").AllowDBNull = True
            End With
            ' .Tables("POTORDRX").Columns.Add("", GetType(System.DateTime))

            ASCMAIN1.sql = "Select ICTATOP2.*" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                & ", CASE WHEN ICTATOP2.PS_CODE = 'P' THEN X.PO_DATE_ETA else POTSHIP1.PO_SHIP_ETA END PS_ETA_NOW" & vbCrLf _
                & " from ICTATOP2, POTORDR1, POTSHIP1, ICTSTYL1" & vbCrLf _
                & ", (SELECT PO_ORDER_NO, STYLE_CODE, COLOR_CODE, MIN (PO_DATE_ETA) PO_DATE_ETA FROM POTORDR2 GROUP BY PO_ORDER_NO, STYLE_CODE, COLOR_CODE) X" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = ICTATOP2.STYLE_CODE" & vbCrLf _
                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = CASE WHEN ICTATOP2.PS_CODE = 'S' THEN PS_NO ELSE '' END" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO (+) = CASE WHEN ICTATOP2.PS_CODE = 'P' THEN PS_NO ELSE '' END" & vbCrLf _
                & "AND X.PO_ORDER_NO (+) = CASE WHEN ICTATOP2.PS_CODE = 'P' THEN PS_NO ELSE '' END" & vbCrLf _
                & "AND X.STYLE_CODE (+) = ICTATOP2.STYLE_CODE" & vbCrLf _
                & "AND X.COLOR_CODE (+) = ICTATOP2.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTATOP2", "**", 0, True,,, "STYLE_ARRIVAL_BUFFER_DAYS,STYLE_AT_ONCE_UNTIL,STYLE_AT_ONCE_ACTIVE")


            ASCMAIN1.sql = "Select ICTATOP1.*" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE ORDR_SHIP_DATE_NOW, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from ICTATOP1, SOTORDR1, ICTSTYL1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = ICTATOP1.STYLE_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO (+) = ICTATOP1.ORDR_NO"
            Create_TDA(.Tables.Add, "ICTATOP1", "**", 0, True,,, "STYLE_SHIP_WINDOW_DAYS,ORDR_SHIP_DATE_PLUS,STYLE_AT_ONCE_UNTIL,STYLE_AT_ONCE_ACTIVE")

            Create_TDA(.Tables.Add, "POTSHIP7", "*", 1)
            Create_TDA(.Tables.Add, "POTSHIP8", "*", 1)
            .Tables("POTSHIP8").Columns.Add("UNITS", GetType(System.Int32), "QTY*IIF(ISNULL(DOZENS,'0')='1',12,1)")

            '  Create_Relation("POTORDRX", "POTSHIP7", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")

            Create_Relation("POTSHIP7", "POTSHIP8", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO,CARTON_NO")
            .Tables("POTSHIP7").Columns.Add("STYLES", GetType(System.Int32), "COUNT(CHILD(POTSHIP7_POTSHIP8).STYLE_CODE)")
            .Tables("POTSHIP7").Columns.Add("UNITS", GetType(System.Int32), "SUM(CHILD(POTSHIP7_POTSHIP8).UNITS)")
            .Tables("POTSHIP7").Columns.Add("PPK_INNER_QTY_CALC", GetType(System.Int32), "SUM(CHILD(POTSHIP7_POTSHIP8).PPK_INNER_QTY)")

            .Tables("POTSHIP8").Columns.Add("CARTONS", GetType(System.Int32), "PARENT(POTSHIP7_POTSHIP8).CARTONS")
            .Tables("POTSHIP8").Columns.Add("TOTAL_UNITS", GetType(System.Int32), "ISNULL(UNITS,0) * ISNULL(CARTONS,0)")
            .Tables("POTSHIP7").Columns.Add("TOTAL_UNITS", GetType(System.Int32), "SUM(CHILD(POTSHIP7_POTSHIP8).TOTAL_UNITS)")
            .Tables("POTSHIP7").Columns.Add("STYLE_CODE_1", GetType(System.String), "MIN(CHILD(POTSHIP7_POTSHIP8).STYLE_CODE)")
            .Tables("POTSHIP7").Columns.Add("COLOR_CODE_1", GetType(System.String), "MIN(CHILD(POTSHIP7_POTSHIP8).COLOR_CODE)")
            .Tables("POTSHIP7").Columns.Add("ITEM_CODE", GetType(System.String), "IIF(ISNULL(PPK_CODE,'')='',ISNULL(STYLE_CODE_1,'') + ISNULL(COLOR_CODE_1,''),PPK_CODE)")
            .Tables("POTSHIP7").Columns.Add("CBM", GetType(System.Decimal), "ISNULL(CARTONS,0) * ISNULL(CARTON_VOLUME,0) / 1000000")
            .Tables("POTSHIP8").Columns.Add("CBM", GetType(System.Decimal), "IIF(ISNULL(PARENT(POTSHIP7_POTSHIP8).TOTAL_UNITS,0) = 0, 0, ISNULL(TOTAL_UNITS,0) * ISNULL(PARENT(POTSHIP7_POTSHIP8).CBM,0) / ISNULL(PARENT(POTSHIP7_POTSHIP8).TOTAL_UNITS,0))")

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO" & vbCrLf _
                & ", SOTORDR2.ORDR_LNO" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & ", 'S' as RECORD_TYPE" & vbCrLf _
                & ", 'O' as RECORD_SUB_TYPE" & vbCrLf _
                & ", SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", '0' ORDR_PRIORITY" & vbCrLf _
                & ", SYSDATE SD_DATE" & vbCrLf _
                & ", 'MM/DD/YY' SD_DATE_X" & vbCrLf _
                & ", SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE" & vbCrLf _
                & ", SYSDATE SHIP_ETA" & vbCrLf _
                & ", 0 SD_QTY" & vbCrLf _
                & ", 0 SD_QTY_ALLO" & vbCrLf _
                & ", 0 SD_QTY_ALLO_CUR" & vbCrLf _
                & ", 0 SD_QTY_ALLO_FUT" & vbCrLf _
                & ", 0 SD_QTY_ALLO_CXL" & vbCrLf _
                & ", 0 BALANCE" & vbCrLf _
                & ", 'X' ORDR_RELEASE" & vbCrLf _
                & ", SYSDATE ORDR_DEMAND_DATE" & vbCrLf _
                & ", SYSDATE ORDR_PRIORITY_DATE" & vbCrLf _
                & ", SYSDATE ORDR_PRIORITY_DATE_ORIG" & vbCrLf _
                & ", SYSDATE ORDR_RELEASE_AVAIL" & vbCrLf _
                & ", '0' ORDR_BACKORDER" & vbCrLf _
                & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "   and SOTORDR2.STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTALLO1", "**", 0, False, "V", 0)
            With .Tables("SOTALLO1")
                .Columns("ORDR_NO").AllowDBNull = True
                .Columns("ORDR_LNO").AllowDBNull = True
                .Columns("CUST_CODE").AllowDBNull = True
                .Columns("ORDR_CUST_PO").AllowDBNull = True
                .Columns("SD_QTY_ALLO").DataType = GetType(System.Int64)
                .Columns("SD_QTY_ALLO_CUR").DataType = GetType(System.Int64)
                .Columns("SD_QTY_ALLO_FUT").DataType = GetType(System.Int64)
                .Columns("SD_QTY_ALLO_CXL").DataType = GetType(System.Int64)
            End With

            'ASCMAIN1.sql = "Select * from SOTRSRV2 where RSRV_NO = :PARM1 and RSRV_LNO = :PARM2"
            'Create_TDA(.Tables.Add, "SOTRSRV2", "**", 0, True, "VN", 2)

            ASCMAIN1.sql = "Select SOTINVH2.CUST_CODE" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) AMT" & vbCrLf _
                & ", MIN (SOTINVH1.INV_DATE) INV_DATE_MIN" & vbCrLf _
                & ", MAX (SOTINVH1.INV_DATE) INV_DATE_MAX" & vbCrLf _
                & ", COUNT (*) INV_LNOS" & vbCrLf _
                & " from SOTINVH2,SOTINVH1" & vbCrLf _
                & " where SOTINVH2.ORDR_YYYYPP_UPDATED >= :PARM1" & vbCrLf _
                & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= :PARM2" & vbCrLf _
                & "   and SOTINVH2.STYLE_CODE = :PARM3" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "   and SOTINVH1.ORDR_TYPE_CODE <> 'XFR'" & vbCrLf _
                & "   and (SOTINVH2.COLOR_CODE = :PARM4 or :PARM4 = '*')" & vbCrLf _
                & " group by SOTINVH2.CUST_CODE" & vbCrLf
            ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME from (" & ASCMAIN1.sql & ") X,ARTCUST1" _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "VVVV", 1)
            .Tables("SOTINVHX").Columns("QTY").DataType = GetType(System.Int64)
            .Tables("SOTINVHX").Columns("INV_LNOS").DataType = GetType(System.Int64)

            ASCMAIN1.sql = "SELECT SOTINVH1.ORDR_CUST_PO, SOTINVH1.CUST_STORE_NO, SOTINVH2.COLOR_CODE" & vbCrLf _
                & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.WHSE_CODE" & vbCrLf _
                & " from SOTINVH2,SOTINVH1" & vbCrLf _
                & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTINVHY", "**", 0, False, "VVV", 0)
            .Tables("SOTINVHY").Columns.Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")

            'ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR1.CUST_STORE_NO" & vbCrLf _
            '               & ", NVL(ARTCUST2.CUST_RANK,0) CUST_RANK" & vbCrLf _
            '               & ", NVL(SOTORDR2.ORDR_QTY_PRE_ALLO,0) ORDR_QTY_PRE_ALLO" & vbCrLf _
            '               & ", NVL(SOTORDR2.ORDR_QTY_ALLO,0) ORDR_QTY_ALLO" & vbCrLf _
            '               & ", SOTORDR2.ORDR_QTY_OPEN" & vbCrLf _
            '               & " from SOTORDR1,ARTCUST2,SOTORDR2" & vbCrLf _
            '               & " where ROWNUM < 1" & vbCrLf
            'Create_TDA(.Tables.Add, "SOTPREA1", "**", 0, False, "V", 2)

            SOTSUPP1 = ASCMAIN1.Temp_Table("Select * from SOTSUPP1")
            ASCMAIN1.sql = "Select * from " & SOTSUPP1
            Create_TDA(.Tables.Add, "SOTSUPP1", "**", 0, False)

            SOTDEMD1 = ASCMAIN1.Temp_Table("Select * from SOTDEMD1")
            ASCMAIN1.sql = "Select * from " & SOTDEMD1
            Create_TDA(.Tables.Add, "SOTDEMD1", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTSTYL3 where STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False, "V")
            ' .Tables("ICTSTYL3").Columns.Add("SEL")

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, WHSE_CODE" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PO_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PO_SUM" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PS_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PS_SUM" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SO_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SO_SUM" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SP_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SP_SUM" & vbCrLf _
                & " from ICTSTAT2"
            Create_TDA(.Tables.Add, "ICTSTATO", "**", 0, False, "", 3)


            With .Tables.Add("ICTPRICX")
                .Columns.Add("TIER", GetType(System.Int32))
                .Columns.Add("PCT", GetType(System.Int32))
                .Columns.Add("DESC")
                .Columns.Add("CASES", GetType(System.Decimal))
                .Columns.Add("ABBR")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("PRICE", GetType(System.Decimal))
            End With

            'ASCMAIN1.sql = "Select * from ICTSTDQ2"
            'Create_TDA(.Tables.Add, "ICTSTDQ2", "**", 0, False)


            ASCMAIN1.sql = "Select * from ICTDISC1"
            Create_TDA(.Tables.Add, "ICTDISC1", "**", 0, False)
            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)
            ASCMAIN1.sql = "Select * from ICTSGRP1"
            Create_TDA(.Tables.Add, "ICTSGRP1", "**", 0, False)

            ASCMAIN1.sql = "Select WHTLOCB1.*,WHTBARC1.LOAD_NO,WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
            & " from WHTLOCB1,WHTBARC1,WHTLOCM1" & vbCrLf _
            & " where WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
            & "   and WHTLOCM1.LOCATION_CODE (+) = WHTLOCB1.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "", 5)
            .Tables("WHTLOCB1").Columns.Add("PERSIST")
            .Tables("WHTLOCB1").Columns.Add("LOCATION_QTY_AVAIL", GetType(System.Int64), "ISNULL(LOCATION_QTY,0)-ISNULL(LOCATION_QTY_WAVE,0)")

            ASCMAIN1.sql = "Select WHTINST1.WAVE_NO, WHTWAVE1.CUST_CODE, WHTWAVE1.WAVE_STATUS" & vbCrLf _
                & ", SUM (WHTINST2.LOCATION_QTY_WAVE) WAVED" & vbCrLf _
                & ", SUM (WHTINST2.LOCATION_QTY_PICK) PICKED" & vbCrLf _
                & " from WHTINST1,WHTINST2,WHTWAVE1" & vbCrLf _
                & " where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
                & "   and WHTINST2.STYLE_CODE = :PARM1 AND WHTINST2.COLOR_CODE = :PARM2" & vbCrLf _
                & "   and WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                & "   and WHTWAVE1.WAVE_STATUS = 'O'" & vbCrLf _
                & " group by WHTINST1.WAVE_NO, WHTWAVE1.CUST_CODE, WHTWAVE1.WAVE_STATUS"
            Create_TDA(.Tables.Add, "WHTINSTX", "**", 0, False, "VV")

            ASCMAIN1.sql = "Select ICTWHSE1.WHSE_CODE, ICTWHSE1.WHSE_DESC, ICTWHSE1.WHSE_MRP_EXC_IND" & vbCrLf _
                & " from ICTWHSE1"
            Create_TDA(.Tables.Add, "ICTWHSES", "**", 0, False)
            .Tables("ICTWHSES").Columns.Add("SEL")
            .Tables("ICTWHSES").Columns("SEL").DefaultValue = "1"



            sqlICTDUTY4 = "Select ICTDUTY4.*, ICTDUTY1.DUTY_RATE_DESC" _
             & " from ICTDUTY1,ICTDUTY4" _
             & " where ICTDUTY1.DUTY_RATE_CODE = ICTDUTY4.DUTY_RATE_CODE"
            ASCMAIN1.sql = sqlICTDUTY4 _
            & "  and ICTDUTY4.DUTY_RATE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTDUTY4", "**", 0, True, "V")


            With .Tables.Add("SOTSUPPA").Columns
                .Add("PO_SEQ", GetType(System.Int32))
                .Add("PO_SHIPMENT_NO")
                .Add("PO_ORDER_NO")
                .Add("PO_ARRIVAL_DATE", GetType(System.DateTime))
                .Add("PO_ARRIVAL_DATE_PLUS", GetType(System.DateTime))
                .Add("PO_QTY", GetType(System.Int64))
                .Add("PO_QTY_USED", GetType(System.Int64))
                .Add("PO_QTY_LEFT", GetType(System.Int64), "ISNULL(PO_QTY,0) - ISNULL(PO_QTY_USED,0)")
                .Add("PO_QTY_CUM", GetType(System.Int64))
            End With
            .Tables("SOTSUPPA").PrimaryKey = New DataColumn() { .Tables("SOTSUPPA").Columns("PO_SEQ")}

            sqlSOTSUPPX = "Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
                & ", NVL(ICTSTYL1.STYLE_COST,ICTSTYV1.PO_COST ) STYLE_COST" & vbCrLf _
                & ", NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) CUR_AVA" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_ON_ORDER, ICTSTAT2.WHSE_QTY_TRAN, ICTSTAT2.WHSE_QTY_OPEN" & vbCrLf _
                & ", NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) FUT_AVA" & vbCrLf _
                & " from ICTSTYL1,ICTSTAT2,ICTCLAS1,ICTSTYV1,ICTSTYC1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = ICTSTAT2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = ICTSTAT2.COLOR_CODE" & vbCrLf _
                & "   and ICTCLAS1.STYLE_CLASS_CODE = ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE"
            ASCMAIN1.sql = sqlSOTSUPPX & vbCrLf _
                & "   and ICTSTAT2.WHSE_CODE = 'MS'" & vbCrLf _
                & "   and NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) > 0 and NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) > 0"
            Create_TDA(.Tables.Add, "SOTSUPPX", "**", 0, False)

            With .Tables("SOTSUPPX")
                .Columns.Add("CUR_AVA_CST", GetType(System.Decimal), "CUR_AVA * STYLE_COST")
                .Columns.Add("FUT_AVA_CST", GetType(System.Decimal), "FUT_AVA * STYLE_COST")
            End With

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                ASCMAIN1.sql = "SELECT" & vbCrLf _
                    & " PROMO_DESC," & vbCrLf _
                    & " PROMO_START_DATE," & vbCrLf _
                    & " PROMO_END_DATE," & vbCrLf _
                    & " P2.PROMO_STYLE_NOTES," & vbCrLf _
                    & " P2.PROMO_UNIT_PRICE" & vbCrLf _
                    & " FROM ICTPROM1 P1, ICTPROM2 P2" & vbCrLf _
                    & " WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO" & vbCrLf _
                    & " AND P2.STYLE_CODE = :PARM1"
                Create_TDA(.Tables.Add, "ICTPROMX", "**", 0, False, "V", 0)
            End If


        End With

        Fill_Records("ICTDISC1")
        Fill_Records("ICTCLAS1")

        Fill_Records("ICTWHSES")
        For Each rowICTWHSES As DataRow In dst.Tables("ICTWHSES").Select("WHSE_MRP_EXC_IND='1'")
            rowICTWHSES.Item("SEL") = "0"
        Next

        If ASCMAIN1.CLIENT = "NYA" Then
            If ASCMAIN1.USER_ID = "wjz" Or ASCMAIN1.USER_ID = "ishalom" Then
                Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSES").Rows.Find("35")
                If rowICTWHSE1 IsNot Nothing Then rowICTWHSE1.Item("SEL") = "1"
            End If
        End If

        With dst.Tables("ICTSTAT1_IMAGES")
            Dim row As DataRow = .NewRow
            row.Item("STYLE_CODE") = "X"
            Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                row.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
            End If
            .Rows.Add(row)
        End With

        'With dst.Tables("ICTQUOT1")
        '    Dim rowICTQUOT1 As DataRow = .NewRow
        '    rowICTQUOT1.Item("QUOTE_NO") = "".PadLeft(6, "0")
        '    Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
        '    If My.Computer.FileSystem.FileExists(FILENAME) Then
        '        rowICTQUOT1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        '    End If
        '    .Rows.Add(rowICTQUOT1)
        'End With

        With dst.Tables("ICTSTDQ1")
            .Columns.Add("QTY_PLUS", GetType(System.Int64))
            .Columns.Add("QTY_PLUS_CUM", GetType(System.Int64))
            If ASCMAIN1.CLIENT = "NOT RGI" Then
                .Columns("QTY_PLUS").Expression = "SUPPLY_QTY"
            Else
                .Columns("QTY_PLUS").Expression = "QTY_ATS"
                .Columns("QTY_PLUS_CUM").Expression = "QTY_ATS_CUM"
            End If
        End With

        Fill_Records("GLTPARM2")

        grdICTCOSTL.DataSource = dst.Tables("ICTCOSTL")
        grdICTCOSTA.DataSource = dst.Tables("ICTCOSTA")
        grdICTCOST1.DataSource = dst.Tables("ICTCOST1")

        grdICTSTYL3.DataSource = dst.Tables("ICTSTYL3")

        grdICTSTYL1_Recent.DataSource = dst.Tables("ICTSTYL1_RECENT")

        grdICTQUOT2.DataSource = dst.Tables("ICTQUOT2")

        grdICTSTATA.DataSource = dst.Tables("ICTSTATA")
        grdICTSTATB.DataSource = dst.Tables("ICTSTATB")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdSOTORDRY.DataSource = dst.Tables("SOTORDRY")
        grdICTTRANX.DataSource = dst.Tables("ICTTRANX")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")

        grdPOTSHIP7.DataSource = dst.Tables("POTSHIP7")
        grdPOTSHIP8.DataSource = dst.Tables("POTSHIP8")

        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdSOTINVHY.DataSource = dst.Tables("SOTINVHY")

        grdSOTPREA1.DataSource = dst.Tables("SOTPREA1")
        grdSOTALLO1.DataSource = dst.Tables("SOTALLO1")
        grdICTSTDQ1.DataSource = dst.Tables("ICTSTDQ1")

        grdICTPRICX.DataSource = dst.Tables("ICTPRICX")
        grdICTSTDQ2.DataSource = dst.Tables("ICTSTDQ2")
        grdICTSTYV1.DataSource = dst.Tables("ICTSTYV1")

        grdICTSTATO.DataSource = dst.Tables("ICTSTATO")

        grdICTQUOTX.DataSource = dst.Tables("ICTQUOTX")
        grdWHTLOCB1.DataSource = dst.Tables("WHTLOCB1")
        grdWHTINSTX.DataSource = dst.Tables("WHTINSTX")

        grdICTWHSES.DataSource = dst.Tables("ICTWHSES")
        Sort_grdColumns(grdICTWHSES, "WHSE_CODE")

        grdICTDUTY4.DataSource = dst.Tables("ICTDUTY4")

        grdSOTSUPPA.DataSource = dst.Tables("SOTSUPPA")
        grdSOTSUPPX.DataSource = dst.Tables("SOTSUPPX")

        grdICTATOP1.DataSource = dst.Tables("ICTATOP1")
        grdICTATOP2.DataSource = dst.Tables("ICTATOP2")


        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            grdICTPROMX.DataSource = dst.Tables("ICTPROMX")

        End If


        Show_Filter(grdICTATOP1, True)
        Create_Summary(grdICTATOP1, "STYLE_CODE", "Count")

        Show_Filter(grdICTATOP2, True)
        Create_Summary(grdICTATOP2, "STYLE_CODE", "Count")


        Show_Filter(grdSOTSUPPX, True)
        Create_Summary(grdSOTSUPPX, "STYLE_CODE", "Count")
        Create_Summary(grdSOTSUPPX, New String() {"CUR_AVA", "FUT_AVA", "CUR_AVA_CST", "FUT_AVA_CST"})

        Create_Summary(grdICTQUOTX, "QUOTE_NO", "Count")
        Create_Summary(grdWHTLOCB1, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCB1, New String() {"LOCATION_QTY", "LOCATION_QTY_WAVE", "LOCATION_QTY_AVAIL"})

        Show_Filter(grdICTSTYL1_Recent, True)
        grdICTSTYL1_Recent.DisplayLayout.GroupByBox.Hidden = False
        Create_Summary(grdICTSTYL1_Recent, "STYLE_CODE", "Count")
        Create_Summary(grdICTSTYL1_Recent, "STYLE_COST_EXT")
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Create_Summary(grdICTSTYL1_Recent, New String() {"OPNQTY", "OPNAMT", "SHPQTY", "SHPAMT"})
        End If

        Create_Summary(grdSOTORDRX, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDRX, New String() {"ORDERS", "ORDR", "OPEN", "SHIP", "PICK", "CANC", "ALLO", "ORDR_AMT"}) ', "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP"
        Create_Summary(grdSOTORDRX, New String() {"QTY_ALLO_0", "QTY_ALLO_1", "QTY_ALLO_2", "QTY_ALLO_3", "QTY_ALLO_4", "QTY_ALLO_5", "QTY_ALLO_6", "QTY_ALLO_7", "QTY_ALLO_8", "QTY_ALLO_9"})

        Create_Summary(grdSOTORDRY, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRY, New String() {"ORDERS", "ORDR", "OPEN", "SHIP", "PICK", "CANC", "ALLO", "ORDR_AMT"})

        'Create_Summary(grdSOTPREA1, New String() {"ORDR_QTY_PRE_ALLO", "ORDR_QTY_ALLO", "ORDR_QTY_OPEN"})

        Create_Summary(grdSOTALLO1, New String() {"SD_QTY_ALLO", "SD_QTY_ALLO_CUR", "SD_QTY_ALLO_FUT", "SD_QTY_ALLO_CXL"})
        grdSOTALLO1.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.InGroupByRows

        Create_Summary(grdICTSTATA, "COLOR_CODE", "Count")
        Create_Summary(grdICTSTATA, New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "PHY",
                                                  "ON_HAND", "ON_ORDER", "TRAN", "OPEN", "PICK", "ALLO", "COMM", "PROD", "OTS_INV", "OTS_WIP", "NET_POS"})

        Create_Summary(grdICTCOSTL, "TRAN_NO", "Count")
        Create_Summary(grdICTCOSTL, New String() {"TRAN_QTY", "LOT_QTY_ONHD", "LOT_AMT_ONHD", "LOT_QTY_SHP", "LOT_AMT_SHP",
                                                  "LOT_QTY_RTN", "LOT_AMT_RTN", "LOT_QTY_ADJ", "LOT_AMT_ADJ"})

        Create_Summary(grdICTSTATB, New String() {"SHP", "RTN", "REC", "ADJ", "XFR", "PHY"})

        Create_Summary(grdICTSTATO, "STYLE_CODE", "Count")

        Create_Summary(grdICTTRANX, "TRAN_NO", "Count")
        Create_Summary(grdICTTRANX, New String() {"TRAN_QTY_X"})

        Create_Summary(grdPOTORDRX, "INIT_DATE", "Count")
        Create_Summary(grdPOTORDRX, New String() {"PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_ORD", "PO_QTY_OPN"})

        Create_Summary(grdSOTINVHX, "CUST_CODE", "Count")
        Create_Summary(grdSOTINVHX, New String() {"QTY", "AMT"})
        Create_Summary(grdSOTINVHY, "ORDR_CUST_PO", "Count")
        Create_Summary(grdSOTINVHY, New String() {"ORDR_QTY_SHIP", "ORDR_AMT_SHIP"})

        Create_Summary(grdICTQUOT2, "STYLE_CODE_PLM", "Count")

        'Create_Summary(grdICTSTYL3, "ATTR_CODE", "Count")
        'Create_Summary(grdICTSTYL3, New String() {"SEL"})

        Style_grdICTSTYL1_Recent()

        chkAutoAllocate.Checked = True


        With grdICTDUTY4.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False
        End With


        grdICTATOP1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdICTATOP1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdICTATOP1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                gcol.CellActivation = Activation.NoEdit
                If New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "STYLE_CLASS_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"ORDR_SHIP_DATE_NOW"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"STYLE_SHIP_WINDOW_DAYS", "ORDR_SHIP_DATE_PLUS", "STYLE_AT_ONCE_UNTIL", "STYLE_AT_ONCE_ACTIVE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With

        grdICTATOP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdICTATOP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdICTATOP2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                gcol.CellActivation = Activation.NoEdit
                If New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "STYLE_CLASS_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"PS_ETA_NOW"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"STYLE_ARRIVAL_BUFFER_DAYS", "STYLE_AT_ONCE_UNTIL", "STYLE_AT_ONCE_ACTIVE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With

        With grdICTCOSTL.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                If New String() {"LOT_QTY_ONHD", "LOT_AMT_ONHD", "QTY_TRAN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"LOT_QTY_SHP", "LOT_AMT_SHP"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"LOT_QTY_RTN", "LOT_AMT_RTN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"LOT_QTY_ADJ", "LOT_AMT_ADJ"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.MediumPurple
                ElseIf New String() {"TRAN_QTY", "TRAN_COST", "LOT_DAYS", "COST_COMPLETE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With


        With grdICTCOSTA.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                If New String() {"USED_COST"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"STYLE_COST", "LOT_DAYS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        grdICTSTATA.DisplayLayout.Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.FixedOnTop
        With grdICTSTATA.DisplayLayout.Bands(0)
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                .Columns("OTS_INV").CellAppearance.BackColor = Color.Empty
            Else
                .Columns("OTS_INV").CellAppearance.BackColor = Color.Yellow
            End If

            .Columns("NET_POS").CellAppearance.BackColor = Color.Yellow
            .Columns("OTS_INV").Width = 65
            .Columns("NET_POS").Width = 65
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor2 = Color.LightGray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

            For Each COLUMN_NAME As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "PHY",
                                                                 "ON_HAND", "ON_ORDER", "TRAN", "OPEN", "PICK", "ALLO", "COMM", "PROD"}
                .Columns(COLUMN_NAME).Width = 65
            Next
        End With

        grdSOTORDRX.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDRX.DisplayLayout.Bands(0)
            If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") _
            Or (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then
            Else
                .Columns("CUST_NAME").Hidden = True
            End If

            For Each COLUMN_NAME In New String() {"SREP_CODE", "ORDR_TYPE", "ORDR_GROUP_NO", "CUST_CODE"}
                Dim gcol As UltraWinGrid.UltraGridColumn = grdSOTORDRX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                gcol.Header.Fixed = True
            Next
            For Each COLUMN_NAME In New String() {"STYLE_SHIP_WINDOW_DAYS", "ORDR_SHIP_DATE_PLUS", "STYLE_AT_ONCE_UNTIL", "STYLE_AT_ONCE_ACTIVE"}
                Dim gcol As UltraWinGrid.UltraGridColumn = grdSOTORDRX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.Orange
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
            Next

            .Columns("ORDR_GROUP_NO").CellAppearance.BackColor = Color.Beige
            '.Columns("ORDR_QTY").CellAppearance.BackColor = Color.Beige

            For Each COLUMN_NAME In New String() {"ORDR", "OPEN", "PICK", "ALLO", "SHIP", "CANC", "ORDERS", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "STYLE_COST", "ORDR_UNIT_PRICE_CALC", "GP_PCT"}
                Dim gcol As UltraWinGrid.UltraGridColumn = grdSOTORDRY.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                gcol.Header.Caption = .Columns(COLUMN_NAME).Header.Caption
                gcol.Width = .Columns(COLUMN_NAME).Width
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next

            For Each COLUMN_NAME In New String() {"SREP_CODE", "WHSE_CODE", "ORDR_TYPE_CODE", "ORDR_TYPE", "ORDR_GROUP_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_DATE_RECD"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.Yellow
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next

            For Each COLUMN_NAME In New String() {"ORDR", "OPEN", "PICK", "ALLO", "SHIP", "CANC", "ORDERS"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.Violet
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
                '  gcol.Width = 90
            Next

            For Each COLUMN_NAME In New String() {"ORDR_UNIT_PRICE_CALC", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next


            For Each COLUMN_NAME In New String() {"SUB", "INV_DATE"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.Pink
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next
            For Each COLUMN_NAME In New String() {"STYLE_COST", "COST_SOURCE", "GP_PCT"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next
        End With

        With grdSOTORDRX.DisplayLayout.Bands(0)
            .Columns("PO_SEQ_MAX_WAIT").Header.Caption = "#PO"
            .Columns("SHIP_SEQ").Header.Caption = "#Shp"
            .Columns("RECD_SEQ").Header.Caption = "#Rec"
            .Columns("SHIP_DATE_PLUS").Header.Caption = "Ship+"
            .Columns("ERROR").Header.Caption = "Allocation Error"
            .Columns("QTY_ALLO_0").Header.Caption = "AtOnce"
            For Each CN As String In New String() {"PO_SEQ_MAX_WAIT", "SHIP_SEQ", "RECD_SEQ", "SHIP_DATE_PLUS", "ERROR",
                                                   "QTY_ALLO_0", "QTY_ALLO_1", "QTY_ALLO_2", "QTY_ALLO_3", "QTY_ALLO_4",
                                                   "QTY_ALLO_5", "QTY_ALLO_6", "QTY_ALLO_7", "QTY_ALLO_8", "QTY_ALLO_9"}
                With .Columns(CN)
                    .Width = 60
                    .Header.Appearance.BackColor2 = Color.Aqua
                    .Header.Appearance.BackColor = Color.White
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
            Next

            .Columns("SHIP_DATE_PLUS").Width = 90
            .Columns("ERROR").Width = 150
        End With

        grdSOTSUPPA.DataSource = dst.Tables("SOTSUPPA")
        With grdSOTSUPPA.DisplayLayout.Bands(0)
            .Columns("PO_SHIPMENT_NO").Hidden = True
            .Columns("PO_ORDER_NO").Hidden = True

            .Columns("PO_ARRIVAL_DATE").Header.Caption = "Arr"
            .Columns("PO_ARRIVAL_DATE").Width = 60
            .Columns("PO_ARRIVAL_DATE").Format = "MM/dd"

            .Columns("PO_ARRIVAL_DATE_PLUS").Header.Caption = "Arr+"
            .Columns("PO_ARRIVAL_DATE_PLUS").Width = 60
            .Columns("PO_ARRIVAL_DATE_PLUS").Format = "MM/dd"
            .Columns("PO_ARRIVAL_DATE_PLUS").Hidden = True

            .Columns("PO_QTY").Header.Caption = "+Qty"
            .Columns("PO_QTY").Width = 60
            .Columns("PO_QTY").Format = "#,##0"
            .Columns("PO_QTY").Hidden = True

            .Columns("PO_QTY_USED").Header.Caption = "Allo"
            .Columns("PO_QTY_USED").Width = 60
            .Columns("PO_QTY_USED").Format = "#,##0"
            .Columns("PO_QTY_USED").Hidden = True

            .Columns("PO_QTY_LEFT").Header.Caption = "+Ava"
            .Columns("PO_QTY_LEFT").Width = 60
            .Columns("PO_QTY_LEFT").Format = "#,##0"

            .Columns("PO_QTY_CUM").Header.Caption = "Cum"
            .Columns("PO_QTY_CUM").Width = 60
            .Columns("PO_QTY_CUM").Format = "#,##0"

            .Columns("PO_SEQ").Hidden = True
        End With


        grdSOTORDRY.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDRY.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC", "STYLE_CODE_SUB"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.Turquoise
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next
            For Each COLUMN_NAME In New String() {"ORDR_NO", "ORDR_CUST_PO", "CUST_STORE_NO"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.Orange
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next
            For Each COLUMN_NAME In New String() {"ORDR", "OPEN", "PICK", "ALLO", "SHIP", "CANC", "ORDERS"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.Violet
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next

            For Each COLUMN_NAME In New String() {"ORDR_UNIT_PRICE_CALC", "ORDR_AMT"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Format = .Columns(COLUMN_NAME).Format
            Next
        End With


        grdICTSTDQ1.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
        With grdICTSTDQ1.DisplayLayout.Bands(0)
            .Columns("WHSE_CODE").HiddenWhenGroupBy = DefaultableBoolean.True
            .Columns("STATUS_DATE").Format = "MM/dd/yy"
            .Columns("STATUS_DATE").Width = 70 ' 75
            .Columns("QTY_ATS").Width = 55 ' 60
            .Columns("QTY_ATS_CUM").Width = 55 ' 60
            .Columns("QTY_ATS").Hidden = False
            .Columns("QTY_ATS_CUM").Hidden = False
            .Columns("QTY_ATS").Header.Caption = "+Avail"
            .Columns("QTY_ATS_CUM").Header.Caption = "Cum ATS"

            If ASCMAIN1.CLIENT = "NOT RGI" Then
                .Columns("QTY_ATS").Hidden = True
                .Columns("QTY_PLUS").Hidden = False
                .Columns("QTY_ATS_CUM").Hidden = False ' True
                .Columns("QTY_PLUS_CUM").Hidden = True ' False
                .Columns("QTY_PLUS").Width = 55
                .Columns("QTY_PLUS_CUM").Width = 55
                .Columns("QTY_ATS_CUM").Width = 55
            End If
        End With

        grdSOTALLO1.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
        grdSOTALLO1.DisplayLayout.UseFixedHeaders = True
        With grdSOTALLO1.DisplayLayout.Bands(0)
            .Columns("WHSE_CODE").HiddenWhenGroupBy = DefaultableBoolean.True
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If New String() {"ORDR_DEMAND_DATE", "ORDR_PRIORITY_DATE", "ORDR_PRIORITY", "ORDR_RELEASE", "ORDR_BACKORDER"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            For Each COLUMN_NAME In New String() {"SD_QTY_ALLO_CUR", "SD_QTY_ALLO_FUT", "SD_QTY_ALLO_CXL"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.CellAppearance.BackColor = Color.LightYellow
            Next
            For Each COLUMN_NAME In New String() {"BALANCE", "SD_DATE_X"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.CellAppearance.BackColor = Color.LightGray
            Next
            For Each COLUMN_NAME In New String() {"SD_QTY", "SD_QTY_ALLO", "SD_QTY_ALLO_CUR", "SD_QTY_ALLO_FUT", "SD_QTY_ALLO_CXL", "BALANCE"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Format = "#,##0"
            Next
        End With

        With grdWHTLOCB1.DisplayLayout.Bands(0)
            .Columns("LOCATION_CODE").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("LOCATION_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LOCATION_QTY_WAVE").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LOCATION_QTY_AVAIL").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
        End With

        Absx1.txtFor("RANGE_STYLE_CODE").Visible = False
        Absx1.txtFor("RANGE_STYLE_DESC").Visible = False
        Absx1.txtFor("RANGE_STYLE_CODE").Top = Absx1.txtFor("STYLE_CODE").Top
        Absx1.txtFor("RANGE_STYLE_DESC").Top = Absx1.txtFor("STYLE_DESC").Top

        'Set_cmbYP("RYP_TO", ASCMAIN1.CYP, -36, 12, 0)
        'Set_cmbYP_Child("RYP_FROM", 12, "RYP_TO", 11)

        'Set_cmbYP("RYP_FROM", Mid(ASCMAIN1.CYP, 1, 4) & "01", -36, 12, 0)
        'Set_cmbYP_Child("RYP_TO", 12, "RYP_FROM", Val(Mid(ASCMAIN1.CYP, 5, 2)) - 1)

        Set_cmbYP("RYP_FROM", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP("RYP_TO", ASCMAIN1.CYP, -60, 0, 0)
        '  Set_cmbYP_Child("RYP_TO", 12, "RYP_FROM", 0)

        'cmbYP_From.ActiveRow = cmbYP_From.Rows(6)
        'cmbYP_To.ActiveRow = cmbYP_To.Rows(6)
        Set_cmbYP("COST_OPS_YYYYPP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -12, 0, 0)

        edi850cust = TAC.SOCMAIN1.Get_EDI_Custs("850")
        Toggle_ALLOCF()

        Absx1.numFor("STYLE_COST").Visible = (ASCMAIN1.USER_SECURITY_CODEs.Contains("X2"))
        lblCost.Visible = (ASCMAIN1.USER_SECURITY_CODEs.Contains("X2"))
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("GP_PCT").Hidden = Not lblCost.Visible

        grdICTSTATA.DisplayLayout.Bands(0).Columns("OTS_WIP").Hidden = True
        grdICTSTATA.DisplayLayout.Bands(1).Columns("OTS_WIP").Hidden = True

        Set_Read_Only(grpICTSTYL1, True)
        Set_Read_Only(grpProcurement, True)


        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            optViewBy.Tag = "X"
            optViewStyles.Value = "S"
            optViewBy.Value = "SCW"
            optViewBy.Tag = ""
        End If

        '   Setup_Recent()


        With grdPOTORDRX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"STYLE_ARRIVAL_BUFFER_DAYS", "STYLE_AT_ONCE_UNTIL", "STYLE_AT_ONCE_ACTIVE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If
            Next
        End With

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            UltraExplorerBar1.Groups("PO / In Transit").Text = "WIP / In Transit"
            tabMain.Tabs("PO / In Transit").Text = "WIP / In Transit"
            grdICTSTATA.DisplayLayout.Bands(0).Columns("ON_ORDER").Header.Caption = "WIP"
            With grdPOTORDRX.DisplayLayout.Bands(0)
                .Columns("PO_DATE_SHIP_BY_REQ").Header.Caption = "Orig Ship By"
                .Columns("PO_DATE_SHIP_BY").Header.Caption = "Revd Ship By"
            End With

            lblSIZE_CODE.Visible = True
            Absx1.txtFor("SIZE_CODE").Visible = True
        End If

        If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
            With grdPOTORDRX.DisplayLayout.Bands(0)
                .Columns("PO_SPEC_ORDR_NO").Hidden = True
                .Columns("FACTORY_CODE").Hidden = True
            End With

        End If

        ASCMAIN1.Add_Value_List(grdICTATOP1, "ORDR_TYPE", Nothing, New String() {":", "R:Reservation", "O:Sales Order"})
        ASCMAIN1.Add_Value_List(grdICTATOP2, "PS_CODE", Nothing, New String() {":", "P:PO", "S:Shipment"})

        ASCMAIN1.Add_Value_List(grdICTCOSTL, "TRAN_TYPE", Nothing, New String() {":", "B:Baseline", "M:Markdown", "R:Receipt", "J:Cost Adj", "Z:Zero Lot"})

        ASCMAIN1.Add_Value_List(grdSOTALLO1, "ORDR_RELEASE", Nothing, New String() {":", "H:Hold if Short", "C:Cancel if Short", "S:Ship Short", "X:X-Cancel Open"})
        'ASCMAIN1.Add_Value_List(grdSOTALLO1, "ORDR_PRIORITY", Nothing, New String() {":", "1:1", "2:2", "3:3", "4:4", "5:5", "6:6", "7:7", "8:8", "9:9"})
        ASCMAIN1.Add_Value_List(grdSOTALLO1, "ORDR_PRIORITY")

        ASCMAIN1.Add_Value_List(grdICTCOST1, "TRAN_TYPE", Nothing, New String() {":", "M:Markdown", "B:Baseline", "J:Adjustment", "Z:Zero Cost"})

        grpATS.Dock = DockStyle.Fill
        cbeWHSE_CODE.DataSource = ASCDATA1.GetDataTable("Select WHSE_CODE,WHSE_DESC from ICTWHSE1 order by WHSE_CODE")
        cbeWHSE_CODE.Value = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")

        splPOTORDRX.Panel2Collapsed = True

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns("STYLE_COST_FIRST").Hidden = False
            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns("STYLE_COST_FIRST").Header.Caption = "1st Cost"

            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns("QTY_NETA").Header.Caption = "Fut Ava"

            grdICTSTATA.DisplayLayout.Bands(0).Columns("NET_POS").Header.Caption = "Fut Ava"
            grdICTSTATA.DisplayLayout.Bands(0).Columns("THEME_DESC").Header.Caption = "Theme"
            Dim last_pos As Integer = grdICTSTATA.DisplayLayout.Bands(0).Columns("UPC_CODE").Header.VisiblePosition
            grdICTSTATA.DisplayLayout.Bands(0).Columns("THEME_DESC").Header.SetVisiblePosition(last_pos + 1, False)
        Else
            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns("CARTONS_PER_UNIT").Hidden = True
            grdICTSTATA.DisplayLayout.Bands(0).Columns("THEME_DESC").Hidden = True
            ' grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns("STYLE_COST_FIRST").Hidden = True
        End If

        grdICTSTATA.DisplayLayout.Bands(0).Columns("COMM").Hidden = True
        grdICTSTATA.DisplayLayout.Bands(0).Columns("PROD").Hidden = True
        grdICTSTATA.DisplayLayout.Bands(1).Columns("COMM").Hidden = True
        grdICTSTATA.DisplayLayout.Bands(1).Columns("PROD").Hidden = True

        Setup_QuoteSheet()

        cmdMulti.Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            grdSOTALLO1.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_ORDER_NO").Hidden = True
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("VEND_CODE").Hidden = True

            lblROYALTY_CODE.Visible = False
            txtROYALTY_CODE.Visible = False
            txtROYALTY_DESC.Visible = False

        Else
            lblSUB_BODY_CODE.Visible = False
            txtSUB_BODY_CODE.Visible = False
            txtSUB_BODY_DESC.Visible = False


            lblCMT_NO.Visible = False
            txtCMT_NO.Visible = False
        End If


        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            lblFABRIC_CODE.Visible = False
            txtFABRIC_CODE.Visible = False
            txtFABRIC_DESC.Visible = False

            lblROYALTY_CODE.Visible = False
            txtROYALTY_CODE.Visible = False
            txtROYALTY_DESC.Visible = False

            Absx1.optFor("FASHION_PROMO").Visible = False

            lblFEPRICE.Visible = True
            lblFEPRICEX.Visible = True

            lblFEMIXPRICE.Visible = True
            lblFEMIXPRICEX.Visible = True

            lblFDMIXPRICE.Visible = True
            lblFDMIXPRICEX.Visible = True

            lblFDPRICE.Visible = True
            lblFDPRICEX.Visible = True

            lblROYALTY_CODE.Visible = True
            txtROYALTY_CODE.Visible = True
            txtROYALTY_DESC.Visible = True
            lblROYALTY_CODE.Text = "Designer"
        Else
            lblUnitsInner.Text = "Units / Inner"
            lblInner2.Visible = False
            numInner2.Visible = False
            lblUnitsInner2.Visible = False

            tabMain.Tabs("Pricing && Availability").Text = "Available to Sell"
            splPA.Panel1Collapsed = True

            lblROYALTY_CODE.Visible = False
        End If

        grdICTSTYL3.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")
        lblSTYLE_MATL_DESC.Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")
        Absx1.txtFor("STYLE_MATL_DESC").Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")
        If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
            Absx1.numFor("STYLE_COST").Visible = False
            lblCost.Visible = False
        End If

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            lblSTYLE_GROUP_CODE.Visible = True
            txtSTYLE_GROUP_CODE.Visible = True
            txtSTYLE_GROUP_DESC.Visible = True

            lblFABRIC_CODE.Visible = False
            txtFABRIC_CODE.Visible = False
            txtFABRIC_DESC.Visible = False

            lblROYALTY_CODE.Visible = True
            txtROYALTY_CODE.Visible = True
            txtROYALTY_DESC.Visible = True

            Absx1.optFor("FASHION_PROMO").Visible = False

            ASCMAIN1.Add_Value_List(grdICTSTYL1_Recent, "STYLE_GROUP_CODE")

            With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
                Dim C As Integer = .Columns("QTY_NETA").Header.VisiblePosition
                '"STYLE_COST", "STYLE_COST_EXT",
                For Each COLUMN_NAME As String In New String() {"CARTON_PACK_QTY", "INNER_PACK_QTY", "STYLE_COST_LDP", "STYLE_COST_ELC", "STYLE_COST_CUM", "SALES_DIVISION_CODE", "STYLE_GROUP_CODE", "VEND_CODE", "CUST_CODE"}
                    C += 1
                    .Columns(COLUMN_NAME).Header.SetVisiblePosition(C, False)
                Next

                C = .Columns("PURCH_NOTES").Header.VisiblePosition
                .Columns("COLOR_CODE").Header.SetVisiblePosition(C, False)
                .Columns("COLOR_DESC").Header.SetVisiblePosition(C, False)

                'Dim pos_QTY_NETA As Integer = .Columns("QTY_NETA").Header.VisiblePosition
                '.Columns("CARTON_PACK_QTY").Header.SetVisiblePosition(pos_QTY_NETA + 1, False)
                '.Columns("INNER_PACK_QTY").Header.SetVisiblePosition(pos_QTY_NETA + 2, False)

                '.Columns("VEND_CODE").Header.Appearance.BackColor = Drawing.Color.LightGray
                '.Columns("VEND_CODE").Header.Appearance.BackColor2 = Drawing.Color.Empty
            End With
        Else
            lblSTYLE_GROUP_CODE.Visible = False
            txtSTYLE_GROUP_CODE.Visible = False
            txtSTYLE_GROUP_DESC.Visible = False
        End If

        ' this option is not ready yet - DO NOT MAKE VISIBLE, AND LEAVE PINNED ON "1"
        ' optASL.Value = "1" ' SET IN DESIGNER
        'If ASCMAIN1.Running_in_VS Then
        '    optASL.Visible = True
        'Else


        'End If
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            optASL.Visible = False ' True
            optASL.Value = "0"
            lblSTYLE_GROUP_CODE.Text = "Family"
            lblSTYLE_GROUP_CODE.Visible = True
            txtSTYLE_GROUP_CODE.Visible = True
            txtSTYLE_GROUP_DESC.Visible = True
            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns.Item("STYLE_GROUP_CODE").Header.Caption = "Family"
        Else
            optASL.Visible = False
        End If


        'Ship+ = Number of days to add to a Ship Date to calculate a more lenient Ship Date
        'ETA+ = Number of days to add to the ETA to get a more conservative ETA
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("QTY_ALLO_0").Header.ToolTipText = "ETA+ = Number of days to add to the ETA to get a more conservative ETA"


        With grdICTSTATA.DisplayLayout.Bands(0)
            .Columns("ON_HAND").Header.ToolTipText = "On Hand"
            .Columns("ON_ORDER").Header.ToolTipText = "Overseas PO (does Not include In Transit"
            .Columns("TRAN").Header.ToolTipText = "PO Qty In Transit (Shipped)"
            .Columns("OPEN").Header.ToolTipText = "Open Customer Orders (does Not include In Pick)"
            .Columns("PICK").Header.ToolTipText = "Customer Orders In Pick"
            .Columns("ALLO").Hidden = True
            .Columns("OTS_INV").Header.ToolTipText = "Open To Ship (On hand less In Pick)"
            .Columns("NET_POS").Header.ToolTipText = "Net Available after All Supply & Demand"
        End With

        With grdICTSTATA.DisplayLayout.Bands(1)
            .Columns("ALLO").Hidden = True
        End With

        ' THESE OPTIONS ARE FOR VAN, BUT NOT TESTED YET
        optDetails.Visible = False
        chkSR.Visible = False

        Bind_Controls(splQuoteContainer.Panel2, "ICTQUOT1")
        Bind_Controls(splQuoteSheet.Panel1, "ICTQUOT1")
        Modes_Quote_Sheet(False)

        clsSOTORDR1 = New TAC.SOCORDR1(Me)

        grdICTTRANX.DisplayLayout.Bands(0).Columns("RUNNING_BALANCE").Hidden = True
        chkShowAll.Visible = False

        grpICTWHSES.BringToFront()
        grdICTWHSES.BringToFront()
        Setup_Recent()

        grdICTDUTY4.Dock = DockStyle.None
        grdICTDUTY4.Dock = DockStyle.Fill
        grdICTSTYV1.Dock = DockStyle.None
        grdICTSTYV1.Dock = DockStyle.Fill

        SO_PARM_SHIP_WINDOW_DAYS = Val(ROWs("SOTPARM1").Item("SO_PARM_SHIP_WINDOW_DAYS") & "")
        SO_PARM_ARRIVAL_BUFFER_DAYS = Val(ROWs("SOTPARM1").Item("SO_PARM_ARRIVAL_BUFFER_DAYS") & "")
        SO_PARM_RELEASE_AT_ONCE = ROWs("SOTPARM1").Item("SO_PARM_RELEASE_AT_ONCE") & ""

        numETA_PLUS.Value = SO_PARM_ARRIVAL_BUFFER_DAYS
        numSHIP_PLUS.Value = SO_PARM_SHIP_WINDOW_DAYS

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            optStockNon.Visible = True
        Else
            optStockNon.Visible = False
        End If

        tabStyles.Tabs("Overages && Shortages").Visible = (ASCMAIN1.CLIENT = "RGI")
        tabStyles.Tabs("At-Once").Visible = (ASCMAIN1.CLIENT = "RGI")

        If (ASCMAIN1.CLIENT = "RGI") Then
        Else
            chkAutoAllocate.Checked = False
            chkAutoCalculate.Checked = False
        End If
        'With chkOverbooked
        '    .Appearance.ForeColor = System.Drawing.Color.White
        '    .Appearance.BackColor = System.Drawing.Color.FromArgb(98, 160, 232)
        '    .Appearance.BackColor2 = System.Drawing.Color.FromArgb(83, 115, 191)
        '    .Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        'End With
        MakeTransparent(chkOverbooked)
        MakeTransparent(chkAnyAva2Ship)
        MakeTransparent(chkAllStyles)

        ASCMAIN1.Add_Value_List(grdSOTSUPPX, "STYLE_COLOR_STATUS")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"
                If multi_style Then

                Else
                    If optSelectBy.Value = "S" Then
                        STYLE_CODE = Absx1.txtFor("STYLE_CODE").Text
                        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)

                        If rowICTSTYL1 Is Nothing Then
                            EMsg &= "Invalid Value entered For Style Code (" & STYLE_CODE & ")"
                        End If

                        If EMsg = "" Then
                            STYLEs.Clear()
                            STYLEs.Add(STYLE_CODE)
                        End If

                    ElseIf optSelectBy.Value = "R" Then
                        RANGE_STYLE_CODE = Absx1.txtFor("RANGE_STYLE_CODE").Text
                        rowICTRSTY1 = LookUp("ICTRSTY1", RANGE_STYLE_CODE)

                        If rowICTSTYL1 Is Nothing Then
                            EMsg &= "Invalid Value entered For Range Style Code (" & RANGE_STYLE_CODE & ")"
                        End If

                        If EMsg = "" Then
                            STYLEs.Clear()
                            ASCMAIN1.sql = "Select STYLE_CODE from ICTRSTY2 where RANGE_STYLE_CODE = : PARM1"
                            For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {RANGE_STYLE_CODE}).Rows
                                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                                STYLEs.Add(STYLE_CODE)
                            Next
                        End If
                    End If
                End If

            Case "Print"
                'If Not grdSOTALLO1.Visible Then
                '    EMsg &= vbCr & "You Must Allocate before Printing"
                'End If

            Case "Cost Update"


            Case "Edit Quote Sheet"
                If grdICTQUOTX.ActiveRow Is Nothing OrElse Not grdICTQUOTX.ActiveRow.IsDataRow Then
                    EMsg &= "You must select (or double-click) a quote appearing in the grid in order to Edit it"
                Else
                    QUOTE_NO = grdICTQUOTX.ActiveRow.Cells("QUOTE_NO").Value
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTQUOT1", QUOTE_NO) Then Exit Sub
                End If

            Case "Cancel Quote Sheet"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete Quote Sheet"
                If ASCMAIN1.USER_ID <> rowICTQUOT1.Item("INIT_OPER") & "" Then
                    EMsg &= vbCr & "Only " & rowICTQUOT1.Item("INIT_OPER") & " may Delete this Quote"
                End If

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete this Quote",
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Print Quote Sheet", "email Quote Sheet"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Code Specified"
                    End If
                End If
                If dst.Tables("ICTQUOT2").Rows.Count = 0 Then
                    EMsg &= vbCr & "No Styles on the Quote Sheet"
                End If

            Case "Save Quote Sheet"
                If txtQUOTE_DESC.Text = "" Then
                    EMsg &= vbCr & "Please enter a Description for the Quote Sheet"
                End If

            Case "Edit At-Once"

                If Not ASCMAIN1.Logical_Lock("ICTATOP1", "*") Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("ICTATOP2", "*") Then Exit Sub

            Case "Update At-Once"

                ' check anything?"


            Case "Cancel At-Once"



        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Select"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

                If multi_style Then
                Else
                End If

                If grdICTSTATA.Rows.Count = 0 Then
                    MsgBox("No Status Records to Display", MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Click_Command("Done")
                End If

            Case "Refresh"
                Dim STYLE_CODE As String = Me.STYLE_CODE
                Click_Command("Done")
                EntryMode = "R"
                Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
                Click_Command("Select")

            Case "Print"
                Print_Record()

            Case "Cost Update"

                If MsgBox("This Option will Update the Sales History for" _
                          & vbCrLf & " All Styles which were updated with 0.00 or 0.01 costs" _
                          & vbCrLf & " to the Current Value of the Style Cost" _
                          & vbCrLf & " in the Style Master File" _
                          & vbCrLf & vbCrLf & "Run Cost Update",
                           MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                           "Verification") = MsgBoxResult.Yes Then

                    BeginTrans()
                    ASCDATA1.ExecuteSQL("BEGIN ICPCOSTX_GABE; END; ")
                    CommitTrans()
                    MsgBox("Your Cost Update is Complete." _
                           & vbCrLf & "Have a Nice Day.",
                            MsgBoxStyle.OkOnly, "Anna's Special Message Box")
                End If
            Case "Import Markdowns"
                Dim openFileDialog1 As New OpenFileDialog
                'openFileDialog1.InitialDirectory = "C:\ABS\icons\iconexperience\48x48\plain\"
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                'openFileDialog1.FilterIndex = 1
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim FILENAME As String = openFileDialog1.FileName
                    Import_Markdowns(FILENAME)
                End If
            Case "Done"
                Mode_Settings(False)

            Case "Update Pre-Allocation"

            Case "Cancel Pre-Allocation"
                frmPreAllocate.Visible = False
                tabMain.Visible = True

            Case "Delete Pre-Allocation"


                'Case "Use Rank"

                '    Dim t As Int64 = Val(dst.Tables("SOTPREA1").Compute("SUM(ORDR_QTY_OPEN)", "") & "")
                '    Dim q As Int64 = 0

                '    For Each row As DataRow In dst.Tables("SOTPREA1").Select("", "CUST_RANK")
                '        If t <> 0 Then
                '            q = Val(row.Item("ORDR_QTY_OPEN") & "")
                '            If q > t Then
                '                q = t
                '            End If
                '            t = t - q
                '            row.Item("ORDR_QTY_PRE_ALLO") = q
                '        End If
                '    Next

            Case "Allocate"
                Allocate()

            Case "Pre-Allocate"
                PreAllocate()

            Case "Find Style by Attribute"
                'Dim STYLE_CODE_selected As String = ""
                Using F As New TAC.ICFATTR2(Me)
                    F.rbadDir = ASCMAIN1.Folders("Work") & "RBAD\"
                    F.IMAGES_FOLDER = "\\192.168.110.221\Shared\rich\MASTER ITEM PHOTO FOLDER\"
                    F.IMAGES_FOLDER = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR")
                    F.ShowDialog()
                    STYLE_CODE = F.STYLE_CODE
                End Using
                If STYLE_CODE <> "" Then
                    Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
                    Click_Command("Select")
                End If

            Case "Integrity Check"

                Integrity_Check()


            Case "Print Quote Sheet", "email Quote Sheet"

                'Update_Record_TDA("ICTQUOT2", "1=1")
                'Dim rowICTQUOT1 As DataRow = dst.Tables("ICTQUOT1").Rows(0)
                'rowICTQUOT1.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
                'rowICTQUOT1.Item("CUST_NAME") = Absx1.txtFor("CUST_NAME").Text
                'rowICTQUOT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                'rowICTQUOT1.Item("INIT_DATE") = DATETIME_STAMP
                'rowICTQUOT1.Item("QUOTE_DESC") = Absx1.txtFor("QUOTE_DESC").Text

                Synch_TABLE_NAME("ICTQUOT1")
                Print_Report_Begin()
                CR_params.Add("CHKOMITAVAIL", IIf(chkOmitAvail.Checked, "1", "0"))
                CR_params.Add("CHKOMITPRICE", IIf(chkOmitPrice.Checked, "1", "0"))
                CR_params.Add("CHKOMITPRICE2", IIf(chkOmitPrice2.Checked, "1", "0"))

                Dim RPT As String = "ICRQUOT1"
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    RPT = "ICRQUOT2"
                End If
                If chk1perPage.Checked Then
                    RPT = "ICRQUOTN"
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        RPT = "ICRQUOTV"
                    End If
                End If

                If eItemKey = "email Quote Sheet" Then
                    Dim tempFileName As String = rowICTQUOT1.Item("QUOTE_NO")

                    Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)
                    ' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
                    Print_Report_End(, True)
                    email_Quote(tempFileName)
                Else
                    Generate_Report(RPT, "Quote Sheet")
                    Print_Report_End()
                End If

            Case "Clear Quote Sheet"
                dst.Tables("ICTQUOT2").Rows.Clear()
                Setup_Style_Quoted()
                ' txtQUOTE_DESC.Text = ""
                'Absx1.txtFor("CUST_CODE").Text = ""

            Case "New Quote Sheet"
                QuoteEntryMode = "N"
                Load_Record_Quote_Sheet()
                Modes_Quote_Sheet(True)

            Case "Edit Quote Sheet"
                QUOTE_NO = grdICTQUOTX.ActiveRow.Cells("QUOTE_NO").Value
                QuoteEntryMode = "E"
                Load_Record_Quote_Sheet()
                Modes_Quote_Sheet(True)

            Case "Cancel Quote Sheet"
                Modes_Quote_Sheet(False)
                Refresh_Documents()

            Case "Delete Quote Sheet"
                BeginTrans()
                ASCDATA1.ExecuteSQL("Delete from ICTQUOT1 where QUOTE_NO = '" & QUOTE_NO & "'")
                ASCDATA1.ExecuteSQL("Delete from ICTQUOT2 where QUOTE_NO = '" & QUOTE_NO & "'")
                CommitTrans("Quote Sheet " & QUOTE_NO & " has been Deleted")

                Modes_Quote_Sheet(False)
                Refresh_Documents()

            Case "Save Quote Sheet"
                BeginTrans()
                Update_Record_TDA("ICTQUOT1", "QUOTE_NO = '" & QUOTE_NO & "'")
                Update_Record_TDA("ICTQUOT2", "QUOTE_NO = '" & QUOTE_NO & "'")
                CommitTrans("Quote Sheet " & QUOTE_NO & " has been Saved")

                Modes_Quote_Sheet(False)
                Refresh_Documents()


            Case "Refresh At-Once"
                Refresh_AtOnce()

            Case "Edit At-Once"
                Refresh_AtOnce()
                Modes_At_Once(True)

            Case "Update At-Once"
                BeginTrans()
                Update_Record_TDA("ICTATOP1")
                Update_Record_TDA("ICTATOP2")
                CommitTrans("At-Once Records have been Updated")
                Modes_At_Once(False)
                Refresh_AtOnce()


            Case "Cancel At-Once"

                Modes_At_Once(False)
                Refresh_AtOnce()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Select").Settings.Enabled = not_iScreenMode
                .Items("Refresh").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode

                .Items("Allocate").Settings.Enabled = iScreenMode
                .Items("Pre-Allocate").Settings.Enabled = iScreenMode
                .Items("Allocate").Visible = False
                .Items("Pre-Allocate").Visible = False

                .Items("Cost Update").Settings.Enabled = iScreenMode

                .Items("Cost Update").Visible = (ASCMAIN1.USER_ID = "wjz" Or ASCMAIN1.USER_ID = "gabe")

                .Items("Import Markdowns").Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") And ASCMAIN1.USER_SECURITY_CODEs.Contains("X5")

                '.Items("New Quote Sheet").Visible = (splQuoteMain.Panel2Collapsed)
                .Items("Find Style by Attribute").Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGO" Or ASCMAIN1.DBS_SERVER = "RGO")
                If ScreenMode Then
                    .Items("Find Style by Attribute").Text = "Find Substitutes"
                Else
                    .Items("Find Style by Attribute").Text = .Items("Find Style by Attribute").Key
                End If

                .Items("Integrity Check").Visible = Not ScreenMode And ASCMAIN1.Running_in_VS

                UltraExplorerBar1.Groups("At-Once").Visible = (ASCMAIN1.CLIENT = "RGI") And Not ScreenMode And (tabStyles.SelectedTab.Key = "At-Once")

            End With

            .Groups("Quote Sheet").Visible = Not ScreenMode
            .Groups("View Styles").Visible = Not ScreenMode

            If Not ScreenMode Then Setup_QuoteSheet()

            If Not ScreenMode Then
                .Groups("Style Image").Visible = False
                .Groups("Available by Date").Visible = False
                .Groups("Orders").Visible = False
                '.Groups("Costs").Visible = False
                .Groups("Pre-Allocate").Visible = False
                .Groups("PO / In Transit").Visible = False
            Else
                grdICTSTATO.Visible = False
            End If
            '.Groups("Available by Date").Visible = ScreenMode '  AndAlso (tabMain.SelectedTab.Key = "Allocate")
            .Groups("Sales History").Visible = False
        End With

        ' optSelectBy.Visible = Not ScreenMode
        lblUPC_CODE.Visible = Not ScreenMode
        txtUPC_CODE.Visible = Not ScreenMode
        '   grpATS.Visible = False

        If Not ScreenMode Then Absx1.txtFor("STYLE_CODE").ReadOnly = False
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        ' grdICTSTYL1_Recent.Visible = Not ScreenMode
        'If Not Me.IsClosing Then
        tabStyles.Visible = Not ScreenMode
        If tabStyles.SelectedTab IsNot Nothing AndAlso tabStyles.SelectedTab.Key = "Overages && Shortages" Then
            ' stay there
        Else
            tabStyles.SelectedTab = tabStyles.Tabs(0)
        End If

        tabMain.Visible = ScreenMode
        SplitContainer1.Visible = ScreenMode
        ' End If


        If optSelectBy.Value = "M" Or Not ScreenMode Then spl.Panel1Collapsed = ScreenMode

        With grdICTSTATA.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Hidden = (optSelectBy.Value = "S")
            .Columns("STYLE_DESC").Hidden = (optSelectBy.Value = "S")
        End With

        Setup_tabMain()

        lblSTYLE_UOM.Visible = ScreenMode
        txtSTYLE_UOM.Visible = ScreenMode
        optSTYLE_STATUS.Visible = ScreenMode

        grdSOTINVHX.Visible = False
        grdSOTINVHY.Visible = False

        If ScreenMode Then
            dst.Tables("ASTSQLX1").Rows.Clear()

        Else
            Clear_Record()
            dst.Tables("ASTSQLX1").Rows.Clear()

            Modes_At_Once(False)

            'If ASCMAIN1.CLIENT = "RGI" Then
            With grdSOTORDRX.DisplayLayout.Bands(0)
                    For Each C As String In New String() {"RECD_SEQ", "SHIP_SEQ", "SHIP_DATE_PLUS", "ERROR", "QTY_ALLO_0"}
                        .Columns(C).Hidden = True
                        If C = "QTY_ALLO_0" Then
                            For I As Integer = 1 To 9
                                Dim CX As String = Replace(C, "0", Format(I, "0"))
                                .Columns(CX).Hidden = True
                            Next
                        End If
                    Next
                End With
            'End If

            'With grdSOTORDRX.DisplayLayout.Bands(0)
            '    For Each CN As String In New String() {"PO_SEQ_MAX_WAIT", "SHIP_SEQ", "RECD_SEQ", "SHIP_DATE_PLUS", "ERROR", _
            '                               "QTY_ALLO_0", "QTY_ALLO_1", "QTY_ALLO_2", "QTY_ALLO_3", "QTY_ALLO_4", _
            '                               "QTY_ALLO_5", "QTY_ALLO_6", "QTY_ALLO_7", "QTY_ALLO_8", "QTY_ALLO_9"}
            '        .Columns(CN).Hidden = True
            '    Next
            'End With
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            ShowZeroStatus = True
            tabMain.Tabs.Item("Promotions").Visible = True
        Else
            tabMain.Tabs.Item("Promotions").Visible = False
        End If


        grdICTSTATA.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = Not multi_style

        If multi_style Then
            grdICTSTATA.DisplayLayout.Bands(0).SortedColumns.Add("STYLE_CODE", False, True)
        End If
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        chkAutoCalculate.Checked = False

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        For Each STYLE_CODE In STYLEs ' ALWAYS 1 STYLE
            rowICTSTYL1 = Fill_Record("ICTSTYL1", STYLE_CODE)
            Dim rowICTSTYL1_RECENT As DataRow = dst.Tables("ICTSTYL1_RECENT").Rows.Find(STYLE_CODE)
            If rowICTSTYL1_RECENT Is Nothing Then
                rowICTSTYL1_RECENT = dst.Tables("ICTSTYL1_RECENT").NewRow
                rowICTSTYL1_RECENT.ItemArray = rowICTSTYL1.ItemArray
                dst.Tables("ICTSTYL1_RECENT").Rows.Add(rowICTSTYL1_RECENT)
            End If

            Dim sqlICTSTYC1 As String = " union " & vbCrLf _
            & "(Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE" & vbCrLf _
            & ", 0 BEG, 0 SHP, 0 RTN, 0 REC, 0 ADJ, 0 XFR, 0 PHY" & vbCrLf _
            & ", 0 ON_HAND, 0 ON_ORDER, 0 TRAN, 0 OPEN, 0 PICK, 0 ALLO, 0 COMM, 0 PROD, ICTSTYC1.UPC_CODE, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
            & " from ICTSTYC1 " & vbCrLf _
            & " where ICTSTYC1.STYLE_CODE = '" & STYLE_CODE & "')" & vbCrLf

            Fill_ICTSTATA_ICTSTATW(STYLE_CODE, sqlICTSTYC1)

            If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
                Fill_THEMES()
            End If
            'If multi_style Then
            'Else
            'End If
        Next

        Dim rowICTSTAT1_IMAGES As DataRow = dst.Tables("ICTSTAT1_IMAGES").Rows(0)
        rowICTSTAT1_IMAGES.Item("STYLE_CODE") = STYLE_CODE

        Fill_Records("ICTSTYL3", STYLE_CODE)
        'For Each rowICTSTYL3 As DataRow In dst.Tables("ICTSTYL3").Select("")
        '    rowICTSTYL3.Item("SEL") = "1"
        'Next

        Fill_Records("ICTSTYV1", STYLE_CODE)

        SHOW_DUTY_EXCEPTIONS()

        numETA_PLUS.Value = SO_PARM_ARRIVAL_BUFFER_DAYS
        numSHIP_PLUS.Value = SO_PARM_SHIP_WINDOW_DAYS

        EnforceConstraints(True)

        Setup_STATA(True)
        Toggle_Show_Cost()

        If tabMain.SelectedTab.Key = "Allocate" Then
            Allocate()
        End If


        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Dim BTB_Price As New FEFDPrice(Me, STYLE_CODE, 1)
            lblFEPRICE.Text = Format(BTB_Price.FEPrice, "#.00")
            lblFEMIXPRICE.Text = Format(BTB_Price.FEMixPrice, "#.00")
            lblFDMIXPRICE.Text = Format(BTB_Price.FDMixPrice, "#.00")
            lblFDPRICE.Text = Format(BTB_Price.FDPrice, "#.00")
            BTB_Price = Nothing

            Dim ECOM_MSG As String = TAC.TACMAIN1.getEcomInfo(Me, STYLE_CODE)
            If ECOM_MSG.Length > 0 Then
                lblEcomStyle.Visible = True
                Dim TTI As New UltraWinToolTip.UltraToolTipInfo
                If Not IsNothing(TTM.GetUltraToolTip(lblEcomStyle)) Then
                    TTI.ToolTipTitle = "E-Commerce Information:"
                    TTM.AutoPopDelay = 20000
                    TTI.ToolTipTextFormatted = ECOM_MSG
                    TTM.SetUltraToolTip(lblEcomStyle, TTI)
                Else
                    TTI.ToolTipTextFormatted = ECOM_MSG
                End If
            Else
                lblEcomStyle.Visible = False
            End If

            ShowPromo(STYLE_CODE)
            Fill_Records("ICTPROMX", STYLE_CODE)
            Sort_grdColumns(grdICTPROMX, "PROMO_START_DATE", False)
        End If

        STYLE_CLASS_CODE = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
        Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)
        If rowICTCLAS1 IsNot Nothing AndAlso rowICTCLAS1.Item("STYLE_CLASS_RELEASE_ATONCE") & "" = "1" Then
            lblAtOnceEligible.Text = "YES"
            lblAtOnceEligible.Appearance.ForeColor = Color.Green
            blnATONCE = (SO_PARM_RELEASE_AT_ONCE = "1")
        Else
            lblAtOnceEligible.Text = "NO"
            lblAtOnceEligible.Appearance.ForeColor = Color.Red
            blnATONCE = False
        End If

        If (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Or (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") Then
            grdICTSTATA.Rows.ExpandAll(True)
        End If


        If ASCMAIN1.CLIENT = "VAN" Then
            Dim SQCs As String = TAC.ICCMAIN1.Get_SIZEs_and_QTYs_and_COLORs(Me, STYLE_CODE)
            Absx1.txtFor("SIZE_SCALE").Text = SQCs
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub Fill_THEMES()
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("ICTSTYC1.STYLE_CODE,")
        sql.AppendLine("ICTSTYC1.COLOR_CODE,")
        sql.AppendLine("ICTSTYC1.THEME_CODE,")
        sql.AppendLine("ICTTHEME.THEME_DESC")
        sql.AppendLine("FROM ICTSTYC1, ICTTHEME")
        sql.AppendLine("WHERE ICTSTYC1.THEME_CODE = ICTTHEME.THEME_CODE")
        sql.AppendLine("AND NVL(ICTSTYC1.THEME_CODE,'NULL') <> 'NULL'")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        For Each rowICTSTATA As DataRow In dst.Tables("ICTSTATA").Select()
            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", rowICTSTATA.Item("STYLE_CODE").ToString, rowICTSTATA.Item("COLOR_CODE").ToString)
            Dim rowtbl As DataRow = tbl.Select(filter).FirstOrDefault
            If Not IsNothing(rowtbl) Then
                rowICTSTATA.Item("THEME_DESC") = rowtbl.Item("THEME_DESC").ToString
            End If
        Next
    End Sub

    Sub Setup_STATA(Optional initial_display As Boolean = False)
        STYLE_CODE = ""
        COLOR_CODE = ""

        Dim sqlw As String = ""
        If Not ShowZeroStatus Then
            sqlw = "ISNULL(BEG,0) <> 0 or ISNULL(ON_HAND,0) <> 0 or ISNULL(ON_ORDER,0) <> 0 or ISNULL(TRAN,0) <> 0 or ISNULL(OPEN,0) <> 0 or ISNULL(PICK,0) <> 0 or ISNULL(ALLO,0) <> 0 or ISNULL(COMM,0) <> 0 or ISNULL(PROD,0) <> 0"
        End If

        DirectCast(grdICTSTATA.DataSource, DataTable).DefaultView.RowFilter = sqlw
        dst.Tables("ICTSTATW").DefaultView.RowFilter = sqlw
        If grdICTSTATA.Rows.Count = 0 And initial_display And Not ShowZeroStatus Then
            DirectCast(grdICTSTATA.DataSource, DataTable).DefaultView.RowFilter = ""
            dst.Tables("ICTSTATW").DefaultView.RowFilter = ""
        End If
        Sort_grdColumns(grdICTSTATA, "STYLE_CODE,COLOR_CODE")
        Sort_grdColumns(grdICTSTATA, "WHSE_CODE", , 1)
        If STYLE_CODE = "" Then
            Setup_SC()
        End If
    End Sub

    Sub Setup_SC()
        Me.Cursor = Cursors.WaitCursor
        If grdICTSTATA.ActiveRow IsNot Nothing AndAlso grdICTSTATA.ActiveRow.IsDataRow Then
            STYLE_CODE = grdICTSTATA.ActiveRow.Cells("STYLE_CODE").Value
            COLOR_CODE = grdICTSTATA.ActiveRow.Cells("COLOR_CODE").Value
            'rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
            rowICTSTYL1 = Fill_Record("ICTSTYL1", STYLE_CODE)
            Setup_Tab()
        Else
            STYLE_CODE = ""
            COLOR_CODE = ""
        End If

        FetchImage(rowICTSTYL1)

        ' WHAT IF MULTIPLE STYLES SELECTED
        If chkSingleColor.Checked Then
            grdSOTINVHX.Visible = False
            grdSOTINVHY.Visible = False
        End If

        If tabMain.SelectedTab IsNot Nothing And tabMain.SelectedTab.Key = "Allocate" Then
        Else
            If STYLE_CODE <> STYLE_CODE_allocated Then AutoAllocate = False
        End If

        blnAtOnceChanged = True
        Setup_tabMain()

        Fill_Records("ICTCOSTL", New String() {STYLE_CODE, COLOR_CODE})
        Fill_Records("ICTCOST1", New String() {STYLE_CODE, COLOR_CODE})
        grdICTCOST1.Text = "Cost Adjustment Records for " & STYLE_CODE & ", Color " & COLOR_CODE

        Fill_Records("ICTCOSTA", New String() {STYLE_CODE, COLOR_CODE})
        Sort_grdColumns(grdICTCOSTA, "OPS_YYYYPP".ToLower)
        Dim dvw As DataView = DirectCast(grdICTCOSTA.DataSource, DataTable).DefaultView
        dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
        grdICTCOSTA.Text = "Style Cost History for " & STYLE_CODE & ", Color " & COLOR_CODE

        'grdICTCOSTL.Visible = False
        ' splCosts.Visible = False

        If chkSingleColor.Checked Then
            If tabMain.SelectedTab.Key = "Sales History" Then
                FetchShippedOrders()
            End If
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDRX", "SOTORDRY", "ICTSTYL1", "ICTSTATA", "ICTSTATB", "ICTSTATW", "ICTATOP1", "ICTATOP2",
             "ICTTRANX", "POTORDRX", "POTSHIP7", "POTSHIP8", "ICTCOSTL", "ICTDUTY4", "WHTLOCB1", "SOTSUPPA"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        blnAtOnceChanged = False

        If EntryMode <> "R" Then
            STYLEs.Clear()
            Absx1.txtFor("STYLE_CODE").Text = ""
            '  ShowZeroStatus = True
            optSelectBy.Value = "S"
        End If

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Maintain Codes/Dates"), UltraWinToolbars.StateButtonTool)
        If tlb_sbt.Checked Then
            tlb_sbt.Checked = False
        End If
        UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        Absx1.txtFor("UPC_CODE").Text = ""

        STYLE_CODE_allocated = ""
        Absx1.txtFor("STYLE_CODE").Focus()

        multi_style = False
        grdICTSTATA.DisplayLayout.Bands(0).SortedColumns.Clear()

        grdWHTLOCB1.Text = ""
        grdWHTINSTX.Text = ""

        lblPromo.Visible = False
        lblPromo.Text = ""
        btnShowPromo.Visible = False

    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        CommitTrans("Update Complete")
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "Select"
                optSelectBy.Value = "S"
                Absx1.txtFor("STYLE_CODE").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ICTSTYL1"
            E.COLUMN_NAME = "STYLE_CODE"
            E.CODE_VALUE = Absx1.txtFor("STYLE_CODE").Text
            E.DESC_VALUE = Absx1.txtFor("STYLE_DESC").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "ICTSTYL1"
        E.TABLE_KEY_CAPTION = "Style"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("STYLE_CODE").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("STYLE_DESC").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "A")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function


    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "SUPPLIER"
                sql_where = "VEND_TYPE = 'S'"
            Case "STYLE_CODE"
                Select Case optStockNon.Value
                    Case "S"
                        sql_where = "NVL(CUST_CODE,'NULL') = 'NULL'"
                    Case "N"
                        sql_where = "NVL(CUST_CODE,'NULL') <> 'NULL'"
                    Case Else
                        sql_where = ""
                End Select
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTLOCB1, "SSSBB", "Show Filter", "Show GroupBox", "Show 0 Qty", "Location Inquiry for Style", "Location Inquiry for Location")
        Load_Popup_Menu(grdSOTORDRX, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Sales Order Entry", "Sales Reservations Inquiry", "Customer Order Inquiry", "Cancel Open Quantity", "Edit Ship+")
        Load_Popup_Menu(grdICTTRANX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Show Transaction", "Show Order")
        Load_Popup_Menu(grdICTSTATA, "SSB", "Show Style/Colors w/Zero Status", "Cardview", "Style Masterfile")
        Load_Popup_Menu(grdPOTORDRX, "SSSBBSB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "Shipments Inquiry", "Show Cartons", "Edit At-Once")
        Load_Popup_Menu(grdICTQUOTX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTSUPPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTALLO1, "SSBBBBB", "Maintain Codes/Dates", "Show Allocation Cur/Fut/Cxl", "Pre-Allocate", "PO Inquiry", "Sales Order Inquiry", "Sales Order Entry", "Sales Reservation Inquiry")

        Load_Popup_Menu(grdICTSTYL1_Recent, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry", "Add to Quote Sheet")
        'Load_Popup_Menu(grdSOTINVHX, "B", "Sales Detail Report")
        Load_Popup_Menu(grdSOTINVHY, "BB", "Show Invoice", "Sales Order Inquiry")
        Load_Popup_Menu(grdICTQUOT2, "B", "Sequence as Shown")
        Load_Popup_Menu(grdICTCOSTA, "B", "Cost Maintenance")
        Load_Popup_Menu(grdWHTINSTX, "B", "Wave Inquiry")
        Load_Popup_Menu(grdICTWHSES, "B", "Uncheck All")
        Load_Popup_Menu(grdICTPRICX, "B", "Calculate With Royalty")
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

        If grd.Name = "grd" Then
            Exit Sub
        End If

        Select Case e.SourceControl.Name
            Case "grdSOTORDRX"
                Dim ORDR_TYPE As String = ""
                If grdSOTORDRX.ActiveRow IsNot Nothing Then
                    ORDR_TYPE = grdSOTORDRX.ActiveRow.Cells("ORDR_TYPE").Value
                End If
                tlb_btn = DirectCast(tlb_pop.Tools("Sales Order Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O")
                tlb_btn = DirectCast(tlb_pop.Tools("Sales Order Entry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O")
                tlb_btn = DirectCast(tlb_pop.Tools("Sales Reservations Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ORDR_TYPE = "R")

                tlb_btn = DirectCast(tlb_pop.Tools("Cancel Open Quantity"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grdSOTORDRX.Selected.Rows.Count > 0

                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow OrElse grd.ActiveRow.IsFilterRow Then
                    tlb_pop.Tools("Edit Ship+").SharedProps.Visible = False
                Else
                    Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Value & ""
                    tlb_pop.Tools("Edit Ship+").SharedProps.Visible = ASCMAIN1.USER_SECURITY_CODEs.Contains("AO") AndAlso WHSE_CODE = "MS"
                End If


            Case "grdSOTALLO1"
                If grdSOTALLO1.ActiveRow Is Nothing OrElse Not grdSOTALLO1.ActiveRow.IsDataRow Then
                Else
                    Dim RECORD_SUB_TYPE As String = grdSOTALLO1.ActiveRow.Cells("RECORD_SUB_TYPE").Value & ""
                    tlb_pop.Tools("PO Inquiry").SharedProps.Visible = (RECORD_SUB_TYPE = "S")
                    tlb_pop.Tools("Sales Order Inquiry").SharedProps.Visible = (RECORD_SUB_TYPE = "O")
                    tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = (RECORD_SUB_TYPE = "O")
                    tlb_pop.Tools("Sales Reservation Inquiry").SharedProps.Visible = (RECORD_SUB_TYPE = "R")
                    ' Permit Maintenance Option only if X2 security is included in User Profile
                    ' tlb_pop.Tools("Maintain Codes/Dates").SharedProps.Visible = ASCMAIN1.USER_SECURITY_CODEs.Contains("X2")
                    ' DISABLING THIS BECAUSE SOMEONE REMOVED X2 FROM JOHN
                    tlb_pop.Tools("Pre-Allocate").SharedProps.Visible = False
                End If

            Case "grdPOTORDRX"

                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow OrElse grd.ActiveRow.IsFilterRow Then
                    tlb_pop.Tools("Edit At-Once").SharedProps.Visible = False
                Else
                    Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Value & ""
                    tlb_pop.Tools("Edit At-Once").SharedProps.Visible = ASCMAIN1.USER_SECURITY_CODEs.Contains("AO") AndAlso WHSE_CODE = "MS"
                End If


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "Show Style/Colors w/Zero Status"
                    tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                    tlb_sbt.Tag = "X"
                    tlb_sbt.Checked = ShowZeroStatus
                    tlb_sbt.Tag = ""


            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show 0 Qty"
                Setup_grdWHTLOCB1()
            Case "Sequence as Shown"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Re-Sequencing by 10's")

                Dim SEQ As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdICTQUOT2.Rows
                    SEQ += 1
                    grow.Cells("SEQ").Value = SEQ
                    grow.Update()
                Next

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("")

            Case "Show Style/Colors w/Zero Status"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then Exit Sub
                ShowZeroStatus = tlb_sbt.Checked
                Setup_STATA()
                Setup_SC()

            Case "Show Cartons"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splPOTORDRX.Panel2Collapsed = Not tlb_sbt.Checked
                If tlb_sbt.Checked Then Setup_POTSHIP7()

            Case "Cardview"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grdICTSTATA.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked
                If tlb_sbt.Checked Then
                    SplitContainer1.SplitterDistance = 350
                Else
                    SplitContainer1.SplitterDistance = 160
                End If

            Case "Maintain Codes/Dates"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    If Not ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE, False, , , 1) Then
                        tlb_sbt.Checked = False
                        Exit Sub
                    End If
                Else
                    ASCMAIN1.MultiTask_Release(, , 1)
                End If

                With grdSOTALLO1.DisplayLayout.Bands(0)
                    If tlb_sbt.Checked Then
                        .Columns("ORDR_PRIORITY").Style = UltraWinGrid.ColumnStyle.DropDownList
                        .Columns("ORDR_RELEASE").Style = UltraWinGrid.ColumnStyle.DropDownList
                    Else
                        .Columns("ORDR_PRIORITY").Style = UltraWinGrid.ColumnStyle.Edit
                        .Columns("ORDR_RELEASE").Style = UltraWinGrid.ColumnStyle.Edit
                    End If
                End With
                Toggle_Maint()

            Case "Show Allocation Cur/Fut/Cxl"
                Toggle_ALLOCF()

            Case "Uncheck All"
                For Each row As DataRow In dst.Tables("ICTWHSES").Rows()
                    row("SEL") = "0"
                Next

            Case "Cost Maintenance"
                If grdICTSTATA.ActiveRow Is Nothing Then Exit Sub

                Using F As New ICFCOSTM
                    F.STYLE_CODE = STYLE_CODE
                    F.STYLE_DESC = Absx1.txtFor("STYLE_DESC").Text
                    F.COLOR_CODE = COLOR_CODE
                    F.COLOR_DESC = grdICTSTATA.ActiveRow.Cells("COLOR_DESC").Value
                    ' F.select_only = True
                    F.ShowDialog()
                    If F.updated Then
                        Dim YP_MIN As String = dst.Tables("ICTCOSTA").Compute("MIN(OPS_YYYYPP)", "") & ""
                        Dim YP_MIN_CLOSED As String = dst.Tables("ICTCOSTA").Compute("MIN(OPS_YYYYPP)", "YP_OPEN = '1'") & ""
                        If YP_MIN_CLOSED = "" Then YP_MIN_CLOSED = ASCMAIN1.CYP
                        ASCDATA1.DeleteRows(dst.Tables("ICTCOSTL"), "OPS_YYYYPP_FIFO >= '" & YP_MIN_CLOSED & "'")
                        ASCDATA1.DeleteRows(dst.Tables("ICTCOSTA"), "OPS_YYYYPP >= '" & YP_MIN_CLOSED & "'")
                        Calculate_Costs()
                        Fetch_Costs()
                    End If
                End Using

                Fill_Records("ICTCOSTL", New String() {STYLE_CODE, COLOR_CODE})
                Fetch_Costs()
            Case "Calculate With Royalty"
                MsgBox("Feature Still in Progress")
        End Select

        If grd Is Nothing OrElse grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Style Masterfile"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim keys As New Dictionary(Of String, Object)
                keys.Add("STYLE_CODE", STYLE_CODE)
                Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "Pre-Allocate"
                PreAllocate()

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = ""
                If grd.Name = "grdSOTALLO1" Then
                    PO_ORDER_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                Else
                    PO_ORDER_NO = grd.ActiveRow.Cells("PO_ORDER_NO").Value & ""
                End If

                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "Shipments Inquiry"
                Dim PO_SHIPMENT_NO As String = Trim(grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value & "")
                If PO_SHIPMENT_NO = "" Then
                    MsgBox("This PO has not been Shipped Yet", MsgBoxStyle.OkOnly, "Cannot Link to PO Shipment")
                    Exit Sub
                End If
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI")

            Case "Customer Order Inquiry"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Context_Launch("Select", CUST_CODE & ":" & ORDR_GROUP_NO, e.Tool.Key, "SOFCORD1")

            Case "Sales Order Inquiry", "Sales Order Entry"
                Dim ORDR_NO As String = ""
                Dim ORDR_GROUP_NO As String = ""
                If grd.Name = "grdSOTALLO1" Then
                    ORDR_GROUP_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                    ASCMAIN1.sql = "Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                    ORDR_NO = ASCDATA1.GetDataValue
                ElseIf grd.Name = "grdSOTINVHY" Then
                    Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value & ""
                    ASCMAIN1.sql = "Select Min (ORDR_NO) from SOTINVH1 where INV_TYPE = 'I' and INV_NO = '" & INV_NO & "'"
                    ORDR_NO = ASCDATA1.GetDataValue
                Else
                    ORDR_GROUP_NO = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                    ASCMAIN1.sql = "Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                    ORDR_NO = ASCDATA1.GetDataValue
                End If

                If e.Tool.Key = "Sales Order Entry" Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDR1")
                Else
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Sales Reservations Inquiry"
                Dim RSRV_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Context_Launch("View", RSRV_NO, e.Tool.Key, "SOFRSRV1")

            Case "Show Transaction"
                Dim TRAN_TYPE As String = grd.ActiveRow.Cells("TRAN_TYPE").Value
                Dim TRAN_NO As String = grd.ActiveRow.Cells("TRAN_NO").Value
                If TRAN_TYPE = "R" Then
                    Dim rowICTIREC1 As DataRow = LookUp("ICTIREC1", TRAN_NO)
                    Dim PO_SHIPMENT_NO As String = rowICTIREC1.Item("PO_SHIPMENT_NO")
                    Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI")
                End If
                If TRAN_TYPE = "A" Then
                    Context_Launch("View", TRAN_NO, e.Tool.Key, "ICFIADJI")
                End If
                If TRAN_TYPE = "T" Then
                    Context_Launch("View", TRAN_NO, e.Tool.Key, "ICFIXFRI")
                End If
                If TRAN_TYPE = "S" Then
                    Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"I", "0000" & TRAN_NO})
                    Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                    'Dim rowICTIREC1 As DataRow = LookUp("ICTIREC1", TRAN_NO)
                    'Dim PO_SHIPMENT_NO As String = rowICTIREC1.Item("PO_SHIPMENT_NO")
                    'Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI")
                End If

            Case "Show Order"
                Dim TRAN_TYPE As String = grd.ActiveRow.Cells("TRAN_TYPE").Value
                Dim TRAN_NO As String = grd.ActiveRow.Cells("TRAN_NO").Value
                Dim TRAN_TYPE_ORIG As String = grd.ActiveRow.Cells("TRAN_TYPE_ORIG").Value & ""
                Dim TRAN_NO_ORIG As String = grd.ActiveRow.Cells("TRAN_NO_ORIG").Value & ""
                Dim TRAN_REF As String = grd.ActiveRow.Cells("TRAN_REF").Value & ""

                If TRAN_TYPE_ORIG = "P" Then
                    Context_Launch("View", TRAN_REF, e.Tool.Key, "POFORDRI")
                    'Context_Launch("View", TRAN_NO_ORIG, e.Tool.Key, "POFSHIPI")
                ElseIf TRAN_TYPE = "T" Then
                    If TRAN_TYPE_ORIG = "S" Then
                        Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"I", TRAN_NO_ORIG})
                        If rowSOTINVH1 IsNot Nothing Then
                            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
                            Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                        End If
                    End If
                ElseIf TRAN_TYPE = "S" Then
                    Dim ORDR_GROUP_NO As String = TRAN_NO.PadLeft(10, "0")
                    Dim row As DataRow = ASCDATA1.GetDataRow("Select Min(ORDR_NO) ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                    If row IsNot Nothing Then
                        Dim ORDR_NO As String = row.Item("ORDR_NO")
                        Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                    End If
                End If

            Case "Show Invoice"
                'InvNos = GetInvoiceList(grd, e.Tool.OwningMenu.Key, INV_NO)
                'If InvNos.Length = 0 Then
                '    Exit Sub
                'End If
                Dim InvNos As String = grd.ActiveRow.Cells("INV_NO").Value
                ' Dim FILENAME As String = Create_Invoice(InvNos)
                Dim FILENAME As String = TAC.SOCMAIN1.Create_Invoice(Me, InvNos)
                Show_Document(FILENAME)

            Case "Style Status Inquiry"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    If rowICTSTYL1 IsNot Nothing Then
                        Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                    End If
                End If

            Case "Add to Quote Sheet"
                If grdICTSTYL1_Recent.Selected.Rows.Count = 0 Then
                    If grdICTSTYL1_Recent.ActiveRow Is Nothing Then
                        Exit Sub
                    Else
                        grdICTSTYL1_Recent.ActiveRow.Selected = True
                    End If
                End If

                For Each grow As UltraWinGrid.UltraGridRow In grdICTSTYL1_Recent.Selected.Rows
                    If grow.IsDataRow Then
                        Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                        Add_to_Quote(STYLE_CODE)
                    End If
                Next

            Case "Cancel Open Quantity"
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    Exit Sub
                End If

                Dim ORDR_GROUP_NO As String = String.Empty
                Dim tbl As DataTable = Nothing
                Dim sql As String = String.Empty
                Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
                Dim COLOR_CODE As String = grdICTSTATA.ActiveRow.Cells("COLOR_CODE").Value & String.Empty
                Dim detailLine As TAC.SOCORDR1.LineDetail

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTORDRX.Selected.Rows
                    ORDR_GROUP_NO = grdRow.Cells("ORDR_GROUP_NO").Value & String.Empty
                    ASCMAIN1.Progress("Processing Order Group: " & ORDR_GROUP_NO, String.Empty)
                    salesOrderLineDetails.Clear()

                    sql = "Select * from SOTORDR2 where STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
                    sql &= " AND ORDR_QTY_OPEN > 0"
                    sql &= " AND ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')"
                    tbl = ASCDATA1.GetDataTable(sql)

                    If tbl.Rows.Count = 0 Then
                        Continue For
                    End If

                    For Each row As DataRow In tbl.Select("", "ORDR_NO,ORDR_LNO")
                        If Val(row.Item("ORDR_QTY_OPEN") & String.Empty) <= 0 Then
                            Continue For
                        End If

                        detailLine = New TAC.SOCORDR1.LineDetail
                        With detailLine
                            .OrderNo = row.Item("ORDR_NO")
                            .OrderLineNo = row.Item("ORDR_LNO")
                            .StyleCode = row.Item("STYLE_CODE")
                            .ColorCode = row.Item("COLOR_CODE")
                            .CancelQuantity = row.Item("ORDR_QTY_OPEN")
                        End With

                        salesOrderLineDetails.Add(detailLine)
                    Next

                    If salesOrderLineDetails.Count > 0 Then
                        clsSOTORDR1.CancelItemsFormSalesOrder(ORDR_GROUP_NO, salesOrderLineDetails)
                    End If
                Next
                ASCMAIN1.Progress(String.Empty, String.Empty)
                Click_Command("Refresh")


            Case "Wave Inquiry"
                Dim WAVE_NO As String = grd.ActiveRow.Cells("WAVE_NO").Value
                Context_Launch("View", WAVE_NO, e.Tool.Key, "WHFWAVEI")


            Case "Location Inquiry for Style", "Location Inquiry for Location"
                Dim KEY As String = ""
                If e.Tool.Key = "Location Inquiry for Style" Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                    KEY = "S:" & STYLE_CODE
                Else
                    Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Value
                    KEY = "L:" & LOCATION_CODE
                End If

                Context_Launch("Select", KEY, e.Tool.Key, "WHFLOCS1")

            Case "Edit Ship+"

                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Text
                Dim ORDR_TYPE As String = grd.ActiveRow.Cells("ORDR_TYPE").Text
                Dim ORDR_SHIP_DATE As Date = grd.ActiveRow.Cells("ORDR_SHIP_DATE").Text
                Dim ORDR_CANCEL_DATE As Date = grd.ActiveRow.Cells("ORDR_CANCEL_DATE").Text

                If Not ASCMAIN1.Logical_Lock("ICTATOP1", $"{ORDR_NO}",,,, 1) Then Exit Sub
                If Not ASCMAIN1.Logical_Open("ICTATOP1", "*",,,, 1) Then Exit Sub

                Using F As New ICFATOP1
                    F.STYLE_CODE = STYLE_CODE
                    F.COLOR_CODE = COLOR_CODE
                    F.ORDR_NO = ORDR_NO
                    F.ORDR_TYPE = ORDR_TYPE
                    F.ORDR_SHIP_DATE = ORDR_SHIP_DATE
                    F.ORDR_CANCEL_DATE = ORDR_CANCEL_DATE

                    Select Case F.ShowDialog
                        Case Windows.Forms.DialogResult.OK
                            ' blnAtOnceChanged = True

                        Case Windows.Forms.DialogResult.Cancel

                    End Select

                    Dim sqlw As String = $"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' and ORDR_GROUP_NO = '{ORDR_NO}'"
                    sqlw = $"ORDR_TYPE = '{ORDR_TYPE}' and ORDR_GROUP_NO = '{ORDR_NO}'"
                    Dim rowICTATOP1 As DataRow = LookUp("ICTATOP1", New String() {STYLE_CODE, COLOR_CODE, ORDR_TYPE, ORDR_NO}, True)

                    Dim rowSOTORDRXs() As DataRow = dst.Tables("SOTORDRX").Select(sqlw)
                    For Each row As DataRow In rowSOTORDRXs
                        row.Item("STYLE_SHIP_WINDOW_DAYS") = rowICTATOP1.Item("STYLE_SHIP_WINDOW_DAYS")
                        row.Item("ORDR_SHIP_DATE_PLUS") = rowICTATOP1.Item("ORDR_SHIP_DATE_PLUS")
                        row.Item("STYLE_AT_ONCE_UNTIL") = rowICTATOP1.Item("STYLE_AT_ONCE_UNTIL")
                        row.Item("STYLE_AT_ONCE_ACTIVE") = rowICTATOP1.Item("STYLE_AT_ONCE_ACTIVE")

                        row.Item("SHIP_DATE_PLUS") = rowICTATOP1.Item("ORDR_SHIP_DATE_PLUS")
                    Next
                End Using

                ASCMAIN1.MultiTask_Release(,, 1)
                CalculateAtOnce()

            Case "Edit At-Once"

                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text

                Dim PS_CODE As String = "P"
                Dim PS_NO As String = PO_ORDER_NO
                Dim PS_ETA As Date = grd.ActiveRow.Cells("PO_SHIP_ETA").Value
                If PO_SHIPMENT_NO <> "" Then
                    PS_CODE = "S"
                    PS_NO = PO_SHIPMENT_NO
                End If

                If Not ASCMAIN1.Logical_Lock("ICTATOP2", $"{PS_CODE}-{PS_NO}",,,, 1) Then Exit Sub
                If Not ASCMAIN1.Logical_Open("ICTATOP2", "*",,,, 1) Then Exit Sub

                Using F As New ICFATOP2
                    F.STYLE_CODE = STYLE_CODE
                    F.COLOR_CODE = COLOR_CODE
                    F.PS_CODE = PS_CODE
                    F.PS_NO = PS_NO
                    F.PS_ETA = PS_ETA

                    Select Case F.ShowDialog
                        Case Windows.Forms.DialogResult.OK
                            blnAtOnceChanged = True

                        Case Windows.Forms.DialogResult.Cancel

                    End Select

                    Dim sqlw As String = $"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'"
                    If PS_CODE = "P" Then
                        sqlw &= $" and PO_ORDER_NO = '{PO_ORDER_NO}' and ISNULL(PO_SHIPMENT_NO,'?') = '?'"
                        sqlw = $"PO_ORDER_NO = '{PO_ORDER_NO}' and ISNULL(PO_SHIPMENT_NO,'?') = '?'"
                    Else
                        sqlw &= $" and PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'"
                        sqlw = $"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'"
                    End If
                    Dim rowICTATOP2 As DataRow = LookUp("ICTATOP2", New String() {STYLE_CODE, COLOR_CODE, PS_CODE, PS_NO}, True)

                    Dim rowPOTORDRXs() As DataRow = dst.Tables("POTORDRX").Select(sqlw)
                    For Each row As DataRow In rowPOTORDRXs
                        row.Item("STYLE_ARRIVAL_BUFFER_DAYS") = rowICTATOP2.Item("STYLE_ARRIVAL_BUFFER_DAYS")
                        row.Item("STYLE_AT_ONCE_UNTIL") = rowICTATOP2.Item("STYLE_AT_ONCE_UNTIL")
                        row.Item("STYLE_AT_ONCE_ACTIVE") = rowICTATOP2.Item("STYLE_AT_ONCE_ACTIVE")
                    Next

                End Using

                ASCMAIN1.MultiTask_Release(,, 1)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                If ctl.Text <> "" Then
                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                        Dim PARTAIALSTYLE As String = PARTIALSTYLE(ctl.Text)
                        If PARTAIALSTYLE.Length > 0 Then
                            ctl.Text = PARTAIALSTYLE
                        End If
                    End If
                    'Call Click_Command("Load Reports")
                End If
            Case "SREP_CODE"
                'Setup_SO()
        End Select
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("Select")
                End If
            Case "STYLE_CODE_CUST"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("STYLE_CODE_CUST").Text <> "" Then
                        Dim STYLE_CODE_CUST As String = Absx1.txtFor("STYLE_CODE_CUST").Text
                        If Add_to_Quote(STYLE_CODE_CUST) = "" Then
                            MsgBox("Style Code " & STYLE_CODE_CUST & " Not on Fle", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                        End If
                    End If
                    Absx1.txtFor("STYLE_CODE_CUST").Text = ""
                    Absx1.txtFor("STYLE_CODE_CUST").Focus()
                End If
            Case "SREP_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Setup_SO()
                End If

            Case "UPC_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Dim UPC_CODE As String = Absx1.txtFor("UPC_CODE").Text
                    If Len(UPC_CODE) = 6 Then
                        UPC_CODE = ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & UPC_CODE
                    End If
                    ASCMAIN1.sql = "Select STYLE_CODE from ICTSTYC1 where UPC_CODE = :PARM1"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", UPC_CODE)
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        If row Is Nothing Then
                            ASCMAIN1.sql = "Select STYLE_CODE from ICTSTYC4 where UPC_CODE = :PARM1"
                            row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", UPC_CODE)
                        End If
                        If row Is Nothing Then
                            ASCMAIN1.sql = "Select STYLE_CODE from ICTSTYC2 where UPC_CODE = :PARM1"
                            row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", UPC_CODE)
                        End If
                    End If
                    If row Is Nothing Then
                        MsgBox("Could not Find any Styles with UPC Code " & UPC_CODE, MsgBoxStyle.OkOnly, "Cannot Process Requested Action")
                    Else
                        Absx1.txtFor("STYLE_CODE").Text = row.Item("STYLE_CODE")
                        Click_Command("Select")
                    End If
                End If

            Case "SUPPLIER"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If txtSupplierOption.Text <> "" Then
                        txtSupplierOption.Tag = "V"
                        Setup_Recent()
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"
                If txtctl.Text <> "" Then
                    Click_Command("Select")
                End If
            Case "STYLE_CODE_CUST"
                If txtctl.Text <> "" Then
                    Add_to_Quote(txtctl.Text)
                    txtctl.Text = ""
                    txtctl.Focus()
                End If
            Case "SREP_CODE"
                Setup_SO()

            Case "SUPPLIER"
                If txtSupplierOption.Text <> "" Then
                    txtSupplierOption.Tag = "V"
                    Setup_Recent()
                End If

        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)

        Select Case COLUMN_NAME
            Case "SUPPLIER"
                If txtSupplierOption.Text = "" Then
                    txtSupplierOption.Tag = ""
                    Setup_Recent()
                End If
        End Select

    End Sub
#End Region

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        ' If IsDone Then Exit Sub

        If tabMain.SelectedTab.Key = "Allocate" And AutoAllocate = False Then
            Allocate()
            AutoAllocate = True
        End If

        If tabMain.SelectedTab.Key = "Orders" Then
            If blnAtOnceChanged Then
                blnAtOnceChanged = False
                If chkAutoCalculate.Checked Then
                    CalculateAtOnce()
                End If
            End If
        End If

        With UltraExplorerBar1
            .Groups("Calculate At-Once").Visible = ScreenMode AndAlso (tabMain.SelectedTab.Key = "Orders") AndAlso ASCMAIN1.CLIENT = "RGI"
            .Groups("Orders").Visible = ScreenMode AndAlso (tabMain.SelectedTab.Key = "Orders")
            '.Groups("Orders").Visible = False ' wjz removing this group while screemode is true
            .Groups("Available by Date").Visible = ScreenMode And (STYLE_CODE_allocated <> "") ' ScreenMode AndAlso (tabMain.SelectedTab.Key = "Allocate")

            If tabMain.SelectedTab.Key = "Allocate" Then
                grpATS.Visible = False
                DirectCast(grdICTSTDQ1.DataSource, DataTable).DefaultView.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
                splATS.FixedPanel = FixedPanel.Panel2
                .Groups("Available by Date").Settings.ContainerHeight = 360
                splATS.SplitterDistance = splATS.Height - 85 ' 275
                grdICTSTDQ1.DisplayLayout.Bands(0).Columns("WHSE_CODE").Hidden = False
                Setup_ICTSTDQ1()
                'optASL.Visible = True
                'If ASCMAIN1.Running_in_VS Then
                'Else

                '    '   optASL.Visible = False ' not working yet - recall MTX43266 on 06/09/13 - showing more than avail at last date
                'End If
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    optASL.Visible = True
                Else
                    optASL.Visible = False
                End If
            Else
                grpATS.Visible = True
                Set_WHSE_CODE()
                splATS.FixedPanel = FixedPanel.Panel2
                .Groups("Available by Date").Settings.ContainerHeight = 200 ' 150
                splATS.SplitterDistance = splATS.Height - 40 ' 110
                ' grdICTSTDQ1.Text = ""
                grdICTSTDQ1.DisplayLayout.Bands(0).Columns("WHSE_CODE").Hidden = True
                Sort_grdColumns(grdICTSTDQ1, "STATUS_DATE")
                optASL.Visible = False
            End If
            .Groups("PO / In Transit").Visible = ScreenMode AndAlso (tabMain.SelectedTab.Key = "PO / In Transit")
            '.Groups("Costs").Visible = ScreenMode AndAlso (tabMain.SelectedTab.Key = "Costs")
            .Groups("Sales History").Visible = ScreenMode AndAlso (tabMain.SelectedTab.Key = "Sales History")

            .Groups("Style Image").Visible = False And ScreenMode And Not (UltraExplorerBar1.Groups("Style Image").Text = "Style Image") _
                And Not (tabMain.SelectedTab.Key = "Sales History") _
                And Not (tabMain.SelectedTab.Key = "Costs") _
                And Not (tabMain.SelectedTab.Key = "Allocate")

        End With

        If tabMain.SelectedTab.Key = "Costs" Then
            Calculate_Costs()
            Fetch_Costs()
        End If

    End Sub

    Sub Calculate_Costs()
        Dim rowICTCOSTA_CYP As DataRow = dst.Tables("ICTCOSTA").Rows.Find(New String() {ASCMAIN1.CYP, STYLE_CODE, COLOR_CODE})
        If rowICTCOSTA_CYP Is Nothing Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Calculating FIFO Costs")

            Dim YP As String = dst.Tables("ICTCOSTA").Compute _
                               ("MAX(OPS_YYYYPP)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & ""
            If YP = "" Then YP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)

            Dim tblICTCOSTA As DataTable = dst.Tables("ICTCOSTA").Copy
            Dim tblICTCOSTL As DataTable = dst.Tables("ICTCOSTL").Copy
            Do
                If YP >= ASCMAIN1.CYP Then Exit Do
                YP = ASCMAIN1.Period_Calc(YP, 1)
                TAC.ICCMAIN1.Calculate_FIFO(Me, YP, True, ICTCOSTA, ICTCOSTL, ICTCOSTU, ICTCOSTG, STYLE_CODE)
                '  Fill_Records("ICTCOSTA", "", False, "Select * from " & ICTCOSTA)
                For Each rowICTCOSTA As DataRow In dst.Tables("ICTCOSTA").Select("")
                    rowICTCOSTA.Item("YP_OPEN") = "1"
                Next
                tblICTCOSTA.Merge(dst.Tables("ICTCOSTA"))

                'ASCMAIN1.sql = Replace(Replace(Replace(sqlICTCOSTL, ":PARM1", "'" & YP & "'"), ":PARM2", "'" & STYLE_CODE & "'"), ":PARM3", "'" & COLOR_CODE & "'")
                'Fill_Records("ICTCOSTA", "", True, ASCMAIN1.sql)
                '
                For Each rowICTCOSTL As DataRow In dst.Tables("ICTCOSTL").Select("TRAN_TYPE = 'R'")
                    ASCMAIN1.sql = "Select ICTIREC1.VEND_CODE,ICTIREC2.PO_ORDER_NO,POTSHIP1.COST_COMPLETE,ICTIREC2.PO_COST,ICTIREC2.INV_COST" & vbCrLf _
                        & ", POTSHIP3.PO_COST_FREIGHT_IN, POTSHIP3.PO_COST_DUTY, POTSHIP3.PO_COST_CUSTOMS, POTSHIP3.PO_COST_TRUCKING, POTSHIP3.DUTY_RATE_CODE, POTSHIP3.DUTY_RATE" & vbCrLf _
                        & ", ICTIREC1.INIT_DATE, ICTIREC1.INIT_OPER, ICTIREC1.LAST_DATE, ICTIREC1.LAST_OPER" & vbCrLf _
                        & " from ICTIREC1,ICTIREC2,POTSHIP1,POTSHIP3" & vbCrLf _
                        & " where POTSHIP3.PO_SHIPMENT_NO (+) = ICTIREC2.PO_SHIPMENT_NO" & vbCrLf _
                        & "   and POTSHIP3.PO_SHIPMENT_LNO (+) = ICTIREC2.PO_SHIPMENT_LNO" & vbCrLf _
                        & "   and POTSHIP3.PO_ORDER_NO (+) = ICTIREC2.PO_ORDER_NO" & vbCrLf _
                        & "   and POTSHIP3.PO_ORDER_LNO (+) = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
                        & "   and ICTIREC2.RECEIPT_NO (+) = '" & rowICTCOSTL.Item("TRAN_NO") & "'" & vbCrLf _
                        & "   and ICTIREC2.RECEIPT_LNO (+) = " & CStr(Val(rowICTCOSTL.Item("TRAN_LNO") & "")) & vbCrLf _
                        & "   and ICTIREC1.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
                        & "   and POTSHIP1.PO_SHIPMENT_NO (+) = ICTIREC2.PO_SHIPMENT_NO" & vbCrLf
                    Dim row As DataRow = ASCDATA1.GetDataRow
                    If row IsNot Nothing Then
                        rowICTCOSTL.Item("VEND_CODE") = row.Item("VEND_CODE")
                        rowICTCOSTL.Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                        rowICTCOSTL.Item("COST_COMPLETE") = row.Item("COST_COMPLETE")
                        rowICTCOSTL.Item("PO_COST") = row.Item("PO_COST")
                        rowICTCOSTL.Item("INV_COST") = row.Item("INV_COST")
                        rowICTCOSTL.Item("PO_COST_FREIGHT_IN") = row.Item("PO_COST_FREIGHT_IN")
                        rowICTCOSTL.Item("PO_COST_DUTY") = row.Item("PO_COST_DUTY")
                        rowICTCOSTL.Item("PO_COST_CUSTOMS") = row.Item("PO_COST_CUSTOMS")
                        rowICTCOSTL.Item("PO_COST_TRUCKING") = row.Item("PO_COST_TRUCKING")
                        rowICTCOSTL.Item("DUTY_RATE_CODE") = row.Item("DUTY_RATE_CODE")
                        rowICTCOSTL.Item("DUTY_RATE") = row.Item("DUTY_RATE")
                        rowICTCOSTL.Item("INIT_OPER") = row.Item("INIT_OPER")
                        rowICTCOSTL.Item("INIT_DATE") = row.Item("INIT_DATE")
                        rowICTCOSTL.Item("LAST_OPER") = row.Item("LAST_OPER")
                        rowICTCOSTL.Item("LAST_DATE") = row.Item("LAST_DATE")
                    End If
                Next

                For Each rowICTCOSTL As DataRow In dst.Tables("ICTCOSTL").Select("")
                    If rowICTCOSTL.Item("TRAN_TYPE") & "" = "R" Then
                        ASCMAIN1.sql = "Select ICTIREC1.VEND_CODE,ICTIREC2.PO_ORDER_NO,POTSHIP1.COST_COMPLETE,ICTIREC2.PO_COST,ICTIREC2.INV_COST" & vbCrLf _
                            & ", POTSHIP3.PO_COST_FREIGHT_IN, POTSHIP3.PO_COST_DUTY, POTSHIP3.PO_COST_CUSTOMS, POTSHIP3.PO_COST_TRUCKING, POTSHIP3.DUTY_RATE_CODE, POTSHIP3.DUTY_RATE" & vbCrLf _
                            & ", ICTIREC1.INIT_DATE, ICTIREC1.INIT_OPER, ICTIREC1.LAST_DATE, ICTIREC1.LAST_OPER" & vbCrLf _
                            & " from ICTIREC1,ICTIREC2,POTSHIP1,POTSHIP3" & vbCrLf _
                            & " where POTSHIP3.PO_SHIPMENT_NO (+) = ICTIREC2.PO_SHIPMENT_NO" & vbCrLf _
                            & "   and POTSHIP3.PO_SHIPMENT_LNO (+) = ICTIREC2.PO_SHIPMENT_LNO" & vbCrLf _
                            & "   and POTSHIP3.PO_ORDER_NO (+) = ICTIREC2.PO_ORDER_NO" & vbCrLf _
                            & "   and POTSHIP3.PO_ORDER_LNO (+) = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
                            & "   and ICTIREC2.RECEIPT_NO (+) = '" & rowICTCOSTL.Item("TRAN_NO") & "'" & vbCrLf _
                            & "   and ICTIREC2.RECEIPT_LNO (+) = " & CStr(Val(rowICTCOSTL.Item("TRAN_LNO") & "")) & vbCrLf _
                            & "   and ICTIREC1.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
                            & "   and POTSHIP1.PO_SHIPMENT_NO (+) = ICTIREC2.PO_SHIPMENT_NO" & vbCrLf

                        Dim row As DataRow = ASCDATA1.GetDataRow

                        If row IsNot Nothing Then
                            rowICTCOSTL.Item("VEND_CODE") = row.Item("VEND_CODE")
                            rowICTCOSTL.Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                            rowICTCOSTL.Item("COST_COMPLETE") = row.Item("COST_COMPLETE")
                            rowICTCOSTL.Item("PO_COST") = row.Item("PO_COST")
                            rowICTCOSTL.Item("INV_COST") = row.Item("INV_COST")
                            rowICTCOSTL.Item("PO_COST_FREIGHT_IN") = row.Item("PO_COST_FREIGHT_IN")
                            rowICTCOSTL.Item("PO_COST_DUTY") = row.Item("PO_COST_DUTY")
                            rowICTCOSTL.Item("PO_COST_CUSTOMS") = row.Item("PO_COST_CUSTOMS")
                            rowICTCOSTL.Item("PO_COST_TRUCKING") = row.Item("PO_COST_TRUCKING")
                            rowICTCOSTL.Item("DUTY_RATE_CODE") = row.Item("DUTY_RATE_CODE")
                            rowICTCOSTL.Item("DUTY_RATE") = row.Item("DUTY_RATE")
                            rowICTCOSTL.Item("INIT_OPER") = row.Item("INIT_OPER")
                            rowICTCOSTL.Item("INIT_DATE") = row.Item("INIT_DATE")
                            rowICTCOSTL.Item("LAST_OPER") = row.Item("LAST_OPER")
                            rowICTCOSTL.Item("LAST_DATE") = row.Item("LAST_DATE")
                        End If
                    End If
                Next
                tblICTCOSTL.Merge(dst.Tables("ICTCOSTL"))
            Loop
            dst.Tables("ICTCOSTA").Rows.Clear()
            dst.Tables("ICTCOSTA").Merge(tblICTCOSTA)
            dst.Tables("ICTCOSTL").Rows.Clear()
            dst.Tables("ICTCOSTL").Merge(tblICTCOSTL)

            Sort_grdColumns(grdICTCOSTA, "OPS_YYYYPP".ToLower)

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        Else
            Dim YP_LAST_CLOSED As String = ASCDATA1.GetDataValue("SELECT MAX(OPS_YYYYPP) FROM ICTCOSTP WHERE UPDATED = '1'")
            For Each rowICTCOSTA As DataRow In dst.Tables("ICTCOSTA").Select("OPS_YYYYPP > '" & YP_LAST_CLOSED & "'")
                rowICTCOSTA.Item("YP_OPEN") = "1"
            Next
        End If

        Sort_grdColumns(grdICTCOSTA, "OPS_YYYYPP".ToLower)
    End Sub

    Sub Fetch_Costs()
        If grdICTCOSTA.ActiveRow Is Nothing Then
            grdICTCOSTL.Visible = False
        Else
            grdICTCOSTL.Visible = True
            Dim OPS_YYYYPP As String = grdICTCOSTA.ActiveRow.Cells("OPS_YYYYPP").Value
            'Fill_Records("ICTCOSTL", New String() {OPS_YYYYPP, STYLE_CODE, COLOR_CODE})

            Dim dvw As DataView = DirectCast(grdICTCOSTL.DataSource, DataTable).DefaultView
            dvw.RowFilter = "OPS_YYYYPP_FIFO = '" & OPS_YYYYPP & "'"

            'If dst.Tables("ICTCOSTL").Rows.Count = 0 And _
            '    (Val(grdICTCOSTA.ActiveRow.Cells("WHSE_QTY_ON_HAND").Value & "") <> 0 Or _
            '     Val(grdICTCOSTA.ActiveRow.Cells("LOT_QTY_ONHD").Value & "") <> 0 Or _
            '     Val(grdICTCOSTA.ActiveRow.Cells("LOT_QTY_USED").Value & "") <> 0) Then
            '    If tabMain.SelectedTab.Key = "Costs" Then
            '        TAC.ICCMAIN1.Calculate_FIFO(Me, OPS_YYYYPP, True, ICTCOSTA, ICTCOSTL, ICTCOSTU, STYLE_CODE)
            '        Fill_Records("ICTCOSTL", New String() {OPS_YYYYPP, STYLE_CODE, COLOR_CODE})
            '    End If
            'End If
            Sort_grdColumns(grdICTCOSTL, "RECORD_NO".ToLower)
            grdICTCOSTL.Text = "Cost Lots for " & OPS_YYYYPP & " for " & STYLE_CODE & ", Color " & COLOR_CODE
        End If
    End Sub

    Sub Set_WHSE_CODE()
        Dim WHSE_CODE As String = cbeWHSE_CODE.Value & ""
        DirectCast(grdICTSTDQ1.DataSource, DataTable).DefaultView.RowFilter = "WHSE_CODE = '" & WHSE_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
    End Sub

    Private Sub grdICTSTYL1_Recent_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTYL1_Recent.AfterRowActivate

    End Sub

    Private Sub grdICTSTYL1_Recent_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTYL1_Recent.DoubleClickRow
        If e.Row.IsDataRow Then
            optSelectBy.Value = "S"
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
            Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
            Click_Command("Select")
        End If
    End Sub

    Private Sub optSelectBy_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSelectBy.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_SelectBy()
    End Sub

    Sub Setup_SelectBy()

        lblSTYLE_CODE.Visible = (optSelectBy.Value = "S" Or optSelectBy.Value = "R")
        lblSTYLE_DESC.Visible = (optSelectBy.Value = "S" Or optSelectBy.Value = "R")

        Absx1.txtFor("STYLE_CODE").Visible = (optSelectBy.Value = "S")
        Absx1.txtFor("STYLE_DESC").Visible = (optSelectBy.Value = "S")

        Absx1.txtFor("UPC_CODE").Visible = (optSelectBy.Value = "S")
        lblUPC_CODE.Visible = (optSelectBy.Value = "S")

        Absx1.txtFor("RANGE_STYLE_CODE").Visible = (optSelectBy.Value = "R")
        Absx1.txtFor("RANGE_STYLE_DESC").Visible = (optSelectBy.Value = "R")

        If optSelectBy.Value = "S" Then
            lblSTYLE_CODE.Text = "Style Code"
        ElseIf optSelectBy.Value = "R" Then
            lblSTYLE_CODE.Text = "Range Style Code"
        End If

        If optSelectBy.Value = "M" Then
            STYLEs.Clear()
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                Using F As New ASFCODE1
                    F.ShowDialog()
                End Using
                Dim CODE_VALUES As String = ""
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    For Each SelectedCode As String In ASCMAIN1.CodeSelector.SelectedCodes
                        STYLEs.Add(SelectedCode)
                    Next
                End If
            End If
            If STYLEs.Count = 0 Then
                optSelectBy.Value = "S"
            Else
                Click_Command("Select")
            End If
        End If
    End Sub

#Region "VB6"

    Sub Clear_Allo()
        dst.Tables("SOTALLO1").Rows.Clear()

        grdSOTALLO1.Visible = False
        optASL.Visible = False

        Toggle_Maint()
    End Sub

    Sub Print_Record()
        Synch_TABLE_NAME("ICTSTYL1")
        Print_Report_Begin()
        Generate_Report("ICRSTYLA", Me.Text, "Style " & STYLE_CODE & " Information Sheet")
        Print_Report_End()
    End Sub

    Function FetchImage(rowICTSTYL1 As DataRow) As Byte()
        Dim STYLE_CODE As String
        If rowICTSTYL1.Table.TableName = "ICTQUOT2" Then
            STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE_PLM") & ""
        Else
            STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & ""
        End If
        Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""


        If IMAGE_NAME = "" Then IMAGE_NAME = STYLE_CODE

        'If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
        '    IMAGE_NAME = ""
        '    IMAGE_NAME = rowICTSTYL1.Item("SEASON_CODE") & "\MENS\" & STYLE_CODE
        'End If
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            IMAGE_NAME = STYLE_CODE & "-" & COLOR_CODE
        End If

        Dim imgba() As Byte = Nothing
        If IMAGE_NAME <> "" Then
            'Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")
            ''If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then FOLDER_NAME = "\\10.0.1.2\Data\Database\Images\"

            ' imgSTYLE.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
            imgSTYLE.Image = Get_Style_Image(IMAGE_NAME, imgba)
            picStyleImage.Image = imgSTYLE.Image
            rowICTSTYL1.Item("IMAGE") = imgba

            UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE & "-" & COLOR_CODE

        Else
            imgSTYLE.Image = Nothing
            rowICTSTYL1.Item("IMAGE") = DBNull.Value
            picStyleImage.Image = Nothing

            UltraExplorerBar1.Groups("Style Image").Text = "Style Image"

        End If

        UltraExplorerBar1.Groups("Style Image").Visible = False And Not (UltraExplorerBar1.Groups("Style Image").Text = "Style Image")
        Return imgba

    End Function

    Function Get_Style_Image(
        ByVal IMAGE_NAME As String,
        Optional ByRef imgba() As Byte = Nothing) As System.Drawing.Bitmap

        ' Dim imgba() As Byte = Nothing

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")
        'If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then FOLDER_NAME = "\\10.0.1.2\Data\Database\Images\"
        ' If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then FOLDER_NAME = "C:\RGI_Images\"
        Return ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)

    End Function

    Private Sub Allocate()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Allocating ... (Please Wait)")

        grdSOTALLO1.Visible = False

        Dim SOTORDR0 As String = TABLE_NAMEs("SOTORDR0")
        Dim SOTORDR1 As String = TABLE_NAMEs("SOTORDR1")
        Dim SOTORDR2 As String = TABLE_NAMEs("SOTORDR2")
        Dim SOTRSRV1 As String = TABLE_NAMEs("SOTRSRV1")
        Dim SOTRSRV2 As String = TABLE_NAMEs("SOTRSRV2")
        Dim ARTCUST1 As String = TABLE_NAMEs("ARTCUST1")

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
        dst.Tables("ICTSTDQ2").Rows.Clear()
        dst.Tables("ICTSTDQ3").Rows.Clear()

        Dim read_only As Boolean = True
        ' If ASCMAIN1.Running_in_VS Then read_only = False
        If ASCMAIN1.CLIENT = "RGI" Then
            read_only = False
            read_only = True
        End If
        TAC.SOCMAIN1.Allocation(Me, False, True, "", "", edi850cust,
                                  SOTSUPP1, SOTDEMD1,
                                  TABLE_NAMEs,
                                  read_only, (optASL.Value = "1"), STYLE_CODE)

        If ASCMAIN1.CLIENT = "RGIX" Then
            ' why RGIX?  this update may be screwing up other colors when the data was prepared for a single color
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is Select * from " & SOTORDR2 & $" where STYLE_CODE = '{STYLE_CODE}';" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update SOTORDR2 Set" & vbCrLf _
                & "      ORDR_QTY_ALLO = R1.ORDR_QTY_ALLO" & vbCrLf _
                & "    , ORDR_RELEASE = R1.ORDR_RELEASE" & vbCrLf _
                & "    , ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" & vbCrLf _
                & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"

            ASCDATA1.ExecuteSQL()
        End If

        'If ASCMAIN1.CLIENT = "RGI" Then
        '    Dim QTY_PLUS_CUM As Int64 = 0
        '    For Each row As DataRow In dst.Tables("ICTSTDQ1").Select("", "STATUS_DATE")
        '        Dim QTY_PLUS As Int64 = Val(row.Item("QTY_PLUS") & "")
        '        QTY_PLUS_CUM += QTY_PLUS
        '        row.Item("QTY_PLUS_CUM") = QTY_PLUS_CUM
        '    Next

        'End If

        ASCMAIN1.sql = "Select SOTORDR7.* from SOTORDR7 where SOTORDR7.STYLE_CODE = '" & STYLE_CODE & "'" _
            & " and SOTORDR7.PICK_BATCH_NO is Null" & vbCrLf
        Fill_Records("SOTORDR7", "", True, ASCMAIN1.sql)

        Load_SOTALLO1()
        STYLE_CODE_allocated = STYLE_CODE

        If ASCMAIN1.CLIENT = "RGI" Then

            For Each rowWHSE As DataRow In ASCDATA1.SelectDistinct("SOTALLO1", "WHSE_CODE").Select("")
                Dim WHSE_CODE As String = rowWHSE.Item("WHSE_CODE")
                Dim BALANCE As Int64 = 0

                For Each row As DataRow In dst.Tables("SOTALLO1").Select($"WHSE_CODE = '{WHSE_CODE}'")

                    If row.Item("RECORD_TYPE") = "0" Then
                        If row.Item("RECORD_SUB_TYPE") = "H" Then
                            BALANCE = row.Item("BALANCE")
                        End If
                    Else
                        row.Item("SD_DATE") = row.Item("ORDR_RELEASE_AVAIL")
                    End If
                Next

                ' Dim rowICTSTDQ1_supply As DataRow = Nothing

                For Each row As DataRow In dst.Tables("SOTALLO1").Select($"WHSE_CODE = '{WHSE_CODE}'",
                    "WHSE_CODE,SD_DATE,RECORD_TYPE,ORDR_PRIORITY_DATE,ORDR_PRIORITY_DATE_ORIG,RECORD_SUB_TYPE")
                    Dim SD_QTY As Int64 = Val(row.Item("SD_QTY"))

                    Dim SD_DATE As Date = Now.Date
                    If row.Item("SD_DATE") & "" = "" Then
                    Else
                        SD_DATE = row.Item("SD_DATE")
                    End If

                    ' Dim rowICTSTDQ1 As DataRow = dst.Tables("ICTSTDQ1").Rows.Find(New Object() {WHSE_CODE, STYLE_CODE, COLOR_CODE, SD_DATE})

                    If row.Item("RECORD_TYPE") = "0" Then
                        'rowICTSTDQ1_supply = rowICTSTDQ1
                    End If

                    If row.Item("RECORD_SUB_TYPE") = "H" Then
                        ' rowICTSTDQ1_supply.Item("SUPPLY_QTY") = SD_QTY
                    Else
                        If row.Item("RECORD_TYPE") = "0" Then
                            BALANCE += SD_QTY
                            ' rowICTSTDQ1_supply.Item("SUPPLY_QTY") = SD_QTY
                        Else
                            BALANCE -= SD_QTY
                            ' If rowICTSTDQ1_supply IsNot Nothing Then rowICTSTDQ1_supply.Item("SUPPLY_QTY") = Val(rowICTSTDQ1_supply.Item("SUPPLY_QTY") & "") - SD_QTY
                        End If
                        row.Item("BALANCE") = BALANCE
                        ' rowICTSTDQ1.Item("QTY_PLUS_CUM") = BALANCE
                    End If
                Next
            Next
        End If

        For Each rowICTSTATA As DataRow In dst.Tables("ICTSTATA").Select("")
            rowICTSTATA.Item("ALLO") = 0
        Next
        For Each rowICTSTATW As DataRow In dst.Tables("ICTSTATW").Select("")
            rowICTSTATW.Item("ALLO") = 0
        Next
        ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and WHSE_QTY_ALLO <> 0"
        For Each rowICTSTAT2 As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowICTSTATA As DataRow = dst.Tables("ICTSTATA").Rows.Find _
                (New String() {rowICTSTAT2.Item("STYLE_CODE"), rowICTSTAT2.Item("COLOR_CODE")})
            If rowICTSTATA IsNot Nothing Then rowICTSTATA.Item("ALLO") += rowICTSTAT2.Item("WHSE_QTY_ALLO")
            Dim rowICTSTATW As DataRow = dst.Tables("ICTSTATW").Rows.Find _
                (New String() {rowICTSTAT2.Item("STYLE_CODE"), rowICTSTAT2.Item("COLOR_CODE"), rowICTSTAT2.Item("WHSE_CODE")})
            If rowICTSTATW IsNot Nothing Then rowICTSTATW.Item("ALLO") += rowICTSTAT2.Item("WHSE_QTY_ALLO")
        Next

        '  Price_and_Availability(STYLE_CODE)
        STYLE_CLASS_CODE = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
        If STYLE_CLASS_CODE = "" Then
            MsgBox("Warning: Style " & STYLE_CODE & " does not have a Class Code",
                   MsgBoxStyle.OkOnly, "Please Assign one Immediately")
        End If
        CARTON_PACK_QTY = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
        STYLE_PRICE = Val(rowICTSTYL1.Item("STYLE_PRICE") & "")
        Price_and_Availability(STYLE_CODE, STYLE_CLASS_CODE, COLOR_CODE, CARTON_PACK_QTY, STYLE_PRICE)

        ' this sort is necessary for the balance to make sense - see RGI MTX11271
        Sort_grdColumns(grdSOTALLO1, "WHSE_CODE,SD_DATE,RECORD_TYPE,RECORD_SUB_TYPE")
        If ASCMAIN1.CLIENT = "RGI" Then
            'Sort_grdColumns(grdSOTALLO1, "WHSE_CODE,SD_DATE,RECORD_TYPE,RECORD_SUB_TYPE,ORDR_PRIORITY_DATE,ORDR_PRIORITY_DATE_ORIG")
            Sort_grdColumns(grdSOTALLO1, "WHSE_CODE,SD_DATE,RECORD_TYPE,ORDR_PRIORITY_DATE,ORDR_PRIORITY_DATE_ORIG,RECORD_SUB_TYPE") ' SD_DATE IS SHIP DATE - NO IT IS NOT - IT IS NOW THE ORDR_RELEASE_AVAIL DATE
            'Sort_grdColumns(grdSOTALLO1, "WHSE_CODE,SD_DATE,RECORD_TYPE,ORDR_PRIORITY_DATE,RECORD_SUB_TYPE") ' SD_DATE IS SHIP DATE
        End If

        grdSOTALLO1.Visible = True
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub FetchShippedOrders()
        If cmbYP_From.Value & "" = "" Or cmbYP_To.Value & "" = "" Then
            MsgBox("You Must First Specify Starting and Ending Shipping Periods", MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        grdSOTINVHY.DisplayLayout.Bands(0).Columns("INV_NO").Header.Caption = "Invoice No"
        grdSOTINVHY.DisplayLayout.Bands(0).Columns("INV_DATE").Header.Caption = "Inv Date"

        Me.Cursor = Cursors.WaitCursor

        grdSOTINVHX.Text = "Customer Shipments, Style " & STYLE_CODE & IIf(chkSingleColor.Checked, ", Color " & COLOR_CODE, ", All Colors")
        grdSOTINVHX.Tag = "S"
        grdSOTINVHY.Text = "Customer Shipment Details" & IIf(chkSingleColor.Checked, ", Color " & COLOR_CODE, ", All Colors")

        Dim sp As String = Mid(cmbYP_From.Value, 1, 4) & Mid(cmbYP_From.Value, 6, 2)
        Dim ep As String = Mid(cmbYP_To.Value, 1, 4) & Mid(cmbYP_To.Value, 6, 2)

        Fill_Records("SOTINVHX", New String() {sp, ep, STYLE_CODE, IIf(chkSingleColor.Checked, COLOR_CODE, "*")})
        Sort_grdColumns(grdSOTINVHX, "CUST_CODE")
        grdSOTINVHX.Visible = True
        'If grdCUST.Rows.Count > 0 Then
        '    Load_CUSTD()
        'End If
        Load_SOTINVHY()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub FetchOpenOrders()
        Me.Cursor = Cursors.WaitCursor

        grdSOTINVHY.DisplayLayout.Bands(0).Columns("INV_NO").Header.Caption = "Status"
        grdSOTINVHY.DisplayLayout.Bands(0).Columns("INV_DATE").Header.Caption = "Ship Date"

        grdSOTINVHX.Text = "Customer Orders, Style " & STYLE_CODE & IIf(chkSingleColor.Checked, ", Color " & COLOR_CODE, ", All Colors")
        grdSOTINVHY.Text = "Customer Order Details"
        grdSOTINVHX.Tag = "O"

        ASCMAIN1.sql = " Select SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SUM (Decode (SOTORDR2.ORDR_STATUS, 'O', SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK)) QTY" & vbCrLf _
            & ", SUM (Decode (SOTORDR2.ORDR_STATUS, 'O', SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK) * SOTORDR2.ORDR_UNIT_PRICE) AMT" & vbCrLf _
            & " from SOTORDR2,SOTORDR1 " & vbCrLf _
            & " where (SOTORDR2.ORDR_STATUS = 'O' or SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf _
            & "   and SOTORDR1.ORDR_TYPE_CODE <> 'XFR'" & vbCrLf _
            & "   and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & IIf(chkSingleColor.Checked, "   and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'", "") & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " group by SOTORDR1.CUST_CODE"
        ASCMAIN1.sql = "Select X.*,ARTCUST1.CUST_NAME from (" & ASCMAIN1.sql & ") X,ARTCUST1" _
            & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
        Fill_Records("SOTINVHX", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdSOTINVHX, "CUST_CODE")
        grdSOTINVHX.Visible = True
        'If grdCUST.Rows.Count > 0 Then
        '    Load_CUSTD()
        'End If
        Load_SOTINVHY()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub PreAllocate()

        Stop

        If grdSOTALLO1.ActiveRow.Cells("RECORD_TYPE").Value & "" <> "1" _
        Or grdSOTALLO1.ActiveRow.Cells("RECORD_SUB_TYPE").Value & "" <> "O" Then
            MsgBox("Wrong Type of Record for Pre-Allocation by Store", MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        If grdSOTALLO1.Selected.Rows.Count = 0 Then
            MsgBox("You Must Select Orders before Pre-Allocating", MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Dim CUST_CODE_pre As String = ""
        Dim ORDR_GROUP_NO_pre As String = ""

        For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLO1.Selected.Rows
            If CUST_CODE_pre = "" Then
                CUST_CODE_pre = grow.Cells("CUST_CODE").Value
            Else
                If CUST_CODE_pre <> grow.Cells("CUST_CODE").Value Then
                    MsgBox("Orders must be for Same Customer", MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Exit Sub
                End If
            End If
            ORDR_GROUP_NO_pre &= ",'" & grow.Cells("ORDR_NO").Value & "'"
        Next

        Me.Cursor = Cursors.WaitCursor

        ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR1.CUST_STORE_NO" & vbCrLf _
            & ", NVL(ARTCUST2.CUST_RANK,0) CUST_RANK" & vbCrLf _
            & ", NVL(SOTORDR2.ORDR_QTY_PRE_ALLO,0) ORDR_QTY_PRE_ALLO" & vbCrLf _
            & ", NVL(SOTORDR2.ORDR_QTY_ALLO,0) ORDR_QTY_ALLO, " & vbCrLf _
            & " SOTORDR2.ORDR_QTY_OPEN" & vbCrLf _
            & " from SOTORDR1,ARTCUST2,SOTORDR2" & vbCrLf _
            & " where SOTORDR1.CUST_CODE = '" & CUST_CODE_pre & "'" & vbCrLf _
            & "   and SOTORDR1.ORDR_GROUP_NO IN (" & Mid$(ORDR_GROUP_NO_pre, 2) & ")" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_ADDR_TYPE = 'MK'" & vbCrLf _
            & "   and ARTCUST2.CUST_ADDR_CODE = SOTORDR1.CUST_STORE_NO" & vbCrLf
        Fill_Records("SOTPREA1", "", True, ASCMAIN1.sql)

        For Each row As DataRow In dst.Tables("SOTPREA1").Select("ISNULL(CUST_RANK,0) = 0", "")
            row.Item("CUST_RANK") = 9999
        Next

        grdSOTPREA1.Text = "Pre-Allo by Store for " & CUST_CODE_pre & " Order Groups " & Mid(ORDR_GROUP_NO_pre, 2)

        Me.Cursor = Cursors.Default

        frmPreAllocate.Visible = True
        tabMain.Visible = False

    End Sub

    Private Sub cmdUPC_Click()

        ASCMAIN1.sql = "Select UPC_CODE, STYLE_CODE, COLOR_CODE, SIZE_CODE FROM(" & vbCrLf _
            & " SELECT STYLE_CODE, COLOR_CODE, COLOR_CODE_UPC, 'ALL' AS SIZE_CODE, UPC_CODE" & vbCrLf _
            & " FROM ICTSTYC2" & vbCrLf _
            & " WHERE UPC_CODE IS NOT NULL" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT ICTSTYC4.STYLE_CODE, ICTSTYC4.COLOR_CODE, ICTSTYC4.COLOR_CODE_UPC," & vbCrLf _
            & " ICTSTYC3.SIZE_CODE, ICTSTYC4.UPC_CODE" & vbCrLf _
            & " FROM ICTSTYC3, ICTSTYC4" & vbCrLf _
            & " WHERE ICTSTYC3.STYLE_CODE (+) = ICTSTYC4.STYLE_CODE" & vbCrLf _
            & " AND ICTSTYC3.COLOR_CODE (+) = ICTSTYC4.COLOR_CODE" & vbCrLf _
            & " AND ICTSTYC3.SIZE_INDEX (+) = ICTSTYC4.SIZE_INDEX)" & vbCrLf _
            & " WHERE UPC_CODE like '049183%'" & vbCrLf _
            & " ORDER BY STYLE_CODE, COLOR_CODE, SIZE_CODE" & vbCrLf
        Dim DT As DataTable = ASCDATA1.GetDataTable

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("UPC_CODE")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            Dim CODE_VALUES As String = ""
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Absx1.txtFor("STYLE_CODE").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STYLE_CODE")
                optSelectBy.Value = "S"
                Click_Command("Select")
            End If
        End If
    End Sub

    Private Sub grdSOTINVHX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTINVHX.AfterRowActivate
        Load_SOTINVHY()
    End Sub

    Sub Load_SOTINVHY()
        If grdSOTINVHX.ActiveRow Is Nothing OrElse Not grdSOTINVHX.ActiveRow.IsDataRow Then
            grdSOTINVHY.Visible = False
        Else
            grdSOTINVHY.Visible = True

            Me.Cursor = Cursors.WaitCursor
            Dim CUST_CODE As String = grdSOTINVHX.ActiveRow.Cells("CUST_CODE").Value
            Dim CUST_NAME As String = grdSOTINVHX.ActiveRow.Cells("CUST_NAME").Value & ""

            If grdSOTINVHX.Tag & "" = "S" Then
                Dim sp As String = cmbYP_From.Value
                Dim ep As String = cmbYP_To.Value
                sp = Mid(sp, 1, 4) & Mid(sp, 6, 2)
                ep = Mid(ep, 1, 4) & Mid(ep, 6, 2)

                If chkSumPO.Checked Then
                    '                        & ", Sum (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _

                    ASCMAIN1.sql = "SELECT SOTINVH1.ORDR_CUST_PO, 'All' CUST_STORE_NO, SOTINVH2.COLOR_CODE" & vbCrLf _
                        & ", Sum (SOTINVH2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                        & ", Min (SOTINVH2.ORDR_UNIT_PRICE) ORDR_UNIT_PRICE" & vbCrLf _
                        & ", Min (SOTINVH2.INV_NO) INV_NO, SOTINVH1.INV_DATE" & vbCrLf _
                        & ", SOTINVH1.WHSE_CODE" & vbCrLf _
                        & " from SOTINVH2,SOTINVH1 where SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                        & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                        & "   and SOTINVH1.ORDR_TYPE_CODE <> 'XFR'" & vbCrLf _
                        & " and SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                        & " and SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & sp & "'" & vbCrLf _
                        & " and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & ep & "'" & vbCrLf _
                        & IIf(chkSingleColor.Checked, " and SOTINVH2.COLOR_CODE = '" & COLOR_CODE & "'", "") & vbCrLf _
                        & " GROUP by SOTINVH1.ORDR_CUST_PO, SOTINVH1.INV_DATE, SOTINVH2.COLOR_CODE, SOTINVH1.WHSE_CODE" & vbCrLf
                Else
                    ASCMAIN1.sql = "SELECT SOTINVH1.ORDR_CUST_PO, SOTINVH1.CUST_STORE_NO, SOTINVH2.COLOR_CODE" & vbCrLf _
                        & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.WHSE_CODE" & vbCrLf _
                        & " from SOTINVH2,SOTINVH1 where SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                        & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                        & " and SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                        & " and SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & sp & "'" & vbCrLf _
                        & " and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & ep & "'" & vbCrLf _
                        & "   and SOTINVH1.ORDR_TYPE_CODE <> 'XFR'" & vbCrLf _
                        & IIf(chkSingleColor.Checked, " and SOTINVH2.COLOR_CODE = '" & COLOR_CODE & "'", "") & vbCrLf _
                        & " order by SOTINVH1.ORDR_CUST_PO, SOTINVH1.CUST_STORE_NO, SOTINVH2.COLOR_CODE" & vbCrLf
                End If
            Else
                ASCMAIN1.sql = "SELECT SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR2.COLOR_CODE, " & vbCrLf _
                    & " Decode (SOTORDR2.ORDR_STATUS, 'O', ORDR_QTY_OPEN, ORDR_QTY_PICK) ORDR_QTY_SHIP " & vbCrLf _
                    & ", SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                    & ", SOTORDR2.ORDR_STATUS INV_NO, SOTORDR1.ORDR_SHIP_DATE INV_DATE, SOTORDR1.WHSE_CODE" & vbCrLf _
                    & " from SOTORDR2,SOTORDR1" & vbCrLf _
                    & " where SOTORDR1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & " and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & " and (SOTORDR2.ORDR_STATUS = 'O' or SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf _
                    & "   and SOTORDR1.ORDR_TYPE_CODE <> 'XFR'" & vbCrLf _
                    & " and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                    & IIf(chkSingleColor.Checked, " and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'", "") & vbCrLf _
                    & " order by SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR2.COLOR_CODE" & vbCrLf
            End If
            Fill_Records("SOTINVHY", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdSOTINVHY, "ORDR_CUST_PO,CUST_STORE_NO,COLOR_CODE")

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                ' need to go thru details looking for a price which has fractions of a cent
                Dim extra_decimals As Boolean = False
                For Each row As DataRow In dst.Tables("SOTINVHY").Select("")
                    Dim ORDR_UNIT_PRICE As Decimal = Val(row.Item("ORDR_UNIT_PRICE") & "")
                    If ORDR_UNIT_PRICE <> Val(Format(ORDR_UNIT_PRICE, "#.00")) Then
                        extra_decimals = True
                        Exit For
                    End If
                Next

                With grdSOTINVHY.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE")
                    If extra_decimals Then
                        .Format = "#,##0.0000"
                    Else
                        .Format = "#,##0.00"
                    End If
                End With
            End If

            grdSOTINVHY.Visible = True
            Me.Cursor = Cursors.Default
            grdSOTINVHY.Text = "Customer " & grdSOTINVHX.ActiveRow.Cells("CUST_CODE").Value & ":" & CUST_NAME & IIf(chkSingleColor.Checked, ", Color " & COLOR_CODE, ", All Colors")
        End If

    End Sub

#Region "grdICTTRANX"
    Private Sub grdICTTRANX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTTRANX.InitializeRow
        'If e.Row.Cells("TRAN_STATUS_UPD").Value = "R" Then
        '    e.Row.Cells("TRAN_STATUS_UPD").Appearance.ForeColor = Color.Red
        'End If
    End Sub
#End Region

#Region "grdICTSTATA"

    Private Sub grdICTSTATA_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTATA.AfterRowActivate
        If STYLE_CODE <> STYLE_CODE_allocated Then AutoAllocate = False
        Setup_SC()
        If chkAutoAllocate.Checked And Not multi_style Then
            Allocate()
        End If
    End Sub

    Private Sub grdICTSTATA_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTATA.InitializeRow
        If e.Row.Band.Key = "ICTSTATA" Then
            Dim STYLE_COLOR_STATUS As String = e.Row.Cells("STYLE_COLOR_STATUS").Value & ""
            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                Select Case STYLE_COLOR_STATUS
                    Case "D"
                        e.Row.Cells("COLOR_CODE").Appearance.BackColor = Color.Red
                        e.Row.Cells("COLOR_DESC").Appearance.BackColor = Color.Red
                        e.Row.ToolTipText = "Style/Color " & e.Row.Cells("STYLE_CODE").Value & "/" & e.Row.Cells("COLOR_CODE").Value & " is Discontinued"
                    Case "N"
                        e.Row.Cells("COLOR_CODE").Appearance.BackColor = Color.Yellow
                        e.Row.Cells("COLOR_DESC").Appearance.BackColor = Color.Yellow
                        e.Row.ToolTipText = "Style/Color " & e.Row.Cells("STYLE_CODE").Value & "/" & e.Row.Cells("COLOR_CODE").Value & " is Do Not Re-Order"
                End Select
            Else
                If STYLE_COLOR_STATUS = "D" Then
                    e.Row.Appearance.ForeColor = Color.Red
                    e.Row.ToolTipText = "Style/Color " & e.Row.Cells("STYLE_CODE").Value & "/" & e.Row.Cells("COLOR_CODE").Value & " is Discontinued"
                End If
            End If
        End If
    End Sub
#End Region

    Sub Setup_Tab()
        EnforceConstraints(False)
        Setup_SO()
        Setup_PO()

        Setup_Allocations()
        Setup_tabMain()
        'Clear_Allo()

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            STYLE_CLASS_CODE = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
            If STYLE_CLASS_CODE = "" Then
                MsgBox("Warning: Style " & STYLE_CODE & " does not have a Class Code",
                       MsgBoxStyle.OkOnly, "Please Assign one Immediately")
            End If
            CARTON_PACK_QTY = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
            STYLE_PRICE = Val(rowICTSTYL1.Item("STYLE_PRICE") & "")
            Price_and_Availability(STYLE_CODE, STYLE_CLASS_CODE, COLOR_CODE, CARTON_PACK_QTY, STYLE_PRICE)
        End If

        Setup_Tran()
        EnforceConstraints(True)
    End Sub

    Sub Load_SOTALLO1()
        ' Load SD Table

        dst.Tables("SOTALLO1").Rows.Clear()

        '        & ", MIN (FORMAT$(ORDR_DEMAND_DATE,'MM/DD/YY')) AS SD_DATE_X" & vbCrLf _

        ASCMAIN1.sql = "Select SOTDEMD1.ORDR_GROUP_NO ORDR_NO" & vbCrLf _
            & ", Decode (SOTDEMD1.DEMAND_TYPE,'R',SOTDEMD1.ORDR_LNO,1) ORDR_LNO" & vbCrLf _
            & ", SOTDEMD1.CUST_CODE, ARTCUST1.CUST_NAME, MIN (SOTDEMD1.ORDR_CUST_PO) ORDR_CUST_PO" & vbCrLf _
            & ", '1' AS RECORD_TYPE" & vbCrLf _
            & ", MIN (SOTDEMD1.DEMAND_TYPE) AS RECORD_SUB_TYPE" & vbCrLf _
            & ", SOTDEMD1.WHSE_CODE, SOTDEMD1.STYLE_CODE, SOTDEMD1.COLOR_CODE" & vbCrLf _
            & ", MIN (SOTDEMD1.ORDR_PRIORITY) ORDR_PRIORITY" & vbCrLf _
            & ", MIN (SOTDEMD1.ORDR_DEMAND_DATE) AS SD_DATE, Null SD_DATE_X " & vbCrLf _
            & ", MIN (SOTDEMD1.ORDR_CANCEL_DATE) ORDR_CANCEL_DATE" & vbCrLf _
            & ", MIN (SOTDEMD1.ORDR_SHIP_DATE) ORDR_SHIP_DATE, NULL SHIP_ETA" & vbCrLf _
            & ", SUM (SOTDEMD1.ORDR_QTY_OPEN) AS SD_QTY" & vbCrLf _
            & ", SUM (SOTDEMD1.ORDR_QTY_ALLO) AS SD_QTY_ALLO " & vbCrLf _
            & ", SUM (SOTDEMD1.ORDR_QTY_ALLO_CUR) AS SD_QTY_ALLO_CUR " & vbCrLf _
            & ", SUM (SOTDEMD1.ORDR_QTY_ALLO_FUT) AS SD_QTY_ALLO_FUT " & vbCrLf _
            & ", SUM (SOTDEMD1.ORDR_QTY_ALLO_CXL) AS SD_QTY_ALLO_CXL " & vbCrLf _
            & ", NULL BALANCE, NULL ORDR_RELEASE" & vbCrLf _
            & ", MIN (SOTDEMD1.ORDR_DEMAND_DATE) AS ORDR_DEMAND_DATE " & vbCrLf _
            & ", MIN (SOTDEMD1.ORDR_PRIORITY_DATE) AS ORDR_PRIORITY_DATE " & vbCrLf _
            & ", MIN (SOTDEMD1.ORDR_PRIORITY_DATE_ORIG) AS ORDR_PRIORITY_DATE_ORIG " & vbCrLf _
            & ", MAX (SOTDEMD1.ORDR_RELEASE_AVAIL) AS ORDR_RELEASE_AVAIL " & vbCrLf _
            & " from " & SOTDEMD1 & " SOTDEMD1,ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTDEMD1.CUST_CODE" & vbCrLf _
            & " group by SOTDEMD1.CUST_CODE, ARTCUST1.CUST_NAME, SOTDEMD1.ORDR_GROUP_NO" & vbCrLf _
            & ", Decode (SOTDEMD1.DEMAND_TYPE,'R',SOTDEMD1.ORDR_LNO,1)" & vbCrLf _
            & ", SOTDEMD1.WHSE_CODE, SOTDEMD1.STYLE_CODE, SOTDEMD1.COLOR_CODE"

        ' "Select " _
        '& ", SUM (ORDR_QTY_ALLO) AS SD_QTY_ALLO " _
        '& ", SUM (ORDR_QTY_ALLO_CUR) AS SD_QTY_ALLO_CUR " _
        '& ", SUM (ORDR_QTY_ALLO_FUT) AS SD_QTY_ALLO_FUT " _
        '& ", MIN (SOTDEMD1.ORDR_PRIORITY_DATE_ORIG) AS ORDR_PRIORITY_DATE_ORIG " _
        '& ", MAX (SOTDEMD1.ORDR_RELEASE_AVAIL) AS ORDR_RELEASE_AVAIL " _
        '& " from ASW29725 SOTDEMD1" _
        '& " group by CUST_CODE, ORDR_GROUP_NO, Decode (DEMAND_TYPE,'R',ORDR_LNO,1), STYLE_CODE, COLOR_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").NewRow
            'rowSOTALLO1.ItemArray = row.ItemArray
            For i As Integer = 0 To row.Table.Columns.Count - 1
                rowSOTALLO1.Item(i) = row.Item(i)
            Next i
            If ASCMAIN1.CLIENT = "RGI" Then
                If ASCMAIN1.Running_in_VS AndAlso rowSOTALLO1.Item("ORDR_NO") = "0000476748" Then Stop
                rowSOTALLO1.Item("SD_DATE") = row.Item("ORDR_PRIORITY_DATE")
            End If
            If row.Item("RECORD_SUB_TYPE") = "O" Then
                rowSOTALLO1.Item("ORDR_RELEASE") = "H"
                Dim ORDR_GROUP_NO As String = row.Item("ORDR_NO")
                'Dim rowSOTORDR7 As DataRow = Fill_Record("SOTORDR7", New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})
                Dim rowSOTORDR7 As DataRow = LookUp("SOTORDR7", New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})
                If rowSOTORDR7 IsNot Nothing Then
                    Dim ORDR_RELEASE As String = rowSOTORDR7.Item("ORDR_RELEASE") & ""
                    If ORDR_RELEASE <> "" Then
                        rowSOTALLO1.Item("ORDR_RELEASE") = ORDR_RELEASE
                    End If
                End If
            End If

            dst.Tables("SOTALLO1").Rows.Add(rowSOTALLO1)
        Next

        '        & ", MID$(SUPPLY_DATE,5,2) & '/' & MID$(SUPPLY_DATE,7,2) & '/' & MID$(SUPPLY_DATE,1,4) AS SD_DATE" & vbCrLf _
        '        & ", MID$(SUPPLY_DATE,5,2) & '/' & MID$(SUPPLY_DATE,7,2) & '/' & MID$(SUPPLY_DATE,3,2) AS SD_DATE_X" & vbCrLf _
        '        & ", '00/00/0000' AS ORDR_DEMAND_DATE" & vbCrLf _

        'Sql = "Insert into SOWALLO1"
        ASCMAIN1.sql = " Select DECODE(SUPPLY_TYPE,'S',PO_ORDER_NO,ORDR_NO) ORDR_NO, DECODE(SUPPLY_TYPE,'S',PO_ORDER_LNO,ORDR_LNO) ORDR_LNO" & vbCrLf _
            & ", SUBSTR(PO_SHIP_VESSEL,1,10) CUST_CODE" & vbCrLf _
            & ", PO_REFERENCE AS ORDR_CUST_PO" & vbCrLf _
            & ", '0' AS RECORD_TYPE, SUPPLY_TYPE AS RECORD_SUB_TYPE" & vbCrLf _
            & ", WHSE_CODE, STYLE_CODE, COLOR_CODE, NULL AS ORDR_PRIORITY" & vbCrLf _
            & ", NULL AS ORDR_DEMAND_DATE" & vbCrLf _
            & ", DECODE(SUPPLY_DATE,'00000000',NULL,TO_DATE(SUBSTR(SUPPLY_DATE,5,2) || '/' || SUBSTR(SUPPLY_DATE,7,2) || '/' || SUBSTR(SUPPLY_DATE,1,4),'MM/DD/YYYY')) AS SD_DATE" & vbCrLf _
            & ", NULL AS ORDR_CANCEL_DATE, NULL AS ORDR_SHIP_DATE " & vbCrLf _
            & ", PO_SHIP_ETA AS SHIP_ETA " & vbCrLf _
            & ", SUPPLY_QTY AS SD_QTY, Null AS SD_QTY_ALLO  " & vbCrLf _
            & " from " & SOTSUPP1 & " SOTSUPP1"
        Fill_Records("SOTALLO1", "", False, ASCMAIN1.sql)

        For Each row As DataRow In dst.Tables("SOTALLO1").Rows
            If row.Item("RECORD_TYPE") = "0" And row.Item("RECORD_SUB_TYPE") <> "H" And row.Item("SD_DATE") & "" <> "" Then
                Dim SD_DATE As Date = row.Item("SD_DATE")
                row.Item("SD_DATE_X") = Format(SD_DATE, "MM/dd/yy")
            End If
            If Val(row.Item("SD_QTY_ALLO_CUR") & "") = 0 Then row.Item("SD_QTY_ALLO_CUR") = DBNull.Value
            If Val(row.Item("SD_QTY_ALLO_FUT") & "") = 0 Then row.Item("SD_QTY_ALLO_FUT") = DBNull.Value
            If Val(row.Item("SD_QTY_ALLO_CXL") & "") = 0 Then row.Item("SD_QTY_ALLO_CXL") = DBNull.Value
        Next

        Setup_ASL()
        Setup_Allocations()
    End Sub

    Sub Setup_Allocations()
        dst.Tables("SOTALLO1").DefaultView.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
        dst.Tables("ICTSTDQ1").DefaultView.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

        With grdSOTALLO1.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("WHSE_CODE", False, True)
            .SortedColumns.Add("SD_DATE", False)
            .SortedColumns.Add("RECORD_TYPE", False)
            If ROWs("SOTPARM1").Item("SO_PARM_ALLO_SEQ") & "" = "1" Then
                .SortedColumns.Add("ORDR_DEMAND_DATE", False)
            Else
                .SortedColumns.Add("ORDR_PRIORITY_DATE", False)
            End If
            ' .SortedColumns.Add("RECORD_TYPE", False)
        End With
        grdSOTALLO1.Rows.ExpandAll(True)
        grdSOTALLO1.Text = "Allocations Plan by Warehouse for Style-Color " & STYLE_CODE & "-" & COLOR_CODE

        Setup_ICTSTDQ1()
    End Sub

    Sub Setup_ICTSTDQ1()
        With grdICTSTDQ1.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("WHSE_CODE", False, True)
            .SortedColumns.Add("STATUS_DATE", False)
        End With
        Dim I As Integer = 0
        Dim EXPANDED As Boolean = False
        Do While Not EXPANDED And I < 5
            Try
                grdICTSTDQ1.Rows.ExpandAll(True)
                EXPANDED = True
            Catch ex As Exception
                I += 1
            End Try
        Loop
        grdICTSTDQ1.Text = STYLE_CODE & "-" & COLOR_CODE
    End Sub

    Sub Set_Table()

        For Each rowWSC As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("SOTALLO1"), New String() {"WHSE_CODE", "STYLE_CODE", "COLOR_CODE"}).Select("")
            Dim WHSE_CODE As String = rowWSC.Item("WHSE_CODE")
            Dim STYLE_CODE As String = rowWSC.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWSC.Item("COLOR_CODE")
            Dim sqlWSC As String = "WHSE_CODE = '" & WHSE_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

            Dim S As Integer
            Dim QTY As Int64
            Dim BALANCE As Int64 = 0
            Dim SD As String = ""
            Dim SD_last As String = ""

            ' Calculate Running Balance

            Dim seq As String = "SD_DATE, RECORD_TYPE, RECORD_SUB_TYPE"
            If ASCMAIN1.CLIENT = "RGI" AndAlso blnATONCE Then
                seq = "SD_DATE, RECORD_TYPE, ORDR_PRIORITY_DATE, ORDR_PRIORITY_DATE_ORIG, RECORD_SUB_TYPE"
            End If
            For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select(sqlWSC, seq)
                If rowSOTALLO1.Item("RECORD_TYPE") & "" = "0" Then
                    S = 1
                    QTY = Val(rowSOTALLO1.Item("SD_QTY") & "")
                    If rowSOTALLO1.Item("SD_DATE") & "" = "" Then
                        SD_last = "00000000"
                        SD &= "00000000"
                    Else
                        Dim SD_DATE As String = Format(rowSOTALLO1.Item("SD_DATE"), "yyyyMMdd")
                        If SD_last <> SD_DATE Then
                            SD &= SD_DATE
                            SD_last = SD_DATE
                        End If
                    End If
                Else
                    S = -1
                    QTY = Val(rowSOTALLO1.Item("SD_QTY_ALLO") & "")
                    If rowSOTALLO1.Item("RECORD_SUB_TYPE") & "" = "O" Then


                        If edi850cust.Contains(rowSOTALLO1.Item("CUST_CODE")) Then
                            rowSOTALLO1.Item("ORDR_BACKORDER") = "0"
                        Else
                            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowSOTALLO1.Item("CUST_CODE"))
                            If rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & "" = "1" _
                            Or (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
                                rowSOTALLO1.Item("ORDR_BACKORDER") = "1"
                            Else
                                rowSOTALLO1.Item("ORDR_BACKORDER") = "0"
                            End If
                        End If

                        Dim rowSOTORDR7 = dst.Tables("SOTORDR7").Rows.Find(New String() {rowSOTALLO1.Item("ORDR_NO"), STYLE_CODE, COLOR_CODE})
                        If rowSOTORDR7 IsNot Nothing Then
                            If rowSOTORDR7.Item("ORDR_BACKORDER") & "" = "Y" Then
                                rowSOTALLO1.Item("ORDR_BACKORDER") = "1"
                            ElseIf rowSOTORDR7.Item("ORDR_BACKORDER") & "" = "N" Then
                                rowSOTALLO1.Item("ORDR_BACKORDER") = "0"
                            End If
                        End If
                    End If
                End If
                BALANCE += S * QTY
                rowSOTALLO1.Item("BALANCE") = BALANCE
            Next
        Next

        ' Prepare for Display
        grdSOTALLO1.Visible = True

    End Sub

#Region "grdSOTORDRX"

    Private Sub grdSOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDRX.AfterRowActivate
        If optDetails.Value <> "A" Then
            optDetails.Value = "A"
        End If
    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow
        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then

            If e.Row.Cells("ORDR_TYPE").Value = "R" Then
                e.Row.Cells("CUST_CODE").Appearance.ForeColor = Color.Red
                e.Row.Cells("ORDR_CUST_PO").Appearance.ForeColor = Color.Red
            End If

            If e.Row.Cells("SUB").Value & "" = "1" Then
                e.Row.Cells("GP_PCT").Appearance.ForeColor = Color.Red
                e.Row.Cells("SUB").Appearance.ForeColor = Color.Red
            End If

            If e.Row.Cells("ERROR").Value & "" <> "" Then
                e.Row.Cells("ERROR").Appearance.ForeColor = Color.Red
            End If

            If e.Row.Cells("SHIP_DATE_PLUS").Value & "" <> e.Row.Cells("ORDR_SHIP_DATE").Value & "" AndAlso e.Row.Cells("STYLE_AT_ONCE_ACTIVE").Value & "" = "1" AndAlso Format(e.Row.Cells("STYLE_AT_ONCE_UNTIL").Value & "", "yyyyMMdd") >= Format(Now, "yyyyMMdd") = "1" Then
                e.Row.Cells("SHIP_DATE_PLUS").Appearance.BackColor = Color.Yellow
                e.Row.Cells("SHIP_DATE_PLUS").ToolTipText = "Original Ship Date = " & e.Row.Cells("ORDR_SHIP_DATE").Value
            Else
                e.Row.Cells("SHIP_DATE_PLUS").Appearance.BackColor = Color.Empty
                e.Row.Cells("SHIP_DATE_PLUS").ToolTipText = ""
            End If


            If e.Row.Cells("STYLE_AT_ONCE_UNTIL").Value & "" <> "" Then

                If Format(e.Row.Cells("STYLE_AT_ONCE_UNTIL").Value, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                    e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Red
                    e.Row.Cells("STYLE_AT_ONCE_UNTIL").ToolTipText = "This At-Once Parameter has expired"
                Else
                    e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Empty
                End If

                If e.Row.Cells("STYLE_AT_ONCE_ACTIVE").Value & "" = "1" Then
                    e.Row.Cells("STYLE_AT_ONCE_ACTIVE").Appearance.BackColor = Color.Empty
                Else
                    e.Row.Cells("STYLE_AT_ONCE_ACTIVE").Appearance.BackColor = Color.Red
                End If
            End If

        End If

    End Sub

#End Region

#Region "grdSOTALLO1"

    Private Sub grdSOTALLO1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTALLO1.AfterRowUpdate
        Dim RECORD_SUB_TYPE As String = e.Row.Cells("RECORD_SUB_TYPE").Value & ""
        Dim ORDR_RELEASE As String = e.Row.Cells("ORDR_RELEASE").Value & ""

        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Value)

        If RECORD_SUB_TYPE = "O" Then
            Dim ORDR_GROUP_NO As String = e.Row.Cells("ORDR_NO").Value
            ASCMAIN1.sql = "Select * from SOTORDR7 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            Fill_Records("SOTORDR7", "", True, ASCMAIN1.sql)

            Dim rowSOTORDR7 As DataRow = Nothing ' Fill_Record("SOTORDR7", New String() {e.Row.Cells("ORDR_NO").Value, STYLE_CODE, COLOR_CODE})
            If dst.Tables("SOTORDR7").Rows.Count = 1 Then
                rowSOTORDR7 = dst.Tables("SOTORDR7").Rows(0)
            End If

            Dim keep_record As Boolean = False
            Dim New_Record As Boolean = False

            If rowSOTORDR7 Is Nothing Then
                rowSOTORDR7 = dst.Tables("SOTORDR7").NewRow
                rowSOTORDR7.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                rowSOTORDR7.Item("STYLE_CODE") = STYLE_CODE
                rowSOTORDR7.Item("COLOR_CODE") = COLOR_CODE
                dst.Tables("SOTORDR7").Rows.Add(rowSOTORDR7)
                New_Record = True
            End If
            If ORDR_RELEASE <> "H" Then
                rowSOTORDR7.Item("ORDR_RELEASE") = ORDR_RELEASE
                keep_record = True
            Else
                rowSOTORDR7.Item("ORDR_RELEASE") = DBNull.Value
            End If
            If e.Row.Cells("ORDR_PRIORITY").Value & "" <> rowARTCUST1.Item("CUST_PRIORITY_CODE") & "" Then
                rowSOTORDR7.Item("ORDR_PRIORITY") = e.Row.Cells("ORDR_PRIORITY").Value
                keep_record = True
            Else
                rowSOTORDR7.Item("ORDR_PRIORITY") = DBNull.Value
            End If

            If Format(e.Row.Cells("ORDR_PRIORITY_DATE").Value, "yyyyMMdd") _
            <> Format(e.Row.Cells("ORDR_PRIORITY_DATE_ORIG").Value, "yyyyMMdd") Then
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = e.Row.Cells("ORDR_PRIORITY_DATE").Value
                keep_record = True
            Else
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = DBNull.Value
            End If

            If Format(e.Row.Cells("ORDR_DEMAND_DATE").Value, "yyyyMMdd") _
            <> Format(DateValue(e.Row.Cells("ORDR_CANCEL_DATE").Value).AddDays(
                 +Val(rowARTCUST1.Item("CUST_CANCEL_GRACE_DAYS") & "")), "yyyyMMdd") Then
                rowSOTORDR7.Item("ORDR_DEMAND_DATE") = e.Row.Cells("ORDR_DEMAND_DATE").Value
                keep_record = True
            Else
                rowSOTORDR7.Item("ORDR_DEMAND_DATE") = DBNull.Value
            End If

            Dim CUST_ALLOW_BACKORDER As Boolean = (rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & "" = "1")
            ' as per GABE 2/21/01
            ' Gabe says that all customers allow back order as default. REMMED OUT THE NEX 5 LINES. I disagree 11/16/01.
            '        If InStr(edi850cust, grdSOWALLO1.Columns("CUST_CODE").Text) = 0 Then ' IF NOT AN EDI CUSTOMER
            '            CUST_ALLOW_BACKORDER = True
            '        Else
            '            CUST_ALLOW_BACKORDER = False
            '        End If
            If Val(e.Row.Cells("ORDR_BACKORDER").Value & "") = 0 And CUST_ALLOW_BACKORDER Then
                rowSOTORDR7.Item("ORDR_BACKORDER") = "N"
                keep_record = True
            ElseIf Val(e.Row.Cells("ORDR_BACKORDER").Value & "") <> 0 And Not CUST_ALLOW_BACKORDER Then
                rowSOTORDR7.Item("ORDR_BACKORDER") = "Y"
                keep_record = True
            Else
                rowSOTORDR7.Item("ORDR_BACKORDER") = DBNull.Value
            End If

            If keep_record Then
                Update_Record_TDA("SOTORDR7")
            Else
                If Not New_Record Then
                    rowSOTORDR7.Delete()
                    Update_Record_TDA("SOTORDR7")
                End If
            End If
        Else
            Dim rowSOTRSRV2 As DataRow = Fill_Record("SOTRSRV2", New Object() {e.Row.Cells("ORDR_NO").Value,
                                                                               e.Row.Cells("ORDR_LNO").Value})
            If e.Row.Cells("ORDR_PRIORITY").Value & "" <> rowARTCUST1.Item("CUST_PRIORITY_CODE") Then
                rowSOTRSRV2.Item("RSRV_PRIORITY") = e.Row.Cells("ORDR_PRIORITY").Value & ""
            Else
                rowSOTRSRV2.Item("RSRV_PRIORITY") = DBNull.Value
            End If

            If Format(e.Row.Cells("ORDR_PRIORITY_DATE").Value & "", "yyyyMMdd") _
            <> Format(e.Row.Cells("ORDR_PRIORITY_DATE_ORIG").Value & "", "yyyyMMdd") Then
                rowSOTRSRV2.Item("RSRV_PRIORITY_DATE") = e.Row.Cells("ORDR_PRIORITY_DATE").Value & ""
            Else
                rowSOTRSRV2.Item("RSRV_PRIORITY_DATE") = DBNull.Value
            End If

            If Format(e.Row.Cells("ORDR_DEMAND_DATE").Value & "", "yyyyMMdd") _
            <> Format(DateValue(e.Row.Cells("ORDR_CANCEL_DATE").Value & "").AddDays(
                 Val(rowARTCUST1.Item("CUST_CANCEL_GRACE_DAYS") & "")), "yyyyMMdd") Then
                rowSOTRSRV2.Item("RSRV_DEMAND_DATE") = e.Row.Cells("ORDR_DEMAND_DATE").Value & ""
            Else
                rowSOTRSRV2.Item("RSRV_DEMAND_DATE") = DBNull.Value
            End If
            Update_Record_TDA("SOTRSRV2")
        End If
    End Sub

    Private Sub grdSOTALLO1_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTALLO1.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case "ORDR_DEMAND_DATE"
                Dim ORDR_DEMAND_DATE As String = Format(e.Cell.EditorResolved.Value, "yyyyMMdd")
                Dim ORDR_CANCEL_DATE As String = Format(e.Cell.Row.Cells("ORDR_CANCEL_DATE").Value, "yyyyMMdd")
                If ORDR_DEMAND_DATE > ORDR_CANCEL_DATE Then
                    If MsgBox("OK to change Demand Date", MsgBoxStyle.YesNo,
                              "New Demand Date is Later than Cancel Date") = MsgBoxResult.No Then
                        e.Cancel = True
                    End If
                End If
        End Select
    End Sub

    Private Sub grdSOTALLO1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTALLO1.ClickCellButton

        Dim z As String = ""

        Select Case e.Cell.Column.Key
            Case "ORDR_PRIORITY"
                If e.Cell.Row.Cells("RECORD_TYPE").Text = "1" Then
                    z = e.Cell.Value
                    If Val(z) >= "9" Then
                        z = "1"
                    Else
                        z = Format(Val(z) + 1, "0")
                    End If
                    e.Cell.Value = z
                End If
            Case "ORDR_RELEASE"
                If e.Cell.Row.Cells("RECORD_TYPE").Text = "1" Then
                    If e.Cell.Row.Cells("RECORD_SUB_TYPE").Text = "O" Then
                        z = e.Cell.Value & ""
                        If Trim(z) = "" Then
                            z = "H"
                        End If
                        If z = "H" Then
                            z = "C"
                        ElseIf z = "C" Then
                            z = "S"
                        ElseIf z = "S" Then
                            z = "X"
                        Else
                            z = "H"
                        End If
                        e.Cell.Value = z
                    End If
                End If
        End Select
    End Sub

    Private Sub grdSOTALLO1_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTALLO1.DoubleClickCell
        'If e.Cell.Column.Key = "ORDR_PRIORITY_DATE" _
        'Or e.Cell.Column.Key = "ORDR_DEMAND_DATE" Then
        '    ssdSOWALLO1.Left = grdSOWALLO1.Left + grdSOWALLO1.Columns(grdSOWALLO1.Col).Left + grdSOWALLO1.Columns(grdSOWALLO1.Col).Width
        '    ssdSOWALLO1.Top = grdSOWALLO1.Top + grdSOWALLO1.Columns(grdSOWALLO1.Col).Top + grdSOWALLO1.RowHeight
        '    ssdSOWALLO1.Text = Format$(grdSOWALLO1.Columns(grdSOWALLO1.Col).Text, "MM/DD/YYYY")
        '    ssdSOWALLO1.DroppedDown = True
        'End If
    End Sub

    Private Sub grdSOTALLO1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTALLO1.DoubleClickRow
    End Sub

    Private Sub grdSOTALLO1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLO1.InitializeRow
        With e.Row
            Dim RECORD_SUB_TYPE As String = .Cells("RECORD_SUB_TYPE").Value & ""
            If RECORD_SUB_TYPE = "H" Then
                .Cells("CUST_CODE").Appearance.BackColor = Color.Yellow
                .Cells("CUST_CODE").ToolTipText = "On Hand Open to Ship"
            ElseIf RECORD_SUB_TYPE = "S" Then
                .Appearance.ForeColor = Color.Blue
                .Cells("CUST_CODE").ToolTipText = "PO Shipment"
            ElseIf RECORD_SUB_TYPE = "R" Then
                .Cells("CUST_CODE").Appearance.ForeColor = Color.Red
                .Cells("ORDR_CUST_PO").Appearance.ForeColor = Color.Red
                .Cells("ORDR_CUST_PO").ToolTipText = "Reservation"
            End If

            If RECORD_SUB_TYPE = "O" _
            Or RECORD_SUB_TYPE = "R" Then
                If Val(.Cells("SD_QTY_ALLO").Value & "") <> Val(.Cells("SD_QTY").Value & "") Then
                    .Cells("SD_QTY_ALLO").Appearance.ForeColor = Color.Red
                    .Cells("SD_QTY_ALLO").ToolTipText = "Unable to Allocate Total Qty Open"
                End If
                If .Cells("ORDR_RELEASE_AVAIL").Value & "" <> "" Then
                    If Format(.Cells("ORDR_RELEASE_AVAIL").Value, "yyyyMMdd") >
                       Format(.Cells("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                        .Cells("SD_QTY_ALLO").Appearance.ForeColor = Color.Red
                        .Cells("SD_QTY_ALLO").ToolTipText = "Supply becomes available after Cancel Date"
                        .Cells("SD_QTY_ALLO_CXL").ToolTipText = "Supply becomes available after Cancel Date"
                    End If
                Else
                    If Format(Now, "yyyyMMdd") >
                       Format(.Cells("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                        .Cells("ORDR_CANCEL_DATE").Appearance.ForeColor = Color.Red
                        .Cells("ORDR_CANCEL_DATE").ToolTipText = "Order is past Cancel Date"
                    End If
                End If
            End If
        End With
    End Sub

#End Region

#Region "grdSOTPREA1"

    Private Sub grdSOTPREA1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPREA1.AfterRowUpdate

    End Sub

    Private Sub grdSOTPREA1_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTPREA1.BeforeCellUpdate
        If e.Cell.Column.Key = "ORDR_QTY_PRE_ALLO" Then
            Dim q As Int64 = Val(e.Cell.Value)
            If q > Val(e.Cell.Row.Cells("ORDR_QTY_OPEN").Value & "") Or q < 0 Then
                e.Cancel = True
            End If
        End If
    End Sub
#End Region


#End Region

    Private Sub optOrders_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optOrders.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_SO()
    End Sub

#Region "Load Style Data"

    Sub Setup_Tran()
        Dim WHSE_CODE As String = ""
        If grdICTSTATA.ActiveRow.Band.Key = "ICTSTATA" Then
            WHSE_CODE = ""
        Else
            WHSE_CODE = grdICTSTATA.ActiveRow.Cells("WHSE_CODE").Value
        End If

        If optTD.Value = "ALL" Then
            ASCMAIN1.sql = "Select ICTSTAT1.OPS_YYYYPP" & vbCrLf
        ElseIf optTD.Value = "YTD" Then
            ASCMAIN1.sql = "Select SUBSTR(ICTSTAT1.OPS_YYYYPP,1,4) OPS_YYYYPP" & vbCrLf
        Else
            ASCMAIN1.sql = "Select ICTSTAT1.OPS_YYYYPP" & vbCrLf
        End If
        ASCMAIN1.sql &= ", ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE" & vbCrLf
        If optTD.Value = "YTD" Then
            ' ASCMAIN1.sql &= ", Sum (Decode(ICTSTAT1.OPS_YYYYPP,'200201',WHSE_QTY_BEG,0)) BEG" & vbCrLf
            ASCMAIN1.sql &= ", Sum (Decode(SUBSTR(ICTSTAT1.OPS_YYYYPP,5,2),'01',WHSE_QTY_BEG,0)) BEG" & vbCrLf
        Else
            ASCMAIN1.sql &= ", Sum (WHSE_QTY_BEG) BEG" & vbCrLf
        End If
        ASCMAIN1.sql &= "" _
            & ", Sum (WHSE_QTY_SHP) SHP, Sum (WHSE_QTY_RTN) RTN, Sum (WHSE_QTY_REC) REC" & vbCrLf _
            & ", Sum (WHSE_QTY_ADJ) ADJ, Sum (WHSE_QTY_XFR) XFR, Sum (WHSE_QTY_PHY) PHY" & vbCrLf
        If optTD.Value = "ALL" Then
            ASCMAIN1.sql &= ", Sum (DECODE (ICTSTAT1.OPS_YYYYPP,'" & ASCMAIN1.CYP & "',ICTSTAT2.WHSE_QTY_ON_HAND,ICTSTAT5.WHSE_QTY_ON_HAND)) ON_HAND" & vbCrLf
        ElseIf optTD.Value = "YTD" Then
            ASCMAIN1.sql &= ", Sum (DECODE (ICTSTAT1.OPS_YYYYPP,'" & ASCMAIN1.CYP & "',ICTSTAT2.WHSE_QTY_ON_HAND,0)) ON_HAND" & vbCrLf
        Else
            ASCMAIN1.sql &= ", Sum (ICTSTAT2.WHSE_QTY_ON_HAND) ON_HAND" & vbCrLf
        End If
        ASCMAIN1.sql &= "" _
            & " from ICTSTAT1,ICTSTAT2,ICTSTAT5" & vbCrLf _
            & " where ICTSTAT1.STYLE_CODE = ICTSTAT2.STYLE_CODE (+)" & vbCrLf _
            & "   and ICTSTAT1.COLOR_CODE = ICTSTAT2.COLOR_CODE (+)" & vbCrLf _
            & "   and ICTSTAT1.WHSE_CODE = ICTSTAT2.WHSE_CODE (+)" & vbCrLf _
            & "   and ICTSTAT1.STYLE_CODE = ICTSTAT5.STYLE_CODE (+)" & vbCrLf _
            & "   and ICTSTAT1.COLOR_CODE = ICTSTAT5.COLOR_CODE (+)" & vbCrLf _
            & "   and ICTSTAT1.WHSE_CODE = ICTSTAT5.WHSE_CODE (+)" & vbCrLf _
            & "   and ICTSTAT1.OPS_YYYYPP = ICTSTAT5.OPS_YYYYPP (+)" & vbCrLf _
            & "   and ICTSTAT1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and ICTSTAT1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & IIf(WHSE_CODE = "", "", "   and ICTSTAT1.WHSE_CODE = '" & WHSE_CODE & "'") & vbCrLf
        If optTD.Value = "MTD" Then
            ASCMAIN1.sql &= "   and ICTSTAT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ElseIf optTD.Value = "YTD" Then
            ASCMAIN1.sql &= " and ICTSTAT1.OPS_YYYYPP >= '" & Mid$(ASCMAIN1.CYP, 1, 4) & "01'"
            ASCMAIN1.sql &= " and ICTSTAT1.OPS_YYYYPP <= '" & Mid$(ASCMAIN1.CYP, 1, 4) & "12'"
        ElseIf optTD.Value = "ALL" Then
        End If
        ASCMAIN1.sql &= " and ICTSTAT1.OPS_YYYYPP >= '200201'"
        If optTD.Value = "ALL" Then
            ASCMAIN1.sql &= " group by ICTSTAT1.OPS_YYYYPP, "
        ElseIf optTD.Value = "YTD" Then
            ASCMAIN1.sql &= " group by substr(ICTSTAT1.OPS_YYYYPP,1,4), "
        Else
            ASCMAIN1.sql &= " group by ICTSTAT1.OPS_YYYYPP, "
        End If
        ASCMAIN1.sql &= " ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE "
        Fill_Records("ICTSTATB", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdICTSTATB, "OPS_YYYYPP".ToLower)

        For Each rowICTSTATB As DataRow In dst.Tables("ICTSTATB").Select("")
            If optTD.Value = "YTD" Then
                rowICTSTATB.Item("LEGEND") = rowICTSTATB.Item("OPS_YYYYPP")
            Else
                Dim rowGLTPARM2 As DataRow = dst.Tables("GLTPARM2").Rows.Find(rowICTSTATB.Item("OPS_YYYYPP"))
                If rowGLTPARM2 IsNot Nothing Then
                    rowICTSTATB.Item("LEGEND") = rowGLTPARM2.Item("LEGEND")
                Else
                    rowICTSTATB.Item("LEGEND") = rowICTSTATB.Item("OPS_YYYYPP")
                End If
            End If
        Next

        grdICTSTATB.Text = "Transaction Summary for " & STYLE_CODE & "-" & COLOR_CODE & ", " & optTD.Text & IIf(WHSE_CODE = "", "", ", Whse " & WHSE_CODE)

        Show_Transaction_Details()

    End Sub

    Sub Show_Transaction_Details()

        Dim WHSE_CODE As String = ""
        If grdICTSTATA.ActiveRow.Band.Key = "ICTSTATA" Then
            WHSE_CODE = ""
        Else
            WHSE_CODE = grdICTSTATA.ActiveRow.Cells("WHSE_CODE").Value
        End If

        Dim YP As String = ASCMAIN1.CYP
        If grdICTSTATB.ActiveRow IsNot Nothing Then
            YP = grdICTSTATB.ActiveRow.Cells("OPS_YYYYPP").Value
        End If

        dst.Tables("ICTTRANX").Rows.Clear()



        ASCMAIN1.sql = "Select ICTIADJ2.OPS_YYYYPP, 'A' TRAN_TYPE, ICTIADJ2.ADJ_NO TRAN_NO, ICTIADJ2.ADJ_LNO TRAN_LNO, 0 TRAN_SLNO, 0 TRAN_XLNO" & vbCrLf _
            & ", ICTIADJ1.ADJ_SOURCE TRAN_SOURCE_DOCUMENT, ICTIADJ1.ADJ_DATE TRAN_DATE, ICTIADJ1.WHSE_CODE TRAN_WHSE_CODE" & vbCrLf _
            & ", NULL TRAN_CUST_CODE, NULL TRAN_VEND_CODE, NULL TRAN_WHSE_CODE_TO, ICTIADJ1.REASON_CODE TRAN_ADJ_REASON_CODE" & vbCrLf _
            & ", ICTIADJ2.ADJ_QTY TRAN_QTY, ICTIADJ1.ADJ_REF TRAN_REF, NULL TRAN_STATUS_UPD" & vbCrLf _
            & ", ICTIADJ1.INIT_DATE, ICTIADJ1.INIT_OPER, ICTIADJ1.RTRN_NO TRAN_NO_ORIG, DECODE(ICTIADJ1.RTRN_NO,NULL,NULL,'R') TRAN_TYPE_ORIG " & vbCrLf _
            & ", NULL TRAN_ORIGINATE" & vbCrLf _
            & " from ICTIADJ1,ICTIADJ2" & vbCrLf _
            & " where ICTIADJ2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and ICTIADJ2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTIADJ1.ADJ_NO = ICTIADJ2.ADJ_NO" & vbCrLf _
            & IIf(WHSE_CODE = "", "", "   and ICTIADJ1.WHSE_CODE = '" & WHSE_CODE & "'") & vbCrLf
        If optTD.Value = "MTD" Then
            ASCMAIN1.sql &= " and ICTIADJ2.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ElseIf optTD.Value = "YTD" Then
            ASCMAIN1.sql &= " and ICTIADJ2.OPS_YYYYPP >= '" & Mid(ASCMAIN1.CYP, 1, 4) & "01'" & vbCrLf
            ASCMAIN1.sql &= " and ICTIADJ2.OPS_YYYYPP <= '" & Mid(ASCMAIN1.CYP, 1, 4) & "12'" & vbCrLf
        ElseIf optTD.Value = "ALL" Then
            If Not chkShowAll.Checked Then
                ASCMAIN1.sql &= " and ICTIADJ2.OPS_YYYYPP = '" & YP & "'"
            End If
        End If
        Fill_Records("ICTTRANX", "", False, ASCMAIN1.sql)

        If WHSE_CODE = "" Then


            ASCMAIN1.sql = "Select ICTIXFR2.OPS_YYYYPP, 'T' TRAN_TYPE, ICTIXFR2.XFR_NO TRAN_NO, ICTIXFR2.XFR_LNO TRAN_LNO, 0 TRAN_SLNO, 0 TRAN_XLNO" & vbCrLf _
                & ", NVL(SOTINVH1.CUST_CODE,ICTIXFR1.XFR_SOURCE) TRAN_SOURCE_DOCUMENT, ICTIXFR1.XFR_DATE TRAN_DATE, ICTIXFR1.WHSE_CODE TRAN_WHSE_CODE" & vbCrLf _
                & ", NULL TRAN_CUST_CODE, NULL TRAN_VEND_CODE, ICTIXFR1.WHSE_CODE_TO TRAN_WHSE_CODE_TO, NULL TRAN_ADJ_REASON_CODE" & vbCrLf _
                & ", ICTIXFR2.XFR_QTY TRAN_QTY, ICTIXFR1.XFR_REF TRAN_REF, NULL TRAN_STATUS_UPD" & vbCrLf _
                & ", ICTIXFR1.INIT_DATE, ICTIXFR1.INIT_OPER, ICTIXFR1.CTL_NO TRAN_NO_ORIG, DECODE(ICTIXFR1.CTL_NO,NULL,NULL,'S') TRAN_TYPE_ORIG " & vbCrLf _
                & ", NULL TRAN_ORIGINATE" & vbCrLf _
                & " from ICTIXFR1,ICTIXFR2,SOTINVH1" & vbCrLf _
                & " where ICTIXFR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and ICTIXFR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                & "   and ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE (+) = 'I' and SOTINVH1.INV_NO (+) = ICTIXFR1.CTL_NO" & vbCrLf _
                & IIf(WHSE_CODE = "", "", "   and ICTIXFR1.WHSE_CODE = '" & WHSE_CODE & "'") & vbCrLf
            If optTD.Value = "MTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
            ElseIf optTD.Value = "YTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP >= '" & Mid(ASCMAIN1.CYP, 1, 4) & "01'" & vbCrLf
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP <= '" & Mid(ASCMAIN1.CYP, 1, 4) & "12'" & vbCrLf
            ElseIf optTD.Value = "ALL" Then
                If Not chkShowAll.Checked Then
                    ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & YP & "'"
                End If
            End If
            Fill_Records("ICTTRANX", "", False, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select ICTIXFR2.OPS_YYYYPP, 'T' TRAN_TYPE, ICTIXFR2.XFR_NO TRAN_NO, ICTIXFR2.XFR_LNO TRAN_LNO, 0 TRAN_SLNO, 0 TRAN_XLNO" & vbCrLf _
                & ", NVL(SOTINVH1.CUST_CODE,ICTIXFR1.XFR_SOURCE) TRAN_SOURCE_DOCUMENT, ICTIXFR1.XFR_DATE TRAN_DATE, ICTIXFR1.WHSE_CODE_TO TRAN_WHSE_CODE" & vbCrLf _
                & ", NULL TRAN_CUST_CODE, NULL TRAN_VEND_CODE, ICTIXFR1.WHSE_CODE TRAN_WHSE_CODE_TO, NULL TRAN_ADJ_REASON_CODE" & vbCrLf _
                & ", -1 * ICTIXFR2.XFR_QTY TRAN_QTY, ICTIXFR1.XFR_REF TRAN_REF, NULL TRAN_STATUS_UPD" & vbCrLf _
                & ", ICTIXFR1.INIT_DATE, ICTIXFR1.INIT_OPER, ICTIXFR1.CTL_NO TRAN_NO_ORIG, DECODE(ICTIXFR1.CTL_NO,NULL,NULL,'S') TRAN_TYPE_ORIG " & vbCrLf _
                & ", NULL TRAN_ORIGINATE" & vbCrLf _
                & " from ICTIXFR1,ICTIXFR2,SOTINVH1" & vbCrLf _
                & " where ICTIXFR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and ICTIXFR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                & "   and ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE (+) = 'I' and SOTINVH1.INV_NO (+) = ICTIXFR1.CTL_NO" & vbCrLf _
                & IIf(WHSE_CODE = "", "", "   and ICTIXFR1.WHSE_CODE = '" & WHSE_CODE & "'") & vbCrLf
            If optTD.Value = "MTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
            ElseIf optTD.Value = "YTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP >= '" & Mid(ASCMAIN1.CYP, 1, 4) & "01'" & vbCrLf
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP <= '" & Mid(ASCMAIN1.CYP, 1, 4) & "12'" & vbCrLf
            ElseIf optTD.Value = "ALL" Then
                If Not chkShowAll.Checked Then
                    ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & YP & "'"
                End If
            End If
            Fill_Records("ICTTRANX", "", False, ASCMAIN1.sql)

            ' REMMED OUT BY WJZ 01/30/2020 - CONFUSING TO SEE THE TRANSACTION WITH QTY - PROB NEED TO SEE BOTH SIDES
        End If

        If WHSE_CODE <> "" Then
            ASCMAIN1.sql = "Select ICTIXFR2.OPS_YYYYPP, 'T' TRAN_TYPE, ICTIXFR2.XFR_NO TRAN_NO, ICTIXFR2.XFR_LNO TRAN_LNO, 0 TRAN_SLNO, 0 TRAN_XLNO" & vbCrLf _
              & ", NVL(ICTIXFR1.XFR_REF,ICTIXFR1.XFR_NOTE) TRAN_SOURCE_DOCUMENT, ICTIXFR1.XFR_DATE TRAN_DATE, ICTIXFR1.WHSE_CODE TRAN_WHSE_CODE" & vbCrLf _
              & ", NULL TRAN_CUST_CODE, NULL TRAN_VEND_CODE, ICTIXFR1.WHSE_CODE_TO TRAN_WHSE_CODE_TO, NULL TRAN_XFR_REASON_CODE" & vbCrLf _
              & ", ICTIXFR2.XFR_QTY TRAN_QTY, ICTIXFR1.XFR_REF TRAN_REF, NULL TRAN_STATUS_UPD" & vbCrLf _
              & ", ICTIXFR1.INIT_DATE, ICTIXFR1.INIT_OPER, ICTIXFR1.CTL_NO TRAN_NO_ORIG, NULL TRAN_TYPE_ORIG " & vbCrLf _
              & ", NULL TRAN_ORIGINATE" & vbCrLf _
              & " from ICTIXFR1,ICTIXFR2" & vbCrLf _
              & " where ICTIXFR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
              & "   and ICTIXFR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
              & "   and ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" & vbCrLf _
              & IIf(WHSE_CODE = "", "", "   and ICTIXFR1.WHSE_CODE = '" & WHSE_CODE & "'") & vbCrLf
            If optTD.Value = "MTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
            ElseIf optTD.Value = "YTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP >= '" & Mid(ASCMAIN1.CYP, 1, 4) & "01'" & vbCrLf
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP <= '" & Mid(ASCMAIN1.CYP, 1, 4) & "12'" & vbCrLf
            ElseIf optTD.Value = "ALL" Then
                If Not chkShowAll.Checked Then
                    ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & YP & "'"
                End If
            End If
            Fill_Records("ICTTRANX", "", False, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select ICTIXFR2.OPS_YYYYPP, 'T' TRAN_TYPE, ICTIXFR2.XFR_NO TRAN_NO, ICTIXFR2.XFR_LNO TRAN_LNO, 0 TRAN_SLNO, 0 TRAN_XLNO" & vbCrLf _
                & ", NVL(ICTIXFR1.XFR_REF,ICTIXFR1.XFR_NOTE) TRAN_SOURCE_DOCUMENT, ICTIXFR1.XFR_DATE TRAN_DATE, ICTIXFR1.WHSE_CODE_TO TRAN_WHSE_CODE" & vbCrLf _
                & ", NULL TRAN_CUST_CODE, NULL TRAN_VEND_CODE, ICTIXFR1.WHSE_CODE TRAN_WHSE_CODE_TO, NULL TRAN_XFR_REASON_CODE" & vbCrLf _
                & ", -1 * ICTIXFR2.XFR_QTY TRAN_QTY, ICTIXFR1.XFR_REF TRAN_REF, NULL TRAN_STATUS_UPD" & vbCrLf _
                & ", ICTIXFR1.INIT_DATE, ICTIXFR1.INIT_OPER, ICTIXFR1.CTL_NO TRAN_NO_ORIG, NULL TRAN_TYPE_ORIG " & vbCrLf _
                & ", NULL TRAN_ORIGINATE" & vbCrLf _
                & " from ICTIXFR1,ICTIXFR2" & vbCrLf _
                & " where ICTIXFR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and ICTIXFR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                & "   and ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO" & vbCrLf _
                & IIf(WHSE_CODE = "", "", "   and ICTIXFR1.WHSE_CODE_TO = '" & WHSE_CODE & "'") & vbCrLf
            If optTD.Value = "MTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
            ElseIf optTD.Value = "YTD" Then
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP >= '" & Mid(ASCMAIN1.CYP, 1, 4) & "01'" & vbCrLf
                ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP <= '" & Mid(ASCMAIN1.CYP, 1, 4) & "12'" & vbCrLf
            ElseIf optTD.Value = "ALL" Then
                If Not chkShowAll.Checked Then
                    ASCMAIN1.sql &= " and ICTIXFR2.OPS_YYYYPP = '" & YP & "'"
                End If
            End If
            Fill_Records("ICTTRANX", "", False, ASCMAIN1.sql)
        End If



        ASCMAIN1.sql = "Select ICTIREC1.OPS_YYYYPP, 'R' TRAN_TYPE, ICTIREC1.RECEIPT_NO TRAN_NO, ICTIREC2.RECEIPT_LNO TRAN_LNO, 0 TRAN_SLNO, 0 TRAN_XLNO" & vbCrLf _
            & ", ICTIREC1.SOURCE_DOC_NO TRAN_SOURCE_DOCUMENT, ICTIREC1.RECEIPT_DATE TRAN_DATE, ICTIREC1.WHSE_CODE TRAN_WHSE_CODE" & vbCrLf _
            & ", NULL TRAN_CUST_CODE, ICTIREC1.VEND_CODE TRAN_VEND_CODE, NULL TRAN_WHSE_CODE_TO, NULL TRAN_ADJ_REASON_CODE" & vbCrLf _
            & ", ICTIREC2.QTY_REC TRAN_QTY, ICTIREC2.PO_ORDER_NO TRAN_REF, CASE WHEN ICTIREC1.REVERSED_BY_RECEIPT_NO IS NULL AND ICTIREC1.REVERSES_RECEIPT_NO IS NULL THEN 'U' ELSE 'R' END TRAN_STATUS_UPD" & vbCrLf _
            & ", INIT_DATE, INIT_OPER, ICTIREC1.REVERSES_RECEIPT_NO TRAN_NO_ORIG, 'P' TRAN_TYPE_ORIG " & vbCrLf _
            & " from ICTIREC1,ICTIREC2" & vbCrLf _
            & " where ICTIREC2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and ICTIREC2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & IIf(WHSE_CODE = "", "", "   and ICTIREC1.WHSE_CODE = '" & WHSE_CODE & "'") & vbCrLf
        If optTD.Value = "MTD" Then
            ASCMAIN1.sql &= " and ICTIREC1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ElseIf optTD.Value = "YTD" Then
            ASCMAIN1.sql &= " and ICTIREC1.OPS_YYYYPP >= '" & Mid(ASCMAIN1.CYP, 1, 4) & "01'" & vbCrLf
            ASCMAIN1.sql &= " and ICTIREC1.OPS_YYYYPP <= '" & Mid(ASCMAIN1.CYP, 1, 4) & "12'" & vbCrLf
        ElseIf optTD.Value = "ALL" Then
            If Not chkShowAll.Checked Then
                ASCMAIN1.sql &= " and ICTIREC1.OPS_YYYYPP = '" & YP & "'"
            End If
        End If
        Fill_Records("ICTTRANX", "", False, ASCMAIN1.sql)


        grdICTTRANX.DisplayLayout.Bands(0).Columns("TRAN_STATUS_UPD").Hidden = True

        ASCMAIN1.sql = "Select SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, 'S' TRAN_TYPE, SUBSTR(SOTORDR1.ORDR_GROUP_NO,5,6) TRAN_NO, 0 TRAN_LNO, 0 TRAN_SLNO, 0 TRAN_XLNO" & vbCrLf _
            & ", Min (SOTINVH1.ORDR_CUST_PO) TRAN_SOURCE_DOCUMENT, Min (SOTINVH1.INV_DATE) TRAN_DATE, SOTINVH1.WHSE_CODE TRAN_WHSE_CODE" & vbCrLf _
            & ", SOTINVH1.CUST_CODE TRAN_CUST_CODE, NULL TRAN_VEND_CODE, NULL TRAN_WHSE_CODE_TO, NULL TRAN_ADJ_REASON_CODE" & vbCrLf _
            & ", Sum (SOTINVH2.ORDR_QTY_SHIP) TRAN_QTY, Null TRAN_REF, 'U' TRAN_STATUS_UPD" & vbCrLf _
            & ", MIN (SOTINVH1.INIT_DATE) INIT_DATE, MIN (SOTINVH1.INIT_OPER) INIT_OPER, NULL TRAN_NO_ORIG, NULL TRAN_TYPE_ORIG " & vbCrLf _
            & " From SOTINVH2, SOTINVH1, SOTORDR1" & vbCrLf _
            & " where SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and SOTINVH2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and NVL(SOTINVH1.ORDR_TYPE_CODE,'?') <> 'XFR'" _
            & IIf(WHSE_CODE = "", "", "   and SOTINVH1.WHSE_CODE = '" & WHSE_CODE & "'") & vbCrLf
        If optTD.Value = "MTD" Then
            ASCMAIN1.sql &= " AND SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf
        ElseIf optTD.Value = "YTD" Then
            ASCMAIN1.sql &= " and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & Mid(ASCMAIN1.CYP, 1, 4) & "01'" & vbCrLf
            ASCMAIN1.sql &= " and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & Mid(ASCMAIN1.CYP, 1, 4) & "12'" & vbCrLf
        ElseIf optTD.Value = "ALL" Then
            ASCMAIN1.sql &= " AND SOTINVH2.ORDR_YYYYPP_UPDATED is Not Null"
            If Not chkShowAll.Checked Then
                ASCMAIN1.sql &= " and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & YP & "'"
            End If
        End If
        ASCMAIN1.sql &= "" _
            & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & " and SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & " group by SOTINVH2.ORDR_YYYYPP_UPDATED, SOTORDR1.ORDR_GROUP_NO, SOTINVH1.INV_DATE, SOTINVH1.CUST_CODE, SOTINVH1.WHSE_CODE" & vbCrLf
        Fill_Records("ICTTRANX", "", False, ASCMAIN1.sql)

        Dim RUNNING_BALANCE As Int64 = 0
        Dim SQLW As String = ""
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Dim rowICTSTAT1 As DataRow = LookUp("ICTSTAT1", New String() {"200201", STYLE_CODE, COLOR_CODE})
            If rowICTSTAT1 IsNot Nothing Then
                RUNNING_BALANCE = Val(rowICTSTAT1.Item("WHSE_QTY_BEG") & "")
            End If
            SQLW = "OPS_YYYYPP >= '200201' and " & "TRAN_STATUS_UPD = 'U'"
        End If

        For Each rowICTTRANX As DataRow In dst.Tables("ICTTRANX").Select(SQLW, "INIT_DATE")
            If rowICTTRANX.Item("TRAN_STATUS_UPD") & "" <> "R" Then
                Dim TRAN_QTY_X As Int64 = Val(rowICTTRANX.Item("TRAN_QTY_X") & "")
                RUNNING_BALANCE = RUNNING_BALANCE + TRAN_QTY_X
                rowICTTRANX.Item("RUNNING_BALANCE") = RUNNING_BALANCE
            End If
        Next
        Sort_grdColumns(grdICTTRANX, "TRAN_DATE".ToLower)

        grdICTTRANX.Text = "Transaction Details for " & STYLE_CODE & "-" & COLOR_CODE & ", " & optTD.Text & IIf(WHSE_CODE = "", "", ", Whse " & WHSE_CODE)

    End Sub
    Sub Setup_SO()
        ASCMAIN1.sql = ""
        If optOrders.Value = "0" Or optOrders.Value = "3" Or optOrders.Value = "1" Then
            ASCMAIN1.sql &= "(" & vbCrLf
        End If
        ASCMAIN1.sql &= "SELECT 'O' ORDR_TYPE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
            & ", MIN(SOTORDR1.SREP_CODE) SREP_CODE, MIN(SOTORDR1.WHSE_CODE) WHSE_CODE, SOTORDR0.ORDR_TYPE_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR, SUM (SOTORDR2.ORDR_QTY_OPEN) OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) PICK, SUM (SOTORDR2.ORDR_QTY_ALLO) ALLO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) SHIP, SUM (SOTORDR2.ORDR_QTY_CANC) CANC" & vbCrLf _
            & ", COUNT (DISTINCT SOTORDR1.ORDR_NO) ORDERS" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY      * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_DATE_RECD) ORDR_DATE_RECD, MIN (SOTORDR1.INIT_DATE) INIT_DATE" & vbCrLf _
            & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & " From SOTORDR2, SOTORDR1, SOTORDR0, ARTCUST1, ICTATOP1" & vbCrLf
        If chkSR.Checked Then
            ASCMAIN1.sql &= "" _
                & " where (SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "     or SOTORDR2.STYLE_CODE_SUB = '" & STYLE_CODE & "'" & vbCrLf _
                & "     or SOTORDR2.RANGE_STYLE_CODE = '" & STYLE_CODE & "')" & vbCrLf _
                & "   and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf
        Else
            ASCMAIN1.sql &= "" _
                & " where SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf
        End If
        If optOrders.Value = "0" Then
            ASCMAIN1.sql &= "   and (SOTORDR2.ORDR_STATUS = 'O' OR SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf
        ElseIf optOrders.Value = "2" Then
            ASCMAIN1.sql &= "   and SOTORDR2.ORDR_QTY_CANC <> 0" & vbCrLf
        ElseIf optOrders.Value = "5" Then
            ASCMAIN1.sql &= "   and SOTORDR2.ORDR_QTY_SHIP <> 0" & vbCrLf
        ElseIf optOrders.Value = "3" Then
            ASCMAIN1.sql &= "" _
                & "   and SOTORDR2.ORDR_STATUS = 'O'" & vbCrLf _
                & "   and SOTORDR2.ORDR_QTY_OPEN <> 0" & vbCrLf
        ElseIf optOrders.Value = "4" Then
            ASCMAIN1.sql &= "" _
                & "   and (SOTORDR2.ORDR_STATUS = 'O' OR SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf _
                & "   and SOTORDR2.ORDR_QTY_PICK <> 0" & vbCrLf
        End If
        ASCMAIN1.sql &= "" _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ICTATOP1.STYLE_CODE (+) = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and ICTATOP1.COLOR_CODE (+) = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTATOP1.ORDR_TYPE (+) = 'O'" & vbCrLf _
            & "   and ICTATOP1.ORDR_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE"
        If Absx1.txtFor("SREP_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and SOTORDR1.CUST_CODE in (Select CUST_CODE from ARTCUST1 where SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "')" & vbCrLf
        End If
        ASCMAIN1.sql &= "" _
            & " group by SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, ARTCUST1.CUST_NAME, SOTORDR0.ORDR_TYPE_CODE" & vbCrLf _
            & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf

        If optOrders.Value = "0" Or optOrders.Value = "3" Or optOrders.Value = "1" Then
            ASCMAIN1.sql &= ") union (" & vbCrLf _
            & "SELECT 'R' ORDR_TYPE, SOTRSRV2.RSRV_NO ORDR_GROUP_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO ORDR_CUST_PO" & vbCrLf _
            & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
            & ", MIN(SOTRSRV1.SREP_CODE) SREP_CODE, MIN(SOTRSRV1.WHSE_CODE) WHSE_CODE, NULL ORDR_TYPE_CODE" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY) ORDR, SUM (SOTRSRV2.RSRV_QTY_OPEN) OPEN" & vbCrLf _
            & ", SUM (0) PICK, SUM (SOTRSRV2.RSRV_QTY_ALLO) ALLO" & vbCrLf _
            & ", 0 SHIP, 0 CANC" & vbCrLf _
            & ", 0 ORDERS" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY_OPEN * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
            & ", SUM (0                      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
            & ", SUM (0                      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY_CANC * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & ", SOTRSRV1.INIT_DATE AS ORDR_DATE_RECD, SOTRSRV1.INIT_DATE" & vbCrLf _
            & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & " From SOTRSRV2, SOTRSRV1, ARTCUST1, ICTATOP1" & vbCrLf _
            & " where SOTRSRV2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and SOTRSRV2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
            & "   and SOTRSRV2.RSRV_QTY_OPEN <> 0" & vbCrLf _
            & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
            & "   and ICTATOP1.STYLE_CODE (+) = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and ICTATOP1.COLOR_CODE (+) = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTATOP1.ORDR_TYPE (+) = 'R'" & vbCrLf _
            & "   and ICTATOP1.ORDR_NO (+) = SOTRSRV2.RSRV_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTRSRV1.CUST_CODE" & vbCrLf
            If Absx1.txtFor("SREP_CODE").Text <> "" Then
                ASCMAIN1.sql &= "   and SOTRSRV1.CUST_CODE in (Select CUST_CODE from ARTCUST1 where SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "')" & vbCrLf
            End If
            ASCMAIN1.sql &= "" _
                & " group by SOTRSRV2.RSRV_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE, ARTCUST1.CUST_NAME, SOTRSRV1.INIT_DATE" & vbCrLf _
                & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & ")" & vbCrLf
        End If

        Fill_Records("SOTORDRX", "", True, ASCMAIN1.sql)
        'sqlSOTORDR2 = ASCMAIN1.sql

        Dim ORDR_GROUP_NO As String = ""
        Dim SUBFLAG As Boolean = False
        Dim rowSOTINVH1 As DataRow = Nothing
        Dim STYLE_COST As Decimal = 0
        Dim rowSOTINVH2 As DataRow = Nothing
        Dim extra_decimals As Boolean = False

        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select("", "ORDR_GROUP_NO")

            Dim ORDR_UNIT_PRICE_CALC As Decimal = Val(rowSOTORDRX.Item("ORDR_UNIT_PRICE_CALC") & "")
            If ORDR_UNIT_PRICE_CALC <> Val(Format(ORDR_UNIT_PRICE_CALC, "#.00")) Then
                extra_decimals = True
            End If
            If ORDR_GROUP_NO <> rowSOTORDRX.Item("ORDR_GROUP_NO") & "" Then
                SUBFLAG = False
                ORDR_GROUP_NO = rowSOTORDRX.Item("ORDR_GROUP_NO") & ""
                ASCMAIN1.sql = " SELECT DISTINCT STYLE_CODE, COLOR_CODE, STYLE_CODE_SUB" & vbCrLf _
                    & " FROM SOTORDR2, SOTORDR1" & vbCrLf _
                    & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                    & " AND SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                    & " AND SOTORDR2.STYLE_CODE_SUB = '" & STYLE_CODE & "'" & vbCrLf _
                    & " AND SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf
                Dim rowSTYLESUB As DataRow = ASCDATA1.GetDataRow
                If rowSTYLESUB IsNot Nothing AndAlso rowSTYLESUB.Item("STYLE_CODE_SUB") & "" <> "" Then
                    SUBFLAG = True
                End If

                ASCMAIN1.sql = "SELECT MAX(SOTINVH1.INV_DATE) INV_DATE" & vbCrLf _
                     & " From SOTINVH1" & vbCrLf _
                     & " where SOTINVH1.ORDR_NO IN (" & vbCrLf _
                     & " SELECT DISTINCT(ORDR_NO)" & vbCrLf _
                     & " From SOTORDR1" & vbCrLf _
                     & " WHERE SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')" & vbCrLf
                rowSOTINVH1 = ASCDATA1.GetDataRow

                If chkCOST.Checked Then
                    ASCMAIN1.sql = " SELECT MAX(SOTINVH2.ORDR_UNIT_COST) COST" & vbCrLf _
                    & " FROM SOTINVH1, SOTINVH2, SOTSHIP1" & vbCrLf _
                    & " WHERE SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                    & " AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                    & " AND SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
                    & " AND SOTINVH1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                    & " AND SOTSHIP1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                    & " AND SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf
                    rowSOTINVH2 = ASCDATA1.GetDataRow
                End If
            End If

            Dim ORDR_TYPE As String = rowSOTORDRX.Item("ORDR_TYPE")

            If rowSOTINVH1.Item("INV_DATE") & "" <> "" Then
                If ORDR_TYPE <> "R" Then
                    rowSOTORDRX.Item("INV_DATE") = rowSOTINVH1.Item(0)
                End If
            End If

            If chkCOST.Checked Then
                If rowSOTINVH2 IsNot Nothing AndAlso Val(rowSOTINVH2.Item("COST") & "") <> 0 Then
                    STYLE_COST = Val(rowSOTINVH2.Item("COST") & "")
                    rowSOTORDRX.Item("STYLE_COST") = STYLE_COST
                    rowSOTORDRX.Item("COST_SOURCE") = "I"
                Else
                    ' Dim W As String = TAC.ICCMAIN1.Calc_Cost_OH(Me, ASCMAIN1.CYP, STYLE_CODE, COLOR_CODE, False)
                    Dim W As String = TAC.ICCMAIN1.Calc_Cost_OH(Me, ASCMAIN1.CYP, STYLE_CODE, COLOR_CODE, Not dst.Tables.Contains("ICTCOST1"))
                    Dim a() As String = Split(W, "|")
                    STYLE_COST = Val(a(0))
                    rowSOTORDRX.Item("STYLE_COST") = STYLE_COST
                    rowSOTORDRX.Item("COST_SOURCE") = "H"
                End If

                If SUBFLAG = True Then
                    rowSOTORDRX.Item("SUB") = "1"
                End If
            Else
                rowSOTORDRX.Item("STYLE_COST") = 0
                rowSOTORDRX.Item("COST_SOURCE") = "X"
            End If

            For Each COLUMN_NAME As String In New String() {"OPEN", "ALLO", "PICK", "SHIP", "CANC"}
                If rowSOTORDRX.Item(COLUMN_NAME) & "" = "0" Then
                    rowSOTORDRX.Item(COLUMN_NAME) = DBNull.Value
                End If
            Next
        Next

        If extra_decimals Then
            grdSOTORDRX.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE_CALC").Format = "#,##0.0000"
        Else
            grdSOTORDRX.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE_CALC").Format = "#,##0.00"
        End If

        Dim SC As String = " for " & STYLE_CODE & "-" & COLOR_CODE
        If optOrders.Value = "0" Then
            grdSOTORDRX.Text = "Open+Pick Sales Orders & Reservations (Shown in Red)" & SC
        ElseIf optOrders.Value = "1" Then
            grdSOTORDRX.Text = "All Sales Orders Ever Taken for this Style && Reservations (Shown in Red)" & SC
        ElseIf optOrders.Value = "2" Then
            grdSOTORDRX.Text = "Sales Orders with a Non-Zero Value for Qty Cancelled" & SC
        ElseIf optOrders.Value = "3" Then
            grdSOTORDRX.Text = "Open Sales Orders & Reservations (Shown in Red)" & SC
        ElseIf optOrders.Value = "4" Then
            grdSOTORDRX.Text = "Orders in Pick" & SC
        ElseIf optOrders.Value = "5" Then
            grdSOTORDRX.Text = "Sales Orders which have been Partially or Totally Shipped" & SC
        End If

        Sort_grdColumns(grdSOTORDRX, "ORDR_CANCEL_DATE".ToLower)

        Setup_SO_Details()
    End Sub

    Sub Setup_PO()

        ' & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN " & vbCrLf _

        ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & "Select POTORDR1.INIT_DATE, POTSHIP1.WHSE_CODE, POTSHIP3.PO_ORDER_NO" & vbCrLf _
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
            & ", POTORDR2.PO_QTY_ORD, 0 PO_QTY_OPN " & vbCrLf _
            & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
            & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
            & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, ICTATOP2 " & vbCrLf _
            & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO " & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTATOP2.STYLE_CODE (+) = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and ICTATOP2.COLOR_CODE (+) = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTATOP2.PS_CODE (+) = 'S'" & vbCrLf _
            & "   and ICTATOP2.PS_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & IIf(optOpen.Value = "O", "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf, "") _
            & ") union (    " & vbCrLf _
            & "Select POTORDR1.INIT_DATE, POTORDR1.WHSE_CODE, POTORDR2.PO_ORDER_NO" & vbCrLf _
            & ", POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
            & ", POTORDR1.FACTORY_CODE, POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ", Null PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf _
            & ", Decode(nvl(POTORDR2.PO_QTY_OPN,0),0,'ClosedPO','OpenPO') PO_SHIP_VESSEL" & vbCrLf _
            & ", POTORDR2.PO_DATE_SHIP_BY, POTORDR2.PO_DATE_ETA" & vbCrLf _
            & ", " & CStr(Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR") & "")) & " PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
            & ", Null PO_SHIP_REF_NO, Null CONTAINER_NO" & vbCrLf _
            & ", NULL PO_DATE_RECEIVED" & vbCrLf _
            & ", 0 PO_QTY_SHP, 0 PO_QTY_REC" & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
            & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN " & vbCrLf _
            & ", POTORDR2.PO_DATE_ETA + " & CStr(Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR") & "")) & " PO_ARRIVAL_DATE" & vbCrLf _
            & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
            & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & " From POTORDR1, POTORDR2, ICTATOP2 " & vbCrLf _
            & " where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO " & vbCrLf _
            & " and POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & " and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTATOP2.STYLE_CODE (+) = '" & STYLE_CODE & "'" & vbCrLf _
            & "   and ICTATOP2.COLOR_CODE (+) = '" & COLOR_CODE & "'" & vbCrLf _
            & "   and ICTATOP2.PS_CODE (+) = 'P'" & vbCrLf _
            & "   and ICTATOP2.PS_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & IIf(optOpen.Value = "O", "   and POTORDR2.PO_QTY_OPN <> 0" & vbCrLf, "") _
            & ")"
        Fill_Records("POTORDRX", "", True, ASCMAIN1.sql)

        If optOpen.Value = "O" Then
            Sort_grdColumns(grdPOTORDRX, "PO_DATE_RECEIVED,PO_ARRIVAL_DATE")
        Else
            Sort_grdColumns(grdPOTORDRX, "PO_DATE_RECEIVED,PO_ARRIVAL_DATE")
        End If



        grdPOTORDRX.Text = "PO & In Transit for " & STYLE_CODE & "-" & COLOR_CODE & " (" & optOpen.Text & ")"

    End Sub

#End Region

    Private Sub grdSOTORDRY_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRY.InitializeRow

    End Sub

    Private Sub optDetails_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optDetails.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_SO_Details()
    End Sub

    Sub Setup_SO_Details()

        Dim ORDR_GROUP_NO As String = ""
        Dim CUST_CODE As String = ""
        Dim ORDR_CUST_PO As String = ""

        If grdSOTORDRX.ActiveRow IsNot Nothing Then
            ORDR_GROUP_NO = grdSOTORDRX.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
            CUST_CODE = grdSOTORDRX.ActiveRow.Cells("CUST_CODE").Value & ""
            ORDR_CUST_PO = grdSOTORDRX.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
        Else
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing View")

        If optDetails.Value = "A" Then
            splDetails.Panel2Collapsed = True
            If grdSOTORDRX.Rows.Count > 0 Then
                grdSOTORDRX.ActiveRowScrollRegion.FirstRow = grdSOTORDRX.Rows(0)
            End If

        ElseIf optDetails.Value = "O" Then
            grdSOTORDRX.ActiveRowScrollRegion.FirstRow = grdSOTORDRX.ActiveRow
            grdSOTORDRY.Text = "All Styles on Customer " & CUST_CODE & " Order " & ORDR_CUST_PO
            If grdSOTORDRX.ActiveRow.Cells("ORDR_TYPE").Value & "" = "O" Then
                ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
                    & ", Null as ORDR_NO, Null as CUST_STORE_NO, Null as ORDR_CUST_PO" & vbCrLf _
                    & ", SOTORDR2.STYLE_CODE_SUB, SOTORDR2.STYLE_DESC" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY) ORDR" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_OPEN) OPEN" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_PICK) PICK" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_ALLO) ALLO" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_SHIP) SHIP" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_CANC) CANC, COUNT (*) ORDERS" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT " & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN " & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_PICK " & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP " & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_CANC " & vbCrLf _
                    & " from SOTORDR2,SOTORDR1,ICTCOLR1" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO " & vbCrLf _
                    & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                    & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                    & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
                    & ", Null, Null, Null" & vbCrLf _
                    & ", SOTORDR2.STYLE_CODE_SUB, SOTORDR2.STYLE_DESC" & vbCrLf
            Else
                ASCMAIN1.sql = "SELECT SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
                    & ", Null as ORDR_NO, Null as CUST_STORE_NO, Null as ORDR_CUST_PO" & vbCrLf _
                    & ", Null STYLE_CODE_SUB, ICTSTYL1.STYLE_DESC" & vbCrLf _
                    & ", SUM (SOTRSRV2.RSRV_QTY) ORDR" & vbCrLf _
                    & ", SUM (SOTRSRV2.RSRV_QTY_OPEN) OPEN" & vbCrLf _
                    & ", SUM (0) PICK" & vbCrLf _
                    & ", SUM (SOTRSRV2.RSRV_QTY_ALLO) ALLO" & vbCrLf _
                    & ", SUM (0) SHIP" & vbCrLf _
                    & ", SUM (SOTRSRV2.RSRV_QTY_CANC) CANC, COUNT (*) ORDERS" & vbCrLf _
                    & ", SUM (SOTRSRV2.RSRV_QTY *      SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT " & vbCrLf _
                    & ", SUM (SOTRSRV2.RSRV_QTY_OPEN * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN " & vbCrLf _
                    & ", SUM (0                      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_PICK " & vbCrLf _
                    & ", SUM (0                      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP " & vbCrLf _
                    & ", SUM (SOTRSRV2.RSRV_QTY_CANC * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_CANC " & vbCrLf _
                    & " from SOTRSRV2,SOTRSRV1,ICTSTYL1,ICTCOLR1 " & vbCrLf _
                    & " where SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO " & vbCrLf _
                    & "   and SOTRSRV1.RSRV_NO = '" & grdSOTORDRX.ActiveRow.Cells("ORDR_GROUP_NO").Text & "'" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = SOTRSRV2.STYLE_CODE" & vbCrLf _
                    & "   and ICTCOLR1.COLOR_CODE = SOTRSRV2.COLOR_CODE" & vbCrLf _
                    & " group by SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
                    & ", Null, Null, Null" & vbCrLf _
                    & ", Null, ICTSTYL1.STYLE_DESC" & vbCrLf
            End If
            Show_Details(True)
            Fill_Records("SOTORDRY", "", True, ASCMAIN1.sql)

        ElseIf optDetails.Value = "S" Then
            grdSOTORDRX.ActiveRowScrollRegion.FirstRow = grdSOTORDRX.ActiveRow
            grdSOTORDRY.Text = "All Stores on " & CUST_CODE & " Order " & ORDR_CUST_PO
            If grdSOTORDRX.ActiveRow.Cells("ORDR_TYPE").Value = "O" Then
                ASCMAIN1.sql = "Select Null as STYLE_CODE, Null as COLOR_CODE, Null as COLOR_DESC" & vbCrLf _
                    & ", SOTORDR1.ORDR_NO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                    & ", Null as STYLE_CODE_SUB, Null as STYLE_DESC" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY ORDR" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_OPEN OPEN" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_PICK PICK" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_ALLO ALLO" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_SHIP SHIP" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_CANC CANC" & vbCrLf _
                    & ", 1 ORDERS" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY      * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_OPEN " & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_PICK " & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP " & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_CANC " & vbCrLf _
                    & " from SOTORDR2,SOTORDR1" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO " & vbCrLf _
                    & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                    & "   and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & "   and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf
                Show_Details(False)
                Fill_Records("SOTORDRY", "", True, ASCMAIN1.sql)
            Else
                MsgBox("No Store Details on Reservations", vbOKOnly, "Cannot Display")
                optDetails.Value = "A"
            End If
        End If

        If optDetails.Value <> "A" Then
            For Each rowSOTORDRY As DataRow In dst.Tables("SOTORDRY").Select("", "")

                If optDetails.Value = "O" Then 'Other styles chosen
                    Dim STYLE_COST As Decimal = 0
                    ASCMAIN1.sql = " SELECT MAX(SOTINVH2.ORDR_UNIT_COST) COST" & vbCrLf _
                        & " FROM SOTINVH1, SOTINVH2, SOTSHIP1" & vbCrLf _
                        & " WHERE SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                        & " AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                        & " AND SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
                        & " AND SOTINVH1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                        & " AND SOTSHIP1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                        & " AND SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf
                    Dim rowSOTINVH2 As DataRow = ASCDATA1.GetDataRow

                    If Val(rowSOTINVH2.Item(0) & "") <> 0 Then
                        STYLE_COST = Val(rowSOTINVH2.Item(0) & "")
                        rowSOTORDRY.Item("STYLE_COST") = STYLE_COST
                        rowSOTORDRY("COST_SOURCE") = "I"
                    Else
                        Dim W As String = TAC.ICCMAIN1.Calc_Cost_OH(Me, ASCMAIN1.CYP, STYLE_CODE, COLOR_CODE, False)
                        Dim a() As String = Split(W, "|")
                        STYLE_COST = Val(a(0))
                        rowSOTORDRY.Item("STYLE_COST") = STYLE_COST
                        rowSOTORDRY.Item("COST_SOURCE") = "H"
                    End If
                End If
            Next
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Show_Details(tf As Boolean)
        With grdSOTORDRY.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Hidden = Not tf
            .Columns("COLOR_CODE").Hidden = Not tf
            .Columns("STYLE_DESC").Hidden = Not tf
            .Columns("COLOR_DESC").Hidden = Not tf
            .Columns("STYLE_CODE_SUB").Hidden = Not tf
            .Columns("ORDR_NO").Hidden = tf
            .Columns("CUST_STORE_NO").Hidden = tf
            .Columns("ORDR_CUST_PO").Hidden = tf
        End With
        splDetails.Panel2Collapsed = False
    End Sub

    Sub Toggle_Maint()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Maintain Codes/Dates"), UltraWinToolbars.StateButtonTool)
        With grdSOTALLO1.DisplayLayout.Bands(0)
            If tlb_sbt.Checked Then
                grdSOTALLO1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                .Columns("ORDR_DEMAND_DATE").CellAppearance.BackColor = Color.LightGreen
                .Columns("ORDR_PRIORITY_DATE").CellAppearance.BackColor = Color.LightGreen
                .Columns("ORDR_PRIORITY").CellAppearance.BackColor = Color.LightGreen
                .Columns("ORDR_RELEASE").CellAppearance.BackColor = Color.LightGreen
                .Columns("ORDR_BACKORDER").CellAppearance.BackColor = Color.LightGreen
            Else
                grdSOTALLO1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                .Columns("ORDR_DEMAND_DATE").CellAppearance.BackColor = Color.Empty
                .Columns("ORDR_PRIORITY_DATE").CellAppearance.BackColor = Color.Empty
                .Columns("ORDR_PRIORITY").CellAppearance.BackColor = Color.Empty
                .Columns("ORDR_RELEASE").CellAppearance.BackColor = Color.Empty
                .Columns("ORDR_BACKORDER").CellAppearance.BackColor = Color.Empty
            End If
        End With
        Setup_tabMain()
    End Sub

    Sub Toggle_ALLOCF()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Allocation Cur/Fut/Cxl"), UltraWinToolbars.StateButtonTool)
        With grdSOTALLO1.DisplayLayout.Bands(0)
            If tlb_sbt.Checked Then
                .Columns("SD_QTY_ALLO_CUR").Hidden = False
                .Columns("SD_QTY_ALLO_FUT").Hidden = False
                .Columns("SD_QTY_ALLO_CXL").Hidden = False
            Else
                .Columns("SD_QTY_ALLO_CUR").Hidden = True
                .Columns("SD_QTY_ALLO_FUT").Hidden = True
                .Columns("SD_QTY_ALLO_CXL").Hidden = True
            End If
        End With
    End Sub

    Private Sub cmdFetchShippedOrders_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetchShippedOrders.Click
        FetchShippedOrders()
    End Sub

    Private Sub cmdFetchOpenOrders_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetchOpenOrders.Click
        FetchOpenOrders()
    End Sub

    Private Sub optASL_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optASL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If Not ScreenMode Then Exit Sub
        Allocate()
        Setup_ASL()
    End Sub

    Sub Setup_ASL()
        For Each row As DataRow In dst.Tables("SOTALLO1").Select("RECORD_TYPE = '1'")
            If optASL.Value = "1" Then
                row.Item("SD_DATE") = row.Item("ORDR_DEMAND_DATE")
            Else
                If row.Item("ORDR_RELEASE_AVAIL") & "" = "" Then
                    row.Item("SD_DATE") = row.Item("ORDR_SHIP_DATE")

                    If ASCMAIN1.CLIENT = "RGI" Then
                        row.Item("SD_DATE") = row.Item("ORDR_PRIORITY_DATE")
                    End If

                Else
                    If Format(row.Item("ORDR_RELEASE_AVAIL"), "yyyyMMdd") _
                     > Format(row.Item("ORDR_SHIP_DATE"), "yyyyMMdd") Then
                        row.Item("SD_DATE") = row.Item("ORDR_RELEASE_AVAIL")
                    Else
                        If ASCMAIN1.CLIENT = "RGI" Then
                            ' THIS MESSES UP THE SORT SEQ IN THE GRID - RGI WANTS SD_DATE TO BE INIT_DATE
                        Else
                            row.Item("SD_DATE") = row.Item("ORDR_SHIP_DATE")
                        End If

                    End If
                End If
            End If
            'row.Item("SD_DATE_X") = Format(row.Item("ORDR_DEMAND_DATE"), "MM/dd/yy")
            row.Item("SD_DATE_X") = Format(row.Item("SD_DATE"), "MM/dd/yy")
        Next
        Set_Table()
    End Sub

    Private Sub optTD_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optTD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        chkShowAll.Visible = (optTD.Value = "ALL")
        EnforceConstraints(False)
        Setup_Tran()
        EnforceConstraints(True)
    End Sub

    Private Sub chkCOST_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkCOST.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Toggle_Show_Cost()
        If chkCOST.Checked Then Setup_SO()
    End Sub

    Sub Toggle_Show_Cost()
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("STYLE_COST").Hidden = Not chkCOST.Checked
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("COST_SOURCE").Hidden = Not chkCOST.Checked Or 1 = 1 ' NOT A GOOD CONCEPT TO DISPLAY
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("GP_PCT").Hidden = Not chkCOST.Checked
        grdSOTORDRY.DisplayLayout.Bands(0).Columns("STYLE_COST").Hidden = Not chkCOST.Checked
        grdSOTORDRY.DisplayLayout.Bands(0).Columns("COST_SOURCE").Hidden = Not chkCOST.Checked Or 1 = 1 ' NOT A GOOD CONCEPT TO DISPLAY
        grdSOTORDRY.DisplayLayout.Bands(0).Columns("GP_PCT").Hidden = Not chkCOST.Checked
    End Sub

    Private Sub optOpen_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optOpen.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_PO()
    End Sub

    Private Sub chkSR_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSR.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_SO()
    End Sub

    Private Sub chkSumPO_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSumPO.CheckedChanged
        If grdSOTINVHY.Visible = True Then
            Load_SOTINVHY()
        End If
    End Sub

    Private Sub optViewStyles_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optViewStyles.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Recent()
    End Sub

    Sub Setup_Recent()
        If optViewBy.Tag & "" = "X" Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Getting " & optViewStyles.Text, "")

        If iosf = "" Then
            iosf = ASCMAIN1.Folders("Temp") & "ISI" & XNO & ".xml"
            grdICTSTYL1_Recent.DisplayLayout.Save(iosf)
        End If

        If txtSupplierOption.Tag & "" <> "V" Then
            txtSupplierOption.Text = ""
        End If

        optViewBy.Visible = (optViewStyles.Value <> "V" And optViewStyles.Value <> "A")

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            optViewBy.Visible = (optViewStyles.Value <> "V")
        End If
        chkOffset.Visible = (optViewBy.Value = "SCW") And (optViewStyles.Value = "N")

        Dim WHSE_CODEs As String = ""
        For Each row As DataRow In dst.Tables("ICTWHSES").Select("SEL = '1'")
            WHSE_CODEs &= ",'" & row.Item(0) & "'"
        Next
        WHSE_CODEs = Mid(WHSE_CODEs, 2)

        If optViewStyles.Value = "V" Then
            'Dim iosf As String = ASCMAIN1.Folders("Temp") & "ISI" & XNO & ".xml"
            'grdICTSTYL1_Recent.DisplayLayout.Save(iosf)

            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Summaries.Clear()
            grdICTSTYL1_Recent.DataSource = Nothing
            grdICTSTYL1_Recent.DataSource = dst.Tables("ICTSTYL1_RECENT")
            grdICTSTYL1_Recent.DisplayLayout.Load(iosf)
            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Summaries.Clear()
            Style_grdICTSTYL1_Recent()
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                Create_Summary(grdICTSTYL1_Recent, New String() {"OPNQTY", "OPNAMT", "SHPQTY", "SHPAMT"}, "Sum", grdICTSTYL1_Recent.DisplayLayout.Bands(0).Key) ' "ICTSTYL1_VIEW")
            End If

        Else

            Dim sqlx As String = "Select STYLE_CODE" _
                                 & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", ",COLOR_CODE", "") & vbCrLf _
                                 & IIf(optViewBy.Value = "SCW", ",WHSE_CODE", "") & vbCrLf _
                                 & ", SUM (WHSE_QTY_ON_HAND) QTY_ONHD" & vbCrLf _
                                 & ", SUM (WHSE_QTY_ON_ORDER) QTY_ONPO" & vbCrLf _
                                 & ", SUM (WHSE_QTY_TRAN) QTY_TRAN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_OPEN) QTY_OPEN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PICK) QTY_PICK" & vbCrLf _
                                 & ", SUM (WHSE_QTY_COMM) QTY_COMM" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PROD) QTY_PROD" & vbCrLf _
                                 & " from ICTSTAT2" & vbCrLf _
                                 & " where WHSE_CODE in (" & WHSE_CODEs & ")" & vbCrLf _
                                 & " group by STYLE_CODE" & vbCrLf _
                                 & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", ",COLOR_CODE", "") _
                                 & IIf(optViewBy.Value = "SCW", ",WHSE_CODE", "")




            Dim SQLW As String = ""
            If optViewStyles.Value = "S" Then ' Show Style/Colors with Status Qtys
                SQLW = " and (NVL(QTY_ONHD,0) <> 0 or NVL(QTY_ONPO,0) <> 0 or NVL(QTY_TRAN,0) <> 0 or NVL(QTY_OPEN,0) <> 0 or NVL(QTY_PICK,0) <> 0 or NVL(QTY_COMM,0) <> 0 or NVL(QTY_PROD,0) <> 0)"

            ElseIf optViewStyles.Value = "N" Then ' Show Style/Colors with Negatives
                SQLW = " and (NVL(QTY_ONHD,0) + NVL(QTY_ONPO,0) + NVL(QTY_TRAN,0) - NVL(QTY_OPEN,0) - NVL(QTY_PICK,0) - NVL(QTY_COMM,0) - NVL(QTY_PROD,0) < 0)"

            End If
            ', SYSDATE LAST_ORDR_DATE, 'X' LAST_ORDR_NO, 'X' LAST_ORDR_CUST_CODE, 'X' LAST_ORDR_PO

            ASCMAIN1.sql = "Select ICTSTYL1.*" & vbCrLf _
                & ", X.QTY_ONHD, X.QTY_ONPO, X.QTY_TRAN, X.QTY_OPEN, X.QTY_PICK, X.QTY_COMM, X.QTY_PROD" & vbCrLf _
                & ", NVL(X.QTY_ONHD,0) + NVL(X.QTY_ONPO,0) + NVL(X.QTY_TRAN,0) - NVL(X.QTY_OPEN,0) - NVL(X.QTY_PICK,0) - NVL(X.QTY_COMM,0) + NVL(X.QTY_PROD,0) QTYAVA" & vbCrLf _
                & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", ",X.COLOR_CODE,ICTCOLR1.COLOR_DESC", "") _
                & IIf(ASCMAIN1.CLIENT = "RGI", ", E.ECOM_PARTNERS", "") _
                & IIf(optViewBy.Value = "SCW", ",X.WHSE_CODE,ICTWHSE1.WHSE_DESC", "") _
                & " from ICTSTYL1,(" & sqlx & ") X" & vbCrLf _
                & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", ",ICTCOLR1", "") _
                & IIf(optViewBy.Value = "SCW", ",ICTWHSE1", "") _
                & IIf(ASCMAIN1.CLIENT = "RGI", sqlECOM, "") _
                & " where X.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", " and ICTCOLR1.COLOR_CODE (+) = X.COLOR_CODE", "") _
                & IIf(optViewBy.Value = "SCW", " and ICTWHSE1.WHSE_CODE (+) = X.WHSE_CODE", "") _
                & IIf(ASCMAIN1.CLIENT = "RGI", " and E.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE", "") _
                & SQLW

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                Dim SQLPO As String = "" _
                    & "Select STYLE_CODE" & vbCrLf _
                    & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", ", COLOR_CODE" & vbCrLf, "") _
                    & IIf(optViewBy.Value = "SCW", ", WHSE_CODE" & vbCrLf, "") _
                    & ", Sum (OPNQTY) OPNQTY, Sum (OPNAMT) OPNAMT" & vbCrLf _
                    & ", Sum (SHPQTY) SHPQTY, Sum (SHPAMT) SHPAMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
                    & ", SUM (POTORDR2.PO_QTY_OPN) OPNQTY" & vbCrLf _
                    & ", SUM (POTORDR2.PO_QTY_OPN * (1 + NVL(ICTDUTY1.DUTY_RATE,0)/100) * POTORDR2.PO_COST) OPNAMT" & vbCrLf _
                    & ", 0 SHPQTY, 0 SHPAMT" & vbCrLf _
                    & " from POTORDR2,POTORDR1,ICTDUTY1,ICTSTYL1" & vbCrLf _
                    & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and ICTDUTY1.DUTY_RATE_CODE (+) = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                    & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
                    & " union " & vbCrLf _
                    & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE" & vbCrLf _
                    & ", 0 OPNQTY, 0 OPNAMT" & vbCrLf _
                    & ", SUM (POTSHIP3.PO_QTY_SHP) SHPQTY" & vbCrLf _
                    & ", SUM (POTSHIP3.PO_QTY_SHP * CASE WHEN NVL(POTSHIP3.PO_COST_LANDED,0) <> 0 AND NVL(POTSHIP3.PO_COST_LANDED,0) <> NVL(POTSHIP3.PO_COST,0) THEN NVL(POTSHIP3.PO_COST_LANDED,0) ELSE (1 + NVL(ICTDUTY1.DUTY_RATE,0)/100) * NVL(POTSHIP3.PO_COST,0) END) SHPAMT" & vbCrLf _
                    & " from POTORDR2,POTORDR1,POTSHIP1,POTSHIP2,POTSHIP3,ICTDUTY1,ICTSTYL1" & vbCrLf _
                    & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                    & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                    & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                    & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                    & "   and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & "   and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                    & "   and ICTDUTY1.DUTY_RATE_CODE (+) = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                    & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE" & vbCrLf _
                    & ") group by STYLE_CODE" & vbCrLf _
                    & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", ", COLOR_CODE" & vbCrLf, "") _
                    & IIf(optViewBy.Value = "SCW", ", WHSE_CODE" & vbCrLf, "")

                ASCMAIN1.sql = "Select A.*, B.OPNQTY, B.OPNAMT, B.SHPQTY, B.SHPAMT" & vbCrLf _
                    & ", A.STYLE_COST STYLE_COST_LDP, CASE WHEN NVL(A.QTYAVA,0) > 0 THEN A.STYLE_COST * NVL(A.QTYAVA,0) ELSE 0 END STYLE_COST_CUM" & vbCrLf _
                    & ", CASE WHEN NVL(B.OPNQTY,0) + NVL(B.SHPQTY,0) = 0 THEN NULL ELSE TRUNC(10000 * ((NVL(B.OPNAMT,0) + NVL(B.SHPAMT,0)) / (NVL(B.OPNQTY,0) + NVL(B.SHPQTY,0)))) / 10000 END STYLE_COST_ELC" & vbCrLf _
                    & " from (" & ASCMAIN1.sql & ") A" _
                    & ", (" & SQLPO & ") B" & vbCrLf _
                    & " where B.STYLE_CODE (+) = A.STYLE_CODE" _
                    & IIf(optViewBy.Value = "SC" Or optViewBy.Value = "SCW", " and B.COLOR_CODE (+) = A.COLOR_CODE", "") _
                    & IIf(optViewBy.Value = "SCW", " and B.WHSE_CODE (+) = A.WHSE_CODE", "")
            End If

            If chkOffset.Checked Then
                ASCMAIN1.sql = "Select * from (" _
                    & ASCMAIN1.sql _
                    & ") where (STYLE_CODE, COLOR_CODE" _
                    & ", -1 * (NVL(QTY_ONHD,0) + NVL(QTY_ONPO,0) + NVL(QTY_TRAN,0)" _
                    & " - NVL(QTY_OPEN,0) - NVL(QTY_PICK,0) - NVL(QTY_COMM,0)" _
                    & " - NVL(QTY_PROD,0))) in " _
                    & "(" _
                    & "Select STYLE_CODE, COLOR_CODE" _
                    & ", NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0)" _
                    & " - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0) - NVL(WHSE_QTY_COMM,0)" _
                    & " - NVL(WHSE_QTY_PROD,0) NET_POS from ICTSTAT2 where " _
                    & "NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0)" _
                    & " - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0) - NVL(WHSE_QTY_COMM,0)" _
                    & " - NVL(WHSE_QTY_PROD,0)> 0)"

            End If

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                If txtSupplierOption.Tag & "" = "V" Then
                    ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where STYLE_CODE in ((Select Distinct STYLE_CODE from POTORDR2,POTORDR1 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO " & IIf(optSupplierOption.Value = "O", " and (QTY_ONPO <> 0 or QTY_TRAN <> 0)", "") & " and POTORDR1.VEND_CODE = '" & txtSupplierOption.Text & "')" & IIf(optSupplierOption.Value = "A", ")", " union (Select Distinct POTORDR2.STYLE_CODE from POTORDR2,POTORDR1,POTSHIP3,POTSHIP2 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO and POTSHIP2.PO_SHIP_STATUS = 'O' and POTORDR1.VEND_CODE = '" & txtSupplierOption.Text & "'))")
                End If
            End If

            ICTSTYL1_Recent = ""
            If ICTSTYL1_Recent = "" Then
                ICTSTYL1_Recent = ASCMAIN1.Temp_Table
                ASCDATA1.ExecuteSQL("Create Index I_" & ICTSTYL1_Recent & "_1 on " & ICTSTYL1_Recent & " (STYLE_CODE)")
            Else
                ASCDATA1.ExecuteSQL("Truncate Table " & ICTSTYL1_Recent)
                ASCDATA1.ExecuteSQL("Insert into " & ICTSTYL1_Recent & " " & ASCMAIN1.sql)
            End If

            'Fill_Records("ICTSTYL1_VIEW", "", True, ASCMAIN1.sql)
            Fill_Records("ICTSTYL1_VIEW", "", True, "Select * from " & ICTSTYL1_Recent)

            With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
                .Columns("LAST_ORDR_DATE").Hidden = Not chkShowDates.Checked
                .Columns("LAST_ORDR_NO").Hidden = Not chkShowDates.Checked
                .Columns("LAST_ORDR_CUST_CODE").Hidden = Not chkShowDates.Checked
                .Columns("LAST_ORDR_CUST_PO").Hidden = Not chkShowDates.Checked
            End With
            If chkShowDates.Checked Then
                Dim sql As String = "Select STYLE_CODE, MAX (ORDR_NO) ORDR_NO" & vbCrLf _
                                    & " from SOTORDR2 " & vbCrLf _
                                    & " where STYLE_CODE in " & vbCrLf _
                                    & " (Select Distinct STYLE_CODE from " & ICTSTYL1_Recent & ")" & vbCrLf _
                                    & " and ORDR_STATUS in ('O','P','F')" & vbCrLf _
                                    & " group by STYLE_CODE"
                sql = "Select LAST_ORDR.*, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                    & " from SOTORDR1, (" & sql & ") LAST_ORDR where SOTORDR1.ORDR_NO = LAST_ORDR.ORDR_NO"
                For Each row As DataRow In ASCDATA1.GetDataTable(sql).Select("")
                    Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                    Dim DATE_LAST_SO As Date = row.Item("ORDR_DATE")
                    For Each row2 As DataRow In dst.Tables("ICTSTYL1_VIEW").Select("STYLE_CODE = '" & STYLE_CODE & "'")
                        row2.Item("LAST_ORDR_DATE") = row.Item("ORDR_DATE")
                        row2.Item("LAST_ORDR_NO") = row.Item("ORDR_NO")
                        row2.Item("LAST_ORDR_CUST_CODE") = row.Item("CUST_CODE")
                        row2.Item("LAST_ORDR_CUST_PO") = row.Item("ORDR_CUST_PO")
                    Next
                Next
            End If

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                For Each row As DataRow In dst.Tables("ICTSTYL1_VIEW").Select("OPNQTY <> 0 OR SHPQTY <> 0")
                    Dim STYLE_COST_LDP As Decimal = Val(row.Item("STYLE_COST_LDP") & "")
                    Dim STYLE_COST_ELC As Decimal = Val(row.Item("STYLE_COST_ELC") & "")
                    If STYLE_COST_LDP = 0 And STYLE_COST_ELC <> 0 Then
                        STYLE_COST_LDP = STYLE_COST_ELC
                        row.Item("STYLE_COST_LDP_CODE") = "E"
                    End If
                    Dim OH As Int64 = Val(row.Item("QTY_ONHD") & "")
                    Dim SUP As Int64 = Val(row.Item("QTY_ONPO") & "") + Val(row.Item("QTY_TRAN") & "")
                    Dim DEM As Int64 = Val(row.Item("QTY_OPEN") & "") + Val(row.Item("QTY_PICK") & "")

                    If SUP < 0 Then SUP = 0
                    If DEM < 0 Then DEM = 0

                    If DEM > 0 Then
                        If OH > 0 Then
                            If OH > DEM Then
                                OH = OH - DEM
                                DEM = 0
                            Else
                                DEM = DEM - OH
                                OH = 0
                            End If
                        End If
                    End If

                    If SUP > 0 Then
                        If DEM > 0 Then
                            If DEM > SUP Then
                                DEM = DEM - SUP
                                SUP = 0
                            Else
                                SUP = SUP - DEM
                                DEM = 0
                            End If
                        End If
                    End If

                    row.Item("STYLE_COST_CUM") = OH * STYLE_COST_LDP + SUP * STYLE_COST_ELC
                Next
            End If

            'Dim iosf As String = ASCMAIN1.Folders("Temp") & "ISI" & XNO & ".xml"
            'grdICTSTYL1_Recent.DisplayLayout.Save(iosf)
            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Summaries.Clear()

            grdICTSTYL1_Recent.DataSource = Nothing
            grdICTSTYL1_Recent.DataSource = dst.Tables("ICTSTYL1_VIEW")
            grdICTSTYL1_Recent.DisplayLayout.Load(iosf)
            'ASCMAIN1.grdInitializeLayout(grdICTSTYL1_Recent, Me)
            Style_grdICTSTYL1_Recent()
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                Create_Summary(grdICTSTYL1_Recent, New String() {"OPNQTY", "OPNAMT", "SHPQTY", "SHPAMT"}, "Sum", grdICTSTYL1_Recent.DisplayLayout.Bands(0).Key) ' "ICTSTYL1_VIEW")
            End If

        End If
        grdICTSTYL1_Recent.Text = optViewStyles.Text
        If optViewStyles.Value = "S" Or optViewStyles.Value = "N" Then
            grdICTSTYL1_Recent.Text &= " Whses:" & Replace(WHSE_CODEs, "'", "")
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            For Each row As DataRow In dst.Tables("ICTSTYL1_VIEW").Select()
                row.Item("STYLE_COST_FIRST") = GetFirstCost(row.Item("STYLE_CODE").ToString)
            Next
            grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns("STYLE_COST_FIRST").Hidden = False
        End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            If txtSupplierOption.Tag & "" = "V" Then
                grdICTSTYL1_Recent.Text &= " - for Supplier " & txtSupplierOption.Text & ": " & optSupplierOption.Text
                txtSupplierOption.Tag = ""
            End If
        End If


        Sort_grdColumns(grdICTSTYL1_Recent, "STYLE_CODE")

        For Each COLUMN_NAME As String In New String() {"QTY_ONHD", "QTY_ONPO", "QTY_TRAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_PROD", "QTY_NETA", "STYLE_COST_EXT"}
            If COLUMN_NAME = "STYLE_COST_EXT" And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
            Else
                grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = (optViewStyles.Value = "V")
                If COLUMN_NAME = "QTY_COMM" Or COLUMN_NAME = "QTY_PROD" Then
                    grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
                End If
            End If
        Next

        With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
            .Columns("COLOR_CODE").Hidden = (optViewStyles.Value = "V") Or Not (optViewBy.Value = "SC" Or optViewBy.Value = "SCW")
            .Columns("COLOR_DESC").Hidden = (optViewStyles.Value = "V") Or Not (optViewBy.Value = "SC" Or optViewBy.Value = "SCW")
            .Columns("WHSE_CODE").Hidden = (optViewStyles.Value = "V") Or Not (optViewBy.Value = "SCW")
            .Columns("WHSE_DESC").Hidden = (optViewStyles.Value = "V") Or Not (optViewBy.Value = "SCW")
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdICTSTYL1_Recent_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYL1_Recent.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("QTY_NETA").Value >= 0 Then
                e.Row.Cells("QTY_NETA").Appearance.ForeColor = Color.Green
            Else
                e.Row.Cells("QTY_NETA").Appearance.ForeColor = Color.Red
            End If
        End If
    End Sub

    Private Sub imgSTYLE_DoubleClick(sender As Object, e As System.EventArgs) Handles imgSTYLE.DoubleClick
        Using F As New ASFMSGBF
            F.Show_img(imgSTYLE.Image, Me, "Style " & STYLE_CODE & ":" & Absx1.txtFor("STYLE_DESC").Text)
        End Using
    End Sub

    Private Sub cmdAllocate_Click(sender As System.Object, e As System.EventArgs) Handles cmdAllocate.Click
        Allocate()
    End Sub

    Private Sub grdPOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDRX.AfterRowActivate
        Setup_POTSHIP7()
    End Sub

    Private Sub grdPOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRX.InitializeRow
        If Trim(e.Row.Cells("PO_SHIPMENT_NO").Value & "") <> "" Then
            e.Row.Appearance.ForeColor = Color.Blue
        Else
            If Val(e.Row.Cells("PO_QTY_OPN").Value & "") <> 0 Then
                e.Row.Appearance.ForeColor = Color.Green
            Else
                e.Row.Appearance.ForeColor = Color.DarkOrange
            End If
        End If

        If e.Row.Cells("STYLE_AT_ONCE_UNTIL").Value & "" <> "" Then

            If Format(e.Row.Cells("STYLE_AT_ONCE_UNTIL").Value, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Red
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").ToolTipText = "This At-Once Parameter has expired"
            Else
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Empty
            End If

            'If e.Row.Cells("PS_ETA_NOW").Value & "" <> "" AndAlso Format(e.Row.Cells("PS_ETA_NOW").Value, "yyyyMMdd") <> Format(e.Row.Cells("PS_ETA").Value, "yyyyMMdd") Then
            '    e.Row.Cells("PS_ETA_NOW").Appearance.BackColor = Color.Yellow

            '    e.Row.Cells("PS_ETA_NOW").ToolTipText = "ETA has changes since the original Parameter record was created"
            'Else
            '    e.Row.Cells("PS_ETA_NOW").Appearance.BackColor = Color.Empty
            'End If

            If e.Row.Cells("STYLE_AT_ONCE_ACTIVE").Value & "" = "1" Then
                e.Row.Cells("STYLE_AT_ONCE_ACTIVE").Appearance.BackColor = Color.Empty
            Else
                e.Row.Cells("STYLE_AT_ONCE_ACTIVE").Appearance.BackColor = Color.Red
            End If
        End If

    End Sub

    Private Sub optViewBy_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optViewBy.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Recent()
    End Sub

    Private Sub chkOffset_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkOffset.CheckedChanged
        Setup_Recent()
    End Sub

    Private Sub grdICTSTDQ1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTDQ1.InitializeRow
        If Format(e.Row.Cells("STATUS_DATE").Value, "yyyyMMdd") = Format(Now.Date, "yyyyMMdd") Then
            e.Row.ToolTipText = "Today"
            e.Row.Appearance.BackColor = Color.Yellow
        ElseIf Format(e.Row.Cells("STATUS_DATE").Value, "yyyyMMdd") < Format(Now.Date, "yyyyMMdd") Then
            e.Row.ToolTipText = "Supply is Past Due"
            e.Row.Appearance.ForeColor = Color.Red
        End If
    End Sub

    Private Sub cbeWHSE_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cbeWHSE_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_WHSE_CODE()
    End Sub

    Sub Setup_POTSHIP7()
        If splPOTORDRX.Panel2Collapsed Then Exit Sub
        If grdPOTORDRX.ActiveRow Is Nothing OrElse Not grdPOTORDRX.ActiveRow.IsDataRow Then
            splPOTSHIP7.Visible = False
        Else
            splPOTSHIP7.Visible = True
            Dim PO_SHIPMENT_NO As String = grdPOTORDRX.ActiveRow.Cells("PO_SHIPMENT_NO").Value & ""
            Dim PO_SHIPMENT_LNO As Int32 = Val(grdPOTORDRX.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            Dim sqlw As String = "Select PO_SHIPMENT_NO, PO_SHIPMENT_LNO, CARTON_NO from POTSHIP8 " _
                                 & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
                                 & "   and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) _
                                 & "   and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            EnforceConstraints(False)
            ASCMAIN1.sql = "Select * from POTSHIP7 where (PO_SHIPMENT_NO, PO_SHIPMENT_LNO, CARTON_NO) in (" & sqlw & ")"
            Fill_Records("POTSHIP7", "", True, ASCMAIN1.sql)
            ASCMAIN1.sql = "Select * from POTSHIP8 where (PO_SHIPMENT_NO, PO_SHIPMENT_LNO, CARTON_NO) in (" & sqlw & ")"
            Fill_Records("POTSHIP8", "", True, ASCMAIN1.sql)
            EnforceConstraints(True)

            Sort_grdColumns(grdPOTSHIP7, "CARTON_NO")

        End If
    End Sub


    Private Sub grdPOTSHIP7_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIP7.AfterRowActivate
        Setup_grdPOTSHIP8()
    End Sub

    Sub Setup_grdPOTSHIP8()
        If grdPOTSHIP7.ActiveRow Is Nothing OrElse Not grdPOTSHIP7.ActiveRow.IsDataRow Then
            grdPOTSHIP8.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdPOTSHIP8.DataSource, DataTable).DefaultView
            Dim CARTON_NO As Integer = Val(grdPOTSHIP7.ActiveRow.Cells("CARTON_NO").Value & "")
            Dim PO_SHIPMENT_NO As String = grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_NO").Value & ""
            Dim PO_SHIPMENT_LNO As Integer = Val(grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            dvw.RowFilter = ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO))
            grdPOTSHIP8.Text = "Carton Configuration by Style/Color for Carton Type " & CStr(CARTON_NO)
            grdPOTSHIP8.Visible = True
        End If
    End Sub


    'Function Create_Invoice(ByVal INV_NO As String) As String
    '    Me.Cursor = Cursors.WaitCursor

    '    ASCMAIN1.Progress("Now Preparing Invoice for Printing")

    '    If Not ROWs.ContainsKey("ARTPARM1") Then
    '        Get_PARM("ARTPARM1")
    '    End If

    '    Dim REPORT_NAME As String = "SORINVP1"
    '    Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
    '    If RPT = "" Then RPT = REPORT_NAME

    '    If Not REPORTS.ContainsKey(REPORT_NAME) Then
    '        REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
    '        REPORTS(REPORT_NAME).Prepare_dst(False, "")
    '    End If

    '    Dim tempFileName As String = INV_NO
    '    Dim sql As String = " and SOTINVH1.INV_TYPE = 'I'"
    '    If InStr(INV_NO, ",") = 0 Then
    '        sql &= " and SOTINVH1.INV_NO = '" & INV_NO & "'"
    '    Else
    '        sql &= " and SOTINVH1.INV_NO in ('" & INV_NO & "')"
    '        tempFileName = Split(INV_NO, ",")(0) '  & DateTime.Now.ToString("yyyyMMddHHmmss")
    '    End If

    '    REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
    '    Dim FILENAME As String = ""
    '    With REPORTS(REPORT_NAME).clsASCBASE1
    '        .Print_Report_Begin()
    '        .CR_params.Add("SUBT", "")
    '        .CR_params.Add("CONS_INV", "")
    '        Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
    '        FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
    '        .Print_Report_End(, True)
    '    End With

    '    Me.Cursor = Cursors.Default
    '    ASCMAIN1.Progress("")

    '    Return FILENAME
    'End Function

    Sub Price_and_Availability(
        STYLE_CODE As String,
        STYLE_CLASS_CODE As String,
        COLOR_CODE As String,
        CARTON_PACK_QTY As Int64,
        STYLE_PRICE As Decimal)

        grdWHTLOCB1.Text = "Location Status "
        dst.Tables("WHTLOCB1").Rows.Clear()
        grdWHTINSTX.Text = "Open Waves"
        dst.Tables("WHTINSTX").Rows.Clear()

        ASCMAIN1.sql = "Select * from ICTSTDQ2 " _
            & " where STYLE_CODE = '" & STYLE_CODE & "'" _
            & IIf(COLOR_CODE = "", "", " and COLOR_CODE = '" & COLOR_CODE & "'")
        Fill_Records("ICTSTDQ2", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdICTSTDQ2, "WHSE_CODE,COLOR_CODE")
        grdICTSTDQ2.Text = "Availability for Style " & STYLE_CODE

        grdICTPRICX.Text = "Price List for Style " & STYLE_CODE & IIf(COLOR_CODE = "", "", ", Color " & COLOR_CODE)

        dst.Tables("ICTPRICX").Rows.Clear()

        Dim STYLE_STATUS As String = Absx1.optFor("STYLE_STATUS").Value & ""
        Dim rowICTSTATA As DataRow = dst.Tables("ICTSTATA").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
        Dim STYLE_COLOR_STATUS As String = ""
        If rowICTSTATA IsNot Nothing Then
            STYLE_COLOR_STATUS = rowICTSTATA.Item("STYLE_COLOR_STATUS") & ""
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
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
                            Dim PRICE As Decimal = STYLE_PRICE * (100 - Val(rowICTDISC1.Item("DISC" & CStr(I) & "_PCT"))) / 100
                            If STYLE_STATUS = "D" Or STYLE_COLOR_STATUS = "D" Then
                                PRICE = 0.3 * STYLE_PRICE
                            End If
                            .Item("PRICE") = PRICE
                        End With
                        dst.Tables("ICTPRICX").Rows.Add(rowICTPRICX)
                    Next
                End If
            End If

        End If

    End Sub


    Private Sub chkSingleColor_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSingleColor.CheckedChanged
        ' grdSOTINVHY.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = chkSingleColor.Checked
    End Sub

    Private Sub btnTest_Click(sender As System.Object, e As System.EventArgs) Handles btnTest.Click

        'TAC.TACMAIN1.Rename_Image_Files(Me)


        'Exit Sub


        'Stop
        'If 1 = 1 Then
        '    Convert_Database("DB", _
        '                     New String() {"Styles", "StyleDetail", "Type"}, _
        '                     New String() {"DB_STYLES", "DB_STYLEDETAIL", "DB_TYPE"})
        '    Convert_Database("PO", _
        '         New String() {"PODetail", "POStyleDetail", "POType"}, _
        '         New String() {"DB_PODETAIL", "DB_POSTYLEDETAIL", "DB_POTYPE"})
        'Else
        '    Catalog_Images_Main()
        'End If
        '' C:\Users\wjz\Desktop\Database

        'Stop

        'Dim IMAGES_FOLDER As String = "S:\Images\"
        'Dim OK As Integer = 0
        'Dim NOT_OK As Integer = 0
        'Dim FOUND As Integer = 0

        'Dim FOUND_IMAGES As New List(Of String)

        'ASCMAIN1.sql = "Select STYLE_CODE_PLM, SALES_DIVISION_CODE from ICTPLIN2"
        'Dim tbl As DataTable = ASCDATA1.GetDataTable
        'For Each row As DataRow In tbl.Select("")
        '    Dim STYLE_CODE_PLM As String = row.Item("STYLE_CODE_PLM") & ""
        '    Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE") & ""
        '    Dim FILENAME As String = IMAGES_FOLDER & SALES_DIVISION_CODE & "\" & STYLE_CODE_PLM & ".JPG"
        '    If My.Computer.FileSystem.FileExists(FILENAME) Then
        '        OK += 1
        '    Else
        '        NOT_OK += 1
        '        If STYLE_CODE_PLM.Contains("/") Then
        '        Else
        '            For Each FILENAME In My.Computer.FileSystem.GetFiles _
        '         ("C:\Users\wjz\Desktop\NYAG\Data\Database\Images", FileIO.SearchOption.SearchAllSubDirectories, STYLE_CODE_PLM & ".JPG")
        '                FOUND += 1
        '                FOUND_IMAGES.Add(FILENAME)
        '                My.Computer.FileSystem.CopyFile(FILENAME, "C:\Users\wjz\Desktop\NYAG\Orphans\" & SALES_DIVISION_CODE & "\" & STYLE_CODE_PLM & ".JPG")
        '                Exit For
        '            Next
        '        End If
        '    End If
        'Next


        'Stop

        'Dim start_writing As Boolean = False
        'Using SR As New System.IO.StreamReader("C:\Users\wjz\Desktop\VCS.TXT")
        '    Using SW As New System.IO.StreamWriter("C:\Users\wjz\Desktop\VCS2.TXT")
        '        Dim D As String = SR.ReadToEnd
        '        For Each LINE As String In D.Split(vbCrLf)
        '            Dim NEWLINE As String = Trim(Mid(LINE, 15))
        '            If NEWLINE.StartsWith("ISA") Then start_writing = True
        '            If start_writing And NEWLINE <> "" Then
        '                If Mid(NEWLINE, 1, 1) >= "A" And Mid(NEWLINE, 1, 1) <= "Z" _
        '                    And Not NEWLINE.StartsWith("Print selected") Then
        '                    SW.WriteLine(NEWLINE)
        '                End If
        '            End If
        '        Next
        '    End Using
        'End Using
        'MsgBox("Done")

    End Sub

    Sub Convert_Database(CONV As String, TBLF() As String, TBLT() As String)
        Dim TT As New Dictionary(Of String, String)

        For Each dbt As String In TBLT
            ASCMAIN1.sql = "Select * from " & dbt & " where ROWNUM < 1"
            TT.Add(dbt, ASCMAIN1.Temp_Table)
            Create_TDA(dst.Tables.Add(dbt), TT(dbt), "*")
        Next

        Dim DBP As String = "C:\Users\wjz\Desktop\datafinal\" & CONV & "\"
        For Each FI As String In My.Computer.FileSystem.GetFiles(DBP)

            Dim DB As String = Mid(FI, Len(DBP) + 1)
            Dim cs As String = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & DBP & DB & ";Uid=Admin;Pwd=;"

            Using c As New System.Data.Odbc.OdbcConnection(cs)
                Dim tc As Integer = -1
                For Each TABLE_NAME As String In TBLF
                    tc += 1
                    Dim sql As String = "Select * from " & TABLE_NAME & " where 1<>1"
                    Using cda As New System.Data.Odbc.OdbcDataAdapter(sql, c)
                        Dim tbl As New DataTable
                        cda.Fill(tbl)
                        Dim cols As String = ""
                        For I As Integer = 0 To tbl.Columns.Count - 1
                            Dim colname As String = tbl.Columns(I).ColumnName
                            If colname <> "Picture" Then
                                cols &= ",[" & colname & "]"
                            End If
                        Next
                        sql = "Select " & Mid(cols, 2) & " from " & TABLE_NAME
                    End Using

                    Using cda As New System.Data.Odbc.OdbcDataAdapter(sql, c)
                        Dim tbl As New DataTable
                        cda.Fill(tbl)

                        Dim DBT As String = TBLT(tc)

                        Dim SC As New Dictionary(Of String, Integer)
                        For I As Integer = 0 To dst.Tables(DBT).Columns.Count - 1
                            SC.Add(dst.Tables(DBT).Columns(I).ColumnName, I)
                        Next

                        For Each row As DataRow In tbl.Select("")
                            Dim row2 As DataRow = dst.Tables(DBT).NewRow
                            row2.Item("DB") = DB
                            If DBT = "DB_STYLES" Then row2.Item("PODATE") = row.Item("Date")
                            If tc = 2 Then row2.Item("XORDER") = row.Item("Order")
                            Dim PFX As String = ""
                            If tc = 1 Then PFX = "X"
                            For i As Integer = 0 To tbl.Columns.Count - 1
                                Dim dc As String = tbl.Columns(i).ColumnName
                                If SC.ContainsKey(PFX & dc.ToUpper) Then
                                    Dim J As Integer = SC(PFX & dc.ToUpper)
                                    row2.Item(J) = row.Item(i)
                                End If
                            Next
                            If tc = 0 Then
                                If Len(row2.Item("NOTESREMINDERS") & "") > 2000 Then
                                    row2.Item("NOTESREMINDERS") = Mid(row2.Item("NOTESREMINDERS"), 1, 2000)
                                End If

                            End If
                            dst.Tables(DBT).Rows.Add(row2)
                        Next

                    End Using
                Next
            End Using
        Next

        For Each dbt As String In TBLT
            Update_Record_TDA(dbt)
        Next

        Stop
    End Sub

    Sub Catalog_Images_Main()
        ASCMAIN1.sql = "Select IMAGE_NAME STYLE_CODE, IMAGE_NAME, IMAGE_NAME IMAGE_FILE, LAST_DATE from ICTSTYL1 where ROWNUM < 1"
        Dim tt As String = ASCMAIN1.Temp_Table

        Create_TDA(dst.Tables.Add("IMAGES"), tt, "*")

        Dim foldername As String = "C:\Users\wjz\Desktop\Data\Database\Images\"
        Catalog_Images(foldername)


        '********************************* INSERTED TO USE PLM STYLE MASTER
        Stop
        For Each row As DataRow In dst.Tables("IMAGES").Select("")
            Dim IMAGE_NAME As String = row.Item("IMAGE_NAME")
            Dim IMAGE_FILE As String = row.Item("IMAGE_FILE")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE)
            If rowICTPLIN2 IsNot Nothing Then
                Dim SALES_DIVISION_CODE As String = rowICTPLIN2.Item("SALES_DIVISION_CODE") & ""
                If SALES_DIVISION_CODE = "" Then Stop
                Debug.Print(STYLE_CODE)
                My.Computer.FileSystem.CopyFile(IMAGE_NAME, foldername & "\DB\" & SALES_DIVISION_CODE & "\" & IMAGE_FILE, True)
            End If
        Next
        Stop
        '*********************************



        Update_Record_TDA("IMAGES")

        Stop

        ASCDATA1.ExecuteSQL("Drop table NYA_IMAGES", True)
        ASCDATA1.ExecuteSQL("Drop table NYA_IMAGES2", True)
        ASCDATA1.ExecuteSQL("Rename " & tt & " to NYA_IMAGES")
        ASCDATA1.ExecuteSQL("UPDATE NYA_IMAGES SET IMAGE_NAME = SUBSTR(IMAGE_NAME,LENGTH('" & foldername & "')+1)")

        Dim sql As String = "" _
                            & "CREATE TABLE NYA_IMAGES2 AS" & vbCrLf _
                            & "SELECT STYLE_CODE, MAX(IMAGE_NAME) IMAGE_NAME, LAST_DATE FROM NYA_IMAGES" & vbCrLf _
                            & "WHERE (STYLE_CODE, LAST_DATE) IN (SELECT STYLE_CODE, MAX(LAST_DATE) LAST_DATE FROM NYA_IMAGES GROUP BY STYLE_CODE)" & vbCrLf _
                            & "GROUP BY STYLE_CODE, LAST_DATE"

        ASCDATA1.ExecuteSQL(sql)
        ASCDATA1.ExecuteSQL("ALTER TABLE NYA_IMAGES2 ADD STYLE_COUNT NUMBER (4,0)")
        ASCDATA1.ExecuteSQL("ALTER TABLE NYA_IMAGES2 drop column IMAGE_FILE ")
        ASCDATA1.ExecuteSQL("ALTER TABLE NYA_IMAGES2 ADD IMAGE_FILE VARCHAR2(60)")
        ASCDATA1.ExecuteSQL("UPDATE NYA_IMAGES2 SET IMAGE_FILE = (SELECT MIN(IMAGE_FILE) FROM NYA_IMAGES WHERE IMAGE_NAME = NYA_IMAGES2.IMAGE_NAME)")
        ASCDATA1.ExecuteSQL("UPDATE NYA_IMAGES2 SET STYLE_COUNT = (SELECT COUNT (*) FROM ICTSTYL1 WHERE STYLE_CODE LIKE NYA_IMAGES2.STYLE_CODE || '%')")
        ASCDATA1.ExecuteSQL("DELETE FROM NYA_IMAGES2 WHERE STYLE_COUNT = 0 OR STYLE_COUNT >= 49")
        ASCDATA1.ExecuteSQL("ALTER TABLE NYA_IMAGES2 ADD IMAGE_NAME_NEW VARCHAR2(200)")
        ASCDATA1.ExecuteSQL("ALTER TABLE NYA_IMAGES2 ADD SALES_DIVISION_CODE VARCHAR2(200)")
        ASCDATA1.ExecuteSQL("CREATE INDEX I_NYA_IMAGES2_1 ON NYA_IMAGES (IMAGE_NAME)")
        ASCDATA1.ExecuteSQL("UPDATE ICTSTYL1 SET IMAGE_NAME = NULL")

        sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS SELECT * FROM NYA_IMAGES2 ORDER BY LENGTH(STYLE_CODE);" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE ICTSTYL1 SET IMAGE_NAME = R1.IMAGE_NAME " & vbCrLf _
            & " WHERE STYLE_CODE LIKE R1.STYLE_CODE || '%' AND IMAGE_NAME IS NULL;" & vbCrLf _
            & "IF LENGTH(R1.STYLE_CODE) >=5 THEN" & vbCrLf _
            & "UPDATE ICTSTYL1 SET IMAGE_NAME = R1.IMAGE_NAME " & vbCrLf _
            & " WHERE STYLE_CODE LIKE '%' || R1.STYLE_CODE || '%' AND IMAGE_NAME IS NULL;" & vbCrLf _
            & "END IF;" & vbCrLf _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

        ASCDATA1.ExecuteSQL("UPDATE NYA_IMAGES2 SET SALES_DIVISION_CODE= (SELECT SALES_DIVISION_CODE FROM ICTSTYL1 WHERE STYLE_CODE = NYA_IMAGES2.STYLE_CODE)")

        Stop

        ASCMAIN1.sql = "Select DISTINCT IMAGE_NAME, IMAGE_FILE, SALES_DIVISION_CODE from NYA_IMAGES2 WHERE SALES_DIVISION_CODE IS NOT NULL"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim IMAGE_NAME As String = row.Item("IMAGE_NAME")
            Dim IMAGE_FILE As String = row.Item("IMAGE_FILE")
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE") & ""
            If SALES_DIVISION_CODE = "" Then Stop
            My.Computer.FileSystem.CopyFile(foldername & "\" & IMAGE_NAME, foldername & "\" & SALES_DIVISION_CODE & "\" & IMAGE_FILE, True)
            '    Stop
        Next
        Stop
    End Sub

    Sub Catalog_Images(folder As String)
        For Each filename As String In My.Computer.FileSystem.GetFiles(folder)
            'Stop
            If filename.EndsWith(".jpg") Then
                Dim f() As String = Split(filename, "\")
                Dim fs As String = f(f.Length - 1)
                If fs.StartsWith("._") Then
                Else
                    Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(filename)
                    Dim SHORTNAME As String = fi.Name
                    dst.Tables("IMAGES").Rows.Add(New String() {Replace(fs, ".jpg", ""), filename, SHORTNAME, fi.LastWriteTime})
                End If
            End If
        Next
        For Each foldername As String In My.Computer.FileSystem.GetDirectories(folder)
            Catalog_Images(foldername)
        Next
    End Sub

    Function Add_to_Quote(STYLE_CODE_PLM As String) As String

        If QUOTE_NO = "" Then
            Click_Command("New Quote Sheet")
        End If
        STYLE_CODE_PLM = STYLE_CODE_PLM.ToUpper
        Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE_PLM})
        If rowICTQUOT2 Is Nothing Then
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_PLM)
            If rowICTSTYL1 Is Nothing Then Return ""
            rowICTQUOT2 = dst.Tables("ICTQUOT2").NewRow()

            'Dim AVAILABILITY As String = ""
            'ASCMAIN1.sql = "Select WHSE_CODE, STATUS_DATE, STATUS_QTY from ICTSTDQ1 where STYLE_CODE = :PARM1"
            'For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {STYLE_CODE}).Select("", "WHSE_CODE,STATUS_DATE")
            '    AVAILABILITY &= ";" & row.Item("WHSE_CODE") & " " & row.Item("STATUS_DATE") & " " & row.Item("STATUS_QTY")
            'Next

            With rowICTQUOT2
                .Item("QUOTE_NO") = QUOTE_NO
                .Item("STYLE_CODE_PLM") = STYLE_CODE_PLM
                .Item("STYLE_CODE_CUST") = STYLE_CODE_PLM
                .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                Dim SEQ As Integer = Val(dst.Tables("ICTQUOT2").Compute("MAX(SEQ)", "") & "") + 1
                .Item("SEQ") = SEQ
                .Item("STYLE_PRICE") = rowICTSTYL1.Item("STYLE_PRICE")
                ' .Item("AVAILABILITY") = Mid(AVAILABILITY, 2)
                If ASCMAIN1.CLIENT = "VAN" Then
                    Dim SQCs As String = TAC.ICCMAIN1.Get_SIZEs_and_QTYs_and_COLORs(Me, STYLE_CODE_PLM)
                    .Item("SIZE_SCALE") = SQCs
                Else
                    .Item("SIZE_SCALE") = rowICTSTYL1.Item("SIZE_SCALE")
                End If
                .Item("STYLE_DESC2") = rowICTSTYL1.Item("STYLE_DESC2")
                .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                .Item("STYLE_GROUP_CODE") = rowICTSTYL1.Item("STYLE_GROUP_CODE")
                .Item("SEASON_CODE") = rowICTSTYL1.Item("SEASON_CODE")
                .Item("IMAGE_NAME") = rowICTSTYL1.Item("IMAGE_NAME")
            End With
            dst.Tables("ICTQUOT2").Rows.Add(rowICTQUOT2)
            'FetchImage(rowICTQUOT2)

            Load_Availability(rowICTQUOT2)
        End If

        Return STYLE_CODE_PLM
    End Function

    Sub Load_Availability(rowICTQUOT2 As DataRow)

        Dim STYLE_CODE_PLM As String = rowICTQUOT2.Item("STYLE_CODE_PLM")
        Dim A As Integer = 0
        ASCMAIN1.sql = "Select WHSE_CODE, STATUS_DATE, STATUS_QTY from ICTSTDQ1 where STYLE_CODE = :PARM1"
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {STYLE_CODE_PLM}).Select("", "WHSE_CODE,STATUS_DATE")
            A += 1
            If A <= 4 Then
                rowICTQUOT2.Item("WHSE_" & Format(A, "00")) = row.Item("WHSE_CODE")
                rowICTQUOT2.Item("DATE_" & Format(A, "00")) = row.Item("STATUS_DATE")
                rowICTQUOT2.Item("QTY_" & Format(A, "00")) = row.Item("STATUS_QTY")
            End If
        Next
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_PLM)
        rowICTQUOT2.Item("IMAGE_NAME") = rowICTSTYL1.Item("IMAGE_NAME")
        FetchImage(rowICTQUOT2)

    End Sub

    Private Sub tabStyles_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabStyles.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_QuoteSheet()
        If tabStyles.SelectedTab IsNot Nothing And tabStyles.SelectedTab.Key = "Overages && Shortages" Then Refresh_Excess_Inventory()
        If tabStyles.SelectedTab IsNot Nothing And tabStyles.SelectedTab.Key = "At-Once" Then Refresh_AtOnce()
        UltraExplorerBar1.Groups("At-Once").Visible = (tabStyles.SelectedTab.Key = "At-Once")

        UltraExplorerBar1.Groups("Screen Control").Visible = Not (tabStyles.SelectedTab.Key = "At-Once")

    End Sub

    Sub Setup_QuoteSheet()
        If Me.IsClosing And tabStyles.SelectedTab Is Nothing Then
        Else
            UltraExplorerBar1.Groups("Quote Sheet").Visible = (tabStyles.SelectedTab.Key = "Quote Sheet")
            UltraExplorerBar1.Groups("View Styles").Visible = (tabStyles.SelectedTab.Key = "Styles") And Not ScreenMode '  and Not ScreenMode was added by WJZ 10/30/19 because View Styles group was showing up while modes was true
            UltraExplorerBar1.Groups("Style Image").Visible = False ' True ' (tabStyles.SelectedTab.Key = "Quote Sheet")
            UltraExplorerBar1.Groups("Screen Control").Visible = Not (tabStyles.SelectedTab.Key = "Quote Sheet")

            spl.Panel1Collapsed = (tabStyles.SelectedTab.Key = "Quote Sheet")
            If tabStyles.SelectedTab.Key = "Quote Sheet" Then
                Setup_Style_Quoted()
                Refresh_Documents()
            End If
        End If
    End Sub

    Sub Refresh_Excess_Inventory()

        ASCMAIN1.sql = sqlSOTSUPPX & vbCrLf _
               & "   and ICTSTAT2.WHSE_CODE = 'MS'" & vbCrLf

        Dim sql_FUT_AVA As String = "  and NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0)"
        If chkOverbooked.Checked Then
            ASCMAIN1.sql &= sql_FUT_AVA & " < 0"
        Else
            ASCMAIN1.sql &= sql_FUT_AVA & " > 0 and (NVL(WHSE_QTY_ON_ORDER,0) > 0 or NVL(WHSE_QTY_TRAN,0) > 0)"
        End If

        ASCMAIN1.sql &= IIf(chkAllStyles.Checked, "", "  and ICTCLAS1.STYLE_CLASS_RELEASE_ATONCE = '1'" & vbCrLf)
        ASCMAIN1.sql &= IIf(chkAnyAva2Ship.Checked, "", "  and NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) > 0" & vbCrLf)

        Fill_Records("SOTSUPPX", "", True, ASCMAIN1.sql)

        If chkOverbooked.Checked Then
            grdSOTSUPPX.Text = "Shortages: Styles with " & "Open Orders > Net Qty Available (ie, Negative Future Available)" _
                & IIf(chkAnyAva2Ship.Checked, "", " with Ava2Ship > 0") _
                & IIf(chkAllStyles.Checked, "", ", At-Once Eligible")
        Else
            grdSOTSUPPX.Text = "Overages: Styles with " _
                & IIf(chkAnyAva2Ship.Checked, "Excess Future Available, ", "Excess On Hand, Ava2Ship > 0, ") _
                & "Future Avail > 0, Qty in Open PO or In Transit" _
                & IIf(chkAllStyles.Checked, "", ", At-Once Eligible")
        End If

        Sort_grdColumns(grdSOTSUPPX, "CUR_AVA_CST".ToLower)

    End Sub

    Sub Refresh_AtOnce()
        Fill_Records("ICTATOP1")
        Sort_grdColumns(grdICTATOP1, "STYLE_CODE,COLOR_CODE,ORDR_SHIP_DATE_ORIG")
        Fill_Records("ICTATOP2")
        Sort_grdColumns(grdICTATOP2, "STYLE_CODE,COLOR_CODE,PS_ETA")

    End Sub

    Private Sub grdICTQUOT2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTQUOT2.AfterRowActivate
        Setup_Style_Quoted()
    End Sub

    Sub Setup_Style_Quoted()
        If grdICTQUOT2.ActiveRow Is Nothing Then
            picStyleImage.Image = Nothing
            UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE & "-" & COLOR_CODE
        Else
            Dim IMAGE_NAME As String = grdICTQUOT2.ActiveRow.Cells("IMAGE_NAME").Value & ""
            Dim STYLE_CODE_PLM As String = grdICTQUOT2.ActiveRow.Cells("STYLE_CODE_PLM").Value & ""
            picStyleImage.Image = Get_Style_Image(IMAGE_NAME)
            UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE_PLM
        End If
    End Sub

    Sub email_Quote(tempFileName As String)
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim CUST_CODE As String = txtQuoteCUST_CODE.Text
        Dim CUST_NAME As String = txtQuoteCUST_NAME.Text
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        ATTACHMENTs.Add(tempFileName & ".pdf", ASCMAIN1.Folders("Temp") & tempFileName & ".pdf")

        Dim SUBJECT As String = "Quote Sheet"
        Dim PFX As String = ""

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        If CUST_CODE <> "" Then
            EMAIL_ADDRESSs.Add(rowARTCUST1.Item("CUST_EMAIL") & "", rowARTCUST1.Item("CUST_CONTACT") & "")
        End If

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, "ICTQUOT1", False, True, CUST_CODE, CUST_NAME, "Customer")
        If SEND_NO <> "" Then
            TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "QUOEML", "Quote Sheet emailed", SEND_NO)
        End If
    End Sub


    Private Sub txtCUST_SKU_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtCUST_SKU.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Dim STYLE_CODE As String = ""

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Searching for Customer SKU " & txtCUST_SKU.Text)

            ASCMAIN1.sql = "Select Distinct SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.CUST_SKU, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME" _
                & " from SOTORDR1,SOTORDR2 where SOTORDR2.CUST_SKU = :PARM1 and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {txtCUST_SKU.Text})
            If tbl.Rows.Count = 0 Then
                MsgBox("No Records Found", vbOKOnly, "Cannot Perform Requested Action")
            Else
                Using frm As New ASFMSGBF
                    frm.Show_grd(tbl, Me, "Select Style")
                    If frm.user_option <> -1 Then
                        STYLE_CODE = frm.grow.Cells("STYLE_CODE").Value & ""
                    End If
                End Using
            End If
            If STYLE_CODE <> "" Then
                Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
                Click_Command("Select")
                txtCUST_SKU.Text = ""
            End If

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If

    End Sub
    Sub Import_Markdowns(fileName As String)

        If fileName <> "" Then
            Dim oWB As SpreadsheetGear.IWorkbook
            Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
            Dim range As SpreadsheetGear.IRange = Nothing
            Try
                oWB = SpreadsheetGear.Factory.GetWorkbook(fileName)
                If oWB.Worksheets.Count = 0 Then
                    MsgBox("No Sheets Found")
                    Exit Sub
                End If

                Dim EXCEL_SHEET As String = ""
                For ws As Integer = 0 To oWB.Worksheets.Count - 1
                    If oWB.Worksheets(ws).Name = "Sheet1" Then
                        EXCEL_SHEET = oWB.Worksheets(ws).Name
                        Exit For
                    End If
                Next
                If EXCEL_SHEET = "" Then
                    MsgBox("No Data Sheet Found")
                    Exit Sub
                End If
                oSheet = oWB.Sheets(EXCEL_SHEET)



                Dim tblICTCOST1_batch As DataTable = ASCDATA1.GetDataTable("SELECT ICTCOST1.*, '' IMPORT_CODE from ICTCOST1 WHERE ROWNUM < 1")
                Dim batchCols As New List(Of String)

                For Each COLUMN_NAME As String In New String() {"COST_NOW", "ON_HAND", "VALUE_NOW", "COST_NEW", "VALUE_NEW", "MARKDOWN"}
                    batchCols.Add(COLUMN_NAME)
                    With tblICTCOST1_batch
                        .Columns.Add(COLUMN_NAME, GetType(System.Double))
                    End With
                Next

                Dim rCount As Int64 = oSheet.UsedRange.RowCount
                Dim rows_added As Int64 = -1
                Dim PX As Integer = 15
                Dim SX As Integer = 0
                Dim TRAN_NO As String = ASCMAIN1.Next_Control_No("ICTCOST1.TRAN_NO", 1)
                Dim totalMarkdown As Double = 0
                Dim OPS_YYYYPP_first As String = ""

                For r As Int64 = 2 To rCount - 1

                    Dim hasNewCost As Boolean = (oSheet.Cells(r, 10).Text & "" <> "")
                    Dim zeroDollar As Boolean = (hasNewCost And Val(oSheet.Cells(r, 10).Text) = 0)

                    Dim OPS_YYYYPP As String = oSheet.Cells(r, 0).Text
                    If OPS_YYYYPP_first = "" Then
                        OPS_YYYYPP_first = OPS_YYYYPP
                    End If

                    Dim STYLE_CODE As String = oSheet.Cells(r, 4).Text
                    Dim COLOR_CODE As String = oSheet.Cells(r, 5).Text
                    Dim TRAN_QTY As Decimal = Val(oSheet.Cells(r, 8).Text & "")
                    Dim TRAN_COST As Decimal = Val(oSheet.Cells(r, 10).Text & "")
                    Dim IMPORT_CODE As String = "0"
                    Dim COST_NOW As Double = CDbl(Val(oSheet.Cells(r, 7).Text))
                    Dim ON_HAND As Double = CDbl(Val(oSheet.Cells(r, 8).Text))

                    Dim VALUE_NOW As Double = Val(oSheet.Cells(r, 9).Text)

                    Dim COST_NEW As Double = 0

                    If hasNewCost Then
                        COST_NEW = Val(oSheet.Cells(r, 10).Text)
                    Else
                        COST_NEW = 0
                    End If
                    Dim VALUE_NEW As Double = CDbl(oSheet.Cells(r, 11).Text)
                    Dim MARKDOWN As Double = CDbl(oSheet.Cells(r, 12).Text)

                    If Not hasNewCost Then
                        IMPORT_CODE = "N"
                    End If
                    If OPS_YYYYPP_first <> OPS_YYYYPP Then
                        IMPORT_CODE = "E"
                    End If

                    Dim rowICTCOST1 As DataRow = tblICTCOST1_batch.NewRow
                    rowICTCOST1.Item("STYLE_CODE") = STYLE_CODE
                    rowICTCOST1.Item("COLOR_CODE") = COLOR_CODE
                    rowICTCOST1.Item("TRAN_NO") = TRAN_NO
                    rowICTCOST1.Item("TRAN_TYPE") = "M"
                    rowICTCOST1.Item("TRAN_DATE") = Mid(OPS_YYYYPP, 5, 2) & "/01/" & Mid(OPS_YYYYPP, 1, 4)
                    rowICTCOST1.Item("OPS_YYYYPP") = OPS_YYYYPP
                    rowICTCOST1.Item("TRAN_QTY") = TRAN_QTY
                    rowICTCOST1.Item("TRAN_COST") = TRAN_COST
                    rowICTCOST1.Item("INIT_DATE") = DATETIME_STAMP
                    rowICTCOST1.Item("INIT_OPER") = ASCMAIN1.USER_ID

                    rowICTCOST1.Item("COST_NOW") = COST_NOW
                    rowICTCOST1.Item("ON_HAND") = ON_HAND
                    rowICTCOST1.Item("VALUE_NOW") = VALUE_NOW
                    rowICTCOST1.Item("COST_NEW") = COST_NEW
                    rowICTCOST1.Item("VALUE_NEW") = VALUE_NEW
                    rowICTCOST1.Item("MARKDOWN") = MARKDOWN
                    rowICTCOST1.Item("IMPORT_CODE") = IMPORT_CODE

                    tblICTCOST1_batch.Rows.Add(rowICTCOST1)
                    tblICTCOST1_batch.AcceptChanges()

                    If rows_added = -1 Then rows_added = 0
                    rows_added += 1
                    totalMarkdown += MARKDOWN

                    If hasNewCost Then

                    Else

                    End If

                Next

                Using F As New ICFCOSTM

                    F.markdownBatch = tblICTCOST1_batch
                    F.batchTranNo = TRAN_NO
                    F.isBatchMarkdown = True
                    F.batchMarkdownColumns = batchCols
                    F.frmBase = Me

                    F.ShowDialog()
                    If F.updated Then
                        'don't think i need to do this if batch
                        'Dim YP_MIN As String = dst.Tables("ICTCOSTA").Compute("MIN(OPS_YYYYPP)", "") & ""
                        'Dim YP_MIN_CLOSED As String = dst.Tables("ICTCOSTA").Compute("MIN(OPS_YYYYPP)", "YP_OPEN = '1'") & ""
                        'If YP_MIN_CLOSED = "" Then YP_MIN_CLOSED = ASCMAIN1.CYP
                        'ASCDATA1.DeleteRows(dst.Tables("ICTCOSTL"), "OPS_YYYYPP_FIFO >= '" & YP_MIN_CLOSED & "'")
                        'ASCDATA1.DeleteRows(dst.Tables("ICTCOSTA"), "OPS_YYYYPP >= '" & YP_MIN_CLOSED & "'")
                        'Calculate_Costs()
                        'Fetch_Costs()
                    End If
                End Using

            Catch ex As Exception

            End Try
        End If




    End Sub
    Sub Integrity_Check()
        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, WHSE_CODE" & vbCrLf _
            & ", SUM (PO_DTL) PO_DTL, SUM (PO_SUM) PO_SUM" & vbCrLf _
            & ", SUM (PS_DTL) PS_DTL, SUM (PS_SUM) PS_SUM" & vbCrLf _
            & ", SUM (SO_DTL) SO_DTL, SUM (SO_SUM) SO_SUM" & vbCrLf _
            & ", SUM (SP_DTL) SP_DTL, SUM (SP_SUM) SP_SUM" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select 'IC' TYPE, STYLE_CODE, COLOR_CODE, WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, SUM (WHSE_QTY_ON_ORDER) PO_SUM" & vbCrLf _
            & ", 0 PS_DTL, SUM (WHSE_QTY_TRAN) PS_SUM" & vbCrLf _
            & ", 0 SO_DTL, SUM (WHSE_QTY_OPEN) SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, SUM (WHSE_QTY_PICK) SP_SUM" & vbCrLf _
            & " from ICTSTAT2" & vbCrLf _
            & " group by STYLE_CODE, COLOR_CODE, WHSE_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'PO' TYPE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_OPN) PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", 0 PS_DTL, 0 PS_SUM" & vbCrLf _
            & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
            & " from POTORDR2,POTORDR1 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'PS' TYPE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", SUM (POTSHIP3.PO_QTY_SHP) PS_DTL, 0 PS_SUM" & vbCrLf _
            & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
            & " from POTORDR2,POTSHIP1,POTSHIP2,POTSHIP3" & vbCrLf _
            & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'SO' TYPE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", 0 PS_DTL, 0 PS_SUM" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) SP_DTL, 0 SP_SUM" & vbCrLf _
            & " from SOTORDR2,SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "  and SOTORDR2.ORDR_STATUS <> 'D' and SOTORDR2.ORDR_STATUS <> 'C'" & vbCrLf _
            & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'SR' TYPE, SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, SOTRSRV1.WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", 0 PS_DTL, 0 PS_SUM" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY_OPEN) SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
            & " from SOTRSRV2,SOTRSRV1 where SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
            & " group by SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, SOTRSRV1.WHSE_CODE" & vbCrLf _
            & ")" & vbCrLf _
            & " group by STYLE_CODE, COLOR_CODE, WHSE_CODE" & vbCrLf _
            & "having SUM (PO_DTL) <> SUM (PO_SUM) or SUM (PS_DTL) <> SUM (PS_SUM) or SUM (SO_DTL) <> SUM (SO_SUM) or SUM (SP_DTL) <> SUM (SP_SUM)"

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Checking Status")

        Dim ICTSTATO As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTSTATO & " Add Primary Key (STYLE_CODE, COLOR_CODE, WHSE_CODE)")

        Fill_Records("ICTSTATO", "", True, "Select * from " & ICTSTATO)
        Sort_grdColumns(grdICTSTATO, "STYLE_CODE,COLOR_CODE,WHSE_CODE")

        Dim TBL As DataTable = ASCDATA1.GetDataTable

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


        If dst.Tables("ICTSTATO").Rows.Count = 0 Then
            MsgBox("All Styles are in Balance", MsgBoxStyle.OkOnly, "Success")
            grdICTSTATO.Visible = False
            tabStyles.Visible = True
        Else
            grdICTSTATO.Visible = True
            tabStyles.Visible = False
            'SplitContainer1.Visible = False
            frmPreAllocate.Visible = False

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")


            If MsgBox("Reset Status Summary Positions to Detail Qtys?", MsgBoxStyle.YesNo, "Option to Fix Summary Status Table ICTSTAT2") = MsgBoxResult.Yes Then
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare" & vbCrLf _
                    & "  Cursor C1 is Select * from " & ICTSTATO & ";" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update ICTSTAT2 Set " & vbCrLf _
                    & "    WHSE_QTY_ON_ORDER = R1.PO_DTL" & vbCrLf _
                    & "  , WHSE_QTY_TRAN = R1.PS_DTL" & vbCrLf _
                    & "  , WHSE_QTY_OPEN = R1.SO_DTL" & vbCrLf _
                    & "  , WHSE_QTY_PICK = R1.SP_DTL" & vbCrLf _
                    & "    where WHSE_CODE = R1.WHSE_CODE and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
                    & "   If SQL%NOTFOUND Then" & vbCrLf _
                    & "    Insert into ICTSTAT2" & vbCrLf _
                    & "     (WHSE_CODE, STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_ORDER, WHSE_QTY_TRAN, WHSE_QTY_OPEN, WHSE_QTY_PICK)" & vbCrLf _
                    & "    Values (R1.WHSE_CODE, R1.STYLE_CODE, R1.COLOR_CODE, R1.PO_DTL, R1.PS_DTL, R1.SO_DTL, R1.SP_DTL);" & vbCrLf _
                    & "   End If;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                MsgBox("Status Summary Qtys have been Reset to match Detail Qty", MsgBoxStyle.OkOnly, "Success")
            End If

        End If

    End Sub

    Private Sub grdICTSTATO_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTATO.DoubleClickRow
        Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
        Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
        Click_Command("Select")
    End Sub


    Private Sub grdICTSTATO_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTATO.InitializeRow
        For Each STAT As String In New String() {"PO", "PS", "SO", "SP"}
            If Val(e.Row.Cells(STAT & "_DTL").Value & "") <> Val(e.Row.Cells(STAT & "_SUM").Value & "") Then
                e.Row.Cells(STAT & "_SUM").Appearance.BackColor = Color.Yellow
            End If
        Next
    End Sub

    Private Sub grdICTCOSTL_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCOSTL.InitializeRow
        If e.Row.Cells("TRAN_TYPE").Value & "" = "J" Then
            e.Row.Cells("TRAN_TYPE").Appearance.ForeColor = Color.Orange
        ElseIf e.Row.Cells("TRAN_TYPE").Value & "" = "Z" Then
            e.Row.Cells("TRAN_TYPE").Appearance.ForeColor = Color.Red
        End If

        If Val(e.Row.Cells("LOT_QTY_ONHD").Value & "") > Val(e.Row.Cells("TRAN_QTY").Value & "") Then
            e.Row.Cells("LOT_QTY_ONHD").Appearance.ForeColor = Color.Red
            e.Row.Cells("LOT_QTY_ONHD").ToolTipText = "Qty On Hand exceeds the Lot Qty used to Establish Cost"
        End If

    End Sub

    Private Sub picStyleImage_Click(sender As System.Object, e As System.EventArgs) Handles picStyleImage.Click

    End Sub

    Private Sub picStyleImage_DoubleClick(sender As Object, e As System.EventArgs) Handles picStyleImage.DoubleClick
        Dim FOLDER_NAME = ASCMAIN1.Folders("Temp") & "Images"
        If Not My.Computer.FileSystem.DirectoryExists(FOLDER_NAME) Then
            My.Computer.FileSystem.CreateDirectory(FOLDER_NAME)
        End If

        Dim IMAGE_NAME = STYLE_CODE & "-" & COLOR_CODE & ".jpg"
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            IMAGE_NAME = rowICTSTYL1.Item("IMAGE_NAME")
        End If
        Dim IMAGE_FOLDER = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR")

        My.Computer.FileSystem.CopyFile(IMAGE_FOLDER & "\" & IMAGE_NAME, FOLDER_NAME & "\" & IMAGE_NAME, True)
        ' Shell("EXPORER " & FOLDER_NAME)
        Process.Start("explorer.exe", FOLDER_NAME)
    End Sub

    Private Sub imgSTYLE_Click(sender As System.Object, e As System.EventArgs) Handles imgSTYLE.Click

    End Sub

    Private Sub cmdAddMultipleStyles_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddMultipleStyles.Click
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                For Each STYLE_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    Add_to_Quote(STYLE_CODE)
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Quotes")
        ASCMAIN1.sql = sqlICTQUOTX
        'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        'If optShow.Value = "A" And CUST_CODE = "" Then
        grdICTQUOTX.Text = "All Quotes"
        'ElseIf optShow.Value = "M" Then
        '    ASCMAIN1.sql &= " and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
        '    grdICTQUOTX.Text = "Quotes entered or modified by Me"
        'ElseIf optShow.Value = "C" Or CUST_CODE <> "" Then
        '    ASCMAIN1.sql &= " and CUST_CODE = '" & CUST_CODE & "'"
        '    grdICTQUOTX.Text = "Quotes associated with " & CUST_CODE
        'End If
        Fill_Records("ICTQUOTX")
        Sort_grdColumns(grdICTQUOTX, "QUOTE_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


    End Sub

    Sub Modes_At_Once(tf As Boolean)

        ScreenMode = tf

        '   & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _

        With UltraExplorerBar1.Groups("At-Once")
            .Items("Edit At-Once").Visible = Not tf
            .Items("Update At-Once").Visible = tf
            .Items("Cancel At-Once").Visible = tf
        End With

        If tf Then
            grdICTATOP1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdICTATOP1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdICTATOP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdICTATOP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
        Else
            grdICTATOP1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdICTATOP1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdICTATOP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdICTATOP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        End If
        With grdICTATOP2.DisplayLayout.Bands(0)
            For Each C As String In New String() {"STYLE_ARRIVAL_BUFFER_DAYS", "STYLE_AT_ONCE_UNTIL", "STYLE_AT_ONCE_ACTIVE"}
                If tf Then
                    .Columns(C).CellActivation = Activation.AllowEdit
                    .Columns(C).CellAppearance.BackColor = Color.LightBlue
                Else
                    .Columns(C).CellActivation = Activation.NoEdit
                    .Columns(C).CellAppearance.BackColor = Color.Empty
                End If
            Next
        End With
        With grdICTATOP1.DisplayLayout.Bands(0)
            For Each C As String In New String() {"STYLE_SHIP_WINDOW_DAYS", "ORDR_SHIP_DATE_PLUS", "STYLE_AT_ONCE_UNTIL", "STYLE_AT_ONCE_ACTIVE"}
                If tf Then
                    .Columns(C).CellActivation = Activation.AllowEdit
                    .Columns(C).CellAppearance.BackColor = Color.LightBlue
                Else
                    .Columns(C).CellActivation = Activation.NoEdit
                    .Columns(C).CellAppearance.BackColor = Color.Empty
                End If
            Next
        End With


        tabStyles.Tabs("Styles").Enabled = Not tf
        tabStyles.Tabs("Quote Sheet").Enabled = Not tf
        tabStyles.Tabs("Overages && Shortages").Enabled = Not tf

        'spl.Panel1Collapsed = tf
        UltraGroupBox1.Visible = Not tf

    End Sub

    Sub Modes_Quote_Sheet(tf As Boolean)

        QuoteEntryMode = tf

        If tf Then
            splQuoteMain.Panel2Collapsed = False
            splQuoteMain.Panel1Collapsed = True
        Else
            splQuoteMain.Panel1Collapsed = False
            splQuoteMain.Panel2Collapsed = True
        End If

        With UltraExplorerBar1.Groups("Quote Sheet")
            .Items("New Quote Sheet").Visible = (splQuoteMain.Panel2Collapsed)
            .Items("Edit Quote Sheet").Visible = (splQuoteMain.Panel2Collapsed)
            .Items("Print Quote Sheet").Visible = Not (splQuoteMain.Panel2Collapsed)
            .Items("email Quote Sheet").Visible = Not (splQuoteMain.Panel2Collapsed)
            .Items("Clear Quote Sheet").Visible = Not (splQuoteMain.Panel2Collapsed)
            .Items("Save Quote Sheet").Visible = Not (splQuoteMain.Panel2Collapsed)
            .Items("Cancel Quote Sheet").Visible = Not (splQuoteMain.Panel2Collapsed)
            .Items("Delete Quote Sheet").Visible = Not (splQuoteMain.Panel2Collapsed)
        End With

        If Not tf Then
            dst.Tables("ICTQUOT1").Rows.Clear()
            dst.Tables("ICTQUOT2").Rows.Clear()
            QUOTE_NO = ""
        End If
    End Sub

    Sub Load_Record_Quote_Sheet()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        ' Save_Header_Fields(UltraGroupBox1)

        If QuoteEntryMode = "N" Then
            rowICTQUOT1 = dst.Tables("ICTQUOT1").NewRow
            QUOTE_NO = ASCMAIN1.Next_Control_No("ICTQUOT1.QUOTE_NO")
            rowICTQUOT1.Item("QUOTE_NO") = QUOTE_NO
            rowICTQUOT1.Item("CUST_CODE") = "" ' HFs("CUST_CODE")
            rowICTQUOT1.Item("QUOTE_DATE") = DATETIME_STAMP.Date '  HFs("QUOTE_DATE")
            rowICTQUOT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTQUOT1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTQUOT1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTQUOT1.Item("LAST_DATE") = DATETIME_STAMP
            rowICTQUOT1.Item("QUOTE_TYPE") = "S"
            dst.Tables("ICTQUOT1").Rows.Add(rowICTQUOT1)
        Else
            rowICTQUOT1 = Fill_Record("ICTQUOT1", QUOTE_NO)
            dst.AcceptChanges()
        End If

        Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
        If My.Computer.FileSystem.FileExists(FILENAME) Then
            rowICTQUOT1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        End If

        Fill_Records("ICTQUOT2", QUOTE_NO)
        Sort_grdColumns(grdICTQUOT2, "SEQ")

        For Each rowICTQUOT2 As DataRow In dst.Tables("ICTQUOT2").Select("")

            ASCMAIN1.Progress("-", rowICTQUOT2.Item("STYLE_CODE_PLM"))
            Load_Availability(rowICTQUOT2)
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdICTQUOTX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTQUOTX.DoubleClickRow
        If e.Row.IsDataRow Then
            QUOTE_NO = e.Row.Cells("QUOTE_NO").Value
            Absx1.txtFor("ICTQUOT1.QUOTE_NO").Text = QUOTE_NO
            Click_Command("Edit Quote Sheet")
        End If
    End Sub

    Private Sub grdICTCOSTA_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTCOSTA.AfterRowActivate
        Fetch_Costs()
    End Sub

    Private Sub grdICTCOSTA_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCOSTA.InitializeRow
        If e.Row.Cells("YP_OPEN").Value & "" = "1" Then
            e.Row.Cells("YP_OPEN").Appearance.BackColor = Color.LightPink
            e.Row.Cells("YP_OPEN").ToolTipText = "Costing is NOT Finalized for this period"
        End If
    End Sub

    Private Sub optSupplierOption_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSupplierOption.ValueChanged
        If txtSupplierOption.Text <> "" Then
            txtSupplierOption.Tag = "V"
            Setup_Recent()
        End If
    End Sub

    Private Sub splPA_SplitterMoved(sender As System.Object, e As System.Windows.Forms.SplitterEventArgs) Handles splPA.SplitterMoved

    End Sub

    Private Sub grdICTSTDQ2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTDQ2.AfterRowActivate
        Dim rowICTWHSE1 As DataRow = clsASCBASE1.LookUp("ICTWHSE1", grdICTSTDQ2.ActiveRow.Cells("WHSE_CODE").Value)

        If rowICTWHSE1 IsNot Nothing Then
            splAVAIL_LOCB2.Panel2Collapsed = IIf(rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1", False, True)

            Dim STYLE_CODE As String = grdICTSTATA.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTATA.ActiveRow.Cells("COLOR_CODE").Value
            Dim WHSE_CODE As String = grdICTSTDQ2.ActiveRow.Cells("WHSE_CODE").Value

            If rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1" Then
                ASCMAIN1.sql = " Select WHTLOCB1.WHSE_CODE" & vbCrLf _
                & " ,WHTLOCB1.LOCATION_CODE,'X' BAR_CODE" & vbCrLf _
                & " ,WHTLOCB1.STYLE_CODE,'X' COLOR_CODE" & vbCrLf _
                & " ,SUM (LOCATION_QTY) LOCATION_QTY" & vbCrLf _
                & " ,NULL INIT_DATE, NULL INIT_OPER, NULL LAST_DATE, NULL LAST_OPER" & vbCrLf _
                & " ,SUM (LOCATION_QTY_WAVE) LOCATION_QTY_WAVE,'X' LOAD_NO" & vbCrLf _
                & " ,WHTLOCM1.LOCATION_LOCKED" & vbCrLf _
                & "  from WHTLOCB1,WHTLOCM1,WHTBARC1 " & vbCrLf _
                & "  where WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "' " & vbCrLf _
                & "  and WHTLOCB1.COLOR_CODE = '" & COLOR_CODE & "' " & vbCrLf _
                & "  and WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "' " & vbCrLf _
                & "  and WHTBARC1.BAR_CODE (+) = WHTLOCB1.BAR_CODE" & vbCrLf _
                & "  and WHTLOCM1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE" & vbCrLf _
                & "  and WHTLOCM1.LOCATION_CODE (+) = WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & "  group by WHTLOCB1.WHSE_CODE,WHTLOCB1.LOCATION_CODE," & vbCrLf _
                & "  WHTLOCB1.STYLE_CODE, WHTLOCM1.LOCATION_LOCKED"
                Fill_Records("WHTLOCB1", , , ASCMAIN1.sql)
                Sort_grdColumns(grdWHTLOCB1, "LOCATION_CODE,STYLE_CODE,COLOR_CODE")
                Setup_grdWHTLOCB1()
                grdWHTLOCB1.Text = "Location Status " & STYLE_CODE & "-" & COLOR_CODE

                Fill_Records("WHTINSTX", New String() {STYLE_CODE, COLOR_CODE})
                Sort_grdColumns(grdWHTINSTX, "WAVE_NO")
                grdWHTINSTX.Text = "Open Waves " & STYLE_CODE & "-" & COLOR_CODE
            End If
        End If

    End Sub

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

    Private Sub cmdRefresh_Click(sender As System.Object, e As System.EventArgs) Handles cmdRefresh.Click
        Setup_Recent()
    End Sub

    Private Function GetFirstCost(ByVal STYLE_CODE As String) As Double
        Dim FIRST_COST As Double = 0
        STYLE_CODE = STYLE_CODE.Replace("'", "")
        Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
        SQLS.AppendLine("SELECT MAX(VEND_CODE) AS VEND_CODE FROM ICTSTYV1 WHERE STYLE_CODE = '" & STYLE_CODE & "'")
        ASCMAIN1.sql = SQLS.ToString()
        Dim VEND_CODE As String = ASCDATA1.GetDataValue

        Dim rowICTSTYV1 As DataRow = LookUp("ICTSTYV1", New String() {STYLE_CODE, VEND_CODE})
        If Not IsNothing(rowICTSTYV1) Then
            If IsDate(rowICTSTYV1.Item("NEW_PO_COST_DATE").ToString & "") Then
                If CDate(rowICTSTYV1.Item("NEW_PO_COST_DATE").ToString & "") < Now() Then
                    FIRST_COST = Val(rowICTSTYV1.Item("NEW_PO_COST").ToString & "")
                End If
            Else
                If IsDate(rowICTSTYV1.Item("PO_COST_DATE").ToString & "") Then
                    If CDate(rowICTSTYV1.Item("PO_COST_DATE").ToString & "") < Now() Then
                        FIRST_COST = Val(rowICTSTYV1.Item("PO_COST").ToString & "")
                    End If
                End If
            End If
        End If
        Return FIRST_COST
    End Function

    Private Function PARTIALSTYLE(STYLE_CODE As String) As String
        Dim RETVAL As String = ""
        Dim NEW_STYLE As String = ""
        ASCMAIN1.sql = String.Format("SELECT COUNT(*) RECCNT FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
        Dim STYLE_COUNT As Int16 = Val(ASCDATA1.GetDataValue)
        If STYLE_COUNT = 1 Then
            ASCMAIN1.sql = String.Format("SELECT STYLE_CODE FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
            NEW_STYLE = ASCDATA1.GetDataValue
            RETVAL = NEW_STYLE
        End If
        Return RETVAL
    End Function

    Private Sub chkShowAll_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowAll.CheckedChanged
        Show_Transaction_Details()
    End Sub

    Private Sub grdICTSTATB_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTATB.AfterRowActivate
        If optTD.Value = "ALL" And Not chkShowAll.Checked Then Show_Transaction_Details()
    End Sub

    Private Sub cmdMulti_Click(sender As System.Object, e As System.EventArgs) Handles cmdMulti.Click


        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            '  ASCMAIN1.CodeSelector.UseDataFromTable = tbl
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                If ASCMAIN1.CodeSelector.SelectedCodes.Count = 1 Then
                    Absx1.txtFor("STYLE_CODE").Text = ASCMAIN1.CodeSelector.SelectedCodes(0)
                    Click_Command("Select")
                Else
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Loading")

                    STYLEs.Clear()

                    For Each STYLE_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                        STYLEs.Add(STYLE_CODE)
                    Next

                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")

                    multi_style = True

                    Click_Command("Select")
                End If
            End If
        End If
    End Sub

    Sub Fill_ICTSTATA_ICTSTATW(STYLE_CODE As String, sqlICTSTYC1 As String)

        For Each TABLE_NAME As String In New String() {"ICTSTATA", "ICTSTATW"}

            ASCMAIN1.sql = "" _
                & IIf(TABLE_NAME = "ICTSTATA",
                    "Select X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC",
                    "Select X.STYLE_CODE, X.COLOR_CODE, X.WHSE_CODE, ICTWHSE1.WHSE_DESC") _
                      & vbCrLf _
                & ", SUM(X.BEG) BEG, SUM(X.SHP) SHP, SUM(X.RTN) RTN, SUM(X.REC) REC" & vbCrLf _
                & ", SUM(X.ADJ) ADJ, SUM(X.XFR) XFR, SUM(X.PHY) PHY, SUM(X.ON_HAND) ON_HAND" & vbCrLf _
                & ", SUM(X.ON_ORDER) ON_ORDER, SUM(X.TRAN) TRAN, SUM(X.OPEN) OPEN" & vbCrLf _
                & ", SUM(X.PICK) PICK, SUM(X.ALLO) ALLO, SUM(X.COMM) COMM, SUM(X.PROD) PROD" & vbCrLf _
                & IIf(TABLE_NAME = "ICTSTATA", ", MAX(UPC_CODE) UPC_CODE", "") _
                & IIf(TABLE_NAME = "ICTSTATA", ", MAX(STYLE_COLOR_STATUS) STYLE_COLOR_STATUS", "") _
                & IIf(TABLE_NAME = "ICTSTATA", " from ICTCOLR1, ICTSTYL1, (", " from ICTWHSE1, (") & vbCrLf _
                & "(Select ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE, ICTSTAT1.WHSE_CODE" & vbCrLf _
                & ", SUM(ICTSTAT1.WHSE_QTY_BEG) BEG" & vbCrLf _
                & ", SUM(ICTSTAT1.WHSE_QTY_SHP) SHP, SUM(ICTSTAT1.WHSE_QTY_RTN) RTN" & vbCrLf _
                & ", SUM(ICTSTAT1.WHSE_QTY_REC) REC, SUM(ICTSTAT1.WHSE_QTY_ADJ) ADJ" & vbCrLf _
                & ", SUM(ICTSTAT1.WHSE_QTY_XFR) XFR, SUM(ICTSTAT1.WHSE_QTY_PHY) PHY" & vbCrLf _
                & ", SUM(0) ON_HAND, SUM (0) ON_ORDER, SUM (0) TRAN, SUM (0) OPEN, SUM (0) PICK, SUM (0) ALLO, SUM (0) COMM, SUM (0) PROD" & vbCrLf _
                & IIf(TABLE_NAME = "ICTSTATA", ", NULL UPC_CODE", "") _
                & IIf(TABLE_NAME = "ICTSTATA", ", NULL STYLE_COLOR_STATUS", "") _
                & " from ICTSTAT1 " & vbCrLf _
                & " where ICTSTAT1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and ICTSTAT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                & " group by ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE, ICTSTAT1.WHSE_CODE)" & vbCrLf _
                & " union " & vbCrLf _
                & "(Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_CODE" & vbCrLf _
                & ", SUM(0) BEG, SUM (0) SHP, SUM (0) RTN, SUM (0) REC, SUM (0) ADJ, SUM (0) XFR, SUM (0) PHY" & vbCrLf _
                & ", SUM(ICTSTAT2.WHSE_QTY_ON_HAND) ON_HAND" & vbCrLf _
                & ", SUM(ICTSTAT2.WHSE_QTY_ON_ORDER) ON_ORDER, SUM(ICTSTAT2.WHSE_QTY_TRAN) TRAN" & vbCrLf _
                & ", SUM(ICTSTAT2.WHSE_QTY_OPEN) OPEN, SUM(ICTSTAT2.WHSE_QTY_PICK) PICK" & vbCrLf _
                & ", SUM(ICTSTAT2.WHSE_QTY_ALLO) ALLO" & vbCrLf _
                & ", SUM(ICTSTAT2.WHSE_QTY_COMM) COMM, SUM(ICTSTAT2.WHSE_QTY_PROD) PROD" & vbCrLf _
                & IIf(TABLE_NAME = "ICTSTATA", ", NULL UPC_CODE", "") _
                & IIf(TABLE_NAME = "ICTSTATA", ", NULL STYLE_COLOR_STATUS", "") _
                & " from ICTSTAT2 " & vbCrLf _
                & " where ICTSTAT2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & " group by ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_CODE)" & vbCrLf _
                & IIf(TABLE_NAME = "ICTSTATA", sqlICTSTYC1, "") _
                & ") X" & vbCrLf _
                & IIf(TABLE_NAME = "ICTSTATA",
                      " where ICTCOLR1.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                    & " group by X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC",
                      " where ICTWHSE1.WHSE_CODE (+) = X.WHSE_CODE" & vbCrLf _
                    & " group by X.STYLE_CODE, X.COLOR_CODE, X.WHSE_CODE, ICTWHSE1.WHSE_DESC")

            If TABLE_NAME = "ICTSTATA" Then
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", ICTSTAT1.WHSE_CODE", "")
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", ICTSTAT2.WHSE_CODE", "")
            End If

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim row2 As DataRow = dst.Tables(TABLE_NAME).Rows.Add(row.ItemArray)
            Next
            '   Fill_Records(TABLE_NAME, "", False, ASCMAIN1.sql)
        Next
        ' Fill_Records("ICTSTATA", STYLE_CODE, False)
    End Sub

    Sub Style_grdICTSTYL1_Recent()

        With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"QTY_ONHD", "QTY_ONPO", "QTY_TRAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_PROD", "QTY_NETA"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If gcol.Key = "QTY_NETA" Then
                        'gcol.CellAppearance.ForeColor = Color.Purple
                    ElseIf New String() {"QTY_ONHD", "QTY_ONPO", "QTY_TRAN"}.Contains(gcol.Key) Then
                        gcol.CellAppearance.ForeColor = Color.Green
                    Else
                        gcol.CellAppearance.ForeColor = Color.Red
                    End If
                    Create_Summary(grdICTSTYL1_Recent, gcol.Key)

                ElseIf New String() {"STYLE_CODE", "STYLE_STATUS", "STYLE_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    gcol.Header.Fixed = True
                ElseIf New String() {"VEND_CODE", "FACTORY_CODE", "STYLE_CART_CUBE", "CASE_CUBE", "VEND_ITEM_CODE", "STYLE_PO_QTY_MIN", "PURCH_NOTES", "REPLENISHMENT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                ElseIf New String() {""}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_XMAS_DATE", "STYLE_SO_QTY_MIN", "STYLE_COST_FIRST"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Else
                        gcol.Hidden = True
                    End If
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                End If
            Next

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                chkOmitPrice.Visible = False
                chkOmitPrice2.Visible = False
            End If

            If ASCMAIN1.CLIENT = "VAN" Then
                .Columns("DUTY_RATE_CODE").Hidden = False
            End If

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Else
                For Each COLUMN_NAME As String In New String() {"SUB_BODY_CODE", "MASTER_BODY_CODE", "SIZE_CODE", "FASHION_PROMO", "SUB_UNIT_PACK_QTY", "STYLE_MATL_DESC", "FABRIC_CODE", "FACTORY_CODE"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End If

            If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                For Each COLUMN_NAME As String In New String() _
                    {"STYLE_COST", "STYLE_COST_EXT"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
                Create_Summary(grdICTSTYL1_Recent, {"STYLE_COST_CUM"})
            Else
                'For Each COLUMN_NAME As String In New String() _
                '    {"STYLE_COST_LDP", "STYLE_COST_LDP_CODE", "STYLE_COST_ELC", "STYLE_COST_CUM"}
                '    .Columns(COLUMN_NAME).Hidden = True
                'Next
            End If

        End With

    End Sub

    Sub SHOW_DUTY_EXCEPTIONS()
        If ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then
            Dim DUTY_RATE_CODE As String = Absx1.txtFor("DUTY_RATE_CODE").Text
            Fill_Records("ICTDUTY4", New String() {DUTY_RATE_CODE})
            grdICTDUTY4.Text = "Duty Rate Modifiers by Country - " & DUTY_RATE_CODE
            Sort_grdColumns(grdICTDUTY4, "COUNTRY_CODE, DUTY_RATE_BEGIN")
        End If
    End Sub

    Private Sub numCARTONS_PER_UNIT_ValueChanged(sender As Object, e As EventArgs) Handles numCARTONS_PER_UNIT.ValueChanged
        If IsNumeric(numCARTONS_PER_UNIT.Value) Then
            If Val(numCARTONS_PER_UNIT.Value) > 0 Then
                numCARTONS_PER_UNIT.Appearance.BackColor = Color.OrangeRed
            Else
                numCARTONS_PER_UNIT.Appearance.BackColor = Color.Empty
            End If
        Else
            numCARTONS_PER_UNIT.Appearance.BackColor = Color.Empty
        End If
    End Sub

    Sub CalculateAtOnce()

        Dim WHSE_CODE As String = "MS"
        Dim SHIP_PLUS As Integer = Val(numSHIP_PLUS.Value & "")
        Dim ETA_PLUS As Integer = Val(numETA_PLUS.Value & "")

        Show_Filter(grdSOTORDRX, True)
        With grdSOTORDRX.DisplayLayout.Bands(0)
            .ColumnFilters.ClearAllFilters()
            .ColumnFilters("WHSE_CODE").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.Equals, WHSE_CODE)
            .ColumnFilters("OPEN").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.GreaterThan, 0)

            .Columns("SHIP_SEQ").Hidden = False
            .Columns("RECD_SEQ").Hidden = False
            .Columns("SHIP_DATE_PLUS").Hidden = False
            .Columns("ERROR").Hidden = False
        End With


        Dim PO_SEQ_MAX As Integer = 0
        dst.Tables("SOTSUPPA").Rows.Clear()

        Dim TOTAL_QTY_SUPPLY As Integer = 0

        Dim rowICTSTATW As DataRow = dst.Tables("ICTSTATW").Rows.Find(New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE})
        Dim QTY As Int64 = 0
        If rowICTSTATW IsNot Nothing Then
            QTY = Val(rowICTSTATW.Item("OTS_INV") & "")
        End If
        Dim row As DataRow = dst.Tables("SOTSUPPA").NewRow
        row.Item("PO_ARRIVAL_DATE") = Now.Date
        row.Item("PO_ARRIVAL_DATE_PLUS") = Now.Date
        row.Item("PO_QTY") = QTY
        TOTAL_QTY_SUPPLY += QTY
        row.Item("PO_QTY_USED") = 0
        row.Item("PO_SEQ") = 0
        dst.Tables("SOTSUPPA").Rows.Add(row)


        Dim C As Integer = 0
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("QTY_ALLO_" & Format(C, "0")).Hidden = False

        For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("WHSE_CODE = '" & WHSE_CODE & "'", "PO_ARRIVAL_DATE")
            Dim PO_ARRIVAL_DATE As Date = rowPOTORDRX.Item("PO_ARRIVAL_DATE")
            Dim ETA_PLUS_2 As Integer = ETA_PLUS
            If rowPOTORDRX.Item("STYLE_AT_ONCE_ACTIVE") & "" = "1" AndAlso rowPOTORDRX.Item("STYLE_AT_ONCE_UNTIL") & "" <> "" _
                AndAlso Format(rowPOTORDRX.Item("STYLE_AT_ONCE_UNTIL"), "yyyyMMdd") >= Format(Now, "yyyyMMdd") Then
                ETA_PLUS_2 = Val(rowPOTORDRX.Item("STYLE_ARRIVAL_BUFFER_DAYS") & "")
            End If

            rowPOTORDRX.Item("PO_ARRIVAL_DATE_PLUS") = PO_ARRIVAL_DATE.AddDays(ETA_PLUS_2)
        Next

        For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("WHSE_CODE = '" & WHSE_CODE & "'", "PO_ARRIVAL_DATE_PLUS")
            Dim PO_SHIPMENT_NO As String = rowPOTORDRX.Item("PO_SHIPMENT_NO") & ""
            Dim PO_ORDER_NO As String = rowPOTORDRX.Item("PO_ORDER_NO") & ""
            Dim PO_ARRIVAL_DATE As Date = rowPOTORDRX.Item("PO_ARRIVAL_DATE")
            Dim PO_ARRIVAL_DATE_PLUS As Date = rowPOTORDRX.Item("PO_ARRIVAL_DATE_PLUS")

            Dim PO_QTY As Int64 = 0
            If PO_SHIPMENT_NO = "" Then
                PO_QTY = Val(rowPOTORDRX.Item("PO_QTY_OPN") & "")
            Else
                PO_QTY = Val(rowPOTORDRX.Item("PO_QTY_SHP") & "")
            End If

            row = dst.Tables("SOTSUPPA").NewRow
            row.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            row.Item("PO_ORDER_NO") = PO_ORDER_NO
            row.Item("PO_ARRIVAL_DATE") = PO_ARRIVAL_DATE
            row.Item("PO_ARRIVAL_DATE_PLUS") = PO_ARRIVAL_DATE_PLUS

            'Dim ETA_PLUS_2 As Integer = ETA_PLUS
            'If rowPOTORDRX.Item("STYLE_AT_ONCE_ACTIVE") & "" = "1" AndAlso rowPOTORDRX.Item("STYLE_AT_ONCE_UNTIL") & "" <> "" _
            '    AndAlso Format(rowPOTORDRX.Item("STYLE_AT_ONCE_UNTIL"), "yyyyMMdd") >= Format(Now, "yyyyMMdd") Then
            '    ETA_PLUS_2 = Val(rowPOTORDRX.Item("STYLE_ARRIVAL_BUFFER_DAYS") & "")
            'End If

            'row.Item("PO_ARRIVAL_DATE_PLUS") = PO_ARRIVAL_DATE.AddDays(ETA_PLUS_2)
            row.Item("PO_QTY") = PO_QTY
            row.Item("PO_QTY_USED") = 0

            TOTAL_QTY_SUPPLY += PO_QTY

            C += 1
            row.Item("PO_SEQ") = C
            PO_SEQ_MAX = C

            dst.Tables("SOTSUPPA").Rows.Add(row)

            With grdSOTORDRX.DisplayLayout.Bands(0).Columns("QTY_ALLO_" & Format(C, "0"))
                .Hidden = False
                '.Header.Caption = Format(PO_ARRIVAL_DATE, "MM/dd") & "->" & Format(PO_ARRIVAL_DATE.AddDays(ETA_PLUS_2), "MM/dd")
                .Header.Caption = Format(PO_ARRIVAL_DATE, "MM/dd") & "->" & Format(PO_ARRIVAL_DATE_PLUS, "MM/dd")
                .Width = 120
                .Header.ToolTipText = "PO " & PO_ORDER_NO & IIf(PO_SHIPMENT_NO = "", "", $", PS {PO_SHIPMENT_NO}") & $", Qty = {CStr(PO_QTY)}"
            End With
        Next

        ' not sure how to implement the ship date plus

        Dim SHIP_SEQ As Integer = 0
        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select("OPEN <> 0 And WHSE_CODE = '" & WHSE_CODE & "'", "ORDR_SHIP_DATE, INIT_DATE") ' "ORDR_SHIP_DATE, ORDR_DATE_RECD")
                SHIP_SEQ += 1
            rowSOTORDRX.Item("SHIP_SEQ") = SHIP_SEQ
            Dim ORDR_SHIP_DATE As Date = rowSOTORDRX.Item("ORDR_SHIP_DATE")



            Dim STYLE_SHIP_WINDOW_DAYS As Integer = SHIP_PLUS
            If rowSOTORDRX.Item("STYLE_SHIP_WINDOW_DAYS") & "" <> "" Then

                Dim STYLE_AT_ONCE_UNTIL As Date = rowSOTORDRX.Item("STYLE_AT_ONCE_UNTIL")
                Dim STYLE_AT_ONCE_ACTIVE As String = rowSOTORDRX.Item("STYLE_AT_ONCE_ACTIVE") & ""

                If STYLE_AT_ONCE_ACTIVE = "1" And Format(STYLE_AT_ONCE_UNTIL, "yyyyMMdd") > Format(Now, "yyyyMMdd") Then
                    STYLE_SHIP_WINDOW_DAYS = Val(rowSOTORDRX.Item("STYLE_SHIP_WINDOW_DAYS") & "")
                End If
            End If

            Dim SHIP_DATE_PLUS As Date = ORDR_SHIP_DATE.AddDays(STYLE_SHIP_WINDOW_DAYS)
            rowSOTORDRX.Item("SHIP_DATE_PLUS") = SHIP_DATE_PLUS
            'If Format(SHIP_DATE_PLUS, "yyyyMMdd") > "20210601" Then Stop
            'If ASCMAIN1.Running_in_VS AndAlso rowSOTORDRX.Item("ORDR_GROUP_NO") = "0000732202" Then Stop
            rowSOTORDRX.Item("ERROR") = ""
            Dim PO_SEQ_MAX_WAIT As Integer = 0
            If PO_SEQ_MAX > 0 Then
                For PO_SEQ_TEST As Integer = PO_SEQ_MAX To 1 Step -1
                    Dim rowSOTSUPPA As DataRow = dst.Tables("SOTSUPPA").Select($"PO_SEQ = {CStr(PO_SEQ_TEST)}")(0)
                    Dim PO_ARRIVAL_DATE_PLUS As Date = rowSOTSUPPA.Item("PO_ARRIVAL_DATE_PLUS")
                    If Format(SHIP_DATE_PLUS, "yyyyMMdd") > Format(PO_ARRIVAL_DATE_PLUS, "yyyyMMdd") Then
                        PO_SEQ_MAX_WAIT = PO_SEQ_TEST
                        Exit For
                    End If
                Next
            End If
            rowSOTORDRX.Item("PO_SEQ_MAX_WAIT") = PO_SEQ_MAX_WAIT
        Next

        Dim TOTAL_QTY_DEMAND As Integer = 0
        Dim RECD_SEQ_MAX As Integer = 0
        Dim RECD_SEQ As Integer = 0
        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select("OPEN <> 0 AND WHSE_CODE = '" & WHSE_CODE & "'", "INIT_DATE, ORDR_GROUP_NO") '  "ORDR_DATE_RECD")
            RECD_SEQ += 1
            rowSOTORDRX.Item("RECD_SEQ") = RECD_SEQ
            Dim OPEN As Integer = Val(rowSOTORDRX.Item("OPEN") & "")
            TOTAL_QTY_DEMAND += OPEN
            If TOTAL_QTY_DEMAND > TOTAL_QTY_SUPPLY Then
            Else
                RECD_SEQ_MAX = RECD_SEQ
            End If
        Next

        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select($"OPEN <> 0 AND WHSE_CODE = '{WHSE_CODE}'", "INIT_DATE, ORDR_GROUP_NO") ' "ORDR_DATE_RECD")
            Dim OPEN As Int64 = Val(rowSOTORDRX.Item("OPEN") & "")
            Dim SHIP_DATE_PLUS As Date = rowSOTORDRX.Item("SHIP_DATE_PLUS")
            'If ASCMAIN1.Running_in_VS AndAlso rowSOTORDRX.Item("CUST_CODE") = "190630" Then Stop
            'If ASCMAIN1.Running_in_VS AndAlso rowSOTORDRX.Item("CUST_CODE") = "021566" Then Stop
            SHIP_SEQ = Val(rowSOTORDRX.Item("SHIP_SEQ") & "")
            RECD_SEQ = Val(rowSOTORDRX.Item("RECD_SEQ") & "")
            Dim PO_SEQ_MAX_WAIT As Integer = Val(rowSOTORDRX.Item("PO_SEQ_MAX_WAIT") & "")

            'rowSOTORDRX.Item("QTY_ALLO_0") = OPEN
            For i As Integer = 0 To 9
                rowSOTORDRX.Item("QTY_ALLO_" & CStr(i)) = DBNull.Value
            Next

            If RECD_SEQ > RECD_SEQ_MAX Then
                rowSOTORDRX.Item("ERROR") = "Ordered too late"
            Else

                Dim sqlPO_SEQ As String = $"PO_SEQ <= {CStr(PO_SEQ_MAX_WAIT)}"
                For Each row In dst.Tables("SOTSUPPA").Select(sqlPO_SEQ, "PO_SEQ DESC") ' PO_QTY_LEFT > 0 
                    Dim PO_ARRIVAL_DATE_PLUS As Date = row.Item("PO_ARRIVAL_DATE_PLUS")
                    Dim PO_SEQ As Integer = Val(row.Item("PO_SEQ") & "")

                    Dim slot As Boolean = False
                    If OPEN <= 0 Then
                        Exit For
                    End If


                    '********************************
                    ' slot = True

                    If Format(SHIP_DATE_PLUS, "yyyyMMdd") > Format(PO_ARRIVAL_DATE_PLUS, "yyyyMMdd") Or PO_SEQ = 0 Then ' Or di = 9 Then
                        slot = True
                        Dim PO_QTY_LEFT As Int32 = Val(row.Item("PO_QTY_LEFT") & "")
                        Dim PO_QTY_USED As Int32 = Val(row.Item("PO_QTY_USED") & "")
                        If PO_QTY_LEFT > OPEN Then
                            rowSOTORDRX.Item("QTY_ALLO_" & CStr(PO_SEQ)) = OPEN
                            PO_QTY_USED = PO_QTY_USED + OPEN
                            OPEN = 0
                        ElseIf PO_QTY_LEFT > 0 Then
                            rowSOTORDRX.Item("QTY_ALLO_" & CStr(PO_SEQ)) = PO_QTY_LEFT
                            PO_QTY_USED = PO_QTY_USED + PO_QTY_LEFT
                            OPEN = OPEN - PO_QTY_LEFT
                        End If

                        row.Item("PO_QTY_USED") = PO_QTY_USED
                    Else
                        'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then Stop
                        'rowSOTORDRX.Item("QTY_ALLO_0") = OPEN
                        'Dim PO_QTY_USED As Int32 = Val(row.Item("PO_QTY_USED") & "")
                        'PO_QTY_USED = PO_QTY_USED + OPEN
                        'row.Item("PO_QTY_USED") = PO_QTY_USED
                    End If

                    'If Format(SHIP_DATE_PLUS, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                    '    rowSOTORDRX.Item("ERROR") = "Already Late"
                    'End If

                    If OPEN > 0 Then ' (Not slot And PO_SEQ = 0) Or (slot And OPEN > 0 And PO_SEQ < PO_SEQ_MAX) Then
                        'If (slot Or PO_SEQ = 0) And OPEN > 0 And PO_SEQ <= PO_SEQ_MAX Then ' MT24796
                        Dim DI As Integer = -1

                        If (slot And PO_SEQ <> 0 And PO_SEQ < PO_SEQ_MAX) Then ' move backward in time to 0

                            For ii As Integer = PO_SEQ To 1 Step -1 ' PO_SEQ_MAX

                                DI = ii
                                row = dst.Tables("SOTSUPPA").Rows.Find(ii)
                                Dim PO_QTY_LEFT As Int32 = Val(row.Item("PO_QTY_LEFT") & "")
                                Dim PO_QTY_USED As Int32 = Val(row.Item("PO_QTY_USED") & "")

                                PO_ARRIVAL_DATE_PLUS = row.Item("PO_ARRIVAL_DATE_PLUS")
                                If Format(SHIP_DATE_PLUS, "yyyyMMdd") > Format(PO_ARRIVAL_DATE_PLUS, "yyyyMMdd") Then
                                Else
                                    If Format(SHIP_DATE_PLUS, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                                        rowSOTORDRX.Item("ERROR") = "Already Late"
                                    Else
                                        rowSOTORDRX.Item("ERROR") = "Arr after Ship+"
                                    End If
                                End If

                                If PO_QTY_LEFT > 0 Then
                                    If PO_QTY_LEFT > OPEN Then
                                        rowSOTORDRX.Item("QTY_ALLO_" & CStr(ii)) = OPEN
                                        PO_QTY_USED = PO_QTY_USED + OPEN
                                        row.Item("PO_QTY_USED") = PO_QTY_USED
                                        OPEN = 0

                                        Exit For
                                    Else
                                        rowSOTORDRX.Item("QTY_ALLO_" & CStr(ii)) = PO_QTY_LEFT
                                        PO_QTY_USED = PO_QTY_USED + PO_QTY_LEFT
                                        row.Item("PO_QTY_USED") = PO_QTY_USED
                                        OPEN = OPEN - PO_QTY_LEFT
                                    End If
                                End If
                            Next ii

                        ElseIf (Not slot Or PO_SEQ = 0) Then ' move forward in time

                            For ii As Integer = PO_SEQ + 1 To PO_SEQ_MAX

                                DI = ii
                                row = dst.Tables("SOTSUPPA").Rows.Find(ii)
                                Dim PO_QTY_LEFT As Int32 = Val(row.Item("PO_QTY_LEFT") & "")
                                Dim PO_QTY_USED As Int32 = Val(row.Item("PO_QTY_USED") & "")

                                PO_ARRIVAL_DATE_PLUS = row.Item("PO_ARRIVAL_DATE_PLUS")
                                If Format(SHIP_DATE_PLUS, "yyyyMMdd") > Format(PO_ARRIVAL_DATE_PLUS, "yyyyMMdd") Then
                                    ' not late
                                Else
                                    If Format(SHIP_DATE_PLUS, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                                        rowSOTORDRX.Item("ERROR") = "Already Late"
                                    Else
                                        rowSOTORDRX.Item("ERROR") = "Arr after Ship+"
                                    End If
                                End If

                                If PO_QTY_LEFT > 0 Then
                                    If PO_QTY_LEFT > OPEN Then
                                        rowSOTORDRX.Item("QTY_ALLO_" & CStr(ii)) = OPEN
                                        PO_QTY_USED = PO_QTY_USED + OPEN
                                        row.Item("PO_QTY_USED") = PO_QTY_USED
                                        OPEN = 0

                                        Exit For
                                    Else
                                        rowSOTORDRX.Item("QTY_ALLO_" & CStr(ii)) = PO_QTY_LEFT
                                        PO_QTY_USED = PO_QTY_USED + PO_QTY_LEFT
                                        row.Item("PO_QTY_USED") = PO_QTY_USED
                                        OPEN = OPEN - PO_QTY_LEFT
                                    End If
                                End If
                            Next
                        End If

                        Dim OVERALLOCATED As Boolean = False

                        If OPEN > 0 Then ' could not satisfy with open POs so try On Hand
                            row = dst.Tables("SOTSUPPA").Rows.Find(0)
                            Dim PO_QTY_LEFT As Int32 = Val(row.Item("PO_QTY_LEFT") & "")
                            Dim PO_QTY_USED As Int32 = Val(row.Item("PO_QTY_USED") & "")

                            If PO_QTY_LEFT >= OPEN Then
                                rowSOTORDRX.Item("QTY_ALLO_" & CStr(0)) = OPEN
                                PO_QTY_USED += OPEN
                                row.Item("PO_QTY_USED") = PO_QTY_USED
                                OPEN = 0
                            Else
                                If PO_QTY_LEFT < 0 Then
                                    rowSOTORDRX.Item("QTY_ALLO_" & CStr(0)) = OPEN
                                Else
                                    rowSOTORDRX.Item("QTY_ALLO_" & CStr(0)) = PO_QTY_LEFT
                                End If
                                If PO_QTY_LEFT < 0 Then
                                    PO_QTY_USED += OPEN
                                    'OPEN = 0
                                Else
                                    PO_QTY_USED += PO_QTY_LEFT
                                End If

                                row.Item("PO_QTY_USED") = PO_QTY_USED

                                If PO_QTY_USED > OPEN Then
                                    OPEN = 0
                                    OVERALLOCATED = True
                                Else
                                    OPEN = OPEN - PO_QTY_USED
                                End If
                            End If
                        End If

                        If OPEN > 0 Or OVERALLOCATED Then ' if there are any left by this time, just chuck them into the last PO
                            If DI < PO_SEQ_MAX Then
                                DI = DI + 1
                            End If

                            rowSOTORDRX.Item("ERROR") = "Past Cancel"

                            row = dst.Tables("SOTSUPPA").Rows.Find(DI)
                            Dim PO_QTY_LEFT As Int32 = Val(row.Item("PO_QTY_LEFT") & "")
                            Dim PO_QTY_USED As Int32 = Val(row.Item("PO_QTY_USED") & "")

                            rowSOTORDRX.Item("QTY_ALLO_" & CStr(DI)) = Val(rowSOTORDRX.Item("QTY_ALLO_" & CStr(DI)) & "") + OPEN
                            PO_QTY_USED += OPEN
                            row.Item("PO_QTY_USED") = PO_QTY_USED
                        End If
                    Else
                        ' rowSOTORDRX.Item("QTY_ALLO_0") = OPEN
                    End If
                Next
            End If
        Next


        Dim PO_QTY_CUM As Int64 = 0
        For Each row In dst.Tables("SOTSUPPA").Select("", "PO_SEQ")
            Dim PO_QTY_LEFT As Int64 = Val(row.Item("PO_QTY_LEFT") & "")
            PO_QTY_CUM += PO_QTY_LEFT
            row.Item("PO_QTY_CUM") = PO_QTY_CUM
        Next

        'Sort_grdColumns(grdSOTORDRX, "SHIP_SEQ")
        Sort_grdColumns(grdSOTORDRX, "RECD_SEQ")
        grdSOTORDRX.ActiveColScrollRegion.Scroll(UltraWinGrid.ColScrollAction.Right)

        If chkAutoAllocateAfterCalculate.Checked Then
            Allocate()
        End If
    End Sub

    Private Sub cmdCalculateAtOnce_Click(sender As Object, e As EventArgs) Handles cmdCalculateAtOnce.Click
        CalculateAtOnce()
    End Sub

    Private Sub grdSOTSUPPX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSUPPX.DoubleClickRow
        If e.Row.IsDataRow Then
            'optSelectBy.Value = "S"
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
            Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
            Click_Command("Select")
        End If
    End Sub

    Private Sub chkOverbooked_CheckedChanged(sender As Object, e As EventArgs) Handles chkOverbooked.CheckedChanged
        Refresh_Excess_Inventory()
    End Sub

    Private Sub grdSOTSUPPX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSUPPX.InitializeRow
        If e.Row.Band.Index = 0 Then
            Dim STYLE_COLOR_STATUS As String = e.Row.Cells("STYLE_COLOR_STATUS").Value & ""

            Select Case STYLE_COLOR_STATUS
                Case "D"
                    e.Row.Cells("STYLE_COLOR_STATUS").Appearance.ForeColor = Color.Red
                    '  e.Row.ToolTipText = "Style/Color " & e.Row.Cells("STYLE_CODE").Value & "/" & e.Row.Cells("COLOR_CODE").Value & " is Discontinued"
                Case "N"
                    e.Row.Cells("STYLE_COLOR_STATUS").Appearance.BackColor = Color.Yellow
                    '  e.Row.ToolTipText = "Style/Color " & e.Row.Cells("STYLE_CODE").Value & "/" & e.Row.Cells("COLOR_CODE").Value & " is Do Not Re-Order"
                Case Else
                    e.Row.Cells("STYLE_COLOR_STATUS").Appearance.ForeColor = Color.Empty
                    e.Row.Cells("STYLE_COLOR_STATUS").Appearance.BackColor = Color.Empty
                    '  e.Row.ToolTipText = ""
            End Select
        End If
    End Sub

    Private Sub chkAllStyles_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllStyles.CheckedChanged
        Refresh_Excess_Inventory()
    End Sub

    Private Sub chkAnyOnHand_CheckedChanged(sender As Object, e As EventArgs) Handles chkAnyAva2Ship.CheckedChanged
        Refresh_Excess_Inventory()
    End Sub

#Region "Promo System"
    Private Sub btnShowPromo_Click(sender As Object, e As EventArgs) Handles btnShowPromo.Click
        Dim F As New ASFMSGBF
        F.grdGroupBy = True
        F.grdFilter = True
        Dim sql As New Text.StringBuilder With {.Length = 0}
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
        Dim OnPromo As Boolean = False
        Dim PROMO_START_DATE As DateTime
        Dim PROMO_END_DATE As DateTime
        Dim sql As New Text.StringBuilder With {.Length = 0}
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
    End Sub

    Private Sub grdICTATOP2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTATOP2.InitializeRow
        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            If Format(e.Row.Cells("STYLE_AT_ONCE_UNTIL").Value, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Red
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").ToolTipText = "This At-Once Parameter has expired"
            Else
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Empty
            End If

            If e.Row.Cells("PS_ETA_NOW").Value & "" <> "" AndAlso Format(e.Row.Cells("PS_ETA_NOW").Value, "yyyyMMdd") <> Format(e.Row.Cells("PS_ETA").Value, "yyyyMMdd") Then
                e.Row.Cells("PS_ETA_NOW").Appearance.BackColor = Color.Yellow
                e.Row.Cells("PS_ETA_NOW").ToolTipText = "ETA has changed since the original Parameter record was created"
            Else
                e.Row.Cells("PS_ETA_NOW").Appearance.BackColor = Color.Empty
            End If

            If e.Row.Cells("PS_CODE").Value & "" = "P" Then
                e.Row.Cells("PS_CODE").Appearance.ForeColor = Color.Green
            Else
                e.Row.Cells("PS_CODE").Appearance.ForeColor = Color.Blue
            End If

        End If
    End Sub

    Private Sub grdICTATOP2_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdICTATOP2.DoubleClickRow
        If ScreenMode Then
            ' editing at-once
        Else
            If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value
                Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
                Click_Command("Select")
                For Each grow As UltraWinGrid.UltraGridRow In grdICTSTATA.Rows
                    If grow.Band.Index = 0 Then
                        If grow.Cells("COLOR_CODE").Value = COLOR_CODE Then
                            grow.Activate()
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub grdICTATOP1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTATOP1.InitializeRow
        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            If Format(e.Row.Cells("STYLE_AT_ONCE_UNTIL").Value, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Red
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").ToolTipText = "This At-Once Parameter has expired"
            Else
                e.Row.Cells("STYLE_AT_ONCE_UNTIL").Appearance.ForeColor = Color.Empty
            End If

            If e.Row.Cells("ORDR_SHIP_DATE_NOW").Value & "" <> "" AndAlso Format(e.Row.Cells("ORDR_SHIP_DATE_NOW").Value, "yyyyMMdd") <> Format(e.Row.Cells("ORDR_SHIP_DATE_ORIG").Value, "yyyyMMdd") Then
                e.Row.Cells("ORDR_SHIP_DATE_NOW").Appearance.BackColor = Color.Yellow
                e.Row.Cells("ORDR_SHIP_DATE_NOW").ToolTipText = "Ship Date has changed since the original Parameter record was created"
            Else
                e.Row.Cells("ORDR_SHIP_DATE_NOW").Appearance.BackColor = Color.Empty
            End If

        End If
    End Sub

    Private Sub grdICTATOP1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdICTATOP1.DoubleClickRow
        If ScreenMode Then
            ' editing at-once
        Else
            If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value
                Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
                Click_Command("Select")
                For Each grow As UltraWinGrid.UltraGridRow In grdICTSTATA.Rows
                    If grow.Band.Index = 0 Then
                        If grow.Cells("COLOR_CODE").Value = COLOR_CODE Then
                            grow.Activate()
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub grdSOTORDRX_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSOTORDRX.InitializeLayout

    End Sub

    Private Sub grdICTATOP1_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdICTATOP1.AfterCellUpdate

        If e.Cell.Column.Key = "STYLE_SHIP_WINDOW_DAYS" Then
            If grdICTATOP1.Tag & "" = "" Then
                grdICTATOP1.Tag = "X"
                Dim STYLE_SHIP_WINDOW_DAYS As Integer = Val(e.Cell.Value & "")
                Dim ORDR_SHIP_DATE_PLUS As Date = e.Cell.Row.Cells("ORDR_SHIP_DATE_PLUS").Value & ""
                Dim ORDR_SHIP_DATE_NOW As Date = e.Cell.Row.Cells("ORDR_SHIP_DATE_NOW").Value & ""
                If STYLE_SHIP_WINDOW_DAYS >= 0 Then
                    e.Cell.Row.Cells("ORDR_SHIP_DATE_PLUS").Value = ORDR_SHIP_DATE_NOW.AddDays(STYLE_SHIP_WINDOW_DAYS)
                Else
                    e.Cell.Row.Cells("STYLE_SHIP_WINDOW_DAYS").Value = ORDR_SHIP_DATE_PLUS.Subtract(ORDR_SHIP_DATE_NOW).TotalDays
                End If
                grdICTATOP1.Tag = ""
            End If
        ElseIf e.Cell.Column.Key = "ORDR_SHIP_DATE_PLUS" Then
            If grdICTATOP1.Tag & "" = "" Then
                grdICTATOP1.Tag = "X"
                Dim ORDR_SHIP_DATE_PLUS As Date = e.Cell.Value & ""
                Dim ORDR_SHIP_DATE_NOW As Date = e.Cell.Row.Cells("ORDR_SHIP_DATE_NOW").Value & ""
                Dim STYLE_SHIP_WINDOW_DAYS_ORIG As Integer = Val(e.Cell.Row.Cells("STYLE_SHIP_WINDOW_DAYS").Value & "")
                Dim STYLE_SHIP_WINDOW_DAYS As Integer = ORDR_SHIP_DATE_PLUS.Subtract(ORDR_SHIP_DATE_NOW).TotalDays
                If STYLE_SHIP_WINDOW_DAYS >= 0 Then
                    e.Cell.Row.Cells("STYLE_SHIP_WINDOW_DAYS").Value = STYLE_SHIP_WINDOW_DAYS
                Else
                    e.Cell.Row.Cells("ORDR_SHIP_DATE_PLUS").Value = ORDR_SHIP_DATE_NOW.AddDays(STYLE_SHIP_WINDOW_DAYS_ORIG)
                End If
                grdICTATOP1.Tag = ""
            End If
        End If

    End Sub

#End Region
End Class