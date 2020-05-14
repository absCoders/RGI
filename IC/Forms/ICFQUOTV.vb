Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

Public Class ICFQUOTV

    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim STYLE_CODE_allocated As String
    Dim AutoAllocate As Boolean
    Dim SOTDEMD1 As String
    Dim SOTSUPP1 As String
    Dim edi850cust As List(Of String)
    Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

    Dim rowICTSTYL1 As DataRow

    Dim STYLE_CLASS_CODE As String
    Dim CARTON_PACK_QTY As Int32
    Dim STYLE_PRICE As Decimal

    Dim sqlICTQUOTX As String
    Dim rowICTQUOT1 As DataRow

    Dim QUOTE_NO As String
    Dim XLS_NO As Integer = 0
    Dim COUNT_COLOR As String = "IIF(QTY_AVA>={0} OR ISNULL(OPEN_PICK_RSRV,0)>0,1,0)"
    Dim CUSTPOSs As New Dictionary(Of String, Integer)
    Dim STYLE_PRICE_copied As Decimal = 0
    Dim STYLE_GROUP_CODE_copied As String = ""

    Dim refresh_required As Boolean = False
    Dim IMG_Error_Reported As Boolean = False
    Dim PRINTING_SHEETS As Boolean = False
    Dim Form_Loading As Boolean = True

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("POTPARM1")

        TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me,
                "",
                False,
                True,
                False,
                "", Now.Date.AddDays(30)) ' using 30 days release date horizon

        With dst


            ASCMAIN1.sql = "Select ICTSTYL1.*,ICTBODY2.MASTER_BODY_CODE" & vbCrLf _
                & " from ICTSTYL1,ICTBODY2" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = :PARM1" & vbCrLf _
                & "   and ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf
            For Each TABLE_NAME As String In New String() {"ICTSTYL1", "ICTSTYL1_RECENT", "ICTSTYL1_VIEW"}
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "V", IIf(TABLE_NAME <> "ICTSTYL1_VIEW", 1, 0))
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
            Next

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

            ASCMAIN1.sql = "Select ICTQUOT2.*, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                & ", ICTSTYL1.SUB_BODY_CODE STYLE_GROUP_CODE, ICTSTYL1.SUB_BODY_CODE, ICTSTYL1.FABRIC_CODE" & vbCrLf _
                & ", ICTSTYL1.SEASON_CODE, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CARTON_PACK_QTY" _
               & " from ICTQUOT2, ICTSTYL1 where ICTSTYL1.STYLE_CODE = ICTQUOT2.STYLE_CODE_PLM" _
               & " and ICTQUOT2.QUOTE_NO = :PARM1"
            'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "ICTSTYL1.STYLE_GROUP_CODE", "ICTSTYL1.SUB_BODY_CODE STYLE_GROUP_CODE")
            'End If
            Create_TDA(.Tables.Add, "ICTQUOT2", "**", 0, True, "V", 2)
            With .Tables("ICTQUOT2")
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
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns.Add("STYLE_COST", GetType(System.Decimal))
                .Columns.Add("STYLE_RETAIL", GetType(System.Decimal))
                .Columns.Add("SALES_DIVISION_CODE_COMB")
                '.Columns.Add("LAST_SHIP_DATE")
            End With

            Create_TDA(.Tables.Add, "ICTQUOT3", "*")

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



            With .Tables.Add("SOTSDIVC")
                .Columns.Add("SALES_DIVISION_CODE")
                .Columns.Add("SALES_DIVISION_CODE_COMB")
            End With
            '  Dim SOTSDIVC As String = ASCMAIN1.Temp_Table
            '  ASCDATA1.ExecuteSQL("Alter Table " & SOTSDIVC & " Add Primary Key (SALES_DIVISION_CODE)")

            SOTSUPP1 = ASCMAIN1.Temp_Table("Select * from SOTSUPP1")
            ASCMAIN1.sql = "Select * from " & SOTSUPP1
            Create_TDA(.Tables.Add, "SOTSUPP1", "**", 0, False)

            SOTDEMD1 = ASCMAIN1.Temp_Table("Select * from SOTDEMD1")
            ASCMAIN1.sql = "Select * from " & SOTDEMD1
            Create_TDA(.Tables.Add, "SOTDEMD1", "**", 0, False)

            'ASCMAIN1.sql = "Select * from ICTSTYL3 where STYLE_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False, "V")


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

            ASCMAIN1.sql = "Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE" & vbCrLf _
                & ", NVL(ICTSTYC1.STYLE_COLOR_DESC,ICTCOLR1.COLOR_DESC) STYLE_COLOR_DESC" & vbCrLf _
                & " from ICTSTYC1,ICTCOLR1" & vbCrLf _
                & " where ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE and ICTSTYC1.STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False, "V", 2)
            With .Tables("ICTSTYC1").Columns
                For iCOL As Integer = 0 To 4
                    .Add("QTY_AVA" & CStr(iCOL), GetType(System.Int64))
                    .Add("DTE" & CStr(iCOL), GetType(System.DateTime))
                Next
                .Add("QTY_AVA", GetType(System.Int64), "ISNULL(QTY_AVA0,0)+ISNULL(QTY_AVA1,0)+ISNULL(QTY_AVA2,0)+ISNULL(QTY_AVA3,0)+ISNULL(QTY_AVA4,0)")
                .Add("OPEN_PICK_RSRV", GetType(System.Int64))
                .Add("COUNT_COLOR", GetType(System.Int32), String.Format(COUNT_COLOR, 0))
                .Add("SKIP_COLOR")
                .Add("LAST_RCD_DATE")
                .Add("EVER_ORDRED", GetType(System.Int64))
                '.Add("LAST_SHIP_DATE")
            End With

            Create_Relation("ICTQUOT2", "ICTSTYC1", "STYLE_CODE_PLM", "STYLE_CODE")



            With .Tables("ICTQUOT2").Columns
                .Add("QTY_AVA", GetType(System.Int64), "SUM(CHILD(ICTQUOT2_ICTSTYC1).QTY_AVA)")
                .Add("COUNT_COLOR", GetType(System.Int64), "SUM(CHILD(ICTQUOT2_ICTSTYC1).COUNT_COLOR)")
            End With

            ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC, SIZE_SCALE" & vbCrLf _
                & " from ICTSTYL1" & vbCrLf _
                & " where INIT_DATE > '01-JAN-2014'" & vbCrLf _
                & "    or STYLE_CODE in (Select STYLE_CODE from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0)"
            Create_TDA(.Tables.Add, "ICTSTYLX", "**", 0, False, "", 1)
            With .Tables("ICTSTYLX").Columns
                Dim QTOTAL As String = ""
                For I As Integer = 1 To 12
                    .Add("S" & CStr(I))
                    .Add("Q" & CStr(I), GetType(System.Int32))
                    QTOTAL &= "+ISNULL(Q" & CStr(I) & ",0)"
                Next
                .Add("SQ")
                .Add("QTOTAL", GetType(System.Int32), Mid(QTOTAL, 2))
            End With

            Create_TDA(.Tables.Add, "ICTSTYLS", "*")

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, STYLE_COLOR_DESC" & vbCrLf _
                & " from ICTSTYC1" & vbCrLf _
                & " where ICTSTYC1.STYLE_CODE in (" & vbCrLf _
                & "Select STYLE_CODE from ICTSTYL1" & vbCrLf _
                & " where INIT_DATE > '01-JAN-2014'" & vbCrLf _
                & "    or STYLE_CODE in (Select STYLE_CODE from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0))"
            Create_TDA(.Tables.Add, "ICTSTYCX", "**", 0, False, "", 2)

            Create_Relation("ICTSTYLX", "ICTSTYCX", "STYLE_CODE")

            ASCMAIN1.sql = "" _
                & "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", 'O' ORDR_TYPE, NVL(SOTORDR1.ORDR_CUST_PO,SOTORDR1.ORDR_NO) ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SUM (NVL(SOTORDR2.ORDR_QTY_OPEN,0)+NVL(SOTORDR2.ORDR_QTY_PICK,0)) QTY" & vbCrLf _
                & " from SOTORDR1,SOTORDR2" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.CUST_CODE = :PARM1" & vbCrLf _
                & "   and SOTORDR2.STYLE_CODE = :PARM2" & vbCrLf _
                & "   and SOTORDR2.ORDR_STATUS >= 'O' and SOTORDR2.ORDR_STATUS <= 'P'" & vbCrLf _
                & "   and NVL(SOTORDR2.ORDR_QTY_OPEN,0)+NVL(SOTORDR2.ORDR_QTY_PICK,0) > 0" & vbCrLf _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", NVL(SOTORDR1.ORDR_CUST_PO,SOTORDR1.ORDR_NO)" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
                & ", 'R' ORDR_TYPE, NVL(SOTRSRV1.ORDR_CUST_PO,SOTRSRV1.RSRV_NO) ORDR_CUST_PO" & vbCrLf _
                & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SUM (NVL(RSRV_QTY_OPEN,0)) QTY" & vbCrLf _
                & " from SOTRSRV1,SOTRSRV2" & vbCrLf _
                & " where SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV1.CUST_CODE = :PARM1" & vbCrLf _
                & "   and SOTRSRV2.STYLE_CODE = :PARM2" & vbCrLf _
                & "   and SOTRSRV1.RSRV_STATUS = 'O' and NVL(SOTRSRV2.RSRV_QTY_OPEN,0) > 0" & vbCrLf _
                & " group by SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
                & ", NVL(SOTRSRV1.ORDR_CUST_PO,SOTRSRV1.RSRV_NO)" & vbCrLf _
                & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE"
            Create_TDA(.Tables.Add, "SOTORDRC", "**", 0, False, "VV", 4)

            Create_TDA(.Tables.Add, "ASTROPT1", "*")

            Create_TDA(.Tables.Add, "ICTQUOH1", "*")
            Create_TDA(.Tables.Add, "ICTQUOH2", "*")
            With .Tables("ICTQUOH2").Columns
                .Add("SALES_DIVISION_NAME")
                .Add("SUB_BODY_DESC")
                .Add("FABRIC_DESC")
            End With

            ASCMAIN1.sql = "Select * from ASTROPT2 where FORM_NAME = :PARM1 and SET_ID = :PARM2 and XNO is Null"
            Create_TDA(.Tables.Add, "ASTROPT2", "**", 0, True, "VV")

            ASCMAIN1.sql = "SELECT SESSION_NO, FILE_NO, FILENAME, HASHVALUE FROM ICTQUOH2 WHERE SESSION_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTQUOHF", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "ASTATTA2", "*")
        End With

        With dst.Tables("ICTSTAT1_IMAGES")
            Dim row As DataRow = .NewRow
            row.Item("STYLE_CODE") = "X"
            Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                row.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
            End If
            .Rows.Add(row)
        End With

        Dim lstIncludeWhse As New List(Of String)
        lstIncludeWhse.Add("All Whse")
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT WHSE_CODE")
        sql.AppendLine("FROM ICTWHSE1")
        sql.AppendLine("WHERE NVL(WHSE_STATUS,'I') = 'A'")
        sql.AppendLine("ORDER BY WHSE_CODE")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        For Each rowICTWHSE1 As DataRow In tbl.Rows
            lstIncludeWhse.Add(rowICTWHSE1.Item("WHSE_CODE").ToString)
        Next
        cboIncludeWhse.DataSource = lstIncludeWhse
        cboIncludeWhse.SelectedIndex = 0

        edi850cust = TAC.SOCMAIN1.Get_EDI_Custs("850")

        grdICTQUOT2.DataSource = dst.Tables("ICTQUOT2")
        grdICTQUOT2B.DataSource = dst.Tables("ICTQUOT2")
        grdICTQUOTX.DataSource = dst.Tables("ICTQUOTX")
        grdICTSTDQ1.DataSource = dst.Tables("ICTSTDQ1")
        grdICTSTYCX.DataSource = dst.Tables("ICTSTYLX")
        grdSOTSDIVC.DataSource = dst.Tables("SOTSDIVC")
        grdICTQUOHF.DataSource = dst.Tables("ICTQUOHF")

        grdSOTSDIVC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        With grdSOTSDIVC.DisplayLayout.Override
            .AllowDelete = DefaultableBoolean.False
            .AllowAddNew = DefaultableBoolean.False
        End With

        With grdICTQUOHF.DisplayLayout.Override
            .AllowDelete = DefaultableBoolean.False
            .AllowAddNew = DefaultableBoolean.False
        End With

        Create_Summary(grdICTQUOTX, "QUOTE_NO", "Count")
        Create_Summary(grdICTQUOT2B, "STYLE_CODE_PLM", "Count")
        Create_Summary(grdICTQUOT2B, "SELECTED")

        grdICTSTYCX.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal

        With grdICTQUOT2.DisplayLayout.Bands(1)
            .Columns("STYLE_CODE").Header.Caption = "Style"
            .Columns("STYLE_CODE").Width = 80
            .Columns("COLOR_CODE").Header.Caption = "Clr"
            .Columns("COLOR_CODE").Width = 50
            .Columns("STYLE_COLOR_DESC").Header.Caption = "Color Description"
            .Columns("STYLE_COLOR_DESC").Width = 140
            .Columns("OPEN_PICK_RSRV").Hidden = True
            .Columns("COUNT_COLOR").Hidden = True
            .Columns("QTY_AVA").Header.Caption = "Avail"
            .Columns("EVER_ORDRED").Header.Caption = "Ever Ordered"
            .Columns("SKIP_COLOR").Hidden = True
        End With

        With grdICTQUOT2.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"QTY_AVA0", "DTE0", "QTY_AVA1", "DTE1", "QTY_AVA2", "DTE2", "QTY_AVA3", "DTE3", "QTY_AVA4", "DTE4", "OPEN_PICK_RSRV"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdICTQUOT2B.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"SEQ", "STYLE_PRICE", "SELECTED", "STYLE_GROUP_CODE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If
            Next
        End With


        With grdICTQUOT2.DisplayLayout.Bands(0)
            .Columns("IMAGE").Editor = New Infragistics.Win.EmbeddableImageRenderer
            .Columns("IMAGE").CellAppearance.ImageHAlign = HAlign.Center
            .Columns("IMAGE").Hidden = True
            For Each g As UltraWinGrid.UltraGridGroup In .Groups
                g.Header.Appearance.BackColor = Color.White
                g.Header.Appearance.BackColor2 = Color.LightBlue
                g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        grdICTQUOT2.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal

        With grdICTQUOT2.DisplayLayout.Bands(1)
            .Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key.StartsWith("QTY_AVA") Then
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                End If
            Next
        End With

        ' grdICTSTDQ1.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
        With grdICTSTDQ1.DisplayLayout.Bands(0)
            .Columns("WHSE_CODE").HiddenWhenGroupBy = DefaultableBoolean.True
            .Columns("STATUS_DATE").Format = "MM/dd/yy"
            .Columns("STATUS_DATE").Width = 85
            .Columns("QTY_ATS_CUM").Width = 65
        End With


        With grdSOTSDIVC.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                '  gcol.Header.Appearance.BackColor = Color.White
                '  gcol.Header.Appearance.BackColor2 = Color.LightGreen
                '  gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"SALES_DIVISION_CODE_COMB"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If
            Next
        End With

        Bind_Controls(grpICTQUOT1_Options, "ICTQUOT1")

        grpATS.Dock = DockStyle.Fill
        cbeWHSE_CODE.DataSource = ASCDATA1.GetDataTable("Select WHSE_CODE,WHSE_DESC from ICTWHSE1 order by WHSE_CODE")
        cbeWHSE_CODE.Value = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")

        chkAutoAllocate.Checked = True

        Dim DT As Date = CDate(Format(Now, "MM") & "/01/" & Format(Now, "yyyy")).AddMonths(1)
        If Now.Day > 15 Then
            DT = DT.AddMonths(1)
        End If

        With Absc1.grdSetup.DisplayLayout.Bands(0)
            '    .Columns("EXCLUDE").Hidden = True
            .Columns("SEQUENCE").Hidden = True
            .Columns("PAGE_BREAK").Hidden = True
            .Columns("SET_SEQ").Hidden = True
            .Columns("COLUMN_CAPTION").Width = 150
        End With
        Absc1.grdSetup.Text = "Set Column Filters"


        dte0.Value = Now.Date
        dte1.Value = DT.AddDays(-1)
        dte2.Value = DT.AddMonths(1).AddDays(-1)
        dte3.Value = DT.AddMonths(2).AddDays(-1)

        Dim datStartPeriod As New List(Of String)
        datStartPeriod.Add("Now")
        datStartPeriod.Add("1st")
        datStartPeriod.Add("2nd")
        datStartPeriod.Add("3rd")
        cboStartPeriod.DataSource = datStartPeriod

        dteInTranAsNow.Value = CDate(Now().ToShortDateString)

        Form_Loading = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                Else
                    Validate_Code("CUST_CODE")

                End If

            Case "View"
                If grdICTQUOTX.ActiveRow Is Nothing OrElse Not grdICTQUOTX.ActiveRow.IsDataRow Then
                    EMsg &= "You must Select (Or Double-click) a quote appearing In the grid In order To View it"
                Else
                    QUOTE_NO = grdICTQUOTX.ActiveRow.Cells("QUOTE_NO").Value
                End If

            Case "Edit"
                If ScreenMode Then

                Else
                    If grdICTQUOTX.ActiveRow Is Nothing OrElse Not grdICTQUOTX.ActiveRow.IsDataRow Then
                        EMsg &= "You must Select (Or Double-click) a quote appearing In the grid In order To View it"
                    Else
                        QUOTE_NO = grdICTQUOTX.ActiveRow.Cells("QUOTE_NO").Value
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTQUOT1", QUOTE_NO) Then Exit Sub
                End If

            Case "Cancel"
                If MsgBox("OK To Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"
                If ASCMAIN1.USER_ID <> rowICTQUOT1.Item("INIT_OPER") & "" Then
                    EMsg &= vbCr & "Only " & rowICTQUOT1.Item("INIT_OPER") & " may Delete this Quote"
                End If

                If EMsg = "" Then
                    If MsgBox("Do you really want To Delete this Quote",
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Excel"
                If refresh_required Then
                    EMsg &= vbCr & "Refresh Required"
                End If
                'If Absx1.txtFor("CUST_CODE").Text = "" Then
                '    If chkShowLastShip.Checked Then
                '        EMsg &= vbCr & "Customer Is Required When Showing Last Ship"
                '    End If
                'End If

            Case "Print", "email"
                If refresh_required Then
                    EMsg &= vbCr & "Refresh Required"
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Code Specified"
                    End If
                Else
                    'If chkShowLastShip.Checked Then
                    '    EMsg &= vbCr & "Customer Is Required When Showing Last Ship"
                    'End If
                End If
                'If chkOmitAvail.Checked And chkShowLastShip.Checked Then
                '    EMsg &= vbCr & "Omit Availability CAD Option Can Not Also Show Last Shipped"
                'End If
                If dst.Tables("ICTQUOT2").Rows.Count = 0 Then
                    EMsg &= vbCr & "No Styles On the Quote Sheet"
                Else
                    If dst.Tables("ICTQUOT2").Select("SELECTED='1'").Length = 0 Then
                        EMsg &= vbCr & "No Styles Selected to Print (You Must Select Styles to print CADs)"
                    End If
                End If

            Case "Update", "Save"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                Else
                    Validate_Code("CUST_CODE")

                End If

                If txtQUOTE_DESC.Text = "" Then
                    EMsg &= vbCr & "Please enter a Description for the Quote Sheet"
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

            Case "View"
                QUOTE_NO = grdICTQUOTX.ActiveRow.Cells("QUOTE_NO").Value
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Excel"

                For Each row As DataRow In dst.Tables("ICTSTYC1").Select("")
                    row.Item("OPEN_PICK_RSRV") = 0
                Next

                If chkShowLastRcd.Checked Then
                    setLastRcdDate()
                End If

                Dim EXCUDE_FUTURE As String = ""
                If chkExcudeFutureWhenZero.Checked Then
                    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select()
                        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                        SQLS.AppendLine("SELECT SUM(NVL(PO_QTY_ORD,0)) AS SUM_PO_QTY_ORD")
                        SQLS.AppendLine("FROM POTORDR2")
                        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", rowICTSTYC1.Item("STYLE_CODE").ToString & String.Empty))
                        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty))
                        ASCMAIN1.sql = SQLS.ToString()
                        rowICTSTYC1.Item("EVER_ORDRED") = Val(ASCDATA1.GetDataValue)
                    Next

                    Dim BegPeriod As Int64 = 0
                    Select Case tkb1.Value
                        Case 3
                            BegPeriod = 1
                        Case 2
                            BegPeriod = 2
                        Case 1
                            BegPeriod = 3
                        Case 0
                            BegPeriod = 4
                    End Select

                    For i As Integer = BegPeriod To 4
                        EXCUDE_FUTURE = EXCUDE_FUTURE & String.Format(" AND ISNULL(QTY_AVA{0},0)=0", i)
                    Next
                    If EXCUDE_FUTURE.Length > 0 Then
                        EXCUDE_FUTURE = " AND (" & EXCUDE_FUTURE.Substring(5, EXCUDE_FUTURE.Length - 5) & ")"
                    End If
                    'Dim COUNT_COLOR_EXCL As String = String.Format("IIF((QTY_AVA>={0} OR ISNULL(OPEN_PICK_RSRV,0)>0) {1},1,0)", Val(numMinQty.Value & ""), EXCUDE_FUTURE)
                    Dim COUNT_COLOR_EXCL As String = String.Format("IIF((QTY_AVA>={0} OR ISNULL(OPEN_PICK_RSRV,0)>0) AND (ISNULL(EVER_ORDRED,0)>0) {1},1,0)", Val(numMinQty.Value & ""), EXCUDE_FUTURE)
                    dst.Tables("ICTSTYC1").Columns("COUNT_COLOR").Expression = COUNT_COLOR_EXCL
                Else
                    dst.Tables("ICTSTYC1").Columns("COUNT_COLOR").Expression = String.Format(COUNT_COLOR, Val(numMinQty.Value & ""))
                End If

                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Save Excel As Link?"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("Would You Like To Save This Excel As A Link?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                Dim fileName As String = ""
                If chkSepDivision.Checked Then
                    For Each row As DataRow In ASCDATA1.SelectDistinct("ICTQUOT2", New String() {"SALES_DIVISION_CODE"}).Select("")
                        Dim SALES_DIVISION_CODE As String = row.Item(0) & ""
                        fileName = Create_Excel(SALES_DIVISION_CODE)
                    Next
                Else
                    fileName = Create_Excel()
                End If

                If fileName <> "" Then
                    If iResult = vbYes Then
                        'Dim SESSION_NO As String = ASCMAIN1.Next_Control_No("ICTQUOH1.SESSION_NO")
                        dst.Tables("ASTATTA2").Rows.Clear()
                        ENTITY.TABLE_NAME = "ICFQUOTV"
                        ENTITY.COLUMN_NAME = "QUOTE_NO"
                        ENTITY.CODE_VALUE = QUOTE_NO

                        MyBase.Attach_File(ASCMAIN1.Folders("Temp") & fileName, "Quote Sheet For Quote " & QUOTE_NO,,,, True)
                        'MyBase.Attach_File(fileName, "Quote Sheet For Quote " & QUOTE_NO,,,, True)

                        'Update ASTATTA2 immediatly.
                        Update_Record_TDA("ASTATTA2")

                        Dim MaxFILE_NO As Int64 = Val(dst.Tables("ICTQUOHF").Compute("Max(FILE_NO)", "") & "")
                        Dim SESSION_NO As String = dst.Tables("ICTQUOHF").Compute("Max(SESSION_NO)", "") & ""
                        MaxFILE_NO += 1
                        Dim newICTQUOHF As DataRow = dst.Tables.Item("ICTQUOHF").NewRow
                        Dim rowASTATTA2 As DataRow = dst.Tables.Item("ASTATTA2").Select().FirstOrDefault
                        newICTQUOHF.Item("SESSION_NO") = SESSION_NO
                        newICTQUOHF.Item("FILE_NO") = MaxFILE_NO
                        newICTQUOHF.Item("FILENAME") = rowASTATTA2.Item("ATTACHMENT_FILENAME").ToString & String.Empty
                        newICTQUOHF.Item("HASHVALUE") = rowASTATTA2.Item("HASHVALUE").ToString
                        dst.Tables.Item("ICTQUOHF").Rows.Add(newICTQUOHF)

                    End If
                Else
                    Dim iMSGX As New System.Text.StringBuilder With {.Length = 0}
                    iMSGX.AppendLine("No File Was Created.")
                    iMSGX.AppendLine("This Is Typically Because")
                    iMSGX.AppendLine("No Styles Qualify Based On")
                    iMSGX.AppendLine("The Selected Cryteria.")
                    iMSGX.AppendLine("")
                    iMSGX.AppendLine("Please Inspect Your Options.")
                    iResult = MsgBox(iMSGX.ToString(), MsgBoxStyle.OkOnly, "No File Created")
                End If
            Case "Buyer Sheet"
                Create_Excel_Buyer()
            Case "Buyer Chart"
                Create_Excel_BuyerChart()
            Case "Print", "email"
                Update_Record(True)
                IMG_Error_Reported = False
                PRINTING_SHEETS = True
                Print_Style_Sheet(eItemKey)
                PRINTING_SHEETS = False
            Case "Clear"
                dst.Tables("ICTQUOT2").Rows.Clear()
                dst.Tables("SOTSDIVC").Rows.Clear()
                dst.Tables("ICTQUOH1").Rows.Clear()
                dst.Tables("ICTQUOH2").Rows.Clear()
                dst.Tables("ICTQUOHF").Rows.Clear()

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                BeginTrans()
                ASCDATA1.ExecuteSQL("Delete from ICTQUOT1 where QUOTE_NO = '" & QUOTE_NO & "'")
                ASCDATA1.ExecuteSQL("Delete from ICTQUOT2 where QUOTE_NO = '" & QUOTE_NO & "'")
                CommitTrans("Quote Sheet " & QUOTE_NO & " has been Deleted")
                Mode_Settings(False)

            Case "Save"
                Update_Record()
            Case "Save as New"
                Dim QUOTE_NO_NEW As String = ASCMAIN1.Next_Control_No("ICTQUOT1.QUOTE_NO")

                'EnforceConstraints(False)
                Dim sqlw As String = "QUOTE_NO = '" & QUOTE_NO & "'"
                For Each row As DataRow In dst.Tables("ICTQUOT1").Select(sqlw)
                    row.Item("QUOTE_NO") = QUOTE_NO_NEW
                Next
                For Each row As DataRow In dst.Tables("ICTQUOT2").Select(sqlw)
                    row.Item("QUOTE_NO") = QUOTE_NO_NEW
                Next
                'EnforceConstraints(True)

                QUOTE_NO = QUOTE_NO_NEW
                Click_Command("Update")
                Mode_Settings(False)

                Absx1.txtFor("QUOTE_NO").Text = QUOTE_NO_NEW
                Click_Command("View")

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Load Last"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Load Last Saved Data"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Load All Last Saved Data")
                iMSG.AppendLine("And Over-Write The Existing Data.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim sql As New System.Text.StringBuilder With {.Length = 0}
                    sql.AppendLine("SELECT *")
                    sql.AppendLine("FROM ICTQUOT3")
                    sql.AppendLine(String.Format("WHERE QUOTE_NO = '{0}'", QUOTE_NO))
                    Dim tblICTQUOT3 As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
                    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("")
                        Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE").ToString & String.Empty
                        Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty
                        Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                        Dim rowICTQUOT3 As DataRow = tblICTQUOT3.Select(filter).FirstOrDefault
                        If Not IsNothing(rowICTQUOT3) Then
                            For Each dcol As DataColumn In tblICTQUOT3.Columns
                                If dcol.ColumnName <> "QUOTE_NO" Then
                                    rowICTSTYC1.Item(dcol.ColumnName) = rowICTQUOT3.Item(dcol.ColumnName)
                                End If
                            Next
                        End If
                    Next
                    Me.Cursor = Cursors.Default
                End If
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        tabStyles.Visible = ScreenMode
        ' grdICTQUOT2.Visible = ScreenMode
        grdICTQUOTX.Visible = Not ScreenMode
        wbv1.Visible = False

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Visible = Not ScreenMode
                '  .Items("Edit").Visible = ScreenMode And (EntryMode = "V")
                .Items("Edit").Visible = Not (EntryMode = "N" Or EntryMode = "E")
                .Items("View").Visible = Not ScreenMode
                .Items("Excel").Visible = ScreenMode
                .Items("Buyer Sheet").Visible = ScreenMode
                .Items("Buyer Chart").Visible = ScreenMode
                .Items("Done").Visible = ScreenMode And (EntryMode = "V")
                .Items("Save").Visible = ScreenMode And (EntryMode = "V")
                .Items("Print").Visible = ScreenMode
                .Items("email").Visible = False ' ScreenMode
                .Items("Clear Styles").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                .Items("Delete").Visible = ScreenMode And (EntryMode = "E")
                .Items("Save As New").Visible = ScreenMode And (EntryMode = "V")
                .Items("Load Last").Visible = ScreenMode And (EntryMode = "E" Or EntryMode = "V")
            End With
            ' .Groups("Options").Visible = ScreenMode
            .Groups("Available by Date").Visible = ScreenMode And (1 <> 1)
            .Groups("Availability Dates").Visible = True
            .Groups("Style Image").Visible = False
            .Groups("Style Color Filter").Visible = False
        End With

        ' If Not ScreenMode Then Setup_QuoteSheet()
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        cmdGetAvailability.Visible = ScreenMode

        cmdAddMultipleStyles.Visible = (EntryMode = "N" Or EntryMode = "E")
        btnByDate.Visible = (EntryMode = "N" Or EntryMode = "E")
        lblQS_STYLE_CODE.Visible = (EntryMode = "N" Or EntryMode = "E")
        txtQS_STYLE_CODE.Visible = (EntryMode = "N" Or EntryMode = "E")

        Set_TrackBar()

        chkShowSelectedOnly.Checked = False

        If ScreenMode Then
            btnFixColors.Visible = False
            'With grdICTQUOT2.DisplayLayout.Bands(0)
            '    .Groups(1).Width = 200
            '    .Groups(1).Width = 190
            'End With
            Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), False)
            Set_Read_Only_for_ctl(Absx1.txtFor("QUOTE_DESC"), False)
            Set_Read_Only_for_ctl(Absx1.dteFor("QUOTE_DATE"), False)
            Set_Read_Only_for_ctl(txtQS_STYLE_CODE, False)

            dst.Tables("ASTSQLX1").Rows.Clear()
            ' opt1Sheet.Enabled = chk1Sheet.Checked
        Else
            Clear_Record()
            dst.Tables("ASTSQLX1").Rows.Clear()
        End If

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        If EntryMode = "N" Then
            Absx1.txtFor("CUST_CODE").Text = ""

            rowICTQUOT1 = dst.Tables("ICTQUOT1").NewRow
            QUOTE_NO = ASCMAIN1.Next_Control_No("ICTQUOT1.QUOTE_NO")
            rowICTQUOT1.Item("QUOTE_NO") = QUOTE_NO
            rowICTQUOT1.Item("QUOTE_DESC") = HFs("QUOTE_DESC")
            rowICTQUOT1.Item("CUST_CODE") = HFs("CUST_CODE")
            rowICTQUOT1.Item("CUST_NAME") = HFs("CUST_NAME") ' NOT SURE WHY THIS BECAME NEC IN THIS FORM
            If HFs("QUOTE_DATE") = "" Then
                rowICTQUOT1.Item("QUOTE_DATE") = Now.Date
            Else
                rowICTQUOT1.Item("QUOTE_DATE") = CDate(HFs("QUOTE_DATE")) '  DATETIME_STAMP.Date '  HFs("QUOTE_DATE")
            End If
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

        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        'SQLS.AppendLine("SELECT MIN(SESSION_NO) AS SESSION_NO") 'This was always returning the original when there were multiple runs.  Changed by WR 9/15/16
        SQLS.AppendLine("SELECT MAX(SESSION_NO) AS SESSION_NO")
        SQLS.AppendLine("FROM ICTQUOH1")
        SQLS.AppendLine("WHERE QUOTE_NO = '" & QUOTE_NO & "'")
        ASCMAIN1.sql = SQLS.ToString()
        Dim SESSION_NO As String = ASCDATA1.GetDataValue

        Fill_Records("ICTQUOHF", SESSION_NO)
        addExtraICTQUOHF()

        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("")
            row.Item("STYLE_GROUP_CODE") = row.Item("STYLE_GROUP_CODE_SAVED")
        Next

        Sort_by_Style()

        dst.Tables("SOTSUPP0").Rows.Clear()
        dst.Tables("SOTSUPPI").Rows.Clear()
        dst.Tables("SOTORDR7").Rows.Clear()
        dst.Tables("ICTSTDQ1").Rows.Clear()
        dst.Tables("ICTSTDQ2").Rows.Clear()

        Retrieve_Settings(QUOTE_NO)
        Get_Availability()
        grdICTQUOT2.Rows.ExpandAll(True)

        ' Retrieve_Settings(QUOTE_NO)

        Setup_tabStyles()

        refresh_required = False
        cmdGetAvailability.Appearance.ForeColor = Color.Empty

        If chkShowLastRcd.Checked Then
            setLastRcdDate()
        End If

        'After binding set this as default if new
        'Make excuding Amazon by default for Kala and Brittany - 0721/19 - WR.
        If EntryMode = "N" Then
            chk1perPage.Checked = True
            chkIncludeWhse.Checked = False
            cboIncludeWhse.SelectedIndex = cboIncludeWhse.Items.IndexOf("AMZN")

        End If

        dteInTranAsNow.Value = CDate(Now().ToShortDateString)
        chkInTranAsNow.Checked = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ICTQUOT1", "ICTQUOT2", "ICTSTYC1", "ICTSTYLX", "ICTSTYCX", "SOTORDRC", "ICTQUOT3", "ICTSTYLS", "SOTSDIVC", "ICTQUOH1", "ICTQUOH2", "ICTQUOHF"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        QUOTE_NO = ""
        STYLE_PRICE_copied = 0
        STYLE_GROUP_CODE_copied = ""


        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("CUST_CODE").Focus()
    End Sub

    Sub Update_Record(Optional ByVal Silent As Boolean = False)

        Me.Cursor = Cursors.WaitCursor
        If Not Silent Then
            ASCMAIN1.Progress("Now Updating")
        End If

        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("")
            row.Item("STYLE_GROUP_CODE_SAVED") = row.Item("STYLE_GROUP_CODE")
        Next

        BeginTrans()

        Update_Record_TDA("ICTQUOT1", "QUOTE_NO = '" & QUOTE_NO & "'")
        Update_Record_TDA("ICTQUOT2", "QUOTE_NO = '" & QUOTE_NO & "'")
        Save_Settings(QUOTE_NO, txtQUOTE_DESC.Text)

        dst.Tables("ICTQUOT3").Rows.Clear()
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("")
            Dim rowICTQUOT3 As DataRow = dst.Tables("ICTQUOT3").NewRow
            For Each dcol As DataColumn In dst.Tables("ICTQUOT3").Columns
                If dcol.ColumnName = "QUOTE_NO" Then
                    rowICTQUOT3.Item(dcol.ColumnName) = QUOTE_NO
                Else
                    rowICTQUOT3.Item(dcol.ColumnName) = rowICTSTYC1.Item(dcol.ColumnName)
                End If
            Next
            dst.Tables("ICTQUOT3").Rows.Add(rowICTQUOT3)
        Next
        Update_Record_TDA("ICTQUOT3", "QUOTE_NO = '" & QUOTE_NO & "'")


        dst.Tables("ICTSTYLS").Rows.Clear()
        For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("")
            Dim rowICTSTYLS As DataRow = dst.Tables("ICTSTYLS").NewRow
            rowICTSTYLS.Item("STYLE_CODE") = rowICTSTYLX.Item("STYLE_CODE")
            'rowICTSTYLS.Item("SIZE_SCALE") = rowICTSTYLX.Item("SIZE_SCALE")
            rowICTSTYLS.Item("SIZE_SCALE") = GET_ONLY_SIZE_SCALE(rowICTSTYLX.Item("STYLE_CODE").ToString & String.Empty)
            rowICTSTYLS.Item("SIZE_QTYS") = rowICTSTYLX.Item("SQ")
            For I As Integer = 1 To 12
                Dim II As String = Format(I, "00")
                rowICTSTYLS.Item("SIZE_" & II) = rowICTSTYLX.Item("S" & CStr(I))
                rowICTSTYLS.Item("QTY_" & II) = rowICTSTYLX.Item("Q" & CStr(I))
            Next
            dst.Tables("ICTSTYLS").Rows.Add(rowICTSTYLS)
        Next
        ' VERIFY THAT THE DELETE CLAUSE IS OK
        ' DISABLING THIS UPDATE FOR NOW - UNTIL WE DO SHOWROOM APP - BECAUSE IT YIELDS A UNIQUE CONSTRAINT VIOLATION
        '   Update_Record_TDA("ICTSTYLS", "STYLE_CODE in (Select STYLE_CODE_PLM from ICTQUOT2 where QUOTE_NO = '" & QUOTE_NO & "')")

        Dim Notice As String = ""
        If Not Silent Then
            Notice = "Quote Sheet " & QUOTE_NO & " has been Saved"
        End If
        CommitTrans(Notice)
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
                Absx1.txtFor("QUOTE_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            'E.TABLE_NAME = "ICTSTYL1"
            'E.COLUMN_NAME = "STYLE_CODE"
            'E.CODE_VALUE = Absx1.txtFor("STYLE_CODE").Text
            'E.DESC_VALUE = Absx1.txtFor("STYLE_DESC").Text
            'E.ATTACHMENT_NOTES = ""
            'E.READ_ONLY = False
            E.TABLE_NAME = "ICFQUOTV"
            E.COLUMN_NAME = "QUOTE_NO"
            E.CODE_VALUE = QUOTE_NO
            'E.DESC_VALUE = Absx1.txtFor("STYLE_DESC").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "ICTQUOT1"
        E.TABLE_KEY_CAPTION = "Quote"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("QUOTE_NO").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("QUOTE_DESC").Text
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

        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTQUOHF, "BBBBB", "View File", "Replace File", "Copy Link", "Copy All Links", "Extend Expiration")
        Load_Popup_Menu(grdICTQUOTX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdICTSTYCX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry", "Style Master")
        Load_Popup_Menu(grdICTQUOT2, "BBBBBBB", "Sequence as Shown", "Select All", "De-Select All", "Style Status Inquiry", "Collapse All", "Expand All", "Sort by Style")
        Load_Popup_Menu(grdICTQUOT2B, "BBBBBBBBBB", "Sequence as Shown", "Select All", "De-Select All", "Select Selected", "Style Status Inquiry", "Sort by Style", "Copy Price", "Paste Price", "Copy Group", "Paste Group")
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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Sequence as Shown"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Re-Sequencing by 10's")

                Dim SEQ As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    SEQ += 1
                    grow.Cells("SEQ").Value = SEQ
                    grow.Update()
                Next

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("")

            Case "Paste Price"
                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("You Must First Select the Rows to Paste the Price to", MsgBoxStyle.OkOnly, "Cannot Paste")
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        grow.Cells("STYLE_PRICE").Value = STYLE_PRICE_copied
                        grow.Update()
                    Next
                End If
            Case "Paste Group"
                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("You Must First Select the Rows to Paste the Group to", MsgBoxStyle.OkOnly, "Cannot Paste")
                Else
                    If STYLE_GROUP_CODE_copied <> "" Then
                        For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                            grow.Cells("STYLE_GROUP_CODE").Value = STYLE_GROUP_CODE_copied
                            grow.Update()
                        Next
                    End If
                End If


            Case "Sort by Style"
                Sort_by_Style()

            Case "Expand All"
                grdICTQUOT2.Rows.ExpandAll(True)

            Case "Collapse All"
                grdICTQUOT2.Rows.CollapseAll(True)

            Case "Select All", "De-Select All"
                For Each grow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Select Selected"
                For Each grow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        If grow.Selected Then
                            grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                            grow.Update()
                        End If
                    End If
                Next
                grd.Selected.Rows.Clear()
            Case "Extend Expiration"
                If grd.Selected.Rows.Count = 1 Then
                    Dim grow As UltraWinGrid.UltraGridRow = grd.Selected.Rows(0)
                    Dim HASHVALUE As String = grow.Cells.Item("HASHVALUE").Text.ToString & String.Empty
                    Dim RetVal As Boolean = EXTEND_LINK(HASHVALUE)
                    If RetVal = True Then
                        MsgBox("Your Link Is Extended For 20 Days From Today.", vbOKOnly, "Extend Expiration")
                    Else
                        MsgBox("Could Not Find The Related Link.  Please Inform ABS.", vbOKOnly, "Extend Expiration")
                    End If
                Else
                    If grd.Selected.Rows.Count > 1 Then
                        MsgBox("You Can Only Update One Link At A Time.", vbOKOnly, "Extend Expiration")
                    Else
                        MsgBox("You Select A Row To Update.", vbOKOnly, "Extend Expiration")
                    End If
                End If
        End Select

        If grd Is Nothing OrElse grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "Copy Price"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    STYLE_PRICE_copied = Val(grd.ActiveRow.Cells("STYLE_PRICE").Value & "")
                End If

            Case "Copy Group"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    STYLE_GROUP_CODE_copied = grd.ActiveRow.Cells("STYLE_GROUP_CODE").Value & ""
                End If

            Case "Style Status Inquiry"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Value
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    If rowICTSTYL1 IsNot Nothing Then
                        Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                    End If
                End If

            Case "Style Master"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Edit", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                End If
            Case "View File"
                If Not IsNothing(grd.ActiveRow) Then
                    Dim FN As String = grd.ActiveRow.Cells.Item("FILENAME").Text
                    Dim SESSION_NO As String = grd.ActiveRow.Cells.Item("SESSION_NO").Text
                    Dim PDFD As String = ASCMAIN1.Folders("Archive") & "QuotePDFs\" & SESSION_NO & "\" & FN & ".pdf"
                    If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                        PDFD = "S:\VDI\ARCHIVE\VAN\QuotePDFs\" & SESSION_NO & "\" & FN & ".pdf"
                    End If
                    Show_Document(PDFD)
                End If
            Case "Replace File"
                Dim openFileDialog1 As New OpenFileDialog()
                If grd.ActiveRow.Cells.Item("FILENAME").Text.EndsWith(".pdf") Then
                    openFileDialog1.Filter = "pdf files (*.pdf)|*.pdf"
                Else
                    openFileDialog1.Filter = "excel files (*.xlsx)|*.xlsx"
                End If
                If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    Dim FN_FROM As String = openFileDialog1.FileName
                    Dim FN_TO As String = grd.ActiveRow.Cells.Item("FILENAME").Text
                    Dim SESSION_NO As String = grd.ActiveRow.Cells.Item("SESSION_NO").Text
                    If FN_TO.EndsWith(".pdf") Then
                        Dim PDFD As String = ASCMAIN1.Folders("Archive") & "QuotePDFs\" & SESSION_NO & "\" & FN_TO & ".pdf"
                        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                            PDFD = "S:\VDI\ARCHIVE\VAN\QuotePDFs\" & SESSION_NO & "\" & FN_TO & ".pdf"
                        End If
                        Dim iResult As MsgBoxResult
                        Dim iTitle As String = "Replace file"
                        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                        iMSG.AppendLine("This Action Will Replace The Generated File")
                        iMSG.AppendLine("With The Following File You Selected:")
                        iMSG.AppendLine(FN_FROM)
                        iMSG.AppendLine("")
                        iMSG.AppendLine("Is That What You Want?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult = MsgBoxResult.Yes Then
                            If IO.File.Exists(PDFD) Then
                                IO.File.Delete(PDFD)
                            End If
                            IO.File.Copy(FN_FROM, PDFD)
                        End If
                    Else
                        'Stop
                        'Dim ATTACH_FILE As String = dst.Tables.Item("ASTATT2").
                        Dim FILENAME As String = grd.ActiveRow.Cells.Item("FILENAME").Text
                        Dim filter As String = String.Format("ATTACHMENT_FILENAME = '{0}'", FILENAME)

                        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                        SQLS.AppendLine(String.Format("SELECT ATTACHMENT_NO FROM ASTATTA2 WHERE ATTACHMENT_FILENAME = '{0}'", FILENAME))
                        ASCMAIN1.sql = SQLS.ToString()
                        Dim ATTACHMENT_NO As String = ASCDATA1.GetDataValue

                        'Dim ATTACHMENT_NO As String = dst.Tables.Item("ASTATTA2").Select(filter).FirstOrDefault.Item("ATTACHMENT_NO").ToString & ""

                        Dim iResult As MsgBoxResult
                        Dim iTitle As String = "Replace file"
                        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                        iMSG.AppendLine("This Action Will Replace The Generated File")
                        iMSG.AppendLine("With The Following File You Selected:")
                        iMSG.AppendLine(FN_FROM)
                        iMSG.AppendLine("")
                        iMSG.AppendLine("Is That What You Want?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        Dim FN_TO_EXCL As String = ""
                        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                            Stop
                            FN_TO_EXCL = "G:\VDI\Attach\VAN\" & ATTACHMENT_NO
                        Else
                            FN_TO_EXCL = ASCMAIN1.Folders("Attach") & ATTACHMENT_NO
                        End If

                        If iResult = MsgBoxResult.Yes Then
                            If IO.File.Exists(FN_TO_EXCL) Then
                                IO.File.Delete(FN_TO_EXCL)
                            End If
                            IO.File.Copy(FN_FROM, FN_TO_EXCL)
                        End If
                    End If

                End If
            Case "Copy Link"
                Dim FILENAME As String = grd.ActiveRow.Cells.Item("FILENAME").Text
                Dim HASH As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                My.Computer.Clipboard.SetText(FILENAME & vbCrLf & LINEPFX)
            Case "Copy All Links"
                Dim clipbrd As String = ""
                For Each grow As UltraWinGrid.UltraGridRow In grdICTQUOHF.Rows
                    Dim FILENAME As String = grow.Cells.Item("FILENAME").Text
                    Dim HASH As String = grow.Cells.Item("HASHVALUE").Text
                    Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                    clipbrd = clipbrd & FILENAME & vbCrLf & LINEPFX & vbCrLf & vbCrLf
                Next
                My.Computer.Clipboard.SetText(clipbrd)
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If ctl.Text <> "" Then
                    'Call Click_Command("Load Reports")
                End If

        End Select
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "QUOTE_NO"
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
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)
            Case "QUOTE_NO"
                If txtctl.Text <> "" Then
                    Click_Command("Select")
                End If
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)

        Select Case COLUMN_NAME
            Case "SUPPLIER"
                'If txtSupplierOption.Text = "" Then
                '    txtSupplierOption.Tag = ""
                '    Setup_Recent()
                'End If
        End Select

    End Sub
#End Region

#Region "VB6"

    Private Sub SetRowImage(row As DataRow)
        Dim STYLE_CODE As String
        If row.Table.TableName = "ICTQUOT2" Then
            STYLE_CODE = row.Item("STYLE_CODE_PLM") & ""
        Else
            STYLE_CODE = row.Item("STYLE_CODE") & ""
        End If

        Dim IMAGE_NAME As String = row.Item("IMAGE_NAME") & ""

        If IMAGE_NAME = "" Then IMAGE_NAME = STYLE_CODE

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            IMAGE_NAME = STYLE_CODE & "-" & COLOR_CODE
        End If

        'Dim imgba() As Byte = Nothing
        Dim imgb As Bitmap = Nothing
        If IMAGE_NAME <> "" Then
            Dim ex_err As Exception = Nothing
            Dim IMAGE_FILE_USED As String = ""
            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")
                If chkLowRes.Checked Then
                    Dim FILE_NAME_LOW_RES As String = String.Format("{0}{1}{2}", FOLDER_NAME, "_lowres\", IMAGE_NAME)
                    If IO.File.Exists(FILE_NAME_LOW_RES) Then
                        FOLDER_NAME = FOLDER_NAME & "_lowres"
                        IMAGE_FILE_USED = FILE_NAME_LOW_RES
                    Else
                        IMAGE_FILE_USED = String.Format("{0}{1}{2}", FOLDER_NAME, "\", IMAGE_NAME)
                    End If
                End If
            End If

            Dim img As System.Drawing.Bitmap = Nothing

            Dim image_file_found As Boolean = True

            If IMAGE_NAME = "\.jpg" Then
                image_file_found = False
                Exit Sub
            End If

            If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
            Dim IMAGE_FILENAME As String = FOLDER_NAME & IMAGE_NAME
            Try
                If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then

                ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".PNG") Then
                    IMAGE_FILE_USED &= ".PNG"
                ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".JPG") Then
                    IMAGE_FILE_USED &= ".JPG"
                Else
                    image_file_found = False
                    img = Nothing
                End If
            Catch ex As Exception
                image_file_found = False
                img = Nothing
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    ex_err = ex
                End If
            End Try

            Dim fs As IO.FileStream = New IO.FileStream(IMAGE_FILENAME, IO.FileMode.Open)
            Dim newBMP As Bitmap = New Bitmap(Image.FromStream(fs))
            Dim scaleFactor As Double = (trkScaleImage.Value / 100)
            Dim newBMP2 As Bitmap = New Bitmap(newBMP, newBMP.Width * scaleFactor, newBMP.Height * scaleFactor)
            Application.DoEvents()
            Try
                'newBMP.MakeTransparent(System.Drawing.Color.White)
                Dim converter As New ImageConverter
                'row.Item("IMAGE") = converter.ConvertTo(newBMP, GetType(Byte()))
                row.Item("IMAGE") = converter.ConvertTo(newBMP2, GetType(Byte()))
                newBMP.Dispose()
                newBMP2.Dispose()
            Catch ex As Exception
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    ex_err = ex
                End If
            End Try
            fs.Close()
            Application.DoEvents()
            If Not IsNothing(ex_err) Then
                If Not IMG_Error_Reported Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Error Getting Image"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("The Following Error Occured While Attempting To ")
                    iMSG.AppendLine("Get An Image:")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Style: " & STYLE_CODE)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Image: " & IMAGE_NAME)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Error: " & ex_err.Message)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Please Relay This Information To Wayne At ABS.")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
                    IMG_Error_Reported = True
                End If
            End If
            'Dim converter As New ImageConverter
            'row.Item("IMAGE") = converter.ConvertTo(imgb, GetType(Byte()))
            'row.Item("IMAGE") = imgb
            'UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE & "-" & COLOR_CODE
        Else
            'row.Item("IMAGE") = DBNull.Value
            'UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        End If

    End Sub

    Function Get_Style_Image(
        ByVal IMAGE_NAME As String,
        Optional ByRef ex_err As Exception = Nothing) As System.Drawing.Bitmap

        Dim IMAGE_FILE_USED As String = ""
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")
            If chkLowRes.Checked Then
                Dim FILE_NAME_LOW_RES As String = String.Format("{0}{1}{2}", FOLDER_NAME, "_lowres\", IMAGE_NAME)
                If IO.File.Exists(FILE_NAME_LOW_RES) Then
                    FOLDER_NAME = FOLDER_NAME & "_lowres"
                    IMAGE_FILE_USED = FILE_NAME_LOW_RES
                Else
                    IMAGE_FILE_USED = String.Format("{0}{1}{2}", FOLDER_NAME, "\", IMAGE_NAME)
                End If
            End If
        End If

        'Return ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba, ex_err)

        Dim img As System.Drawing.Bitmap = Nothing

        Dim image_file_found As Boolean = True

        If IMAGE_NAME = "\.jpg" Then
            image_file_found = False
            Return Nothing
        End If

        If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
        Dim IMAGE_FILENAME As String = FOLDER_NAME & IMAGE_NAME
        Try
            If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                'img = System.Drawing.Image.FromFile(IMAGE_FILENAME)
                Dim fs As IO.FileStream = New IO.FileStream(IMAGE_FILENAME, IO.FileMode.Open)
                img = New Bitmap(Image.FromStream(fs))
                fs.Close()
            ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".PNG") Then
                'img = System.Drawing.Image.FromFile(IMAGE_FILENAME & ".PNG")
                Dim fs As IO.FileStream = New IO.FileStream(IMAGE_FILENAME & ".PNG", IO.FileMode.Open)
                img = New Bitmap(Image.FromStream(fs))
                fs.Close()
                IMAGE_FILE_USED &= ".PNG"
            ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".JPG") Then
                'img = System.Drawing.Image.FromFile(IMAGE_FILENAME & ".JPG")
                Dim fs As IO.FileStream = New IO.FileStream(IMAGE_FILENAME & ".JPG", IO.FileMode.Open)
                img = New Bitmap(Image.FromStream(fs))
                fs.Close()
                IMAGE_FILE_USED &= ".JPG"
            Else
                image_file_found = False
                img = Nothing
            End If

        Catch ex As Exception
            image_file_found = False
            img = Nothing
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                ex_err = ex
            End If
        End Try

        Try
            'img.MakeTransparent(System.Drawing.Color.White)
        Catch ex As Exception
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                ex_err = ex
            End If
        End Try

        'If img IsNot Nothing And image_file_found And return_byte_array Then
        '    byte_array = GetImageData(FOLDER_NAME & IMAGE_FILE_USED)
        'End If

        Return img
    End Function

    Private Sub Allocate(Optional Silent As Boolean = False)
        If Not Silent Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Allocating ... (Please Wait)")
        End If

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

        'dst.Tables("SOTSUPP0").Rows.Clear()
        'dst.Tables("SOTSUPPI").Rows.Clear()
        'dst.Tables("SOTORDR7").Rows.Clear()
        'dst.Tables("ICTSTDQ1").Rows.Clear()
        'dst.Tables("ICTSTDQ2").Rows.Clear()

        TAC.SOCMAIN1.Allocation(Me, False, True, "", "", edi850cust, SOTSUPP1, SOTDEMD1, TABLE_NAMEs, True, ("1" = "1"), STYLE_CODE, "", False, False, Not Silent) ' optASL.Value = "1"

        ' Truncate SOTORDR1 SOTORDR0 ARTCUST1 ICTSTDQ1 SOTORDR2 SOTRSRV1 SOTRSRV2
        ' Execute all sql's loaded into TABLE_NAMEs dictionary, in the order that they were placed
        ' Clear Rows for SOTSUPP0 SOTSUPPI SOTORDR7 and refill as necessary

        ASCMAIN1.sql = "Select SOTORDR7.* from SOTORDR7 where SOTORDR7.STYLE_CODE = '" & STYLE_CODE & "'" _
            & " and SOTORDR7.PICK_BATCH_NO is Null" & vbCrLf
        Fill_Records("SOTORDR7", "", True, ASCMAIN1.sql)

        Load_SOTALLO1()
        STYLE_CODE_allocated = STYLE_CODE

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

        If Not Silent Then
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
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
            For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select(sqlWSC, "SD_DATE, RECORD_TYPE, RECORD_SUB_TYPE")
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
    End Sub

#End Region

#Region "Custom Methods"
    Private Sub addExtraICTQUOHF()
        Dim MaxFILE_NO As Int64 = Val(dst.Tables("ICTQUOHF").Compute("Max(FILE_NO)", "") & "")
        Dim SESSION_NO As String = dst.Tables("ICTQUOHF").Compute("Max(SESSION_NO)", "") & ""
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT * ")
        sql.AppendLine("FROM ASTATTA2")
        sql.AppendLine("WHERE TABLE_NAME = 'ICFQUOTV'")
        sql.AppendLine("AND COLUMN_NAME = 'QUOTE_NO'")
        sql.AppendLine("AND CODE_VALUE = :PARM")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", QUOTE_NO)
        For Each rowASTATTA2 As DataRow In tbl.Rows
            MaxFILE_NO += 1
            Dim newICTQUOHF As DataRow = dst.Tables.Item("ICTQUOHF").NewRow
            newICTQUOHF.Item("SESSION_NO") = SESSION_NO
            newICTQUOHF.Item("FILE_NO") = MaxFILE_NO
            newICTQUOHF.Item("FILENAME") = rowASTATTA2.Item("ATTACHMENT_FILENAME").ToString
            newICTQUOHF.Item("HASHVALUE") = rowASTATTA2.Item("HASHVALUE").ToString
            dst.Tables.Item("ICTQUOHF").Rows.Add(newICTQUOHF)
        Next
    End Sub
#End Region
    Private Function EXTEND_LINK(ByVal HASHVALUE As String) As Boolean
        Dim RetVal As Boolean = False
        For Each TABLE_NAME As String In New String() {"ASTATTA2", "ICTQUOH2", "WEBLINKS"}
            If RetVal = False Then
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine(String.Format("SELECT COUNT(*) FROM {0} WHERE HASHVALUE = '{1}'", TABLE_NAME, HASHVALUE))
                ASCMAIN1.sql = SQLS.ToString()
                Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                If REC_CNT > 0 Then
                    RetVal = True
                    Dim SQLE As New System.Text.StringBuilder With {.Length = 0}
                    Dim NEW_DATE As String = Format(Now(), "dd-MMM-yyyy")
                    SQLE.AppendLine(String.Format("UPDATE {0} SET NEW_HASH_EXP = '{1}' WHERE HASHVALUE = '{2}'", TABLE_NAME, NEW_DATE, HASHVALUE))
                    ASCMAIN1.sql = SQLE.ToString
                    ASCDATA1.ExecuteSQL()
                End If
            End If
        Next
        Return RetVal
    End Function

    Private Sub grdICTSTDQ1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTDQ1.InitializeRow
        If Format(e.Row.Cells("STATUS_DATE").Value, "yyyyMMdd") = Format(Now.Date, "yyyyMMdd") Then
            e.Row.Appearance.BackColor = Color.Yellow
        End If
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
                '.Item("SIZE_SCALE") = rowICTSTYL1.Item("SIZE_SCALE")
                If ASCMAIN1.CLIENT = "VAN" Then
                    Dim SQCs As String = TAC.ICCMAIN1.Get_SIZEs_and_QTYs_and_COLORs(Me, STYLE_CODE_PLM)
                    .Item("SIZE_SCALE") = SQCs
                Else
                    .Item("SIZE_SCALE") = rowICTSTYL1.Item("SIZE_SCALE")
                End If
                .Item("STYLE_DESC2") = rowICTSTYL1.Item("STYLE_DESC2")
                .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                ' .Item("STYLE_GROUP_CODE") = rowICTSTYL1.Item("STYLE_GROUP_CODE")
                .Item("STYLE_GROUP_CODE") = rowICTSTYL1.Item("SUB_BODY_CODE")
                .Item("SUB_BODY_CODE") = rowICTSTYL1.Item("SUB_BODY_CODE")
                .Item("FABRIC_CODE") = rowICTSTYL1.Item("FABRIC_CODE")
                .Item("SEASON_CODE") = rowICTSTYL1.Item("SEASON_CODE")


                .Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
                .Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")

                .Item("IMAGE_NAME") = rowICTSTYL1.Item("IMAGE_NAME")
            End With
            dst.Tables("ICTQUOT2").Rows.Add(rowICTQUOT2)

            Load_Availability(rowICTQUOT2, True)

            grdICTQUOT2.Rows.ExpandAll(True)
        End If

        Return STYLE_CODE_PLM
    End Function

    Sub Load_Availability(rowICTQUOT2 As DataRow, Optional ByRef Silent As Boolean = False)

        STYLE_CODE = rowICTQUOT2.Item("STYLE_CODE_PLM")
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
        rowICTQUOT2.Item("IMAGE_NAME") = rowICTSTYL1.Item("IMAGE_NAME")
        rowICTQUOT2.Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
        rowICTQUOT2.Item("STYLE_RETAIL") = rowICTSTYL1.Item("STYLE_RETAIL")

        ' out of memory happens here
        'FetchImage(rowICTQUOT2)

        If dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "'").Length = 0 Then
            Fill_Records("ICTSTYC1", STYLE_CODE, False)

            ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC, SIZE_SCALE" & vbCrLf _
                & " from ICTSTYL1" & vbCrLf _
                & " where STYLE_CODE = '" & STYLE_CODE & "'"
            Fill_Records("ICTSTYLX", "", False, ASCMAIN1.sql)
            SET_NEW_SIZE_SCALE(STYLE_CODE)

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, STYLE_COLOR_DESC" & vbCrLf _
                & " from ICTSTYC1" & vbCrLf _
                & " where ICTSTYC1.STYLE_CODE = '" & STYLE_CODE & "'"
            Fill_Records("ICTSTYCX", "", False, ASCMAIN1.sql)

            'Fix_Colors(STYLE_CODE)
            Do While Fix_Colors(STYLE_CODE)

            Loop
            Fix_Size(STYLE_CODE)

            Dim rowICTSTYLX As DataRow = dst.Tables("ICTSTYLX").Rows.Find(STYLE_CODE)
            Dim SQ As String = ""
            For I As Integer = 1 To 12
                If rowICTSTYLX.Item("S" & CStr(I)) & "" <> "" Then
                    SQ &= " " & rowICTSTYLX.Item("S" & CStr(I)) & "/" & CStr(rowICTSTYLX.Item("Q" & CStr(I)))
                Else
                    Exit For
                End If
            Next
            Dim NEW_CAD_SIZE_SCALE As String = GET_CAD_SIZE_SCALE(STYLE_CODE)
            If NEW_CAD_SIZE_SCALE.Length = 0 Then
                rowICTSTYLX.Item("SQ") = Mid(SQ, 2)
            Else
                rowICTSTYLX.Item("SQ") = NEW_CAD_SIZE_SCALE
            End If

            'rowICTQUOT2.Item("SIZE_SCALE") = rowICTSTYLX.Item("SIZE_SCALE")
            rowICTQUOT2.Item("SIZE_SCALE") = GET_ONLY_SIZE_SCALE(STYLE_CODE)
        End If

        ASCDATA1.DeleteRows("ICTSTDQ1", "STYLE_CODE = '" & STYLE_CODE & "'")
        Allocate(Silent)

        For Each row As DataRow In dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "'")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim rowICTSTYCX As DataRow = dst.Tables("ICTSTYCX").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If rowICTSTYCX IsNot Nothing Then
                Dim STYLE_COLOR_DESC As String = rowICTSTYCX.Item("STYLE_COLOR_DESC") & ""
                If STYLE_COLOR_DESC <> "" Then
                    row.Item("STYLE_COLOR_DESC") = STYLE_COLOR_DESC
                End If
            End If

            ' IF WE DON'T DO THIS, THE QTYS DOUBLE AND TRIPLE, ETC
            For i As Integer = 0 To 4
                row.Item("QTY_AVA" & CStr(i)) = DBNull.Value
                row.Item("DTE" & CStr(i)) = DBNull.Value
            Next

            Dim fltrICTSTDQ1 As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            If cboIncludeWhse.Text <> "All Whse" Then
                If chkIncludeWhse.Checked Then
                    fltrICTSTDQ1 = fltrICTSTDQ1 + String.Format(" and WHSE_CODE = '{0}'", cboIncludeWhse.Text)
                Else
                    fltrICTSTDQ1 = fltrICTSTDQ1 + String.Format(" and WHSE_CODE <> '{0}'", cboIncludeWhse.Text)
                End If
            End If
            For Each rowICTSTDQ1 As DataRow In dst.Tables("ICTSTDQ1").Select(fltrICTSTDQ1, "STATUS_DATE")
                Dim STYLE_CODE_S As String = rowICTSTDQ1.Item("STYLE_CODE").ToString & String.Empty
                Dim COLOR_CODE_S As String = rowICTSTDQ1.Item("COLOR_CODE").ToString & String.Empty

                If ASCMAIN1.Running_in_VS Then
                    If rowICTSTDQ1.Item("STYLE_CODE").ToString = "VCO51509" Then Stop
                End If

                Dim QTY_ATS As Int64 = Val(rowICTSTDQ1.Item("QTY_ATS") & "")
                Dim STATUS_DATE As Date = rowICTSTDQ1.Item("STATUS_DATE")
                If STATUS_DATE < CDate(Now.ToShortDateString) Then
                    STATUS_DATE = CDate(Now.ToShortDateString)
                End If
                If chkInTranAsNow.Checked = True Then
                    Dim INTRANDATE As Date = CDate(dteInTranAsNow.Value)

                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine("SELECT (NVL(P3.PO_QTY_SHP,0) - NVL(P3.PO_QTY_REC,0)) AS IN_TRAN")
                    SQLS.AppendLine("FROM POTORDR2 P2, POTSHIP3 P3")
                    SQLS.AppendLine("WHERE P2.PO_ORDER_NO = P3.PO_ORDER_NO")
                    SQLS.AppendLine("AND P2.PO_ORDER_LNO = P3.PO_ORDER_LNO")
                    SQLS.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE_S))
                    SQLS.AppendLine(String.Format("AND P2.COLOR_CODE = '{0}'", COLOR_CODE_S))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim IN_TRAN As Int16 = Val(ASCDATA1.GetDataValue)
                    If IN_TRAN > 0 Then
                        If STATUS_DATE > CDate(Now.ToShortDateString) And STATUS_DATE <= INTRANDATE Then
                            STATUS_DATE = CDate(Now.ToShortDateString)
                        End If
                    End If
                End If

                Dim S As String = Format(STATUS_DATE, "yyyyMMdd")
                Dim i As Integer = 4
                If STATUS_DATE >= CDate(Now.ToShortDateString) Then
                    If Not chkNowOHonly.Checked And S <= Format(dte0.Value, "yyyyMMdd") Then
                        i = 0
                    ElseIf dte1.Visible And S <= Format(dte1.Value, "yyyyMMdd") Then
                        i = 1
                    ElseIf dte2.Visible And S <= Format(dte2.Value, "yyyyMMdd") Then
                        i = 2
                    ElseIf dte3.Visible And S <= Format(dte3.Value, "yyyyMMdd") Then
                        i = 3
                    End If

                    row.Item("QTY_AVA" & CStr(i)) = Val(row.Item("QTY_AVA" & CStr(i)) & "") + QTY_ATS
                    row.Item("DTE" & CStr(i)) = STATUS_DATE
                End If
            Next
        Next

        Dim A As Integer = 0
        If cboIncludeWhse.Text = "All Whse" Then
            ASCMAIN1.sql = "Select WHSE_CODE, STATUS_DATE, STATUS_QTY from ICTSTDQ1 where STYLE_CODE = :PARM1"
        Else
            If chkIncludeWhse.Checked Then
                ASCMAIN1.sql = String.Format("Select WHSE_CODE, STATUS_DATE, STATUS_QTY from ICTSTDQ1 where WHSE_CODE = '{0}' AND STYLE_CODE = :PARM1", cboIncludeWhse.Text)
            Else
                ASCMAIN1.sql = String.Format("Select WHSE_CODE, STATUS_DATE, STATUS_QTY from ICTSTDQ1 where WHSE_CODE <> '{0}' AND STYLE_CODE = :PARM1", cboIncludeWhse.Text)
            End If
        End If
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {STYLE_CODE}).Select("", "WHSE_CODE,STATUS_DATE")
            A += 1
            If A <= 4 Then
                rowICTQUOT2.Item("WHSE_" & Format(A, "00")) = row.Item("WHSE_CODE")
                rowICTQUOT2.Item("DATE_" & Format(A, "00")) = row.Item("STATUS_DATE")
                rowICTQUOT2.Item("QTY_" & Format(A, "00")) = row.Item("STATUS_QTY")
            End If
        Next

    End Sub

    Private Function GET_CAD_SIZE_SCALE(ByVal STYLE_CODE As String) As String
        Dim RetVal As String = ""
        Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
        If Not IsNothing(rowICTSTYLS) Then
            If rowICTSTYLS.Item("SIZE_01") & "" <> "" Then
                For iSZ As Integer = 1 To 24
                    If rowICTSTYLS.Item("SIZE_" & Format(iSZ, "00")) & "" = "" Then
                        Exit For
                    Else
                        Dim SIZE As String = rowICTSTYLS.Item("SIZE_" & Format(iSZ, "00")) & String.Empty
                        Dim UNITS As String = CStr(Val(rowICTSTYLS.Item("QTY_" & Format(iSZ, "00")) & ""))
                        RetVal = RetVal & String.Format("{0}-{1}{2}", SIZE, UNITS, vbCrLf)
                    End If
                Next
            End If
        End If
        Return RetVal
    End Function

    Private Sub SET_NEW_SIZE_SCALE(Optional ByVal STYLE_CODE As String = "")
        'If STYLE_CODE = "VCO51279" Then
        '    Stop
        'End If
        Dim Filter As String = ""
        If STYLE_CODE.Length > 0 Then
            Filter = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        End If
        For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select(Filter)
            Dim SC As String = rowICTSTYLX.Item("STYLE_CODE").ToString & String.Empty
            rowICTSTYLX.Item("SIZE_SCALE") = TAC.ICCMAIN1.Get_SIZEs_and_QTYs_and_COLORs(Me, SC)
        Next
    End Sub

    Private Function GET_ONLY_SIZE_SCALE(ByVal STYLE_CODE As String) As String
        Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
        'If STYLE_CODE = "VCO51279" Then
        '    Stop
        'End If
        Dim SIZEs As String = ""
        Dim QTYs As String = ""
        Dim SIZEs_And_QTYs As String = ""
        If rowICTSTYLS IsNot Nothing Then
            If rowICTSTYLS.Item("SIZE_01") & "" <> "" Then
                For iSZ As Integer = 1 To 24
                    If rowICTSTYLS.Item("SIZE_" & Format(iSZ, "00")) & "" = "" Then
                        Exit For
                    Else
                        SIZEs &= "-" & rowICTSTYLS.Item("SIZE_" & Format(iSZ, "00")) & ""
                        QTYs &= "/" & CStr(Val(rowICTSTYLS.Item("QTY_" & Format(iSZ, "00")) & ""))
                    End If
                Next
                SIZEs = Mid(SIZEs, 2) ' just the sizes
                If Not QTYs.StartsWith("/0") Then
                    SIZEs_And_QTYs = SIZEs & " = " & Mid(QTYs, 2)
                Else
                    SIZEs_And_QTYs = SIZEs
                End If
            End If
        End If
        Return SIZEs_And_QTYs
    End Function

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

    Private Sub cmdAddMultipleStyles_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddMultipleStyles.Click
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.Custom_sql_where = " and STYLE_CODE in (Select Distinct STYLE_CODE from ICTSTAT2 where NVL(WHSE_QTY_ON_HAND,0) <> 0 or NVL(WHSE_QTY_ON_ORDER,0) <> 0 or NVL(WHSE_QTY_TRAN,0) <> 0)"
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

            Sort_by_Style()
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Quotes")
        ASCMAIN1.sql = sqlICTQUOTX

        grdICTQUOTX.Text = "All Quotes"

        Fill_Records("ICTQUOTX")
        Sort_grdColumns(grdICTQUOTX, "QUOTE_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdICTQUOTX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTQUOTX.DoubleClickRow
        If e.Row.IsDataRow Then
            QUOTE_NO = e.Row.Cells("QUOTE_NO").Value
            Absx1.txtFor("QUOTE_NO").Text = QUOTE_NO
            Click_Command("View")
        End If
    End Sub

    Private Sub chkShowSelectedOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowSelectedOnly.CheckedChanged
        If Not Form_Loading Then
            If SELECTION_NO = 0 Then Exit Sub
            Dim dvw As DataView = DirectCast(grdICTQUOT2.DataSource, DataTable).DefaultView
            If chkShowSelectedOnly.Checked Then
                dvw.RowFilter = "SELECTED = '1'"
            Else
                dvw.RowFilter = ""
            End If
        End If
    End Sub

    Private Sub grdICTQUOT2_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdICTQUOT2.DoubleClickCell
        'If grdICTQUOT2.ActiveRow IsNot Nothing Then
        '    STYLE_CODE = e.Cell.Row.Cells("STYLE_CODE_PLM").Value & ""
        '    If e.Cell.Column.Key = "IMAGE" Then
        '        PRINTING_SHEETS = True
        '        Print_Style_Sheet("Print", STYLE_CODE)
        '        PRINTING_SHEETS = False
        '    End If
        '    If e.Cell.Column.Key = "IMAGE" Then
        '        Allocate()
        '    End If
        'End If
    End Sub

    Private Sub grdICTQUOT2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTQUOT2.InitializeRow
        If e.Row.Band.Key = "ICTQUOT2_ICTSTYC1" Then
            Exit Sub
        End If
        e.Row.Cells("lblDivision").Value = "Division"
        e.Row.Cells("lblSeason").Value = "Season"
        e.Row.Cells("lblClass").Value = "Class"
        'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        '    e.Row.Cells("lblGroup").Value = "Sub-Body"
        'Else
        e.Row.Cells("lblGroup").Value = "Group"
        'End If


        e.Row.Cells("lblStyle").Value = "Style"
        e.Row.Cells("lblPrice").Value = "Price"
        e.Row.Cells("lblSelect").Value = "Select"
        e.Row.Cells("lblSeq").Value = "Seq"

        e.Row.Cells("lblDescription").Value = "Description"
        e.Row.Cells("lblComment").Value = "Comment"
        e.Row.Cells("lblSizeScale").Value = "Size Scale"
        e.Row.Cells("lblCustomerStyle").Value = "Customer Style"
    End Sub

    Private Function Create_Excel(Optional SALES_DIVISION_CODE As String = "") As String
        Dim RetVal As String = ""

        RESEQ()

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim sqlWB As String = ""
        If SALES_DIVISION_CODE <> "" Then
            sqlWB = " and SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
            ASCMAIN1.Progress("Now Creating Workbook for Divison " & SALES_DIVISION_CODE, "")
        Else
            ASCMAIN1.Progress("Now Creating Workbook", "")
        End If

        Dim sql0 As String = " and COUNT_COLOR > 0" ' & Val(numMinQty.Value & "")
        If chkShowSelectedOnly.Checked Then
            sql0 &= " and SELECTED = '1'"
        End If

        CUSTPOSs.Clear()

        Dim CUSTPOi As Integer = 0
        dst.Tables("SOTORDRC").Rows.Clear()

        For Each row As DataRow In dst.Tables("ICTSTYC1").Select("")
            row.Item("OPEN_PICK_RSRV") = 0
        Next

        If chkShowPOs.Checked Then
            For Each row As DataRow In dst.Tables("ICTQUOT2").Select("")
                STYLE_CODE = row.Item("STYLE_CODE_PLM")
                Fill_Records("SOTORDRC", New String() {txtQuoteCUST_CODE.Text, STYLE_CODE}, False)
            Next
            For Each row As DataRow In dst.Tables("SOTORDRC").Select("", "ORDR_CANCEL_DATE")
                Dim OPO As String = row.Item("ORDR_TYPE") & vbTab & row.Item("ORDR_CUST_PO") & vbTab & Format(row.Item("ORDR_SHIP_DATE"), "MM/dd/yyyy") & vbTab & Format(row.Item("ORDR_CANCEL_DATE"), "MM/dd/yyyy")
                If Not CUSTPOSs.ContainsKey(OPO) Then
                    CUSTPOi += 1
                    CUSTPOSs.Add(OPO, CUSTPOi)
                End If
                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                Dim QTY As Int64 = Val(row.Item("QTY") & "")
                Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 IsNot Nothing Then
                    rowICTSTYC1.Item("OPEN_PICK_RSRV") = Val(rowICTSTYC1.Item("OPEN_PICK_RSRV") & "") + QTY
                End If
            Next
        End If

        Dim XLS_CREATED As Boolean = False

        If chk1Sheet.Checked Then
            Dim wsi As Integer = 0
            'Dim WJZ As Integer = dst.Tables("ICTQUOT2").Rows.Count

            Dim CODES As String = ""
            If opt1Sheet.Value = "S" Then
                CODES = "SUB_BODY_CODE"
            ElseIf opt1Sheet.Value = "FS" Then
                CODES = "FABRIC_CODE,SUB_BODY_CODE,STYLE_GROUP_CODE"
            ElseIf opt1Sheet.Value = "G" Then
                CODES = "STYLE_GROUP_CODE,FABRIC_CODE,SUB_BODY_CODE"
                ' CODES = "STYLE_GROUP_CODE"
                ' DGJ
            End If

            For Each rowSB As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ICTQUOT2").Select(Mid(sqlWB & sql0, 6)), Split(CODES, ",")).Select("")
                Dim SHEET_NAME As String = ""
                Dim sqlSB As String = ""
                For Each COLUMN_NAME As String In Split(CODES, ",")
                    Dim CODE_VALUE As String = rowSB.Item(COLUMN_NAME) & ""
                    SHEET_NAME &= "-" & CODE_VALUE
                    If CODE_VALUE = "" Then
                        sqlSB &= " and " & COLUMN_NAME & " IS NULL"
                    Else
                        sqlSB &= " and " & COLUMN_NAME & " = '" & CODE_VALUE & "'"
                    End If
                Next

                SHEET_NAME = Mid(SHEET_NAME, 2)

                If dst.Tables("ICTQUOT2").Select(Mid(sqlWB & sqlSB & sql0, 6)).Length > 0 Then
                    Dim worksheet As SpreadsheetGear.IWorksheet
                    If wsi = 0 Then
                        worksheet = workbook.Worksheets(0)
                    Else
                        worksheet = workbook.Worksheets.Add
                    End If
                    wsi += 1
                    If SHEET_NAME <> "" Then
                        worksheet.Name = SHEET_NAME
                    Else
                        worksheet.Name = "Unknown"
                    End If

                    Create_Excel_WorkSheet(worksheet, sqlWB & sqlSB & sql0)
                    XLS_CREATED = True
                End If
            Next
        Else
            If dst.Tables("ICTQUOT2").Select(Mid(sqlWB & sql0, 6)).Length > 0 Then
                Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
                worksheet.Name = "Style Info"
                Create_Excel_WorkSheet(worksheet, sqlWB & sql0)
                XLS_CREATED = True
            End If
        End If

        If XLS_CREATED Then
            Dim XLS_FILENAME As String = ""
            Dim success As Boolean = False

            ASCMAIN1.Progress("Now Saving Workbook")

            Do Until success
                Try
                    XLS_NO += 1
                    XLS_FILENAME = Absx1.txtFor("QUOTE_NO").Text
                    If SALES_DIVISION_CODE <> "" Then
                        XLS_FILENAME &= "-" & SALES_DIVISION_CODE
                    End If
                    XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".XLSX"
                    workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    'workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    RetVal = XLS_FILENAME
                    success = True
                Catch ex As Exception

                End Try
            Loop

            Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)
        End If

        ASCMAIN1.Progress("")
        Return RetVal
    End Function

    Private Function Create_Excel_Buyer() As String
        Dim RetVal As String = ""
        Me.Cursor = Cursors.WaitCursor

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Name = "Buyer Spreadsheet"
        ASCMAIN1.Progress("Now Creating Buyer Workbook", "")
        'Make Headers
        'Set Column widths
        Dim CWC() As String = Split("A, B, C, D, E, F,G,H,I,J,K, L,M, N, O", ",")
        Dim CWS() As String = Split("6,20,40,10,40,10,8,5,5,5,5,15,5,10,25", ",")
        Dim CWA() As String = Split("C, L, L, L, L, C,C,C,C,C,L, L,C, C, C", ",")
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
            If Trim(CWA(CWCi)) = "C" Then
                worksheet.Cells(Trim(CWC(CWCi)) & "1").HorizontalAlignment = SpreadsheetGear.HAlign.Center
            Else
                worksheet.Cells(Trim(CWC(CWCi)) & "1").HorizontalAlignment = SpreadsheetGear.HAlign.Left
            End If
        Next

        With worksheet
            'Paint cell colors, borders and col width
            With .Cells("A1:O2")
                .Interior.Color = SpreadsheetGear.Colors.LightBlue
                .Font.Bold = True
            End With
            With .Cells("A1:C1")
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            End With
            With .Cells("A2:O2")
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            End With
            With .Cells("D1:O1")
                .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            End With

            'Fill In the Captions.
            .Cells("A1").Value = "Categ"
            .Cells("B1").Value = "Style"
            .Cells("C1").Value = "Description"
            .Cells("F1").Value = "Start Ship"
            .Cells("G1").Value = "Consol" & vbCrLf & "cxl"
            .Cells("H1").Value = "LP"
            .Cells("I1").Value = "YJ"
            .Cells("J1").Value = "Marsh"
            .Cells("L1").Value = "Units"
            .Cells("M1").Value = "Cost"
            .Cells("N1").Value = "Comp"
            .Cells("O1").Value = "Comments"
            .Cells("C2").Value = "Group A"
            .Cells("O2").Value = Now().ToShortDateString
            .Cells("O2").HorizontalAlignment = SpreadsheetGear.HAlign.Right

        End With

        'Fill Rows
        'worksheet.Cells("D1").EntireColumn.NumberFormat = SpreadsheetGear.NumberFormatType.Text

        Dim curRow As Int64 = 3
        For Each rowSB As DataRow In dst.Tables.Item("ICTSTYC1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowSB.Item("STYLE_CODE").ToString & String.Empty
            Dim fltrICTQUOT2 As String = String.Format("STYLE_CODE_PLM = '{0}'", STYLE_CODE)
            Dim rowICTQUOT2 As DataRow = dst.Tables.Item("ICTQUOT2").Select(fltrICTQUOT2).FirstOrDefault
            worksheet.Cells("B" & curRow.ToString).Value = STYLE_CODE
            worksheet.Cells("C" & curRow.ToString).Value = rowICTQUOT2.Item("STYLE_DESC").ToString & String.Empty
            worksheet.Cells("D" & curRow.ToString).NumberFormat = "@"
            worksheet.Cells("D" & curRow.ToString).Value = (rowSB.Item("COLOR_CODE").ToString & String.Empty).ToString
            worksheet.Cells("E" & curRow.ToString).Value = rowSB.Item("STYLE_COLOR_DESC").ToString & String.Empty
            worksheet.Cells("A" & curRow.ToString & ":" & "O" & curRow.ToString).Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            Dim START_SHIP As Date = CDate("01/01/1900")
            Dim UNITS As Int64 = 0
            For i As Int64 = 0 To 3
                If IsDate(rowSB.Item("DTE" & i).ToString & String.Empty) And Val(rowSB.Item("QTY_AVA" & i).ToString & String.Empty) > 0 Then
                    If CDate(rowSB.Item("DTE" & i).ToString & String.Empty) > START_SHIP Then
                        START_SHIP = CDate(rowSB.Item("DTE" & i).ToString & String.Empty)
                        UNITS += Val(rowSB.Item("QTY_AVA" & i).ToString & String.Empty)
                    End If
                End If
            Next
            If START_SHIP <> CDate("01/01/1900") Then
                'Start Ship (F)
                worksheet.Cells("F" & curRow.ToString).Value = START_SHIP.ToShortDateString
                'Units (L)
                worksheet.Cells("L" & curRow.ToString).Value = UNITS
            End If
            curRow += 1
        Next

        'Show Workbook
        Dim XLS_FILENAME As String = Format(XLS_NO, "000") & ".XLSX"
        workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

        Return RetVal
    End Function

    Private Function Create_Excel_BuyerChart() As String
        Dim RetVal As String = ""
        Me.Cursor = Cursors.WaitCursor

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Name = "Buyer Spreadsheet"
        ASCMAIN1.Progress("Now Creating Buyer Workbook", "")
        'Make Headers
        worksheet.Cells("A1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("B1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("C1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("D1").EntireColumn.ColumnWidth = 0
        If chkShowFactoryBC.Checked Then
            worksheet.Cells("E1").EntireColumn.ColumnWidth = 13.17
        Else
            worksheet.Cells("E1").EntireColumn.ColumnWidth = 0
        End If
        worksheet.Cells("F1").EntireColumn.ColumnWidth = 20.33
        worksheet.Cells("G1").EntireColumn.ColumnWidth = 17.83
        worksheet.Cells("H1").EntireColumn.ColumnWidth = 27.33
        worksheet.Cells("I1").EntireColumn.ColumnWidth = 17.33
        worksheet.Cells("J1").EntireColumn.ColumnWidth = 19.83
        worksheet.Cells("K1").EntireColumn.ColumnWidth = 29.83
        If chkShowCountry.Checked Then
            worksheet.Cells("L1").EntireColumn.ColumnWidth = 14.83
        Else
            worksheet.Cells("L1").EntireColumn.ColumnWidth = 0
        End If
        worksheet.Cells("M1").EntireColumn.ColumnWidth = 15.83
        worksheet.Cells("N1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("O1").EntireColumn.ColumnWidth = 12
        worksheet.Cells("P1").EntireColumn.ColumnWidth = 13
        worksheet.Cells("Q4").EntireColumn.ColumnWidth = 12.83
        worksheet.Cells("E1: J1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells("K1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
        worksheet.Cells("L1: M1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells("O1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        worksheet.Cells("P1: Q1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells("J1").EntireColumn.WrapText = True
        worksheet.Cells("K1").EntireColumn.WrapText = True
        worksheet.Cells("A1").RowHeight = 12
        worksheet.Cells("A2").RowHeight = 48.75
        worksheet.Cells("A3").RowHeight = 12
        worksheet.Cells("A4").RowHeight = 56.5
        worksheet.Cells("A4").Value = "Department Number"
        worksheet.Cells("B4").Value = "Season"
        worksheet.Cells("C4").Value = "Class"
        worksheet.Cells("D4").Value = "Category Number"
        If chkShowFactoryBC.Checked Then
            worksheet.Cells("E4").Value = "Factory"
        End If
        worksheet.Cells("F4").Value = "Brand"
        worksheet.Cells("G4").Value = "Size Ratio"
        worksheet.Cells("H4").Value = "Photo"
        worksheet.Cells("I4").Value = "Style Code"
        worksheet.Cells("J4").Value = "Product Description"
        worksheet.Cells("K4").Value = "Color"
        If chkShowCountry.Checked Then
            worksheet.Cells("L4").Value = "Country"
        Else
            worksheet.Cells("L4").Value = ""
        End If
        worksheet.Cells("M4").Value = "Start"
        worksheet.Cells("N4").Value = "TKM"
        worksheet.Cells("O4").Value = "Avail"
        worksheet.Cells("P4").Value = "Vandale Cost"
        worksheet.Cells("Q4").Value = ""
        worksheet.Cells("F2").Value = "Buyer Chart"
        With worksheet.Cells("E2:Q2")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Font.Bold = True
            .Font.Size = 18
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.LightCyan
        End With
        With worksheet.Cells("A4:D4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.LightGray
        End With
        With worksheet.Cells("E4:N4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
        End With
        With worksheet.Cells("O4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.Yellow
        End With
        With worksheet.Cells("P4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
        End With
        With worksheet.Cells("Q4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.Yellow
        End With

        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

        Dim curRow As Int64 = 5
        For Each rowSB As DataRow In dst.Tables.Item("ICTSTYC1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowSB.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowSB.Item("COLOR_CODE").ToString & String.Empty
            Dim sql As New System.Text.StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT")
            sql.AppendLine("ST1.FACTORY_CODE,")
            sql.AppendLine("CN1.COUNTRY_NAME,")
            sql.AppendLine("SD1.SALES_DIVISION_NAME")
            sql.AppendLine("FROM ICTSTYL1 ST1, SOTSDIV1 SD1, TATCNTRY CN1")
            sql.AppendLine("WHERE ST1.SALES_DIVISION_CODE = SD1.SALES_DIVISION_CODE")
            sql.AppendLine("AND ST1.COUNTRY_CODE = CN1.COUNTRY_CODE (+)")
            sql.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
            Dim tblSTYLE As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
            Dim FACTORY_CODE As String = ""
            Dim COUNTRY_NAME As String = ""
            Dim SALES_DIVISION_NAME As String = ""
            If tblSTYLE.Rows.Count = 1 Then
                FACTORY_CODE = tblSTYLE.Rows(0).Item("FACTORY_CODE").ToString & String.Empty
                COUNTRY_NAME = tblSTYLE.Rows(0).Item("COUNTRY_NAME").ToString & String.Empty
                SALES_DIVISION_NAME = tblSTYLE.Rows(0).Item("SALES_DIVISION_NAME").ToString & String.Empty
            End If
            Dim STYLE_COLOR_DESC As String = rowSB.Item("STYLE_COLOR_DESC").ToString & String.Empty
            Dim fltrICTQUOT2 As String = String.Format("STYLE_CODE_PLM = '{0}'", STYLE_CODE)
            Dim rowICTQUOT2 As DataRow = dst.Tables.Item("ICTQUOT2").Select(fltrICTQUOT2).FirstOrDefault
            Dim STYLE_DESC As String = rowICTQUOT2.Item("STYLE_DESC").ToString & String.Empty
            Dim SIZE_SCALE As String = rowICTQUOT2.Item("SIZE_SCALE").ToString & String.Empty
            Dim IMAGE_NAME As String = rowICTQUOT2.Item("IMAGE_NAME") & ""
            Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME
            Dim HasImage As Boolean = False
            Dim imageStyle As System.Drawing.Image = Nothing
            If My.Computer.FileSystem.FileExists(imageFileStyle) Then
                imageStyle = System.Drawing.Image.FromFile(imageFileStyle)
                HasImage = True
            End If
            worksheet.Cells("A" & curRow.ToString & ":" & "Q" & curRow.ToString).Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            worksheet.Cells("A" & curRow.ToString).RowHeight = 100.5
            worksheet.Cells("E" & curRow.ToString).Value = FACTORY_CODE
            worksheet.Cells("F" & curRow.ToString).Value = SALES_DIVISION_NAME
            worksheet.Cells("G" & curRow.ToString).Value = SIZE_SCALE
            If HasImage Then
                Dim leftStyle As Integer = windowInfoStyle.ColumnToPoints(7)
                Dim topStyle As Integer = windowInfoStyle.RowToPoints(curRow - 1) + 0.1
                Dim WidthStyle As Integer = 100
                Dim HeightStyle As Integer = 99
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle + 20, topStyle + 1, WidthStyle, HeightStyle)
            End If
            worksheet.Cells("I" & curRow.ToString).Value = STYLE_CODE
            worksheet.Cells("J" & curRow.ToString).Value = STYLE_DESC
            worksheet.Cells("K" & curRow.ToString).Value = COLOR_CODE & " - " & STYLE_COLOR_DESC
            worksheet.Cells("L" & curRow.ToString).Value = COUNTRY_NAME
            Dim TOT_AVAIL As Int64 = 0
            Dim DATES As New System.Text.StringBuilder With {.Length = 0}
            For w As Int64 = 0 To 4
                Dim THIS_AVAIL As Int64 = Val(rowSB.Item("QTY_AVA" & w).ToString & String.Empty)
                If THIS_AVAIL > 0 Then
                    TOT_AVAIL = TOT_AVAIL + THIS_AVAIL
                End If
                Dim THIS_DATE As String = rowSB.Item("DTE" & w).ToString & String.Empty
                If IsDate(THIS_DATE) Then
                    DATES.AppendLine(Format(CDate(THIS_DATE), "MM/dd/yy"))
                End If
            Next
            Dim DATES_STRING As String = ""
            If DATES.ToString.Length > 2 Then
                DATES_STRING = DATES.ToString.Substring(0, DATES.Length - 2)
            End If
            With worksheet.Cells("M" & curRow.ToString)
                .Value = DATES_STRING
                .Font.Color = SpreadsheetGear.Colors.Red
            End With
            With worksheet.Cells("O" & curRow.ToString)
                .Value = TOT_AVAIL
                .NumberFormat = "###,##0"
            End With
            With worksheet.Cells("P" & curRow.ToString)
                '.Value = 3.4 'Get Vandale Cost Here
                .NumberFormat = "$###,##0.00"
            End With
            With worksheet.Cells("Q" & curRow.ToString)
                ' .Value = 3.3 'Get TKMAX OFFER Here
                .NumberFormat = "$###,##0.00"
                .Interior.Color = SpreadsheetGear.Colors.Yellow
                .Font.Color = SpreadsheetGear.Colors.Red
                .VerticalAlignment = SpreadsheetGear.VAlign.Center
            End With
            curRow += 1
        Next

        'Show Workbook
        Dim XLS_FILENAME As String = "5000"
        Dim success As Boolean = False
        Dim RPT_PREFIX As String = Absx1.txtFor("QUOTE_NO").Text
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = RPT_PREFIX & "_" & Format(XLS_NO, "000") & ".XLSX"
                workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                RetVal = XLS_FILENAME
                success = True
            Catch ex As Exception
                If XLS_NO > 5000 Then
                    success = True
                End If
            End Try
        Loop
        If XLS_FILENAME = "5000" Then
            MsgBox("Reports In Temp Folder Exceeded", vbCritical, "Log Out Of ABS And Get Back In")
        Else
            Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)
        End If

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

        Return RetVal
    End Function

    Sub Create_Excel_WorkSheet(
                              worksheet As SpreadsheetGear.IWorksheet,
                              Optional sqlWB As String = "")

        Dim BegAlloPeriod As Int64 = CalculateBegAlloPeriod()
        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            '  Stop
            IMAGE_FOLDER = "C:\Users\wjz\Desktop\Clients\VAN\images"
        End If

        Dim CAP_QTY As Int64 = Val(numCapQty.Value & "")

        'Dim xnames As SpreadsheetGear.INames = worksheet.Workbook.Names
        'xnames.Add("Jan", "='2005 Sales'!$B$2:$B$4")
        'xnames.Add("Feb", "='2005 Sales'!$C$2:$C$4")

        'Dim rr As SpreadsheetGear.IRange = worksheet.Cells(0, 0)
        'Dim aa As SpreadsheetGear.IAreas = worksheet.Cells(0, 0)


        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange


        worksheet.Cells("A1:Z1").EntireColumn.Font.Size = 16


        Dim CX As Integer = 0
        Dim RX As Integer = 0

        Dim I As Integer = 0
        I += 4

        Dim CWC() As String = Split("A,B, C,D,E,F,G,H,I,J,K,L, M", ",")
        Dim CWS() As String = Split("1,1,40,6,6,10,6,6,6,6,6,6,20", ",")
        If optPP.Value & "" = "4/5" Then
            CWS(2) = 45
        End If
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
        Next

        'range.EntireColumn.NumberFormat = SpreadsheetGear.NumberFormatType.Text
        'range.EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center

        worksheet.Cells(0, 0).EntireColumn.Hidden = True
        worksheet.Cells(0, 1).EntireColumn.Hidden = True

        Dim COL0 As Integer = 6 + 6

        Dim COL As Integer = COL0

        Dim ColVisible(4) As Boolean
        ColVisible(0) = True
        ColVisible(1) = (tkb1.Value <= 2)
        ColVisible(2) = (tkb1.Value <= 1)
        ColVisible(3) = (tkb1.Value <= 0)
        ColVisible(4) = chkBeyond.Checked

        Dim ColColors(4) As SpreadsheetGear.Color
        ColColors(0) = SpreadsheetGear.Colors.ForestGreen
        ColColors(1) = SpreadsheetGear.Colors.Magenta
        ColColors(2) = SpreadsheetGear.Colors.DodgerBlue
        ColColors(3) = SpreadsheetGear.Colors.OrangeRed
        ColColors(4) = SpreadsheetGear.Colors.Purple

        For iCol As Integer = BegAlloPeriod To 4
            If ColVisible(iCol) Then
                COL += 1
                With worksheet.Cells(I - 1, COL)
                    .ColumnWidth = 9
                    .EntireColumn.NumberFormat = "#,##0"
                    .EntireColumn.Font.Color = ColColors(iCol)
                    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "Avail"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With
                COL += 1
                With worksheet.Cells(I - 1, COL)
                    .ColumnWidth = 8
                    .EntireColumn.NumberFormat = "MM/dd"
                    .EntireColumn.Font.Color = ColColors(iCol)
                    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Value = "Date"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With
                With worksheet.Cells(I, COL)
                    .Font.Size = 12
                    Dim nth As String = ""
                    Select Case iCol - BegAlloPeriod
                        Case 0
                            nth = "1st"
                        Case 1
                            nth = "2nd"
                        Case 2
                            nth = "3rd"
                        Case 3
                        Case Else
                            nth = CStr(iCol) & "th"
                    End Select

                    .Value = nth & " Del"
                End With
            End If
        Next

        COL += 1
        With worksheet.Cells(I - 1, COL)
            .ColumnWidth = 9
            .EntireColumn.NumberFormat = "#,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .Value = "All"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With

        If chkShowPOs.Checked Then
            If CUSTPOSs.Count > 0 Then
                For Each CUSTPOkey As String In CUSTPOSs.Keys
                    Dim CUSTPOi As Integer = CUSTPOSs(CUSTPOkey)
                    Dim CUSTPO() As String = Split(CUSTPOkey, vbTab)
                    worksheet.Cells(I - 1, COL + CUSTPOi).Value = CUSTPO(1) ' PO
                    If CUSTPO(0) = "R" Then
                        worksheet.Cells(I - 1, COL + CUSTPOi).Font.Color = SpreadsheetGear.Colors.Red
                    End If
                    worksheet.Cells(I - 1, COL + CUSTPOi).HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    If CUSTPOi Mod 2 = 0 Then
                        worksheet.Cells(I - 1, COL + CUSTPOi).EntireColumn.Interior.Color = SpreadsheetGear.Colors.LavenderBlush
                    Else
                        worksheet.Cells(I - 1, COL + CUSTPOi).EntireColumn.Interior.Color = SpreadsheetGear.Colors.MistyRose
                    End If
                Next
            End If
        End If


        With worksheet.Cells(I, 0, I, COL)
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        If chkShowCost.Checked Then
            With worksheet.Cells(2, 3)
                .Value = "INTERNAL USE"
                .Font.Color = SpreadsheetGear.Colors.Red
                .Font.Bold = True
                .Font.Size = 16
            End With
        End If





        Dim I0 As Integer = 0
        Dim IA As Integer = 0

        Dim sqlw As String = sqlWB
        'If chkShowSelectedOnly.Checked Then
        '    sqlw &= " and SELECTED = '1'"
        'End If
        If sqlw <> "" Then
            sqlw = Mid(sqlw, 6)
        End If

        Dim RT(5) As String

        Dim ROW0 As Integer = I

        Dim style_count As Integer = 0
        Dim pages As Integer = 0

        For Each row As DataRow In dst.Tables("ICTQUOT2").Select(sqlw, "STYLE_CODE_PLM") ' SEQ")

            If opt1Sheet.Value = "G" And chk1Sheet.Checked = True Then


                With worksheet.Cells(2, 7)
                    .Value = "Group " & row.Item("STYLE_GROUP_CODE")
                    .Font.Color = SpreadsheetGear.Colors.Red
                    .Font.Bold = True
                    .Font.Size = 16
                End With
            End If

            STYLE_CODE = row.Item("STYLE_CODE_PLM")
            rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)

            If BegAlloPeriod <> 0 Then
                Dim totalAvaliable As Int64 = 0
                For Each rowA As DataRow In dst.Tables.Item("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                    For tcnt As Integer = BegAlloPeriod To 4
                        totalAvaliable += Val(rowA.Item(String.Format("QTY_AVA{0}", tcnt.ToString)) & "")
                    Next
                Next
                If totalAvaliable = 0 Then
                    Continue For
                End If
            End If

            ASCMAIN1.Progress("-", STYLE_CODE)

            I += 1
            I0 = I

            COL = COL0

            worksheet.Cells(I, COL - 1).Value = "Color"
            worksheet.Cells(I, COL - 0).Value = "Description"

            For iCol As Integer = BegAlloPeriod To 4
                If ColVisible(iCol) Then
                    COL += 1
                    With worksheet.Cells(I, COL)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Left
                        If iCol = 0 Then
                            If Format(dte0.DateTime, "MM/dd/yy") = Format(Now, "MM/dd/yy") Then
                                .Value = "At Once"
                                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                            Else
                                .Value = "At Once-" & Format(dte0.DateTime, "MM/dd")
                            End If

                        End If
                        'If iCol = 1 Then .Value = Format(dte1.Value, "MM/dd/yy")
                        'If iCol = 2 Then .Value = Format(dte2.Value, "MM/dd/yy")
                        'If iCol = 3 Then .Value = Format(dte3.Value, "MM/dd/yy")

                        If iCol = 1 Then .Value = Format(dte0.DateTime.AddDays(1), "MM/dd") & "-" & Format(dte1.Value, "MM/dd")
                        If iCol = 2 Then .Value = Format(dte1.DateTime.AddDays(1), "MM/dd") & "-" & Format(dte2.Value, "MM/dd")
                        If iCol = 3 Then .Value = Format(dte2.DateTime.AddDays(1), "MM/dd") & "-" & Format(dte3.Value, "MM/dd")

                        If iCol = 4 Then .Value = "Beyond"

                        If iCol = 0 Or iCol = 4 Then
                        Else
                            .NumberFormat = "MM/dd/yy"
                        End If

                    End With

                    COL += 1
                    'With worksheet.Cells(I, COL)
                    '    .NumberFormat = "MM/dd"
                    '    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    'End With
                End If
            Next

            COL += 1
            With worksheet.Cells(I, COL)
                .Value = "Total"
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            'If chkShowLastShip.Checked Then
            '    COL += 1
            '    With worksheet.Cells(I, COL)
            '        .Value = "Last Ship"
            '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            '        .ColumnWidth = 16
            '        .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            '        .Borders.Color = SpreadsheetGear.Colors.Black
            '        'range = worksheet.Cells(I, COL0 - 1, I, COL)
            '        'interior = range.Interior
            '        'interior.Color = SpreadsheetGear.Colors.Gold
            '    End With
            'End If
            If chkShowLastRcd.Checked Then
                COL += 1
                With worksheet.Cells(I, COL)
                    .Value = "Last Rcvd"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .ColumnWidth = 16
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders.Color = SpreadsheetGear.Colors.Black
                End With
            End If

            range = worksheet.Cells(I, COL0 - 1, I, COL)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.Gold

            If chkShowPOs.Checked Then
                If CUSTPOSs.Count > 0 Then
                    For Each CUSTPOkey As String In CUSTPOSs.Keys
                        Dim CUSTPOi As Integer = CUSTPOSs(CUSTPOkey)
                        Dim CUSTPO() As String = Split(CUSTPOkey, vbTab)
                        With worksheet.Cells(I, COL + CUSTPOi)
                            .NumberFormat = "MM/dd"
                            .Value = Format(CDate(CUSTPO(3)), "MM/dd") ' Cancel Date
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With

                    Next
                End If
            End If

            I += 1


            'Dim SourceImagePath As String
            'Dim DestinationPath As String
            'Dim original As Image = Image.FromFile(SourceImagePath)
            'Dim resized As Image = ResizeImage(original, New Size(1024, 768))
            'SaveImageWithQuality(resized, DestinationPath, 75L)


            Dim IMAGE_NAME As String = row.Item("IMAGE_NAME") & ""


            Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME


            'Dim SourceImagePath As String = IMAGE_FOLDER & "\" & IMAGE_NAME
            'Dim imageFileStyle As String = IMAGE_FOLDER & "\resized\" & IMAGE_NAME
            'Dim original As Image = Image.FromFile(SourceImagePath)
            'Dim resized As Image = ResizeImage(original, New Size(1024, 1024))
            'SaveImageWithQuality(resized, imageFileStyle, 75L)



            Dim ImageRows As Integer = 0
            Dim ImageRowsBig As Integer = 0

            If Not chkNoPictures.Checked And IMAGE_NAME <> "" _
                AndAlso My.Computer.FileSystem.FileExists(imageFileStyle) Then

                Dim widthStyle As Double
                Dim heightStyle As Double

                Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
                Try
                    If optPP.Value = "4/5" Then
                        widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution / 3
                        heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution / 3
                    Else
                        widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution / 4
                        heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution / 4
                    End If


                Finally
                    imageStyle.Dispose()
                End Try

                ' Calculate the left and top placement of the picture by converting 
                ' row and column coordinates to points.  Use fractional values to 
                ' get coordinates anywhere in between row and column boundaries.
                Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

                Dim col_adj As Decimal = 0
                If optPP.Value = "4/5" Then
                    If heightStyle > widthStyle Then
                        col_adj = 0.3
                    Else
                        col_adj = 0.05
                    End If
                Else
                    If heightStyle > widthStyle Then
                        col_adj = 0.15
                    Else
                        col_adj = 0.05
                    End If
                End If


                Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0) + col_adj
                Dim topStyle As Double = windowInfoStyle.RowToPoints(I - 1) + 0.1 ' 1.5)

                ImageRows = windowInfoStyle.PointsToRow(heightStyle)

                ' Add the picture from file.

                If optPP.Value & "" <> "2" Then
                    worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
                Else

                    ImageRowsBig = 14

                    widthStyle = widthStyle * 2
                    heightStyle = heightStyle * 2

                    Dim column_adj As Integer = 0
                    If widthStyle < heightStyle Then
                        column_adj = 2
                    End If

                    leftStyle = windowInfoStyle.ColumnToPoints(4 + column_adj + 0.1) ' 1.5)
                    topStyle = windowInfoStyle.RowToPoints(I + 7 - 0.5) ' 1.5)

                    worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
                End If

            End If

            CX = 1

            With worksheet.Cells(I - 1, 3)
                .Value = "'" & row.Item("STYLE_CODE_PLM")
                .Font.Color = SpreadsheetGear.Colors.Purple
                .Font.Size = 24
                .Font.Bold = True
            End With

            CX = 3

            worksheet.Cells(I + 2, CX).Value = "Case Qty"
            ' worksheet.Cells(I + 2, CX).Value = ""
            'If Not chkOmitPrice.Checked And Val(row.Item("STYLE_PRICE") & "") <> 0 Then worksheet.Cells(I + 1, CX + 4).Value = "Price"
            'If chkShowRetail.Checked And Val(row.Item("STYLE_RETAIL") & "") <> 0 Then worksheet.Cells(I + 2, CX + 4).Value = "MSRP"
            If chkShowRetail.Checked And Val(row.Item("STYLE_RETAIL") & "") <> 0 Then worksheet.Cells(I + 7, CX).Value = "MSRP"

            range = worksheet.Cells(I + 1, 3, I + 2, 4)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.LightGray
            'range = worksheet.Cells(I + 1, 3, I + 3, 4)

            range = worksheet.Cells(I + 7, 3, I + 7, 4)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.LightGray

            range = worksheet.Cells(I + 1, 3, I + 1, 5)
            range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous

            range = worksheet.Cells(I + 2, 3, I + 2, 5)
            range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous

            'range = worksheet.Cells(I + 1, 3 + 4, I + 1, 4 + 5)
            'range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            'range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            'range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            'range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous

            range = worksheet.Cells(I + 7, 3, I + 7, 5)
            range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            'range.Borders.Color = SpreadsheetGear.Colors.Black

            Dim FACTORY_CODE As String = rowICTSTYL1.Item("FACTORY_CODE") & String.Empty
            If FACTORY_CODE.Length = 0 Then
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT MAX(FACTORY_CODE) AS FACTORY_CODE")
                SQLS.AppendLine("FROM POTORDR1")
                SQLS.AppendLine("WHERE PO_ORDER_NO")
                SQLS.AppendLine("IN")
                SQLS.AppendLine("(")
                SQLS.AppendLine("SELECT MAX(P1.PO_ORDER_NO)")
                SQLS.AppendLine("FROM POTORDR1 P1, POTORDR2 P2")
                SQLS.AppendLine("WHERE P1.PO_ORDER_NO = P2.PO_ORDER_NO")
                SQLS.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", rowICTSTYL1.Item("STYLE_CODE") & String.Empty))
                SQLS.AppendLine("AND NVL(P1.FACTORY_CODE, 'NULLRECORD') <> 'NULLRECORD'")
                SQLS.AppendLine(")")
                ASCMAIN1.sql = SQLS.ToString()
                Dim RETVAL As String = ASCDATA1.GetDataValue
                If RETVAL.Length > 0 Then
                    FACTORY_CODE = RETVAL
                End If
            End If

            If chkShowFactory.Checked And FACTORY_CODE.Length <> 0 Then
                worksheet.Cells(I + 9, CX).Value = "Factory"
                worksheet.Cells(I + 9, CX + 2).Value = FACTORY_CODE
                range = worksheet.Cells(I + 9, 3, I + 9, 4)
                interior = range.Interior
                interior.Color = SpreadsheetGear.Colors.LightGray
                range = worksheet.Cells(I + 9, 3, I + 9, 5)
                range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End If

            CX = 5
            worksheet.Cells(I, CX - 2).Value = row.Item("STYLE_DESC")
            '  worksheet.Cells(I, CX - 2).Font.Size = 14
            ' worksheet.Cells(I + 0, CX).Value = row.Item("STYLE_DESC2")

            ' worksheet.Cells(I + 0, CX).Value = rowICTSTYL1
            worksheet.Cells(I + 2, CX).Value = rowICTSTYL1.Item("CARTON_PACK_QTY")

            'If Not chkOmitPrice.Checked And Val(row.Item("STYLE_PRICE") & "") <> 0 Then
            '    With worksheet.Cells(I + 1, CX + 4)
            '        .Value = row.Item("STYLE_PRICE")
            '        .NumberFormat = "$#,##0.00"
            '        .Font.Color = SpreadsheetGear.Colors.Green
            '        .Font.Bold = True
            '    End With
            'End If

            If chkShowRetail.Checked And Val(row.Item("STYLE_RETAIL") & "") <> 0 Then
                With worksheet.Cells(I + 7, CX)
                    .Value = row.Item("STYLE_RETAIL")
                    .NumberFormat = "$#,##0.00"
                    .Font.Color = SpreadsheetGear.Colors.Blue
                End With
            End If

            Dim SZMAX As Integer = 0
            Dim SZTOT As Integer = 0

            Dim rowICTSTYLX As DataRow = dst.Tables("ICTSTYLX").Rows.Find(STYLE_CODE)

            Dim OFFSET As Integer = -2

            Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
            If Not IsNothing(rowICTSTYLS) Then
                For SZ As Integer = 1 To 8
                    Dim SIZE As String = rowICTSTYLS.Item("SIZE_" & Format(SZ, "00")) & String.Empty
                    Dim UNITS As String = CStr(Val(rowICTSTYLS.Item("QTY_" & Format(SZ, "00")) & ""))
                    If SIZE.Length = 0 Then
                        Exit For
                    End If
                    worksheet.Cells(I + 4, CX + OFFSET + SZ - 1).Value = SIZE
                    worksheet.Cells(I + 5, CX + OFFSET + SZ - 1).Value = UNITS
                    SZMAX = SZ
                Next
            End If
            'For SZ As Integer = 1 To 8
            '    'worksheet.Cells(I + 4, CX + OFFSET + SZ - 1).Value = rowICTSTYLX.Item("S" & CStr(SZ))
            '    'worksheet.Cells(I + 5, CX + OFFSET + SZ - 1).Value = rowICTSTYLX.Item("Q" & CStr(SZ))
            '    'If rowICTSTYLX.Item("S" & CStr(SZ)) & "" <> "" Then
            '    '    SZMAX = SZ
            '    '    SZTOT += Val(rowICTSTYLX.Item("Q" & CStr(SZ)) & "")
            '    'End If
            'Next

            If SZMAX <> 0 Then
                With worksheet.Cells(I + 4, CX + OFFSET - 1 + 1, I + 4, CX + OFFSET - 1 + SZMAX)
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Interior.Color = SpreadsheetGear.Colors.LightBlue
                    .Interior.TintAndShade = -1
                End With

                With worksheet.Cells(I + 5, CX + OFFSET - 1 + 1, I + 5, CX + OFFSET - 1 + SZMAX)
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Interior.Color = SpreadsheetGear.Colors.LightBlue
                End With

                'worksheet.Cells(I + 4, CX - 1).Value = SZTOT
                'worksheet.Cells(I + 4, CX - 1).HorizontalAlignment = SpreadsheetGear.HAlign.Right
                'worksheet.Cells(I + 1, CX).Value = SZTOT
                'worksheet.Cells(I + 1, 3).Value = "Inner"


                With worksheet.Cells(I + 4, CX + OFFSET + 0, I + 5, CX + OFFSET + SZMAX - 1)
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders.Color = SpreadsheetGear.Colors.Black

                    '.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With

            End If

            Dim T As String = ""

            Dim CI As Integer = 0
            'Dim sqlstatus As String = " and (ISNULL(QTY_AVA0,0) <> 0 or ISNULL(QTY_AVA1,0) <> 0 or ISNULL(QTY_AVA2,0) <> 0 or ISNULL(QTY_AVA3,0) <> 0 or ISNULL(QTY_AVA4,0) <> 0)"
            Dim sqlstatus As String = " and COUNT_COLOR = 1"
            For Each row2 As DataRow In dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & row.Item("STYLE_CODE_PLM") & "'" & sqlstatus, "COLOR_CODE")
                CI += 1
                COL = COL0

                worksheet.Cells(I + CI - 1, COL - 1).Value = "'" & row2.Item("COLOR_CODE")
                worksheet.Cells(I + CI - 1, COL - 0).Value = "'" & row2.Item("STYLE_COLOR_DESC")


                If chkShowCost.Checked Then
                    With worksheet.Cells(I + CI - 1, COL - 2) '  worksheet.Cells(I, CX + 5)

                        ASCMAIN1.sql = "Select STYLE_COST from (" & vbCrLf _
                            & "Select STYLE_COST from ICTCOSTA " & vbCrLf _
                            & "where (STYLE_CODE, COLOR_CODE) in (" & vbCrLf _
                            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                            & " from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & row2.Item("COLOR_CODE") & "'" _
                            & " and WHSE_QTY_ON_HAND > 0)" & vbCrLf _
                            & " order by OPS_YYYYPP DESC) where ROWNUM < 2"
                        Dim STYLE_COST As Decimal = Val(ASCDATA1.GetDataValue)

                        If STYLE_COST = 0 Then
                            ASCMAIN1.sql = "Select NVL(PO_COST_LANDED,PO_COST) STYLE_COST" & vbCrLf _
                                & " from (" & vbCrLf _
                                & " Select POTSHIP3.PO_SHIPMENT_NO, POTORDR2.PO_ORDER_NO, " & vbCrLf _
                                & " POTORDR2.PO_COST, POTSHIP3.PO_COST_LANDED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
                                & " from POTORDR2,POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
                                & " where POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' and POTORDR2.COLOR_CODE = '" & row2.Item("COLOR_CODE") & "'" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_LNO (+) = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & " order by POTSHIP3.PO_SHIPMENT_NO DESC, POTORDR2.PO_ORDER_NO DESC" & vbCrLf _
                                & ") where ROWNUM <2"
                            STYLE_COST = Val(ASCDATA1.GetDataValue)
                        End If

                        If STYLE_COST = 0 Then
                            STYLE_COST = Val(row.Item("STYLE_COST") & "")
                        End If
                        .Value = STYLE_COST
                        .NumberFormat = "$#,##0.00"
                        .Font.Size = 12
                        .Font.Color = SpreadsheetGear.Colors.Red
                    End With
                End If


                T = ""
                Dim CARRY_OVER As Integer = 0
                For iCOL As Integer = BegAlloPeriod To 4
                    If ColVisible(iCOL) Then
                        Dim QTY_AVA As Int64 = Val(row2.Item("QTY_AVA" & Format(iCOL, "0")) & "")
                        If QTY_AVA < 0 Then
                            CARRY_OVER += QTY_AVA
                            QTY_AVA = 0
                        Else
                            QTY_AVA += CARRY_OVER
                            CARRY_OVER = 0
                        End If
                        If CAP_QTY <> 0 And QTY_AVA > CAP_QTY Then QTY_AVA = CAP_QTY

                        COL += 1
                        If QTY_AVA <> 0 Then
                            worksheet.Cells(I + CI - 1, COL).Value = QTY_AVA
                        End If
                        T &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                        COL += 1
                        If QTY_AVA <> 0 Then
                            If iCOL = 0 Then
                                'If STYLE_CODE = "802294JS" Then Stop
                                If Format(row2.Item("DTE" & Format(iCOL, "0")), "MM/dd/yy") = Format(Now, "MM/dd/yy") Then
                                    'If Format(dte0.DateTime, "MM/dd/yy") = Format(Now, "MM/dd/yy") Then
                                    worksheet.Cells(I + CI - 1, COL).Value = "Now"
                                Else
                                    worksheet.Cells(I + CI - 1, COL).Value = row2.Item("DTE" & Format(iCOL, "0"))
                                End If
                            Else
                                worksheet.Cells(I + CI - 1, COL).Value = row2.Item("DTE" & Format(iCOL, "0"))
                            End If
                        End If
                    End If

                Next
                COL += 1
                worksheet.Cells(I + CI - 1, COL).Formula = "=" & Mid(T, 2)

                'If chkShowLastShip.Checked Then
                '    COL += 1
                '    Dim filterQ2 As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", row2.Item("STYLE_CODE"), row2.Item("COLOR_CODE"))
                '    Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Select(filterQ2).FirstOrDefault
                '    If Not IsNothing(rowICTSTYC1) Then
                '        If IsDate(rowICTSTYC1.Item("LAST_SHIP_DATE").ToString & String.Empty) Then
                '            Dim LAST_SHIPPED As Date = CDate(rowICTSTYC1.Item("LAST_SHIP_DATE").ToString & String.Empty)
                '            With worksheet.Cells(I + CI - 1, COL)
                '                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                '                .Value = Format(LAST_SHIPPED, "MM/dd/yy")
                '            End With

                '        End If
                '    End If
                'End If
                If chkShowLastRcd.Checked Then
                    COL += 1
                    Dim filterQ2 As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", row2.Item("STYLE_CODE"), row2.Item("COLOR_CODE"))
                    Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Select(filterQ2).FirstOrDefault
                    If Not IsNothing(rowICTSTYC1) Then
                        If IsDate(rowICTSTYC1.Item("LAST_RCD_DATE").ToString & String.Empty) Then
                            Dim LAST_SHIPPED As Date = CDate(rowICTSTYC1.Item("LAST_RCD_DATE").ToString & String.Empty)
                            With worksheet.Cells(I + CI - 1, COL)
                                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                                .Value = Format(LAST_SHIPPED, "MM/dd/yy")
                            End With
                        Else
                            Dim LAST_SHIPPED As String = rowICTSTYC1.Item("LAST_RCD_DATE").ToString & String.Empty
                            With worksheet.Cells(I + CI - 1, COL)
                                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                                .Value = LAST_SHIPPED
                            End With
                        End If
                    End If
                End If

                If chkShowPOs.Checked Then

                    If CUSTPOSs.Count > 0 Then
                        For Each rowSOTORDRC As DataRow In dst.Tables("SOTORDRC").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & row2.Item("COLOR_CODE") & "'")

                            Dim CUSTPOkey As String = rowSOTORDRC.Item("ORDR_TYPE") & vbTab _
                                              & rowSOTORDRC.Item("ORDR_CUST_PO") & vbTab _
                                              & Format(rowSOTORDRC.Item("ORDR_SHIP_DATE"), "MM/dd/yyyy") & vbTab _
                                              & Format(rowSOTORDRC.Item("ORDR_CANCEL_DATE"), "MM/dd/yyyy")
                            Dim CUSTPOi As Integer = CUSTPOSs(CUSTPOkey)

                            Dim QTY As Int64 = Val(rowSOTORDRC.Item("QTY") & "")
                            worksheet.Cells(I + CI - 1, COL + CUSTPOi).Value = QTY
                        Next
                    End If
                End If
            Next


            CI += 1
            COL = COL0

            worksheet.Cells(I - 1, COL - 1, I + CI - 1, COL - 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center

            worksheet.Cells(I + CI - 1, COL - 1).Value = "'" & "***"
            worksheet.Cells(I + CI - 1, COL - 0).Value = "'" & "Total"

            If Not chkOmitPrice.Checked And Val(row.Item("STYLE_PRICE") & "") <> 0 Then
                With worksheet.Cells((I + CI - 1) + 2, COL - 0)
                    .Value = row.Item("STYLE_PRICE")
                    .NumberFormat = "$#,##0.00"
                    .Font.Color = SpreadsheetGear.Colors.Black
                    .Font.Size = 20
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                End With
            End If

            T = ""
            For iCOL As Integer = BegAlloPeriod To 4
                If ColVisible(iCOL) Then
                    COL += 1
                    If CI = 1 Then ' NO COLORS
                        worksheet.Cells(I + CI - 1, COL).Value = 0
                    Else
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                    End If

                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")

                    T &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                    COL += 1
                End If
            Next
            COL += 1
            worksheet.Cells(I + CI - 1, COL).Formula = "=" & Mid(T, 2)

            RT(5) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")

            worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

            With worksheet.Cells(I, COL0 - 1, I + CI - 1, COL)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
            End With

            ' I += 1 ' to add a little more space

            I += ImageRowsBig

            Dim CJ As Integer = ImageRows ' - 1

            If CJ < 6 Then CJ = 6

            If CI > CJ Then
                I += CI
            Else
                I += CJ
            End If

            style_count += 1

            If (optPP.Value & "" <> "2" And ((I - 5) Mod 80) < ((I0 - 5) Mod 80)) Or (optPP.Value & "" = "2" And style_count = 3) Or (optPP.Value & "" = "4/5" And style_count >= 5) Or style_count >= 9 Then
                Dim R As SpreadsheetGear.IRange = worksheet.Cells(I0, 0).EntireRow
                worksheet.HPageBreaks.Add(R)
                style_count = 1
                pages += 1
            End If

            With worksheet.Cells(I0, 0, I + 1 - 1, COL)
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                '.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
        Next


        I += 2
        COL = COL0

        worksheet.Cells(I - 1, COL - 1).Value = "'" & "All"
        worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = BegAlloPeriod To 4
            If ColVisible(iCOL) Then
                COL += 1
                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)

                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                COL += 1
            End If
        Next
        COL += 1
        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(GT, 2)


        worksheet.Cells(I - 1, COL0 - 1, I - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray




        Dim H0 As Integer = 8 + 6

        worksheet.Cells(0, H0).Value = "Prep"
        worksheet.Cells(1, H0).Value = "By"
        worksheet.Cells(2, H0).Value = "XNo"

        worksheet.Cells(0, H0, 2, H0).Interior.Color = SpreadsheetGear.Colors.LightGray


        worksheet.Cells(0, H0 + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Left
        worksheet.Cells(0, H0 + 1).Value = Now
        worksheet.Cells(0, H0 + 1).NumberFormat = "MM/dd/yy" ' SpreadsheetGear.NumberFormatType.Date

        worksheet.Cells(1, H0 + 1).Value = ASCMAIN1.USER_ID
        worksheet.Cells(2, H0 + 1).Value = "'" & Mid(XNO, 5)

        With worksheet.Cells(0, H0, 2, H0 + 1)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        If Not chkNoBoxes.Checked Then
            COL = 0
            For iCol As Integer = BegAlloPeriod To 4
                If ColVisible(iCol) Then
                    COL += 1
                    With worksheet.Cells(ROW0, COL0 + (COL - 1) * 2 + 1, I, COL0 + (COL - 1) * 2 + 2)
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Font.Color = SpreadsheetGear.Colors.Black
                    End With
                End If
            Next
        End If

        'With worksheet.Cells(0, 11, 1, 13)
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .Font.Color = Nothing
        'End With

        Dim H1 As Integer = 11

        worksheet.Cells(0, 2).Value = "Quote Sheet"
        worksheet.Cells(0, 2).Font.Bold = True

        worksheet.Cells(0, H1).Value = "Quote"
        worksheet.Cells(1, H1).Value = "For"

        worksheet.Cells(0, H1, 2, H1).Interior.Color = SpreadsheetGear.Colors.LightGray

        worksheet.Cells(0, H1 + 1).Value = Format(Absx1.dteFor("QUOTE_DATE").Value, "MM/dd/yyyy")
        worksheet.Cells(0, H1 + 1).NumberFormat = "MM/dd/yy"
        worksheet.Cells(0, H1 + 2).Value = "'" & Absx1.txtFor("QUOTE_NO").Text
        worksheet.Cells(1, H1 + 1).Value = Absx1.txtFor("CUST_NAME").Text
        worksheet.Cells(1, H1 + 2).Value = Absx1.txtFor("CUST_CODE").Text


        With worksheet.Cells(0, H1, 2, H1 + 2)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With worksheet.Cells(3, 3)
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Size = 20
            .Font.Bold = True
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = Absx1.txtFor("QUOTE_DESC").Text
        End With

        With worksheet.PageSetup
            .TopMargin = 0.25
            .LeftMargin = 0.25
            .RightMargin = 0.25
            .BottomMargin = 0.25
            .FitToPagesWide = 1
            .FitToPagesTall = Nothing
            .PrintTitleRows = "A1:S5"

            .CenterFooter = "&P"
        End With

        Dim imageFile As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & "_EMAIL.PNG"

        ' Get the width and height of the picture in pixels and convert to
        ' points for use in the call to AddPicture.  This step is only
        ' necessary if the actual picture size is to be used and that size
        ' is unknown.  Another option is to calculate the width and height 
        ' in row and column coordinates in the same manner as left and top below.

        ' ''Dim width As Double
        ' ''Dim height As Double
        ' ''Dim imageCAD As System.Drawing.Image = System.Drawing.Image.FromFile(imageFile)
        ' ''Try
        ' ''    width = imageCAD.Width * 72.0 / imageCAD.HorizontalResolution / 1
        ' ''    height = imageCAD.Height * 72.0 / imageCAD.VerticalResolution / 1
        ' ''Finally
        ' ''    imageCAD.Dispose()
        ' ''End Try

        '' '' Calculate the left and top placement of the picture by converting 
        '' '' row and column coordinates to points.  Use fractional values to 
        '' '' get coordinates anywhere in between row and column boundaries.
        ' ''Dim windowInfo As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo
        ' ''Dim left As Double = windowInfo.ColumnToPoints(3 + 0.15)
        ' ''Dim top As Double = windowInfo.RowToPoints(0) ' 1.5)

        ' ''windowInfo.ScrollRow = 0
        ' ''windowInfo.SplitRows = 5
        ' ''windowInfo.FreezePanes = True

        '' '' Add the picture from file.
        ' ''worksheet.Shapes.AddPicture(imageFile, left, top, width, height)
        '' '  dgj

        ' Auto size all worksheet columns which contain data
        ' worksheet.UsedRange.Columns.AutoFit()

        If optPP.Value & "" = "2" Then
            worksheet.Cells("C1").EntireColumn.Hidden = True
        End If
    End Sub

    Private Function CalculateBegAlloPeriod() As Long
        Dim RetVal As Int64 = 0
        If chkALLOSTDT.Checked Then
            Select Case cboStartPeriod.Text
                Case "Now"
                    RetVal = 0
                Case "1st"
                    RetVal = 1
                Case "2nd"
                    RetVal = 2
                Case "3rd"
                    RetVal = 3
            End Select
        End If
        Return RetVal
    End Function

    Sub Print_Style_Sheet(eItemKey As String, Optional STYLE_CODE As String = "")
        Dim ListPDFSheets As New List(Of String)
        Dim MISSING_IMAGES As New List(Of String)

        Dim BegAlloPeriod As Int64 = CalculateBegAlloPeriod()

        Synch_TABLE_NAME("ICTQUOT1")

        Dim blnShowSelected As Boolean = False
        If Not chkShowSelectedOnly.Checked Then
            blnShowSelected = True
            chkShowSelectedOnly.Checked = True
        End If

        RESEQ()

        'For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
        '    FetchImage(row)
        'Next

        ' COPYING THE SAME LOGIC USED FOR EXCEL
        For Each row As DataRow In dst.Tables("ICTSTYC1").Select("")
            row.Item("OPEN_PICK_RSRV") = 0
            row.Item("SKIP_COLOR") = "0"
        Next

        Dim EXCUDE_FUTURE As String = ""
        If chkExcudeFutureWhenZero.Checked Then
            Dim BegPeriod As Int64 = 0
            Select Case tkb1.Value
                Case 3
                    BegPeriod = 1
                Case 2
                    BegPeriod = 2
                Case 1
                    BegPeriod = 3
                Case 0
                    BegPeriod = 4
            End Select

            For i As Integer = BegPeriod To 4
                EXCUDE_FUTURE = EXCUDE_FUTURE & String.Format(" AND ISNULL(QTY_AVA{0},0)=0", i)
            Next
            If EXCUDE_FUTURE.Length > 0 Then
                EXCUDE_FUTURE = " AND (" & EXCUDE_FUTURE.Substring(5, EXCUDE_FUTURE.Length - 5) & ")"
            End If
            Dim COUNT_COLOR_EXCL As String = String.Format("IIF((QTY_AVA>={0} OR ISNULL(OPEN_PICK_RSRV,0)>0) {1},1,0)", Val(numMinQty.Value & ""), EXCUDE_FUTURE)
            dst.Tables("ICTSTYC1").Columns("COUNT_COLOR").Expression = COUNT_COLOR_EXCL
        Else
            dst.Tables("ICTSTYC1").Columns("COUNT_COLOR").Expression = String.Format(COUNT_COLOR, Val(numMinQty.Value & ""))
        End If
        'dst.Tables("ICTSTYC1").Columns("COUNT_COLOR").Expression = String.Format(COUNT_COLOR, Val(numMinQty.Value & ""))

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
        FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")

        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
            'If row.Item("IMAGE") Is DBNull.Value Then
            '    row.Item("SELECTED") = "0"
            'End If
            Dim STYLE_CODE_PLM As String = row.Item("STYLE_CODE_PLM")
            'If STYLE_CODE_PLM = "500498AVR" And ASCMAIN1.Running_in_VS Then Stop

            If Not My.Computer.FileSystem.FileExists(FOLDER_NAME & row.Item("IMAGE_NAME")) Then
                row.Item("SELECTED") = "0"
                MISSING_IMAGES.Add(STYLE_CODE_PLM)
            End If
            If Val(row.Item("COUNT_COLOR") & "") > 0 Then
            Else
                If Not chkShowZero.Checked Then
                    row.Item("SELECTED") = "0"
                End If
            End If

            If BegAlloPeriod <> 0 Then
                Dim totalAvaliable As Int64 = 0
                For Each rowA As DataRow In dst.Tables.Item("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE_PLM))
                    For tcnt As Integer = BegAlloPeriod To 4
                        totalAvaliable += Val(rowA.Item(String.Format("QTY_AVA{0}", tcnt.ToString)) & "")
                    Next
                Next
                If totalAvaliable = 0 Then
                    If Not chkShowZero.Checked Then
                        row.Item("SELECTED") = "0"
                    End If
                End If
            End If

        Next


        '  Print_Report_Begin()

        'CR_params.Add("CHKOMITAVAIL", IIf(chkOmitAvail.Checked, "1", "0"))
        'CR_params.Add("CHKOMITPRICE", IIf(chkOmitPrice.Checked, "1", "0"))
        'CR_params.Add("CHKOMITPRICE2", "0")
        'CR_params.Add("CHKSHOWRETAIL", IIf(chkShowRetail.Checked, "1", "0"))
        'CR_params.Add("CHKSHOWSELECTEDONLY", IIf(chkShowSelectedOnly.Checked, "1", "0"))
        'CR_params.Add("IMAGES_FOLDER", FOLDER_NAME)

        Dim RPT As String = "ICRQUOT1"
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            RPT = "ICRQUOT2"
        End If

        Dim ColVisible(4) As Boolean
        ColVisible(0) = True
        ColVisible(1) = (tkb1.Value <= 2)
        ColVisible(2) = (tkb1.Value <= 1)
        ColVisible(3) = (tkb1.Value <= 0)
        ColVisible(4) = chkBeyond.Checked

        'For iCol As Integer = 0 To 4
        '    Dim sValue As String = ""
        '    If iCol = 0 Then sValue = "At Once"
        '    If iCol = 1 Then sValue = Format(dte1.Value, "MM/dd")
        '    If iCol = 2 Then sValue = Format(dte2.Value, "MM/dd")
        '    If iCol = 3 Then sValue = Format(dte3.Value, "MM/dd")
        '    If iCol = 4 Then sValue = "Beyond"
        '    If ColVisible(iCol) Then
        '        CR_params.Add("DTE" & CStr(iCol), sValue)
        '    Else
        '        CR_params.Add("DTE" & CStr(iCol), "")
        '    End If
        'Next

        'setLastShipDate()
        'setLastRcdDate()

        If chk1perPage.Checked Then ' Or STYLE_CODE <> "" Then
            'CR_params.Add("TXTSTYLE_CODE", "") '  STYLE_CODE)

            RPT = "ICRQUOTN"
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                If chkOmitAvail.Checked Then
                    RPT = "ICRQUOTX"
                Else
                    RPT = "ICRQUOTV"
                End If
            End If
        End If

        If eItemKey = "email" Then
            Dim tempFileName As String = rowICTQUOT1.Item("QUOTE_NO")

            Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)
            ' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
            Print_Report_End(, True)
            email_Quote(tempFileName)
        Else

            For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
                row.Item("SELECTED") = "2"

                ASCMAIN1.sql = " SALES_DIVISION_CODE = '" & row.Item("SALES_DIVISION_CODE") & "'"
                Dim rows() As DataRow = dst.Tables("SOTSDIVC").Select(ASCMAIN1.sql)
                If rows.Length <> 0 Then
                    Dim rowSOTSDIVC As DataRow = dst.Tables("SOTSDIVC").Rows(0)
                    row.Item("SALES_DIVISION_CODE_COMB") = rows(0).Item("SALES_DIVISION_CODE_COMB")
                Else
                    row.Item("SALES_DIVISION_CODE_COMB") = row.Item("SALES_DIVISION_CODE")
                End If


            Next

            Dim REPORT_INDEX As Integer = 0
            Dim PDF_FN As String = ""
            Dim PDF_LINKS As String = ""
            Dim SUB_BODY_DESC As String = ""
            Dim SALES_DIVISION_NAME As String = ""
            Dim FABRIC_DESC As String = ""
            Dim DESCHASH As String = ""

            Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/"
            Dim SESSION_NO As String = ASCMAIN1.Next_Control_No("ICTQUOH1.SESSION_NO")
            Dim FILE_NO As Integer = 0


            If chkPublishPDF.Checked Then
                Dim rowICTQUOH1 As DataRow = dst.Tables("ICTQUOH1").NewRow
                rowICTQUOH1.Item("SESSION_NO") = SESSION_NO
                rowICTQUOH1.Item("QUOTE_NO") = QUOTE_NO
                rowICTQUOH1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                rowICTQUOH1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                dst.Tables("ICTQUOH1").Rows.Add(rowICTQUOH1)

            End If

            Do While dst.Tables("ICTQUOT2").Select("SELECTED='2'").Length <> 0

                Print_Report_Begin()

                If chkPublishPDF.Checked Then
                    ' 1 report for every Comb-Division / Sub-Body / Fabric



                    Dim row2() As DataRow = dst.Tables("ICTQUOT2").Select("SELECTED='2'", "SEQ")
                    Dim SALES_DIVISION_CODE_COMB As String = row2(0).Item("SALES_DIVISION_CODE_COMB") & ""
                    Dim SALES_DIVISION_CODE As String = row2(0).Item("SALES_DIVISION_CODE")
                    Dim STYLE_GROUP_CODE As String = row2(0).Item("STYLE_GROUP_CODE")
                    Dim FABRIC_CODE As String = row2(0).Item("FABRIC_CODE")
                    Dim SUB_BODY_CODE As String = row2(0).Item("SUB_BODY_CODE")

                    '   PDF_FN = SALES_DIVISION_CODE & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE
                    PDF_FN = SALES_DIVISION_CODE_COMB & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE

                    ASCMAIN1.sql = "Select SALES_DIVISION_NAME from SOTSDIV1 where SALES_DIVISION_CODE = :PARM1"
                    Dim rowSOTDIV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SALES_DIVISION_CODE_COMB)
                    If rowSOTDIV1 IsNot Nothing Then
                        SALES_DIVISION_NAME = rowSOTDIV1.Item("SALES_DIVISION_NAME")
                    Else
                        SALES_DIVISION_NAME = ""
                    End If

                    '            If chkGROUP.Checked Then
                    '                SALES_DIVISION_NAME = "GROUP" & STYLE_GROUP_CODE
                    '               PDF_FN = SALES_DIVISION_NAME & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE
                    '           End If

                    '     PDF_FN = SALES_DIVISION_CODE_COMB & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE

                    ASCMAIN1.sql = "Select SUB_BODY_DESC from ICTBODY2 where SUB_BODY_CODE = :PARM1"


                    Dim rowICTBODY2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", STYLE_GROUP_CODE)
                    If rowICTBODY2 IsNot Nothing Then
                        SUB_BODY_DESC = rowICTBODY2.Item("SUB_BODY_DESC")
                    Else
                        SUB_BODY_DESC = ""
                    End If

                    If chkGROUP.Checked Then
                        ASCMAIN1.sql = "Select SUB_BODY_DESC from ICTBODY2 where SUB_BODY_CODE = :PARM1"
                        Dim rowICTBODY2A As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SUB_BODY_CODE)
                        If rowICTBODY2A IsNot Nothing Then
                            SUB_BODY_DESC = rowICTBODY2A.Item("SUB_BODY_DESC")
                        Else
                            SUB_BODY_DESC = ""
                        End If
                        ' SUB_BODY_DESC = row2(0).Item("SUB_BODY_CODE")
                        SALES_DIVISION_NAME = "GROUP" & STYLE_GROUP_CODE
                        PDF_FN = SALES_DIVISION_NAME & "-" & SUB_BODY_CODE & "-" & FABRIC_CODE
                    End If


                    ASCMAIN1.sql = "Select FABRIC_DESC from ICTFABR1 where FABRIC_CODE = :PARM1"
                    Dim rowICTFABR1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", FABRIC_CODE)
                    If rowICTFABR1 IsNot Nothing Then
                        FABRIC_DESC = rowICTFABR1.Item("FABRIC_DESC")
                    Else
                        FABRIC_DESC = ""
                    End If


                    DESCHASH = SALES_DIVISION_NAME & SUB_BODY_DESC & FABRIC_DESC
                    DESCHASH = Replace(DESCHASH, " ", "")
                    DESCHASH = Replace(DESCHASH, "/", "")
                    DESCHASH = Replace(DESCHASH, ".", "")
                    DESCHASH = Replace(DESCHASH, ",", "")
                    DESCHASH = Replace(DESCHASH, "&", "")
                    ' DGJ 
                    PDF_FN = DESCHASH & PDF_FN


                    '      ASCMAIN1.sql = "Select * from SOTSDIV1 " _
                    '      & " where SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE_COMB & "'" _
                    '      Dim rowSOTSDIV1 As DataRow = ASCDATA1.GetDataRow

                    '   If rowSOTSDIV1 IsNot Nothing Then
                    'End If
                    Dim sqlw As String

                    If chkGROUP.Checked Then
                        sqlw = "SELECTED='2'" _
                         & " and ISNULL(STYLE_GROUP_CODE,'') = '" & STYLE_GROUP_CODE & "'" _
                         & " and ISNULL(SUB_BODY_CODE,'') = '" & SUB_BODY_CODE & "'" _
                         & " and ISNULL(FABRIC_CODE,'') = '" & FABRIC_CODE & "'"
                    Else
                        sqlw = "SELECTED='2'" _
                        & " and ISNULL(SALES_DIVISION_CODE_COMB,'') = '" & SALES_DIVISION_CODE_COMB & "'" _
                        & " and ISNULL(STYLE_GROUP_CODE,'') = '" & STYLE_GROUP_CODE & "'" _
                        & " and ISNULL(FABRIC_CODE,'') = '" & FABRIC_CODE & "'"

                    End If


                    Dim STYLE_count As Integer = 0
                    For Each row As DataRow In dst.Tables("ICTQUOT2").Select(sqlw, "SEQ")
                        STYLE_count += 1
                        row.Item("SELECTED") = "1"
                        SetRowImage(row)
                    Next
                Else
                    ' 1 report for every 10 Styles
                    'For Each row As DataRow In dst.Tables("ICTQUOT2").Select()
                    '    row.Item("IMAGE") = Null
                    'Next
                    Dim STYLE_count As Integer = 0
                    For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='2'", "SEQ")
                        STYLE_count += 1
                        row.Item("SELECTED") = "1"
                        SetRowImage(row)
                        If STYLE_count >= 10 And Not chkPublishPDF.Checked Then Exit For
                    Next
                    Application.DoEvents()
                End If

                'Runtime.GCSettings.LargeObjectHeapCompactionMode = Runtime.GCLargeObjectHeapCompactionMode.CompactOnce
                'GC.Collect()

                CR_params.Add("CHKOMITAVAIL", IIf(chkOmitAvail.Checked, "1", "0"))
                CR_params.Add("CHKOMITPRICE", IIf(chkOmitPrice.Checked, "1", "0"))
                CR_params.Add("CHKOMITPRICE2", "0")
                'CR_params.Add("CHKSHOWRETAIL", IIf(chkShowRetail.Checked, "1", "0"))
                CR_params.Add("CHKSHOWRETAIL", IIf(chkShowRetailCAD.Checked, "1", "0"))
                CR_params.Add("CHKSHOWSELECTEDONLY", IIf(chkShowSelectedOnly.Checked, "1", "0"))
                CR_params.Add("IMAGES_FOLDER", FOLDER_NAME)

                If RPT = "ICRQUOTV" Then
                    Dim SHOWLASTSHIP = "0"
                    'If chkShowLastShip.Checked Then
                    '    SHOWLASTSHIP = "1"
                    'End If
                    CR_params.Add("SHOWLASTSHIP", SHOWLASTSHIP)
                End If

                For iCol As Integer = BegAlloPeriod To 4
                    Dim sValue As String = ""
                    If iCol = 0 Then sValue = "At Once"
                    If iCol = 1 Then sValue = Format(dte1.Value, "MM/dd")
                    If iCol = 2 Then sValue = Format(dte2.Value, "MM/dd")
                    If iCol = 3 Then sValue = Format(dte3.Value, "MM/dd")
                    If iCol = 4 Then sValue = "Beyond"
                    If ColVisible(iCol) And (iCol >= BegAlloPeriod) Then
                        CR_params.Add("DTE" & CStr(iCol), sValue)
                    Else
                        'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        CR_params.Add("DTE" & CStr(iCol), "")
                        'Else
                        '    CR_params.Add("DTE" & CStr(iCol), "")
                        'End If
                    End If
                Next
                If BegAlloPeriod > 0 Then
                    For bp As Int64 = 0 To BegAlloPeriod - 1
                        CR_params.Add("DTE" & CStr(bp), "")
                    Next
                End If
                For Each row As DataRow In dst.Tables("ICTSTYC1").Select()
                    Dim skipRow As Boolean = True
                    Dim TotalAval As Int64 = 0
                    For iCol As Integer = 0 To 4
                        If ColVisible(iCol) And (iCol >= BegAlloPeriod) Then
                            TotalAval += Val(row.Item("QTY_AVA" & iCol) & "")
                        End If
                    Next
                    Dim minQ As Int64 = 0
                    If IsNumeric(numMinQty.Value) Then
                        minQ = Val(numMinQty.Value)
                    End If
                    Dim capQ As Int64 = 1000000
                    If IsNumeric(numCapQty.Value) Then
                        If Val(numCapQty.Value) > 0 Then
                            capQ = Val(numCapQty.Value)
                        End If
                    End If
                    If (Val(TotalAval) > minQ And Val(TotalAval) < capQ) And skipRow Then
                        skipRow = False
                    End If
                    If skipRow Then
                        If chkShowBlank.Checked Then
                            row.Item("SKIP_COLOR") = "0"
                        Else
                            row.Item("SKIP_COLOR") = "1"
                        End If
                    Else
                        row.Item("SKIP_COLOR") = "0"
                    End If
                Next

                If chk1perPage.Checked Then
                    CR_params.Add("TXTSTYLE_CODE", "")
                End If

                Dim tempFileName As String = ""
                Do
                    REPORT_INDEX += 1
                    tempFileName = rowICTQUOT1.Item("QUOTE_NO") & "-" & Format(REPORT_INDEX, "000")
                Loop While My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")

                Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)

                Dim tempNotMade As Boolean = Not IO.File.Exists(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
                If chkPublishPDF.Checked Then

                    If tempNotMade Then
                        Dim msg As New System.Text.StringBuilder With {.Length = 0}
                        msg.AppendLine(String.Format("Too Many Large Images In File {0}.", PDF_FN))
                        msg.AppendLine("Try Using Low Res Images Or Option")
                        msg.AppendLine("To Scale High Res Images")
                        msg.AppendLine("")
                        msg.AppendLine("This PDF Will Be Skipped But The")
                        msg.AppendLine("Rest Will Be Generated.")
                        MsgBox(msg.ToString, vbOKOnly, "Images Too Large")
                    End If
                    If Not tempNotMade Then
                        Dim PDFD As String = ASCMAIN1.Folders("Archive") & "QuotePDFs\" & SESSION_NO
                        If Not My.Computer.FileSystem.DirectoryExists(PDFD) Then
                            My.Computer.FileSystem.CreateDirectory(PDFD)
                        End If

                        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF", PDFD & "\" & PDF_FN & ".PDF", True)
                        Dim urlpfx As String = "http: //dataservice.absolut1.net/Pictures/StyleCADs/"
                        Dim link As String = "<a href='" & urlpfx & PDF_FN & ".PDF'>Click here for " & PDF_FN & "</a>"
                        PDF_LINKS &= vbCrLf & link

                        FILE_NO += 1

                        Dim strToHash As String = ASCMAIN1.Get_Hash(SESSION_NO & FILE_NO & PDF_FN)

                        Dim ICTQUOH2 As DataRow = dst.Tables("ICTQUOH2").NewRow
                        ICTQUOH2.Item("SESSION_NO") = SESSION_NO
                        ICTQUOH2.Item("FILE_NO") = FILE_NO
                        ICTQUOH2.Item("FILENAME") = PDF_FN
                        'DGJ
                        ICTQUOH2.Item("HASHVALUE") = DESCHASH & strToHash
                        ' ICTQUOH2.Item("HASHVALUE") = strToHash
                        ICTQUOH2.Item("SUB_BODY_DESC") = SUB_BODY_DESC
                        ICTQUOH2.Item("SALES_DIVISION_NAME") = SALES_DIVISION_NAME
                        ICTQUOH2.Item("FABRIC_DESC") = FABRIC_DESC
                        dst.Tables("ICTQUOH2").Rows.Add(ICTQUOH2)

                        '                  Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow
                        '                  If rowARTCUST1 IsNot Nothing Then
                        '                      Dim rowARTCUSTX As DataRow = dst.Tables("ARTCUSTX").NewRow
                        '                      With rowARTCUSTX
                        '.Item("CUST_SHIP_TO_CODE") = rowARTCUST1.Item("CUST_SHIP_TO_CODE") & 
                        '                      End With
                        '                      dst.Tables("ARTCUSTX").Rows.Add(rowARTCUSTX)
                        '                  End If
                        '                  dst.Tables("ARTCUSTX").AcceptChanges()

                    End If
                End If

                If Not tempNotMade Then
                    'Show_Document(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
                    ListPDFSheets.Add(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
                    Print_Report_End(, True)
                End If

                ' Generate_Report(RPT, "Quote Sheet")

                For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
                    row.Item("SELECTED") = "3"
                    row.Item("IMAGE") = DBNull.Value
                Next
            Loop

            'Print_Report_End(, True)

            For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='3'")
                row.Item("SELECTED") = "1"
            Next

            If chkPublishPDF.Checked Then
                Dim ATTACHMENTs As New Dictionary(Of String, String)
                Dim CUST_CODE As String = txtQuoteCUST_CODE.Text
                Dim CUST_NAME As String = txtQuoteCUST_NAME.Text
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                '  ATTACHMENTs.Add(tempFileName & ".pdf", ASCMAIN1.Folders("Temp") & tempFileName & ".pdf")
                ATTACHMENTs.Add("BODY", PDF_LINKS)

                Dim SUBJECT As String = "Quote Sheet"
                Dim PFX As String = ""

                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                If CUST_CODE <> "" Then
                    '   EMAIL_ADDRESSs.Add(rowARTCUST1.Item("CUST_EMAIL") & "", rowARTCUST1.Item("CUST_CONTACT") & "")
                    EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL & "", ASCMAIN1.USER_NAME & "")
                End If

                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                        SUBJECT, "ICFQUOTV", False, True, CUST_CODE, CUST_NAME, "Customer")
                If SEND_NO <> "" Then
                    TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "QUOEML", "Quote Sheet emailed", SEND_NO)
                End If
                Dim sqlDelete As String = "SESSION_NO = '" & SESSION_NO & "'"
                ' & " and SET_ID = '" & SET_ID & "'"
                Update_Record_TDA("ICTQUOH1", sqlDelete)
                Update_Record_TDA("ICTQUOH2", sqlDelete)

                Fill_Records("ICTQUOHF", SESSION_NO)
                addExtraICTQUOHF()

                Using sw As New System.IO.StreamWriter(ASCMAIN1.Folders("Temp") & SESSION_NO & ".TXT")


                    For Each rowICTQUOH2 As DataRow In dst.Tables("ICTQUOH2").Select("")
                        sw.WriteLine(rowICTQUOH2.Item("SALES_DIVISION_NAME") & " - " & rowICTQUOH2.Item("SUB_BODY_DESC") & " - " & rowICTQUOH2.Item("FABRIC_DESC"))
                        sw.WriteLine(LINEPFX & rowICTQUOH2.Item("HASHVALUE"))
                        sw.WriteLine()
                    Next

                End Using
                dst.Tables("ICTQUOH1").Rows.Clear()
                dst.Tables("ICTQUOH2").Rows.Clear()

                Show_Document(ASCMAIN1.Folders("Temp") & SESSION_NO & ".TXT")

            End If

        End If

        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
            row.Item("IMAGE") = Nothing
        Next

        For Each PDF As String In ListPDFSheets
            Show_Document(PDF)
        Next

        If blnShowSelected Then
            chkShowSelectedOnly.Checked = False
        End If

        If MISSING_IMAGES.Count > 0 Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Missing Images"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("The Following Styles Did Not Have")
            iMSG.AppendLine("Set-up In The Style Masterfile:")
            For Each MI As String In MISSING_IMAGES
                iMSG.AppendLine("-> " & MI)
            Next
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        End If
    End Sub

    Private Sub btnFixColors_Click(sender As System.Object, e As System.EventArgs) Handles btnFixColors.Click

        Fill_Records("ICTSTYLX")
        SET_NEW_SIZE_SCALE()
        Fill_Records("ICTSTYCX")
        grdICTSTYCX.Visible = True
        grdICTQUOTX.Visible = False

        For Each row As DataRow In dst.Tables("ICTSTYLX").Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Do While Fix_Colors(STYLE_CODE)

            Loop
            Fix_Size(STYLE_CODE)
        Next
        grdICTSTYCX.Rows.ExpandAll(True)
    End Sub

    Private Sub grdICTSTYCX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTYCX.DoubleClickRow
        If e.Row.Band.Key = "ICTSTYLX" Then
            If grdICTSTYCX IsNot Nothing AndAlso grdICTSTYCX.ActiveCell.Column.Key = "SIZE_SCALE" Then
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
                Do While Fix_Colors(STYLE_CODE)

                Loop
                Fix_Size(STYLE_CODE)
            End If
        End If
    End Sub

    Sub Fix_Size(STYLE_CODE As String)
        Dim rowICTSTYLX As DataRow = dst.Tables("ICTSTYLX").Rows.Find(STYLE_CODE)
        Dim SIZE_SCALE As String = rowICTSTYLX.Item("SIZE_SCALE") & ""
        Do While SIZE_SCALE.EndsWith(vbCrLf)
            SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SIZE_SCALE.Length - 2))
        Loop
        If InStr(SIZE_SCALE, vbCrLf) = 0 Then
            Dim SIZE_CODEs As New List(Of String)
            Dim SIZE_QTYs As New List(Of Integer)
            Try
                If InStr(SIZE_SCALE, "=") <> 0 Then
                    Dim sizes As String = Trim(Split(SIZE_SCALE, "=")(0))
                    Do While InStr(sizes, "//") <> 0
                        sizes = Replace(sizes, "//", "/")
                    Loop
                    Do While InStr(sizes, vbTab & vbTab) <> 0
                        sizes = Replace(sizes, vbTab & vbTab, vbTab)
                    Loop
                    Do While InStr(sizes, "  ") <> 0
                        sizes = Replace(sizes, "  ", " ")
                    Loop
                    Dim qtys As String = Trim(Split(SIZE_SCALE, "=")(1))
                    Dim dl As String = " "
                    If InStr(sizes, vbTab) <> 0 Then dl = vbTab

                    Dim SI As Integer = 0
                    For Each S As String In Split(sizes, dl)
                        SI += 1
                        If SI > 6 Then
                            Debug.Print(STYLE_CODE)
                            Exit For
                        End If
                        rowICTSTYLX.Item("S" & CStr(SI)) = S
                        rowICTSTYLX.Item("Q" & CStr(SI)) = Val(Split(qtys & "//////", "/")(SI - 1) & "")
                    Next
                    rowICTSTYLX.Item("SIZE_SCALE") = ""
                Else
                    Do While InStr(SIZE_SCALE, "  ") <> 0
                        SIZE_SCALE = Replace(SIZE_SCALE, "  ", " ")
                    Loop
                    Dim SS() As String = Split(SIZE_SCALE, " ")
                    Dim SI As Integer = 0
                    For Each S As String In SS
                        If S.Length >= 3 AndAlso Mid(S, S.Length - 1, 1) = "/" AndAlso InStr("123456789", Mid(S, S.Length, 1)) <> 0 Then
                            SIZE_CODEs.Add(Mid(S, 1, S.Length - 2))
                            SIZE_QTYs.Add(Val(Mid(S, S.Length, 1)))
                            SI += 1
                            If SI > 6 Then
                                Debug.Print(STYLE_CODE)
                                Exit For
                            End If
                            rowICTSTYLX.Item("S" & CStr(SI)) = Mid(S, 1, S.Length - 2)
                            rowICTSTYLX.Item("Q" & CStr(SI)) = Val(Mid(S, S.Length, 1))
                            Dim SX As Integer = InStr(SIZE_SCALE, S)
                            SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SX - 1) & Mid(SIZE_SCALE, SX + S.Length))

                        End If
                        SIZE_SCALE = Trim(SIZE_SCALE)
                    Next
                    rowICTSTYLX.Item("SIZE_SCALE") = SIZE_SCALE
                End If
            Catch ex As Exception
                For SI As Integer = 1 To 12

                    rowICTSTYLX.Item("S" & CStr(SI)) = DBNull.Value
                    rowICTSTYLX.Item("Q" & CStr(SI)) = DBNull.Value
                Next
            End Try

        End If
    End Sub

    Function Fix_Colors(STYLE_CODE As String) As Boolean
        Dim fixed As Boolean = False
        Dim rowICTSTYLX As DataRow = dst.Tables("ICTSTYLX").Rows.Find(STYLE_CODE)
        Dim SIZE_SCALE As String = rowICTSTYLX.Item("SIZE_SCALE") & ""
        Dim MAX_LENGTH As Integer = 60
        If SIZE_SCALE <> "" Then
            Dim COLOR_CODEs As New List(Of String)
            For Each row As DataRow In rowICTSTYLX.GetChildRows("ICTSTYLX_ICTSTYCX")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)

                Dim I As Integer = InStr(SIZE_SCALE, COLOR_CODE)
                If I <> 0 Then
                    Dim S As String = Trim(Mid(SIZE_SCALE, I + 3))
                    Dim J As Integer = InStr(Mid(S & "  ", 1, MAX_LENGTH), "  ")
                    Dim K As Integer = InStr(Mid(S & vbCrLf, 1, MAX_LENGTH), vbCrLf)
                    If J = 0 And K = 0 Then
                        J = InStr(Mid(S & " ", 1, MAX_LENGTH), " ")
                    End If
                    If J = 0 Or J > K Then J = K
                    Dim SC As String = ""
                    If J <> 0 Then
                        fixed = True
                        SC = Mid(S, 1, J)
                        SIZE_SCALE = Mid(SIZE_SCALE, 1, I - 1) & Mid(S, J)
                        For C As Integer = 1 To SC.Length - 1
                            If C = 1 Or (C > 1 AndAlso Mid(SC, C + 1, 1) <> " " AndAlso (Mid(SC, C - 1, 1) = " " Or Mid(SC, C - 1, 1) = "/")) Then
                                Mid(SC, C, 1) = Mid(SC, C, 1).ToUpper
                            End If
                        Next
                        If Trim(SC) <> "" Then
                            row.Item("STYLE_COLOR_DESC") = SC
                        End If
                    End If
                End If
            Next
            Dim TF As Boolean = False
            Do
                TF = False
                Do While InStr(SIZE_SCALE, vbCrLf & vbCrLf) <> 0
                    SIZE_SCALE = Replace(SIZE_SCALE, vbCrLf & vbCrLf, vbCrLf)
                    TF = True
                Loop
                Do While SIZE_SCALE.EndsWith(vbCrLf)
                    SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SIZE_SCALE.Length - 2))
                    TF = True
                Loop
                Do While SIZE_SCALE.EndsWith("#")
                    SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SIZE_SCALE.Length - 1))
                    TF = True
                Loop
            Loop While TF

            rowICTSTYLX.Item("SIZE_SCALE") = Trim(SIZE_SCALE)
        End If
    End Function

    Private Sub cmdAllocate_Click(sender As System.Object, e As System.EventArgs) Handles cmdAllocate.Click
        If grdICTQUOT2.ActiveRow Is Nothing OrElse Not grdICTQUOT2.ActiveRow.IsDataRow Or grdICTQUOT2.ActiveRow.IsFilterRow Then
        Else
            STYLE_CODE = grdICTQUOT2.ActiveRow.Cells("STYLE_CODE_PLM").Value
            rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
            Allocate()
        End If
    End Sub

    Private Sub optASL_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optASL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If Not ScreenMode Then Exit Sub
        Allocate()
        Setup_ASL()
    End Sub

    Private Sub cbeWHSE_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cbeWHSE_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_WHSE_CODE()
    End Sub

    Sub Set_WHSE_CODE()
        Dim WHSE_CODE As String = cbeWHSE_CODE.Value & ""
        DirectCast(grdICTSTDQ1.DataSource, DataTable).DefaultView.RowFilter = "WHSE_CODE = '" & WHSE_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
    End Sub

    Sub Setup_ASL()
        For Each row As DataRow In dst.Tables("SOTALLO1").Select("RECORD_TYPE = '1'")
            If optASL.Value = "1" Then
                row.Item("SD_DATE") = row.Item("ORDR_DEMAND_DATE")
            Else
                If row.Item("ORDR_RELEASE_AVAIL") & "" = "" Then
                    row.Item("SD_DATE") = row.Item("ORDR_SHIP_DATE")
                Else
                    If Format(row.Item("ORDR_RELEASE_AVAIL"), "yyyyMMdd") _
                     > Format(row.Item("ORDR_SHIP_DATE"), "yyyyMMdd") Then
                        row.Item("SD_DATE") = row.Item("ORDR_RELEASE_AVAIL")
                    Else
                        row.Item("SD_DATE") = row.Item("ORDR_SHIP_DATE")
                    End If
                End If
            End If
            'row.Item("SD_DATE_X") = Format(row.Item("ORDR_DEMAND_DATE"), "MM/dd/yy")
            row.Item("SD_DATE_X") = Format(row.Item("SD_DATE"), "MM/dd/yy")
        Next
        Set_Table()
    End Sub

    Private Sub tkb1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles tkb1.ValueChanged
        Set_TrackBar()
        refresh_required = True
        cmdGetAvailability.Appearance.ForeColor = Color.Red
    End Sub

    Sub Set_TrackBar()
        lbl1.Visible = (tkb1.Value <= 2)
        dte1.Visible = (tkb1.Value <= 2)
        lbl2.Visible = (tkb1.Value <= 1)
        dte2.Visible = (tkb1.Value <= 1)
        lbl3.Visible = (tkb1.Value <= 0)
        dte3.Visible = (tkb1.Value <= 0)
    End Sub

    Private Sub cmdGetAvailability_Click(sender As System.Object, e As System.EventArgs) Handles cmdGetAvailability.Click
        Get_Availability()
    End Sub

    Sub Get_Availability()

        With grdICTQUOT2.DisplayLayout.Bands(1)
            .Columns("QTY_AVA0").Header.Caption = "At Once" ' Format(dte0.Value, "MM/dd")
            .Columns("QTY_AVA1").Header.Caption = Format(dte1.Value, "MM/dd")
            .Columns("QTY_AVA2").Header.Caption = Format(dte2.Value, "MM/dd")
            .Columns("QTY_AVA3").Header.Caption = Format(dte3.Value, "MM/dd")
            .Columns("QTY_AVA4").Header.Caption = "Beyond"

            .Columns("DTE0").Header.Caption = "Dates"
            .Columns("DTE1").Header.Caption = "Dates"
            .Columns("DTE2").Header.Caption = "Dates"
            .Columns("DTE3").Header.Caption = "Dates"
            .Columns("DTE4").Header.Caption = "Dates"

            ' ENABLING THIS CODE MAKES THE ROWHEIGHT OF BAND1 CRAZY

            'grdICTQUOT2.DisplayLayout.Override.RowSizing = UltraWinGrid.RowSizing.Free
            'grdICTQUOT2.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy

            '.Columns("QTY_AVA0").Hidden = False
            '.Columns("QTY_AVA1").Hidden = Not dte1.Visible
            '.Columns("QTY_AVA2").Hidden = Not dte2.Visible
            '.Columns("QTY_AVA3").Hidden = Not dte3.Visible
            '.Columns("QTY_AVA4").Hidden = Not chkBeyond.Checked

            'If Not dte1.Visible Then
            '    .Columns("QTY_AVA1").Width = 1
            'Else
            '    .Columns("QTY_AVA1").Width = 80
            'End If

            'If Not dte2.Visible Then
            '    .Columns("QTY_AVA2").Width = 1
            'Else
            '    .Columns("QTY_AVA2").Width = 80
            'End If

            'If Not dte3.Visible Then
            '    .Columns("QTY_AVA3").Width = 1
            'Else
            '    .Columns("QTY_AVA3").Width = 80
            'End If

            'If Not chkBeyond.Checked Then
            '    .Columns("QTY_AVA4").Width = 1
            'Else
            '    .Columns("QTY_AVA4").Width = 80
            'End If
            ''grdICTQUOT2.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal
            '.Override.MinRowHeight = 1
            '.Override.ResetMinRowHeight()
            '.Override.DefaultRowHeight = 1
            '.Override.ResetDefaultRowHeight()

            '  .Override.DefaultRowHeight = 4


        End With


        dst.Tables("ICTSTYC1").Columns("QTY_AVA").Expression = "0"
        For Each rowICTQUOT2 As DataRow In dst.Tables("ICTQUOT2").Select("")
            Load_Availability(rowICTQUOT2)
        Next

        Dim MinGrpOpt As Int64 = 0
        If chkALLOSTDT.Checked Then
            MinGrpOpt = cboStartPeriod.SelectedIndex
        End If

        Dim ColVisible(4) As Boolean
        If MinGrpOpt < 1 Then
            ColVisible(0) = True
        End If
        If MinGrpOpt < 2 Then
            ColVisible(1) = (tkb1.Value <= 2)
        End If
        If MinGrpOpt < 3 Then
            ColVisible(2) = (tkb1.Value <= 1)
        End If
        If MinGrpOpt < 4 Then
            ColVisible(3) = (tkb1.Value <= 0)
        End If
        If MinGrpOpt < 5 Then
            ColVisible(4) = chkBeyond.Checked
        End If

        Dim EX As String = ""
        For I As Integer = 0 To 4
            If ColVisible(I) Then
                EX &= "+ISNULL(QTY_AVA" & CStr(I) & ",0)"
            End If
        Next
        dst.Tables("ICTSTYC1").Columns("QTY_AVA").Expression = Mid(EX, 2)

        refresh_required = False
        cmdGetAvailability.Appearance.ForeColor = Color.Empty

    End Sub

    Sub Sort_by_Style()
        Dim SEQ As Integer = 0
        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("", "STYLE_CODE_PLM")
            SEQ += 10
            row.Item("SEQ") = SEQ
        Next
        Sort_grdColumns(grdICTQUOT2, "SEQ")
        Sort_grdColumns(grdICTQUOT2B, "SEQ")
    End Sub

    'Private Sub setLastShipDate()
    '    For Each row As DataRow In dst.Tables("ICTSTYC1").Select()
    '        Dim S As New System.Text.StringBuilder With {.Length = 0}
    '        S.AppendLine("SELECT MAX(I1.INV_DATE) LAST_SHIP_DATE")
    '        S.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
    '        S.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
    '        S.AppendLine(String.Format("AND I2.STYLE_CODE = '{0}'", row.Item("STYLE_CODE").ToString & String.Empty))
    '        S.AppendLine(String.Format("AND I2.COLOR_CODE = '{0}'", row.Item("COLOR_CODE").ToString & String.Empty))
    '        S.AppendLine(String.Format("AND I1.CUST_CODE = '{0}'", Absx1.txtFor("CUST_CODE").Text.ToString & String.Empty))
    '        ASCMAIN1.sql = S.ToString()
    '        Dim LAST_SHIP_DATE As String = ASCDATA1.GetDataValue
    '        If IsDate(LAST_SHIP_DATE) Then
    '            LAST_SHIP_DATE = Format(CDate(LAST_SHIP_DATE), "MM/dd/yy")
    '        Else
    '            LAST_SHIP_DATE = ""
    '        End If
    '        row.Item("LAST_SHIP_DATE") = LAST_SHIP_DATE
    '    Next
    'End Sub

    Private Sub setLastRcdDate()
        For Each row As DataRow In dst.Tables("ICTSTYC1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim S As New System.Text.StringBuilder With {.Length = 0}
            Dim STYLE_CODE As String = row.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = row.Item("COLOR_CODE").ToString & String.Empty
            ASCMAIN1.Progress("Calculating Last Rcd Date", String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE))
            S.AppendLine("SELECT NVL(TO_CHAR(MAX(POTSHIP2.PO_DATE_RECEIVED),'MM/DD/YY'),'') PO_DATE_RECEIVED")
            S.AppendLine("FROM POTORDR2, POTSHIP3, POTSHIP2")
            S.AppendLine("WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
            S.AppendLine("AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO")
            S.AppendLine("AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO")
            S.AppendLine("AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO")
            S.AppendLine(String.Format("AND POTORDR2.STYLE_CODE = '{0}'", STYLE_CODE))
            S.AppendLine(String.Format("AND POTORDR2.COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = S.ToString()
            Dim LAST_RCD_DATE As String = ASCDATA1.GetDataValue
            'If STYLE_CODE = "VCO72239A" Then Stop
            If IsDate(LAST_RCD_DATE) Then
                If chkRECDATES.Checked Then
                    If CDate(LAST_RCD_DATE) < dteRECDATEFR.DateTime Or CDate(LAST_RCD_DATE) > dteRECDATETO.DateTime Then
                        LAST_RCD_DATE = ""
                    End If
                End If
            End If
            If IsDate(LAST_RCD_DATE) Then
                LAST_RCD_DATE = Format(CDate(LAST_RCD_DATE), "MM/dd/yy")
            Else
                S.Length = 0
                S.AppendLine("SELECT SUM(NVL(WHSE_QTY_TRAN,0)) AS IN_TRAN")
                S.AppendLine("FROM ICTSTAT2")
                S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                If cboIncludeWhse.Text <> "All Whse" Then
                    If chkIncludeWhse.Checked Then
                        S.AppendLine(String.Format("AND WHSE_CODE = '{0}'", cboIncludeWhse.Text))
                    Else
                        S.AppendLine(String.Format("AND WHSE_CODE <> '{0}'", cboIncludeWhse.Text))
                    End If
                End If
                ASCMAIN1.sql = S.ToString()
                Dim IN_TRAN As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                If IN_TRAN > 0 Then
                    LAST_RCD_DATE = "In-Tran"
                Else
                    S.Length = 0
                    S.AppendLine("SELECT SUM(NVL(WHSE_QTY_ON_ORDER,0)) AS IN_WIP")
                    S.AppendLine("FROM ICTSTAT2")
                    S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                    If cboIncludeWhse.Text <> "All Whse" Then
                        If chkIncludeWhse.Checked Then
                            S.AppendLine(String.Format("AND WHSE_CODE = '{0}'", cboIncludeWhse.Text))
                        Else
                            S.AppendLine(String.Format("AND WHSE_CODE <> '{0}'", cboIncludeWhse.Text))
                        End If
                    End If
                    ASCMAIN1.sql = S.ToString()
                    Dim IN_WIP As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                    If IN_WIP > 0 Then
                        LAST_RCD_DATE = "In-WIP"
                    Else
                        LAST_RCD_DATE = ""
                    End If
                End If
            End If
            row.Item("LAST_RCD_DATE") = LAST_RCD_DATE
        Next
        If chkShowLastRcd.Checked Then
            grdICTQUOT2.DisplayLayout.Bands(1).Columns("LAST_RCD_DATE").Hidden = False
            grdICTQUOT2.DisplayLayout.Bands(1).Columns("LAST_RCD_DATE").Header.Caption = "Last Rcd Date"
        Else
            grdICTQUOT2.DisplayLayout.Bands(1).Columns("LAST_RCD_DATE").Hidden = True
        End If
        ASCMAIN1.Progress("", "")
    End Sub

#Region "Save and Load Settings - these should be in ASFBASE0"

    Sub Retrieve_Settings(ByRef SET_ID As String)

        Dim grdSetup As UltraWinGrid.UltraGrid = Absc1.grdSetup
        Dim FORM_NAME As String = Me.Name
        Dim tblASTROPT1 As DataTable = dst.Tables("ASTROPT1")

        Dim SET_CTL_NAME As String
        Dim SET_CTL_TYPE As String
        Dim SET_CTL_TAG As String
        Dim SET_CTL_DATA As String

        Dim rowASTROPT1 As DataRow = Fill_Record("ASTROPT1", New String() {FORM_NAME, SET_ID})

        Clear_grdSetup(grdSetup, True)

        Fill_Records("ASTROPT2", New String() {FORM_NAME, SET_ID})

        For Each rowASTROPT2 As DataRow In dst.Tables("ASTROPT2").Select("") ' ASCDATA1.GetDataTable(sql).Select("", "SET_CTL_TAG")
            SET_CTL_NAME = rowASTROPT2.Item("SET_CTL_NAME") & ""
            SET_CTL_TYPE = rowASTROPT2.Item("SET_CTL_TYPE") & ""
            SET_CTL_TAG = rowASTROPT2.Item("SET_CTL_TAG") & ""
            SET_CTL_DATA = rowASTROPT2.Item("SET_CTL_DATA") & ""

            Dim gDR As DataRow

            If SET_CTL_NAME = "grdSetup" Then
                If SET_CTL_TAG = "" Then
                    Dim GRDCOLS() As String = Split(SET_CTL_DATA, vbTab)
                    gDR = DirectCast(grdSetup.DataSource, DataTable).Rows.Find(GRDCOLS(0))
                    If gDR IsNot Nothing Then
                        If Val(GRDCOLS(1) & "") <> 0 Then
                            gDR.Item("SEQUENCE") = Val(GRDCOLS(1) & "")
                        End If
                        gDR.Item("PAGE_BREAK") = GRDCOLS(2)
                        gDR.Item("EXCLUDE") = GRDCOLS(3)
                        gDR.Item("GROUP_ALL_OTHERS") = GRDCOLS(4)
                    End If
                Else
                    Dim COLUMN_NAME As String = SET_CTL_TAG
                    gDR = DirectCast(grdSetup.DataSource, DataTable).Rows.Find(COLUMN_NAME)
                    If gDR.Item("CODE_VALUES") & "" = "" Then
                        gDR.Item("CODE_VALUES") = SET_CTL_DATA
                    Else
                        gDR.Item("CODE_VALUES") &= "," & SET_CTL_DATA
                    End If
                End If
            Else
                Dim C As Control = Absx1.CtlFor(SET_CTL_TAG, True)
                If C IsNot Nothing Then
                    Select Case SET_CTL_TYPE
                        Case "UltraCheckEditor"
                            Absx1.chkFor(SET_CTL_TAG).Checked = (SET_CTL_DATA = "True")
                        Case "UltraOptionSet"
                            Absx1.optFor(SET_CTL_TAG).Value = SET_CTL_DATA
                        Case "UltraTrackBar"
                            If SET_CTL_DATA <> "" Then
                                DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinEditors.UltraTrackBar).Value = SET_CTL_DATA
                            End If
                        Case "ABSCheckBox"
                            DirectCast(Absx1.CtlFor(SET_CTL_TAG), ABSCS.ABSCheckBox).ABSChecked = SET_CTL_DATA
                        Case "UltraCombo"
                            Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinGrid.UltraCombo)
                            'cmbctl.Text = SET_CTL_DATA
                            cmbctl.Value = SET_CTL_DATA
                            If SET_CTL_TAG = "RYP" Or SET_CTL_TAG = "RYP0" Or SET_CTL_TAG = "RYP1" Then
                                If SET_ID <> "0000000000" AndAlso rowASTROPT1.Item("SET_YP_REL") & "" = "1" AndAlso rowASTROPT1.Item("SET_YP_BASE") & "" <> "" Then
                                    Dim RYP As String = Mid(SET_CTL_DATA, 1, 4) & Mid(SET_CTL_DATA, 6, 2)
                                    Dim NP As Integer = ASCMAIN1.Period_Diff(rowASTROPT1.Item("SET_YP_BASE") & "", RYP)
                                    cmbctl.Text = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, NP)), 1, 16)
                                End If
                            End If
                            If SET_CTL_TAG = "RYW" Or SET_CTL_TAG = "RYW0" Or SET_CTL_TAG = "RYW1" Then
                                If SET_ID <> "0000000000" AndAlso rowASTROPT1.Item("SET_YP_REL") & "" = "1" AndAlso rowASTROPT1.Item("SET_YW_BASE") & "" <> "" Then
                                    Dim RYW As String = Mid(SET_CTL_DATA, 1, 4) & Mid(SET_CTL_DATA, 6, 2)
                                    Dim NW As Integer = ASCMAIN1.Week_Diff(rowASTROPT1.Item("SET_YW_BASE") & "", RYW)
                                    cmbctl.Text = Mid(ASCMAIN1.Get_Legend_Wk(ASCMAIN1.Week_Calc(ASCMAIN1.CYW, NW)), 1, 17)
                                End If
                            End If
                        Case "UltraComboEditor"
                            Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinEditors.UltraComboEditor)
                            cbectl.Value = SET_CTL_DATA
                        Case "UltraNumericEditor"
                            Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinEditors.UltraNumericEditor)
                            numctl.Value = Val(SET_CTL_DATA)
                        Case "UltraDateTimeEditor"
                            Dim dtectl As UltraWinEditors.UltraDateTimeEditor = DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinEditors.UltraDateTimeEditor)
                            dtectl.Value = SET_CTL_DATA
                        Case Else
                            Absx1.CtlFor(SET_CTL_TAG).Text = SET_CTL_DATA
                    End Select

                End If
            End If
        Next

        If grdSetup.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If

    End Sub

    Sub Save_Settings(ByRef SET_ID As String, SET_DESC As String)

        Dim grdSetup As UltraWinGrid.UltraGrid = Absc1.grdSetup
        Dim tblASTROPT1 As DataTable = dst.Tables("ASTROPT1")
        Dim FORM_NAME As String = Me.Name

        ' BeginTrans()

        Dim rowASTROPT1 As DataRow
        Dim LAST_DATE As Date = DATETIME_STAMP
        rowASTROPT1 = tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})

        If rowASTROPT1 Is Nothing Then
            rowASTROPT1 = tblASTROPT1.NewRow()
            ' SET_ID = ASCMAIN1.Next_Control_No("ASTROPT1.SET_ID")
            rowASTROPT1.Item("FORM_NAME") = FORM_NAME
            rowASTROPT1.Item("SET_ID") = SET_ID
            rowASTROPT1.Item("SET_YP_BASE") = ASCMAIN1.CYP
            rowASTROPT1.Item("SET_YP_REL") = "1"
            rowASTROPT1.Item("SET_ALLOW_OTHERS") = "0"
            rowASTROPT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowASTROPT1.Item("INIT_DATE") = LAST_DATE
            tblASTROPT1.Rows.Add(rowASTROPT1)
        Else
            rowASTROPT1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowASTROPT1.Item("LAST_DATE") = LAST_DATE
        End If

        Dim sqlDelete As String = "FORM_NAME = '" & FORM_NAME & "'" _
            & " and SET_ID = '" & SET_ID & "'"

        rowASTROPT1.Item("SET_DESC") = SET_DESC
        rowASTROPT1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowASTROPT1.Item("LAST_DATE") = LAST_DATE

        Update_Record_TDA("ASTROPT1", sqlDelete)

        Dim rowASTROPT2 As DataRow
        Dim tblASTROPT2 As DataTable = dst.Tables("ASTROPT2")
        dst.Tables("ASTROPT2").Rows.Clear()

        Save_Settings_ctls(grpICTQUOT1_Options, FORM_NAME, SET_ID, XNO, tblASTROPT2)
        Save_Settings_ctls(grpExcelOptions, FORM_NAME, SET_ID, XNO, tblASTROPT2)
        Save_Settings_ctls(grpAvailabilityDates, FORM_NAME, SET_ID, XNO, tblASTROPT2)

        For Each gr As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSetup.Rows
            rowASTROPT2 = tblASTROPT2.NewRow()
            With rowASTROPT2
                .Item("FORM_NAME") = FORM_NAME
                .Item("SET_ID") = SET_ID
                .Item("SET_CTL_NAME") = grdSetup.Name
                .Item("SET_CTL_TYPE") = grdSetup.GetType.Name
                .Item("SET_CTL_DATA") = gr.Cells("COLUMN_NAME").Text _
                            & vbTab & gr.Cells("SEQUENCE").Text _
                            & vbTab & gr.Cells("PAGE_BREAK").Value _
                            & vbTab & gr.Cells("EXCLUDE").Value _
                            & vbTab & gr.Cells("GROUP_ALL_OTHERS").Value
                .Item("SET_CTL_TAG") = ""
                '.Item("XNO") = XNO
            End With
            tblASTROPT2.Rows.Add(rowASTROPT2)

            If gr.Cells("CODE_VALUES").Text <> "" Then
                Dim CODE_VALUES() As String = Split(gr.Cells("CODE_VALUES").Text, ",")
                For Each CODE_VALUE As String In CODE_VALUES
                    rowASTROPT2 = tblASTROPT2.NewRow()
                    With rowASTROPT2
                        .Item("FORM_NAME") = FORM_NAME
                        .Item("SET_ID") = SET_ID
                        .Item("SET_CTL_NAME") = grdSetup.Name
                        .Item("SET_CTL_TYPE") = grdSetup.GetType.Name
                        .Item("SET_CTL_DATA") = CODE_VALUE
                        .Item("SET_CTL_TAG") = gr.Cells("COLUMN_NAME").Text
                        ' .Item("XNO") = XNO
                    End With
                    tblASTROPT2.Rows.Add(rowASTROPT2)
                Next
            End If
        Next

        Update_Record_TDA("ASTROPT2", sqlDelete & " and XNO is Null")

        ' CommitTrans()

        ' Sort_grdColumns(grdASTROPT1, "LAST_DATE".ToLower)

        'If XNO = "" And SET_ID <> "0000000000" Then
        '    MsgBox("Settings have been Saved", MsgBoxStyle.OkOnly, "Verification")
        'End If

    End Sub

    Sub Save_Settings_ctls(
    ByRef cc As Control,
    ByRef FORM_NAME As String,
    ByRef SET_ID As String,
    ByRef XNO As String,
    ByRef tblASTROPT2 As DataTable)

        Dim rowASTROPT2 As DataRow
        For Each ctl As Control In cc.Controls
            If ctl.Controls.Count > 0 Then
                Call Save_Settings_ctls(ctl, FORM_NAME, SET_ID, XNO, tblASTROPT2)
            End If
            Dim ABSCOLUMN_NAME As String = Absx1.GetABSColumnName(ctl)
            If ABSCOLUMN_NAME <> "" Then
                rowASTROPT2 = tblASTROPT2.NewRow()
                With rowASTROPT2
                    .Item("FORM_NAME") = FORM_NAME
                    .Item("SET_ID") = SET_ID
                    .Item("SET_CTL_NAME") = ctl.Name
                    .Item("SET_CTL_TYPE") = ctl.GetType.Name
                    Select Case ctl.GetType.Name
                        Case "UltraCheckEditor"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraCheckEditor).Checked
                        Case "UltraOptionSet"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraOptionSet).Value
                        Case "ABSCheckBox"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, ABSCS.ABSCheckBox).ABSChecked
                        Case "UltraTrackBar"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraTrackBar).Value
                        Case "UltraCombo"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinGrid.UltraCombo).Value
                        Case "UltraComboEditor"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraComboEditor).Value
                        Case "UltraNumericEditor"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value
                        Case "UltraDateTimeEditor"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value
                        Case Else
                            .Item("SET_CTL_DATA") = ctl.Text
                    End Select
                    .Item("SET_CTL_TAG") = ABSCOLUMN_NAME
                    '.Item("XNO") = XNO
                End With

                tblASTROPT2.Rows.Add(rowASTROPT2)
            End If
        Next
    End Sub

#End Region

    Private Sub btnRefreshStyles_Click(sender As System.Object, e As System.EventArgs) Handles btnRefreshStyles.Click
        If chkAgedPast.Checked Then
            If chkAgedLess.Checked = False And chkAgedGreater.Checked = False Then
                MsgBox("You Must Pick An Aging Range", vbOKOnly, "Aging Option")
                Exit Sub
            End If
        End If

        Absc1.Get_SQL("*") ' ,"ICTQUOTV")
        Dim SQL As String = ""
        If chkRECDATES.Checked Then
            SQL = "(Select Distinct STYLE_CODE from ICTSTAT2) ICTQUOTV"
        Else
            SQL = "(Select Distinct STYLE_CODE from ICTSTAT2 where NVL(WHSE_QTY_ON_HAND,0) <> 0 or NVL(WHSE_QTY_ON_ORDER,0) <> 0 or NVL(WHSE_QTY_TRAN,0) <> 0) ICTQUOTV"
        End If

        Dim sqlwhere As String = ""
        If optASN.Value = "S" Then
            sqlwhere = " and ICTSTYL1.CUST_CODE is Null"
        ElseIf optASN.Value = "N" Then
            sqlwhere = " and ICTSTYL1.CUST_CODE is NOT Null"
        End If
        Dim SQLUnion As String = ""
        If chkInTranAsNow.Checked Then
            If IsDate(dteInTranAsNow.Value) Then
                SQLUnion = SQLUnion & "    UNION"
                SQLUnion = SQLUnion & "    SELECT"
                SQLUnion = SQLUnion & "    P2.STYLE_CODE"
                SQLUnion = SQLUnion & "    FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2, POTSHIP1 SP1"
                SQLUnion = SQLUnion & "    WHERE P2.STYLE_CODE = S1.STYLE_CODE"
                SQLUnion = SQLUnion & "    AND P2.PO_ORDER_NO = S3.PO_ORDER_NO"
                SQLUnion = SQLUnion & "    AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO"
                SQLUnion = SQLUnion & "    AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO"
                SQLUnion = SQLUnion & "    AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO"
                SQLUnion = SQLUnion & "    AND SP1.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO"
                SQLUnion = SQLUnion & "    AND PO_SHIP_STATUS = 'O'"
                If CDate(dte0.DateTime.ToShortDateString) > CDate(dteInTranAsNow.DateTime.ToShortDateString) Then
                    SQLUnion = SQLUnion & "    AND (NVL(SP1.PO_SHIP_ETA,'01-JAN-2100') + PO_SHIP_LANDING_LEAD_DAYS) <= '" & Format(dte0.DateTime, "dd-MMM-yyyy") & "'"
                Else
                    SQLUnion = SQLUnion & "    AND (NVL(SP1.PO_SHIP_ETA,'01-JAN-2100') + PO_SHIP_LANDING_LEAD_DAYS) <= '" & Format(dteInTranAsNow.DateTime, "dd-MMM-yyyy") & "'"
                End If
                SQLUnion = SQLUnion & "    GROUP BY"
                SQLUnion = SQLUnion & "    P2.STYLE_CODE"

                If chkLOADWIP.Checked Then
                    SQLUnion = SQLUnion & "    UNION"
                    SQLUnion = SQLUnion & "    SELECT P2.STYLE_CODE"
                    SQLUnion = SQLUnion & "    FROM ICTSTYL1 S1, POTORDR2 P2"
                    SQLUnion = SQLUnion & "    WHERE P2.STYLE_CODE = S1.STYLE_CODE"
                    SQLUnion = SQLUnion & "    AND P2.PO_STATUS = 'O'"
                    SQLUnion = SQLUnion & String.Format("    AND (NVL(P2.PO_DATE_ETA,'01-JAN-2100') + 10) <= '{0}'", Format(dte0.DateTime, "dd-MMM-yyyy"))
                    SQLUnion = SQLUnion & "    GROUP BY P2.STYLE_CODE"
                End If

            End If
        End If
        If chkRECDATES.Checked Then
            If IsDate(dteRECDATEFR.DateTime) And IsDate(dteRECDATETO.DateTime) Then
                sqlwhere = sqlwhere & " and ICTSTYL1.STYLE_CODE IN ("
                sqlwhere = sqlwhere & "    SELECT"
                sqlwhere = sqlwhere & "    P2.STYLE_CODE"
                sqlwhere = sqlwhere & "    FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2"
                sqlwhere = sqlwhere & "    WHERE P2.STYLE_CODE = S1.STYLE_CODE"
                sqlwhere = sqlwhere & "    AND P2.PO_ORDER_NO = S3.PO_ORDER_NO"
                sqlwhere = sqlwhere & "    AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO"
                sqlwhere = sqlwhere & "    AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO"
                sqlwhere = sqlwhere & "    AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO"
                sqlwhere = sqlwhere & "    AND PO_SHIP_STATUS = 'C'"
                sqlwhere = sqlwhere & "    AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-1900') >= '" & Format(dteRECDATEFR.DateTime, "dd-MMM-yyyy") & "'"
                sqlwhere = sqlwhere & "    AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2099') <= '" & Format(dteRECDATETO.DateTime, "dd-MMM-yyyy") & "'"
                sqlwhere = sqlwhere & "    GROUP BY"
                sqlwhere = sqlwhere & "    P2.STYLE_CODE"
                sqlwhere = sqlwhere & SQLUnion
                sqlwhere = sqlwhere & " )"
            End If
        End If
        If chkBABefore.Checked Then
            If IsDate(dteBABefore.Value) Then
                sqlwhere = sqlwhere & " and ICTSTYL1.STYLE_CODE IN ("
                sqlwhere = sqlwhere & "    SELECT"
                sqlwhere = sqlwhere & "    P2.STYLE_CODE"
                sqlwhere = sqlwhere & "    FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2"
                sqlwhere = sqlwhere & "    WHERE P2.STYLE_CODE = S1.STYLE_CODE"
                sqlwhere = sqlwhere & "    AND P2.PO_ORDER_NO = S3.PO_ORDER_NO"
                sqlwhere = sqlwhere & "    AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO"
                sqlwhere = sqlwhere & "    AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO"
                sqlwhere = sqlwhere & "    AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO"
                sqlwhere = sqlwhere & "    AND PO_SHIP_STATUS = 'C'"
                sqlwhere = sqlwhere & "    AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2100') < '" & Format(dteBABefore.DateTime, "dd-MMM-yyyy") & "'"
                sqlwhere = sqlwhere & "    GROUP BY"
                sqlwhere = sqlwhere & "    P2.STYLE_CODE"
                sqlwhere = sqlwhere & SQLUnion
                sqlwhere = sqlwhere & " )"
            End If
        End If

        If chkBAAfter.Checked Then
            If IsDate(dteBAAfter.Value) Then
                sqlwhere = sqlwhere & " and ICTSTYL1.STYLE_CODE IN ("
                sqlwhere = sqlwhere & "    SELECT"
                sqlwhere = sqlwhere & "    P2.STYLE_CODE"
                sqlwhere = sqlwhere & "    FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2"
                sqlwhere = sqlwhere & "    WHERE P2.STYLE_CODE = S1.STYLE_CODE"
                sqlwhere = sqlwhere & "    AND P2.PO_ORDER_NO = S3.PO_ORDER_NO"
                sqlwhere = sqlwhere & "    AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO"
                sqlwhere = sqlwhere & "    AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO"
                sqlwhere = sqlwhere & "    AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO"
                sqlwhere = sqlwhere & "    AND PO_SHIP_STATUS = 'C'"
                sqlwhere = sqlwhere & "    AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-1900') > '" & Format(dteBAAfter.DateTime, "dd-MMM-yyyy") & "'"
                sqlwhere = sqlwhere & "    GROUP BY"
                sqlwhere = sqlwhere & "    P2.STYLE_CODE"
                sqlwhere = sqlwhere & SQLUnion
                sqlwhere = sqlwhere & " )"
            End If
        End If

        If chkAgedPast.Checked Then
            Dim GreaterDate As Date = Now
            If chkAgedGreater.Checked Then
                GreaterDate = Now.AddDays(numAgedGreater.Value * -1)
            Else

            End If
            Dim GreaterDateF As String = Format(GreaterDate, "dd-MMM-yyyy")

            Dim LessDate As Date = Now
            If chkAgedLess.Checked Then
                LessDate = Now.AddDays(numAgedLess.Value * -1)
            End If
            Dim LessDateF As String = Format(LessDate, "dd-MMM-yyyy")

            sqlwhere = sqlwhere & " and ICTSTYL1.STYLE_CODE IN ("
            sqlwhere = sqlwhere & "     SELECT POTORDR2.STYLE_CODE"
            sqlwhere = sqlwhere & "     FROM POTORDR2, POTSHIP3, POTSHIP2, ICTSTAT2"
            sqlwhere = sqlwhere & "     WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO"
            sqlwhere = sqlwhere & "     AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO"
            sqlwhere = sqlwhere & "     AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO"
            sqlwhere = sqlwhere & "     AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO"
            sqlwhere = sqlwhere & "     AND POTORDR2.STYLE_CODE = ICTSTAT2.STYLE_CODE"
            sqlwhere = sqlwhere & "     AND POTORDR2.COLOR_CODE = ICTSTAT2.COLOR_CODE"
            sqlwhere = sqlwhere & "     AND NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) > 0"
            sqlwhere = sqlwhere & "     GROUP BY POTORDR2.STYLE_CODE"
            If chkInclWIPTRAN.Checked Then
                sqlwhere = sqlwhere & "    UNION"
                sqlwhere = sqlwhere & "    SELECT STYLE_CODE"
                sqlwhere = sqlwhere & "    FROM ICTSTAT2"
                sqlwhere = sqlwhere & "    WHERE (NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0)) <> 0"
            End If
            If chkAgedLess.Checked And chkAgedGreater.Checked Then
                sqlwhere = sqlwhere & String.Format("     HAVING (MAX(POTSHIP2.PO_DATE_RECEIVED) >= '{0}' AND MAX(POTSHIP2.PO_DATE_RECEIVED) <= '{1}')", LessDateF, GreaterDateF)
            Else
                If chkAgedGreater.Checked Then
                    sqlwhere = sqlwhere & String.Format("     HAVING MAX(POTSHIP2.PO_DATE_RECEIVED) <= '{0}'", GreaterDateF)
                Else
                    sqlwhere = sqlwhere & String.Format("     HAVING MAX(POTSHIP2.PO_DATE_RECEIVED) >= '{0}'", LessDateF)
                End If
            End If
            sqlwhere = sqlwhere & " )"
        End If

        If chkBAAfter.Checked Then
            Absc1.sql_JOIN = Absc1.sql_JOIN.Replace("AND ICTSTYL1.STYLE_CODE = ICTQUOTV.STYLE_CODE", "AND ICTSTYL1.STYLE_CODE = ICTQUOTV.STYLE_CODE (+)")
            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE from  " & SQL & Absc1.sql_TABLE_NAMEs & ASCMAIN1.SQL_Add_WHERE(Replace(Absc1.sql_JOIN, "ICTQUOTV.SUB_BODY_CODE", "ICTSTYL1.SUB_BODY_CODE") & Absc1.sql_WHERE & sqlwhere)
        Else
            ASCMAIN1.sql = "Select ICTQUOTV.STYLE_CODE from  " & SQL & Absc1.sql_TABLE_NAMEs & ASCMAIN1.SQL_Add_WHERE(Replace(Absc1.sql_JOIN, "ICTQUOTV.SUB_BODY_CODE", "ICTSTYL1.SUB_BODY_CODE") & Absc1.sql_WHERE & sqlwhere)
        End If

        Dim NEW_STYLES As Integer = 0

        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE})
            If rowICTQUOT2 Is Nothing Then
                NEW_STYLES += 1
                Add_to_Quote(STYLE_CODE)
            End If
        Next

        If chkBABefore.Checked Then
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select
                Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE").ToString & String.Empty
                Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT SUM(P2.PO_QTY_REC) PO_QTY_REC_SUM")
                SQLS.AppendLine("FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2")
                SQLS.AppendLine("WHERE P2.STYLE_CODE = S1.STYLE_CODE")
                SQLS.AppendLine("AND P2.PO_ORDER_NO = S3.PO_ORDER_NO")
                SQLS.AppendLine("AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO")
                SQLS.AppendLine("AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
                SQLS.AppendLine("AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO")
                SQLS.AppendLine("AND PO_SHIP_STATUS = 'C'")
                SQLS.AppendLine("AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2100') < '" & Format(dteBABefore.DateTime, "dd-MMM-yyyy") & "'")
                SQLS.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(String.Format("AND P2.COLOR_CODE = '{0}'", COLOR_CODE))
                ASCMAIN1.sql = SQLS.ToString()
                Dim PO_QTY_REC_SUM As Int64 = Val(ASCDATA1.GetDataValue)
                If PO_QTY_REC_SUM = 0 Then
                    'If STYLE_CODE = "VCO51222" Then
                    '    Stop
                    'End If
                    Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                    Dim rowICTSTYCX As DataRow = dst.Tables("ICTSTYCX").Select(filter).FirstOrDefault
                    If Not IsNothing(rowICTSTYCX) Then
                        rowICTSTYCX.Delete()
                    End If
                    rowICTSTYC1.Delete()
                End If
            Next
        End If

        If chkBAAfter.Checked Then
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select
                Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE").ToString & String.Empty
                Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT SUM(PO_QTY_REC_SUM) PO_QTY_REC_SUM")
                SQLS.AppendLine("FROM (")
                SQLS.AppendLine("SELECT SUM(P2.PO_QTY_REC) PO_QTY_REC_SUM")
                SQLS.AppendLine("FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2")
                SQLS.AppendLine("WHERE P2.STYLE_CODE = S1.STYLE_CODE")
                SQLS.AppendLine("AND P2.PO_ORDER_NO = S3.PO_ORDER_NO")
                SQLS.AppendLine("AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO")
                SQLS.AppendLine("AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
                SQLS.AppendLine("AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO")
                SQLS.AppendLine("AND PO_SHIP_STATUS = 'C'")
                SQLS.AppendLine("AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2100') > '" & Format(dteBAAfter.DateTime, "dd-MMM-yyyy") & "'")
                SQLS.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(String.Format("AND P2.COLOR_CODE = '{0}'", COLOR_CODE))
                If chkInTranAsNow.Checked Then
                    If IsDate(dteInTranAsNow.Value) Then
                        SQLS.AppendLine("UNION")
                        SQLS.AppendLine("SELECT SUM(S3.PO_QTY_SHP) PO_QTY_REC_SUM")
                        SQLS.AppendLine("FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2, POTSHIP1 SP1")
                        SQLS.AppendLine("WHERE P2.STYLE_CODE = S1.STYLE_CODE")
                        SQLS.AppendLine("AND P2.PO_ORDER_NO = S3.PO_ORDER_NO")
                        SQLS.AppendLine("AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO")
                        SQLS.AppendLine("AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
                        SQLS.AppendLine("AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO")
                        SQLS.AppendLine("AND SP1.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
                        SQLS.AppendLine("AND PO_SHIP_STATUS = 'O'")
                        If CDate(dte0.DateTime.ToShortDateString) > CDate(dteInTranAsNow.DateTime.ToShortDateString) Then
                            SQLS.AppendLine(String.Format("AND (NVL(SP1.PO_SHIP_ETA,'01-JAN-2100') + PO_SHIP_LANDING_LEAD_DAYS) <= '{0}'", Format(dte0.DateTime, "dd-MMM-yyyy")))
                        Else
                            SQLS.AppendLine(String.Format("AND (NVL(SP1.PO_SHIP_ETA,'01-JAN-2100') + PO_SHIP_LANDING_LEAD_DAYS) <= '{0}'", Format(dteInTranAsNow.DateTime, "dd-MMM-yyyy")))
                        End If

                        SQLS.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE))
                        SQLS.AppendLine(String.Format("AND P2.COLOR_CODE = '{0}'", COLOR_CODE))
                    End If
                End If
                SQLS.AppendLine(")")
                If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    If STYLE_CODE = "VCO51665" And COLOR_CODE = "260" Then
                        Stop
                    End If
                End If
                ASCMAIN1.sql = SQLS.ToString()
                Dim PO_QTY_REC_SUM As Int64 = Val(ASCDATA1.GetDataValue)
                If PO_QTY_REC_SUM = 0 Then
                    'Note - If there is future based on parameter leave.
                    Dim SC As New System.Text.StringBuilder With {.Length = 0}
                    SC.AppendLine("SELECT COUNT(*) FROM")
                    SC.AppendLine("(")
                    SC.AppendLine("SELECT P2.STYLE_CODE, P2.COLOR_CODE")
                    SC.AppendLine("FROM ICTSTYL1 S1, POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2, POTSHIP1 SP1")
                    SC.AppendLine("WHERE P2.STYLE_CODE = S1.STYLE_CODE")
                    SC.AppendLine("AND P2.PO_ORDER_NO = S3.PO_ORDER_NO")
                    SC.AppendLine("AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO")
                    SC.AppendLine("AND S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
                    SC.AppendLine("AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO")
                    SC.AppendLine("AND SP1.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
                    SC.AppendLine("AND PO_SHIP_STATUS = 'O'")
                    If CDate(dte0.DateTime.ToShortDateString) > CDate(dteInTranAsNow.DateTime.ToShortDateString) Then
                        SC.AppendLine("AND (NVL(SP1.PO_SHIP_ETA,'01-JAN-2100') + PO_SHIP_LANDING_LEAD_DAYS) <= '" & Format(dte0.DateTime, "dd-MMM-yyyy") & "'")
                    Else
                        SC.AppendLine("AND (NVL(SP1.PO_SHIP_ETA,'01-JAN-2100') + PO_SHIP_LANDING_LEAD_DAYS) <= '" & Format(dteInTranAsNow.DateTime, "dd-MMM-yyyy") & "'")
                    End If
                    SC.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE))
                    SC.AppendLine(String.Format("AND P2.COLOR_CODE = '{0}'", COLOR_CODE))
                    SC.AppendLine("GROUP BY P2.STYLE_CODE, COLOR_CODE")
                    SC.AppendLine("UNION")
                    SC.AppendLine("SELECT P2.STYLE_CODE, P2.COLOR_CODE")
                    SC.AppendLine("FROM ICTSTYL1 S1, POTORDR2 P2")
                    SC.AppendLine("WHERE P2.STYLE_CODE = S1.STYLE_CODE")
                    SC.AppendLine("AND P2.PO_STATUS = 'O'")
                    SC.AppendLine(String.Format("AND (NVL(P2.PO_DATE_ETA,'01-JAN-2100') + 10) <= '{0}'", Format(dte0.DateTime, "dd-MMM-yyyy")))
                    SC.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE))
                    SC.AppendLine(String.Format("AND P2.COLOR_CODE = '{0}'", COLOR_CODE))
                    SC.AppendLine("GROUP BY P2.STYLE_CODE, COLOR_CODE")
                    SC.AppendLine(")")
                    ASCMAIN1.sql = SC.ToString()
                    Dim F_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                    If F_CNT = 0 Then
                        Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{0}'", STYLE_CODE, COLOR_CODE)
                        Dim rowICTSTYCX As DataRow = dst.Tables("ICTSTYCX").Select(filter).FirstOrDefault
                        If Not IsNothing(rowICTSTYCX) Then
                            rowICTSTYCX.Delete()
                        End If
                        rowICTSTYC1.Delete()
                    End If
                    'Note - End.
                End If
            Next
        End If


        MsgBox(CStr(NEW_STYLES) & " New Styles Added", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    'Private Sub btnRefreshStyles_Click(sender As System.Object, e As System.EventArgs) Handles btnRefreshStyles.Click
    '    Dim eMsg As String = ""
    '    If chkAgedPast.Checked Then
    '        If chkAgedLess.Checked = False And chkAgedGreater.Checked = False Then
    '            eMsg = vbCrLf & "You Must Pick An Aging Range."
    '            Exit Sub
    '        End If
    '    End If
    '    If dst.Tables("ICTQUOT2").Rows.Count > 0 Then
    '        eMsg = vbCrLf & "You Can Only Load New Styles On New Quotes."
    '    End If
    '    If eMsg.Length > 0 Then
    '        MsgBox(eMsg.ToString, vbOKOnly, "Problem With Selections")
    '        Exit Sub
    '    End If
    '    Absc1.Get_SQL("*") ' ,"ICTQUOTV")

    '    Dim Sql As String = "(Select Distinct STYLE_CODE from ICTSTYL1) ICTQUOTV"
    '    Dim sqlwhere As String = ""
    '    If optASN.Value = "S" Then
    '        sqlwhere = " and ICTSTYL1.CUST_CODE is Null"
    '    ElseIf optASN.Value = "N" Then
    '        sqlwhere = " and ICTSTYL1.CUST_CODE is NOT Null"
    '    End If
    '    ASCMAIN1.sql = "Select ICTQUOTV.STYLE_CODE from  " & SQL & Absc1.sql_TABLE_NAMEs & ASCMAIN1.SQL_Add_WHERE(Replace(Absc1.sql_JOIN, "ICTQUOTV.SUB_BODY_CODE", "ICTSTYL1.SUB_BODY_CODE") & Absc1.sql_WHERE & sqlwhere)

    '    Dim NEW_STYLES As Integer = 0
    '    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
    '    For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "STYLE_CODE")
    '        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
    '        Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE})
    '        If rowICTQUOT2 Is Nothing Then
    '            ASCMAIN1.Progress("Style", STYLE_CODE)
    '            Dim AddStyle As Boolean = False
    '            Dim AgedPast As Boolean = False
    '            If chkAgedPast.Checked Then
    '                Dim GreaterDate As Date = Now
    '                If chkAgedGreater.Checked Then
    '                    GreaterDate = Now.AddDays(numAgedGreater.Value * -1)
    '                End If
    '                Dim GreaterDateF As String = Format(GreaterDate, "dd-MMM-yyyy")

    '                Dim LessDate As Date = Now
    '                If chkAgedLess.Checked Then
    '                    LessDate = Now.AddDays(numAgedLess.Value * -1)
    '                End If
    '                Dim LessDateF As String = Format(LessDate, "dd-MMM-yyyy")

    '                SQLS.Length = 0
    '                SQLS.AppendLine("SELECT COUNT(*) FROM (")
    '                SQLS.AppendLine(" SELECT POTORDR2.STYLE_CODE")
    '                SQLS.AppendLine(" FROM POTORDR2, POTSHIP3, POTSHIP2, ICTSTAT2")
    '                SQLS.AppendLine(" WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
    '                SQLS.AppendLine(" AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO")
    '                SQLS.AppendLine(" AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO")
    '                SQLS.AppendLine(" AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO")
    '                SQLS.AppendLine(" AND POTORDR2.STYLE_CODE = ICTSTAT2.STYLE_CODE")
    '                SQLS.AppendLine(" AND POTORDR2.COLOR_CODE = ICTSTAT2.COLOR_CODE")
    '                SQLS.AppendLine(" AND NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) > 0")
    '                SQLS.AppendLine(" AND (NVL(I2.WHSE_QTY_TRAN,0) + NVL(I2.WHSE_QTY_ON_ORDER,0)) <> 0")
    '                SQLS.AppendLine(" GROUP BY POTORDR2.STYLE_CODE")
    '                If chkAgedLess.Checked And chkAgedGreater.Checked Then
    '                    SQLS.AppendLine(String.Format(" HAVING (MAX(POTSHIP2.PO_DATE_RECEIVED) >= '{0}' AND MAX(POTSHIP2.PO_DATE_RECEIVED) <= '{1}')", LessDateF, GreaterDateF))
    '                Else
    '                    If chkAgedGreater.Checked Then
    '                        SQLS.AppendLine(String.Format(" HAVING MAX(POTSHIP2.PO_DATE_RECEIVED) <= '{0}'", GreaterDateF))
    '                    Else
    '                        SQLS.AppendLine(String.Format(" HAVING MAX(POTSHIP2.PO_DATE_RECEIVED) >= '{0}'", LessDateF))
    '                    End If
    '                End If
    '                SQLS.AppendLine(String.Format(") WHERE STYLE_CODE = '{0}'", STYLE_CODE))
    '                ASCMAIN1.sql = SQLS.ToString()
    '                AgedPast = Val(ASCDATA1.GetDataValue) > 0
    '                If AgedPast Then
    '                    AddStyle = True
    '                End If
    '            Else
    '                SQLS.Length = 0
    '                SQLS.AppendLine("SELECT COUNT(*) FROM (")
    '                SQLS.AppendLine("SELECT P2.STYLE_CODE")
    '                SQLS.AppendLine("FROM POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2, ICTSTAT2 I2")
    '                SQLS.AppendLine("WHERE S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
    '                SQLS.AppendLine("AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO")
    '                SQLS.AppendLine("AND P2.PO_ORDER_NO = S3.PO_ORDER_NO")
    '                SQLS.AppendLine("AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO")
    '                SQLS.AppendLine("AND P2.STYLE_CODE = I2.STYLE_CODE")
    '                SQLS.AppendLine("AND P2.COLOR_CODE = I2.COLOR_CODE")
    '                If chkRECDATES.Checked Then
    '                    If chkInclWIPTRAN.Checked Then
    '                        SQLS.AppendLine(String.Format("AND ((NVL(S2.PO_DATE_RECEIVED,'01-JAN-1900') >= '{0}'", Format(dteRECDATEFR.DateTime, "dd-MMM-yyyy")))
    '                        SQLS.AppendLine(String.Format("AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2900') <= '{0}')", Format(dteRECDATETO.DateTime, "dd-MMM-yyyy")))
    '                        SQLS.AppendLine("OR (NVL(I2.WHSE_QTY_TRAN,0) + NVL(I2.WHSE_QTY_ON_ORDER,0)) <> 0)")
    '                    Else
    '                        SQLS.AppendLine(String.Format("AND (NVL(S2.PO_DATE_RECEIVED,'01-JAN-1900') >= '{0}'", Format(dteRECDATEFR.DateTime, "dd-MMM-yyyy")))
    '                        SQLS.AppendLine(String.Format("AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2900') <= '{0}')", Format(dteRECDATETO.DateTime, "dd-MMM-yyyy")))
    '                    End If
    '                End If
    '                If chkLIMIT_NETPOS.Checked And chkLIMIT_NETPOS_G.Checked Then
    '                    SQLS.AppendLine(String.Format("AND (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) >= {0}", Val(numLIMIT_NETPOS_G.Value & String.Empty)))
    '                End If
    '                If chkLIMIT_NETPOS.Checked And chkLIMIT_NETPOS_L.Checked Then
    '                    SQLS.AppendLine(String.Format("AND (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) <= {0}", Val(numLIMIT_NETPOS_L.Value & String.Empty)))
    '                End If
    '                SQLS.AppendLine("GROUP BY P2.STYLE_CODE")
    '                SQLS.AppendLine(String.Format(") WHERE STYLE_CODE = '{0}'", STYLE_CODE))
    '                ASCMAIN1.sql = SQLS.ToString()
    '                If Val(ASCDATA1.GetDataValue) > 0 Then
    '                    AddStyle = True
    '                End If
    '            End If
    '            If AddStyle Then
    '                NEW_STYLES += 1
    '                Add_to_Quote(STYLE_CODE)
    '            End If
    '        End If
    '    Next

    '    RemoveBadColors()

    '    MsgBox(CStr(NEW_STYLES) & " New Styles Added", MsgBoxStyle.OkOnly, "Verification")

    'End Sub

    Private Sub RemoveBadColors()
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select()
            Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty
            'If STYLE_CODE = "803089IZ" Then Stop
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT COUNT(*) FROM (")
            SQLS.AppendLine("SELECT P2.STYLE_CODE, P2.COLOR_CODE")
            SQLS.AppendLine("FROM POTORDR2 P2, POTSHIP3 S3, POTSHIP2 S2, ICTSTAT2 I2")
            SQLS.AppendLine("WHERE S3.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO")
            SQLS.AppendLine("AND S3.PO_SHIPMENT_LNO = S2.PO_SHIPMENT_LNO")
            SQLS.AppendLine("AND P2.PO_ORDER_NO = S3.PO_ORDER_NO")
            SQLS.AppendLine("AND P2.PO_ORDER_LNO = S3.PO_ORDER_LNO")
            SQLS.AppendLine("AND P2.STYLE_CODE = I2.STYLE_CODE")
            SQLS.AppendLine("AND P2.COLOR_CODE = I2.COLOR_CODE")
            If chkRECDATES.Checked Then
                If chkInclWIPTRAN.Checked Then
                    SQLS.AppendLine(String.Format("AND ((NVL(S2.PO_DATE_RECEIVED,'01-JAN-1900') >= '{0}'", Format(dteRECDATEFR.DateTime, "dd-MMM-yyyy")))
                    SQLS.AppendLine(String.Format("AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2900') <= '{0}')", Format(dteRECDATETO.DateTime, "dd-MMM-yyyy")))
                    SQLS.AppendLine("OR (NVL(I2.WHSE_QTY_TRAN,0) + NVL(I2.WHSE_QTY_ON_ORDER,0)) <> 0)")
                Else
                    SQLS.AppendLine(String.Format("AND (NVL(S2.PO_DATE_RECEIVED,'01-JAN-1900') >= '{0}'", Format(dteRECDATEFR.DateTime, "dd-MMM-yyyy")))
                    SQLS.AppendLine(String.Format("AND NVL(S2.PO_DATE_RECEIVED,'01-JAN-2900') <= '{0}')", Format(dteRECDATETO.DateTime, "dd-MMM-yyyy")))
                End If
            End If
            If chkLIMIT_NETPOS.Checked And chkLIMIT_NETPOS_G.Checked Then
                SQLS.AppendLine(String.Format("AND (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) >= {0}", Val(numLIMIT_NETPOS_G.Value & String.Empty)))
            End If
            If chkLIMIT_NETPOS.Checked And chkLIMIT_NETPOS_L.Checked Then
                SQLS.AppendLine(String.Format("AND (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) <= {0}", Val(numLIMIT_NETPOS_L.Value & String.Empty)))
            End If
            SQLS.AppendLine("GROUP BY P2.STYLE_CODE, P2.COLOR_CODE")
            SQLS.AppendLine(String.Format(") WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQLS.ToString()
            If Val(ASCDATA1.GetDataValue) = 0 Then
                rowICTSTYC1.Delete()
            End If
        Next
    End Sub

    Private Sub tabStyles_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabStyles.SelectedTabChanged
        Setup_tabStyles()
    End Sub

    Sub Setup_tabStyles()
        If Me.SELECTION_NO = 0 Then Exit Sub
        If tabStyles.SelectedTab Is Nothing Then Exit Sub

        With UltraExplorerBar1
            .Groups("Style Image").Visible = (tabStyles.SelectedTab.Key = "Styles")
            .Groups("Availability Dates").Visible = Not (tabStyles.SelectedTab.Key = "Styles")
            .Groups("Style Color Filter").Visible = (tabStyles.SelectedTab.Key = "Styles && Colors")
        End With
    End Sub

    Function FetchImage(IMAGE_NAME As String, STYLE_CODE As String) As Byte()
        Dim imgba() As Byte = Nothing
        If IMAGE_NAME <> "" Then
            If Not IsNothing(imgSTYLE.Image) Then
                imgSTYLE.Image.dispose
            End If
            imgSTYLE.Image = Get_Style_Image(IMAGE_NAME)
            UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE & "-" & COLOR_CODE
        Else
            imgSTYLE.Image = Nothing
            UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        End If
        Return imgba
    End Function

    Private Sub grdICTQUOT2B_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTQUOT2B.AfterRowActivate
        If grdICTQUOT2B.ActiveRow Is Nothing OrElse Not grdICTQUOT2B.ActiveRow.IsDataRow Then
        Else
            If Not PRINTING_SHEETS Then
                FetchImage(grdICTQUOT2B.ActiveRow.Cells("IMAGE_NAME").Value & "", grdICTQUOT2B.ActiveRow.Cells("STYLE_CODE_PLM").Value)
            End If
        End If
    End Sub

    Private Sub chk1Sheet_CheckedChanged(sender As Object, e As System.EventArgs) Handles chk1Sheet.CheckedChanged
        '   opt1Sheet.Enabled = chk1Sheet.Checked
    End Sub

    Sub RESEQ()

        Dim SEQ As Integer = 0

        Dim sqlWB As String = ""
        'If chkSepDivision.Checked <> "" Then
        sqlWB = ",SALES_DIVISION_CODE"
        'End If
        If opt1Sheet.Value = "S" Then
            sqlWB &= ",SUB_BODY_CODE"
        ElseIf opt1Sheet.Value = "FS" Then
            sqlWB &= ",FABRIC_CODE,SUB_BODY_CODE,STYLE_GROUP_CODE"
        ElseIf opt1Sheet.Value = "G" Then
            sqlWB &= ",STYLE_GROUP_CODE,FABRIC_CODE,SUB_BODY_CODE"

        End If
        sqlWB &= ",STYLE_CODE_PLM"

        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("", Mid(sqlWB, 2))
            SEQ += 10
            row.Item("SEQ") = SEQ
        Next
    End Sub

    Private Sub chkNowOHonly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkNowOHonly.CheckedChanged
        refresh_required = True
        cmdGetAvailability.Appearance.ForeColor = Color.Red
        chkALLOSTDT.Visible = Not chkNowOHonly.Checked
        cboStartPeriod.Visible = Not chkNowOHonly.Checked
        dte0.Visible = Not chkNowOHonly.Checked
        dte1.Visible = Not chkNowOHonly.Checked
        dte2.Visible = Not chkNowOHonly.Checked
        dte3.Visible = Not chkNowOHonly.Checked
        lblNOW.Visible = Not chkNowOHonly.Checked
        lbl1.Visible = Not chkNowOHonly.Checked
        lbl2.Visible = Not chkNowOHonly.Checked
        lbl3.Visible = Not chkNowOHonly.Checked
        chkBeyond.Visible = Not chkNowOHonly.Checked
        tkb1.Visible = Not chkNowOHonly.Checked
    End Sub

    Private Sub chkBeyond_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBeyond.CheckedChanged
        refresh_required = True
        cmdGetAvailability.Appearance.ForeColor = Color.Red
    End Sub

    Private Sub dte0_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dte0.ValueChanged
        refresh_required = True
        cmdGetAvailability.Appearance.ForeColor = Color.Red
    End Sub

    Private Sub dte1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dte1.ValueChanged
        refresh_required = True
        cmdGetAvailability.Appearance.ForeColor = Color.Red
    End Sub

    Private Sub dte2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dte2.ValueChanged
        refresh_required = True
        cmdGetAvailability.Appearance.ForeColor = Color.Red
    End Sub

    Private Sub dte3_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dte3.ValueChanged
        refresh_required = True
        cmdGetAvailability.Appearance.ForeColor = Color.Red
    End Sub

    Private Sub txtQS_STYLE_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtQS_STYLE_CODE.ValueChanged

    End Sub

    Private Sub chk1perPage_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chk1perPage.CheckedChanged

    End Sub

    Private Sub UltraButton1_Click(sender As System.Object, e As System.EventArgs) Handles CMDGETDIV.Click

        dst.Tables("SOTSDIVC").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct("ICTQUOT2", "SALES_DIVISION_CODE").Rows
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE") & ""
            If SALES_DIVISION_CODE <> "" Then

                Dim rowSOTSDIVC As DataRow = dst.Tables("SOTSDIVC").NewRow
                rowSOTSDIVC.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                rowSOTSDIVC.Item("SALES_DIVISION_CODE_COMB") = SALES_DIVISION_CODE
                dst.Tables("SOTSDIVC").Rows.Add(rowSOTSDIVC)

            End If
        Next
    End Sub

    Private Sub chkALLOSTDT_CheckedChanged(sender As Object, e As EventArgs) Handles chkALLOSTDT.CheckedChanged
        If chkALLOSTDT.Checked Then
            chkALLOSTDT.Width = 44
            chkALLOSTDT.Text = "SD"
            cboStartPeriod.Visible = True
            cboStartPeriod.SelectedIndex = 0
            lblNOW.Text = "SD"
        Else
            chkALLOSTDT.Width = 100
            chkALLOSTDT.Text = "Start Date"
            cboStartPeriod.Visible = False
            cboStartPeriod.SelectedIndex = 0
            lblNOW.Text = "Now"
        End If
    End Sub

    Private Sub btnByDate_Click(sender As Object, e As EventArgs) Handles btnByDate.Click
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("S.STYLE_CODE,")
        sql.AppendLine("S.STYLE_DESC,")
        sql.AppendLine("MAX(P.INIT_DATE) AS ENTERED")
        sql.AppendLine("FROM POTORDR2 P, ICTSTYL1 S")
        sql.AppendLine("WHERE P.STYLE_CODE = S.STYLE_CODE")
        sql.AppendLine("AND NVL(P.INIT_DATE,'01-JAN-1900') <> '01-JAN-1900'")
        sql.AppendLine("GROUP BY S.STYLE_CODE,")
        sql.AppendLine("S.STYLE_DESC")
        sql.AppendLine("ORDER BY MAX(P.INIT_DATE) DESC, STYLE_CODE")

        With ASCMAIN1.CodeSelector
            .SQL = sql.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Select By Date"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
        End With

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

        Sort_by_Style()
    End Sub

    Private Sub chkRECDATES_CheckedChanged(sender As Object, e As EventArgs) Handles chkRECDATES.CheckedChanged
        If Not Form_Loading Then
            If chkRECDATES.Checked Then
                chkAgedPast.Checked = False
                grpRECDATES.Visible = True
                chkInclWIPTRAN.Checked = True
                dteRECDATETO.DateTime = Now
                dteRECDATEFR.DateTime = Now
            Else
                grpRECDATES.Visible = False
                chkInclWIPTRAN.Checked = False
            End If
        End If
    End Sub

    Private Sub chkAgedPast_CheckedChanged(sender As Object, e As EventArgs) Handles chkAgedPast.CheckedChanged
        If Not Form_Loading Then
            If chkAgedPast.Checked Then
                chkRECDATES.Checked = False
                chkLIMIT_NETPOS.Checked = False
                grpAgedPast.Visible = True
                numAgedGreater.Value = 0
                numAgedLess.Value = 180
            Else
                grpAgedPast.Visible = False
            End If
        End If
    End Sub

    Private Sub chkShowLastRcd_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowLastRcd.CheckedChanged
        If Not Form_Loading Then
            If chkShowLastRcd.Checked Then
                setLastRcdDate()
            Else
                grdICTQUOT2.DisplayLayout.Bands(1).Columns("LAST_RCD_DATE").Hidden = True
            End If
        End If
    End Sub

    Private Sub chkAgedGreater_CheckedChanged(sender As Object, e As EventArgs) Handles chkAgedGreater.CheckedChanged
        If Not Form_Loading Then
            If chkAgedGreater.Checked Then
                numAgedGreater.Value = 0
            Else
                numAgedGreater.Value = Null
            End If
        End If
    End Sub

    Private Sub chkAgedLess_CheckedChanged(sender As Object, e As EventArgs) Handles chkAgedLess.CheckedChanged
        If Not Form_Loading Then
            If chkAgedLess.Checked Then
                numAgedLess.Value = 180
            Else
                numAgedLess.Value = Null
            End If
        End If
    End Sub

    Private Sub chkLIMIT_NETPOS_CheckedChanged(sender As Object, e As EventArgs) Handles chkLIMIT_NETPOS.CheckedChanged
        If Not Form_Loading Then
            If chkLIMIT_NETPOS.Checked Then
                chkAgedPast.Checked = False
                grpLIMIT_NETPOS.Visible = True
                numLIMIT_NETPOS_G.Value = 0
                numLIMIT_NETPOS_L.Value = 100
            Else
                grpLIMIT_NETPOS.Visible = False
                numLIMIT_NETPOS_G.Value = 0
                numLIMIT_NETPOS_L.Value = 100
            End If
        End If
    End Sub

    Private Sub chkScaleImages_CheckedChanged(sender As Object, e As EventArgs) Handles chkScaleImages.CheckedChanged
        If chkScaleImages.Checked Then
            trkScaleImage.Visible = True
        Else
            trkScaleImage.Visible = False
            trkScaleImage.Value = 100
        End If
    End Sub

    Private Sub trkScaleImage_ValueChanged(sender As Object, e As EventArgs) Handles trkScaleImage.ValueChanged
        chkScaleImages.Text = String.Format("Scale Images ({0}%)", trkScaleImage.Value)
    End Sub

    Private Sub numMinQty_ValueChanged(sender As Object, e As EventArgs)
        If IsNumeric(numMinQty.Value) Then
            If numMinQty.Value <= 0 Then
                chkExcudeFutureWhenZero.Checked = True
            Else
                chkExcudeFutureWhenZero.Checked = False
            End If
        Else
            chkExcudeFutureWhenZero.Checked = False
        End If
    End Sub

    Private Sub btnSCFilter_Click(sender As Object, e As EventArgs) Handles btnSCFilter.Click
        Dim filter As String = ""
        If txtSCFilter.Text.Length > 0 Then
            filter = String.Format("STYLE_CODE_PLM = '{0}'", txtSCFilter.Text)
        End If
        Dim dvw As DataView = DirectCast(grdICTQUOT2.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Format(filter)
    End Sub

    Private Sub btnSCShowAll_Click(sender As Object, e As EventArgs) Handles btnSCShowAll.Click
        Dim dvw As DataView = DirectCast(grdICTQUOT2.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Format("")
        txtSCFilter.Text = ""
    End Sub

    Private Sub chkBADates_CheckedChanged(sender As Object, e As EventArgs) Handles chkBADates.CheckedChanged
        If chkBADates.Checked Then
            grpBADates.Visible = True
            chkBABefore.Checked = False
            chkBAAfter.Checked = False
        Else
            grpBADates.Visible = False
            chkBABefore.Checked = False
            chkBAAfter.Checked = False
        End If
    End Sub

    Private Sub chkBABefore_CheckedChanged(sender As Object, e As EventArgs) Handles chkBABefore.CheckedChanged
        If chkBABefore.Checked Then
            dteBABefore.Visible = True
        Else
            dteBABefore.Visible = False
        End If
        dteBABefore.Value = Null
    End Sub

    Private Sub chkBAAfter_CheckedChanged(sender As Object, e As EventArgs) Handles chkBAAfter.CheckedChanged
        If chkBAAfter.Checked Then
            dteBAAfter.Visible = True
        Else
            dteBAAfter.Visible = False
        End If
        dteBAAfter.Value = Null
    End Sub
End Class

Public Module ImageUtils
    Public Function ResizeImage(ByVal image As Image,
  ByVal size As Size, Optional ByVal preserveAspectRatio As Boolean = True) As Image

        Dim newWidth As Integer
        Dim newHeight As Integer
        If preserveAspectRatio Then
            Dim originalWidth As Integer = image.Width
            Dim originalHeight As Integer = image.Height
            Dim percentWidth As Single = CSng(size.Width) / CSng(originalWidth)
            Dim percentHeight As Single = CSng(size.Height) / CSng(originalHeight)
            Dim percent As Single = If(percentHeight < percentWidth, percentHeight, percentWidth)
            newWidth = CInt(originalWidth * percent)
            newHeight = CInt(originalHeight * percent)
        Else
            newWidth = size.Width
            newHeight = size.Height
        End If

        Dim newImage As Image = New Bitmap(newWidth, newHeight)

        Using graphicsHandle As Graphics = Graphics.FromImage(newImage)
            graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic
            graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight)
        End Using

        Return newImage

    End Function
    ' Compression
    Public Sub SaveImageWithQuality(ByVal bmp1 As Image, ByVal destinationPath As String, ByVal quality As Long)
        'Or you can use build-in method
        'Dim jgpEncoder As ImageCodecInfo = GetEncoderInfo("image/jpeg");
        Dim jgpEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)

        ' Create an Encoder object based on the GUID
        ' for the Quality parameter category.
        Dim myEncoder As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.Quality

        ' Create an EncoderParameters object.
        ' An EncoderParameters object has an array of EncoderParameter
        ' objects. In this case, there is only one
        ' EncoderParameter object in the array.
        Dim myEncoderParameters As New EncoderParameters(1)

        ' Save with 100% quality
        Dim myEncoderParameter As New EncoderParameter(myEncoder, quality)
        myEncoderParameters.Param(0) = myEncoderParameter
        bmp1.Save(destinationPath, jgpEncoder, myEncoderParameters)

    End Sub

    Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo

        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()

        Dim codec As ImageCodecInfo
        For Each codec In codecs
            If codec.FormatID = format.Guid Then
                Return codec
            End If
        Next codec
        Return Nothing

    End Function

End Module

