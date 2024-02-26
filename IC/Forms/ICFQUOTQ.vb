Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

Public Class ICFQUOTQ

    'Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim STYLE_CODE_allocated As String
    Dim AutoAllocate As Boolean
    Dim SOTDEMD1 As String
    Dim SOTSUPP1 As String
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
    Dim LOAD_TEMP As String = ""
    Dim S As New System.Text.StringBuilder With {.Length = 0}
    Dim dayBreaks As Integer = 120

    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        'FtpS.RuntimeLicense = "31484E46414431535542323032333033313352415331544531414D483134323600000000000000003335384A30543346000059554A4336594E46335047530000"

        chkNewLinks.Checked = True

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
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "V", IIf(TABLE_NAME = "ICTSTYL1_RECENT", 1, 0))
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
              & " where X.QUOTE_NO = ICTQUOT1.QUOTE_NO and ICTQUOT1.QUOTE_TYPE = 'S'" _
              & " and SOLDOUT_OPT = '2'"
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
                .Columns.Add("EXCLUDE_STYLE")
                '.Columns.Add("LAST_SHIP_DATE")
            End With

            S.Length = 0
            S.AppendLine("SELECT * FROM ICTQUOT3 WHERE ICTQUOT3.QUOTE_NO = :PARM1")
            ASCMAIN1.sql = S.ToString
            'Create_TDA(.Tables.Add, "ICTQUOT3", "*")
            Create_TDA(.Tables.Add, "ICTQUOT3", "**", 0, True, "V", 3)
            With .Tables("ICTQUOT3")
                .Columns.Add("EXCLUDE_COLOR")
                .Columns.Add("WIP_TRAN", GetType(System.Int64))
            End With

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

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("T2.STYLE_CODE,")
            S.AppendLine("T2.COLOR_CODE,")
            S.AppendLine("C1.COLOR_DESC,")
            S.AppendLine("SUM(CASE")
            S.AppendLine("     WHEN ROUND((SYSDATE - T1.TRAN_DATE)) <= :PARM1")
            S.AppendLine("     THEN T2.TRAN_QTY")
            S.AppendLine("     ELSE 0")
            S.AppendLine("END) AS RECEIVED_01,")
            S.AppendLine("SUM(CASE")
            S.AppendLine("     WHEN ROUND((SYSDATE - T1.TRAN_DATE)) > :PARM1 AND ROUND((SYSDATE - T1.TRAN_DATE)) <= (:PARM1 * 2)")
            S.AppendLine("     THEN T2.TRAN_QTY")
            S.AppendLine("     ELSE 0")
            S.AppendLine("END) AS RECEIVED_02,")
            S.AppendLine("SUM(CASE")
            S.AppendLine("     WHEN ROUND((SYSDATE - T1.TRAN_DATE)) > (:PARM1 * 2)")
            S.AppendLine("     THEN T2.TRAN_QTY")
            S.AppendLine("     ELSE 0")
            S.AppendLine("END) AS RECEIVED_03")
            S.AppendLine("FROM ICTTRAN1 T1, ICTTRAN2 T2, ICTCOLR1 C1")
            S.AppendLine("WHERE T1.OPS_YYYYPP = T2.OPS_YYYYPP")
            S.AppendLine("AND T1.TRAN_TYPE = T2.TRAN_TYPE")
            S.AppendLine("AND T1.TRAN_NO = T2.TRAN_NO")
            S.AppendLine("AND T1.TRAN_TYPE = 'R'")
            S.AppendLine("AND T2.COLOR_CODE = C1.COLOR_CODE")
            S.AppendLine("GROUP BY")
            S.AppendLine("T2.STYLE_CODE,")
            S.AppendLine("T2.COLOR_CODE,")
            S.AppendLine("C1.COLOR_DESC")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(dst.Tables.Add, "ICTQUOTD", "**", 0, False, "I")

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("ICTSTYL1.STYLE_CODE,")
            S.AppendLine("ICTSTAT2.COLOR_CODE,")
            S.AppendLine("ICTCOLR1.COLOR_DESC,")
            S.AppendLine("SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) AS NET_POS,")
            S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) AS IN_TRANS,")
            S.AppendLine("(SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) - SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0))) AS NOW_OH,")
            S.AppendLine("0 AS RECEIVED_01,")
            S.AppendLine("0 AS RECEIVED_02,")
            S.AppendLine("0 AS RECEIVED_03,")
            S.AppendLine("0 AS AGED_01,")
            S.AppendLine("0 AS AGED_02,")
            S.AppendLine("0 AS AGED_03")
            S.AppendLine("FROM ICTSTYL1, ICTSTAT2, ICTCOLR1")
            S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
            S.AppendLine("AND ICTSTAT2.COLOR_CODE = ICTCOLR1.COLOR_CODE")
            S.AppendLine("HAVING SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) > 0")
            S.AppendLine("GROUP BY")
            S.AppendLine("ICTSTYL1.STYLE_CODE,")
            S.AppendLine("ICTSTAT2.COLOR_CODE,")
            S.AppendLine("ICTCOLR1.COLOR_DESC")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(dst.Tables.Add, "ICTQUOTQ", "**", 0, False)
            With dst.Tables("ICTQUOTQ").Columns
                .Add("LAST_RCD_DATE")
            End With


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

        grdICTQUOT2B.DataSource = dst.Tables("ICTQUOT2")
        grdICTQUOT3.DataSource = dst.Tables("ICTQUOT3")
        grdICTQUOTX.DataSource = dst.Tables("ICTQUOTX")
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

        Create_Summary(grdICTQUOT3, "QTY_AVA1")
        Create_Summary(grdICTQUOT3, "QTY_AVA2")
        Create_Summary(grdICTQUOT3, "QTY_AVA3")

        grdICTSTYCX.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal

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
                Dim rowSEL As Int64 = dst.Tables.Item("ICTQUOT2").Select("SELECTED = '1'").Count
                If rowSEL = 0 Then
                    EMsg &= vbCr & "You Must Select Styles To Include On Excel"
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
                    Dim rowARTCUST1 As DataRow = Lookup("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
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

                CALC_EXCLUDED_COLORS()

                If chkShowLastRcd.Checked Then
                    setLastRcdDate()
                End If

                dst.Tables("ICTSTYC1").Columns("COUNT_COLOR").Expression = String.Format(COUNT_COLOR, 0)

                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Save Excel As Link?"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("Would You Like To Save This Excel As A Link?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                Dim fileName As String = ""
                Dim fileNameB As String = ""
                If chkSepDivision.Checked Then
                    For Each row As DataRow In ASCDATA1.SelectDistinct("ICTQUOT2", New String() {"SALES_DIVISION_CODE"}).Select("")
                        Dim SALES_DIVISION_CODE As String = row.Item(0) & ""
                        fileName = Create_Excel(SALES_DIVISION_CODE, False)
                        If chkBuyerOutput.Checked Then
                            fileNameB = Create_Excel(SALES_DIVISION_CODE, True)
                        End If
                    Next
                Else
                    fileName = Create_Excel("", False)
                    fileNameB = Create_Excel("", True)
                End If

                If iResult = vbYes Then
                    Dim fileCnt As Int64 = 1
                    If fileNameB.Length >= 0 Then
                        fileCnt = 2
                    End If
                    For FC As Int64 = 1 To fileCnt
                        'Dim SESSION_NO As String = ASCMAIN1.Next_Control_No("ICTQUOH1.SESSION_NO")
                        dst.Tables("ASTATTA2").Rows.Clear()
                        ENTITY.TABLE_NAME = "ICFQUOTV"
                        ENTITY.COLUMN_NAME = "QUOTE_NO"
                        ENTITY.CODE_VALUE = QUOTE_NO

                        If FC = 1 Then
                            MyBase.Attach_File(ASCMAIN1.Folders("Temp") & fileName, "Quote Sheet For Quote " & QUOTE_NO,,,, True)
                        Else
                            MyBase.Attach_File(ASCMAIN1.Folders("Temp") & fileNameB, "Quote Sheet For Quote " & QUOTE_NO,,,, True)
                        End If
                        'MyBase.Attach_File(fileName, "Quote Sheet For Quote " & QUOTE_NO,,,, True)

                        'Update ASTATTA2 immediatly.
                        Update_Record_TDA("ASTATTA2")

                        Dim MaxFILE_NO As Int64 = Val(dst.Tables("ICTQUOHF").Compute("Max(FILE_NO)", "") & "")
                        Dim SESSION_NO As String = dst.Tables("ICTQUOHF").Compute("Max(SESSION_NO)", "") & ""
                        MaxFILE_NO += 1

                        Dim newICTQUOHF As DataRow = dst.Tables.Item("ICTQUOHF").NewRow
                        Dim rowASTATTA2 As DataRow = dst.Tables.Item("ASTATTA2").Select().FirstOrDefault
                        Dim ATTACHMENT_FILENAME As String = rowASTATTA2.Item("ATTACHMENT_FILENAME").ToString
                        Dim ext As String = getFileExt(ATTACHMENT_FILENAME)

                        newICTQUOHF.Item("SESSION_NO") = SESSION_NO
                        newICTQUOHF.Item("FILE_NO") = MaxFILE_NO
                        newICTQUOHF.Item("FILENAME") = ATTACHMENT_FILENAME
                        newICTQUOHF.Item("HASHVALUE") = rowASTATTA2.Item("HASHVALUE").ToString
                        dst.Tables.Item("ICTQUOHF").Rows.Add(newICTQUOHF)

                        Dim FileNameLocalFull As String = rowASTATTA2.Item("ATTACHMENT_FILENAME").ToString
                        Dim FileNameRemote As String = $"{SESSION_NO}-{MaxFILE_NO}{ext}"
                        Dim eMsg As Text.StringBuilder = FTP_BLUEHOST(FileNameLocalFull, FileNameRemote)
                        If eMsg.Length > 0 Then
                            MsgBox(eMsg.ToString, vbCritical, "Error Sending To Remote Server")
                        End If
                    Next
                End If
            Case "Print", "email"
                Update_Record(True)
                IMG_Error_Reported = False
                PRINTING_SHEETS = True
                'Print_Style_Sheet(eItemKey)
                'This has to be re-worked because the quotes are all forward thinking allocaitons and htis report is all about historical receipts
                'You will find the old code remmed out
                PRINTING_SHEETS = False
            Case "Clear"
                dst.Tables("ICTQUOT2").Rows.Clear()
                dst.Tables("ICTQUOT3").Rows.Clear()
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
                .Items("Edit").Visible = Not (EntryMode = "N" Or EntryMode = "E")
                .Items("View").Visible = Not ScreenMode
                .Items("Excel").Visible = ScreenMode
                .Items("Done").Visible = ScreenMode And (EntryMode = "V")
                .Items("Save").Visible = ScreenMode And (EntryMode = "V")
                .Items("Print").Visible = False
                .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
            End With
        End With

        ' If Not ScreenMode Then Setup_QuoteSheet()
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'cmdGetAvailability.Visible = ScreenMode


        cmdAddMultipleStyles.Visible = (EntryMode = "N" Or EntryMode = "E")
        lblQS_STYLE_CODE.Visible = (EntryMode = "N" Or EntryMode = "E")
        txtQS_STYLE_CODE.Visible = (EntryMode = "N" Or EntryMode = "E")

        chkShowSelectedOnly.Checked = False

        If ScreenMode Then
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
            rowICTQUOT1.Item("QUOTE_SHOW_LAST_REC") = "1"
            rowICTQUOT1.Item("SOLDOUT_OPT") = "2"
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
        Fill_Records("ICTQUOT3", QUOTE_NO)
        FILL_WIPTRANS()

        Fill_Records("ICTCOLR1")

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

        ' Retrieve_Settings(QUOTE_NO)

        Setup_tabStyles()

        refresh_required = False

        If chkShowLastRcd.Checked Then
            setLastRcdDate()
        End If

        Setup_ICTQUOT3()

        setSizeScales()

        If dst.Tables.Item("ICTQUOT2").Rows.Count = 0 Then
            btnLoadStyles.Enabled = True
            btnRefreshStyles.Enabled = False
        Else
            btnLoadStyles.Enabled = False
            btnRefreshStyles.Enabled = True
        End If
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub setSizeScales()
        dst.Tables.Item("ICTSTYLX").Clear()
        dst.Tables.Item("ICTSTYC1").Clear()

        For Each rowICTQUOT2 As DataRow In dst.Tables("ICTQUOT2").Select()
            ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC, SIZE_SCALE" & vbCrLf _
        & " from ICTSTYL1" & vbCrLf _
        & " where STYLE_CODE = '" & rowICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty & "'"
            Fill_Records("ICTSTYLX", "", False, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, STYLE_COLOR_DESC" & vbCrLf _
                    & " from ICTSTYC1" & vbCrLf _
                    & " where ICTSTYC1.STYLE_CODE = '" & rowICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty & "'"
            Fill_Records("ICTSTYCX", "", False, ASCMAIN1.sql)

            'Fix_Colors(STYLE_CODE)
            Do While Fix_Colors(rowICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)

            Loop
            Fix_Size(rowICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)

            Dim rowICTSTYLX As DataRow = dst.Tables("ICTSTYLX").Rows.Find(rowICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)
            Dim SQ As String = ""
            For I As Integer = 1 To 12
                If rowICTSTYLX.Item("S" & CStr(I)) & "" <> "" Then
                    SQ &= " " & rowICTSTYLX.Item("S" & CStr(I)) & "/" & CStr(rowICTSTYLX.Item("Q" & CStr(I)))
                Else
                    Exit For
                End If
            Next
            rowICTSTYLX.Item("SQ") = Mid(SQ, 2)
            'rowICTQUOT2.Item("SIZE_SCALE") = rowICTSTYLX.Item("SIZE_SCALE")
            rowICTQUOT2.Item("SIZE_SCALE") = GET_ONLY_SIZE_SCALE(rowICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)

            Fill_Records("ICTSTYC1", rowICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty, False)
        Next
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
        Update_Record_TDA("ICTQUOT3", "QUOTE_NO = '" & QUOTE_NO & "'")
        Save_Settings(QUOTE_NO, txtQUOTE_DESC.Text)

        'dst.Tables("ICTQUOT3").Rows.Clear()
        'For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("")
        '    Dim rowICTQUOT3 As DataRow = dst.Tables("ICTQUOT3").NewRow
        '    For Each dcol As DataColumn In dst.Tables("ICTQUOT3").Columns
        '        If dcol.ColumnName = "QUOTE_NO" Then
        '            rowICTQUOT3.Item(dcol.ColumnName) = QUOTE_NO
        '        Else
        '            rowICTQUOT3.Item(dcol.ColumnName) = rowICTSTYC1.Item(dcol.ColumnName)
        '        End If
        '    Next
        '    dst.Tables("ICTQUOT3").Rows.Add(rowICTQUOT3)
        'Next
        'Update_Record_TDA("ICTQUOT3", "QUOTE_NO = '" & QUOTE_NO & "'")


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
                    Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
                    If rowICTSTYL1 IsNot Nothing Then
                        Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                    End If
                End If

            Case "Style Master"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
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
                If chkNewLinks.Checked Then
                    Dim FILENAME As String = grd.ActiveRow.Cells.Item("FILENAME").Text
                    Dim ext As String = getFileExt(FILENAME)
                    Dim SESSION_NO As String = grd.ActiveRow.Cells.Item("SESSION_NO").Text
                    Dim FILE_NO As String = grd.ActiveRow.Cells.Item("FILE_NO").Text
                    Dim LINEPFX As String = $"https://docs.vandalequotes.com/{SESSION_NO}-{FILE_NO}{ext}"
                    My.Computer.Clipboard.SetText(FILENAME & vbCrLf & LINEPFX)
                Else
                    Dim FILENAME As String = grd.ActiveRow.Cells.Item("FILENAME").Text
                    Dim HASH As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                    Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                    My.Computer.Clipboard.SetText(FILENAME & vbCrLf & LINEPFX)
                End If
            Case "Copy All Links"
                Dim clipbrd As String = ""
                If chkNewLinks.Checked Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdICTQUOHF.Rows
                        Dim FILENAME As String = grow.Cells.Item("FILENAME").Text
                        Dim ext As String = getFileExt(FILENAME)
                        Dim SESSION_NO As String = grow.Cells.Item("SESSION_NO").Text
                        Dim FILE_NO As String = grow.Cells.Item("FILE_NO").Text
                        Dim LINEPFX As String = $"https://docs.vandalequotes.com/{SESSION_NO}-{FILE_NO}.{ext}"

                        clipbrd = clipbrd & FILENAME & vbCrLf & LINEPFX & vbCrLf & vbCrLf
                    Next
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grdICTQUOHF.Rows
                        Dim FILENAME As String = grow.Cells.Item("FILENAME").Text
                        Dim HASH As String = grow.Cells.Item("HASHVALUE").Text
                        Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                        clipbrd = clipbrd & FILENAME & vbCrLf & LINEPFX & vbCrLf & vbCrLf
                    Next
                End If
                My.Computer.Clipboard.SetText(clipbrd)
                'Case "Copy Link"
                '    Dim FILENAME As String = grd.ActiveRow.Cells.Item("FILENAME").Text
                '    Dim HASH As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                '    Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                '    My.Computer.Clipboard.SetText(FILENAME & vbCrLf & LINEPFX)
                'Case "Copy All Links"
                '    Dim clipbrd As String = ""
                '    For Each grow As UltraWinGrid.UltraGridRow In grdICTQUOHF.Rows
                '        Dim FILENAME As String = grow.Cells.Item("FILENAME").Text
                '        Dim HASH As String = grow.Cells.Item("HASHVALUE").Text
                '        Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                '        clipbrd = clipbrd & FILENAME & vbCrLf & LINEPFX & vbCrLf & vbCrLf
                '    Next
                '    My.Computer.Clipboard.SetText(clipbrd)
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
            Application.DoEvents()
            Try
                'newBMP.MakeTransparent(System.Drawing.Color.White)
                Dim converter As New ImageConverter
                row.Item("IMAGE") = converter.ConvertTo(newBMP, GetType(Byte()))
                newBMP.Dispose()
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

#End Region

#Region "Custom Methods"

    'Function Add_to_Quote(STYLE_CODE As String) As String
    '    STYLE_CODE = STYLE_CODE.ToUpper
    '    Dim QUOTE_NO As String = Absx1.txtFor("QUOTE_NO").Text
    '    Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE})
    '    If rowICTQUOT2 Is Nothing Then
    '        Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE)
    '        If rowICTPLIN2 Is Nothing Then Return ""
    '        rowICTQUOT2 = dst.Tables("ICTQUOT2").NewRow()

    '        With rowICTQUOT2
    '            .Item("QUOTE_NO") = QUOTE_NO
    '            .Item("STYLE_CODE_PLM") = STYLE_CODE
    '            .Item("STYLE_CODE_CUST") = STYLE_CODE
    '            .Item("STYLE_DESC") = rowICTPLIN2.Item("STYLE_DESC")
    '            Dim SEQ As Integer = Val(dst.Tables("ICTQUOT2").Compute("MAX(SEQ)", "") & "") + 10
    '            .Item("SEQ") = SEQ

    '            ASCMAIN1.sql = "Select MAX(STYLE_PRICE) STYLE_PRICE, COUNT (*) COMPS " _
    '                  & " from ICTPLIN3 where STYLE_CODE_PLM = '" & STYLE_CODE & "'"
    '            Dim row As DataRow = ASCDATA1.GetDataRow
    '            If row IsNot Nothing AndAlso Val(row.Item("COMPS") & "") = 1 Then
    '                .Item("STYLE_PRICE") = row.Item("STYLE_PRICE")
    '            End If

    '            .Item("SIZE_SCALE") = Get_Colors(STYLE_CODE)
    '            '.Item("STYLE_DESC2") = rowICTPLIN2.Item("STYLE_DESC2")
    '            .Item("SALES_DIVISION_CODE") = rowICTPLIN2.Item("SALES_DIVISION_CODE")
    '            .Item("STYLE_CLASS_CODE") = rowICTPLIN2.Item("STYLE_CLASS_CODE")
    '            .Item("STYLE_GROUP_CODE") = rowICTPLIN2.Item("STYLE_GROUP_CODE")
    '            .Item("SEASON_CODE") = rowICTPLIN2.Item("SEASON_CODE")
    '        End With
    '        dst.Tables("ICTQUOT2").Rows.Add(rowICTQUOT2)
    '        ' FetchImage(rowICTQUOT2)
    '        Load_Pricing(rowICTQUOT2)
    '    End If
    '    txtQS_STYLE_CODE.Text = ""
    '    Return STYLE_CODE_PLM
    'End Function

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
            Dim ATTACHMENT_FILENAME As String = rowASTATTA2.Item("ATTACHMENT_FILENAME").ToString
            Dim ext As String = getFileExt(ATTACHMENT_FILENAME)

            MaxFILE_NO += 1
            Dim newICTQUOHF As DataRow = dst.Tables.Item("ICTQUOHF").NewRow
            newICTQUOHF.Item("SESSION_NO") = SESSION_NO
            newICTQUOHF.Item("FILE_NO") = MaxFILE_NO
            newICTQUOHF.Item("FILENAME") = ATTACHMENT_FILENAME
            newICTQUOHF.Item("HASHVALUE") = rowASTATTA2.Item("HASHVALUE").ToString
            dst.Tables.Item("ICTQUOHF").Rows.Add(newICTQUOHF)

            Dim FileNameLocalFull As String = rowASTATTA2.Item("ATTACHMENT_FILENAME").ToString
            Dim FileNameRemote As String = $"{SESSION_NO}-{MaxFILE_NO}{ext}"
            Dim eMsg As Text.StringBuilder = FTP_BLUEHOST(FileNameLocalFull, FileNameRemote)
            If eMsg.Length > 0 Then
                MsgBox(eMsg.ToString, vbCritical, "Error Sending To Remote Server")
            End If
        Next
    End Sub

    Private Function FTP_BLUEHOST(ByRef FileNameLocalFull As String, ByRef FileNameRemote As String) As Text.StringBuilder
        Dim RetVal As New Text.StringBuilder With {.Length = 0}
        Dim FTPUser As String = "abs@vandalequotes.com"
        Dim FTPPassword As String = "0ff1c3ABS#"
        Dim FTPHost As String = "ftp.tzn.lnr.mybluehost.me"
        Dim FTPRemoteFull As String = $"/public_html/FTP/{FileNameRemote}"

        If Not System.IO.File.Exists(FileNameLocalFull) Then
            RetVal.AppendLine($"FTP File Provided Does Not Exist: {FileNameLocalFull}")
        End If

        If RetVal.Length = 0 Then
            Try
                If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    Stop
                End If

                Ftp1.User = FTPUser
                Ftp1.Password = FTPPassword
                Ftp1.RemoteHost = FTPHost

                Ftp1.Logon()

                Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                Ftp1.LocalFile = FileNameLocalFull
                Ftp1.RemoteFile = FTPRemoteFull
                'Ftp1.Timeout = 0 'Don't Timeout
                Ftp1.Overwrite = True

                Ftp1.Upload()

                Ftp1.Logoff()
            Catch ex As Exception
                RetVal.AppendLine($"FTP Error: {ex.Message} : {ex.InnerException}")
                'Just bail out for now.  We eventually need some kind of tracking.
            End Try
        End If
        Return RetVal
    End Function

    Private Sub CALC_EXCLUDED_COLORS()
        For Each row As DataRow In dst.Tables("ICTQUOT3").Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = row.Item("COLOR_CODE").ToString & String.Empty
            Dim EXCLUDE_COLOR = "0"
            If chkExcludeWIP.Checked Then
                If Val(row.Item("WIP_TRAN").ToString & String.Empty) > 0 Then
                    EXCLUDE_COLOR = "1"
                End If
            End If
            If chkGreaterThan.Checked Then
                Dim valGreaterThan As Integer = numGreaterThan.Value
                Dim actGreaterThan As Integer = 0
                If chkAGED_01.Checked Then
                    actGreaterThan = actGreaterThan + Val(row.Item("QTY_AVA1").ToString & String.Empty)
                End If
                If chkAGED_02.Checked Then
                    actGreaterThan = actGreaterThan + Val(row.Item("QTY_AVA2").ToString & String.Empty)
                End If
                If chkAGED_03.Checked Then
                    actGreaterThan = actGreaterThan + Val(row.Item("QTY_AVA3").ToString & String.Empty)
                End If
                If actGreaterThan <= valGreaterThan Then
                    EXCLUDE_COLOR = "1"
                End If
            End If
            row.Item("EXCLUDE_COLOR") = EXCLUDE_COLOR
        Next
        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE_PLM").ToString & String.Empty
            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND EXCLUDE_COLOR = '0'", STYLE_CODE)
            Dim recCnt As Integer = dst.Tables("ICTQUOT3").Select(filter).Count
            If recCnt = 0 Then
                row.Item("EXCLUDE_STYLE") = "1"
            Else
                row.Item("EXCLUDE_STYLE") = "0"
            End If
        Next
    End Sub

    Private Sub FILL_WIPTRANS()
        For Each row As DataRow In dst.Tables("ICTQUOT3").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = row.Item("COLOR_CODE").ToString & String.Empty
            ASCMAIN1.Progress("Calculating WIP For", String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE))
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT SUM(NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0)) AS WIP_TRAN")
            SQLS.AppendLine("FROM ICTSTAT2")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQLS.ToString()
            Dim WIP_TRAN As Int16 = Val(ASCDATA1.GetDataValue)
            row.Item("WIP_TRAN") = WIP_TRAN
        Next
        ASCMAIN1.Progress("", "")
    End Sub

    Private Function getCostForStyleColor(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As Double
        Dim Retval As Double = 0
        ASCMAIN1.sql = "Select STYLE_COST from (" & vbCrLf _
                            & "Select STYLE_COST from ICTCOSTA " & vbCrLf _
                            & "where (STYLE_CODE, COLOR_CODE) in (" & vbCrLf _
                            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                            & " from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'" _
                            & " and WHSE_QTY_ON_HAND > 0)" & vbCrLf _
                            & " order by OPS_YYYYPP DESC) where ROWNUM < 2"
        Dim STYLE_COST As Decimal = Val(ASCDATA1.GetDataValue)

        If STYLE_COST = 0 Then
            ASCMAIN1.sql = "Select NVL(PO_COST_LANDED,PO_COST) STYLE_COST" & vbCrLf _
                                & " from (" & vbCrLf _
                                & " Select POTSHIP3.PO_SHIPMENT_NO, POTORDR2.PO_ORDER_NO, " & vbCrLf _
                                & " POTORDR2.PO_COST, POTSHIP3.PO_COST_LANDED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
                                & " from POTORDR2,POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
                                & " where POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_LNO (+) = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & " order by POTSHIP3.PO_SHIPMENT_NO DESC, POTORDR2.PO_ORDER_NO DESC" & vbCrLf _
                                & ") where ROWNUM <2"
            STYLE_COST = Val(ASCDATA1.GetDataValue)
        End If

        If STYLE_COST <> 0 Then
            Retval = Math.Round(STYLE_COST, 2)
        End If

        Return Retval
    End Function
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

    Private Sub grdICTSTDQ1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
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

    Function Add_to_Quote(STYLE_CODE_PLM As String, Optional ByRef Style_cnt As Int64 = 0, Optional ByRef color_cnt As Int64 = 0) As String

        If QUOTE_NO = "" Then
            Click_Command("New Quote Sheet")
        End If
        STYLE_CODE_PLM = STYLE_CODE_PLM.ToUpper
        Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE_PLM})
        If rowICTQUOT2 Is Nothing Then
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_PLM)
            If rowICTSTYL1 Is Nothing Then Return ""

            'Find Qualified Colors that Match Criteria and Add Them To ICTQUOT3
            Dim HasQualifiedColor As Boolean = False
            Dim numGreaterThan As Int64 = Val(numGreaterThan.ToString & String.Empty)
            Dim filter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE_PLM)
            For Each rowICTQUOTQ As DataRow In dst.Tables("ICTQUOTQ").Select(filter, "STYLE_CODE, COLOR_CODE")
                Dim ThisColorQualifies As Boolean = False
                Dim AGED_01 As Int64 = Val(rowICTQUOTQ.Item("AGED_01") & String.Empty)
                If chkAGED_01.Checked And AGED_01 > numGreaterThan Then
                    ThisColorQualifies = True
                End If

                Dim AGED_02 As Int64 = Val(rowICTQUOTQ.Item("AGED_02") & String.Empty)
                If chkAGED_02.Checked And AGED_02 > numGreaterThan Then
                    ThisColorQualifies = True
                End If

                Dim AGED_03 As Int64 = Val(rowICTQUOTQ.Item("AGED_03") & String.Empty)
                If chkAGED_03.Checked And AGED_03 > numGreaterThan Then
                    ThisColorQualifies = True
                End If
                If ThisColorQualifies Then
                    HasQualifiedColor = True
                    Dim rowICTQUOT3 As DataRow = dst.Tables("ICTQUOT3").NewRow()
                    Dim COLOR_CODE As String = rowICTQUOTQ.Item("COLOR_CODE").ToString & String.Empty
                    Dim COLOR_DESC As String = dst.Tables("ICTCOLR1").Select(String.Format("COLOR_CODE = '{0}'", COLOR_CODE)).FirstOrDefault.Item("COLOR_DESC").ToString & String.Empty
                    rowICTQUOT3.Item("QUOTE_NO") = QUOTE_NO
                    rowICTQUOT3.Item("STYLE_CODE") = STYLE_CODE_PLM
                    rowICTQUOT3.Item("COLOR_CODE") = COLOR_CODE
                    rowICTQUOT3.Item("STYLE_COLOR_DESC") = COLOR_DESC
                    rowICTQUOT3.Item("QTY_AVA1") = AGED_01
                    rowICTQUOT3.Item("QTY_AVA2") = AGED_02
                    rowICTQUOT3.Item("QTY_AVA3") = AGED_03
                    dst.Tables("ICTQUOT3").Rows.Add(rowICTQUOT3)
                    color_cnt += 1
                End If
            Next

            'If There are any Qulified Colors Then Add The Style.
            If HasQualifiedColor Then
                Dim newICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").NewRow()
                Dim SEQ As Integer = Val(dst.Tables("ICTQUOT2").Compute("MAX(SEQ)", "") & "") + 1

                newICTQUOT2.Item("QUOTE_NO") = QUOTE_NO
                newICTQUOT2.Item("STYLE_CODE_PLM") = STYLE_CODE_PLM
                newICTQUOT2.Item("STYLE_CODE_CUST") = STYLE_CODE_PLM
                newICTQUOT2.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                newICTQUOT2.Item("SEQ") = SEQ
                newICTQUOT2.Item("STYLE_PRICE") = rowICTSTYL1.Item("STYLE_PRICE")
                'newICTQUOT2.Item("SIZE_SCALE") = rowICTSTYL1.Item("SIZE_SCALE")
                If ASCMAIN1.CLIENT = "VAN" Then
                    Dim SQCs As String = TAC.ICCMAIN1.Get_SIZEs_and_QTYs_and_COLORs(Me, STYLE_CODE_PLM)
                    newICTQUOT2.Item("SIZE_SCALE") = SQCs
                Else
                    newICTQUOT2.Item("SIZE_SCALE") = rowICTSTYL1.Item("SIZE_SCALE")
                End If
                newICTQUOT2.Item("STYLE_DESC2") = rowICTSTYL1.Item("STYLE_DESC2")
                newICTQUOT2.Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                newICTQUOT2.Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                newICTQUOT2.Item("STYLE_GROUP_CODE") = rowICTSTYL1.Item("SUB_BODY_CODE")
                newICTQUOT2.Item("SUB_BODY_CODE") = rowICTSTYL1.Item("SUB_BODY_CODE")
                newICTQUOT2.Item("FABRIC_CODE") = rowICTSTYL1.Item("FABRIC_CODE")
                newICTQUOT2.Item("SEASON_CODE") = rowICTSTYL1.Item("SEASON_CODE")
                newICTQUOT2.Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
                newICTQUOT2.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
                newICTQUOT2.Item("IMAGE_NAME") = rowICTSTYL1.Item("IMAGE_NAME")
                dst.Tables("ICTQUOT2").Rows.Add(newICTQUOT2)
                Style_cnt += 1
            End If

        End If

        Return STYLE_CODE_PLM
    End Function

    Private Sub cmdAddMultipleStyles_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddMultipleStyles.Click
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.Custom_sql_where = " And STYLE_CODE In (Select Distinct STYLE_CODE from ICTSTAT2 where NVL(WHSE_QTY_ON_HAND,0) <> 0 Or NVL(WHSE_QTY_ON_ORDER,0) <> 0 Or NVL(WHSE_QTY_TRAN,0) <> 0)"
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
            Dim dvw As DataView = DirectCast(grdICTQUOT2B.DataSource, DataTable).DefaultView
            If chkShowSelectedOnly.Checked Then
                dvw.RowFilter = "SELECTED = '1'"
            Else
                dvw.RowFilter = ""
            End If
        End If
    End Sub

    Private Function Create_Excel(ByVal SALES_DIVISION_CODE As String, ByVal MakeBuyerVersion As Boolean) As String
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

        'Dim sql0 As String = " and COUNT_COLOR > 0" ' & Val(numMinQty.Value & "")
        Dim sql0 As String = "" ' & Val(numMinQty.Value & "")
        'If chkShowSelectedOnly.Checked Then
        '    sql0 &= " and SELECTED = '1'"
        'End If
        sql0 = " and SELECTED = '1'"

        CUSTPOSs.Clear()

        Dim CUSTPOi As Integer = 0
        dst.Tables("SOTORDRC").Rows.Clear()

        For Each row As DataRow In dst.Tables("ICTSTYC1").Select("")
            row.Item("OPEN_PICK_RSRV") = 0
        Next

        If chkShowPOs.Checked Then
            For Each row As DataRow In dst.Tables("ICTQUOT2").Select("")
                Fill_Records("SOTORDRC", New String() {txtQuoteCUST_CODE.Text, row.Item("STYLE_CODE_PLM").ToString & String.Empty}, False)
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
            Dim filter As String = Mid(sqlWB & sql0, 6) & " AND EXCLUDE_STYLE = '0'"
            For Each rowSB As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ICTQUOT2").Select(filter), Split(CODES, ",")).Select()
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

                    'Create_Excel_WorkSheet(worksheet, sqlWB & sqlSB & sql0)
                    Create_Excel_WorkSheet(worksheet, MakeBuyerVersion, sqlSB, sql0)
                    XLS_CREATED = True
                End If
            Next
        Else
            If dst.Tables("ICTQUOT2").Select(Mid(sqlWB & sql0, 6)).Length > 0 Then
                Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
                worksheet.Name = "Style Info"
                'Create_Excel_WorkSheet(worksheet, sqlWB & sql0)
                Create_Excel_WorkSheet(worksheet, MakeBuyerVersion, "", sql0)
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
                    If MakeBuyerVersion Then
                        XLS_FILENAME &= "-" & Format(XLS_NO, "000") & "_B.XLSX"
                    Else
                        XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".XLSX"
                    End If

                    If IO.File.Exists(ASCMAIN1.Folders("Temp") & XLS_FILENAME) Then
                        success = False
                    Else
                        workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    End If
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

    Sub Create_Excel_WorkSheet(
                              worksheet As SpreadsheetGear.IWorksheet,
                               ByVal MakeBuyerVersion As Boolean,
                              ByVal sqlSB As String,
                              ByVal sql0 As String)
        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        If (ASCMAIN1.Running_in_VS) Then
            If Not IO.Directory.Exists(IMAGE_FOLDER) Then
                Stop 'You Need to Set up Image Folder.
            End If
        End If
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        worksheet.Cells("A1:Z1").EntireColumn.Font.Size = 16

        Dim CX As Integer = 0
        Dim RX As Integer = 0

        Dim I As Integer = 0
        I += 4

        Dim CWC() As String = Split("A,B, C,D,E,F,G,H,I,J,K,L, M", ",")
        Dim CWS() As String = Split("1,1,40,6,6,6,6,6,6,6,6,6,20", ",")
        'If optPP.Value & "" = "4/5" Then
        '    CWS(2) = 45
        'End If
        CWS(2) = 45
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
        Next

        worksheet.Cells(0, 0).EntireColumn.Hidden = True
        worksheet.Cells(0, 1).EntireColumn.Hidden = True

        Dim COL0 As Integer = 6 + 6

        Dim COL As Integer = COL0

        Dim ColVisible(4) As Boolean
        ColVisible(0) = True
        If MakeBuyerVersion Then
            ColVisible(1) = False
            ColVisible(2) = False
            ColVisible(3) = False
        Else
            ColVisible(1) = chkAGED_01.Checked
            ColVisible(2) = chkAGED_02.Checked
            ColVisible(3) = chkAGED_03.Checked
        End If
        ColVisible(4) = False

        For iCol As Integer = 1 To 4
            If ColVisible(iCol) Then
                COL += 1
                With worksheet.Cells(I - 1, COL)
                    .ColumnWidth = 15
                    .EntireColumn.NumberFormat = "#,##0"
                    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With
                If chkExtended.Checked Then
                    COL += 1
                    With worksheet.Cells(I - 1, COL)
                        .ColumnWidth = 15
                        .EntireColumn.NumberFormat = "#,##0"
                        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    End With
                End If
            End If
        Next

        COL += 1
        With worksheet.Cells(I - 1, COL)
            .ColumnWidth = 15
            .EntireColumn.NumberFormat = "#,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            '.Value = "All"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With

        If chkExtended.Checked Then
            COL += 1
            With worksheet.Cells(I - 1, COL)
                .ColumnWidth = 15
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                '.Value = "All"
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
        End If

        With worksheet.Cells(I, 0, I, COL)
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        Dim I0 As Integer = 0
        Dim IA As Integer = 0
        Dim RT(5) As String
        Dim ROW0 As Integer = I
        Dim style_count As Integer = 0
        Dim pages As Integer = 0

        Dim fltrICTQUOT2 As String = ""
        If (sqlSB & sql0).Length > 5 Then
            fltrICTQUOT2 = (sqlSB & sql0).Substring(5) & " AND EXCLUDE_STYLE = '0'"
        Else
            fltrICTQUOT2 = "EXCLUDE_STYLE = '0'"
        End If

        Dim GrandTotal As Int64 = 0
        For Each selICTQUOT2 As DataRow In dst.Tables.Item("ICTQUOT2").Select(fltrICTQUOT2)
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", selICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)
            ASCMAIN1.Progress("-", selICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)

            I += 1
            I0 = I

            COL = COL0

            worksheet.Cells(I, COL - 1).Value = "Color"
            worksheet.Cells(I, COL - 0).Value = "Description"

            For iCol As Integer = 1 To 4
                If ColVisible(iCol) Then
                    COL += 1
                    With worksheet.Cells(I, COL)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        If iCol = 0 Then
                            .Value = "COL 0"
                        End If
                        If iCol = 1 Then .Value = String.Format("0 - {0}", dayBreaks)
                        If iCol = 2 Then .Value = String.Format("{0} - {1}", dayBreaks + 1, dayBreaks * 2)
                        If iCol = 3 Then .Value = String.Format("Greater {0}", dayBreaks * 2)
                        If iCol = 4 Then .Value = "Never Used"
                    End With
                    If chkExtended.Checked Then
                        COL += 1
                        With worksheet.Cells(I, COL)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                            .Value = "Ext Cost"
                        End With
                    End If
                End If
            Next

            COL += 1
            With worksheet.Cells(I, COL)
                If MakeBuyerVersion Then
                    .Value = "Now"
                Else
                    .Value = "Total"
                End If
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            If chkExtended.Checked And Not MakeBuyerVersion Then
                COL += 1
                With worksheet.Cells(I, COL)
                    .Value = "Ext Total"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With
            End If

            If chkShowLastRcd.Checked And Not MakeBuyerVersion Then
                COL += 1
                With worksheet.Cells(I, COL)
                    .Value = "Last Recd"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .ColumnWidth = 18
                End With
            End If

            range = worksheet.Cells(I, COL0 - 1, I, COL)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.Gold

            I += 1

            Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""

            Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME
            If Not IO.File.Exists(imageFileStyle) Then
                IMAGE_NAME = ""
            End If

            Dim ImageRows As Integer = 0
            Dim ImageRowsBig As Integer = 0

            If IMAGE_NAME <> "" _
                AndAlso My.Computer.FileSystem.FileExists(imageFileStyle) Then

                Dim widthStyle As Double
                Dim heightStyle As Double

                Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
                Try
                    widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution / 3
                    heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution / 3
                Finally
                    imageStyle.Dispose()
                End Try

                Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

                Dim col_adj As Decimal = 0
                If heightStyle > widthStyle Then
                    col_adj = 0.3
                Else
                    col_adj = 0.05
                End If

                Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0) + col_adj
                Dim topStyle As Double = windowInfoStyle.RowToPoints(I - 1) + 0.1 ' 1.5)

                ImageRows = windowInfoStyle.PointsToRow(heightStyle)
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
            End If

            CX = 1

            With worksheet.Cells(I - 1, 3)
                .Value = "'" & selICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty
                .Font.Color = SpreadsheetGear.Colors.Purple
                .Font.Size = 24
                .Font.Bold = True
            End With

            CX = 3

            worksheet.Cells(I + 2, CX).Value = "Case Qty"

            range = worksheet.Cells(I + 1, 3, I + 2, 4)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.LightGray

            range = worksheet.Cells(I + 1, 3 + 4, I + 2, 4 + 4)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.LightGray

            CX = 5
            worksheet.Cells(I, CX - 2).Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
            worksheet.Cells(I + 2, CX).Value = rowICTSTYL1.Item("CARTON_PACK_QTY")

            Dim SZMAX As Integer = 0
            Dim SZTOT As Integer = 0

            Dim rowICTSTYLX As DataRow = dst.Tables("ICTSTYLX").Rows.Find(selICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)

            Dim OFFSET As Integer = -2

            For SZ As Integer = 1 To 8
                worksheet.Cells(I + 4, CX + OFFSET + SZ - 1).Value = rowICTSTYLX.Item("S" & CStr(SZ)) & String.Empty
                worksheet.Cells(I + 5, CX + OFFSET + SZ - 1).Value = rowICTSTYLX.Item("Q" & CStr(SZ)) & String.Empty
                If rowICTSTYLX.Item("S" & CStr(SZ)) & "" <> "" Then
                    SZMAX = SZ
                    SZTOT += Val(rowICTSTYLX.Item("Q" & CStr(SZ)) & "")
                End If
            Next

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
                worksheet.Cells(I + 1, CX).Value = SZTOT
                worksheet.Cells(I + 1, 3).Value = "Inner"


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
            Dim styleTotal As Int64 = 0
            'Dim filter As String = "STYLE_CODE = '" & selICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty & "'"
            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND EXCLUDE_COLOR = '0'", selICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty)
            For Each rowICTQUOT3 As DataRow In dst.Tables("ICTQUOT3").Select(filter, "COLOR_CODE")
                Dim STYLE_COLOR_COST As Double = getCostForStyleColor(rowICTQUOT3.Item("STYLE_CODE") & String.Empty, rowICTQUOT3.Item("COLOR_CODE") & String.Empty)
                CI += 1
                COL = COL0

                If chkShowCost.Checked And Not MakeBuyerVersion Then
                    worksheet.Cells(I + CI - 1, COL - 2).Value = STYLE_COLOR_COST
                    worksheet.Cells(I + CI - 1, COL - 2).NumberFormat = "###,##0.00"
                    worksheet.Cells(I + CI - 1, COL - 2).Font.Bold = True
                    worksheet.Cells(I + CI - 1, COL - 2).Font.Color = SpreadsheetGear.Colors.Red
                End If

                worksheet.Cells(I + CI - 1, COL - 1).Value = "'" & rowICTQUOT3.Item("COLOR_CODE")
                'worksheet.Cells(I + CI - 1, COL - 0).Value = rowICTQUOT3.Item("STYLE_COLOR_DESC")
                worksheet.Cells(I + CI - 1, COL - 0).Value = GetAltColorCode(selICTQUOT2.Item("STYLE_CODE_PLM").ToString & String.Empty, rowICTQUOT3.Item("COLOR_CODE").ToString & String.Empty, rowICTQUOT3.Item("STYLE_COLOR_DESC").ToString & String.Empty)

                T = ""
                Dim VisCount As Integer = 0
                Dim rowTOTAL As Int64 = 0
                For iCOL As Integer = 1 To 4
                    If MakeBuyerVersion Then
                        If iCOL = 1 And chkAGED_01.Checked Then
                            rowTOTAL += Val(rowICTQUOT3.Item("QTY_AVA1").ToString & String.Empty)
                        End If
                        If iCOL = 2 And chkAGED_02.Checked Then
                            rowTOTAL += Val(rowICTQUOT3.Item("QTY_AVA2").ToString & String.Empty)
                        End If
                        If iCOL = 3 And chkAGED_03.Checked Then
                            rowTOTAL += Val(rowICTQUOT3.Item("QTY_AVA3").ToString & String.Empty)
                        End If
                    End If
                    If ColVisible(iCOL) Then
                        VisCount += 1
                        worksheet.Cells(I + CI - 1, COL + VisCount).Value = Val(rowICTQUOT3.Item("QTY_AVA" & iCOL).ToString & String.Empty)
                        T &= "+" & Replace(worksheet.Cells(I + CI - 1, COL + VisCount).Address, "$", "")
                        If chkExtended.Checked Then
                            VisCount += 1
                            worksheet.Cells(I + CI - 1, COL + VisCount).Value = Val(rowICTQUOT3.Item("QTY_AVA" & iCOL).ToString & String.Empty) * STYLE_COLOR_COST
                        End If
                    End If
                Next
                styleTotal += rowTOTAL
                GrandTotal += styleTotal
                COL += 1
                'This is where you can figure out the row total
                If MakeBuyerVersion Then
                    worksheet.Cells(I + CI - 1, COL + VisCount).Value = rowTOTAL
                Else
                    worksheet.Cells(I + CI - 1, COL + VisCount).Formula = "=" & Mid(T, 2)
                    If chkExtended.Checked Then
                        COL += 1
                        worksheet.Cells(I + CI - 1, COL + VisCount).Value = worksheet.Cells(I + CI - 1, COL - 1 + VisCount).Value * STYLE_COLOR_COST
                    End If
                End If

                If chkShowLastRcd.Checked And Not MakeBuyerVersion Then
                    COL += 1
                    With worksheet.Cells(I + CI - 1, COL + VisCount)
                        Dim fltrICTSTYC1 As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'",
                                                                   rowICTQUOT3.Item("STYLE_CODE").ToString & String.Empty,
                                                                   rowICTQUOT3.Item("COLOR_CODE").ToString & String.Empty)
                        Dim rowICTSTYC1 As DataRow = dst.Tables.Item("ICTSTYC1").Select(fltrICTSTYC1).FirstOrDefault
                        Dim LAST_RCD_DATE As String = ""
                        If Not IsNothing(rowICTSTYC1) Then
                            LAST_RCD_DATE = rowICTSTYC1.Item("LAST_RCD_DATE").ToString & String.Empty
                        End If
                        .Value = LAST_RCD_DATE
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                    End With
                End If
            Next

            CI += 1
            COL = COL0

            worksheet.Cells(I - 1, COL - 1, I + CI - 1, COL - 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center

            worksheet.Cells(I + CI - 1, COL - 1).Value = "'" & "***"
            worksheet.Cells(I + CI - 1, COL - 0).Value = "'" & "Total"

            T = ""
            Dim ET As Double = 0
            For iCOL As Integer = 1 To 4
                If ColVisible(iCOL) Then
                    COL += 1
                    If CI = 1 Then ' NO COLORS
                        worksheet.Cells(I + CI - 1, COL).Value = 0
                    Else
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                    End If

                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")

                    T &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                    'COL += 1
                    If chkExtended.Checked Then
                        COL += 1
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        ET += worksheet.Cells(I + CI - 1, COL).Value
                    End If
                End If
            Next
            COL += 1

            If MakeBuyerVersion Then
                worksheet.Cells(I + CI - 1, COL).Value = styleTotal
            Else
                worksheet.Cells(I + CI - 1, COL).Formula = "=" & Mid(T, 2)
                If chkExtended.Checked Then
                    COL += 1
                    'worksheet.Cells(I + CI - 1, COL).Value = 100
                    worksheet.Cells(I + CI - 1, COL).Value = ET
                End If
            End If

            RT(5) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")

            worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

            With worksheet.Cells(I, COL0 - 1, I + CI - 1, COL)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
            End With

            I += ImageRowsBig

            Dim CJ As Integer = ImageRows ' - 1

            If CJ < 6 Then CJ = 6

            If CI > CJ Then
                I += CI
            Else
                I += CJ
            End If

            style_count += 1

            If (((I - 5) Mod 80) < ((I0 - 5) Mod 80)) Or (style_count >= 5) Or style_count >= 9 Then
                Dim R As SpreadsheetGear.IRange = worksheet.Cells(I0, 0).EntireRow
                worksheet.HPageBreaks.Add(R)
                style_count = 1
                pages += 1
            End If

            With worksheet.Cells(I0, 0, I + 1 - 1, COL)
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
        Next

        I += 2
        COL = COL0

        'worksheet.Cells(I - 1, COL - 1).Value = "'" & "All"
        worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = 1 To 4
            If ColVisible(iCOL) Then
                COL += 1
                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)

                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                'COL += 1
                If chkExtended.Checked Then
                    COL += 1
                    worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                End If
            End If
        Next
        COL += 1
        If MakeBuyerVersion Then
            worksheet.Cells(I - 1, COL).Value = GrandTotal
        Else
            worksheet.Cells(I - 1, COL).Formula = "=" & Mid(GT, 2)
        End If

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

        Dim H1 As Integer = 11

        worksheet.Cells(0, 2).Value = "Quick Quote Sheet"
        worksheet.Cells(0, 2).Font.Bold = True

        worksheet.Cells(0, H1).Value = "Note"
        worksheet.Cells(1, H1).Value = "For"

        worksheet.Cells(0, H1, 2, H1).Interior.Color = SpreadsheetGear.Colors.LightGray

        'worksheet.Cells(0, H1 + 1).Value = Format(Absx1.dteFor("QUOTE_DATE").Value, "MM/dd/yyyy")
        worksheet.Cells(0, H1 + 1).NumberFormat = "MM/dd/yy"
        worksheet.Cells(0, H1 + 1).Value = "" 'Add txtNotes.Text here later
        worksheet.Cells(1, H1 + 1).Value = txtQuoteCUST_CODE.Text
        'worksheet.Cells(1, H1 + 2).Value = Absx1.txtFor("CUST_CODE").Text


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
            '.Value = Absx1.txtFor("QUOTE_DESC").Text
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
    End Sub

    Private Function GetAltColorCode(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal COLOR_DESC_ORIG As String) As String
        Dim RetVal As String = COLOR_DESC_ORIG
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        Dim SIZE_SCALE As String = rowICTSTYL1.Item("SIZE_SCALE") & String.Empty
        Dim MAX_LENGTH As Integer = 60
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
                SC = Mid(S, 1, J)
                SIZE_SCALE = Mid(SIZE_SCALE, 1, I - 1) & Mid(S, J)
                For C As Integer = 1 To SC.Length - 1
                    If C = 1 Or (C > 1 AndAlso Mid(SC, C + 1, 1) <> " " AndAlso (Mid(SC, C - 1, 1) = " " Or Mid(SC, C - 1, 1) = "/")) Then
                        Mid(SC, C, 1) = Mid(SC, C, 1).ToUpper
                    End If
                Next
                If Trim(SC) <> "" Then
                    If SC.Length > 35 Then
                        RetVal = SC.Substring(0, 34)
                    Else
                        RetVal = SC
                    End If

                End If
            End If
        End If
        If RetVal = COLOR_DESC_ORIG Then
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT NVL(STYLE_COLOR_DESC,'') STYLE_COLOR_DESC")
            SQLS.AppendLine("FROM ICTSTYC1")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQLS.ToString()
            Dim COLOR_DESC_MF As String = ASCDATA1.GetDataValue
            If COLOR_DESC_MF.Length > 35 Then
                COLOR_DESC_MF = COLOR_DESC_MF.Substring(0, 35)
            End If
            If COLOR_DESC_MF.Length > 0 Then
                RetVal = COLOR_DESC_MF
            End If
        End If
        Return RetVal
    End Function


    Private Sub btnFixColors_Click(sender As System.Object, e As System.EventArgs)

        Fill_Records("ICTSTYLX")
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
                    Dim SZ As String = Trim(Mid(SIZE_SCALE, I + 3))
                    Dim J As Integer = InStr(Mid(SZ & "  ", 1, MAX_LENGTH), "  ")
                    Dim K As Integer = InStr(Mid(SZ & vbCrLf, 1, MAX_LENGTH), vbCrLf)
                    If J = 0 And K = 0 Then
                        J = InStr(Mid(SZ & " ", 1, MAX_LENGTH), " ")
                    End If
                    If J = 0 Or J > K Then J = K
                    Dim SC As String = ""
                    If J <> 0 Then
                        fixed = True
                        SC = Mid(SZ, 1, J)
                        SIZE_SCALE = Mid(SIZE_SCALE, 1, I - 1) & Mid(SZ, J)
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

    Sub Sort_by_Style()
        Dim SEQ As Integer = 0
        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("", "STYLE_CODE_PLM")
            SEQ += 10
            row.Item("SEQ") = SEQ
        Next
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
            S.Length = 0
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
            'If STYLE_CODE = "500469XIZ" Then Stop
            If IsDate(LAST_RCD_DATE) Then
                'If chkRECDATES.Checked Then
                '    If CDate(LAST_RCD_DATE) < dteRECDATEFR.DateTime Or CDate(LAST_RCD_DATE) > dteRECDATETO.DateTime Then
                '        LAST_RCD_DATE = ""
                '    End If
                'End If
            End If
            If IsDate(LAST_RCD_DATE) Then
                LAST_RCD_DATE = Format(CDate(LAST_RCD_DATE), "MM/dd/yy")
            Else
                S.Length = 0
                S.AppendLine("SELECT NVL(WHSE_QTY_TRAN,0) AS IN_TRAN")
                S.AppendLine("FROM ICTSTAT2")
                S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                ASCMAIN1.sql = S.ToString()
                Dim IN_TRAN As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                If IN_TRAN > 0 Then
                    LAST_RCD_DATE = "In-Tran"
                Else
                    S.Length = 0
                    S.AppendLine("SELECT NVL(WHSE_QTY_ON_ORDER,0) AS IN_WIP")
                    S.AppendLine("FROM ICTSTAT2")
                    S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
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

    Private Sub btnLoadStyles_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadStyles.Click
        If txtQUOTE_DESC.Text = "" Then
            MsgBox("Please enter a Description for the Quote Sheet", vbOKOnly, "Can Not Load")
            Exit Sub
        End If
        Absc1.Get_SQL("*") ' ,"ICTQUOTV")
        '----------------------
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTCOLR1.COLOR_DESC,")
        S.AppendLine("SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) AS NET_POS,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) AS IN_TRANS,")
        S.AppendLine("(SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) - SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0))) AS NOW_OH,")
        S.AppendLine("0 AS RECEIVED_01,")
        S.AppendLine("0 AS RECEIVED_02,")
        S.AppendLine("0 AS RECEIVED_03,")
        S.AppendLine("0 AS AGED_01,")
        S.AppendLine("0 AS AGED_02,")
        S.AppendLine("0 AS AGED_03")
        S.AppendLine("FROM ICTSTYL1, ICTSTAT2, ICTCOLR1")
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
        S.AppendLine("AND ICTSTAT2.COLOR_CODE = ICTCOLR1.COLOR_CODE")
        If Absx1.optFor("ASN_OPT").Value = "S" Then
            S.AppendLine("AND ICTSTYL1.CUST_CODE IS NULL")
        ElseIf Absx1.optFor("ASN_OPT").Value = "N" Then
            S.AppendLine("AND ICTSTYL1.CUST_CODE IS NOT NULL")
        End If
        S.AppendLine(Absc1.sql_WHERE)
        'S.AppendLine(Absc1.SQLA_filter)
        S.AppendLine("HAVING SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) > 0")
        S.AppendLine("GROUP BY")
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTCOLR1.COLOR_DESC")
        ASCMAIN1.sql = S.ToString()
        'Create_TDA(dst.Tables.Add, "ICTQUOTQ", "**", 0, False)
        'With dst.Tables("ICTQUOTQ").Columns
        '    .Add("LAST_RCD_DATE")
        'End With
        Dim SQL As String = S.ToString
        LOAD_TEMP = ASCMAIN1.Temp_Table(S.ToString)

        Fill_Records("ICTQUOTD", dayBreaks)

        SetICTQUOTQAging(SQL)

        Dim NEW_STYLES As Integer = 0
        Dim NEW_COLORS As Integer = 0
        dst.Tables.Item("ICTQUOT2").Clear()
        dst.Tables.Item("ICTQUOT3").Clear()
        S.Length = 0
        S.AppendLine(String.Format("   SELECT DISTINCT STYLE_CODE FROM {0}", LOAD_TEMP))
        ASCMAIN1.sql = S.ToString
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "STYLE_CODE")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE})
            If rowICTQUOT2 Is Nothing Then
                ASCMAIN1.Progress("Style", STYLE_CODE)
                Dim Style_cnt As Int64 = 0
                Dim Color_cnt As Int64 = 0
                Add_to_Quote(STYLE_CODE, Style_cnt, Color_cnt)
                NEW_STYLES += Style_cnt
                NEW_COLORS += Color_cnt
            End If
        Next

        'RemoveBadColors()
        Update_Record(True)

        MsgBox(CStr(NEW_STYLES) & " New Styles Added", MsgBoxStyle.OkOnly, "Verification")
        MsgBox(CStr(NEW_COLORS) & " New Colors Added", MsgBoxStyle.OkOnly, "Verification")
        btnLoadStyles.Enabled = False
        btnRefreshStyles.Enabled = True
    End Sub

    Private Sub SetICTQUOTQAging(ByVal SQL As String)
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTCOLR1.COLOR_DESC,")
        S.AppendLine("SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) AS NET_POS,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) AS IN_TRANS,")
        S.AppendLine("(SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) - SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0))) AS NOW_OH,")
        S.AppendLine("0 AS RECEIVED_01,")
        S.AppendLine("0 AS RECEIVED_02,")
        S.AppendLine("0 AS RECEIVED_03,")
        S.AppendLine("0 AS AGED_01,")
        S.AppendLine("0 AS AGED_02,")
        S.AppendLine("0 AS AGED_03")
        S.AppendLine("FROM ICTSTYL1, ICTSTAT2, ICTCOLR1")
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
        S.AppendLine("AND ICTSTAT2.COLOR_CODE = ICTCOLR1.COLOR_CODE")
        S.AppendLine(String.Format("AND ICTSTYL1.STYLE_CODE IN (SELECT STYLE_CODE_PLM FROM ICTQUOT2 WHERE QUOTE_NO = '{0}')", Absx1.txtFor("QUOTE_NO").Text.ToString))
        S.AppendLine("GROUP BY")
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTCOLR1.COLOR_DESC")
        If SQL.Length = 0 Then
            SQL = S.ToString
        End If
        Fill_Records("ICTQUOTQ", , , SQL)
        For Each rowICTQUOTQ As DataRow In dst.Tables("ICTQUOTQ").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowICTQUOTQ.Item("STYLE_CODE").ToString()
            Dim COLOR_CODE As String = rowICTQUOTQ.Item("COLOR_CODE").ToString()
            Dim NOW_OH As Int64 = Val(rowICTQUOTQ.Item("NOW_OH").ToString & String.Empty)
            If NOW_OH > 0 Then
                ASCMAIN1.Progress("Now Calculating Aging For Style", String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE))
                Dim rowFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                Dim rowICTQUOTD As DataRow = dst.Tables("ICTQUOTD").Select(rowFilter).FirstOrDefault
                Dim RECEIVED_TOTAL As Int64 = 0
                If Not IsNothing(rowICTQUOTD) Then
                    Dim RECEIVED_01 As Int64 = rowICTQUOTD.Item("RECEIVED_01")
                    rowICTQUOTQ.Item("RECEIVED_01") = RECEIVED_01
                    Dim RECEIVED_02 As Int64 = rowICTQUOTD.Item("RECEIVED_02")
                    rowICTQUOTQ.Item("RECEIVED_02") = RECEIVED_02
                    Dim RECEIVED_03 As Int64 = rowICTQUOTD.Item("RECEIVED_03")
                    rowICTQUOTQ.Item("RECEIVED_03") = RECEIVED_03
                    RECEIVED_TOTAL = RECEIVED_01 + RECEIVED_02 + RECEIVED_03
                    If NOW_OH <= RECEIVED_01 Then
                        rowICTQUOTQ.Item("AGED_01") = NOW_OH
                    Else
                        rowICTQUOTQ.Item("AGED_01") = RECEIVED_01
                        NOW_OH = NOW_OH - RECEIVED_01
                        If NOW_OH <= RECEIVED_02 Then
                            rowICTQUOTQ.Item("AGED_02") = NOW_OH
                        Else
                            rowICTQUOTQ.Item("AGED_02") = RECEIVED_02
                            NOW_OH = NOW_OH - RECEIVED_02
                            rowICTQUOTQ.Item("AGED_03") = NOW_OH
                        End If
                    End If
                End If
                If chkGreaterThan.Checked Then
                    If RECEIVED_TOTAL < numGreaterThan.Value Then
                        rowICTQUOTQ.Delete()
                    End If
                End If
            Else
                rowICTQUOTQ.Delete()
            End If
        Next
        dst.Tables("ICTQUOTQ").AcceptChanges()
    End Sub

    Private Sub tabStyles_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabStyles.SelectedTabChanged
        Setup_tabStyles()
    End Sub

    Sub Setup_tabStyles()
        If Me.SELECTION_NO = 0 Then Exit Sub
        If tabStyles.SelectedTab Is Nothing Then Exit Sub

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

    Private Sub chkShowLastRcd_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowLastRcd.CheckedChanged
        If Not Form_Loading Then
            If chkShowLastRcd.Checked Then
                setLastRcdDate()
            End If
        End If
    End Sub

    Private Sub chkUse180Days_CheckedChanged(sender As Object, e As EventArgs) Handles chkUse180Days.CheckedChanged
        If chkUse180Days.Checked Then
            chkCustomBreaks.Checked = False
        End If
        If chkCustomBreaks.Checked Then
            numCustomBreaks.ReadOnly = False
        Else
            numCustomBreaks.ReadOnly = True
        End If
        setDayBreaks()
    End Sub

    Private Sub chkCustomBreaks_CheckedChanged(sender As Object, e As EventArgs) Handles chkCustomBreaks.CheckedChanged
        If chkCustomBreaks.Checked Then
            numCustomBreaks.ReadOnly = False
        Else
            numCustomBreaks.ReadOnly = True
        End If
        chkUse180Days.Checked = False
        setDayBreaks()
    End Sub

    Private Sub numCustomBreaks_ValueChanged(sender As Object, e As EventArgs) Handles numCustomBreaks.ValueChanged
        setDayBreaks()
    End Sub

    Private Sub setDayBreaks()
        If chkCustomBreaks.Checked Then
            dayBreaks = numCustomBreaks.Value
        Else
            If chkUse180Days.Checked Then
                dayBreaks = 180
                numCustomBreaks.Value = 180
            Else
                dayBreaks = 120
                numCustomBreaks.Value = 120
            End If
        End If
        grpAGED_01.Text = String.Format("0 to {0} Days", dayBreaks)
        grpAGED_02.Text = String.Format("{0} to {1} Days", dayBreaks + 1, dayBreaks * 2)
        grpAGED_03.Text = String.Format("Greater that {0} Days", (dayBreaks * 2) + 1)
    End Sub

    Sub Setup_ICTQUOT3()
        If grdICTQUOT2B.ActiveRow Is Nothing OrElse (Not grdICTQUOT2B.ActiveRow.IsDataRow Or grdICTQUOT2B.ActiveRow.IsAddRow) Then
        Else
            Dim dvw As DataView = DirectCast(grdICTQUOT3.DataSource, DataTable).DefaultView
            Dim STYLE_CODE As String = grdICTQUOT2B.ActiveRow.Cells("STYLE_CODE_PLM").Value & String.Empty
            dvw.RowFilter = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            grdICTQUOT3.Text = "Color Details for Style " & STYLE_CODE
        End If
    End Sub

    Private Sub grdICTQUOT2B_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTQUOT2B.AfterRowActivate
        Setup_ICTQUOT3()
    End Sub

    Private Sub btnRefreshStyles_Click(sender As Object, e As EventArgs) Handles btnRefreshStyles.Click
        Fill_Records("ICTQUOTD", dayBreaks)

        SetICTQUOTQAging("")

        Me.Cursor = Cursors.WaitCursor
        Dim numGreaterThan As Int64 = Val(numGreaterThan.ToString & String.Empty)

        For Each rowICTQUOT3 As DataRow In dst.Tables("ICTQUOT3").Select()
            Dim THIS_STYLE As String = rowICTQUOT3.Item("STYLE_CODE").ToString & String.Empty
            Dim THIS_COLOR As String = rowICTQUOT3.Item("COLOR_CODE").ToString & String.Empty
            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", THIS_STYLE, THIS_COLOR)

            Dim AGED_01 As Int64 = 0
            Dim AGED_02 As Int64 = 0
            Dim AGED_03 As Int64 = 0
            Dim rowICTQUOTQ As DataRow = dst.Tables.Item("ICTQUOTQ").Select(filter).FirstOrDefault
            If Not IsNothing(rowICTQUOTQ) Then
                AGED_01 = Val(rowICTQUOTQ.Item("AGED_01") & String.Empty)
                AGED_02 = Val(rowICTQUOTQ.Item("AGED_02") & String.Empty)
                AGED_03 = Val(rowICTQUOTQ.Item("AGED_03") & String.Empty)
            End If

            rowICTQUOT3.Item("QTY_AVA1") = AGED_01
            rowICTQUOT3.Item("QTY_AVA2") = AGED_02
            rowICTQUOT3.Item("QTY_AVA3") = AGED_03

        Next
        Me.Cursor = Cursors.Default
        MsgBox("Styles Are Now Fresh!")
    End Sub

    Private Sub txtQS_STYLE_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtQS_STYLE_CODE.ValueChanged

    End Sub

    Private Function GET_ONLY_SIZE_SCALE(ByVal STYLE_CODE As String) As String
        Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)

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

    Private Function getFileExt(ByVal ATTACHMENT_FILENAME As String) As String
        Dim RetVal As String = ""
        Dim dotLoc As Int64 = ATTACHMENT_FILENAME.IndexOf(".", ATTACHMENT_FILENAME.Length - 5)
        If dotLoc > 0 Then
            RetVal = ATTACHMENT_FILENAME.Substring(dotLoc, ATTACHMENT_FILENAME.Length - dotLoc)
        End If
        Return RetVal
    End Function

    'Sub Print_Style_Sheet(eItemKey As String, Optional STYLE_CODE As String = "")
    '    Dim ListPDFSheets As New List(Of String)

    '    Dim BegAlloPeriod As Int64 = CalculateBegAlloPeriod()

    '    Synch_TABLE_NAME("ICTQUOT1")

    '    Dim blnShowSelected As Boolean = False
    '    If Not chkShowSelectedOnly.Checked Then
    '        blnShowSelected = True
    '        chkShowSelectedOnly.Checked = True
    '    End If

    '    RESEQ()

    '    'For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
    '    '    FetchImage(row)
    '    'Next

    '    ' COPYING THE SAME LOGIC USED FOR EXCEL
    '    For Each row As DataRow In dst.Tables("ICTSTYC1").Select("")
    '        row.Item("OPEN_PICK_RSRV") = 0
    '        row.Item("SKIP_COLOR") = "0"
    '    Next
    '    dst.Tables("ICTSTYC1").Columns("COUNT_COLOR").Expression = String.Format(COUNT_COLOR, Val(numMinQty.Value & ""))

    '    Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
    '    If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
    '    FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")

    '    For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
    '        'If row.Item("IMAGE") Is DBNull.Value Then
    '        '    row.Item("SELECTED") = "0"
    '        'End If
    '        Dim STYLE_CODE_PLM As String = row.Item("STYLE_CODE_PLM")
    '        'If STYLE_CODE_PLM = "500498AVR" And ASCMAIN1.Running_in_VS Then Stop

    '        If Not My.Computer.FileSystem.FileExists(FOLDER_NAME & row.Item("IMAGE_NAME")) Then
    '            row.Item("SELECTED") = "0"
    '        End If
    '        If Val(row.Item("COUNT_COLOR") & "") > 0 Then
    '        Else
    '            If Not chkShowZero.Checked Then
    '                row.Item("SELECTED") = "0"
    '            End If
    '        End If

    '        If BegAlloPeriod <> 0 Then
    '            Dim totalAvaliable As Int64 = 0
    '            For Each rowA As DataRow In dst.Tables.Item("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE_PLM))
    '                For tcnt As Integer = BegAlloPeriod To 4
    '                    totalAvaliable += Val(rowA.Item(String.Format("QTY_AVA{0}", tcnt.ToString)) & "")
    '                Next
    '            Next
    '            If totalAvaliable = 0 Then
    '                If Not chkShowZero.Checked Then
    '                    row.Item("SELECTED") = "0"
    '                End If
    '            End If
    '        End If

    '    Next

    '    Dim RPT As String = "ICRQUOT1"
    '    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
    '        RPT = "ICRQUOT2"
    '    End If

    '    Dim ColVisible(4) As Boolean
    '    ColVisible(0) = True
    '    ColVisible(1) = (tkb1.Value <= 2)
    '    ColVisible(2) = (tkb1.Value <= 1)
    '    ColVisible(3) = (tkb1.Value <= 0)
    '    ColVisible(4) = chkBeyond.Checked

    '    'For iCol As Integer = 0 To 4
    '    '    Dim sValue As String = ""
    '    '    If iCol = 0 Then sValue = "At Once"
    '    '    If iCol = 1 Then sValue = Format(dte1.Value, "MM/dd")
    '    '    If iCol = 2 Then sValue = Format(dte2.Value, "MM/dd")
    '    '    If iCol = 3 Then sValue = Format(dte3.Value, "MM/dd")
    '    '    If iCol = 4 Then sValue = "Beyond"
    '    '    If ColVisible(iCol) Then
    '    '        CR_params.Add("DTE" & CStr(iCol), sValue)
    '    '    Else
    '    '        CR_params.Add("DTE" & CStr(iCol), "")
    '    '    End If
    '    'Next

    '    'setLastShipDate()
    '    'setLastRcdDate()

    '    If chk1perPage.Checked Then ' Or STYLE_CODE <> "" Then
    '        'CR_params.Add("TXTSTYLE_CODE", "") '  STYLE_CODE)

    '        RPT = "ICRQUOTN"
    '        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
    '            If chkOmitAvail.Checked Then
    '                RPT = "ICRQUOTX"
    '            Else
    '                RPT = "ICRQUOTV"
    '            End If
    '        End If
    '    End If

    '    If eItemKey <> "email" Then

    '        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
    '            row.Item("SELECTED") = "2"

    '            ASCMAIN1.sql = " SALES_DIVISION_CODE = '" & row.Item("SALES_DIVISION_CODE") & "'"
    '            Dim rows() As DataRow = dst.Tables("SOTSDIVC").Select(ASCMAIN1.sql)
    '            If rows.Length <> 0 Then
    '                Dim rowSOTSDIVC As DataRow = dst.Tables("SOTSDIVC").Rows(0)
    '                row.Item("SALES_DIVISION_CODE_COMB") = rows(0).Item("SALES_DIVISION_CODE_COMB")
    '            Else
    '                row.Item("SALES_DIVISION_CODE_COMB") = row.Item("SALES_DIVISION_CODE")
    '            End If


    '        Next

    '        Dim REPORT_INDEX As Integer = 0
    '        Dim PDF_FN As String = ""
    '        Dim PDF_LINKS As String = ""
    '        Dim SUB_BODY_DESC As String = ""
    '        Dim SALES_DIVISION_NAME As String = ""
    '        Dim FABRIC_DESC As String = ""
    '        Dim DESCHASH As String = ""

    '        Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/"
    '        Dim SESSION_NO As String = ASCMAIN1.Next_Control_No("ICTQUOH1.SESSION_NO")
    '        Dim FILE_NO As Integer = 0


    '        If chkPublishPDF.Checked Then
    '            Dim rowICTQUOH1 As DataRow = dst.Tables("ICTQUOH1").NewRow
    '            rowICTQUOH1.Item("SESSION_NO") = SESSION_NO
    '            rowICTQUOH1.Item("QUOTE_NO") = QUOTE_NO
    '            rowICTQUOH1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
    '            rowICTQUOH1.Item("INIT_OPER") = ASCMAIN1.USER_ID
    '            dst.Tables("ICTQUOH1").Rows.Add(rowICTQUOH1)

    '        End If

    '        Do While dst.Tables("ICTQUOT2").Select("SELECTED='2'").Length <> 0

    '            Print_Report_Begin()

    '            If chkPublishPDF.Checked Then
    '                ' 1 report for every Comb-Division / Sub-Body / Fabric



    '                Dim row2() As DataRow = dst.Tables("ICTQUOT2").Select("SELECTED='2'", "SEQ")
    '                Dim SALES_DIVISION_CODE_COMB As String = row2(0).Item("SALES_DIVISION_CODE_COMB") & ""
    '                Dim SALES_DIVISION_CODE As String = row2(0).Item("SALES_DIVISION_CODE")
    '                Dim STYLE_GROUP_CODE As String = row2(0).Item("STYLE_GROUP_CODE")
    '                Dim FABRIC_CODE As String = row2(0).Item("FABRIC_CODE")
    '                Dim SUB_BODY_CODE As String = row2(0).Item("SUB_BODY_CODE")

    '                '   PDF_FN = SALES_DIVISION_CODE & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE
    '                PDF_FN = SALES_DIVISION_CODE_COMB & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE

    '                ASCMAIN1.sql = "Select SALES_DIVISION_NAME from SOTSDIV1 where SALES_DIVISION_CODE = :PARM1"
    '                Dim rowSOTDIV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SALES_DIVISION_CODE_COMB)
    '                If rowSOTDIV1 IsNot Nothing Then
    '                    SALES_DIVISION_NAME = rowSOTDIV1.Item("SALES_DIVISION_NAME")
    '                Else
    '                    SALES_DIVISION_NAME = ""
    '                End If

    '                '            If chkGROUP.Checked Then
    '                '                SALES_DIVISION_NAME = "GROUP" & STYLE_GROUP_CODE
    '                '               PDF_FN = SALES_DIVISION_NAME & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE
    '                '           End If

    '                '     PDF_FN = SALES_DIVISION_CODE_COMB & "-" & STYLE_GROUP_CODE & "-" & FABRIC_CODE

    '                ASCMAIN1.sql = "Select SUB_BODY_DESC from ICTBODY2 where SUB_BODY_CODE = :PARM1"


    '                Dim rowICTBODY2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", STYLE_GROUP_CODE)
    '                If rowICTBODY2 IsNot Nothing Then
    '                    SUB_BODY_DESC = rowICTBODY2.Item("SUB_BODY_DESC")
    '                Else
    '                    SUB_BODY_DESC = ""
    '                End If

    '                If chkGROUP.Checked Then
    '                    ASCMAIN1.sql = "Select SUB_BODY_DESC from ICTBODY2 where SUB_BODY_CODE = :PARM1"
    '                    Dim rowICTBODY2A As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SUB_BODY_CODE)
    '                    If rowICTBODY2A IsNot Nothing Then
    '                        SUB_BODY_DESC = rowICTBODY2A.Item("SUB_BODY_DESC")
    '                    Else
    '                        SUB_BODY_DESC = ""
    '                    End If
    '                    ' SUB_BODY_DESC = row2(0).Item("SUB_BODY_CODE")
    '                    SALES_DIVISION_NAME = "GROUP" & STYLE_GROUP_CODE
    '                    PDF_FN = SALES_DIVISION_NAME & "-" & SUB_BODY_CODE & "-" & FABRIC_CODE
    '                End If


    '                ASCMAIN1.sql = "Select FABRIC_DESC from ICTFABR1 where FABRIC_CODE = :PARM1"
    '                Dim rowICTFABR1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", FABRIC_CODE)
    '                If rowICTFABR1 IsNot Nothing Then
    '                    FABRIC_DESC = rowICTFABR1.Item("FABRIC_DESC")
    '                Else
    '                    FABRIC_DESC = ""
    '                End If


    '                DESCHASH = SALES_DIVISION_NAME & SUB_BODY_DESC & FABRIC_DESC
    '                DESCHASH = Replace(DESCHASH, " ", "")
    '                DESCHASH = Replace(DESCHASH, "/", "")
    '                DESCHASH = Replace(DESCHASH, ".", "")
    '                DESCHASH = Replace(DESCHASH, ",", "")
    '                DESCHASH = Replace(DESCHASH, "&", "")
    '                ' DGJ 
    '                PDF_FN = DESCHASH & PDF_FN


    '                '      ASCMAIN1.sql = "Select * from SOTSDIV1 " _
    '                '      & " where SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE_COMB & "'" _
    '                '      Dim rowSOTSDIV1 As DataRow = ASCDATA1.GetDataRow

    '                '   If rowSOTSDIV1 IsNot Nothing Then
    '                'End If
    '                Dim sqlw As String

    '                If chkGROUP.Checked Then
    '                    sqlw = "SELECTED='2'" _
    '                     & " and ISNULL(STYLE_GROUP_CODE,'') = '" & STYLE_GROUP_CODE & "'" _
    '                     & " and ISNULL(SUB_BODY_CODE,'') = '" & SUB_BODY_CODE & "'" _
    '                     & " and ISNULL(FABRIC_CODE,'') = '" & FABRIC_CODE & "'"
    '                Else
    '                    sqlw = "SELECTED='2'" _
    '                    & " and ISNULL(SALES_DIVISION_CODE_COMB,'') = '" & SALES_DIVISION_CODE_COMB & "'" _
    '                    & " and ISNULL(STYLE_GROUP_CODE,'') = '" & STYLE_GROUP_CODE & "'" _
    '                    & " and ISNULL(FABRIC_CODE,'') = '" & FABRIC_CODE & "'"

    '                End If


    '                Dim STYLE_count As Integer = 0
    '                For Each row As DataRow In dst.Tables("ICTQUOT2").Select(sqlw, "SEQ")
    '                    STYLE_count += 1
    '                    row.Item("SELECTED") = "1"
    '                    SetRowImage(row)
    '                Next
    '            Else
    '                ' 1 report for every 10 Styles
    '                'For Each row As DataRow In dst.Tables("ICTQUOT2").Select()
    '                '    row.Item("IMAGE") = Null
    '                'Next
    '                Dim STYLE_count As Integer = 0
    '                For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='2'", "SEQ")
    '                    STYLE_count += 1
    '                    row.Item("SELECTED") = "1"
    '                    SetRowImage(row)
    '                    If STYLE_count >= 10 And Not chkPublishPDF.Checked Then Exit For
    '                Next
    '                Application.DoEvents()
    '            End If

    '            'Runtime.GCSettings.LargeObjectHeapCompactionMode = Runtime.GCLargeObjectHeapCompactionMode.CompactOnce
    '            'GC.Collect()

    '            CR_params.Add("CHKOMITAVAIL", IIf(chkOmitAvail.Checked, "1", "0"))
    '            CR_params.Add("CHKOMITPRICE", IIf(chkOmitPrice.Checked, "1", "0"))
    '            CR_params.Add("CHKOMITPRICE2", "0")
    '            CR_params.Add("CHKSHOWRETAIL", IIf(chkShowRetail.Checked, "1", "0"))
    '            CR_params.Add("CHKSHOWSELECTEDONLY", IIf(chkShowSelectedOnly.Checked, "1", "0"))
    '            CR_params.Add("IMAGES_FOLDER", FOLDER_NAME)

    '            If RPT = "ICRQUOTV" Then
    '                Dim SHOWLASTSHIP = "0"
    '                'If chkShowLastShip.Checked Then
    '                '    SHOWLASTSHIP = "1"
    '                'End If
    '                CR_params.Add("SHOWLASTSHIP", SHOWLASTSHIP)
    '            End If

    '            For iCol As Integer = BegAlloPeriod To 4
    '                Dim sValue As String = ""
    '                If iCol = 0 Then sValue = "At Once"
    '                If iCol = 1 Then sValue = Format(dte1.Value, "MM/dd")
    '                If iCol = 2 Then sValue = Format(dte2.Value, "MM/dd")
    '                If iCol = 3 Then sValue = Format(dte3.Value, "MM/dd")
    '                If iCol = 4 Then sValue = "Beyond"
    '                If ColVisible(iCol) And (iCol >= BegAlloPeriod) Then
    '                    CR_params.Add("DTE" & CStr(iCol), sValue)
    '                Else
    '                    'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
    '                    CR_params.Add("DTE" & CStr(iCol), "")
    '                    'Else
    '                    '    CR_params.Add("DTE" & CStr(iCol), "")
    '                    'End If
    '                End If
    '            Next
    '            If BegAlloPeriod > 0 Then
    '                For bp As Int64 = 0 To BegAlloPeriod - 1
    '                    CR_params.Add("DTE" & CStr(bp), "")
    '                Next
    '            End If
    '            For Each row As DataRow In dst.Tables("ICTSTYC1").Select()
    '                Dim skipRow As Boolean = True
    '                For iCol As Integer = 0 To 4
    '                    If ColVisible(iCol) And (iCol >= BegAlloPeriod) Then
    '                        If Val(row.Item("QTY_AVA" & iCol) & "") > 0 And skipRow Then
    '                            skipRow = False
    '                        End If
    '                    End If
    '                Next
    '                If skipRow Then
    '                    If chkShowBlank.Checked Then
    '                        row.Item("SKIP_COLOR") = "0"
    '                    Else
    '                        row.Item("SKIP_COLOR") = "1"
    '                    End If
    '                Else
    '                    row.Item("SKIP_COLOR") = "0"
    '                End If
    '            Next

    '            If chk1perPage.Checked Then
    '                CR_params.Add("TXTSTYLE_CODE", "")
    '            End If

    '            Dim tempFileName As String = ""
    '            Do
    '                REPORT_INDEX += 1
    '                tempFileName = rowICTQUOT1.Item("QUOTE_NO") & "-" & Format(REPORT_INDEX, "000")



    '            Loop While My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")

    '            Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)

    '            If chkPublishPDF.Checked Then




    '                Dim PDFD As String = ASCMAIN1.Folders("Archive") & "QuotePDFs\" & SESSION_NO
    '                If Not My.Computer.FileSystem.DirectoryExists(PDFD) Then
    '                    My.Computer.FileSystem.CreateDirectory(PDFD)
    '                End If

    '                My.Computer.FileSystem.CopyFile _
    '            (ASCMAIN1.Folders("Temp") & tempFileName & ".PDF",
    '             PDFD & "\" & PDF_FN & ".PDF", True)
    '                Dim urlpfx As String = "http: //dataservice.absolut1.net/Pictures/StyleCADs/"
    '                Dim link As String = "<a href='" & urlpfx & PDF_FN & ".PDF'>Click here for " & PDF_FN & "</a>"
    '                PDF_LINKS &= vbCrLf & link

    '                FILE_NO += 1

    '                Dim strToHash As String = ASCMAIN1.Get_Hash(SESSION_NO & FILE_NO & PDF_FN)

    '                Dim ICTQUOH2 As DataRow = dst.Tables("ICTQUOH2").NewRow
    '                ICTQUOH2.Item("SESSION_NO") = SESSION_NO
    '                ICTQUOH2.Item("FILE_NO") = FILE_NO
    '                ICTQUOH2.Item("FILENAME") = PDF_FN
    '                'DGJ
    '                ICTQUOH2.Item("HASHVALUE") = DESCHASH & strToHash
    '                ' ICTQUOH2.Item("HASHVALUE") = strToHash
    '                ICTQUOH2.Item("SUB_BODY_DESC") = SUB_BODY_DESC
    '                ICTQUOH2.Item("SALES_DIVISION_NAME") = SALES_DIVISION_NAME
    '                ICTQUOH2.Item("FABRIC_DESC") = FABRIC_DESC
    '                dst.Tables("ICTQUOH2").Rows.Add(ICTQUOH2)


    '                '                  Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow
    '                '                  If rowARTCUST1 IsNot Nothing Then
    '                '                      Dim rowARTCUSTX As DataRow = dst.Tables("ARTCUSTX").NewRow
    '                '                      With rowARTCUSTX
    '                '.Item("CUST_SHIP_TO_CODE") = rowARTCUST1.Item("CUST_SHIP_TO_CODE") & 
    '                '                      End With
    '                '                      dst.Tables("ARTCUSTX").Rows.Add(rowARTCUSTX)
    '                '                  End If
    '                '                  dst.Tables("ARTCUSTX").AcceptChanges()


    '            End If

    '            'Show_Document(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
    '            ListPDFSheets.Add(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
    '            Print_Report_End(, True)

    '            ' Generate_Report(RPT, "Quote Sheet")

    '            For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
    '                row.Item("SELECTED") = "3"
    '                row.Item("IMAGE") = DBNull.Value
    '            Next
    '        Loop

    '        'Print_Report_End(, True)

    '        For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='3'")
    '            row.Item("SELECTED") = "1"
    '        Next

    '        If chkPublishPDF.Checked Then
    '            Dim ATTACHMENTs As New Dictionary(Of String, String)
    '            Dim CUST_CODE As String = txtQuoteCUST_CODE.Text
    '            Dim CUST_NAME As String = txtQuoteCUST_NAME.Text
    '            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
    '            '  ATTACHMENTs.Add(tempFileName & ".pdf", ASCMAIN1.Folders("Temp") & tempFileName & ".pdf")
    '            ATTACHMENTs.Add("BODY", PDF_LINKS)

    '            Dim SUBJECT As String = "Quote Sheet"
    '            Dim PFX As String = ""

    '            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
    '            If CUST_CODE <> "" Then
    '                '   EMAIL_ADDRESSs.Add(rowARTCUST1.Item("CUST_EMAIL") & "", rowARTCUST1.Item("CUST_CONTACT") & "")
    '                EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL & "", ASCMAIN1.USER_NAME & "")
    '            End If

    '            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
    '                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
    '                    SUBJECT, "ICFQUOTV", False, True, CUST_CODE, CUST_NAME, "Customer")
    '            If SEND_NO <> "" Then
    '                TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "QUOEML", "Quote Sheet emailed", SEND_NO)
    '            End If
    '            Dim sqlDelete As String = "SESSION_NO = '" & SESSION_NO & "'"
    '            ' & " and SET_ID = '" & SET_ID & "'"
    '            Update_Record_TDA("ICTQUOH1", sqlDelete)
    '            Update_Record_TDA("ICTQUOH2", sqlDelete)

    '            Fill_Records("ICTQUOHF", SESSION_NO)
    '            addExtraICTQUOHF()

    '            Using sw As New System.IO.StreamWriter(ASCMAIN1.Folders("Temp") & SESSION_NO & ".TXT")


    '                For Each rowICTQUOH2 As DataRow In dst.Tables("ICTQUOH2").Select("")
    '                    sw.WriteLine(rowICTQUOH2.Item("SALES_DIVISION_NAME") & " - " & rowICTQUOH2.Item("SUB_BODY_DESC") & " - " & rowICTQUOH2.Item("FABRIC_DESC"))
    '                    sw.WriteLine(LINEPFX & rowICTQUOH2.Item("HASHVALUE"))
    '                    sw.WriteLine()
    '                Next

    '            End Using
    '            dst.Tables("ICTQUOH1").Rows.Clear()
    '            dst.Tables("ICTQUOH2").Rows.Clear()

    '            Show_Document(ASCMAIN1.Folders("Temp") & SESSION_NO & ".TXT")

    '        End If

    '    End If

    '    For Each row As DataRow In dst.Tables("ICTQUOT2").Select("SELECTED='1'")
    '        row.Item("IMAGE") = Nothing
    '    Next

    '    For Each PDF As String In ListPDFSheets
    '        Show_Document(PDF)
    '    Next

    '    If blnShowSelected Then
    '        chkShowSelectedOnly.Checked = False
    '    End If
    'End Sub

End Class
