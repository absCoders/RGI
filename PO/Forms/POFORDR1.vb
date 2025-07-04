Public Class POFORDR1

    ' SEE Prepare_for_View_Lookup_Special - WE ARE NOT DOING CODE LOOKUPS FOR SELECT DISTINCT
    'see proceed_prereq allowing mtc to pos forever - no check on status

    Dim PO_ORDER_NO As String
    Dim rowAPTVEND1 As DataRow
    Dim rowPOTORDR1 As DataRow
    Dim rowPOTORDR2 As DataRow
    Dim COLOR_CODEs As New List(Of String)
    Dim PPK_CODE_ctr As Int32
    Dim sqlPOTORDRS As String

    Dim STYLE_CODE As String
    Dim CMT_NO As String
    Dim PO_ORDER_LNO As Integer
    Dim user_assigned As Boolean
    Dim OOBAL As Boolean
    Dim No_Costs As Boolean

    Dim APPR_NOTES As String
    Dim APPR_DECISION As String
    Dim APPR_BY As String

    Dim confirm_notes_mode As Boolean

    Dim ETD_to_ETA As Integer

    Dim PO_ORDER_NO_clone As String
    Dim rowPOTORDR1_clone As DataRow

    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing

    Dim PO_PARM_AUTOPO_FOLDER As String = "R:\AutoPO\"
    Dim rowPOTXLSF1 As DataRow = Nothing

    Dim blnAutomatic As Boolean = False
    Dim fix_ICTSTYL1_packs As Boolean = False
    Dim PO_PARM_PO_IMG_DIR As String = ""
    Dim subUPCSupport As Boolean = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")



#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            cmdAutoPO.Visible = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")
        Get_PARM("SOTPARM1")

        Get_PARM("POTPARM1")
        PO_PARM_PO_IMG_DIR = ROWs("POTPARM1").Item("PO_PARM_PO_IMG_DIR") & ""

        With dst
            ASCMAIN1.sql = "SELECT POTORDR1.*,SOTORDR1.CUST_NAME,SOTORDR1.ORDR_CANCEL_DATE from POTORDR1,SOTORDR1 where SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "", 0)
            With .Tables("POTORDRX")
                .Columns.Add("PO_QTY_ORD", GetType(System.Int64))
                .Columns.Add("PO_QTY_SHP", GetType(System.Int64))
                .Columns.Add("PO_QTY_REC", GetType(System.Int64))
                .Columns.Add("PO_QTY_OPN", GetType(System.Int64))
                .Columns.Add("PO_AMT_ORD", GetType(System.Decimal))
                .Columns.Add("PO_AMT_SHP", GetType(System.Decimal))
                .Columns.Add("PO_AMT_REC", GetType(System.Decimal))
                .Columns.Add("PO_AMT_OPN", GetType(System.Decimal))
                .Columns.Add("PO_LINES_CONF", GetType(System.Int64))
                .Columns.Add("PO_LINES", GetType(System.Int64))
                .Columns.Add("PO_CTNS_ORD", GetType(System.Decimal))
                .Columns.Add("PO_CTNS_SHP", GetType(System.Decimal))
                .Columns.Add("PO_CTNS_OPN", GetType(System.Decimal))
                .Columns.Add("PO_CUBE_ORD", GetType(System.Decimal))
                .Columns.Add("PO_CUBE_SHP", GetType(System.Decimal))
                .Columns.Add("PO_CUBE_OPN", GetType(System.Decimal))
                .Columns.Add("PO_DATE_SHIP_BY_MIN", GetType(System.DateTime))
                .Columns.Add("PO_DATE_ETA_MIN", GetType(System.DateTime))
                .Columns.Add("PO_DATE_SHIP_BY_MAX", GetType(System.DateTime))
                .Columns.Add("PO_DATE_ETA_MAX", GetType(System.DateTime))
                .Columns.Add("PO_COST_COMM_MIN", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "POTORDR1", "*")
            With .Tables("POTORDR1")
                .Columns.Add("PO_DATE_RECEIVED", GetType(System.DateTime))
                .Columns.Add("PO_SOURCE_DOC")
            End With

            Create_TDA(.Tables.Add, "ICTCOLR1", "*", 0, False)

            ASCMAIN1.sql = "Select POTORDR2.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CASE_CUBE" & vbCrLf _
                & " from POTORDR2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "  and POTORDR2.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, True, "V", 2)
            .Tables("POTORDR2").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(PO_QTY_ORD,0) / ISNULL(CARTON_PACK_QTY,0))")
            .Tables("POTORDR2").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")
            .Tables("POTORDR2").Columns.Add("CONFIRMED")
            .Tables("POTORDR2").Columns("CONFIRMED").DefaultValue = "0"
            .Tables("POTORDR2").Columns.Add("COST_COMPLETE")
            .Tables("POTORDR2").Columns("COST_COMPLETE").DefaultValue = "0"

            With .Tables("POTORDR2").Columns
                .Add("CUST_UPC")
                .Add("CUST_SKU")
                .Add("CUST_STYLE_CODE")
                .Add("CUST_COLOR_CODE")
                .Add("CUST_SIZE_CODE")
                .Add("STYLE_RETAIL", GetType(System.Decimal))
            End With

            dst.Tables.Add("POTORDR2_LINE")
            dst.Tables("POTORDR2_LINE").Merge(dst.Tables("POTORDR2"))

            For Each T As String In New String() {"POTORDR2", "POTORDR2_LINE"}
                With .Tables(T)
                    .Columns.Add("LINE_CLOSED")
                    .Columns.Add("PO_QTY_ORD_DZ", GetType(System.Decimal))
                    .Columns.Add("PO_QTY_SHP_DZ", GetType(System.Decimal), "PO_QTY_SHP / (12 / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,SUB_UNIT_PACK_QTY))")
                    .Columns.Add("PO_QTY_REC_DZ", GetType(System.Decimal), "PO_QTY_REC / (12 / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,SUB_UNIT_PACK_QTY))")
                    .Columns.Add("PO_QTY_OPN_DZ", GetType(System.Decimal), "PO_QTY_OPN / (12 / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,SUB_UNIT_PACK_QTY))")

                    .Columns.Add("PO_COST_QUOTA_UN", GetType(System.Decimal))
                    .Columns.Add("PO_COST_OTHER_UN", GetType(System.Decimal))

                    .Columns.Add("PO_COST_COMM_UN", GetType(System.Decimal), "(ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER_UN,0)) * ISNULL(PO_COST_COMM,0) / 100")
                    .Columns.Add("PO_COST_BUFFER_UN", GetType(System.Decimal), "(ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER_UN,0)) * ISNULL(PO_COST_BUFFER,0) / 100")
                    .Columns.Add("PO_COST_COMM_DZ", GetType(System.Decimal), "PO_COST_COMM_UN * (12 / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,SUB_UNIT_PACK_QTY))")
                    .Columns.Add("PO_COST_BUFFER_DZ", GetType(System.Decimal), "PO_COST_BUFFER_UN * (12 / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,SUB_UNIT_PACK_QTY))")

                    .Columns.Add("PO_FIRST_COST_UN", GetType(System.Decimal), "ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER_UN,0)")
                    .Columns.Add("PO_FIRST_COST_DZ", GetType(System.Decimal), "ISNULL(PO_COST_VCOST_DZ,0) + ISNULL(PO_COST_MATLS_DZ,0) + ISNULL(PO_COST_OTHER,0)")

                    .Columns.Add("PO_COST_DZ", GetType(System.Decimal), "PO_COST * (12 / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,SUB_UNIT_PACK_QTY))")
                    .Columns.Add("PO_AMT_ORD", GetType(System.Decimal), "PO_COST * PO_QTY_ORD")
                    .Columns.Add("PO_AMT_SHP", GetType(System.Decimal), "PO_COST * PO_QTY_SHP")
                    .Columns.Add("PO_AMT_REC", GetType(System.Decimal), "PO_COST * PO_QTY_REC")
                    .Columns.Add("PO_AMT_OPN", GetType(System.Decimal), "PO_COST * PO_QTY_OPN")
                End With
            Next

            'Create_Relation("POTORDR1", "POTORDR2", "PO_ORDER_NO")

            'With .Tables("POTORDR1")
            '    .Columns.Add("PO_QTY_ORD", GetType(System.Int64), "SUM(CHILD(POTORDR1_POTORDR2).PO_QTY_ORD)")
            '    .Columns.Add("PO_QTY_SHP", GetType(System.Int64), "SUM(CHILD(POTORDR1_POTORDR2).PO_QTY_SHP)")
            '    .Columns.Add("PO_QTY_REC", GetType(System.Int64), "SUM(CHILD(POTORDR1_POTORDR2).PO_QTY_REC)")
            '    .Columns.Add("PO_QTY_OPN", GetType(System.Int64), "SUM(CHILD(POTORDR1_POTORDR2).PO_QTY_OPN)")
            '    .Columns.Add("PO_AMT_ORD", GetType(System.Decimal), "SUM(CHILD(POTORDR1_POTORDR2).PO_AMT_ORD)")
            '    .Columns.Add("PO_AMT_SHP", GetType(System.Decimal), "SUM(CHILD(POTORDR1_POTORDR2).PO_AMT_SHP)")
            '    .Columns.Add("PO_AMT_REC", GetType(System.Decimal), "SUM(CHILD(POTORDR1_POTORDR2).PO_AMT_REC)")
            '    .Columns.Add("PO_AMT_OPN", GetType(System.Decimal), "SUM(CHILD(POTORDR1_POTORDR2).PO_AMT_OPN)")
            'End With

            ASCMAIN1.sql = "Select POTORDR3.*, ICTCMTM4.COLOR_DESC, ICTCMTM4.FABRIC_NO, ICTCMTM4.COLOR_CODE" & vbCrLf _
               & " from POTORDR3,POTORDR2,ICTCMTM4" & vbCrLf _
               & " where POTORDR2.PO_ORDER_NO = POTORDR3.PO_ORDER_NO" & vbCrLf _
               & "   and POTORDR2.PO_ORDER_LNO = POTORDR3.PO_ORDER_LNO" & vbCrLf _
               & "   and ICTCMTM4.CMT_NO = POTORDR2.CMT_NO" & vbCrLf _
               & "   and ICTCMTM4.COLOR_NO = POTORDR3.COLOR_NO" & vbCrLf _
               & "   and POTORDR3.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR3", "**", 0, True, "V", 3)

            dst.Tables.Add("POTORDR3_LINE")
            dst.Tables("POTORDR3_LINE").Merge(dst.Tables("POTORDR3"))

            For Each T As String In New String() {"POTORDR3", "POTORDR3_LINE"}
                With .Tables(T)
                    .Columns.Add("TOTAL_COST", GetType(System.Decimal), "YARDS_CONSUMED * FABRIC_COST")
                    .Columns.Add("CONSUMPTION", GetType(System.Decimal), "YARDS_CONSUMED * YIELD_QTY")
                End With
            Next
            Create_Relation("POTORDR2", "POTORDR3", "PO_ORDER_NO,PO_ORDER_LNO")
            Create_Relation("POTORDR2_LINE", "POTORDR3_LINE", "PO_ORDER_NO,PO_ORDER_LNO")

            With dst.Tables("POTORDR2").Columns
                .Add("CONSUMPTION", GetType(System.Decimal), "SUM(CHILD(POTORDR2_POTORDR3).CONSUMPTION)")
            End With
            With dst.Tables("POTORDR2_LINE").Columns
                .Add("CONSUMPTION", GetType(System.Decimal), "SUM(CHILD(POTORDR2_LINE_POTORDR3_LINE).CONSUMPTION)")
            End With

            For Each T As String In New String() {"POTORDR2", "POTORDR2_LINE"}
                With .Tables(T)
                    .Columns.Add("PO_QTY_UOM_FACTOR", GetType(System.Int32), "13 - PO_QTY_UOM")
                    '.Columns.Add("CONSUMPTION", GetType(System.Decimal), "SUM(CHILD(" & T & "_POTORDR3).CONSUMPTION)")
                    .Columns.Add("CONSUMPTION_X", GetType(System.Decimal), "(13 - PO_QTY_UOM) * CONSUMPTION")
                End With
            Next

            ASCMAIN1.sql = "Select POTORDR4.*, ICTCMTM5.FABRIC_DESC" & vbCrLf _
                & " from POTORDR4,POTORDR2,ICTCMTM5" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTORDR4.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTORDR4.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTCMTM5.CMT_NO = POTORDR2.CMT_NO" & vbCrLf _
                & "   and ICTCMTM5.FABRIC_NO = POTORDR4.FABRIC_NO" & vbCrLf _
                & "   and POTORDR4.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR4", "**", 0, True, "V", 3)

            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select TATEVNT1.*" & vbCrLf _
                    & " from TATEVNT1" & vbCrLf _
                    & " where TATEVNT1.TABLE_NAME = 'POTORDR1'" & vbCrLf _
                    & "   and TATEVNT1.TABLE_KEY = :PARM1" & vbCrLf _
                    & "   and TATEVNT1.EVENT_TYPE IN ('LBL_PRT','LBL_VIEW','LBL_ACC')"
                Create_TDA(.Tables.Add, "POTORDRL", "**", 0, False, "V")

            End If

            If subUPCSupport Then
                ASCMAIN1.sql = "Select ICTXLSPS.* 
                    from ICTXLSPS, POTORDR2
                    where ICTXLSPS.STYLE_CODE = POTORDR2.STYLE_CODE 
                    AND ICTXLSPS.COLOR_CODE = POTORDR2.COLOR_CODE
                    AND POTORDR2.PO_ORDER_NO = :PARM1"
                Create_TDA(.Tables.Add, "ICTXLSPS", "**", 0, False, "V")
            End If

            dst.Tables.Add("POTORDR4_LINE")
            dst.Tables("POTORDR4_LINE").Merge(dst.Tables("POTORDR4"))

            For Each T As String In New String() {"POTORDR4", "POTORDR4_LINE"}
                With .Tables(T)
                    .Columns.Add("YIELD_QTY", GetType(System.Decimal))
                    .Columns.Add("PRODUCTION", GetType(System.Decimal), "CONSUMPTION_RATE * 1")
                End With
            Next

            Create_Relation("POTORDR2", "POTORDR4", "PO_ORDER_NO,PO_ORDER_LNO")
            Create_Relation("POTORDR2_LINE", "POTORDR4_LINE", "PO_ORDER_NO,PO_ORDER_LNO")

            ASCMAIN1.sql = "Select POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
                & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO" & vbCrLf _
                & ", POTSHIP2.TRAN_NO, POTSHIP2.PO_SOURCE_DOC" & vbCrLf _
                & ", POTSHIP2.PO_DATE_RECEIVED" & vbCrLf _
                & ", POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_REC" & vbCrLf _
                & " From POTSHIP1, POTSHIP2, POTSHIP3" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTSHIP3.PO_ORDER_NO = :PARM1" & vbCrLf
            Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "V", 4)

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'POTORDR1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V")
            .Tables("TATEVNT1").Columns.Add("ATTACHMENT_EXT")

            ASCMAIN1.sql = "Select * from POTORDXR where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDXR", "**", 0, True, "V")

            Create_TDA(.Tables.Add, "ICTSTYL1", "*", 1, True)


            With .Tables.Add("POTORDRT")
                .Columns.Add("LNO", GetType(System.Int64))
                .Columns.Add("TOTAL")
                .Columns.Add("UNITS", GetType(System.Int64))
                .Columns.Add("AMOUNT", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "POTORDR6", "*", 1)
            With .Tables("POTORDR6")
                .Columns.Add("PO_MESSAGE_IMG", GetType(System.Byte()))
            End With

            Create_TDA(.Tables.Add, "POTORDR7", "*", 1)
            Create_TDA(.Tables.Add, "POTORDR8", "*", 1)
            .Tables("POTORDR8").Columns.Add("UNITS", GetType(System.Int32), "QTY*IIF(ISNULL(DOZENS,'0')='1',12,1)")

            Create_Relation("POTORDR7", "POTORDR8", "PO_ORDER_NO,CARTON_NO")
            .Tables("POTORDR7").Columns.Add("STYLES", GetType(System.Int32), "COUNT(CHILD(POTORDR7_POTORDR8).STYLE_CODE)")
            .Tables("POTORDR7").Columns.Add("UNITS", GetType(System.Int32), "SUM(CHILD(POTORDR7_POTORDR8).UNITS)")
            .Tables("POTORDR7").Columns.Add("PPK_INNER_QTY_CALC", GetType(System.Int32), "SUM(CHILD(POTORDR7_POTORDR8).PPK_INNER_QTY)")

            .Tables("POTORDR8").Columns.Add("CARTONS", GetType(System.Int32), "PARENT(POTORDR7_POTORDR8).CARTONS")
            .Tables("POTORDR8").Columns.Add("TOTAL_UNITS", GetType(System.Int32), "ISNULL(UNITS,0) * ISNULL(CARTONS,0)")
            .Tables("POTORDR7").Columns.Add("TOTAL_UNITS", GetType(System.Int32), "SUM(CHILD(POTORDR7_POTORDR8).TOTAL_UNITS)")
            .Tables("POTORDR7").Columns.Add("STYLE_CODE_1", GetType(System.String), "MIN(CHILD(POTORDR7_POTORDR8).STYLE_CODE)")
            .Tables("POTORDR7").Columns.Add("COLOR_CODE_1", GetType(System.String), "MIN(CHILD(POTORDR7_POTORDR8).COLOR_CODE)")
            .Tables("POTORDR7").Columns.Add("ITEM_CODE", GetType(System.String), "IIF(ISNULL(PPK_CODE,'')='',ISNULL(STYLE_CODE_1,'') + ISNULL(COLOR_CODE_1,''),PPK_CODE)")


            With .Tables.Add("POTORDRR")
                .Columns.Add("PO_ORDER_NO")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("QTY_ORD", GetType(System.Int32))
                .Columns.Add("QTY_CTN", GetType(System.Int32))
                .Columns.Add("QTY_VAR", GetType(System.Int32), "ISNULL(QTY_ORD,0) - ISNULL(QTY_CTN,0)")
                .Columns.Add("COLOR_DESC")
                .PrimaryKey = New DataColumn() { .Columns("PO_ORDER_NO"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With

            Create_Relation("POTORDRR", "POTORDR2", "PO_ORDER_NO,STYLE_CODE,COLOR_CODE")
            .Tables("POTORDRR").Columns("QTY_ORD").Expression = "SUM(CHILD(POTORDRR_POTORDR2).PO_QTY_ORD)"
            Create_Relation("POTORDRR", "POTORDR8", "PO_ORDER_NO,STYLE_CODE,COLOR_CODE")
            .Tables("POTORDRR").Columns("QTY_CTN").Expression = "SUM(CHILD(POTORDRR_POTORDR8).TOTAL_UNITS)"

            sqlPOTORDRS = "Select STYLE_CODE" & vbCrLf _
                & ", MIN (PO_ORDER_NO) PO_ORDER_NO" & vbCrLf _
                & ", SUM (PO_QTY_OPN) PO_QTY_OPN, COUNT (DISTINCT PO_ORDER_NO) POS" & vbCrLf _
                & ", MIN (PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MIN" & vbCrLf _
                & ", MAX (PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MAX" & vbCrLf _
                & ", MIN (SUB_UNIT_PACK_QTY) SUB_UNIT_PACK_QTY_MIN" & vbCrLf _
                & ", MAX (SUB_UNIT_PACK_QTY) SUB_UNIT_PACK_QTY_MAX" & vbCrLf _
                & ", MIN (CARTON_PACK_QTY) CARTON_PACK_QTY_MIN" & vbCrLf _
                & ", MAX (CARTON_PACK_QTY) CARTON_PACK_QTY_MAX" & vbCrLf _
                & ", MIN (INNER_PACK_QTY) INNER_PACK_QTY_MIN" & vbCrLf _
                & ", MAX (INNER_PACK_QTY) INNER_PACK_QTY_MAX" & vbCrLf _
                & " from POTORDR2 where PO_STATUS = 'O' group by STYLE_CODE"
            sqlPOTORDRS = "Select X.*" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_UNIT_PACK_QTY, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CARTON_PACK_QTY CARTON_PACK_QTY_STYLE" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE" & vbCrLf _
                & " from (" & sqlPOTORDRS & ") X, ICTSTYL1, POTORDR1" _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" _
                & "   and POTORDR1.PO_ORDER_NO (+) = X.PO_ORDER_NO"
            ASCMAIN1.sql = Replace(sqlPOTORDRS, "group by", " and ROWNUM < 1 group by")
            Create_TDA(.Tables.Add, "POTORDRS", "**", 0, False, "", 1)
            .Tables("POTORDRS").Columns.Add("STYLE_ACTION")
            .Tables("POTORDRS").Columns.Add("POS_UPDATED")
            .Tables("POTORDRS").Columns.Add("POS_SKIPPED")

            With .Tables.Add("ICTCOLRS")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("COLOR_DESC")
                .PrimaryKey = New DataColumn() { .Columns("COLOR_CODE")}
            End With

            ASCMAIN1.sql = "Select * from SOTWORK1 where WO_REF_TYPE = 'P' and WO_REF_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTWORK1", "**", 0, , "V", 1)
            ASCMAIN1.sql = "Select * from SOTWORK2 where WO_NO in " _
                & " (Select WO_NO from SOTWORK1 where WO_REF_TYPE = 'P' and WO_REF_NO = :PARM1)"
            Create_TDA(.Tables.Add, "SOTWORK2", "**", 0, , "V", 1)

            Create_TDA(.Tables.Add, "POTORDRN", "*", 1)
            Create_TDA(.Tables.Add, "POTORDRH", "*", 1)

            Create_TDA(.Tables.Add, "POTORDRZ", "*", 1, False)
            With .Tables("POTORDRZ").Columns
                .Add("STYLE_CODE_PREV")
                .Add("COLOR_CODE_PREV")
                .Add("PO_QTY_ORD_PREV", GetType(System.Int64))
                .Add("PO_COST_PREV", GetType(System.Decimal))
                .Add("PO_DATE_SHIP_BY_PREV", GetType(System.DateTime))
                .Add("PO_STATUS_PREV")
                .Add("CARTON_PACK_QTY_PREV", GetType(System.Int64))
            End With

            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, True, "V")


            ASCMAIN1.sql = "Select POTLCST2.*" & vbCrLf _
                & ", POTLCST1.VEND_CODE, POTLCST1.COST_CATGY_CODE, POTLCST1.COST_ACT, POTLCST1.VOUCHER_NO" & vbCrLf _
                & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE" & vbCrLf _
                & " from POTLCST2,POTLCST1,APTINVH1" & vbCrLf _
                & " where POTLCST2.PO_ORDER_NO = :PARM1 " & vbCrLf _
                & "   and POTLCST1.CTL_NO = POTLCST2.CTL_NO" & vbCrLf _
                & "   and APTINVH1.VOUCHER_NO (+) = POTLCST1.VOUCHER_NO"
            Create_TDA(.Tables.Add, "POTLCST2", "**", 0, True, "V")

            Create_TDA(.Tables.Add, "ASTATTA2", "*")

            ASCMAIN1.sql = "Select * from POTXLSF0 where PO_XLS_STATUS = '0'"
            Create_TDA(.Tables.Add, "POTXLSF0", "**", 0)

            Create_TDA(.Tables.Add, "POTXLSF1", "*", 1)

            Create_TDA(.Tables.Add, "POTXLSF2", "*", 1)
            With .Tables("POTXLSF2")
                .Columns.Add("STYLE_DESC")
                .Columns.Add("COLOR_DESC")
                .Columns.Add("SUB_UNIT_PACK_QTY", GetType(System.Int64))

                .Columns.Add("EXT_COST", GetType(System.Decimal), "ISNULL(PO_QTY,0) * ISNULL(PO_COST,0)")
                .Columns.Add("EXT_OTHER", GetType(System.Decimal), "ISNULL(PO_QTY,0) * ISNULL(PO_OTHER,0)")
                .Columns.Add("EXT_OTHER2", GetType(System.Decimal), "ISNULL(PO_QTY,0) * ISNULL(PO_OTHER2,0)")
                .Columns.Add("PO_AMT", GetType(System.Decimal), "EXT_COST + EXT_OTHER")
                .Columns.Add("PO_AMT2", GetType(System.Decimal), "EXT_COST + EXT_OTHER + EXT_OTHER2")
            End With

            Create_Relation("POTXLSF1", "POTXLSF2", "PO_XLS_NO,XLS_ORDER_NO")

            With .Tables("POTXLSF1")
                .Columns.Add("TOTAL_COST", GetType(System.Decimal), "SUM(CHILD.EXT_COST)")
                .Columns.Add("TOTAL_OTHER", GetType(System.Decimal), "SUM(CHILD.EXT_OTHER)")
                .Columns.Add("TOTAL_AMT", GetType(System.Decimal), "SUM(CHILD.PO_AMT)")
                .Columns.Add("TOTAL_OTHER2", GetType(System.Decimal), "SUM(CHILD.EXT_OTHER2)")
                .Columns.Add("TOTAL_AMT2", GetType(System.Decimal), "SUM(CHILD.PO_AMT2)")
                .Columns.Add("TOTAL_DZS", GetType(System.Decimal), "SUM(CHILD.PO_DZS)")
                .Columns.Add("TOTAL_QTY", GetType(System.Int64), "SUM(CHILD.PO_QTY)")
                .Columns.Add("COUNT_STYLES", GetType(System.Int64), "COUNT(CHILD.STYLE_CODE)")
            End With

            Create_TDA(.Tables.Add, "ICTFACT1", "*", 1)

            If ASCMAIN1.CLIENT = "VAN" Then
                If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    'Stop
                Else
                    ASCMAIN1.sql = "SELECT POTORDRA.*" & vbCrLf _
                    & ", x.`SentSeq` SENDNO, x.`SendRemarks` REMARKS, x.`StyleNo` STYLE_NO, x.`Style` STYLE_DESC, x.`OrderDate` ORDR_DATE, x.`OrderCfmDate` ORDR_CONF_DATE" & vbCrLf _
                    & ", x.`VandaleShipDate` SHIP_DATE, x.`TotalQty` TOTAL_QTY, x.`TotalVandaleAmount` TOTAL_AMT" & vbCrLf _
                    & ", x.`CreateBy` CREATED_BY, x.`VandaleUser` VANDALE_USER, x.`FollowBy` FOLLOWED_BY, x.`FollowByEmail` FOLLOWED_BY_EMAIL" & vbCrLf _
                    & " from POTORDRA, AT.`pohdr` X where X.VAN_REF = POTORDRA.VAN_REF" & vbCrLf _
                    & " and POTORDRA.PONO = :PARM1"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", ChrW(34))
                    Create_TDA(.Tables.Add, "POTORDRA", "**", 0, False, "V", 0)
                End If

                With .Tables.Add("POTORDS1")
                    .Columns.Add("RECORD_NO")
                    .Columns.Add("VESSEL_NAME")
                    .Columns.Add("SHIPPED", GetType(System.DateTime))
                    .Columns.Add("ETD", GetType(System.DateTime))
                    .Columns.Add("ETA", GetType(System.DateTime))
                    .Columns.Add("PORT_CODE_ORIG")
                    .Columns.Add("PORT_CODE_DEST")
                    .PrimaryKey = New DataColumn() { .Columns("RECORD_NO")}
                End With

                With .Tables.Add("POTORDS2")
                    .Columns.Add("RECORD_NO")
                    .Columns.Add("LNO", GetType(System.Int32))
                    .Columns.Add("FACTORY_NAME")
                    .Columns.Add("CONTAINER_SIZE")
                    .Columns.Add("SHIPPER")
                    .Columns.Add("INVOICE_NO")
                    .Columns.Add("CONTAINER")
                    .Columns.Add("CBM", GetType(System.Decimal))
                    .PrimaryKey = New DataColumn() { .Columns("RECORD_NO"), .Columns("LNO")}
                End With

                With .Tables.Add("POTORDS3")
                    .Columns.Add("RECORD_NO")
                    .Columns.Add("LNO", GetType(System.Int32))
                    .Columns.Add("LNO2", GetType(System.Int32))
                    .Columns.Add("MERCHANDISER")
                    .Columns.Add("PO_REFERENCE")
                    .Columns.Add("STYLE")
                    .Columns.Add("QTY", GetType(System.Int32))
                    .Columns.Add("UM")
                    .Columns.Add("CTN", GetType(System.Int32))
                    .Columns.Add("UNITS", GetType(System.Int32), "IIF(UM='DZ',QTY*12,QTY)")
                    .Columns.Add("UNITS_PO", GetType(System.Int32))
                    .Columns.Add("DZ_PO", GetType(System.Decimal))
                    .PrimaryKey = New DataColumn() { .Columns("RECORD_NO"), .Columns("LNO"), .Columns("LNO2")}
                End With

                With .Tables.Add("POTORDS4")
                    .Columns.Add("RECORD_NO")
                    .Columns.Add("LNO", GetType(System.Int32))
                    .Columns.Add("LNO2", GetType(System.Int32))
                    .Columns.Add("PO_ORDER_NO")
                    .Columns.Add("PO_ORDER_LNO", GetType(System.Int32))
                    .Columns.Add("STYLE_CODE")
                    .Columns.Add("COLOR_CODE")
                    .Columns.Add("PO_QTY_OPN", GetType(System.Int32))
                    .Columns.Add("SUB_UNIT_PACK_QTY", GetType(System.Int32))
                    .Columns.Add("PO_QTY_OPN_DZ", GetType(System.Decimal))
                    .Columns.Add("PO_DATE_SHIP_BY", GetType(System.DateTime))
                    .Columns.Add("PO_DATE_ETA", GetType(System.DateTime))
                    .Columns.Add("LAST_OPER_SHIP_BY")
                    .Columns.Add("LAST_DATE_SHIP_BY", GetType(System.DateTime))
                    .Columns.Add("PO_CONF_NO")
                    .Columns.Add("PO_CONF_DATE", GetType(System.DateTime))
                    .Columns.Add("PO_LINE_NOTE_INT")
                    .Columns.Add("SEL")
                    .Columns("SEL").DefaultValue = "0"
                    .PrimaryKey = New DataColumn() { .Columns("RECORD_NO"), .Columns("LNO"), .Columns("LNO2"), .Columns("PO_ORDER_NO"), .Columns("PO_ORDER_LNO")}
                End With

                Create_Relation("POTORDS1", "POTORDS2", "RECORD_NO")
                Create_Relation("POTORDS2", "POTORDS3", "RECORD_NO,LNO")
                Create_Relation("POTORDS3", "POTORDS4", "RECORD_NO,LNO,LNO2")

                .Tables("POTORDS3").Columns("UNITS_PO").Expression = "SUM(CHILD.PO_QTY_OPN)"
                '.Tables("POTORDS4").Columns("PO_QTY_OPN_DZ").Expression = "ISNULL(SUB_UNIT_PACK_QTY,1)*ISNULL(PO_QTY_OPN,0)/12"
                .Tables("POTORDS3").Columns("DZ_PO").Expression = "SUM(CHILD.PO_QTY_OPN_DZ)"
            End If
        End With

        If ASCMAIN1.CLIENT = "VAN" Then
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                Stop
            Else
                grdPOTORDRA.DataSource = dst.Tables("POTORDRA")
            End If
        End If

        grdPOTXLSF0.DataSource = dst.Tables("POTXLSF0")
        grdPOTXLSF1.DataSource = dst.Tables("POTXLSF1")

        Dim dvw As DataView = DirectCast(grdPOTXLSF1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "XLS_ORDER_STATUS = '0'"

        If ASCMAIN1.CLIENT = "VAN" Then
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                'Stop
            Else
                grdPOTORDS1.DataSource = dst.Tables("POTORDS1")
                grdPOTORDS4.DataSource = dst.Tables("POTORDS4")

                Create_Summary(grdPOTORDS4, "SEL")
                Create_Summary(grdPOTORDS4, "PO_ORDER_NO", "Count")

                'With grdPOTORDS1.DisplayLayout
                '    .Override.AllowColSizing = UltraWinGrid.AllowColSizing.Free
                '    .PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                'End With

                grdPOTORDS4.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTORDS4.DisplayLayout.Bands(0).Columns
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                    gcol.Header.Appearance.BackColor2 = Drawing.Color.DodgerBlue
                    If gcol.Key = "SEL" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                Next

                For B As Integer = 0 To 3
                    With grdPOTORDS1.DisplayLayout.Bands(B).Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        If B = 3 Then
                            .AllowUpdate = DefaultableBoolean.True
                        Else
                            .AllowUpdate = DefaultableBoolean.False
                        End If

                    End With
                    For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTORDS1.DisplayLayout.Bands(B).Columns
                        gcol.Header.Appearance.BackColor = Drawing.Color.White
                        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        If B = 3 Then
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.DodgerBlue
                            If gcol.Key = "SEL" Then
                                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                            Else
                                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                            End If
                        Else
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        End If
                    Next
                Next
            End If
        End If

        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdPOTORDR2.DataSource = dst.Tables("POTORDR2")
        'grdPOTORDR2.DataMember = "POTORDR2"
        'grdPOTORDR2.DataSource = dst ' .Tables("POTORDR2")
        grdPOTORDR3.DataSource = dst.Tables("POTORDR3_LINE")
        grdPOTORDR4.DataSource = dst.Tables("POTORDR4_LINE")
        grdPOTSHIPX.DataSource = dst.Tables("POTSHIPX")
        grdPOTORDRT.DataSource = dst.Tables("POTORDRT")

        grdPOTORDR6.DataSource = dst.Tables("POTORDR6")
        grdPOTORDR7.DataSource = dst.Tables("POTORDR7")
        grdPOTORDR8.DataSource = dst.Tables("POTORDR8")
        grdPOTORDRR.DataSource = dst.Tables("POTORDRR")
        grdPOTORDRS.DataSource = dst.Tables("POTORDRS")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdPOTORDXR.DataSource = dst.Tables("POTORDXR")
        grdPOTORDRN.DataSource = dst.Tables("POTORDRN")
        grdPOTORDRH.DataSource = dst.Tables("POTORDRH")
        grdPOTORDRZ.DataSource = dst.Tables("POTORDRZ")

        grdPOTLCST2.DataSource = dst.Tables("POTLCST2")

        If ASCMAIN1.CLIENT = "RGI" Then
            tabPOTORDR2.Tabs("Labels").Visible = True
            grdPOTORDRL.DataSource = dst.Tables("POTORDRL")
        Else
            tabPOTORDR2.Tabs("Labels").Visible = False
        End If

        If subUPCSupport Then
            grdICTXLSPS.DataSource = dst.Tables("ICTXLSPS")
            Create_Summary(grdICTXLSPS, "SET_LNO", "Count")
            Sort_grdColumns(grdICTXLSPS, "SET_LNO", True)
        End If

        tabPOTORDR2.Tabs("XLS").Visible = False ' (ASCMAIN1.CLIENT = "VAN")

        Bind_Controls(grpPOTORDR2, "POTORDR2_LINE")
        Bind_Controls(grpPOTORDR2X, "POTORDR2_LINE")
        Bind_Controls(grpDetails, "POTORDR2_LINE")
        Bind_Controls(grpSOTORDR1, "SOTORDR1")

        Set_Read_Only(grpSOTORDR1, True)

        With grdPOTXLSF1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_XLS_NO", "XLS_ORDER_NO",
                                 "TOTAL_COST", "TOTAL_OTHER", "TOTAL_AMT", "TOTAL_OTHER2", "TOTAL_AMT2", "TOTAL_DZS", "TOTAL_QTY", "COUNT_STYLES"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next
        End With

        With grdPOTXLSF1.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_XLS_NO", "XLS_ORDER_NO", "XLS_ORDER_LNO", "XLS_STYLE_CODE", "XLS_COLOR_CODE",
                                 "STYLE_DESC", "COLOR_DESC", "SUB_UNIT_PACK_QTY", "PO_DZS", "EXT_COST", "EXT_OTHER", "PO_AMT", "EXT_OTHER2", "PO_AMT2"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next
        End With

        With grdPOTORDR6.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "PO_ORDER_MLNO" Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If gcol.Key = "PO_MESSAGE_COST" Then
                    gcol.MaskInput = "nnn.nnnnnn"
                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        gcol.Format = "#.0000"
                    Else
                        gcol.Format = "#.000000"
                    End If
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.Tomato
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With


        With grdPOTORDR3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "YIELD_QTY" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdPOTORDR4.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "CONSUMPTION_RATE" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdPOTORDR2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN",
                                 "PO_QTY_ORD_DZ", "PO_QTY_SHP_DZ", "PO_QTY_REC_DZ", "PO_QTY_OPN_DZ",
                                 "LINE_CLOSED"}.Contains(gcol.Key) Then
                    If gcol.Key.EndsWith("_DZ") Then
                        gcol.Format = "#.00"
                    ElseIf gcol.Key <> "LINE_CLOSED" Then
                        gcol.Format = "#,##0"
                    End If

                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    If gcol.Key.EndsWith("_DZ") Then gcol.Header.Appearance.BackColor2 = Drawing.Color.Turquoise

                ElseIf New String() {"PO_COST_VCOST", "PO_COST_MATLS", "PO_COST",
                                     "PO_COST_VCOST_DZ", "PO_COST_MATLS_DZ", "PO_COST_DZ",
                                     "PO_COST_COMM", "PO_COST_OTHER", "PO_COST_OTHER_UN", "PO_FIRST_COST_UN", "PO_FIRST_COST_DZ"}.Contains(gcol.Key) Then
                    If gcol.Key.EndsWith("_DZ") Or gcol.Key = "PO_COST_OTHER" Or gcol.Key = "PO_COST_COMM" Then
                        gcol.Format = "#.00"
                    Else
                        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                            gcol.Format = "#.0000"
                        Else
                            gcol.Format = "#.000000"
                        End If
                    End If
                    If gcol.Key = "PO_COST" Or gcol.Key = "PO_COST_DZ" Or gcol.Key = "PO_FIRST_COST_DZ" Or gcol.Key = "PO_FIRST_COST_DZ" Then gcol.CellAppearance.BackColor = Drawing.Color.LightPink
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Fuchsia
                    ' If gcol.Key.EndsWith("_DZ") Or gcol.Key = "PO_COST_OTHER" Then gcol.Header.Appearance.ForeColor = Drawing.Color.White
                ElseIf New String() {"PO_COST_QUOTA", "PO_COST_QUOTA_UN", "PO_COST_BUFFER", "DFQUOTA"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                    If gcol.Key = "PO_COST_QUOTA" Or gcol.Key = "PO_COST_BUFFER" Then
                        gcol.Format = "#.00"
                    ElseIf gcol.Key = "PO_COST_QUOTA_UN" Then
                        gcol.Format = "#.000000"
                    End If
                    ' If gcol.Key.EndsWith("_DZ") Or gcol.Key = "PO_COST_QUOTA" Then gcol.Header.Appearance.ForeColor = Drawing.Color.White
                ElseIf New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA", "CONFIRMED", "PO_CONF_NO", "PO_CONF_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE", "PO_SHIPMENT_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                For Each COLUMN_NAME As String In New String() {"CUST_UPC", "CUST_SKU", "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "STYLE_RETAIL"}
                    With .Columns(COLUMN_NAME)
                        .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                        .Header.Appearance.BackColor = Drawing.Color.White
                        .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                    End With
                Next
            Next

            For Each COLUMN_NAME In New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN",
                                                "PO_QTY_ORD_DZ", "PO_QTY_SHP_DZ", "PO_QTY_REC_DZ", "PO_QTY_OPN_DZ"}
                .Columns(COLUMN_NAME).Width = 80

            Next
            For Each COLUMN_NAME In New String() {"PO_COST_VCOST", "PO_COST_MATLS", "PO_COST",
                                     "PO_COST_VCOST_DZ", "PO_COST_MATLS_DZ", "PO_COST_DZ",
                                     "PO_COST_COMM", "PO_COST_OTHER", "PO_COST_OTHER_UN",
                                     "PO_COST_QUOTA", "PO_COST_QUOTA_UN", "PO_COST_BUFFER", "PO_FIRST_COST_UN", "PO_FIRST_COST_DZ"}

                .Columns(COLUMN_NAME).Width = 80
            Next
        End With


        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRX, New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN",
                                                  "PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN",
                                                  "PO_CTNS_ORD", "PO_CTNS_SHP", "PO_CTNS_OPN",
                                                  "PO_CUBE_ORD", "PO_CUBE_SHP", "PO_CUBE_OPN"}, , , "#,##0")

        Create_Summary(grdPOTORDR2, "PO_ORDER_LNO", "Count")
        Create_Summary(grdPOTORDR2, New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN", "TOTAL_CARTONS", "TOTAL_CUBE",
                                                  "PO_QTY_ORD_DZ", "PO_QTY_SHP_DZ", "PO_QTY_REC_DZ", "PO_QTY_OPN_DZ",
                                                  "PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"})

        Create_Summary(grdPOTORDR3, "COLOR_NO", "Count")
        Create_Summary(grdPOTORDR3, New String() {"YARDS_CONSUMED"}) ' , "YIELD_QTY", "TOTAL_COST"})

        'Create_Summary(grdPOTORDR4, "FABRIC_NO", "Count")
        'Create_Summary(grdPOTORDR4, New String() {"CONSUMPTION_RATE", "PRODUCTION"})

        Create_Summary(grdPOTORDR6, "PO_ORDER_MLNO", "Count")
        Create_Summary(grdPOTORDR6, New String() {"PO_MESSAGE_COST"}) ' , "YIELD_QTY", "TOTAL_COST"})


        Create_Summary(grdPOTLCST2, "CTL_NO", "Count")
        Create_Summary(grdPOTLCST2, New String() {"COST_ACT_PO"})

        Create_Summary(grdPOTORDRR, "STYLE_CODE", "Count")
        Create_Summary(grdPOTORDRR, New String() {"QTY_ORD", "QTY_CTN", "QTY_VAR"})

        Create_Summary(grdPOTORDR7, "CARTON_NO", "Count")
        Create_Summary(grdPOTORDR7, New String() {"CARTONS", "UNITS", "TOTAL_UNITS"})

        Create_Summary(grdPOTORDR8, "STYLE_CODE", "Count")
        Create_Summary(grdPOTORDR8, New String() {"QTY", "UNITS", "TOTAL_UNITS"})

        Create_Summary(grdPOTORDRS, "STYLE_CODE", "Count")

        With grdPOTORDRS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "CARTON_PACK_QTY" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next
        End With

        With grdPOTORDRZ.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"PO_QTY_ORD", "PO_COST", "PO_DATE_SHIP_BY", "PO_STATUS", "CARTON_PACK_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"PO_QTY_ORD_PREV", "PO_COST_PREV", "PO_DATE_SHIP_BY_PREV", "PO_STATUS_PREV", "CARTON_PACK_QTY_PREV"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdPOTORDRN.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_ORDER_COMMENT"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("PO_ORDER_NO").Header.Fixed = True
            For Each COLUMN_NAME As String In New String() {"PO_STATUS", "FOB_CMT"}
                .Columns(COLUMN_NAME).Header.Appearance.TextHAlign = HAlign.Center
                .Columns(COLUMN_NAME).CellAppearance.TextHAlign = HAlign.Center
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"PO_CTNS_ORD", "PO_CTNS_SHP", "PO_CTNS_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"PO_CUBE_ORD", "PO_CUBE_SHP", "PO_CUBE_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_NO", "CUST_CODE", "CUST_NAME", "ORDR_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With
        grdPOTORDRX.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grdPOTORDRX, True)

        'grdPOTORDR2.DisplayLayout.UseFixedHeaders = True
        With grdPOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                    {"PO_ORDER_LNO", "STYLE_CODE", "COLOR_CODE", "STYLE_DESC",
                     "SUB_UNIT_PACK_QTY", "CARTON_PACK_QTY"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() _
                    {"PO_ORDER_LNO", "STYLE_DESC", "SUB_UNIT_PACK_QTY", "PO_COST", "PO_COST_DZ", "PO_FIRST_COST_UN", "PO_FIRST_COST_DZ",
                     "PO_QTY_SHP_DZ", "PO_QTY_SHP", "PO_QTY_REC_DZ", "PO_QTY_REC", "PO_QTY_OPN_DZ", "PO_QTY_OPN",
                     "INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE", "LAST_OPER_SHIP_BY", "LAST_DATE_SHIP_BY",
                     "SHIP_COST_CHANGE_USER", "SHIP_COST_CHANGE_DATE", "PO_ORIG_DATE_SHIP_BY", "PO_ORIG_DATE_ETA"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("LINE_FINISHED").CellActivation = UltraWinGrid.Activation.NoEdit ' NEED TO UNDERSTAND THIS ONE BETTER - MAYBE SOME SHOULD HAVE ACCESS
        End With

        With grdPOTSHIPX.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdPOTORDRT.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.Goldenrod
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("TOTAL").Header.Caption = ""
        End With

        With grdPOTORDR7.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"CARTON_NO", "PPK_CODE", "STYLES", "UNITS", "TOTAL_UNITS", "PPK_INNER_QTY_CALC"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        With grdPOTORDR7.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"CARTON_NO", "CARTONS", "CARTON_COMMENTS", "CUSTOM_PPK", "PPK_CODE", "PO_QTY_PER_CTN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.MediumAquamarine
                ElseIf New String() {"STYLE_CODE", "COLOR_CODE", "PPK_INNER_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.BurlyWood
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        With grdPOTORDR8.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"QTY", "DOZENS", "PPK_INNER_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"CARTON_NO", "STYLE_CODE", "COLOR_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME In New String() {"CARTON_NO", "STYLE_CODE", "COLOR_CODE", "UNITS", "TOTAL_UNITS"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        With grdPOTORDRR.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"QTY_CTN", "QTY_ORD", "QTY_VAR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                ElseIf New String() {"STYLE_CODE", "COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        If MENU_ITEM_OBJECT = "POFORDRI" Then
            grdPOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If

        ASCMAIN1.Add_Value_List(grdPOTORDRX, "PO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "FOB_CMT", Nothing, New String() {":", "F:FOB", "C:NonInv-CMT", "I:Invty-CMT", "B:BTB"})
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "LABEL_RESP_CODE")
        ASCMAIN1.Add_Value_List(grdPOTLCST2, "CHARGEBACK_STATUS", Nothing, New String() {":", "0:Absorb", "1:Pending", "2:Re-Billed"})
        ASCMAIN1.Add_Value_List(grdPOTORDR6, "PO_MESSAGE_TYPE")
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "PO_APPR_PENDING", Nothing, New String() {":", "0:WIP", "1:Queued"})


        InquiryMode = (MENU_ITEM_OBJECT = "POFORDRI")
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            No_Costs = Not (ASCMAIN1.USER_SECURITY_CODEs.Contains("X5") Or ASCMAIN1.USER_SECURITY_CODEs.Contains("X2"))
        Else
            No_Costs = False
        End If

        '' grdPOTORDR2.DisplayLayout.Bands(0).Columns("TOTAL_CUBE").Hidden = False

        If ASCMAIN1.CLIENT = "VAN" Then
            With grdPOTORDR2.DisplayLayout.Bands(0)
                .Columns("CASE_CUBE").Hidden = True
                .Columns("TOTAL_CARTONS").Hidden = True
                .Columns("TOTAL_CUBE").Hidden = True
            End With
            ASCMAIN1.Add_Value_List(grdPOTORDRA, "STATUS", Nothing, New String() {":", "W:Submitted", "X:Superceded", "A:Approved", "I:Imported", "R:Rejected", "D:Deleted"})



        Else
            Dim VL As New ValueList
            VL.ValueListItems.Add(New ValueListItem("F", "FOB"))
            VL.ValueListItems.Add(New ValueListItem("B", "BTB"))
            Absx1.optFor("FOB_CMT").ValueList = VL

            grdPOTORDR2.DisplayLayout.Bands(0).Columns("SUB_UNIT_PACK_QTY").Hidden = True
            grdPOTORDRS.DisplayLayout.Bands(0).Columns("SUB_UNIT_PACK_QTY").Hidden = True

            If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                With grdPOTORDR2.DisplayLayout.Bands(0)
                    .Columns("CASE_CUBE").Hidden = True
                    .Columns("TOTAL_CUBE").Hidden = True
                    ''  .Columns("TOTAL_CUBE").Hidden = False
                End With
            End If
        End If

        If ASCMAIN1.CLIENT = "NYA" Then
            lblPO_SPEC_ORDR_NO.Text = "Customer"
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_SPEC_ORDR_NO").Header.Caption = "Customer"
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY").Header.Caption = "ETD"
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY_MIN").Header.Caption = "ETD Min"
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY_MAX").Header.Caption = "ETD Max"
            lblPO_DATE_SHIP_BY.Text = "Orig ETD"

            grdPOTORDR2.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY").Header.Caption = "ETD"
            grdPOTORDR2.DisplayLayout.Bands(0).Columns("PO_ORIG_DATE_SHIP_BY").Header.Caption = "Orig ETD"

        End If

        chkBuyersCommision.Visible = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
        chkPPK.Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")
        Toggle_Ship_ETA_from_PO_Detail(False)
        tabPO.Tabs("XLS").Visible = (ASCMAIN1.CLIENT = "VAN")

        If ASCMAIN1.CLIENT = "VAN" Then
            Absx1.txtFor("PO_NOTES").Parent = grpHeaderData
            Absx1.txtFor("PO_NOTES").Left = Absx1.txtFor("PO_CONTACT").Left
            Absx1.txtFor("PO_NOTES").Top = Absx1.txtFor("PO_CONTACT").Top + Absx1.txtFor("PO_CONTACT").Height
        End If

        lblCUST_CODE.Visible = (ASCMAIN1.CLIENT = "VAN")
        txtCUST_CODE.Visible = (ASCMAIN1.CLIENT = "VAN")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("VEND_CODE")
                If cdr IsNot Nothing Then
                    If cdr.Item("VEND_STATUS") & "" <> "A" Then
                        EMsg &= vbCr & "Vendor is not Active"
                    End If
                    If cdr.Item("VEND_STOP_PURCHASE") & "" = "1" Then
                        EMsg &= vbCr & "Vendor is On Hold for Purchasing"
                    End If

                End If

                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    If Absx1.optFor("FOB_CMT").Value = "" Then
                        EMsg &= vbCr & "You Must First Select a PO Type (FOB, CMT, etc)"
                    End If
                End If

                If Absx1.txtFor("PO_REFERENCE").Text = "" Then
                    EMsg &= vbCr & "You Must First Supply A PO Reference"
                End If

                If ASCMAIN1.CLIENT = "VAN" And EMsg = "" Then
                    Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
                    If VEND_CODE = "AT" Then

                        Dim PO_REFERENCE As String = Absx1.txtFor("PO_REFERENCE").Text
                        ASCMAIN1.sql = "Select MAX(VAN_REF) from POTORDRA where (STATUS = 'I' or STATUS = 'A' or STATUS = 'W') and PONO = :PARM1"
                        Dim VAN_REF As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {PO_REFERENCE})
                        Dim msg As String = ""
                        If VAN_REF <> "" Then
                            ' msg = "Warning: This PO was automatically created from an AT Transmission"
                        Else
                            ' msg = "Warning: You know you should not be Entering New POs from AT - use PO Import"
                        End If
                        If msg <> "" Then
                            If MsgBox(msg, vbOKCancel, "Warning - PO Data may not reconcile with AT") = MsgBoxResult.Cancel Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

            Case "Clone"
                PO_ORDER_NO = Absx1.txtFor("PO_ORDER_NO").Text
                rowPOTORDR1 = LookUp("POTORDR1", PO_ORDER_NO)
                If rowPOTORDR1.Item("ORDR_NO") & "" <> "" Then
                    EMsg &= vbCr & "PO No " & PO_ORDER_NO & " was created by Back-to-Back Sales Order " & rowPOTORDR1.Item("ORDR_NO") _
                            & vbCrLf & "You man not clone a BTB PO"
                End If
                If rowPOTORDR1.Item("FOB_CMT") & "" <> "F" Then
                    EMsg &= vbCr & "PO No " & PO_ORDER_NO & " is Not an FOB PO - you cannot clone a PO unless the source PO is an FOB PO"
                End If

            Case "View", "Edit"
                If tabPOTORDR2.Tabs("Line Item Details").Visible Then Setup_POLine(False)

                Validate_Code("PO_ORDER_NO")

                If Absx1.txtFor("PO_ORDER_NO").Text = "" Then
                    EMsg &= vbCr & "You Must First Select a PO No"
                Else
                    PO_ORDER_NO = Absx1.txtFor("PO_ORDER_NO").Text
                    rowPOTORDR1 = LookUp("POTORDR1", PO_ORDER_NO)

                    If rowPOTORDR1 Is Nothing Then
                        EMsg &= vbCr & "PO No " & PO_ORDER_NO & " Not on File"
                    Else

                        If eItemKey = "Edit" Then
                            If ASCMAIN1.CLIENT = "VAN" Then
                                Dim VEND_CODE As String = rowPOTORDR1.Item("VEND_CODE") & ""
                                If VEND_CODE = "AT" Then
                                    '  msg = "Warning: You know you should not be Editing POs from AT"
                                    Dim PO_REFERENCE As String = rowPOTORDR1.Item("PO_REFERENCE") & ""
                                    ASCMAIN1.sql = "Select MAX(VAN_REF) from POTORDRA where (STATUS = 'I' or STATUS = 'A' or STATUS = 'W') and PONO = :PARM1"
                                    Dim VAN_REF As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {PO_REFERENCE})
                                    Dim msg As String = ""
                                    If VAN_REF <> "" Then
                                        '  msg = "Warning: This PO was automatically created from an AT Transmission"
                                        'If MsgBox(msg, vbOKCancel, "Warning - PO Data may not reconcile with AT") = MsgBoxResult.Cancel Then
                                        '    Exit Sub
                                        'End If
                                    End If

                                End If
                            End If

                            If rowPOTORDR1.Item("PO_STATUS") <> "O" Then
                                Select Case rowPOTORDR1.Item("PO_STATUS")
                                    Case "C"
                                        Dim special_exception = (ASCMAIN1.CLIENT = "NYA" And (ASCMAIN1.USER_ID = "lshalom" Or ASCMAIN1.USER_ID = "wjz"))
                                        If rowPOTORDR1.Item("ORDR_NO") & "" <> "" And Not special_exception Then
                                            EMsg = EMsg & vbCr & "PO " & PO_ORDER_NO & " is No Longer Open"
                                        Else
                                            Dim msgbox_caption = "PO Closed"
                                            If special_exception Then msgbox_caption = "PO Closed - Special Exception to Change PO Cost"
                                            If MsgBox("PO " & PO_ORDER_NO & " has been Closed or Cancelled." _
                                                       & vbCrLf & vbCrLf & "Would you like to re-open it?",
                                                       MsgBoxStyle.YesNo, msgbox_caption) = MsgBoxResult.Yes Then
                                                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                                                Record_Event("PO-OPN", "PO Re-Opened", True)
                                                ASCMAIN1.sql = "Update POTORDR1 Set PO_REVISION_NOTE = NULL, LAST_DATE = :PARM1, LAST_OPER = :PARM2, PO_DATE_CANCELLED = NULL where PO_ORDER_NO = :PARM3"
                                                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, PO_ORDER_NO})
                                                MsgBox("PO " & PO_ORDER_NO & " has been Re-Opened, but lines are still Closed (0 Qty Open)." _
                                                       & vbCrLf & vbCrLf & "You must Re-Open Lines individually.", MsgBoxStyle.OkOnly, "Success")
                                            Else
                                                EMsg = EMsg & vbCr & "PO " & PO_ORDER_NO & " has been Closed - Edit not Permitted"
                                            End If
                                        End If

                                    Case Else
                                        EMsg = EMsg & vbCr & "PO " & PO_ORDER_NO & " is No Longer Open"
                                End Select
                            End If
                        End If

                        If EMsg = "" Then
                            If rowPOTORDR1.Item("ORDR_NO") & "" <> "" And eItemKey = "Edit" Then
                                MsgBox("PO No " & PO_ORDER_NO & " was created by Back-to-Back Sales Order " & rowPOTORDR1.Item("ORDR_NO") _
                                        & vbCrLf & "Changes are limited", MsgBoxStyle.OkOnly, "You may not add or delete Styles or Change Style Qty")
                            End If
                        End If

                    End If
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                        If Not ASCMAIN1.Logical_Open("POTORDR1", "PO:" & rowPOTORDR1.Item("VEND_CODE")) Then Exit Sub
                    End If
                End If

            Case "Update"
                If tabPOTORDR2.Tabs("Line Item Details").Visible Then
                    MsgBox("You must first Update or Cancel Line Item Details", MsgBoxStyle.OkOnly, "Line Item Details Page is Open")
                    Exit Sub
                End If

                If Absx1.txtFor("PO_REFERENCE").Text = "" Then
                    EMsg &= vbCr & "PO Reference is Required"
                End If

                If dst.Tables("POTORDR2").Rows.Count = 0 Then
                    EMsg = EMsg & vbCr & "No Details Entered"
                End If

                If Absx1.txtFor("LABEL_RESP_CODE").Text = "V" Then
                    If rowAPTVEND1.Item("LABEL_RESP_CODE_NOT") & "" = "1" Then
                        EMsg &= vbCr & "This supplier does not accept responsibility for Labels"
                    End If
                End If

                Dim SQLW As String = "PO_DATE_SHIP_BY is Null or PO_DATE_ETA is Null or PO_DATE_SHIP_BY > PO_DATE_ETA"
                Dim C As Integer = Val(dst.Tables("POTORDR2").Compute("COUNT(PO_ORDER_LNO)", SQLW) & "")
                If C <> 0 Then
                    EMsg &= vbCr & "Detail Ship By and ETA Dates are Required for All Details"
                    EMsg &= vbCr & "Detail Ship By Dates must not be later than ETA Dates"
                End If

                If Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" = "" _
                Or Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Then
                    EMsg &= vbCr & "Header Ship By and ETA Dates are Required for all Purchase Orders"
                Else
                    If Format(Absx1.dteFor("PO_DATE_SHIP_BY").Value, "yyyyMMdd") _
                    >= Format(Absx1.dteFor("PO_DATE_ETA").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Header Ship By Date must not be later than ETA Date"
                    Else
                        If Absx1.dteFor("PO_DATE_CANCEL").Value & "" <> "" Then
                            If Format(Absx1.dteFor("PO_DATE_SHIP_BY").Value, "yyyyMMdd") _
                            >= Format(Absx1.dteFor("PO_DATE_CANCEL").Value, "yyyyMMdd") Then
                                EMsg &= vbCr & "Cancellation Date must be later than Ship By Date"
                            End If
                        End If
                    End If
                End If

                Check_ETA()

                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    SQLW = "PO_COST is Null"
                    C = Val(dst.Tables("POTORDR2").Compute("COUNT(PO_ORDER_LNO)", SQLW) & "")
                    If C <> 0 Then EMsg &= vbCr & "A Value for PO Cost is Required for All Details"
                    SQLW = "ISNULL(PO_QTY_ORD,0) = 0"
                    C = Val(dst.Tables("POTORDR2").Compute("COUNT(PO_ORDER_LNO)", SQLW) & "")
                    If C <> 0 Then EMsg &= vbCr & "A Non-Zero Value for PO Qty is Required for All Details"
                End If


                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Destination Warehouse Specified"
                Else
                    If EntryMode = "N" And rowICTWHSE1.Item("WHSE_TYPE") & "" = "P" Then
                        EMsg &= vbCr & "Cannot Start a New PO with a Warehouse that is actually a Port"
                    End If
                End If

                If txtPO_REVISION_NOTE.Visible And txtPO_REVISION_NOTE.Text = "" Then
                    If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    Else
                        EMsg &= vbCr & "You Must Specify a Reason for Revision (will print on PO)"
                    End If
                End If

                If rowPOTORDR1.Item("ORDR_NO") & "" <> "" Then
                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowPOTORDR1.Item("ORDR_NO"))
                    If rowSOTORDR1.Item("WHSE_CODE") <> Absx1.txtFor("WHSE_CODE").Text Then
                        EMsg &= vbCr & "Back-to-Back Order Destination Warehouse must be " & rowSOTORDR1.Item("WHSE_CODE")
                    End If
                End If

                fix_ICTSTYL1_packs = False

                If ASCMAIN1.CLIENT = "NYA" Then
                    ASCMAIN1.sql = "Select PO_ORDER_NO, PO_ORDER_LNO, STYLE_CODE" & vbCrLf _
                        & ", CARTON_PACK_QTY, CARTON_PACK_QTY CARTON_PACK_QTY_STYLE" & vbCrLf _
                        & ", INNER_PACK_QTY, INNER_PACK_QTY INNER_PACK_QTY_STYLE" & vbCrLf _
                        & " from POTORDR2 where ROWNUM < 1"
                    Dim tblPack As DataTable = ASCDATA1.GetDataTable
                    For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                        Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE")
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        Dim CARTON_PACK_QTY_po As Integer = Val(rowPOTORDR2.Item("CARTON_PACK_QTY") & "")
                        Dim CARTON_PACK_QTY_style As Integer = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
                        Dim INNER_PACK_QTY_po As Integer = Val(rowPOTORDR2.Item("INNER_PACK_QTY") & "")
                        Dim INNER_PACK_QTY_style As Integer = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")

                        If (CARTON_PACK_QTY_po <> 0 And CARTON_PACK_QTY_po <> CARTON_PACK_QTY_style And CARTON_PACK_QTY_style <> 0) _
                        Or (INNER_PACK_QTY_po <> INNER_PACK_QTY_style And INNER_PACK_QTY_style <> 0) Then
                            tblPack.Rows.Add(New Object() {rowPOTORDR2.Item("PO_ORDER_NO"), rowPOTORDR2.Item("PO_ORDER_LNO"),
                                                           STYLE_CODE, CARTON_PACK_QTY_po, CARTON_PACK_QTY_style, INNER_PACK_QTY_po, INNER_PACK_QTY_style})
                        End If
                    Next
                    If tblPack.Rows.Count <> 0 Then
                        Using frmmsg As New ASFMSGBF
                            frmmsg.Show_grd(tblPack, Me, "Styles on This PO with Carton Pack or Inner Pack Qty at odds with Style Table")

                            Dim msg_answer As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Do you want to correct the Style Master Qtys from this PO",
                                                                                          MsgBoxStyle.YesNoCancel,
                                                                                          "Option to Correct Carton and Inner Pack Qtys in Style Table")
                            If msg_answer = MsgBoxResult.Cancel Then
                                Exit Sub
                            ElseIf msg_answer = MsgBoxResult.Yes Then
                                fix_ICTSTYL1_packs = True
                            End If
                        End Using
                    End If
                End If

                If ASCMAIN1.CLIENT = "NYA" Then
                    Dim SEG4_CODE As String = TAC.TACMAIN1.Check_Division_MixMatch(Me, EMsg, "POTORDR2", "", Absx1.txtFor("WHSE_CODE").Text)

                    'Dim SALES_DIVISION_CODEs As New List(Of String)
                    'Dim SEG4_CODEs As New List(Of String)
                    'For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                    '    Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE")
                    '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    '    Dim SALES_DIVISION_CODE As String = rowICTSTYL1.Item("SALES_DIVISION_CODE") & ""
                    '    If SALES_DIVISION_CODE = "" Then
                    '        EMsg &= vbCr & "No Sales Divison Code set up for Style " & STYLE_CODE
                    '    Else
                    '        If Not SALES_DIVISION_CODEs.Contains(SALES_DIVISION_CODE) Then
                    '            Dim rowSOTSDIV1 As DataRow = LookUp("SOTSDIV1", SALES_DIVISION_CODE)
                    '            Dim SEG4_CODE As String = rowSOTSDIV1.Item("SEG4_CODE") & ""
                    '            If Not SEG4_CODEs.Contains(SEG4_CODE) Then
                    '                If SEG4_CODEs.Count > 0 Then
                    '                    EMsg &= vbCr & "Cannot have Styles with Sales Divisions from multiple Companies - " & STYLE_CODE & " (" & SALES_DIVISION_CODE & ")"
                    '                    Exit For
                    '                End If
                    '                SEG4_CODEs.Add(SEG4_CODE)
                    '            End If
                    '            'If SALES_DIVISION_CODEs.Count > 0 Then
                    '            'EMsg &= vbCr & "Cannot have multiple Sales Division Codes on a Single PO"
                    '            'End If
                    '            SALES_DIVISION_CODEs.Add(SALES_DIVISION_CODE)
                    '        End If
                    '    End If
                    'Next
                    If SEG4_CODE = "001" Then ' SEG4_CODEs.Count = 1 AndAlso SEG4_CODEs(0) = "001" Then
                        If chkBuyersCommision.Checked Then
                            If MsgBox("This PO appears to be for Canadian Product." & vbCrLf & vbCrLf & "Are you sure you want to Apply Buyers Commission?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If


                If ASCMAIN1.CLIENT = "VAN" Then
                    If txtCUST_CODE.Text <> "" Then
                        Dim CUST_CODE_STYLE As String = txtCUST_CODE.Text
                        If LookUp("ARTCUST1", CUST_CODE_STYLE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value for Customer Code"
                        Else
                            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR2"), "STYLE_CODE").Select()
                                Dim STYLE_CODE_PO As String = row.Item(0)
                                Dim rowSTYLE As DataRow = LookUp("ICTSTYL1", STYLE_CODE_PO)
                                If rowSTYLE.Item("CUST_CODE") & "" <> CUST_CODE_STYLE Then
                                    EMsg &= vbCr & $"Style Code {STYLE_CODE_PO} is not coded for Customer {CUST_CODE_STYLE}"
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If txtSTYLE_CODE_PFX.Visible And txtSTYLE_CODE_PFX.Text = "" Then
                        EMsg &= vbCr & "Please provide a Style Prefix"
                    End If
                    If txtCARTON_COUNT.Visible And Absx1.txtFor("PO_SPEC_ORDR_NO").Text.ToUpper.StartsWith("INITIAL") Then
                        Dim CARTON_COUNT As Integer = Val(txtCARTON_COUNT.Value & "")
                        If CARTON_COUNT <= 0 Then
                            EMsg &= vbCr & "Please provide a Carton Count (Initial Orders Only)"
                        Else
                            Dim OPNs As New Dictionary(Of String, Int32)
                            Dim ORDs As New Dictionary(Of String, Int32)
                            For Each ROW As DataRow In dst.Tables("POTORDR2").Select("")
                                Dim PO_ORDER_LNO As Integer = Val(ROW.Item("PO_ORDER_LNO") & "")
                                Dim PO_QTY_ORD As Integer = Val(ROW.Item("PO_QTY_ORD") & "")
                                Dim PO_QTY_OPN As Integer = Val(ROW.Item("PO_QTY_OPN") & "")

                                Dim STYLE_CODE As String = ROW.Item("STYLE_CODE") & ""
                                Dim COLOR_CODE As String = ROW.Item("COLOR_CODE") & ""

                                If OPNs.ContainsKey(STYLE_CODE & ":" & COLOR_CODE) Then
                                    OPNs(STYLE_CODE & ":" & COLOR_CODE) += PO_QTY_OPN
                                Else
                                    OPNs.Add(STYLE_CODE & ":" & COLOR_CODE, PO_QTY_OPN)
                                End If

                                If ORDs.ContainsKey(STYLE_CODE & ":" & COLOR_CODE) Then
                                    ORDs(STYLE_CODE & ":" & COLOR_CODE) += PO_QTY_ORD
                                Else
                                    ORDs.Add(STYLE_CODE & ":" & COLOR_CODE, PO_QTY_ORD)
                                End If
                            Next

                            'For Each SC As String In OPNs.Keys
                            '    Dim PO_QTY_OPN As Int32 = OPNs(SC)
                            '    If PO_QTY_OPN Mod CARTON_COUNT <> 0 Then
                            '        EMsg &= vbCr & $"PO Qty Open {PO_QTY_OPN} is not evenly divisible by Carton Count {CARTON_COUNT} For Style:Color {SC}"
                            '        Exit For
                            '    End If
                            'Next

                            For Each SC As String In ORDs.Keys
                                Dim PO_QTY_ORD As Int32 = ORDs(SC)
                                If PO_QTY_ORD Mod CARTON_COUNT <> 0 Then
                                    EMsg &= vbCr & $"PO Qty Ordered {PO_QTY_ORD} is not evenly divisible by Carton Count {CARTON_COUNT} For Style:Color {SC}"
                                    Exit For
                                Else
                                    Dim RATIO As Decimal = PO_QTY_ORD / CARTON_COUNT
                                    Dim PO_QTY_OPN As Int32 = OPNs(SC)
                                    If PO_QTY_OPN Mod RATIO <> 0 Then
                                        EMsg &= vbCr & $"PO Qty Open {PO_QTY_OPN} is not evenly divisible by Ratio {RATIO} For Style:Color {SC}"
                                        Exit For
                                    End If
                                End If
                            Next

                            'For Each ROW As DataRow In dst.Tables("POTORDR2").Select("")
                            '    Dim PO_ORDER_LNO As Integer = Val(ROW.Item("PO_ORDER_LNO") & "")
                            '    Dim PO_QTY_ORD As Integer = Val(ROW.Item("PO_QTY_ORD") & "")
                            '    Dim PO_QTY_OPN As Integer = Val(ROW.Item("PO_QTY_OPN") & "")
                            '    If PO_QTY_ORD Mod CARTON_COUNT <> 0 Then
                            '        EMsg &= vbCr & $"PO Qty Ordered {PO_QTY_ORD} is not evenly divisible by Carton Count {CARTON_COUNT} on Line {PO_ORDER_LNO}"
                            '        Exit For
                            '    End If
                            '    If PO_QTY_OPN Mod CARTON_COUNT <> 0 Then
                            '        EMsg &= vbCr & $"PO Qty Open {PO_QTY_OPN} is not evenly divisible by Carton Count {CARTON_COUNT} on Line {PO_ORDER_LNO}"
                            '        Exit For
                            '    End If
                            'Next
                        End If
                    End If
                End If
                If EMsg = "" Then
                    If dst.Tables("POTORDR2").Select("PO_QTY_SHP <> 0").Length <> 0 Then
                        Dim e As String = Change_Style_Color_in_Shipment(True)
                        If e <> "" Then
                            EMsg &= e
                        End If
                    End If
                End If

                If EMsg = "" Then
                    Dim PO_AMT_ORD As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
                    ' Dim PO_NINV_AMOUNT As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMOUNT)", "") & "")
                    Dim PO_TOTAL_AMT As Decimal = PO_AMT_ORD '+ PO_NINV_AMOUNT

                    If ROWs("POTPARM1").Item("PO_PARM_APPR_REQD") & "" = "1" Then

                        APPR_NOTES = ""
                        APPR_DECISION = ""
                        APPR_BY = ""
                        If chkReadyForApproval.Checked Then
                            If PO_TOTAL_AMT <= Val(rowPOTORDR1.Item("PO_APPR_AMOUNT") & "") Then
                                APPR_NOTES = "Previously Approved for " & Format(Val(rowPOTORDR1.Item("PO_APPR_AMOUNT") & ""), "$#,##0.00")
                                APPR_DECISION = "A"
                                APPR_BY = rowPOTORDR1.Item("PO_APPR_BY") & ""
                            Else
                                If PO_TOTAL_AMT <= Val(ROWs("POTPARM1").Item("PO_PARM_APPR_LIMIT") & "") Then
                                    Seek_Approval("")
                                End If
                            End If

                            If APPR_DECISION = "" Then
                                rowPOTORDR1.Item("PO_APPR_DATE") = DBNull.Value
                                rowPOTORDR1.Item("PO_APPR_BY") = DBNull.Value
                                rowPOTORDR1.Item("PO_APPR_AMOUNT") = DBNull.Value
                                ' rowPOTORDR1.Item("PO_APPR_NOTES") = DBNull.Value
                            End If
                        End If
                    End If

                End If

            Case "Cancel"
                If tabPOTORDR2.Tabs("Line Item Details").Visible Then
                    MsgBox("You must first Update or Cancel Line Item Details", MsgBoxStyle.OkOnly, "Line Item Details Page is Open")
                    Exit Sub
                End If


                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If


                'If tabPOTORDR2.Tabs("Line Item Details").Visible Then Setup_POLine(False)

                'If splPOLine.Visible Then
                '    Cancel_Changes()
                '    Exit Sub
                'End If

            Case "Done"
                If tabPOTORDR2.Tabs("Line Item Details").Visible Then Setup_POLine(False)

                If splPOLine.Visible Then
                    Cancel_Changes()
                    Exit Sub
                End If

            Case "Delete"
                If tabPOTORDR2.Tabs("Line Item Details").Visible Then
                    MsgBox("You must first Update or Cancel Line Item Details", MsgBoxStyle.OkOnly, "Line Item Details Page is Open")
                    Exit Sub
                End If

                If Check_Shipped(PO_ORDER_NO) = True Then
                    EMsg &= vbCr & "You May Not Delete an Order which has been Shipped"
                End If


                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    ASCMAIN1.sql = "Select * from POTPACK1 where PO_ORDER_NO = '" & PO_ORDER_NO & "' OR PO_ORDER_NO2 = '" & PO_ORDER_NO & "' or PO_ORDER_NO3 = '" & PO_ORDER_NO & "' or PO_ORDER_NO4 = '" & PO_ORDER_NO & "' Or PO_ORDER_NO5 = '" & PO_ORDER_NO & "' or PO_ORDER_NO6 = '" & PO_ORDER_NO & "' or PO_ORDER_NO7 = '" & PO_ORDER_NO & "' or PO_ORDER_NO8 = '" & PO_ORDER_NO & "'"
                    Dim tblPOTPACK1 As DataTable = ASCDATA1.GetDataTable()
                    If tblPOTPACK1.Rows.Count > 0 Then
                        EMsg &= vbCr & "You May Not Delete an Order which has Pack Lists Assigned To it"
                    End If
                    ' dgj new


                End If


                If EMsg = "" Then
                    If MsgBox("Do You Really Want To Delete this PO?",
                              MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "WARNING! - Answering 'Yes' will PERMANENTLY DELETE THIS PO!!") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If

            Case "Print"
                If splPOLine.Visible Then
                    EMsg = EMsg & vbCr & "You May Not Print While In Detail View"
                End If


            Case "Cancel PO"
                If tabPOTORDR2.Tabs("Line Item Details").Visible Then
                    MsgBox("You must first Update or Cancel Line Item Details", MsgBoxStyle.OkOnly, "Line Item Details Page is Open")
                    Exit Sub
                End If

                Dim lines_shipped As Integer = dst.Tables("POTORDR2").Select("PO_QTY_SHP <> 0").Length
                If lines_shipped = 0 Then
                    If MsgBox("Do You Still want to Cancel?", MsgBoxStyle.YesNo,
                              "No Shipments Recorded Yet; You May want to use the Delete Option") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

                Dim lines_open As Integer = dst.Tables("POTORDR2").Select("PO_QTY_OPN <> 0").Length
                If lines_open = 0 Then
                    EMsg &= vbCr & "Nothing Open to Cancel - If you Update, PO Status will be updated as Closed"
                End If

                If txtPO_REVISION_NOTE.Visible And txtPO_REVISION_NOTE.Text = "" Then
                    If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    Else
                        EMsg &= vbCr & "You Must Specify a Reason for Cancellation (will print on PO)"
                    End If
                End If


            Case "Add Line"
                If Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" = "" Or Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Then
                    MsgBox("Please provide default values for Ship-By and ETA Date above in the PO Header before entering PO Details",
                           MsgBoxStyle.OkOnly, "Cannot Enter New PO Details")
                    Exit Sub
                End If


            Case "Change Line"
                If grdPOTORDR2.ActiveRow Is Nothing OrElse Not grdPOTORDR2.ActiveRow.IsDataRow Then
                    EMsg &= vbCr & "You must Select a Row to Change the PO Data on the Line"
                End If

            Case "Update Line"

                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text)
                If rowICTSTYL1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Style Code"
                Else
                    If Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") <= 0 Then
                        EMsg &= vbCr & "Invalid Unit Pack Qty"
                    Else
                        If Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "") <> Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") Then
                            EMsg &= vbCr & "Sub-Unit Pack in Style Table does not Match Order"
                        End If
                    End If
                End If

                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", Absx1.txtFor("COLOR_CODE").Text)
                If rowICTCOLR1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Color Code"
                Else
                    Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() _
                                                        {Absx1.txtFor("STYLE_CODE").Text,
                                                         Absx1.txtFor("COLOR_CODE").Text})
                    If rowICTSTYC1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Style / Color Combination"
                    End If
                End If

                Dim rowICTCMTM1 As DataRow = LookUp("ICTCMTM1", Absx1.txtFor("CMT_NO").Text)
                If rowICTCMTM1 Is Nothing Then
                    EMsg &= vbCr & "Invalid CMT No"
                End If

                If OOBAL Then
                    EMsg &= vbCr & "Production is Out of Balance"
                End If

                If Absx1.dteFor("POTORDR2_LINE.PO_DATE_SHIP_BY").Value & "" = "" Or Absx1.dteFor("POTORDR2_LINE.PO_DATE_ETA").Value & "" = "" Then
                    EMsg = EMsg & vbCr & "Ship by and ETA Dates Can Not Be Blank"
                End If

                If EMsg <> "" Then
                    MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Update Line for Reasons Indicated")
                    Exit Sub
                End If

            Case "Approve", "Reject"
                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                    Exit Sub
                End If

                Seek_Approval(eItemKey)

                If APPR_DECISION = "" Then
                    'ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

            Case "Change Vendor"
                ASCMAIN1.sql = "Select * from POTSHIP3 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                Dim rowPOTSHIP3 As DataRow = ASCDATA1.GetDataRow
                If rowPOTSHIP3 IsNot Nothing Then
                    EMsg &= vbCr & "Cannot Change Vendor - PO has already been shipped or received"
                End If

                'Case "Load Styles"

                '    If Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" = "" Or Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Then

                '        WorkbookView1.GetLock()

                '        workbook = WorkbookView1.ActiveWorkbook
                '        worksheet = workbook.ActiveSheet

                '        For r As Integer = 0 To 10
                '            Dim S As String = worksheet.Cells(r, 0).Value & ""
                '            If S <> "" And S.StartsWith("Delivery") Then
                '                Dim D As String = worksheet.Cells(r, 1).Value & ""

                '                Dim dt0 As Date = "01/01/1900"
                '                Dim DT As Date = dt0.AddDays(Val(D) - 2)
                '                Absx1.dteFor("PO_DATE_SHIP_BY").Value = DT

                '            End If
                '        Next

                '        WorkbookView1.ReleaseLock()
                '    End If

                '    If Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" = "" Or Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Then
                '        EMsg &= vbCr & "Please provide default values for Ship-By and ETA Date above in the PO Header before entering PO Details"
                '    End If
            Case "Ship Confirmation"
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.InitialDirectory = "C:\"
                    openFileDialog1.Title = "Select a File to Upload Ship Confirmation Data"
                    openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        Dim FILENAME As String = openFileDialog1.FileName
                        EMsg = Load_Ship_Confirmation(FILENAME)

                        If EMsg = "" Then
                            If Not ASCMAIN1.Logical_Open("POTORDR1", "PO:" & "AT") Then Exit Sub

                            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDS4"), New String() {"PO_ORDER_NO"}).Select("")
                                Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                            Next
                        End If
                    Else
                        EMsg = "Nothing Selected"
                    End If
                End Using

            Case "Ship Confirmation Update"
                Dim LINES As Integer = dst.Tables("POTORDS4").Rows.Count
                Dim LINES2UPDATE As Integer = dst.Tables("POTORDS4").Select("SEL = '1'").Length
                If LINES2UPDATE = 0 Then
                    EMsg &= vbCr & "No Lines Selected for Update"
                Else
                    If MsgBox(String.Format("There are {0} PO Lines shown, and {1} have been selected for Update", LINES, LINES2UPDATE) _
                              & vbCrLf & vbCrLf & "OK to Proceed to Update ETD/ETA?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Clone"
                PO_ORDER_NO_clone = Absx1.txtFor("PO_ORDER_NO").Text
                rowPOTORDR1_clone = LookUp("POTORDR1", PO_ORDER_NO_clone)
                Click_Command("Done")
                Absx1.txtFor("VEND_CODE").Text = rowPOTORDR1_clone.Item("VEND_CODE")
                Absx1.txtFor("PO_REFERENCE").Text = rowPOTORDR1_clone.Item("PO_REFERENCE")
                Click_Command("New")
                Clone_PO()

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Confirm/Notes"
                confirm_notes_mode = True
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                If confirm_notes_mode Then
                    Update_Confirmed()
                Else
                    Update_Record()
                    If APPR_DECISION <> "" Then
                        Update_Approval("")
                    End If
                End If

                Mode_Settings(False)

                TAC.POCMAIN1.Check_Status(Me)

            Case "Cancel", "Done"

                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel PO"
                Cancel_Order()
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "email", "email to myself"
                Transmit_POs(PO_ORDER_NO, eItemKey = "email to myself")

            Case "Add Line"
                Add_Line()

            Case "Change Line", "View Line"
                Change_Line()

            Case "Refresh Line"
                Refresh_Line()

            Case "Duplicate Line"
                Duplicate_Line()

            Case "Cancel Changes", "Return to Summary"
                Cancel_Changes()

                'Case "Return to Summary"
                '    Done_with_Line()

            Case "Update Line"
                Update_Line()

            Case "Work Orders"
                Using F As New TAC.SOFWORK1(Me, "P", PO_ORDER_NO, (EntryMode = "V" Or InquiryMode),
                                            Absx1.txtFor("VEND_CODE").Text,
                                            Absx1.txtFor("PO_REFERENCE").Text,
                                            Absx1.dteFor("PO_DATE_SHIP_BY").Value,
                                             Absx1.dteFor("PO_DATE_ETA").Value,
                                            "Work Orders relating to PO " & PO_ORDER_NO)
                    F.ShowDialog()
                End Using

            Case "Approve", "Reject"
                Update_Approval(eItemKey)
                Mode_Settings(False)

            Case "Sales Order Entry", "Sales Order Inquiry"
                Dim ORDR_NO As String = rowPOTORDR1.Item("ORDR_NO")
                If ORDR_NO <> "" Then
                    Context_Launch("View", ORDR_NO, eItemKey, IIf(eItemKey = "Sales Order Inquiry", "SOFORDRI", "SOFORDR1"))
                End If

            Case "Integrity Check"
                Integrity_Check()

            Case "Change Vendor"
                Dim sql_where As String = "VEND_TYPE = 'S'"
                Dim VEND_CODE As String = Get_Code("VEND_CODE", , sql_where)
                If VEND_CODE <> "" Then
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
                    If MsgBox("OK to replace Vendor Code and Name with" _
                              & vbCrLf & rowAPTVEND1.Item("VEND_CODE") _
                              & ":" & rowAPTVEND1.Item("VEND_NAME") & "?",
                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Absx1.txtFor("VEND_CODE").Text = rowAPTVEND1.Item("VEND_CODE")
                        Absx1.txtFor("VEND_NAME").Text = rowAPTVEND1.Item("VEND_NAME")
                    End If
                End If

            Case "Ship Confirmation"

                grdPOTORDS1.Rows.ExpandAll(False)

                EntryMode = "S"
                Mode_Settings(True)
                Toggle_SC()

            Case "Ship Confirmation Update"

                Update_Ship_Confirmation()

                Trigger_Excel_Export(grdPOTORDS1)

                Mode_Settings(False)
                Toggle_SC()

            Case "Ship Confirmation Cancel"

                Mode_Settings(False)
                Toggle_SC()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("New").Settings.Enabled = not_iScreenMode
                        If ScreenMode Then
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                        Else
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                        End If
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode
                    If (EntryMode = "E" And ScreenMode) Then
                        .Items("Delete").Settings.Enabled = DefaultableBoolean.True
                        .Items("Cancel PO").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                        .Items("Cancel PO").Settings.Enabled = DefaultableBoolean.False
                    End If
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("email").Settings.Enabled = iScreenMode
                    .Items("email to myself").Settings.Enabled = iScreenMode

                    .Items("View").Visible = InquiryMode Or (EntryMode = "V" Or Not ScreenMode)
                    .Items("Done").Visible = Not confirm_notes_mode And (InquiryMode Or (EntryMode = "V" And ScreenMode))

                    .Items("Update").Visible = ((EntryMode = "N" Or EntryMode = "E") And ScreenMode) Or confirm_notes_mode
                    .Items("Cancel").Visible = ((EntryMode = "N" Or EntryMode = "E") And ScreenMode) Or confirm_notes_mode
                    .Items("Delete").Visible = ((EntryMode = "E") And ScreenMode) AndAlso (rowPOTORDR1.Item("ORDR_NO") & "" = "")
                    .Items("Cancel PO").Visible = ((EntryMode = "E") And ScreenMode) AndAlso (rowPOTORDR1.Item("ORDR_NO") & "" = "")
                    .Items("Approve").Settings.Enabled = iScreenMode
                    .Items("Reject").Settings.Enabled = iScreenMode
                    .Items("Approve").Visible = (Not InquiryMode And EntryMode = "V" And ScreenMode And optStatus.Value = "A") And ASCMAIN1.USER_SECURITY_CODEs.Contains("OM") And Not confirm_notes_mode
                    .Items("Reject").Visible = (Not InquiryMode And EntryMode = "V" And ScreenMode And optStatus.Value = "A") And ASCMAIN1.USER_SECURITY_CODEs.Contains("OM") And Not confirm_notes_mode
                    .Items("Work Orders").Text = "Work Orders" & IIf(dst.Tables("SOTWORK1").Rows.Count = 0, "", " (" & CStr(dst.Tables("SOTWORK1").Rows.Count) & ")")
                    .Items("Work Orders").Visible = ScreenMode And Not (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") And Not (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")
                    .Items("Work Orders").Visible = False
                    .Items("Print").Visible = InquiryMode Or (EntryMode = "V" And ScreenMode) And Not confirm_notes_mode

                    .Items("email").Visible = InquiryMode Or (EntryMode = "V" And ScreenMode) And Not confirm_notes_mode
                    .Items("email to myself").Visible = (ASCMAIN1.CLIENT = "NYA") And (InquiryMode Or (EntryMode = "V" And ScreenMode) And Not confirm_notes_mode)

                    .Items("Confirm/Notes").Visible = Not InquiryMode And Not confirm_notes_mode And ((EntryMode = "V" And ScreenMode) And Not (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN"))
                    .Items("Integrity Check").Visible = Not ScreenMode And ASCMAIN1.Running_in_VS

                    .Items("Change Vendor").Visible = ((EntryMode = "E") And ScreenMode) AndAlso (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")

                    .Items("New").Visible = Not InquiryMode And Not (ScreenMode And confirm_notes_mode) And (.Items("New").Settings.Enabled = DefaultableBoolean.True)
                    .Items("Edit").Visible = Not InquiryMode And (.Items("Edit").Settings.Enabled = DefaultableBoolean.True)
                    .Items("Clone").Visible = Not (ASCMAIN1.CLIENT = "VAN") And Not InquiryMode And (EntryMode = "V" And ScreenMode) And (.Items("Edit").Settings.Enabled = DefaultableBoolean.True)
                    .Items("Ship Confirmation").Visible = Not InquiryMode And Not ScreenMode And (ASCMAIN1.CLIENT = "VAN")

                End With

                .Groups("Totals").Visible = ScreenMode And (EntryMode <> "S")
                .Groups("Ship Confirmation").Visible = ScreenMode And (EntryMode = "S")

                '   .Groups("Status Filter").Visible = Not ScreenMode ' And InquiryMode
                .Groups("Line Item Commands").Visible = ScreenMode And (EntryMode <> "S") And Not No_Costs AndAlso (Absx1.optFor("FOB_CMT").Value = "I")
                ' .Groups("Cost Calculation").Visible = False ' ScreenMode AndAlso (Absx1.optFor("FOB_CMT").Value = "C")
                ' .Groups("Style Info").Visible = ScreenMode And (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
                Setup_tabPOTORDR2()
            End With

        End If

        If ScreenMode And EntryMode = "S" Then
            Exit Sub
        End If

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            If ScreenMode And optStatus.Value = "A" Then
                With grdPOTORDR2.DisplayLayout.Bands(0)
                    Dim P As Integer = .Columns("PO_DATE_SHIP_BY").Header.VisiblePosition
                    .Columns("PO_QTY_ORD").Header.SetVisiblePosition(P + 1, False)
                    .Columns("PO_COST").Header.SetVisiblePosition(P + 2, False)
                    .Columns("PO_LINE_NOTE_INT").Header.SetVisiblePosition(P + 3, False)
                    .Columns("PO_AMT_ORD").Header.SetVisiblePosition(P + 4, False)
                End With
            End If
        End If

        Dim blnShowYintak As Boolean = ScreenMode And ASCMAIN1.CLIENT = "VAN" And (Absx1.txtFor("VEND_CODE").Text = "YINTAK" Or Absx1.txtFor("VEND_CODE").Text = "CIVIC")
        lblSTYLE_CODE_PFX.Visible = blnShowYintak : txtSTYLE_CODE_PFX.Visible = blnShowYintak
        Set_Visible_CARTON_COUNT(blnShowYintak)

        tabDetails.Tabs("Style").Visible = ScreenMode And (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
        tabDetails.Tabs("Msg").Visible = ScreenMode And (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
        tabDetails.Tabs("Set").Visible = False

        tabPOTORDR2.Tabs("Imports").Visible = EntryMode <> "S" AndAlso ScreenMode AndAlso (ASCMAIN1.CLIENT = "VAN") AndAlso (rowPOTORDR1.Item("VEND_CODE") = "AT") And EntryMode <> "N"

        Absx1.chkFor("PO_WEB_VISIBLE").Visible = (ASCMAIN1.CLIENT = "NYA") And ScreenMode
        If (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") And Not ScreenMode Then
            If IsNothing(optStatus.ValueList.FindByDataValue("N")) Then
                Dim vli As New ValueListItem
                vli.DataValue = "N"
                vli.DisplayText = "Not On Portal"
                optStatus.ValueList.ValueListItems.Add(vli)
            End If
            If IsNothing(optStatus.ValueList.FindByDataValue("P")) Then
                Dim vli As New ValueListItem
                vli.DataValue = "P"
                vli.DisplayText = "On Portal"
                optStatus.ValueList.ValueListItems.Add(vli)
            End If
            optStatus.Height = optStatus.ValueList.ValueListItems.Count * 19
            UltraExplorerBarContainerControl3.Height = optStatus.Height + chkSplitByShipDate.Height
        End If

        Set_Read_Only(grpHeaderData, Not (EntryMode = "N" Or EntryMode = "E"))
        Set_Read_Only(grpPOTORDR2, Not (EntryMode = "N" Or EntryMode = "E"))
        Set_Read_Only(grpPOTORDR2X, Not (EntryMode = "N" Or EntryMode = "E"))

        Set_Read_Only_for_ctl(Absx1.numFor("YIELD_QTY"), Not (EntryMode = "N" Or EntryMode = "E"))
        cmdCostInc.Visible = (EntryMode = "N" Or EntryMode = "E")
        cmdPercent.Visible = (EntryMode = "N" Or EntryMode = "E")

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(ssdDZGRD, False)
        'Set_Read_Only_for_ctl(Absx1.txtFor("PO_MESSAGE"), ScreenMode And Not (EntryMode = "N" Or EntryMode = "E"))
        Set_Read_Only(grpPO_MESSAGE, ScreenMode And Not (EntryMode = "N" Or EntryMode = "E"))
        Set_Read_Only(grpLabel, ScreenMode And Not (EntryMode = "N" Or EntryMode = "E"))
        'Set_Read_Only_for_ctl(Absx1.optFor("PO_STATUS"), True)
        Set_Read_Only_for_ctl(Absx1.txtFor("PO_REFERENCE"), ScreenMode And Not (EntryMode = "N" Or EntryMode = "E"))
        Set_Read_Only_for_ctl(Absx1.txtFor("PO_SPEC_ORDR_NO"), ScreenMode And Not (EntryMode = "N" Or EntryMode = "E"))

        lblStatus.Visible = ScreenMode

        If EntryMode = "N" Then
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                lblPO_DATE_SHIP_BY.Text = "Ship by"
            ElseIf ASCMAIN1.CLIENT = "NYA" Then
            Else
                lblPO_DATE_SHIP_BY.Text = "Req Ship by"
            End If
        Else
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ElseIf ASCMAIN1.CLIENT = "NYA" Then
            Else
                lblPO_DATE_SHIP_BY.Text = "Req Ship by"
            End If
        End If


        'grdPOTORDRX.Visible = Not tf
        tabPO.Visible = Not tf
        tabPO.Tabs("Integrity Check").Visible = False
        splPOTSHIP1.Visible = tf
        Setup_tabPO()

        Absx1.optFor("FOB_CMT").Visible = tf Or (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")




        If ScreenMode Then

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdPOTORDR2, grdPOTORDR3, grdPOTORDR4, grdPOTORDR6, grdPOTORDR7, grdPOTORDR8, grdPOTORDRN}
                With grd.DisplayLayout.Override
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Or (confirm_notes_mode And grd.Name = "grdPOTORDRN") Then
                        If grd.Name = "grdPOTORDR6" Or grd.Name = "grdPOTORDRN" Or (grd.Name = "grdPOTORDR2" And Absx1.optFor("FOB_CMT").Value <> "I") Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        End If

                        .AllowUpdate = DefaultableBoolean.True
                        .AllowDelete = DefaultableBoolean.True

                        If grd.Name = "grdPOTORDR2" And rowPOTORDR1.Item("ORDR_NO") & "" <> "" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                        End If
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    End If
                End With
            Next

            If subUPCSupport Then
                With grdICTXLSPS.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End With
            End If
            If confirm_notes_mode Then
                grdPOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                ' DON'T KNOW WHY i DON'T HAVE TO WORRY ABOUT FIXING THIS VALUE SO THAT WHEN A PO IS CALLED UP TO BE EDITED THE EDITABLE FIELDS ARE EDITABLE
                For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTORDR2.DisplayLayout.Bands(0).Columns
                    If gcol.Key = "PO_CONF_NO" Or gcol.Key = "PO_CONF_DATE" Or gcol.Key = "PO_DATE_SHIP_BY" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                Next
            End If

            LOAD_CMT_TYPE()

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                splPOTORDR2.Panel2Collapsed = (EntryMode = "N" Or EntryMode = "E")
            End If

            If EntryMode = "N" Or EntryMode = "E" Then
                With grdPOTORDR2.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() _
                        {"PO_ORIG_DATE_SHIP_BY", "PO_ORIG_DATE_ETA", "LAST_OPER_SHIP_BY", "LAST_DATE_SHIP_BY",
                         "SHIP_COST_CHANGE_USER", "SHIP_COST_CHANGE_DATE",
                         "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER",
                         "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN", "LINE_FINISHED"}
                        If EntryMode = "E" And COLUMN_NAME = "PO_QTY_OPN" Then
                        Else
                            .Columns(COLUMN_NAME).Hidden = True
                        End If
                    Next
                End With
            End If

            Dim showPortalVisibilityControls As Boolean = False
            Dim WHSE_CODE_PO As String = rowPOTORDR1.Item("WHSE_CODE") & ""
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE_PO)

            If rowICTWHSE1 IsNot Nothing Then
                Dim WHSE_POS_WEB_VISIBLE As String = rowICTWHSE1.Item("WHSE_POS_WEB_VISIBLE") & ""
                showPortalVisibilityControls = (WHSE_POS_WEB_VISIBLE = "1")
            End If

            If ASCMAIN1.DBS_COMPANY = "RGI" AndAlso showPortalVisibilityControls Then
                tabPOTORDR2.Tabs("Labels").Visible = True
                Set_Read_Only(grpURL, Not (EntryMode = "N" Or EntryMode = "E"))
                If EntryMode = "N" Or EntryMode = "E" Then
                    cmdBuildURL.Visible = (rowPOTORDR1.Item("PO_LABEL_URL") & "" = "")
                Else
                    cmdBuildURL.Visible = False
                End If
                cmdPRINTLABEL.Visible = (rowPOTORDR1.Item("PO_LABEL_URL") & "" <> "")
            Else
                tabPOTORDR2.Tabs("Labels").Visible = False
            End If


        Else
            Clear_Record()
        End If



        rowPOTXLSF1 = Nothing

        If ScreenMode Then
            If EntryMode = "S" Then
            Else


                Toggle_Customer_Style_Fields(False)

                ssdDZGRD.Value = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_PO_UM"))

                lblPO_REVISION_NOTE.Visible = (EntryMode = "E" And Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "") > 0) Or txtPO_REVISION_NOTE.Text <> ""
                txtPO_REVISION_NOTE.Visible = (EntryMode = "E" And Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "") > 0) Or txtPO_REVISION_NOTE.Text <> ""

                If rowPOTORDR1.Item("PO_DATE_CANCELLED") & "" <> "" Then
                    lblPO_REVISION_NOTE.Text = "Cancelled on " & Format(rowPOTORDR1.Item("PO_DATE_CANCELLED"), "MM/dd/yyyy")
                Else
                    lblPO_REVISION_NOTE.Text = "Reason for Revision"
                End If


                With grdPOTORDR2.DisplayLayout.Bands(0)
                    .Columns("CONFIRMED").Hidden = Not (EntryMode = "E")
                    .Columns("COST_COMPLETE").Hidden = Not (EntryMode = "E")
                    .Columns("COST_COMPLETE").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("PO_CONF_NO").Hidden = Not (EntryMode = "E" Or EntryMode = "V")
                    .Columns("PO_CONF_DATE").Hidden = Not (EntryMode = "E" Or EntryMode = "V")
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        .Columns("PO_CONF_NO").Hidden = Not (EntryMode = "E")
                        .Columns("PO_CONF_DATE").Hidden = Not (EntryMode = "E")
                    End If
                End With


                With grdPOTORDR2.DisplayLayout.Bands(0)
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        .Columns("PO_COST_VCOST").Header.Caption = IIf(Absx1.optFor("FOB_CMT").Value = "F", "Cost", "CMT")
                        .Columns("PO_COST_VCOST_DZ").Header.Caption = IIf(Absx1.optFor("FOB_CMT").Value = "F", "Cost/Dz", "CMT/Dz")
                    Else
                        .Columns("PO_COST_VCOST").Header.Caption = "Cost"
                        .Columns("PO_COST_VCOST_DZ").Header.Caption = "Cost/Dz"
                        .Columns("COST_COMPLETE").Hidden = True
                    End If
                End With

                Set_grdPOTORDR2_cols_Visibility()

                lblBack2Back.Visible = (rowPOTORDR1.Item("ORDR_NO") & "" <> "")
                tabPOTORDR2.Tabs("Back-to-Back").Visible = (rowPOTORDR1.Item("ORDR_NO") & "" <> "")
            End If

        Else
            lblBack2Back.Visible = False
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            tabPOTORDR2.Tabs("AP").Visible = False
        Else
            tabPOTORDR2.Tabs("Cartonization").Visible = False
            tabPOTORDR2.Tabs("AP").Visible = ScreenMode And (InquiryMode Or EntryMode = "V")

            For Each COLUMN_NAME As String In New String() {"PO_COST_QUOTA", "PO_COST_QUOTA_UN", "PO_COST_BUFFER", "DFQUOTA", "LINE_FINISHED"}
                grdPOTORDR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
            Next

            Show_Audit_Fields(False)
        End If


    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTORDR1", "POTORDR2", "POTORDR3", "POTORDR4", "POTORDRR", "POTORDR6", "POTORDR7", "POTORDR8", "POTSHIPX",
                 "TATEVNT1", "POTORDXR", "POTORDRT", "SOTWORK1", "SOTWORK2", "ASTATTA2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        If ASCMAIN1.CLIENT = "RGI" Then
            For Each TABLE_NAME As String In New String() {"POTORDRL"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        End If

        If subUPCSupport Then
            dst.Tables("ICTXLSPS").Rows.Clear()
        End If

        EnforceConstraints(True)

        confirm_notes_mode = False

        dtePO_CONF_SHIP_BY.Value = Now.Date
        dtePO_CONF_DATE.Value = Now.Date
        txtPO_CONF_NO.Text = Format(Now.Date, "yyyyMMdd")

        APPR_NOTES = ""
        APPR_DECISION = ""
        APPR_BY = ""

        fix_ICTSTYL1_packs = False

        lblPO_ORDER_NO.Text = "PO No"
        tabPO.SelectedTab = tabPO.Tabs("Open POs")

        If ASCMAIN1.CLIENT = "VAN" Then
            If WorkbookView1.ActiveWorkbook IsNot Nothing Then
                Try
                    WorkbookView1.GetLock()
                    WorkbookView1.ActiveWorkbook.Close()
                    WorkbookView1.ReleaseLock()
                Catch ex As Exception

                End Try

            End If
        End If

        Load_POTORDRX()
        Absx1.txtFor("PO_ORDER_NO").Focus()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor

        If EntryMode = "N" Then
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                PO_ORDER_NO = ASCMAIN1.Next_Control_No("PO_ORDER_NO")
            Else
                PO_ORDER_NO = ASCMAIN1.Next_Control_No("POTORDR1.PO_ORDER_NO")
            End If
        Else
            PO_ORDER_NO = Absx1.txtFor("PO_ORDER_NO").Text
        End If

        EnforceConstraints(False)

        grdPOTORDR2.DisplayLayout.Bands(0).Columns("LINE_CLOSED").Hidden = (EntryMode = "N" Or EntryMode = "V")

        'Dim rowPOTORDR1 As DataRow
        If EntryMode = "N" Then
            rowPOTORDR1 = dst.Tables("POTORDR1").NewRow
            rowPOTORDR1.Item("PO_ORDER_NO") = PO_ORDER_NO

            If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                ' NO DEFAULT WAREHOUSE
            Else
                rowPOTORDR1.Item("WHSE_CODE") = ROWs("POTPARM1").Item("PO_PARM_DEF_WHSE_CODE")
            End If

            rowPOTORDR1.Item("VEND_CODE") = HFs("VEND_CODE")

            Dim VEND_NAME As String = HFs("VEND_NAME")
            If VEND_NAME.Length > 35 And ASCMAIN1.CLIENT = "VAN" Then
                VEND_NAME = Mid(VEND_NAME, 1, 35)
            End If
            rowPOTORDR1.Item("VEND_NAME") = VEND_NAME

            rowPOTORDR1.Item("PO_REFERENCE") = HFs("PO_REFERENCE")
            rowPOTORDR1.Item("PO_SPEC_ORDR_NO") = HFs("PO_SPEC_ORDR_NO")
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                rowPOTORDR1.Item("FOB_CMT") = HFs("FOB_CMT")
            Else
                If HFs("FOB_CMT") <> "" Then
                    rowPOTORDR1.Item("FOB_CMT") = HFs("FOB_CMT")
                Else
                    rowPOTORDR1.Item("FOB_CMT") = "F"
                End If
            End If

            Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
            rowPOTORDR1.Item("TERM_CODE") = rowAPTVEND1.Item("TERM_CODE")

            rowPOTORDR1.Item("PORT_CODE_ORIG") = rowAPTVEND1.Item("PORT_CODE")
            ' rowPOTORDR1.Item("PORT_CODE_DEST") = rowICTWHSE1.Item("PORT_CODE")

            rowPOTORDR1.Item("COST_CODE") = rowAPTVEND1.Item("COST_CODE")
            rowPOTORDR1.Item("PO_FOB_DESC") = rowAPTVEND1.Item("VEND_PURCH_FOB_DESC")
            rowPOTORDR1.Item("PO_SHIP_VIA") = rowAPTVEND1.Item("VEND_PURCH_SHIP_VIA")

            rowPOTORDR1.Item("PO_DATE_ORDERED") = DATETIME_STAMP.Date
            rowPOTORDR1.Item("PO_STATUS") = "O"
            rowPOTORDR1.Item("PO_XMIT_IND") = "0"
            rowPOTORDR1.Item("PO_WEB_VISIBLE") = "1"

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                rowPOTORDR1.Item("PO_COMM_PAYABLE_TO_BRKR") = "1"
                rowPOTORDR1.Item("PO_COMM_CHGBACK_TO_SUPP") = "1"
                rowPOTORDR1.Item("PO_COMM_PCT") = 2 ' BE CAREFUL - IF YOU PUT THIS INTO PARAMETER FILE IT MIGHT BE ADDED TO PO_COST
            End If

            dst.Tables("POTORDR1").Rows.Add(rowPOTORDR1)

        Else
            ASCMAIN1.Progress("Now Loading PO")
            rowPOTORDR1 = Fill_Record("POTORDR1", PO_ORDER_NO)

        End If
        Calculate_ETD_to_ETA()

        Fill_Records("POTORDR2", PO_ORDER_NO)
        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
            If Val(rowPOTORDR2.Item("PO_QTY_OPN") & "") = 0 Then
                If Val(rowPOTORDR2.Item("PO_QTY_ORD") & "") - Val(rowPOTORDR2.Item("PO_QTY_SHP") & "") > 0 Then
                    rowPOTORDR2.Item("LINE_CLOSED") = "1"
                End If
            End If
            Dim SUB_UNIT_PACK_QTY As Int16 = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")
            If SUB_UNIT_PACK_QTY = 0 Then SUB_UNIT_PACK_QTY = 1
            rowPOTORDR2.Item("PO_QTY_ORD_DZ") = Val(rowPOTORDR2.Item("PO_QTY_ORD") & "") / (12 / SUB_UNIT_PACK_QTY)
        Next

        ASCMAIN1.sql = "Select Distinct POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & " from POTSHIP3,POTSHIP1" & vbCrLf _
            & " where POTSHIP3.PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
            & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP1.COST_COMPLETE = '1'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {row.Item("PO_ORDER_NO"), row.Item("PO_ORDER_LNO")})
            rowPOTORDR2.Item("COST_COMPLETE") = "1"
        Next

        grdPOTORDR2.Text = "Details for PO " & Absx1.txtFor("PO_ORDER_NO").Text & ", Reference " & Absx1.txtFor("PO_REFERENCE").Text

        Dim PO_STATUS As String = rowPOTORDR1.Item("PO_STATUS")
        ' Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
        lblStatus.Text = IIf(PO_STATUS = "O", "Open", "Closed")
        If PO_STATUS = "O" Then
            If rowPOTORDR1.Item("PO_XMIT_IND") & "" = "1" Then
                lblStatus.Text &= ", Transmitted"
            ElseIf rowPOTORDR1.Item("PO_APPR_BY") & "" <> "" Then
                lblStatus.Text &= ", Pending Transmit"
            ElseIf rowPOTORDR1.Item("PO_APPR_PENDING") & "" = "1" Then
                lblStatus.Text &= ", Pending Approval"
            Else
                lblStatus.Text &= ", Work in Process"
            End If
        Else
            If rowPOTORDR1.Item("PO_DATE_CANCELLED") & "" <> "" Then
                If dst.Tables("POTORDR2").Select("PO_QTY_SHP <> 0 or PO_QTY_REC <> 0").Length <> 0 Then
                    lblStatus.Text &= ", " & Format(rowPOTORDR1.Item("PO_DATE_CANCELLED"), "MM/dd/yy")
                Else
                    lblStatus.Text &= ", Cancelled " & Format(rowPOTORDR1.Item("PO_DATE_CANCELLED"), "MM/dd/yy")
                End If
            End If
        End If
        ' lblRevision.Text = IIf(PO_HDR_CTR_REV = 0, "Original", "Rev #" & CStr(PO_HDR_CTR_REV))

        ASCMAIN1.sql = "Select * from ASTATTA2 where TABLE_NAME = 'POTORDR6' and COLUMN_NAME = 'PO_MESSAGE_ATTACHMENT' and CODE_VALUE = '" & PO_ORDER_NO & "'"
        Fill_Records("ASTATTA2", "", True, ASCMAIN1.sql)

        Fill_Records("POTORDR3", New Object() {PO_ORDER_NO})
        Fill_Records("POTORDR4", PO_ORDER_NO)

        Fill_Records("POTORDR6", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDR6, "PO_ORDER_MLNO")

        Fill_Records("POTORDR7", PO_ORDER_NO)
        Fill_Records("POTORDR8", PO_ORDER_NO)

        Fill_Records("SOTWORK1", PO_ORDER_NO)
        Fill_Records("SOTWORK2", PO_ORDER_NO)

        Fill_Records("TATEVNT1", PO_ORDER_NO)
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
        Fill_Records("POTORDXR", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDXR, "INIT_DATE".ToLower)

        If ASCMAIN1.CLIENT = "RGI" Then
            Fill_Records("POTORDRL", PO_ORDER_NO)
        End If

        If subUPCSupport Then
            Fill_Records("ICTXLSPS", PO_ORDER_NO)
        End If

        Fill_Records("POTLCST2", PO_ORDER_NO)
        ' Sort_grdColumns(grdPOTLCST2, "INIT_DATE".ToLower)

        rowAPTVEND1 = LookUp("APTVEND1", rowPOTORDR1.Item("VEND_CODE"))

        Fill_Records("POTORDRN", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDRN, "PO_ORDER_CLNO")
        Fill_Records("POTORDRH", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDRH, "PO_HDR_CTR_REV")

        If rowPOTORDR1.Item("ORDR_NO") & "" <> "" Then
            Fill_Records("SOTORDR1", rowPOTORDR1.Item("ORDR_NO"))

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU, SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.STYLE_RETAIL" _
                & " from SOTORDR2,POTORDR2" _
                & " where SOTORDR2.ORDR_NO = POTORDR2.ORDR_NO" _
                & "   and SOTORDR2.ORDR_LNO = POTORDR2.ORDR_LNO" _
                & "   and POTORDR2.PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim ORDR_NO As String = rowSOTORDR2.Item("ORDR_NO")
                Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO))(0)
                For Each COLUMN_NAME As String In New String() {"CUST_UPC", "CUST_SKU", "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "STYLE_RETAIL"}
                    rowPOTORDR2.Item(COLUMN_NAME) = rowSOTORDR2.Item(COLUMN_NAME)
                Next
            Next
        End If

        dst.Tables("POTORDRR").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR2"), New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
            Create_POTORDRR(row.Item("STYLE_CODE"), row.Item("COLOR_CODE"))
        Next
        Sort_grdColumns(grdPOTORDRR, "STYLE_CODE,COLOR_CODE")

        Sort_grdColumns(grdPOTORDR2, "PO_ORDER_LNO")

        dst.Tables("POTORDRT").Rows.Clear()


        If ASCMAIN1.CLIENT = "VAN" And EntryMode <> "N" Then
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                'stop
            Else
                Fill_Records("POTORDRA", Absx1.txtFor("PO_REFERENCE").Text)
                Sort_grdColumns(grdPOTORDRA, "VAN_REF".ToLower)
            End If

        End If

        EnforceConstraints(True)


        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select()
            With rowPOTORDR2
                If .Item("PO_STATUS") = "C" And Val(.Item("PO_QTY_ORD") & "") > Val(.Item("PO_QTY_SHP") & "") Then .Item("LINE_CLOSED") = "1"
                Dim SUB_UNIT_PACK_QTY As Int16 = Val(.Item("SUB_UNIT_PACK_QTY") & "")
                If SUB_UNIT_PACK_QTY = 0 Then SUB_UNIT_PACK_QTY = 1
                .Item("PO_COST_OTHER_UN") = Val(.Item("PO_COST_OTHER") & "") / (12 / SUB_UNIT_PACK_QTY)
                .Item("PO_COST_QUOTA_UN") = Val(.Item("PO_COST_QUOTA") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))
            End With
        Next

        If grdPOTORDR2.ActiveRow IsNot Nothing Then
            dtePO_CONF_SHIP_BY.Value = grdPOTORDR2.ActiveRow.Cells("PO_DATE_SHIP_BY").Value
        End If

        If InquiryMode Or EntryMode = "V" Then
            Fill_Records("POTSHIPX", PO_ORDER_NO)
        End If
        Set_POTSHIPX()

        If Not InquiryMode Then
            If Absx1.optFor("FOB_CMT").Value = "I" Then
                With grdPOTORDR2.DisplayLayout.Bands(0)
                    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    Next
                    .Columns("PO_DATE_SHIP_BY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("PO_DATE_ETA").CellActivation = UltraWinGrid.Activation.AllowEdit

                End With
                grdPOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdPOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Else
                grdPOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdPOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If

            Set_Read_Only_for_ctl(Absx1.txtFor("PO_REFERENCE"), Check_Shipped(PO_ORDER_NO))
        End If

        Setup_POLine(False)
        UltraExplorerBar1.Groups("Line Item Commands").Items("Duplicate Line").Settings.Enabled = DefaultableBoolean.True
        Display_Totals()

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
        If EntryMode = "N" Then
            If PO_ORDER_NO_clone <> "" Then
                lblPO_ORDER_NO.Text = "PO No - Clone"
                Record_Event("INIT", "PO Entry Started")
                Record_Event("CLONE", "PO Cloned from " & PO_ORDER_NO_clone, , PO_ORDER_NO_clone)
            Else
                lblPO_ORDER_NO.Text = "PO No - New"
                Record_Event("INIT", "PO Entry Started")
            End If

        ElseIf EntryMode = "E" Then
            If rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1" Then
                rowPOTORDR1.Item("PO_PRINTED_IND") = "0"
                rowPOTORDR1.Item("PO_XMIT_IND") = "0"
                PO_HDR_CTR_REV += 1
                rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV
                rowPOTORDR1.Item("PO_REVISION_NOTE") = ""
            End If
            lblPO_ORDER_NO.Text = "PO No - Rev#" & CStr(PO_HDR_CTR_REV)
            Record_Event("LAST", "PO Edit Started")
        Else
            lblPO_ORDER_NO.Text = "PO No - Rev#" & CStr(PO_HDR_CTR_REV)
        End If

        grdPOTORDRR.Text = "Style/Color Unit Recap for PO" & PO_ORDER_NO
        grdPOTORDR7.Text = "Carton Types for PO " & PO_ORDER_NO
        Setup_grdPOTORDR8()

        Sort_grdColumns(grdPOTORDR3, "COLOR_NO")
        Sort_grdColumns(grdPOTORDR4, "FABRIC_NO")
        Setup_Style_Image()
        PPK_CODE_ctr = 0

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Cancel_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()
        Dependent_Updates(-1, PO_ORDER_NO, True)
        Record_Event("PO-CXL", "PO Cancelled", True)

        ASCMAIN1.sql = "Update POTORDR1 Set PO_REVISION_NOTE = :PARM1, LAST_DATE = :PARM2, LAST_OPER = :PARM3 where PO_ORDER_NO = :PARM4"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VDVV", New Object() {txtPO_REVISION_NOTE.Text, DATETIME_STAMP, ASCMAIN1.USER_ID, PO_ORDER_NO})

        CommitTrans("Order " & PO_ORDER_NO & " has been Cancelled")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()
        Dependent_Updates(-1, PO_ORDER_NO)
        For Each TABLE_NAME In New String() {"POTORDR1", "POTORDR2", "POTORDR3", "POTORDR4"}
            ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'")
        Next
        CommitTrans("PO " & PO_ORDER_NO & " has been Deleted")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Dependent_Updates(S As Integer, PO_ORDER_NO As String, Optional cancel_po As Boolean = False)
        TAC.POCMAIN1.Dependent_Updates(S, PO_ORDER_NO, cancel_po)
    End Sub

    Sub Update_Record()
        BeginTrans()

        If dst.Tables("POTORDR2").Select("PO_QTY_SHP <> 0").Length <> 0 Then
            Change_Style_Color_in_Shipment(False)
        End If

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("(LINE_CLOSED = '1' or ISNULL(PO_QTY_OPN,0) = 0) AND ISNULL(PO_STATUS,'?') <> 'C'")
            rowPOTORDR2.Item("PO_STATUS") = "C"
        Next

        Dim OPEN_LINES As Integer = Val(dst.Tables("POTORDR2").Compute("COUNT(PO_ORDER_LNO)", "PO_STATUS = 'O'") & "")
        rowPOTORDR1.Item("PO_STATUS") = IIf(OPEN_LINES = 0, "C", "O")

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("PO_ORIG_DATE_SHIP_BY IS NULL")
            rowPOTORDR2.Item("PO_ORIG_DATE_SHIP_BY") = rowPOTORDR2.Item("PO_DATE_SHIP_BY")
        Next
        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("PO_ORIG_DATE_ETA IS NULL")
            rowPOTORDR2.Item("PO_ORIG_DATE_ETA") = rowPOTORDR2.Item("PO_DATE_ETA")
        Next

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select _
            ("PO_ORDER_NO = '" & PO_ORDER_NO & "'", "", DataViewRowState.ModifiedCurrent)

            If Val(rowPOTORDR2.Item("PO_COST_VCOST") & "") <> Val(rowPOTORDR2.Item("PO_COST_VCOST", DataRowVersion.Original) & "") _
            Or Val(rowPOTORDR2.Item("PO_COST_MATLS") & "") <> Val(rowPOTORDR2.Item("PO_COST_MATLS", DataRowVersion.Original) & "") _
            Or Val(rowPOTORDR2.Item("PO_COST_OTHER") & "") <> Val(rowPOTORDR2.Item("PO_COST_OTHER", DataRowVersion.Original) & "") _
            Or Val(rowPOTORDR2.Item("PO_COST_COMM") & "") <> Val(rowPOTORDR2.Item("PO_COST_COMM", DataRowVersion.Original) & "") _
            Or Val(rowPOTORDR2.Item("PO_COST_BUFFER") & "") <> Val(rowPOTORDR2.Item("PO_COST_BUFFER", DataRowVersion.Original) & "") _
            Then
                ASCMAIN1.sql = "Update POTSHIP3 Set COST_CHANGED = '1' where PO_ORDER_NO = '" & PO_ORDER_NO & "'" _
                & " and PO_ORDER_LNO = " & rowPOTORDR2.Item("PO_ORDER_LNO")
                ASCDATA1.ExecuteSQL()
            End If
        Next

        If EntryMode <> "N" Then
            Dependent_Updates(-1, PO_ORDER_NO)

            ' CHANGE THIS TO TAC.POCMAIN1.Check_Changed_Fields after we know this works from POFORDRA

            If Check_Changed_Fields() Then ' this function records audit trail, so it is important to call even if we do nothing with the results returned
                'Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
                'rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV + 1
            End If
        End If

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            rowPOTORDR1.Item("PO_COMM_CHGBACK_TO_SUPP") = rowPOTORDR1.Item("PO_COMM_PAYABLE_TO_BRKR")
        End If

        Dim sqlx As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        INIT_LAST("POTORDR1", False, sqlx, True)
        ' INIT_LAST("POTORDR2", False, sqlx, True)

        Record_Event("UPDT", "PO Updated", True)

        Update_Record_TDA("POTORDR1", sqlx)
        Update_Record_TDA("POTORDR2", sqlx)
        Update_Record_TDA("POTORDR3", sqlx)
        Update_Record_TDA("POTORDR4", sqlx)
        Update_Record_TDA("POTORDR6", sqlx)
        Update_Record_TDA("POTORDR7", sqlx)
        Update_Record_TDA("POTORDR8", sqlx)
        Update_Record_TDA("POTORDRN", sqlx)

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Delete from POTORDR5 where PO_ORDER_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PO_ORDER_NO)

            If txtCARTON_COUNT.Visible And Absx1.txtFor("PO_SPEC_ORDR_NO").Text.ToUpper.StartsWith("INITIAL") Then
                Dim CARTON_COUNT As Integer = Val(txtCARTON_COUNT.Value & "")
                ASCMAIN1.sql = "Insert into POTORDR5" & vbCrLf _
                    & $"Select PO_ORDER_NO, STYLE_CODE, COLOR_CODE, SUM (PO_QTY_ORD) / {CARTON_COUNT} PO_QTY_ORD from POTORDR2" & vbCrLf _
                    & " where PO_ORDER_NO = :PARM1" & vbCrLf _
                    & " group by PO_ORDER_NO, STYLE_CODE, COLOR_CODE"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PO_ORDER_NO)
            End If
        End If

        Update_Record_TDA("ASTATTA2", "TABLE_NAME = 'POTORDR6' and COLUMN_NAME = 'PO_MESSAGE_ATTACHMENT' and CODE_VALUE = '" & PO_ORDER_NO & "'")

        Update_Record_TDA("POTORDXR")

        Update_Record_TDA("SOTWORK1")
        Update_Record_TDA("SOTWORK2")

        Dependent_Updates(1, PO_ORDER_NO)

        ' Fix Carton Pack and Inner Pack

        dst.Tables("ASTAUDT1").Rows.Clear()
        dst.Tables("ICTSTYL1").Rows.Clear()
        Dim styles_changed As Boolean = False

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
            Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE")
            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
            If rowICTSTYL1 Is Nothing Then rowICTSTYL1 = Fill_Record("ICTSTYL1", STYLE_CODE, False, False)
            Dim CARTON_PACK_QTY_po As Integer = Val(rowPOTORDR2.Item("CARTON_PACK_QTY") & "")
            Dim CARTON_PACK_QTY_style As Integer = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
            Dim INNER_PACK_QTY_po As Integer = Val(rowPOTORDR2.Item("INNER_PACK_QTY") & "")
            Dim INNER_PACK_QTY_style As Integer = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")

            ' the data is suspect if carton pack qty is not 0 - maybe also suspect if inner pack qty is not 0, but less concerned about inner
            If CARTON_PACK_QTY_po <> 0 Then ' And INNER_PACK_QTY_po <> 0 Then

                Dim style_changed As Boolean = False
                If (CARTON_PACK_QTY_po <> 0 And CARTON_PACK_QTY_po <> CARTON_PACK_QTY_style And (CARTON_PACK_QTY_style = 0 Or fix_ICTSTYL1_packs)) Then
                    rowICTSTYL1.Item("CARTON_PACK_QTY") = CARTON_PACK_QTY_po
                    rowICTSTYL1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    rowICTSTYL1.Item("LAST_DATE") = DATETIME_STAMP
                    styles_changed = True
                    Dim rowASTAUDT1 As DataRow = Write_Audit_Special("ICTSTYL1", STYLE_CODE, "CARTON_PACK_QTY", "E", rowICTSTYL1)
                    rowASTAUDT1.Item("NOTES") = "PO Fix CARTON_PACK_QTY"
                    rowASTAUDT1.Item("KEY_VALUE2") = rowPOTORDR2.Item("PO_ORDER_NO")
                    rowASTAUDT1.Item("KEY_LNO") = rowPOTORDR2.Item("PO_ORDER_LNO")
                End If
                If (INNER_PACK_QTY_po <> 0 And INNER_PACK_QTY_po <> INNER_PACK_QTY_style And (INNER_PACK_QTY_style = 0 Or fix_ICTSTYL1_packs)) Then
                    rowICTSTYL1.Item("INNER_PACK_QTY") = INNER_PACK_QTY_po
                    rowICTSTYL1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    rowICTSTYL1.Item("LAST_DATE") = DATETIME_STAMP
                    styles_changed = True
                    Dim rowASTAUDT1 As DataRow = Write_Audit_Special("ICTSTYL1", STYLE_CODE, "INNER_PACK_QTY", "E", rowICTSTYL1)
                    rowASTAUDT1.Item("NOTES") = "PO Fix INNER_PACK_QTY"
                    rowASTAUDT1.Item("KEY_VALUE2") = rowPOTORDR2.Item("PO_ORDER_NO")
                    rowASTAUDT1.Item("KEY_LNO") = rowPOTORDR2.Item("PO_ORDER_LNO")
                End If

            End If

        Next

        If styles_changed Then
            Update_Record_TDA("ICTSTYL1")
            Update_Record_TDA("ASTAUDT1")
        End If


        'Dim sqlAudit As String = ""

        'ASCMAIN1.sql = "" _
        '    & "Begin Declare Cursor C1 is " _
        '    & " Select POTORDR2.* from POTORDR2,ICTSTYL1 " & vbCrLf _
        '    & "  where ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE " & vbCrLf _
        '    & "    and NVL(POTORDR2.CARTON_PACK_QTY,0) <> 0 " & vbCrLf _
        '    & "    and NVL(ICTSTYL1.CARTON_PACK_QTY,0) = 0 " & vbCrLf _
        '    & "    and POTORDR2.PO_ORDER_NO = '" & PO_ORDER_NO & "';" & vbCrLf _
        '    & " Begin For R1 in C1 Loop" & vbCrLf _
        '    & "  Update ICTSTYL1 Set CARTON_PACK_QTY = R1.CARTON_PACK_QTY, LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
        '    & "   where STYLE_CODE = R1.STYLE_CODE" & IIf(fix_ICTSTYL1_packs, " and NVL(CARTON_PACK_QTY,0) <> NVL(R1.CARTON_PACK_QTY,0)", " and NVL(CARTON_PACK_QTY,0) = 0") & ";" & vbCrLf _
        '    & "  If SQL%NOTFOUND then " & Replace(sqlAudit, "CARTON_PACK_QTY", "CARTON_PACK_QTY") & " Endif;" & vbCrLf _
        '    & " End Loop; End; End;"
        'ASCDATA1.ExecuteSQL()
        'ASCMAIN1.sql = "" _
        '    & "Begin Declare Cursor C1 is " _
        '    & " Select POTORDR2.* from POTORDR2,ICTSTYL1 " & vbCrLf _
        '    & "  where ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE " & vbCrLf _
        '    & "    and NVL(POTORDR2.INNER_PACK_QTY,0) <> 0 " & vbCrLf _
        '    & "    and NVL(ICTSTYL1.INNER_PACK_QTY,0) = 0 " & vbCrLf _
        '    & "    and POTORDR2.PO_ORDER_NO = '" & PO_ORDER_NO & "';" & vbCrLf _
        '    & " Begin For R1 in C1 Loop" & vbCrLf _
        '    & "  Update ICTSTYL1 Set INNER_PACK_QTY = R1.INNER_PACK_QTY, LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
        '    & "   where STYLE_CODE = R1.STYLE_CODE" & IIf(fix_ICTSTYL1_packs, " and NVL(INNER_PACK_QTY,0) <> NVL(R1.INNER_PACK_QTY,0)", " and NVL(INNER_PACK_QTY,0) = 0") & ";" & vbCrLf _
        '    & " End Loop; End; End;"
        'ASCDATA1.ExecuteSQL()

        If rowPOTORDR1.Item("FOB_CMT") & "" = "I" Then
            ASCMAIN1.sql = "Update ICTSTYL1 set CMT_NO = " _
            & "(Select MIN(CMT_NO) from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "' and STYLE_CODE = ICTSTYL1.STYLE_CODE)" _
            & " where STYLE_CODE in (Select DISTINCT STYLE_CODE from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "')" _
            & "   and CMT_NO is Null"
            ASCDATA1.ExecuteSQL()
        End If


        If rowPOTXLSF1 IsNot Nothing Then
            Dim PO_XLS_NO As String = rowPOTXLSF1.Item("PO_XLS_NO")
            Dim rowPOTXLSF0 As DataRow = dst.Tables("POTXLSF0").Rows.Find(PO_XLS_NO)
            dst.Tables("POTXLSF1").AcceptChanges()
            rowPOTXLSF1.Item("XLS_ORDER_STATUS") = "1"
            Update_Record_TDA("POTXLSF1")
            rowPOTXLSF1.Delete()
            dst.Tables("POTXLSF1").AcceptChanges()

            If grdPOTXLSF1.Rows.Count = 0 Then
                dst.Tables("POTXLSF0").AcceptChanges()
                rowPOTXLSF0.Item("PO_XLS_STATUS") = "1"
                Dim FILENAME As String = rowPOTXLSF0.Item("FILENAME")
                Dim FILENAME_NEW As String = Replace(FILENAME, "XLS_New\", "XLS_Archived\" & PO_XLS_NO & "_")
                My.Computer.FileSystem.CopyFile(FILENAME, FILENAME_NEW)
                My.Computer.FileSystem.DeleteFile(FILENAME)
                Update_Record_TDA("POTXLSF0")
                rowPOTXLSF0.Delete()
                dst.Tables("POTXLSF0").AcceptChanges()

            End If
        End If

        If blnAutomatic Then
            CommitTrans()
            blnAutomatic = False
        Else
            CommitTrans("Update Complete")
        End If


        If ASCMAIN1.Running_in_VS Or ASCMAIN1.USER_ID = "rgomez" Then
            Dim sqlIC = TAC.POCMAIN1.Get_sql_Integrity_Check
            Dim tbl As DataTable = ASCDATA1.GetDataTable(sqlIC)
            If tbl.Rows.Count <> 0 Then
                MsgBox("Please email a Screenshot to Walter, and describe your work on PO " & PO_ORDER_NO, MsgBoxStyle.OkOnly, "PO Shipments are Out of Balance")
            End If
        End If

    End Sub

    Function Write_Audit_Special(TABLE_NAME As String, KEY_VALUE As String, COLUMN_NAME As String, FM_MODE As String, row As DataRow) As DataRow
        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
        rowASTAUDT1.Item("TABLE_NAME") = TABLE_NAME
        rowASTAUDT1.Item("KEY_VALUE") = KEY_VALUE
        rowASTAUDT1.Item("COLUMN_NAME") = COLUMN_NAME
        rowASTAUDT1.Item("USER_ID") = ASCMAIN1.USER_ID
        rowASTAUDT1.Item("INIT_DATE") = DATETIME_STAMP

        Dim OLD_VALUE As String = row.Item(COLUMN_NAME, DataRowVersion.Original) & ""
        If Len(OLD_VALUE & "") > 255 Then
            OLD_VALUE = Mid(OLD_VALUE, 1, 255)
        End If
        rowASTAUDT1.Item("OLD_VALUE") = OLD_VALUE

        If FM_MODE <> "D" Then
            Dim NEW_VALUE As String = row.Item(COLUMN_NAME)
            If Len(NEW_VALUE & "") > 255 Then
                NEW_VALUE = Mid(NEW_VALUE, 1, 255)
            End If
            rowASTAUDT1.Item("NEW_VALUE") = NEW_VALUE
        End If

        rowASTAUDT1.Item("FM_MODE") = FM_MODE
        rowASTAUDT1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
        rowASTAUDT1.Item("SELECTION_NO") = SELECTION_NO
        rowASTAUDT1.Item("XNO") = XNO
        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)

        Return rowASTAUDT1
    End Function

    Sub Update_Confirmed()
        BeginTrans()

        Dim sqlx As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        INIT_LAST("POTORDR1", False, sqlx, True)
        ' INIT_LAST("POTORDR2", False, sqlx, True)

        Record_Event("UPDC", "PO Updated Conf", True)

        Update_Record_TDA("POTORDR1", sqlx)
        Update_Record_TDA("POTORDR2", sqlx)
        Update_Record_TDA("POTORDRN", sqlx)

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

            Case "View", "Edit"
                Absx1.txtFor("PO_ORDER_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTORDR1"
            E.COLUMN_NAME = "PO_ORDER_NO"
            E.CODE_VALUE = Absx1.txtFor("PO_ORDER_NO").Text
            E.DESC_VALUE = Absx1.txtFor("VEND_CODE").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "POTORDR1"
        E.TABLE_KEY_CAPTION = "PO"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PO_ORDER_NO").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "N")
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

            Case "PO_ORDER_NO"
                If InquiryMode Then
                    If optStatus.Value = "O" Then
                        sql_where = " AND PO_ORDER_NO in (Select DISTINCT PO_ORDER_NO from POTORDR2 where PO_STATUS = 'O') "
                    End If
                Else
                    sql_where = " AND PO_ORDER_NO in (Select DISTINCT PO_ORDER_NO from POTORDR2 where PO_STATUS = 'O') "
                End If
                If Absx1.txtFor("VEND_CODE").Text <> "" Then
                    sql_where &= " AND VEND_CODE = '" & Replace(Absx1.txtFor("VEND_CODE").Text, "'", "") & "'"
                End If
                If Absx1.txtFor("PO_REFERENCE").Text <> "" Then
                    ' HOW DO WE PROTECT AGAINST SINGLE QUOTES?
                    sql_where &= " AND PO_REFERENCE like '" & Replace(Absx1.txtFor("PO_REFERENCE").Text, "'", "") & "%'"
                End If
                If Absx1.txtFor("PO_SPEC_ORDR_NO").Text <> "" Then
                    sql_where &= " AND PO_SPEC_ORDR_NO like '" & Replace(Absx1.txtFor("PO_SPEC_ORDR_NO").Text, "'", "") & "%'"
                End If

            Case "COLOR_CODE"
                If COLOR_CODEs.Count = 0 Then
                    sql_where = "COLOR_CODE = ''"
                    Exit Sub
                Else
                    sql_where = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                End If

            Case "CMT_NO"
                If STYLE_CODE = "" Then
                    sql_where = "COLOR_CODE = ''"
                    Exit Sub
                Else

                End If

            Case "VEND_CODE"
                sql_where = "VEND_TYPE = 'S'"

            Case "PO_REFERENCE"
                sql_where = "DISTINCT PO_REFERENCE from POTORDR1"

            Case "PO_SPEC_ORDR_NO"
                sql_where = "DISTINCT PO_SPEC_ORDR_NO from POTORDR1"

        End Select

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTORDRX, "SSSBBBS", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "Sales Order Entry", "Sales Order Inquiry", "Show Ship/ETA from PO Detail", "Show Min Comm")
        Load_Popup_Menu(grdPOTORDR2, "SSBBBBBBBBBSBB", "Show Sub-Details", "Show Audit Fields", "Update All ETA Dates", "Update All Ship Dates",
                        "Update ETA Dates of Selected Rows", "Update Ship Dates of Selected Rows", "Style Status Inquiry",
                        "Style Multi-Color", "Copy Line", "Split Line", "Cost Calculator", "Show UPC/SKU", "Copy DF Quota", "Paste DF Quota")
        Load_Popup_Menu(grdPOTORDRS, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry",
                        "Update Style Master (+ POs) with Changes made to Case Packs",
                        "Update POs with Case Pack for Selected Styles",
                        "Update Style Master with Case Pack with Changes made to Case Packs",
                        "Update Style Master with Case Pack for Selected Styles",
                        "Refresh Styles with Open POs")
        Load_Popup_Menu(grdPOTORDRR, "BB",
                     "Create a carton type containing all selected Style/Colors",
                     "Create an individual carton type for All Style/Colors",
                     "Create an individual carton type for each selected Style/Color")
        Load_Popup_Menu(grdPOTORDRH, "B", "Show PO")
        Load_Popup_Menu(grdTATEVNT1, "B", "Show email")
        Load_Popup_Menu(grdPOTORDR8, "S", "Show All Carton Details")
        Load_Popup_Menu(grdPOTSHIPX, "B", "PO Shipment Inquiry")
        Load_Popup_Menu(grdPOTORDXR, "SS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdPOTLCST2, "B", "Voucher Inquiry")
        Load_Popup_Menu(txtPO_MESSAGE, "B", "PO Messages")
        Load_Popup_Menu(grdPOTXLSF1, "B", "Style Master File")
        If ASCMAIN1.CLIENT = "VAN" Then
            Load_Popup_Menu(grdPOTORDRA, "SSSBBBS", "Show Filter", "Show GroupBox", "Show Pins", "PO Import Inquiry", "AT PDF", "FOB Cost Sheet")
        End If
        Load_Popup_Menu(grdPOTORDS1, "SSSBSS", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "Horizontal View", "Fill Screen")
        Load_Popup_Menu(grdPOTORDS4, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All")
        If subUPCSupport Then
            Load_Popup_Menu(grdICTXLSPS, "B", "Style Master File")
        End If
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.Key = "txtPO_MESSAGE" Then
            If EntryMode = "N" Or EntryMode = "E" Then
            Else
                e.Cancel = True
            End If
            Exit Sub
        End If

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
            Case "grdPOTORDRR"
                If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool

        Select Case grd.Name


            Case "grdPOTORDRX"
                tlb_sbt = DirectCast(tlb_pop.Tools("Show Ship/ETA from PO Detail"), UltraWinToolbars.StateButtonTool)

                If chkSplitByShipDate.Checked Then
                    tlb_sbt.SharedProps.Visible = False
                Else
                    tlb_sbt.SharedProps.Visible = True
                    tlb_sbt.Tag = "X"
                    tlb_sbt.Checked = Not grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY_MIN").Hidden
                    tlb_sbt.Tag = ""
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Show Min Comm"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = ASCMAIN1.CLIENT = "VAN"

            Case "grdPOTORDR2"
                tlb_sbt = DirectCast(tlb_pop.Tools("Show Sub-Details"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Tag = "X"
                tlb_sbt.Checked = Not splPOTORDR2.Panel2Collapsed
                tlb_sbt.Tag = ""

                tlb_sbt = DirectCast(tlb_pop.Tools("Show UPC/SKU"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = True ' (Absx1.optFor("ORDR_SOURCE").Value = "K")
                tlb_sbt.Tag = "X"
                tlb_sbt.Checked = Not grdPOTORDR2.DisplayLayout.Bands(0).Columns("CUST_UPC").Hidden
                tlb_sbt.Tag = ""

                tlb_sbt = DirectCast(tlb_pop.Tools("Show Audit Fields"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Tag = "X"
                tlb_sbt.Checked = Not grdPOTORDR2.DisplayLayout.Bands(0).Columns("LAST_OPER_SHIP_BY").Hidden
                tlb_sbt.Tag = ""

                tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And (rowPOTORDR1.Item("ORDR_NO") & "" = "")
                tlb_btn = DirectCast(tlb_pop.Tools("Copy Line"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And (rowPOTORDR1.Item("ORDR_NO") & "" = "")
                tlb_btn = DirectCast(tlb_pop.Tools("Split Line"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And (rowPOTORDR1.Item("ORDR_NO") & "" = "")
                tlb_btn = DirectCast(tlb_pop.Tools("Cost Calculator"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And (Absx1.optFor("FOB_CMT").Value = "C") And (ssdDZGRD.Value = 12)

                tlb_btn = DirectCast(tlb_pop.Tools("Copy DF Quota"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And ASCMAIN1.CLIENT = "VAN"
                tlb_btn = DirectCast(tlb_pop.Tools("Paste DF Quota"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N") And ASCMAIN1.CLIENT = "VAN" And (tlb_btn.Tag & "").ToString.StartsWith(PO_ORDER_NO)

                'Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
                'tlb_btn = DirectCast(tlb_pop.Tools("Update AT Commissions"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (EntryMode = "E" And ASCMAIN1.CLIENT = "VAN" And VEND_CODE = "AT")

            Case "grdTATEVNT1"
                tlb_btn = DirectCast(tlb_pop.Tools("Show email"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("EVENT_TYPE").Value = "PO-XMIT" Or grd.ActiveRow.Cells("EVENT_TYPE").Value = "PO-XPED"))

            Case "grdPOTORDRS"
                tlb_btn = DirectCast(tlb_pop.Tools("Update Style Master (+ POs) with Changes made to Case Packs"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden And Not (MENU_ITEM_OBJECT = "POFORDRI") And (txtPO_SHIPMENT_NO.Text = "")
                tlb_btn = DirectCast(tlb_pop.Tools("Update POs with Case Pack for Selected Styles"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grdPOTORDRS.DisplayLayout.Bands(0).Columns("POS_UPDATED").Hidden And Not (MENU_ITEM_OBJECT = "POFORDRI") And (txtPO_SHIPMENT_NO.Text = "")

                tlb_btn = DirectCast(tlb_pop.Tools("Update Style Master with Case Pack with Changes made to Case Packs"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden And Not (MENU_ITEM_OBJECT = "POFORDRI") And (txtPO_SHIPMENT_NO.Text <> "")
                tlb_btn = DirectCast(tlb_pop.Tools("Update Style Master with Case Pack for Selected Styles"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grdPOTORDRS.DisplayLayout.Bands(0).Columns("POS_UPDATED").Hidden And Not (MENU_ITEM_OBJECT = "POFORDRI") And (txtPO_SHIPMENT_NO.Text <> "")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked

                Case "grdPOTXLSF1"
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsAddRow Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    If grd.ActiveRow.Band.Index <> 1 Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value & ""
                        If STYLE_CODE = "" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                Case "grdPOTORDR2"
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsAddRow Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    Dim shipDate As Date
                    Dim etaDate As Date
                    Dim blnShowDateChange As Boolean = False
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") And grd.ActiveCell IsNot Nothing _
                        AndAlso grd.ActiveRow IsNot Nothing _
                        AndAlso (grd.ActiveCell.Column.Key = "PO_DATE_ETA" Or grd.ActiveCell.Column.Key = "PO_DATE_SHIP_BY") Then
                        shipDate = grd.ActiveRow.Cells("PO_DATE_SHIP_BY").Value
                        etaDate = shipDate.AddDays(ETD_to_ETA)
                        blnShowDateChange = True
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Update All ETA Dates"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = blnShowDateChange AndAlso grd.ActiveCell.Column.Key = "PO_DATE_ETA"
                    If tlb_btn.SharedProps.Visible Then tlb_btn.SharedProps.Caption = "Update All ETA Dates to " & Format(etaDate, "MM/dd/yy")

                    tlb_btn = DirectCast(tlb_pop.Tools("Update ETA Dates of Selected Rows"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = blnShowDateChange AndAlso grd.ActiveCell.Column.Key = "PO_DATE_ETA"
                    If tlb_btn.SharedProps.Visible Then tlb_btn.SharedProps.Caption = "Update ETA Dates of Selected Rows to " & Format(etaDate, "MM/dd/yy")

                    tlb_btn = DirectCast(tlb_pop.Tools("Update All Ship Dates"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = blnShowDateChange AndAlso grd.ActiveCell.Column.Key = "PO_DATE_SHIP_BY"
                    If tlb_btn.SharedProps.Visible Then tlb_btn.SharedProps.Caption = "Update All Ship Dates to " & Format(shipDate, "MM/dd/yy")

                    tlb_btn = DirectCast(tlb_pop.Tools("Update Ship Dates of Selected Rows"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = blnShowDateChange AndAlso grd.ActiveCell.Column.Key = "PO_DATE_SHIP_BY"
                    If tlb_btn.SharedProps.Visible Then tlb_btn.SharedProps.Caption = "Update Ship Dates of Selected Rows to " & Format(shipDate, "MM/dd/yy")
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.Key <> "PO Messages" Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Paste DF Quota"
                tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)

                If tlb_btn.Tag & "" = "" Then
                    MsgBox("You must first indicate a line ot Copy before Pasting", MsgBoxStyle.OkOnly, "Cannot Paste")
                    Exit Sub
                End If

                Dim copy_from_row As DataRow = Nothing
                Dim copy_from As String = tlb_btn.Tag & ""
                If copy_from.StartsWith(PO_ORDER_NO) And copy_from.Contains(vbTab) Then
                    Dim line As Integer = Val(Split(copy_from, vbTab)(1))
                    copy_from_row = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, line})
                End If

                If copy_from_row Is Nothing Then
                    MsgBox("Cannot Determine the Copy From Row (" & copy_from & ")", MsgBoxStyle.OkOnly, "Cannot Paste")
                    Exit Sub
                End If

                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("No Lines Selected to Paste To", MsgBoxStyle.OkOnly,
                           "Cannot Paste DF Quota from PO " & PO_ORDER_NO & " Line " & copy_from_row.Item("PO_ORDER_LNO"))
                    Exit Sub
                End If

                If MsgBox("OK to Paste the DF Quota Amount and Selection from Line " _
                          & copy_from_row.Item("PO_ORDER_LNO") & " to " _
                          & CStr(grd.Selected.Rows.Count) & " Selected Rows?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        grow.Cells("DFQUOTA").Value = copy_from_row.Item("DFQUOTA")
                        grow.Cells("PO_COST_QUOTA").Value = copy_from_row.Item("PO_COST_QUOTA")

                        grow.Update()
                    Next
                    tlb_btn.Tag = ""
                End If



            Case "Horizontal View"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    grdPOTORDS1.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal
                Else
                    grdPOTORDS1.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy
                End If

            Case "Fill Screen"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    splPOTORDR1.Panel2Collapsed = True
                    grdPOTORDS1.Rows.ExpandAll(True)
                Else
                    splPOTORDR1.Panel2Collapsed = False
                    grdPOTORDS1.Rows.CollapseAll(True)
                    grdPOTORDS1.Rows.ExpandAll(False)
                End If

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "De-Select All", "0", "1")
                        grow.Update()
                    End If
                Next


            Case "Show Sub-Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then
                    Exit Sub
                End If
                splPOTORDR2.Panel2Collapsed = Not tlb_sbt.Checked


            Case "Show Audit Fields"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then
                    Exit Sub
                End If

                Show_Audit_Fields(tlb_sbt.Checked)

            Case "PO Messages"

                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PO_MESSAGE_CODE")
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                        Dim PO_MESSAGE_CODE As String = ASCMAIN1.CodeSelector.SelectedCode
                        Dim rowPOTMESS1 As DataRow = LookUp("POTMESS1", PO_MESSAGE_CODE)
                        If rowPOTMESS1 IsNot Nothing Then
                            Dim PO_MESSAGE_DESC As String = rowPOTMESS1.Item("PO_MESSAGE_DESC") & ""
                            txtPO_MESSAGE.Text &= IIf(txtPO_MESSAGE.Text = "", "", vbCrLf) & PO_MESSAGE_DESC
                        End If
                    End If
                End If

                Exit Sub

            Case "Show Ship/ETA from PO Detail"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    Toggle_Ship_ETA_from_PO_Detail(tlb_sbt.Checked)
                End If

            Case "Show UPC/SKU"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    Toggle_Customer_Style_Fields(tlb_sbt.Checked)
                End If

            Case "Style Multi-Color"
                Using F As New TAC.ICFSTYCX
                    F.STYLE_CODE = ""
                    F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
                    F.ShowDialog()
                    If F.STYLE_CODE <> "" Then
                        Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
                    End If
                End Using

            Case "Refresh Styles with Open POs"
                Refresh_Styles_Open_POs()

            Case "Update Style Master with Case Pack for Selected Styles"

                If grdPOTORDRS.Selected.Rows.Count = 0 Then
                    MsgBox("No Rows Selected", vbOKOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                Dim STYLE_CODEs As New List(Of String)
                Dim msg As String = "All Styles have been Updated Successfully"

                For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDRS.Selected.Rows
                    Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                    Dim rowPOTORDRS As DataRow = dst.Tables("POTORDRS").Rows.Find(STYLE_CODE)
                    Dim CARTON_PACK_QTY As Int64 = Val(rowPOTORDRS.Item("CARTON_PACK_QTY") & "")

                    If Not ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE, , False, False, 1) Then
                        rowPOTORDRS.Item("STYLE_ACTION") = "Failed to Lock Style"
                        msg = "Some Styles may NOT have been Updated - Check Grid for Results"
                    Else
                        ASCMAIN1.sql = "Update ICTSTYL1 Set CARTON_PACK_QTY = :PARM1 where STYLE_CODE = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NV", New Object() {Val(rowPOTORDRS.Item("CARTON_PACK_QTY") & ""), STYLE_CODE})
                        STYLE_CODEs.Add(STYLE_CODE)
                        rowPOTORDRS.Item("STYLE_ACTION") = "Style Updated"
                    End If

                Next

                grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden = False
                ASCMAIN1.MultiTask_Release(, , 1)
                MsgBox(msg & vbCrLf & vbCrLf & "Export Grid and Save for record of Updates",
                       vbOKOnly, "Process Complete")

            Case "Update Style Master (+ POs) with Changes made to Case Packs", "Update Style Master with Case Pack with Changes made to Case Packs"
                Dim rows() As DataRow = dst.Tables("POTORDRS").Select _
                    ("ISNULL(CARTON_PACK_QTY,0) <> 0", "STYLE_CODE", DataViewRowState.ModifiedCurrent)

                If rows.Length = 0 Then
                    MsgBox("No Changes Made", vbOKOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                Dim STYLE_CODEs As New List(Of String)
                Dim msg As String = "All Styles have been Updated Successfully"
                For Each rowPOTORDRS As DataRow In rows
                    Dim STYLE_CODE As String = rowPOTORDRS.Item("STYLE_CODE")
                    If Not ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE, , False, False, 1) Then
                        rowPOTORDRS.Item("STYLE_ACTION") = "Failed to Lock Style"
                        msg = "Some Styles may NOT have been Updated - Check Grid for Results"
                    Else
                        ASCMAIN1.sql = "Update ICTSTYL1 Set CARTON_PACK_QTY = :PARM1 where STYLE_CODE = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NV", New Object() {Val(rowPOTORDRS.Item("CARTON_PACK_QTY") & ""), STYLE_CODE})
                        STYLE_CODEs.Add(STYLE_CODE)
                        rowPOTORDRS.Item("STYLE_ACTION") = "Style Updated"
                    End If
                Next

                grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden = False
                ASCMAIN1.MultiTask_Release(, , 1)
                MsgBox(msg & vbCrLf & vbCrLf & "Export Grid and Save for record of Updates",
                       vbOKOnly, "Process Complete")

                If txtPO_SHIPMENT_NO.Text = "" Then
                    If MsgBox("Would you like the POs for these Styles Updated as well?",
                           MsgBoxStyle.YesNo,
                           "Option to Take Case Pack Updates into Open POs") = MsgBoxResult.Yes Then
                        grdPOTORDRS.Selected.Rows.Clear()
                        For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDRS.Rows
                            If STYLE_CODEs.Contains(grow.Cells("STYLE_CODE").Value) Then
                                grow.Selected = True
                            End If
                        Next
                        Update_POs_with_Case_Pack()
                    End If
                End If



            Case "Update POs with Case Pack for Selected Styles"
                Update_POs_with_Case_Pack()

            'Case "Update AT Commissions"
            '    Dim AT_PO_COST_COMM As Decimal = 2.0 'New AT Commissions as of 1/1/21
            '    Dim iResult As MsgBoxResult
            '    Dim iTitle As String = "Update Commissions to " & Format(AT_PO_COST_COMM, "###,##0.00")
            '    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            '    iMSG.AppendLine("This Will Update All Open Lines On This Order.")
            '    iMSG.AppendLine("Is That What You Want?")
            '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            '    If iResult = MsgBoxResult.Yes Then
            '        For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDR2.Rows
            '            grow.Cells.Item("PO_COST_COMM").Value = AT_PO_COST_COMM
            '        Next
            '    End If

            Case "Show Min Comm"
                Me.Cursor = Cursors.WaitCursor
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select()
                    Dim PO_ORDER_NO As String = rowPOTORDRX.Item("PO_ORDER_NO").ToString & String.Empty
                    'If PO_ORDER_NO = "140319" Then Stop
                    SQLS.Length = 0
                    SQLS.AppendLine("SELECT MIN(PO_COST_COMM) PO_COST_COMM_MIN")
                    SQLS.AppendLine("FROM POTORDR2")
                    SQLS.AppendLine(String.Format("WHERE PO_ORDER_NO = '{0}'", PO_ORDER_NO))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim PO_COST_COMM_MIN As Double = Val(ASCDATA1.GetDataValue)
                    rowPOTORDRX.Item("PO_COST_COMM_MIN") = Format(PO_COST_COMM_MIN, "###,##0.00")
                Next
                grdPOTORDRX.DisplayLayout.Bands(0).Columns.Item("PO_COST_COMM_MIN").Hidden = False
                Me.Cursor = Cursors.Default
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "PO Import Inquiry"
                Dim VAN_REF As String = grd.ActiveRow.Cells("VAN_REF").Text
                Context_Launch("View", VAN_REF, e.Tool.Key, "POFORDRA", "F", "POAT")

            Case "Copy DF Quota"

                tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                tlb_btn = DirectCast(tlb_btn.OwningMenu.Tools("Paste DF Quota"), UltraWinToolbars.ButtonTool)

                tlb_btn.Tag = PO_ORDER_NO & vbTab & CStr(grd.ActiveRow.Cells("PO_ORDER_LNO").Value)

                MsgBox("DF Quota for Line " & CStr(grd.ActiveRow.Cells("PO_ORDER_LNO").Value) & " has been Copied", MsgBoxStyle.OkOnly, "Success")

            Case "AT PDF"

                Dim POKEY As String = grd.ActiveRow.Cells("POKEY").Text

                Dim PONO As String = grd.ActiveRow.Cells("PONO").Text
                Dim SENDNO As String = grd.ActiveRow.Cells("SENDNO").Text

                Dim FILENAME As String = PONO & "-" & SENDNO & ".PDF"
                Dim FULLPATH As String = PO_PARM_PO_IMG_DIR & "\" & POKEY & "\" & FILENAME
                If Not My.Computer.FileSystem.FileExists(FULLPATH) Then
                    FILENAME = PONO & ".PDF"
                    FULLPATH = PO_PARM_PO_IMG_DIR & "\" & POKEY & "\" & FILENAME
                    If Not My.Computer.FileSystem.FileExists(FULLPATH) Then
                        MsgBox("No PDF Found", MsgBoxStyle.OkOnly, "Cannot find AT PDF")
                        Exit Sub
                    End If
                End If

                Dim file As String = ASCMAIN1.Folders("Temp") & "\" & FILENAME
                Try
                    If My.Computer.FileSystem.FileExists(file) Then
                        My.Computer.FileSystem.DeleteFile(file)
                    End If
                    My.Computer.FileSystem.CopyFile(PO_PARM_PO_IMG_DIR & "\" & POKEY & "\" & FILENAME, ASCMAIN1.Folders("Temp") & "\" & FILENAME)
                    'Show_Document(PO_PARM_PO_IMG_DIR & "\" & POKEY & "\" & FILENAME)
                    Show_Document(ASCMAIN1.Folders("Temp") & "\" & FILENAME)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Show PDF")
                End Try

            Case "FOB Cost Sheet"

            Case "PO Inquiry"
                If grd.Name = "grdPOTORDS1" AndAlso grd.ActiveRow.Band.Index <> 3 Then Exit Sub

                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Text
                Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")

            Case "Sales Order Entry", "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                If ORDR_NO <> "" Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, IIf(e.Tool.Key = "Sales Order Inquiry", "SOFORDRI", "SOFORDR1"))
                End If

            Case "Update All ETA Dates"
                If grd.ActiveRow.IsDataRow Then
                    If grd.ActiveRow.DataChanged Then grd.ActiveRow.Update()
                    Dim etaDate As Date = grd.ActiveRow.Cells("PO_DATE_ETA").Value
                    For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select
                        rowPOTORDR2.Item("PO_DATE_ETA") = etaDate
                    Next
                End If

            Case "Update All Ship Dates"
                If grd.ActiveRow.IsDataRow Then
                    If grd.ActiveRow.DataChanged Then grd.ActiveRow.Update()
                    Dim shipDate As Date = grd.ActiveRow.Cells("PO_DATE_SHIP_BY").Value
                    Dim etaDate As Date = shipDate.AddDays(ETD_to_ETA)

                    For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select
                        rowPOTORDR2.Item("PO_DATE_SHIP_BY") = shipDate
                        rowPOTORDR2.Item("LAST_DATE_SHIP_BY") = Now + ASCMAIN1.NowTSD
                        rowPOTORDR2.Item("LAST_OPER_SHIP_BY") = ASCMAIN1.USER_ID
                        rowPOTORDR2.Item("PO_DATE_ETA") = etaDate
                    Next
                End If

            Case "Update ETA Dates of Selected Rows"
                If grd.ActiveRow.IsDataRow Then
                    If grd.ActiveRow.DataChanged Then grd.ActiveRow.Update()
                    Dim etaDate As Date = grd.ActiveRow.Cells("PO_DATE_ETA").Value
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find _
                                                   (New Object() {grow.Cells("PO_ORDER_NO").Value, grow.Cells("PO_ORDER_LNO").Value})
                        rowPOTORDR2.Item("PO_DATE_ETA") = etaDate
                    Next
                End If

            Case "Update Ship Dates of Selected Rows"
                If grd.ActiveRow.IsDataRow Then
                    If grd.ActiveRow.DataChanged Then grd.ActiveRow.Update()
                    Dim shipDate As Date = grd.ActiveRow.Cells("PO_DATE_SHIP_BY").Value
                    Dim etaDate As Date = DateAdd(DateInterval.Month, 1, shipDate)
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find _
                                                     (New Object() {grow.Cells("PO_ORDER_NO").Value, grow.Cells("PO_ORDER_LNO").Value})
                        rowPOTORDR2.Item("PO_DATE_SHIP_BY") = shipDate
                        rowPOTORDR2.Item("LAST_DATE_SHIP_BY") = Now + ASCMAIN1.NowTSD
                        rowPOTORDR2.Item("LAST_OPER_SHIP_BY") = ASCMAIN1.USER_ID
                        rowPOTORDR2.Item("PO_DATE_ETA") = etaDate
                    Next
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Create a carton type containing all selected Style/Colors"
                If grdPOTORDRR.Selected.Rows.Count = 0 Then
                    If grdPOTORDRR.ActiveRow IsNot Nothing Then
                        grdPOTORDRR.Selected.Rows.Add(grdPOTORDRR.ActiveRow)
                    End If
                End If
                If grdPOTORDRR.Selected.Rows.Count = 0 Then
                    MsgBox("You must first select rows before creating a carton", MsgBoxStyle.OkOnly, "Cannot Create Carton")
                ElseIf grdPOTORDRR.ActiveRow IsNot Nothing AndAlso Not grdPOTORDRR.ActiveRow.Selected Then
                    MsgBox("Active Row is not Selected", MsgBoxStyle.OkOnly, "Cannot Create Carton")
                Else
                    Create_Carton_for_Selected_Styles()
                    grdPOTORDRR.Selected.Rows.Clear()
                End If

            Case "Create an individual carton type for All Style/Colors"
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDRR.Rows
                    grdPOTORDRR.Selected.Rows.Clear()
                    grdPOTORDRR.Selected.Rows.Add(grow)
                    Create_Carton_for_Selected_Styles()
                Next
                grdPOTORDRR.Selected.Rows.Clear()

            Case "Create an individual carton type for each selected Style/Color"
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDRR.Selected.Rows
                    grdPOTORDRR.Selected.Rows.Clear()
                    grdPOTORDRR.Selected.Rows.Add(grow)
                    Create_Carton_for_Selected_Styles()
                Next
                grdPOTORDRR.Selected.Rows.Clear()

            Case "Show All Carton Details"
                Setup_grdPOTORDR8()

            Case "Copy Line"
                Dim PO_ORDER_LNO As Int64 = Val(grd.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
                Copy_Line(PO_ORDER_LNO)

            Case "Split Line"
                Dim PO_QTY_OPN As Int64 = 0
                If ssdDZGRD.Value = 1 Then
                    PO_QTY_OPN = Val(grd.ActiveRow.Cells("PO_QTY_OPN").Value & "")
                Else
                    PO_QTY_OPN = Val(grd.ActiveRow.Cells("PO_QTY_OPN_DZ").Value & "")
                End If

                Dim PO_ORDER_LNO As Int64 = Val(grd.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
                Dim qty As Int64 = ASCMAIN1.Get_num_from_User("Qty " & IIf(ssdDZGRD.Value = 1, "", " (in Dz)") & "to Split to a New Line:",
                                                              "Enter Qty to Split from PO Line " & CStr(PO_ORDER_LNO) & " (Max = " & CStr(PO_QTY_OPN) & ")",
                                                              0, PO_QTY_OPN, 1, 0)
                If ASCMAIN1.response = -1 Then
                    Exit Sub
                End If
                Copy_Line(PO_ORDER_LNO, qty)

                For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDR2.Rows
                    If Val(grow.Cells("PO_ORDER_LNO").Value & "") = PO_ORDER_LNO Then
                        grow.Activate()
                    End If
                Next

            Case "Cost Calculator"
                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing Then
                        grd.ActiveRow.Selected = True
                    End If
                End If
                If grd.Selected.Rows.Count <> 0 Then Cost_Calculator()

            Case "Show email"
                If grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "PO-XMIT" _
                    Or grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "PO-XPED" Then
                    ' WHEN RUNNING FROM VS YOU MAY NEED TO SET THE ARCHIVE
                    ' ASCMAIN1.Folders("Archive")= "S:\NYA\Archive\NYA"
                    Dim FILENAME As String = grd.ActiveRow.Cells("EVENT_KEY").Value & ".EML"
                    Show_Document(ASCMAIN1.Folders("Archive") & "\email\Sent\" & FILENAME)
                End If

            Case "Show PO"
                Dim FILENAME As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value & "_" & CStr(Val(grd.ActiveRow.Cells("PO_HDR_CTR_REV").Value & "")) & ".PDF"
                Show_Document(ASCMAIN1.Folders("Archive") & "PO\" & FILENAME)

            Case "PO Shipment Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")


            Case "Style Master File"
                If grd.ActiveRow.Band.Index <> 1 Then Exit Sub
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If

        End Select
    End Sub

    Sub Show_Audit_Fields(tf As Boolean)
        For Each COLUMN_NAME As String In New String() {"PO_ORIG_DATE_SHIP_BY", "PO_ORIG_DATE_ETA", "LAST_OPER_SHIP_BY", "LAST_DATE_SHIP_BY", "SHIP_COST_CHANGE_USER", "SHIP_COST_CHANGE_DATE", "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}
            grdPOTORDR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tf
        Next
    End Sub

    Sub Copy_Line(PO_ORDER_LNO As Int64, Optional QTY As Int64 = 0)
        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
        Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE") ' grdPOTORDR2.ActiveRow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = rowPOTORDR2.Item("COLOR_CODE") ' grdPOTORDR2.ActiveRow.Cells("COLOR_CODE").Value
        Dim PO_COST_VCOST As Decimal = 0
        Dim C As String = IIf(ssdDZGRD.Value = 1, "PO_QTY_ORD", "PO_QTY_ORD_DZ")
        If ssdDZGRD.Value = 1 Then
            PO_COST_VCOST = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "") ' Val(grdPOTORDR2.ActiveRow.Cells("PO_COST_VCOST").Value & "")
        Else
            PO_COST_VCOST = Val(rowPOTORDR2.Item("PO_COST_VCOST_DZ") & "") 'Val(grdPOTORDR2.ActiveRow.Cells("PO_COST_VCOST_DZ").Value & "")
        End If
        If QTY <> 0 Then ' REDUCE QTY ON CURRENT LINE
            grdPOTORDR2.ActiveRow.Cells(C).Value = Val(grdPOTORDR2.ActiveRow.Cells(C).Value) - QTY
            grdPOTORDR2.ActiveRow.Update()
        End If
        If QTY = 0 Then QTY = Val(grdPOTORDR2.ActiveRow.Cells(C).Value & "")
        Dim grow As UltraWinGrid.UltraGridRow = Add_grdPOTORDR2(STYLE_CODE, COLOR_CODE, ssdDZGRD.Value, QTY, PO_COST_VCOST)

        grow.Cells("PO_COST_OTHER").Value = Val(rowPOTORDR2.Item("PO_COST_OTHER") & "")
        grow.Cells("PO_COST_QUOTA").Value = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "")
        grow.Update()

        Dim rowPOTORDR2_new As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, grow.Cells("PO_ORDER_LNO").Value})
        Dim PO_ORDER_LNO_new As Integer = Val(rowPOTORDR2_new.Item("PO_ORDER_LNO") & "")

        rowPOTORDR2_new.Item("PO_DATE_SHIP_BY") = rowPOTORDR2.Item("PO_DATE_SHIP_BY")
        rowPOTORDR2_new.Item("PO_DATE_ETA") = rowPOTORDR2.Item("PO_DATE_ETA")
        rowPOTORDR2_new.Item("PO_ORIG_DATE_SHIP_BY") = rowPOTORDR2.Item("PO_ORIG_DATE_SHIP_BY")
        rowPOTORDR2_new.Item("PO_ORIG_DATE_ETA") = rowPOTORDR2.Item("PO_ORIG_DATE_ETA")
        'rowPOTORDR2_new.Item("INIT_OPER") = ASCMAIN1.USER_ID
        'rowPOTORDR2_new.Item("LAST_OPER") = ASCMAIN1.USER_ID
        'rowPOTORDR2_new.Item("INIT_DATE") = DATETIME_STAMP
        'rowPOTORDR2_new.Item("LAST_DATE") = DATETIME_STAMP
        rowPOTORDR2_new.Item("LAST_OPER_SHIP_BY") = rowPOTORDR2.Item("LAST_OPER_SHIP_BY")
        rowPOTORDR2_new.Item("LAST_DATE_SHIP_BY") = rowPOTORDR2.Item("LAST_DATE_SHIP_BY")

        rowPOTORDR2_new.Item("PO_COST_MATLS") = rowPOTORDR2.Item("PO_COST_MATLS")
        rowPOTORDR2_new.Item("STYLE_NOTES") = rowPOTORDR2.Item("STYLE_NOTES")
        rowPOTORDR2_new.Item("CMT_NO") = rowPOTORDR2.Item("CMT_NO")
        rowPOTORDR2_new.Item("YIELD_QTY") = rowPOTORDR2.Item("YIELD_QTY")
        rowPOTORDR2_new.Item("PO_COST_MATLS_DZ") = rowPOTORDR2.Item("PO_COST_MATLS_DZ")
        rowPOTORDR2_new.Item("YARDS_CONSUMED") = rowPOTORDR2.Item("YARDS_CONSUMED")
        rowPOTORDR2_new.Item("FABRIC_COST") = rowPOTORDR2.Item("FABRIC_COST")
        rowPOTORDR2_new.Item("YARDS_CONSUMED") = rowPOTORDR2.Item("YARDS_CONSUMED")
        rowPOTORDR2_new.Item("YARDS_CONSUMED") = rowPOTORDR2.Item("YARDS_CONSUMED")
        rowPOTORDR2_new.Item("YARDS_CONSUMED") = rowPOTORDR2.Item("YARDS_CONSUMED")

        'PO_COST_QUOTA()
        'DFQUOTA()
        'CARTON_PACK_QTY()


        Sort_grdColumns(grdPOTORDR2, "PO_ORDER_LNO", , , False)

        For Each TABLE_NAME As String In New String() {"POTORDR3", "POTORDR4"}
            Dim sqlw As String = "PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlw)
                Dim row2 As DataRow = dst.Tables(TABLE_NAME).NewRow
                row2.ItemArray = row.ItemArray
                row2.Item("PO_ORDER_LNO") = PO_ORDER_LNO_new
                dst.Tables(TABLE_NAME).Rows.Add(row2)
            Next
        Next

        'dst.Tables("POTORDR3_LINE").Rows.Clear()
        'dst.Tables("POTORDR4_LINE").Rows.Clear()

    End Sub

    Sub Add_Colors(STYLE_CODE As String, tbl As DataTable, PRICE As Decimal)
        If tbl.Select("ISNULL(QTY,0)<>0").Length = 0 Then
            MsgBox("No Qty's Entered", MsgBoxStyle.OkOnly, "Cannot Add Colors")
            Exit Sub
        End If

        For Each rowICTCOLRM As DataRow In tbl.Select("ISNULL(QTY,0)<>0", "COLOR_CODE")
            Add_grdPOTORDR2(STYLE_CODE, rowICTCOLRM.Item("COLOR_CODE"), ssdDZGRD.Value, rowICTCOLRM.Item("QTY"), PRICE)
        Next
        Sort_grdColumns(grdPOTORDR2, "PO_ORDER_LNO")
    End Sub

    Function Add_grdPOTORDR2(STYLE_CODE As String, COLOR_CODE As String,
                        UM As Integer, PO_QTY_ORD As Int64, PO_COST_VCOST As Decimal) As UltraWinGrid.UltraGridRow

        If Absx1.optFor("FOB_CMT").Value = "I" Then
            If grdPOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No Then
                grdPOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            End If
        End If
        If grdPOTORDR2.ActiveRow IsNot Nothing AndAlso grdPOTORDR2.ActiveRow.IsAddRow Then
            grdPOTORDR2.ActiveRow.CancelUpdate()
        End If
        grdPOTORDR2.DisplayLayout.Bands(0).AddNew()
        With grdPOTORDR2.ActiveRow
            .Cells("STYLE_CODE").Value = STYLE_CODE
            .Cells("COLOR_CODE").Value = COLOR_CODE
            If UM = 1 Then
                .Cells("PO_QTY_ORD").Value = PO_QTY_ORD
                .Cells("PO_COST_VCOST").Value = PO_COST_VCOST
            Else
                .Cells("PO_QTY_ORD_DZ").Value = PO_QTY_ORD
                .Cells("PO_COST_VCOST_DZ").Value = PO_COST_VCOST
            End If
            .Update()
        End With

        If Absx1.optFor("FOB_CMT").Value = "I" Then
            If grdPOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop Then
                grdPOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            End If
        End If

        Return grdPOTORDR2.ActiveRow
    End Function

    Sub Update_POs_with_Case_Pack()
        If grdPOTORDRS.Selected.Rows.Count = 0 Then
            MsgBox("No Rows Selected", vbOKOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim PO_ORDER_NOs As New List(Of String)
        Dim PO_ORDER_NOs_LOCKED As New List(Of String)
        Dim msg As String = "All POs for Selected Styles have been Updated Successfully"
        Dim sqlw As String = "ISNULL(CARTON_PACK_QTY,0) <> 0 and (ISNULL(CARTON_PACK_QTY,0) <> ISNULL(CARTON_PACK_QTY_MIN,0) OR ISNULL(CARTON_PACK_QTY,0) <> ISNULL(CARTON_PACK_QTY_MAX,0))"
        For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDRS.Selected.Rows
            Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
            Dim rowPOTORDRS As DataRow = dst.Tables("POTORDRS").Rows.Find(STYLE_CODE)
            Dim CARTON_PACK_QTY As Int64 = Val(rowPOTORDRS.Item("CARTON_PACK_QTY") & "")
            Dim POS_UPDATED As String = ""
            Dim POS_SKIPPED As String = ""
            If CARTON_PACK_QTY = 0 Then
                POS_SKIPPED = ",All Skipped - Case Pack = 0"
            Else
                Dim PO_SHIPMENT_NO As String = grdPOTORDRS.Tag & ""
                ASCMAIN1.sql = "Select Distinct PO_ORDER_NO from POTORDR2" _
                   & " where STYLE_CODE = '" & STYLE_CODE & "'" _
                   & IIf(PO_SHIPMENT_NO = "",
                         "   and PO_STATUS = 'O'",
                         "   and (PO_ORDER_NO,PO_ORDER_LNO) in (Select PO_ORDER_NO,PO_ORDER_LNO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')") _
                   & "   and NVL(CARTON_PACK_QTY,0) <> " & CStr(CARTON_PACK_QTY)
                For Each rowPOTORDR1 As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim PO_ORDER_NO As String = rowPOTORDR1.Item("PO_ORDER_NO")
                    If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then
                        PO_ORDER_NOs.Add(PO_ORDER_NO)
                        If ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, , False, False, 1) Then
                            PO_ORDER_NOs_LOCKED.Add(PO_ORDER_NO)
                        End If
                    End If
                    If PO_ORDER_NOs_LOCKED.Contains(PO_ORDER_NO) Then
                        ASCMAIN1.sql = "Update POTORDR2 Set CARTON_PACK_QTY = :PARM1" _
                            & " where PO_ORDER_NO = :PARM2 and STYLE_CODE = :PARM3" _
                            & IIf(PO_SHIPMENT_NO = "",
                                    "   and PO_STATUS = 'O'",
                                    "   and (PO_ORDER_NO,PO_ORDER_LNO) in (Select PO_ORDER_NO,PO_ORDER_LNO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')")
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NVV", New Object() {CARTON_PACK_QTY, PO_ORDER_NO, STYLE_CODE})
                        POS_UPDATED &= "," & PO_ORDER_NO
                    Else
                        POS_SKIPPED &= "," & PO_ORDER_NO
                        msg = "Some POs may NOT have been Updated - Check Grid for Results"
                    End If
                Next
            End If

            rowPOTORDRS.Item("POS_UPDATED") = Mid(POS_UPDATED, 2)
            rowPOTORDRS.Item("POS_SKIPPED") = Mid(POS_SKIPPED, 2)
        Next

        grdPOTORDRS.DisplayLayout.Bands(0).Columns("POS_UPDATED").Hidden = False
        grdPOTORDRS.DisplayLayout.Bands(0).Columns("POS_SKIPPED").Hidden = False
        ASCMAIN1.MultiTask_Release(, , 1)
        MsgBox(msg & vbCrLf & vbCrLf & "Export Grid and Save for record of Updates",
               vbOKOnly, "Process Complete")

    End Sub
#End Region

    Sub Create_Carton_for_Selected_Styles()

        Dim rowPOTORDR7 As DataRow = dst.Tables("POTORDR7").NewRow
        rowPOTORDR7.Item("PO_ORDER_NO") = PO_ORDER_NO
        Dim CARTON_NO As Int32 = Val(dst.Tables("POTORDR7").Compute("MAX (CARTON_NO)", "") & "") + 1
        rowPOTORDR7.Item("CARTON_NO") = CARTON_NO
        rowPOTORDR7.Item("CARTON_COMMENTS") = ""
        dst.Tables("POTORDR7").Rows.Add(rowPOTORDR7)

        Dim CARTONS As Int32 = 0
        For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDRR.Selected.Rows
            Dim rowPOTORDR8 As DataRow = dst.Tables("POTORDR8").NewRow
            rowPOTORDR8.Item("PO_ORDER_NO") = PO_ORDER_NO
            rowPOTORDR8.Item("CARTON_NO") = CARTON_NO
            Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value
            rowPOTORDR8.Item("STYLE_CODE") = STYLE_CODE
            rowPOTORDR8.Item("COLOR_CODE") = COLOR_CODE
            dst.Tables("POTORDR8").Rows.Add(rowPOTORDR8)
            Dim rowPOTORDRR As DataRow = dst.Tables("POTORDRR").Rows.Find(New Object() {PO_ORDER_NO, STYLE_CODE, COLOR_CODE})
            Dim rows() As DataRow = dst.Tables("POTORDR2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
            Dim CARTON_PACK_QTY As Int32 = IIf(rows.Length = 0, 1, Val(rows(0).Item("CARTON_PACK_QTY") & ""))
            'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            'Dim CARTON_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
            If CARTON_PACK_QTY = 0 Then CARTON_PACK_QTY = 1
            rowPOTORDR8.Item("QTY") = CARTON_PACK_QTY
            If CARTONS = 0 Then
                Dim QTY As Int32 = Val(rowPOTORDRR.Item("QTY_VAR") & "")
                If CARTON_PACK_QTY <> 0 And QTY > 0 Then
                    CARTONS = QTY / CARTON_PACK_QTY
                    rowPOTORDR7.Item("CARTONS") = CARTONS
                End If
            End If
            For Each grow7 As UltraWinGrid.UltraGridRow In grdPOTORDR7.Rows
                If Val(grow7.Cells("CARTON_NO").Value) = CARTON_NO Then
                    grdPOTORDR7.ActiveRow = grow7
                    Check_for_PPK(grow7)
                End If
            Next
        Next
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "PO_ORDER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not ScreenMode Then Click_Command("View", e)
                End If

            Case "PO_REFERENCE", "PO_SPEC_ORDR_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    ASCMAIN1.sql = "Select PO_ORDER_NO, PO_REFERENCE, PO_SPEC_ORDR_NO, VEND_CODE " _
                        & " from POTORDR1 where " & COLUMN_NAME & " = '" & Absx1.txtFor(COLUMN_NAME).Text.ToUpper & "'"
                    Dim tbl As DataTable = ASCDATA1.GetDataTable

                    If tbl.Rows.Count <> 0 Then
                        If tbl.Rows.Count > 1 Then
                            Using frmmsg As New ASFMSGBF
                                frmmsg.Show_grd(tbl, Me, "There Is More Than One PO That Matches This Selection")
                                If frmmsg.grow Is Nothing Then
                                    Exit Sub
                                End If
                                Absx1.txtFor("PO_ORDER_NO").Text = frmmsg.grow.Cells("PO_ORDER_NO").Value
                                Click_Command("View")
                            End Using
                        Else
                            Absx1.txtFor("PO_ORDER_NO").Text = tbl.Rows(0).Item("PO_ORDER_NO")
                            Click_Command("View")
                        End If
                    Else
                        MsgBox("There Are No PO's That Matches This Selection.", MsgBoxStyle.OkOnly, "No Selections")
                    End If
                End If


        End Select
    End Sub

    Public Overrides Sub num_Leave(sender As Object, e As System.EventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Dependent_Calculations(COLUMN_NAME)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        With Absx1.txtFor(COLUMN_NAME)
            Select Case COLUMN_NAME

                Case "STYLE_CODE"
                    If .Text <> "" And STYLE_CODE <> .Text Then
                        Dim rowICTSTYL1 As DataRow = Validate_Style(.Text, True)
                        If STYLE_CODE <> "" Then
                            Absx1.numFor("SUB_UNIT_PACK_QTY").Value = rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & ""
                            Absx1.txtFor("CMT_NO").Text = rowICTSTYL1.Item("CMT_NO") & ""
                            CMT_NO = rowICTSTYL1.Item("CMT_NO") & ""
                            Setup_CMT()
                            If COLOR_CODEs.Count = 1 Then
                                Absx1.txtFor("COLOR_CODE").Text = COLOR_CODEs(0)
                            End If
                        End If
                    End If

                Case "CMT_NO"
                    If .Text <> "" And CMT_NO <> .Text Then
                        If Validate_Code("CMT_NO") Then
                            ' NEED TO SET UP A VIEW FOR VALIDATE TO WORK
                            CMT_NO = .Text
                            Call Setup_CMT()
                        End If
                    Else
                    End If

                Case "WHSE_CODE"
                    If Not IsLoading Then Check_ETA()

                Case "PORT_CODE_ORIG"
                    If Not IsLoading Then Check_ETA()

                Case "PO_SHIP_VIA"
                    If Absx1.txtFor(COLUMN_NAME).Text <> "" Then
                        Calculate_ETD_to_ETA()
                        If EntryMode = "N" Or EntryMode = "E" Then
                            If Absx1.dteFor("PO_DATE_ETA").Value & "" <> "" Then
                                Absx1.dteFor("PO_DATE_ETA").Value = CDate(Absx1.dteFor("PO_DATE_SHIP_BY").Value).AddDays(ETD_to_ETA)
                            End If
                        End If
                    End If

                Case Else

            End Select

        End With
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PO_ORDER_NO"
                Click_Command("View")
            Case "PO_SHIPMENT_NO"
                Dim PO_SHIPMENT_NO As String = txtPO_SHIPMENT_NO.Text
                Refresh_Styles_Open_POs(PO_SHIPMENT_NO)
        End Select
    End Sub

    Public Overrides Sub CheckedChanged_Special(COLUMN_NAME As String, chk As Infragistics.Win.UltraWinEditors.UltraCheckEditor)
        MyBase.CheckedChanged_Special(COLUMN_NAME, chk)
        If Me.IsDone Or Me.IsLoading Then Exit Sub
        '   Exit Sub ' GETTING ERRORS ALL OF THE TIME IN HERE
        Select Case COLUMN_NAME

            Case "LINE_CLOSED"
                If tabPOTORDR2.SelectedTab.Key = "Line Item Details" Then
                    Dim PO_QTY_OPN As Int32 = 0
                    ' Dim PO_QTY_OPN_DZ As Int32 = 0
                    If chk.Checked Then
                        PO_QTY_OPN = 0
                        ' PO_QTY_OPN_DZ = 0

                        Absx1.numFor("YIELD_QTY").Value = Val(Absx1.numFor("PO_QTY_SHP").Value & "") / (optUD.Value * Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & ""))
                    Else
                        PO_QTY_OPN = Val(Absx1.numFor("PO_QTY_ORD").Value & "") - Val(Absx1.numFor("PO_QTY_SHP").Value & "")
                        'PO_QTY_OPN_DZ = Val(Absx1.numFor("PO_QTY_ORD_DZ").Value & "") - Val(Absx1.numFor("PO_QTY_SHP_DZ").Value & "")
                        Absx1.numFor("YIELD_QTY").Value = Val(Absx1.numFor("PO_QTY_ORD").Value & "") / (Val(optUD.Value) * Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & ""))
                        If PO_QTY_OPN < 0 Then
                            PO_QTY_OPN = 0
                            'PO_QTY_OPN_DZ = 0
                        End If
                    End If
                    Absx1.numFor("PO_QTY_OPN").Value = PO_QTY_OPN
                    ' Absx1.numFor("PO_QTY_OPN_DZ").Value = PO_QTY_OPN_DZ

                End If

            Case "LINE_FINISHED"
                'Don't forget there is a similar area in the grid for CMT with Inv and FOB types.
                Dim EMsg As String = ""

                'Check all the reasons why we can not un-check this.
                If chkFINISHED.Checked Then 'Tried to close the line.
                    'Can't if They used a Dummy CMT.
                    If UCase(Absx1.txtFor("CMT_NO").Text) = "DUM" And Not chkFINISHED.Checked Then
                        EMsg = EMsg & "CMT is a Dummy" & vbCrLf
                        chkFINISHED.Checked = False
                    End If

                    If Absx1.txtFor("STYLE_CODE").Text = "" Then EMsg = EMsg & "Style Code Is Blank" & vbCrLf
                    If Absx1.txtFor("COLOR_CODE").Text = "" Then EMsg = EMsg & "Color Code Is Blank" & vbCrLf
                    If Val(Absx1.numFor("PO_QTY_ORD").Value) = 0 Then EMsg = EMsg & "Qty Ordered is 0" & vbCrLf
                    If Val(Absx1.numFor("YIELD_QTY").Value) = 0 Then EMsg = EMsg & "Production is 0" & vbCrLf
                    If Absx1.txtFor("STYLE_CODE").Text = "" Then EMsg = EMsg & "Style Code Is Blank" & vbCrLf
                    If Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Then EMsg = EMsg & "ETA Date Is Blank" & vbCrLf

                    If 1 <> 1 Then 'Remove this when Gabe Says Costs are Manditory.
                        If Val(Absx1.numFor("PO_COST_VCOST").Value) = 0 Then EMsg = EMsg & "Vendor Cost is 0" & vbCrLf
                        If Val(Absx1.numFor("PO_COST_MATLS").Value) = 0 Then EMsg = EMsg & "Material Cost is 0" & vbCrLf
                        If Val(Absx1.numFor("PO_COST").Value) = 0 Then EMsg = EMsg & "PO Cost is 0" & vbCrLf
                        If Val(Absx1.numFor("PO_COST_VCOST_DZ").Value) = 0 Then EMsg = EMsg & "Vendor Cost (Dz) is 0" & vbCrLf
                        If Val(Absx1.numFor("PO_COST_MATLS_DZ").Value) = 0 Then EMsg = EMsg & "Material Cost (Dz) is 0" & vbCrLf
                    End If

                Else
                    'Don't let them un-check if it's been used on shipment.
                    ' SHOULDNT THIS BE CHECKING SHIPPED?
                    Dim sqlw As String = "PO_ORDER_NO = " & PO_ORDER_NO & " AND PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
                    If Val(dst.Tables("POTSHIPX").Compute("SUM(PO_QTY_REC)", sqlw) & "") <> 0 Then
                        EMsg = EMsg & "You Can Not Re-Open A Line That Has Been Recv'd On A Shipment" & vbCrLf
                    End If
                End If

                If EMsg <> "" Then
                    MsgBox(EMsg, vbOKOnly, "You Can Not Perform This Action For The Following Reasons")
                    If chkFINISHED.Checked Then
                        chkFINISHED.Checked = False
                    Else
                        chkFINISHED.Checked = True
                    End If
                Else

                End If
        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "PO_DATE_SHIP_BY"
                If Absx1.dteFor(COLUMN_NAME).Value & "" <> "" Then
                    Calculate_ETD_to_ETA()
                    If EntryMode = "N" Or EntryMode = "E" Then Absx1.dteFor("PO_DATE_ETA").Value = CDate(Absx1.dteFor(COLUMN_NAME).Value).AddDays(ETD_to_ETA)
                End If

            Case "POTORDR2_LINE.PO_DATE_SHIP_BY"
                If Absx1.dteFor(COLUMN_NAME).Value & "" <> "" Then
                    Absx1.dteFor("POTORDR2_LINE.PO_DATE_ETA").Value = CDate(Absx1.dteFor(COLUMN_NAME).Value).AddDays(ETD_to_ETA)
                    Stop
                    ' ALSO NEED TO UPDATE THESE:
                    Absx1.txtFor("POTORDR2_LINE.LAST_OPER_SHIP_BY").Text = ASCMAIN1.USER_ID
                    Absx1.dteFor("POTORDR2_LINE.LAST_DATE_SHIP_BY").Value = Now + ASCMAIN1.NowTSD
                End If
        End Select

    End Sub

    Public Overrides Sub num_ValueChanged(sender As Object, e As System.EventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If COLUMN_NAME = "PO_COST_COMM" Or COLUMN_NAME = "PO_COST_BUFFER" Then
            Dependent_Calculations(COLUMN_NAME)
        End If

        Select Case COLUMN_NAME
            Case "PO_COST_COMM"

            Case "PO_COST_BUFFER"

            Case "PO_COST_OTHER"

            Case "PO_COST_QUOTA"


        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "FREIGHT_ENTERED_BY"
            '    Dim blnFREIGHT_ENTERED_BY_Container As Boolean = (Absx1.optFor(COLUMN_NAME).Value = "C")
            '    With grdPOTORDR3.DisplayLayout.Bands(0)
            '        .Columns("CBM_RATE").Hidden = blnFREIGHT_ENTERED_BY_Container
            '        .Columns("CBM").Hidden = blnFREIGHT_ENTERED_BY_Container

            '        .Columns("BOL_FEE").Hidden = blnFREIGHT_ENTERED_BY_Container
            '        .Columns("FREIGHT_AMT").Hidden = Not blnFREIGHT_ENTERED_BY_Container
            '        .Columns("CBM").Hidden = Not blnFREIGHT_ENTERED_BY_Container
            '        .Columns("TRUCKING").Hidden = Not blnFREIGHT_ENTERED_BY_Container
            '    End With
        End Select
    End Sub
#End Region

    Private Sub grdPOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDRX.AfterRowActivate

    End Sub

    Private Sub grdPOTORDRX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTORDRX.DoubleClickRow
        If grdPOTORDRX.ActiveRow IsNot Nothing AndAlso grdPOTORDRX.ActiveRow.IsDataRow Then
            Absx1.txtFor("PO_ORDER_NO").Text = grdPOTORDRX.ActiveRow.Cells("PO_ORDER_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Load_POTORDRX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim sqlw As String = ""
        If optStatus.Value = "T" Then
            sqlw = " where POTORDR1.PO_STATUS = 'O' and NVL(POTORDR1.PO_XMIT_IND,'0') = '0' and PO_APPR_BY is Not Null"
            grdPOTORDRX.Text = "Open POs NOT Transmitted"
        ElseIf optStatus.Value = "A" Then
            sqlw = " where POTORDR1.PO_STATUS = 'O' and PO_APPR_BY is Null and NVL(PO_APPR_PENDING,'0') = '1'"
            grdPOTORDRX.Text = "Open POs NOT Approved"
        ElseIf optStatus.Value = "W" Then
            sqlw = " where POTORDR1.PO_STATUS = 'O' and PO_APPR_BY is Null and NVL(PO_APPR_PENDING,'0') = '0'"
            grdPOTORDRX.Text = "Open POs NOT yet Ready for Approval Queue"
        ElseIf optStatus.Value = "O" Then
            sqlw = " where POTORDR1.PO_STATUS = 'O'"
            grdPOTORDRX.Text = "Open POs"
        ElseIf optStatus.Value = "P" Then
            sqlw = " where POTORDR1.PO_STATUS = 'O' and NVL(POTORDR1.PO_WEB_VISIBLE,'0') = '1' and NVL(POTORDR1.PO_XMIT_IND,'0') = '1' and POTORDR1.PO_APPR_BY is Not Null"
            grdPOTORDRX.Text = "Open POs Visible on the Portal"
        ElseIf optStatus.Value = "N" Then
            sqlw = " where POTORDR1.PO_STATUS = 'O' and (NVL(POTORDR1.PO_WEB_VISIBLE,'0') = '0' or NVL(POTORDR1.PO_XMIT_IND,'0') = '0' and POTORDR1.PO_APPR_BY is Null)"
            grdPOTORDRX.Text = "Open POs NOT Visible on the Portal"
        Else
            grdPOTORDRX.Text = "All POs"
        End If

        Dim sqlDTL As String = "Select POTORDR2.PO_ORDER_NO" & vbCrLf _
            & IIf(chkSplitByShipDate.Checked, ",POTORDR2.PO_DATE_SHIP_BY", "") _
            & ", SUM (POTORDR2.PO_QTY_ORD) PO_QTY_ORD" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_REC) PO_QTY_REC" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST) PO_AMT_ORD" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_SHP * POTORDR2.PO_COST) PO_AMT_SHP" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_REC * POTORDR2.PO_COST) PO_AMT_REC" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST) PO_AMT_OPN" & vbCrLf _
            & ", SUM (DECODE(POTORDR2.PO_CONF_NO,NULL,0,1)) PO_LINES_CONF" & vbCrLf _
            & ", COUNT (*) PO_LINES" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_ORD / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CTNS_ORD" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_SHP / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CTNS_SHP" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_OPN / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CTNS_OPN" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_ORD  * NVL(ICTSTYL1.CASE_CUBE,0) / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_ORD" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_SHP  * NVL(ICTSTYL1.CASE_CUBE,0) / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_SHP" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_OPN  * NVL(ICTSTYL1.CASE_CUBE,0) / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_OPN" & vbCrLf _
            & ", MIN (POTORDR2.PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MIN" & vbCrLf _
            & ", MIN (POTORDR2.PO_DATE_ETA) PO_DATE_ETA_MIN" & vbCrLf _
            & ", MAX (POTORDR2.PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MAX" & vbCrLf _
            & ", MAX (POTORDR2.PO_DATE_ETA) PO_DATE_ETA_MAX" & vbCrLf _
            & " from  POTORDR2,ICTSTYL1 where POTORDR2.PO_ORDER_NO in " & vbCrLf _
            & " (Select PO_ORDER_NO from POTORDR1 " & sqlw & ")" _
            & " and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
            & " group by POTORDR2.PO_ORDER_NO" & vbCrLf _
            & IIf(chkSplitByShipDate.Checked, ",POTORDR2.PO_DATE_SHIP_BY", "")


        'ASCMAIN1.sql = "Select POTORDR2.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CASE_CUBE" & vbCrLf _
        '     & " from POTORDR2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
        '     & "  and POTORDR2.PO_ORDER_NO = :PARM1"
        'Create_TDA(.Tables.Add, "POTORDR2", "**", 0, True, "V", 2)
        '.Tables("POTORDR2").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(PO_QTY_ORD,0) / ISNULL(CARTON_PACK_QTY,0))")
        '.Tables("POTORDR2").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")
        '.Tables("POTORDR2").Columns.Add("CONFIRMED")

        Dim sqlMyPOs As String = ""
        If optMyPOs.Value = "M" Then
            If optStatus.Value = "A" Then
                sqlMyPOs = " and (NVL(POTORDR1.PO_APPR_QUEUE,'" & ASCMAIN1.USER_ID & "') = '" & ASCMAIN1.USER_ID & "')"
            Else
                sqlMyPOs = " and (POTORDR1.INIT_OPER = '" & ASCMAIN1.USER_ID & "' or POTORDR1.LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            End If
        End If

        ASCMAIN1.sql = "Select POTORDR1.*" & vbCrLf _
            & ", X.PO_QTY_ORD, X.PO_QTY_SHP, X.PO_QTY_REC, X.PO_QTY_OPN " & vbCrLf _
            & ", X.PO_AMT_ORD, X.PO_AMT_SHP, X.PO_AMT_REC, X.PO_AMT_OPN " & vbCrLf _
            & ", X.PO_LINES_CONF, X.PO_LINES" & vbCrLf _
            & ", X.PO_CTNS_ORD, X.PO_CTNS_SHP, X.PO_CTNS_OPN" & vbCrLf _
            & ", X.PO_CUBE_ORD, X.PO_CUBE_SHP, X.PO_CUBE_OPN " & vbCrLf _
            & ", X.PO_DATE_SHIP_BY_MIN, X.PO_DATE_ETA_MIN, X.PO_DATE_SHIP_BY_MAX, X.PO_DATE_ETA_MAX" & vbCrLf _
            & ", SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & " from (" & sqlDTL & ") X,POTORDR1,SOTORDR1" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sqlw & " and SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO and POTORDR1.PO_ORDER_NO = X.PO_ORDER_NO" & sqlMyPOs)

        Fill_Records("POTORDRX", "", True, ASCMAIN1.sql)

        If chkSplitByShipDate.Checked Then
            Sort_grdColumns(grdPOTORDRX, "PO_DATE_ETA_MIN".ToLower)
        Else
            Sort_grdColumns(grdPOTORDRX, "PO_DATE_ETA".ToLower)
        End If

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            Dim portalVisibilityColumnLayout As Boolean = (optStatus.Value = "N")
            Dim defaultApprByPosition As Integer = grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_APPR_DATE").Header.VisiblePosition - 1
            Dim portalVisibilityColumnsStart As Integer = grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_REFERENCE").Header.VisiblePosition - 1
            With grdPOTORDRX.DisplayLayout.Bands(0)
                With .Columns("PO_WEB_VISIBLE")
                    .Style = UltraWinGrid.ColumnStyle.CheckBox
                    .Hidden = Not portalVisibilityColumnLayout
                    .Width = 60
                    .Header.Caption = "Visible"
                    .Header.SetVisiblePosition(portalVisibilityColumnsStart + 1, False)
                End With
                With .Columns("PO_XMIT_IND")
                    .Style = UltraWinGrid.ColumnStyle.CheckBox
                    .Hidden = Not portalVisibilityColumnLayout
                    .Width = 70
                    .Header.Caption = "Transmitted"
                    .Header.SetVisiblePosition(portalVisibilityColumnsStart + 2, False)
                End With
                With .Columns("PO_APPR_BY")
                    .Header.SetVisiblePosition(IIf(portalVisibilityColumnLayout, portalVisibilityColumnsStart + 3, defaultApprByPosition), False)
                End With
            End With
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        tabPO.SelectedTab = tabPO.Tabs("Open POs")
        'grdPOTORDR2.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
    End Sub

    Public Overrides Function CustomSummary_End(
      ByVal summarySettings As UltraWinGrid.SummarySettings,
          ByVal rows As UltraWinGrid.RowsCollection,
          ByVal CustomValue As Double,
          ByVal grd As UltraWinGrid.UltraGrid) As Double

        '    CustomValue = 0
        '    'Dim TOTALS As New Dictionary(Of String, Decimal)

        '    Select Case grd.Name

        '        Case "grdPOTORDR3"
        '            Dim KEY As String = summarySettings.Key

        '            Dim COLUMN_NAME_QTY As String = "PO_QTY_SR"
        '            If KEY.EndsWith("_DZ") Then COLUMN_NAME_QTY = "PO_QTY_SR_DZ"

        '            For Each grow As UltraWinGrid.UltraGridRow In rows
        '                CustomValue += Val(grow.Cells(COLUMN_NAME_QTY).Value & "") * Val(grow.Cells(KEY).Value & "")
        '            Next

        '        Case Else
        '            MsgBox("CustomSummary_End " & grd.Name)
        '    End Select

        '    Return CustomValue

    End Function

#Region "grdPOTORDR2"

    Private Sub grdPOTORDR2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR2.AfterCellUpdate

        Dim UMF As Decimal = 0
        Dim SUB_UNIT_PACK_QTY As Integer = Val(e.Cell.Row.Cells("SUB_UNIT_PACK_QTY").Value & "")
        If SUB_UNIT_PACK_QTY = 0 Then
            UMF = 1
        Else
            UMF = 12 / SUB_UNIT_PACK_QTY
        End If

        ' when we speak of dozens, we mean dozens of pcs, not units
        ' - that is why we need to factor in SUB_UNIT_PACK_QTY

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim rowICTSTYL1 As DataRow = Validate_Style(e.Cell.Value & "", True)
                If STYLE_CODE <> "" Then
                    ' SEE SAVE_LNO
                    e.Cell.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
                    e.Cell.Row.Cells("INNER_PACK_QTY").Value = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")
                    e.Cell.Row.Cells("SUB_UNIT_PACK_QTY").Value = Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "")
                    e.Cell.Row.Cells("CASE_CUBE").Value = Val(rowICTSTYL1.Item("CASE_CUBE") & "")
                    ' NOTE THAT IS IS SIMPLY RECORDING THE UOM SETTING THAT WAS ACTIVE WHEN THE LINE WAS ENTERED
                    ' IT IS NOT TOO MEANINFUL
                    e.Cell.Row.Cells("PO_QTY_UOM").Value = Val(ssdDZGRD.Value & "")
                    If COLOR_CODEs.Count = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                    End If

                    If Absx1.txtFor("PO_APPR_QUEUE").Text = "" Then
                        Dim rowSOTSDIV1 As DataRow = LookUp("SOTSDIV1", rowICTSTYL1("SALES_DIVISION_CODE") & "")
                        If rowSOTSDIV1 IsNot Nothing Then
                            Absx1.txtFor("PO_APPR_QUEUE").Text = rowSOTSDIV1.Item("SALES_DIVISION_MGR") & ""
                        End If
                    End If

                    Dim rowICTSTYV1 As DataRow = LookUp("ICTSTYV1", New String() {STYLE_CODE, Absx1.txtFor("VEND_CODE").Text})
                    If rowICTSTYV1 IsNot Nothing Then
                        If rowICTSTYV1.Item("NEW_PO_COST_DATE") & "" <> "" AndAlso Format(Absx1.dteFor("PO_DATE_ORDERED").Value, "yyyyMMdd") >= Format(rowICTSTYV1.Item("NEW_PO_COST_DATE"), "yyyyMMdd") Then
                            e.Cell.Row.Cells("PO_COST_VCOST").Value = rowICTSTYV1.Item("NEW_PO_COST")
                        Else
                            e.Cell.Row.Cells("PO_COST_VCOST").Value = rowICTSTYV1.Item("PO_COST")
                        End If

                    End If


                    If rowICTSTYL1.Item("VEND_CODE") & "" <> "" Then
                        If rowICTSTYL1.Item("VEND_CODE") <> Absx1.txtFor("VEND_CODE").Text Then
                            e.Cell.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
                            e.Cell.Row.Cells("STYLE_CODE").ToolTipText = "Style's Default Supplier is " & rowICTSTYL1.Item("VEND_CODE")
                        End If
                    End If
                End If

                If e.Cell.Row.IsAddRow Then
                    e.Cell.Row.Cells("PO_DATE_SHIP_BY").Value = Absx1.dteFor("PO_DATE_SHIP_BY").Value
                    e.Cell.Row.Cells("PO_DATE_ETA").Value = Absx1.dteFor("PO_DATE_ETA").Value
                    If rowPOTORDR1.Item("VEND_CODE") = "AT" And (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
                        ' MAYBE WE SET THESE UP IN APTVEND1 AS OVERRIDES, BUT HOW TO OVERRIDE WITH A 0?
                        'Changed to 3 from 4 on 12/30/14 Per Anna - WR.
                        'Changed from 3 To 2.5 Per Anna 1/20/17 - WR.
                        'Changed from 2.5 to 2.0 Per Anna 7/7/19 - WR.
                        'Changed from 2.0 to 2.5 Per Anna 7/17/19 - WR.
                        'Changed from 2.5 to 2.0 Per Anna 1/26/21 - WR.
                        'Changed from 2.0 to 0.0 Per Anna 8/16/24 - WR.
                        e.Cell.Row.Cells("PO_COST_COMM").Value = 0.0
                        e.Cell.Row.Cells("PO_COST_BUFFER").Value = 1
                    Else
                        ' NOTE THAT VAN KEEPS 0 COMM IN PARM FILE SO THAT COMM IS 0 FOR ALL BUT AT
                        e.Cell.Row.Cells("PO_COST_COMM").Value = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_COMM") & "")
                        e.Cell.Row.Cells("PO_COST_BUFFER").Value = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_BUFFER") & "")
                    End If
                End If

            Case "PO_QTY_ORD"
                grdPOTORDR2.Tag = "PO_QTY_ORD"
                Dim PO_QTY_OPN As Int64 = Val(e.Cell.Row.Cells("PO_QTY_ORD").Value & "") _
                                        - Val(e.Cell.Row.Cells("PO_QTY_SHP").Value & "")
                If e.Cell.Row.Cells("LINE_CLOSED").Value & "" = "1" Or PO_QTY_OPN < 0 Then
                    PO_QTY_OPN = 0
                End If
                e.Cell.Row.Cells("PO_QTY_OPN").Value = PO_QTY_OPN
                If grdPOTORDR2.Tag <> "PO_QTY_ORD_DZ" Then
                    e.Cell.Row.Cells("PO_QTY_ORD_DZ").Value = Val(e.Cell.Row.Cells("PO_QTY_ORD").Value & "") / UMF
                End If
                grdPOTORDR2.Tag = ""

            Case "PO_QTY_ORD_DZ"
                If grdPOTORDR2.Tag <> "PO_QTY_ORD" Then
                    grdPOTORDR2.Tag = "PO_QTY_ORD_DZ"
                    e.Cell.Row.Cells("PO_QTY_ORD").Value = Val(e.Cell.Row.Cells("PO_QTY_ORD_DZ").Value & "") * UMF
                    grdPOTORDR2.Tag = ""
                End If

            Case "LINE_CLOSED"
                If e.Cell.Row.Cells("LINE_CLOSED").Value & "" = "1" Then
                    e.Cell.Row.Cells("PO_QTY_OPN").Value = 0
                Else
                    Dim QTY As Int64 = Val(e.Cell.Row.Cells("PO_QTY_ORD").Value & "") - Val(e.Cell.Row.Cells("PO_QTY_SHP").Value & "")
                    If QTY < 0 Then
                        QTY = 0
                    End If
                    e.Cell.Row.Cells("PO_QTY_OPN").Value = QTY ' PO_QTY_OPN_DZ IS CALCULATED
                End If

            Case "PO_COST_VCOST", "PO_COST_MATLS"
                Dim C_UN As String = e.Cell.Column.Key
                Dim C_DZ As String = C_UN & "_DZ"
                If grdPOTORDR2.Tag <> C_DZ Then
                    grdPOTORDR2.Tag = C_UN
                    e.Cell.Row.Cells(C_DZ).Value = System.Math.Round(Val(e.Cell.Row.Cells(C_UN).Value & "") * UMF, 6)
                    ReCalculate_PO_Cost()
                    grdPOTORDR2.Tag = ""
                End If

            Case "PO_COST_VCOST_DZ", "PO_COST_MATLS_DZ"
                Dim C_DZ As String = e.Cell.Column.Key
                Dim C_UN As String = Mid(C_DZ, 1, C_DZ.Length - 3)
                If grdPOTORDR2.Tag <> C_UN Then
                    grdPOTORDR2.Tag = C_DZ
                    e.Cell.Row.Cells(C_UN).Value = System.Math.Round(Val(e.Cell.Row.Cells(C_DZ).Value & "") / UMF, 6)
                    ReCalculate_PO_Cost()
                    grdPOTORDR2.Tag = ""
                End If

            Case "PO_COST_OTHER", "PO_COST_QUOTA"  ' THESE COSTS ARE STORED IN DB AS PER DZ PCS
                Dim C_DZ As String = e.Cell.Column.Key
                Dim C_UN As String = C_DZ & "_UN"
                If grdPOTORDR2.Tag <> C_UN Then
                    grdPOTORDR2.Tag = C_DZ
                    e.Cell.Row.Cells(C_UN).Value = System.Math.Round(Val(e.Cell.Row.Cells(C_DZ).Value & "") / UMF, 6)
                    ReCalculate_PO_Cost()
                    grdPOTORDR2.Tag = ""
                End If

            Case "PO_COST_OTHER_UN", "PO_COST_QUOTA_UN"  ' THESE (PER UNIT) VALUES ARE NOT STORED IN DB
                Dim C_UN As String = e.Cell.Column.Key
                Dim C_DZ As String = Mid(C_UN, 1, C_UN.Length - 3)
                If grdPOTORDR2.Tag <> C_DZ Then
                    grdPOTORDR2.Tag = C_UN
                    e.Cell.Row.Cells(C_DZ).Value = System.Math.Round(Val(e.Cell.Row.Cells(C_UN).Value & "") * UMF, 6)
                    ReCalculate_PO_Cost()
                    grdPOTORDR2.Tag = ""
                End If

            Case "PO_COST_COMM", "PO_COST_BUFFER"  ' THESE ARE %
                ReCalculate_PO_Cost()

            Case "YARDS_CONSUMED", "FABRIC_COST"
                e.Cell.Row.Cells("PO_COST_MATLS_DZ").Value = Val(e.Cell.Row.Cells("FABRIC_COST").Value & "") _
                                                           * Val(e.Cell.Row.Cells("YARDS_CONSUMED").Value & "")
                'CONSUMPTION IS ALWAYS IN DZ OF PCS

            Case "CONFIRMED"
                If e.Cell.Row.IsAddRow Then
                Else
                    e.Cell.Row.Cells("LAST_DATE_SHIP_BY").Value = Now + ASCMAIN1.NowTSD
                    e.Cell.Row.Cells("LAST_OPER_SHIP_BY").Value = ASCMAIN1.USER_ID
                    e.Cell.Row.Update()
                End If

            Case "PO_DATE_SHIP_BY"
                If e.Cell.Value & "" <> "" Then
                    e.Cell.Row.Cells("LAST_DATE_SHIP_BY").Value = Now + ASCMAIN1.NowTSD
                    e.Cell.Row.Cells("LAST_OPER_SHIP_BY").Value = ASCMAIN1.USER_ID
                    ' If e.Cell.Row.Cells("PO_DATE_ETA").Value & "" = "" Then
                    e.Cell.Row.Cells("PO_DATE_ETA").Value = CDate(e.Cell.Value).AddDays(ETD_to_ETA)
                    'End If
                End If

            Case "LINE_FINISHED"
                'Don't forget there is code in the CMT detail grid that handles a similar function there.
                Dim SQL As String = ""
                Dim EMsg As String = ""

                'Check all the reasons why we can not un-check this.
                With e.Cell.Row
                    If .Cells("LINE_FINISHED").Value & "" = "1" Then  'Tried to close the line.
                        If .Cells("STYLE_CODE").Value & "" = "" Then EMsg &= "Style Code Is Blank" & vbCrLf
                        If .Cells("COLOR_CODE").Value & "" = "" Then EMsg &= "Color Code Is Blank" & vbCrLf
                        If Val(.Cells("PO_QTY_ORD").Value & "") = 0 Then EMsg &= "Qty Ordered is 0" & vbCrLf
                        If .Cells("PO_DATE_SHIP_BY").Value & "" = "" Then EMsg &= "Ship By Date Is Blank" & vbCrLf
                        If .Cells("PO_DATE_ETA").Value & "" = "" Then EMsg &= "ETA Date Is Blank" & vbCrLf

                        If "Remove this when Mgmt says to make costs mandatory before saying that a line is finished" <> "" Then
                            'do nothing
                        Else
                            'PO_COST = 0
                            If Val(.Cells("PO_COST").Value) = 0 Then EMsg &= "PO Cost is 0" & vbCrLf
                            If Val(.Cells("PO_COST_VCOST").Value) = 0 Then EMsg &= "Vendor Cost is 0" & vbCrLf
                            If Val(.Cells("PO_COST_MATLS").Value) = 0 Then EMsg &= "Material Cost is 0" & vbCrLf
                            If Val(.Cells("PO_COST_VCOST_DZ").Value) = 0 Then EMsg &= "Vendor Cost (Dz) is 0" & vbCrLf
                            If Val(.Cells("PO_COST_MATLS_DZ").Value) = 0 Then EMsg &= "Material Cost (Dz) is 0" & vbCrLf
                        End If
                    Else
                        If Check_Shipped(PO_ORDER_NO, PO_ORDER_LNO) Then
                            If .Cells("LINE_FINISHED").Value & "" = "" Then
                                ' if the value in the column was changed to unchecked
                                EMsg &= "You Can Not Re-Open A Line That Has Been Used On A Shipment" & vbCrLf
                            End If
                        End If
                    End If

                End With

                If EMsg <> "" Then
                    MsgBox(EMsg, MsgBoxStyle.OkOnly, "You Can Not Perform This Action For The Following Reasons")
                    If e.Cell.Row.Cells("LINE_FINISHED").Value = "1" Then
                        e.Cell.Row.Cells("LINE_FINISHED").Value = "0"
                    Else
                        e.Cell.Row.Cells("LINE_FINISHED").Value = "1"
                    End If
                    e.Cell.Row.Update()
                End If

            Case Else

        End Select
    End Sub

    Private Sub grdPOTORDR2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTORDR2.AfterRowActivate
        Setup_grdPOTORDR2_ActiveRow()

        '    UltraExplorerBar1.Groups("Cost Calculation").Visible = False

        If grdPOTORDR2.ActiveRow.Cells("LINE_FINISHED").Value & "" = "1" Then
            Lock_Line(True)
        Else
            Lock_Line(False)
        End If

        Set_POTSHIPX()

        If Absx1.optFor("FOB_CMT").Value = "I" Then
            Exit Sub
        End If

        With grdPOTORDR2.DisplayLayout.Bands(0)
            If grdPOTORDR2.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

                If Val(grdPOTORDR2.ActiveRow.Cells("PO_QTY_REC").Value & "") <> 0 _
                Or grdPOTORDR2.ActiveRow.Cells("COST_COMPLETE").Value & "" = "1" Then
                    .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If

                If rowPOTORDR1.Item("ORDR_NO") & "" <> "" Then
                    .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("PO_QTY_ORD").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("PO_QTY_OPN").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End If
        End With

        If grdPOTORDR2.ActiveRow.IsAddRow Then
            If grdPOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                grdPOTORDR2.ActiveCell = grdPOTORDR2.ActiveRow.Cells("STYLE_CODE")
            End If

            grdPOTORDR6.Visible = False
        Else

            If grdPOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & "" <> "" And
                grdPOTORDR2.ActiveRow.Cells("STYLE_CODE").Value <> STYLE_CODE Then
                Validate_Style(grdPOTORDR2.ActiveRow.Cells("STYLE_CODE").Value, False)
            End If

            grdPOTORDR6.Visible = True

            Dim dvw As DataView = DirectCast(grdPOTORDR6.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PO_ORDER_LNO = " & grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value
        End If

    End Sub

    Sub Setup_Style_Image()
        If grdPOTORDR2.ActiveRow Is Nothing Then
            imgSTYLE.Visible = False
            Exit Sub
        End If

        'If tabDetails.Tabs("Style").Visible Then 'UltraExplorerBar1.Groups("Style Info").Visible Then
        Dim STYLE_CODE As String = grdPOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = grdPOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
        grpImage.Text = "Style " & STYLE_CODE
        Dim rowICTSTYL1 As DataRow = Nothing
        If STYLE_CODE <> "" Then rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
        If rowICTSTYL1 IsNot Nothing Then

            Absx1.txtFor("SIZE_SCALE").Text = rowICTSTYL1.Item("SIZE_SCALE") & ""
            Absx1.txtFor("PURCH_NOTES").Text = rowICTSTYL1.Item("PURCH_NOTES") & ""
            imgSTYLE.Image = TAC.ICCMAIN1.Get_Image(Me, rowICTSTYL1, Nothing)

            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            Dim IMAGE_NAME As String = ""
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                IMAGE_NAME = rowICTSTYL1.Item("IMAGE_NAME") & ""
            ElseIf ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                IMAGE_NAME = ""
            End If
            imgSTYLE.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , ) ' imgba)
            imgSTYLE.Visible = True
            ' tabDetails.SelectedTab = tabDetails.Tabs("Style")
        Else
            Absx1.txtFor("SIZE_SCALE").Text = ""
            Absx1.txtFor("PURCH_NOTES").Text = ""
            imgSTYLE.Visible = False
            'tabDetails.SelectedTab = tabDetails.Tabs("Style")
        End If


        'If txtIMAGE_NAME.Text = "" Then
        '    imgSTYLE.Image = Nothing
        'Else
        '    Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        '    Dim IMAGE_NAME As String = txtIMAGE_NAME.Text
        '    imgSTYLE.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , ) ' imgba)
        'End If

        ' End If
    End Sub

    Sub Setup_grdPOTORDR2_ActiveRow()

        If (grdPOTORDR2.ActiveRow Is Nothing OrElse Not grdPOTORDR2.ActiveRow.IsDataRow) Then
            'grdPOTSHIPX.Visible = False
            tabDetails.Visible = False

        Else
            Dim PO_ORDER_LNO As Int32 = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
            'grdPOTSHIPX.Visible = True
            tabDetails.Visible = True

            If EntryMode = "N" Then
                tabDetails.Tabs("Shp").Visible = False
            Else
                tabDetails.Tabs("Shp").Visible = True
                grdPOTSHIPX.Text = "Shipments for PO Line " & CStr(PO_ORDER_LNO)
                Dim dvw As DataView = DirectCast(grdPOTSHIPX.DataSource, DataTable).DefaultView
                dvw.RowFilter = "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)

                If tabDetails.Tabs("Style").Visible Then
                    Setup_Style_Image()
                End If
            End If
            If ScreenMode And subUPCSupport Then
                Setup_Sub_UPC_Grid()
            End If
        End If

    End Sub

    Sub Setup_Sub_UPC_Grid()

        Dim STYLE_CODE As String = grdPOTORDR2.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = grdPOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""

        Dim dvw As DataView = DirectCast(grdICTXLSPS.DataSource, DataTable).DefaultView
        dvw.RowFilter = $"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'"

        tabDetails.Tabs("Set").Visible = Line_Has_Sub_UPCs(STYLE_CODE, COLOR_CODE)

        If tabDetails.Tabs("Set").Visible Then
            Dim PO_ORDER_LNO As Int32 = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
            Sort_grdColumns(grdICTXLSPS, "SET_LNO")
            grdICTXLSPS.Text = $"Sub UPCS for PO Line {PO_ORDER_LNO}"
        End If

    End Sub

    Function Line_Has_Sub_UPCs(STYLE_CODE As String, COLOR_CODE As String) As Boolean
        Return dst.Tables("ICTXLSPS").Select($"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'").Count > 0
    End Function
    Private Sub grdPOTORDR2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTORDR2.AfterRowsDeleted
        Display_Totals()

        dst.Tables("POTORDRR").AcceptChanges()
        Dim DT As DataTable = ASCDATA1.SelectDistinct(dst.Tables("POTORDR2").Select(""), New String() {"STYLE_CODE", "COLOR_CODE"})
        DT.PrimaryKey = New DataColumn() {DT.Columns("STYLE_CODE"), DT.Columns("COLOR_CODE")}
        If dst.Tables("POTORDRR").Rows.Count > 0 Then
            For R As Integer = dst.Tables("POTORDRR").Rows.Count - 1 To 0 Step -1
                Dim rowPOTORDRR As DataRow = dst.Tables("POTORDRR").Rows(R)
                If DT.Rows.Count = 0 OrElse DT.Rows.Find(New String() {rowPOTORDRR.Item("STYLE_CODE"), rowPOTORDRR.Item("COLOR_CODE")}) Is Nothing Then
                    rowPOTORDRR.Delete()
                End If
            Next
        End If

        dst.Tables("POTORDR6").AcceptChanges()
        For R As Integer = dst.Tables("POTORDR6").Rows.Count - 1 To 0 Step -1
            Dim rowPOTORDR6 As DataRow = dst.Tables("POTORDR6").Rows(R)
            If dst.Tables("POTORDR2").Rows.Find(New String() {PO_ORDER_NO, rowPOTORDR6.Item("PO_ORDER_LNO")}) Is Nothing Then
                rowPOTORDR6.Delete()
            End If
        Next

    End Sub

    Private Sub grdPOTORDR2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDR2.AfterRowUpdate
        Display_Totals()
        If EntryMode = "N" Then
            grdPOTORDR2.ActiveColScrollRegion.Position = 0 ' .Scroll(UltraWinGrid.ColScrollAction.PageLeft)
        End If
    End Sub

    Private Sub grdPOTORDR2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdPOTORDR2.BeforeCellUpdate

        If e.Cell.Row.IsAddRow Then
            If Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" = "" Or Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Then
                MsgBox("Please provide default values for Ship-By and ETA Date above in the PO Header before entering PO Details",
                       MsgBoxStyle.OkOnly, "Cannot Enter New PO Details")
                e.Cancel = True
                Exit Sub
            End If
        End If



        If e.Cell.Row.Cells("PO_STATUS").Value & "" = "C" And Val(e.Cell.Row.Cells("PO_QTY_REC").Value & "") <> 0 Then
            If New String() {"LAST_OPER", "LAST_DATE", "PO_COST_VCOST_DZ", "PO_COST_VCOST", "PO_COST_OTHER", "PO_COST_OTHER_UN", "PO_COST", "PO_COST_QUOTA", "PO_COST_QUOTA_UN", "PO_COST_COMM", "DFQUOTA", "PO_COST_BUFFER"}.Contains(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("COST_COMPLETE").Value & "" = "1" Then
                    ' If e.Cell.Row.Cells("PO_STATUS").Value & "" = "C" And Val(e.Cell.Row.Cells("PO_QTY_REC").Value & "") <> 0 Then
                    '"You may Not Change a PO Line that Has Been Closed and Received"
                    MsgBox("You may Not Change a PO Line that Has Been Costed",
                           MsgBoxStyle.OkOnly, "This PO Line has already been Costed")
                    e.Cancel = True
                    Exit Sub
                End If
            Else
                MsgBox("You may Not Change a PO Line that Has Been Closed and Received",
                       MsgBoxStyle.OkOnly, "This PO Line has already been Closed and Received")
                e.Cancel = True
                Exit Sub
            End If
        End If

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                If e.NewValue & "" <> "" Then
                    Validate_Style(e.NewValue, True)
                End If

            Case "COLOR_CODE"
                If e.NewValue & "" <> "" Then
                    Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {e.Cell.Row.Cells("STYLE_CODE").Value & "", e.NewValue})
                    If rowICTSTYC1 Is Nothing Then
                        MsgBox("Color " & e.NewValue & " is not valid with Style " & e.Cell.Row.Cells("STYLE_CODE").Value,
                               MsgBoxStyle.OkOnly, "Invalid Value for Color")
                        e.Cancel = True
                    End If
                End If

            Case "PO_QTY_UOM"
                If Val(e.NewValue) <> 1 And Val(e.NewValue) <> 12 Then
                    MsgBox("Value for " & e.Cell.Column.Header.Caption & " Must be either 1 or 12", MsgBoxStyle.OkOnly, "Invalid Entry")
                    e.Cancel = True
                End If

            Case "CMT_NO"
                If e.NewValue & "" <> "" Then
                    Dim rowICTCMTM1 As DataRow = LookUp("ICTCMTM1", e.NewValue)
                    If rowICTCMTM1 Is Nothing Then
                        MsgBox("CMT " & e.NewValue & " is Not on a valid CMT No", MsgBoxStyle.OkOnly, "Invalid Value")
                        e.Cancel = True
                    End If
                End If
        End Select
    End Sub

    Private Sub grdPOTORDR2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTORDR2.BeforeExitEditMode
        If grdPOTORDR2.ActiveCell IsNot Nothing Then
            With grdPOTORDR2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        'If .EditorResolved.Value & "" <> "" AndAlso .EditorResolved.Value <> CStr(.EditorResolved.Value & "").ToUpper Then
                        If .EditorResolved.Value & "" <> "" Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value, .Column.Key)
                        End If
                        'End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdPOTORDR2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTORDR2.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Not grow.IsAddRow Then

                Dim PO_ORDER_LNO As Integer = Val(grow.Cells("PO_ORDER_LNO").Value & "")

                If Val(grow.Cells("PO_QTY_SHP_DZ").Value & "") <> 0 Then
                    MsgBox("Shipments have been entered", MsgBoxStyle.OkOnly, "Deletion Denied")
                    e.Cancel = True
                Else
                    ASCMAIN1.sql = "Select COUNT (*) from ICTTRAN2 " _
                        & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'" _
                        & " and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
                    If Val(ASCDATA1.GetDataValue) <> 0 Then
                        MsgBox("Receipts have been entered", MsgBoxStyle.OkOnly, "Deletion Denied")
                        e.Cancel = True
                    End If
                End If

                If grow.Cells("PO_STATUS").Value <> "O" Then
                    MsgBox("The Selected Line Of This PO Is Not Open", MsgBoxStyle.OkOnly, "Deletion Denied")
                    e.Cancel = True
                End If
            End If
        Next
    End Sub

    Private Sub grdPOTORDR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTORDR2.BeforeRowUpdate

        With grdPOTORDR2
            ' WHY NOT USE VALIDATE STYLE HERE?
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
            If STYLE_CODE = "" Then
                e.Cancel = True
                Exit Sub
            End If
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", New String() {STYLE_CODE})
            If rowICTSTYL1 Is Nothing Then
                MsgBox("Invalid Style: " & STYLE_CODE)
                e.Cancel = True
            Else
                If Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") <= 0 Then
                    MsgBox("Style " & STYLE_CODE & " does not have a Unit Pack")
                    e.Cancel = True
                End If
            End If

            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            If rowICTSTYC1 Is Nothing Then
                MsgBox("Invalid Color: " & COLOR_CODE)
                e.Cancel = True
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PO_DATE_SHIP_BY").Value & "" = "" And Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" <> "" Then
                    e.Row.Cells("PO_DATE_SHIP_BY").Value = Absx1.dteFor("PO_DATE_SHIP_BY").Value
                End If
                If e.Row.Cells("PO_DATE_ETA").Value & "" = "" And Absx1.dteFor("PO_DATE_ETA").Value & "" <> "" Then
                    e.Row.Cells("PO_DATE_ETA").Value = Absx1.dteFor("PO_DATE_ETA").Value
                End If

                If e.Row.IsAddRow Then
                    e.Row.Cells("PO_ORDER_NO").Value = PO_ORDER_NO
                    e.Row.Cells("PO_ORDER_LNO").Value = Val(dst.Tables("POTORDR2").Compute("Max(PO_ORDER_LNO)", "") & "") + 1
                    e.Row.Cells("PO_STATUS").Value = "O"
                    If dst.Tables("POTORDRR").Rows.Find(New String() {PO_ORDER_NO, e.Row.Cells("STYLE_CODE").Value, e.Row.Cells("COLOR_CODE").Value}) Is Nothing Then
                        dst.Tables("POTORDRR").Rows.Add(New String() {PO_ORDER_NO, e.Row.Cells("STYLE_CODE").Value, e.Row.Cells("COLOR_CODE").Value})
                    End If
                    e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                    e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
                Else ' because we now allow change to color as long as not already received
                    If dst.Tables("POTORDRR").Rows.Find(New String() {PO_ORDER_NO, e.Row.Cells("STYLE_CODE").Value, e.Row.Cells("COLOR_CODE").Value}) Is Nothing Then
                        dst.Tables("POTORDRR").Rows.Add(New String() {PO_ORDER_NO, e.Row.Cells("STYLE_CODE").Value, e.Row.Cells("COLOR_CODE").Value})
                    End If

                    If Val(e.Row.Cells("PO_QTY_OPN").Value & "") <> 0 Then
                        e.Row.Cells("PO_STATUS").Value = "O"
                    End If
                End If


                e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP


            End If
        End With
    End Sub

    Private Sub grdPOTORDR2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "CMT_NO"
                Dim sql_where As String = Get_Code_SQL("CMT_NO:0")
                grdClickCellButton(grdPOTORDR2, sql_where)

            Case "COLOR_CODE"
                Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                grdClickCellButton(grdPOTORDR2, sql_where)

            Case Else
                Dim sql_where As String = ""
                grdClickCellButton(grdPOTORDR2, sql_where)
        End Select
    End Sub

    Private Sub grdPOTORDR2_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdPOTORDR2.DoubleClickCell
        'If Not (EntryMode = "V") Then
        '    Select Case e.Cell.Column.Key
        '        Case Is = "PO_COST_MATLS_DZ", "FABRIC_COST", "YARDS_CONSUMED"
        '            If Absx1.optFor("FOB_CMT").Value = "C" Then
        '                If ssdDZGRD.Value = 12 Then
        '                    Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, Val(e.Cell.Row.Cells("PO_ORDER_LNO").Value & "")})
        '                    If rowPOTORDR2 IsNot Nothing Then
        '                        Using F As New POFORDRM
        '                            F.rowPOTORDR2 = rowPOTORDR2
        '                            F.Text = "PO Line " & rowPOTORDR2.Item("PO_ORDER_LNO") & " - Cost Calculator"
        '                            F.ShowDialog()

        '                            If F.ok2update Then
        '                                e.Cell.Row.Cells("YARDS_CONSUMED").Value = F.CONSUMPTION
        '                                e.Cell.Row.Cells("FABRIC_COST").Value = F.FABRIC_COST
        '                                e.Cell.Row.Cells("PO_COST_MATLS_DZ").Value = F.TOTAL_COST
        '                                e.Cell.Row.Update()
        '                                ReCalculate_PO_Cost()
        '                            End If
        '                        End Using
        '                    End If
        '                Else
        '                    MsgBox("You Must Be In Dozens To Use The Cost Calculator", MsgBoxStyle.OkOnly, "Cost Calculator")
        '                End If
        '            End If
        '    End Select
        'End If
    End Sub

    Private Sub grdPOTORDR2_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTORDR2.DoubleClickRow
        If Not e.Row.IsAddRow And Absx1.optFor("FOB_CMT").Value = "I" And Not No_Costs Then
            If InquiryMode Or EntryMode = "V" Then
                Click_Command("View Line")
            Else
                Click_Command("Change Line")
            End If
        End If
    End Sub

    Private Sub grdPOTORDR2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdPOTORDR2.Error
        If grdPOTORDR2.ActiveRow IsNot Nothing Then grdPOTORDR2.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdPOTORDR2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDR2.InitializeRow

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
        Else
            If Not e.Row.IsAddRow Then
                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1.Item("VEND_CODE") & "" <> "" Then
                    If rowICTSTYL1.Item("VEND_CODE") <> Absx1.txtFor("VEND_CODE").Text Then
                        e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
                        e.Row.Cells("STYLE_CODE").ToolTipText = "Style's Default Supplier is " & rowICTSTYL1.Item("VEND_CODE")
                    End If
                End If

                Dim TOTAL_CARTONS As Decimal = Val(e.Row.Cells("TOTAL_CARTONS").Value & "")
                If TOTAL_CARTONS - CLng(TOTAL_CARTONS) <> 0 Then
                    e.Row.Cells("TOTAL_CARTONS").Appearance.BackColor = Drawing.Color.Pink
                    e.Row.Cells("TOTAL_CARTONS").ToolTipText = "Qty Ordered will result in a Partial Carton"
                Else
                    e.Row.Cells("TOTAL_CARTONS").Appearance.BackColor = Drawing.Color.Empty
                    e.Row.Cells("TOTAL_CARTONS").ToolTipText = ""
                End If

            End If
        End If

        'Following code added by RDW 02/09/04
        'If component costs have been changed via the shipment costing screen
        'and the first cost is no long equal to vend + matls + other, to a precision of .0000
        'i change the fore color of the first cost cell to let the user know
        'to hit the 'refresh line' cmd button to see the correct First Cost

        If e.Row.Cells("SHIP_COST_CHANGE_DATE").Value & "" <> "" Then

            ' what about other, comm, and the new field, buffer
            ' ISNT PO_COST RECALCULATED AS THE SUM OF THESE FIELDS ANYWAY?

            If e.Row.Cells("LAST_DATE").Value & "" = "" OrElse
                Format(e.Row.Cells("SHIP_COST_CHANGE_DATE").Value, "yyyyMMdd") >
                Format(e.Row.Cells("LAST_DATE").Value, "yyyyMMdd") Then

                Dim FIRST_COST As Decimal = Val(e.Row.Cells("PO_COST_VCOST").Value & "") _
                                          + Val(e.Row.Cells("PO_COST_MATLS").Value & "") _
                                          + Val(e.Row.Cells("PO_COST_OTHER").Value & "")
                If System.Math.Round(FIRST_COST, 4) <> System.Math.Round(Val(e.Row.Cells("PO_FIRST_COST").Value & ""), 4) Then
                    e.Row.Cells("PO_FIRST_COST").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("PO_FIRST_COST").ToolTipText = ""
                End If

                Dim FIRST_COST_DZ As Decimal = Val(e.Row.Cells("PO_COST_VCOST_DZ").Value & "") _
                                             + Val(e.Row.Cells("PO_COST_MATLS_DZ").Value & "") _
                                             + Val(e.Row.Cells("PO_COST_OTHER_DZ").Value & "")
                If System.Math.Round(FIRST_COST_DZ, 4) <> System.Math.Round(Val(e.Row.Cells("PO_FIRST_COST_DZ").Value & ""), 4) Then
                    e.Row.Cells("PO_FIRST_COST_DZ").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("PO_FIRST_COST_DZ").ToolTipText = ""
                End If
            End If

        End If

        If subUPCSupport Then
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
            If Line_Has_Sub_UPCs(STYLE_CODE, COLOR_CODE) Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Blue
            End If
        End If
    End Sub

    Private Sub grdPOTORDR2_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdPOTORDR2.MouseUp
        If grdPOTORDR2.ActiveCell IsNot Nothing And e.Button = Windows.Forms.MouseButtons.Left Then
            If grdPOTORDR2.ActiveCell.Column.Key = "LINE_CLOSED" Then
                grdPOTORDR2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
            End If
        End If
    End Sub
#End Region

#Region "grdPOTORDR6"

    Private Sub grdPOTORDR6_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR6.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "PO_MESSAGE_CODE"
                Dim rowPOTMESS1 As DataRow = LookUp("POTMESS1", e.Cell.Value & "")
                If rowPOTMESS1 IsNot Nothing Then
                    e.Cell.Row.Cells("PO_MESSAGE_DESC").Value = rowPOTMESS1.Item("PO_MESSAGE_DESC")
                End If
            Case Else

        End Select
    End Sub

    Private Sub grdPOTORDR6_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTORDR6.AfterRowActivate

    End Sub

    Private Sub grdPOTORDR6_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTORDR6.AfterRowsDeleted
        Dim PO_ORDER_LNO As Integer = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
        Dim PO_COST_OTHER_UN As Decimal = Val(dst.Tables("POTORDR6").Compute("SUM(PO_MESSAGE_COST)", "PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)) & "")
        grdPOTORDR2.ActiveRow.Cells("PO_COST_OTHER_UN").Value = PO_COST_OTHER_UN
        grdPOTORDR2.ActiveRow.Update()
        '  grdPOTORDR2.Update()
    End Sub

    Private Sub grdPOTORDR6_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDR6.AfterRowUpdate
        Dim PO_ORDER_LNO As Integer = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
        Dim PO_COST_OTHER_UN As Decimal = Val(dst.Tables("POTORDR6").Compute("SUM(PO_MESSAGE_COST)", "PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)) & "")
        grdPOTORDR2.ActiveRow.Cells("PO_COST_OTHER_UN").Value = PO_COST_OTHER_UN
        grdPOTORDR2.ActiveRow.Update()
        '  grdPOTORDR2.Update()
    End Sub

    Private Sub grdPOTORDR6_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdPOTORDR6.BeforeCellUpdate

        'Select Case e.Cell.Column.Key
        '    Case "PO_MESSAGE_CODE"
        '        If e.NewValue & "" <> "" Then
        '            Validate_Style(e.NewValue, True)
        '        End If

        'End Select
    End Sub

    Private Sub grdPOTORDR6_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTORDR6.BeforeExitEditMode
        If grdPOTORDR6.ActiveCell IsNot Nothing Then
            With grdPOTORDR6.ActiveCell
                Select Case .Column.Key
                    Case "PO_MESSAGE_CODE"
                        If .EditorResolved.Value & "" <> "" Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value, .Column.Key)
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdPOTORDR6_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTORDR6.BeforeRowsDeleted

    End Sub

    Private Sub grdPOTORDR6_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTORDR6.BeforeRowUpdate

        With grdPOTORDR6

            Dim PO_MESSAGE_CODE As String = e.Row.Cells("PO_MESSAGE_CODE").Value & ""
            If PO_MESSAGE_CODE = "" Then
            Else
                Dim rowPOTMESS1 As DataRow = LookUp("POTMESS1", New String() {PO_MESSAGE_CODE})
                If rowPOTMESS1 Is Nothing Then
                    MsgBox("Invalid PO Message Code: " & PO_MESSAGE_CODE)
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.IsAddRow Then
                    e.Row.Cells("PO_ORDER_NO").Value = PO_ORDER_NO
                    e.Row.Cells("PO_ORDER_LNO").Value = grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value
                    e.Row.Cells("PO_ORDER_MLNO").Value = Val(dst.Tables("POTORDR6").Compute("Max(PO_ORDER_MLNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTORDR6_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR6.ClickCellButton
        Select Case e.Cell.Column.Key

            Case "PO_MESSAGE_CODE"
                Dim sql_where As String = ""

                Dim PO_MESSAGE_TYPE As String = e.Cell.Row.Cells("PO_MESSAGE_TYPE").Value & ""
                If PO_MESSAGE_TYPE <> "" Then
                    sql_where = "PO_MESSAGE_TYPE = '" & PO_MESSAGE_TYPE & "'"
                End If

                grdClickCellButton(grdPOTORDR6, sql_where)

            Case "ATTACH"

                If ScreenMode And (EntryMode = "E" Or EntryMode = "N") Then
                Else
                    Exit Sub
                End If
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.InitialDirectory = "c:\"
                    openFileDialog1.Title = "Select a File to Attach to this Record"
                    openFileDialog1.Filter = "jpg files (*.jpg)|*.jpg|png files (*.png)|*.png|bmp files (*.bmp)|*.bmp"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        ENTITY = New Dropped_On_Entity
                        ENTITY.ATTACHMENT_NO = ""
                        ENTITY.COLUMN_NAME = "PO_MESSAGE_ATTACHMENT"
                        ENTITY.TABLE_NAME = "POTORDR6"
                        ENTITY.CODE_VALUE = PO_ORDER_NO
                        Dim Msg As String = Attach_File(openFileDialog1.FileName, , , , True)
                        If Msg <> "" Then
                            MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                        Else
                            Dim PO_MESSAGE_ATTACHMENT As String = ENTITY.ATTACHMENT_NO
                            'Do
                            '    PO_MESSAGE_ATTACHMENT = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")
                            'Loop While My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Attach") & PO_MESSAGE_ATTACHMENT)

                            '  My.Computer.FileSystem.CopyFile(openFileDialog1.FileName, ASCMAIN1.Folders("Attach") & PO_MESSAGE_ATTACHMENT)

                            grdPOTORDR6.ActiveRow.Cells("PO_MESSAGE_ATTACHMENT").Value = PO_MESSAGE_ATTACHMENT
                            grdPOTORDR6.ActiveRow.Update()
                        End If
                    End If
                End Using

            Case "VIEW"
                Dim PO_MESSAGE_ATTACHMENT As String = e.Cell.Row.Cells("PO_MESSAGE_ATTACHMENT").Value & ""
                If PO_MESSAGE_ATTACHMENT <> "" Then
                    Dim rowS() As DataRow = dst.Tables("ASTATTA2").Select("ATTACHMENT_NO = '" & PO_MESSAGE_ATTACHMENT & "'")
                    If rowS.Length = 1 Then
                        Dim ATTACHMENT_EXT As String = rowS(0).Item("ATTACHMENT_EXT").ToUpper
                        ASCMAIN1.Launch_Attachment(PO_MESSAGE_ATTACHMENT, ATTACHMENT_EXT)
                        'Using F As New ASFMSGBF
                        '    Dim x As System.Drawing.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Attach"), PO_MESSAGE_ATTACHMENT, False)
                        '    F.Show_img(imgSTYLE.Image, Me, "Message Attachment")
                        'End Using
                    End If
                End If

            Case Else
                Dim sql_where As String = ""
                grdClickCellButton(grdPOTORDR6, sql_where)
        End Select
    End Sub

    Private Sub grdPOTORDR6_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdPOTORDR6.DoubleClickCell

    End Sub

    Private Sub grdPOTORDR6_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTORDR6.DoubleClickRow

    End Sub

    Private Sub grdPOTORDR6_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdPOTORDR6.Error
        If grdPOTORDR6.ActiveRow IsNot Nothing Then grdPOTORDR6.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdPOTORDR6_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDR6.InitializeRow


    End Sub

#End Region
    Function Validate_Style(STYLE_CODE_candidate As String, Show_Msgs As Boolean) As DataRow

        ' IF THE STYLE CODE IS "", THEN DO NOT PERMIT THE ENTRY OF ANYTHING ELSE

        STYLE_CODE = ""
        Dim Msgs As String = ""

        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_candidate)

        If rowICTSTYL1 Is Nothing Then
            Msgs = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then Msgs &= "Style Status is not Active" & vbCrLf
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then Msgs &= "Style does not have a valid Unit of Measure" & vbCrLf
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then Msgs &= "Style does not have a valid Division Code" & vbCrLf
            If Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") = 0 Then Msgs &= "Styles With Empty Unit Packs Are Not Allowed. Please Change the Masterfile" & vbCrLf
        End If

        COLOR_CODEs.Clear()

        If Msgs = "" Then
            ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE_candidate & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "COLOR_CODE")
                COLOR_CODEs.Add(row.Item("COLOR_CODE"))
            Next
        End If

        If Msgs <> "" And grdPOTORDR2.ActiveRow IsNot Nothing AndAlso grdPOTORDR2.ActiveRow.IsAddRow Then
            If Show_Msgs Then
                MsgBox(Msgs, vbOKOnly, "Style Code Entered is Invalid because ...")
            End If
        Else
            If Msgs = "" Then
                STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE")
            End If
        End If
        Return rowICTSTYL1
    End Function
    Function Validate_Color(COLOR_CODE_candidate As String) As Boolean
        Dim colorIsValid As Boolean = True
        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE_candidate)
        colorIsValid = (rowICTCOLR1 IsNot Nothing)
        Return colorIsValid
    End Function
    Sub Dependent_Calculations(COLUMN_NAME As String)
        ' used in line mode

        Dim UMF As Integer = Val(optUD.Value)
        Dim UMFEA As Integer = 13 - UMF
        Dim SUB_UNIT_PACK_QTY As Integer = Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "")
        If UMF = 12 Then
            If SUB_UNIT_PACK_QTY <> 0 Then
                UMFEA = 12 / SUB_UNIT_PACK_QTY
            End If
        End If

        Select Case COLUMN_NAME

            Case "YIELD_QTY"
                Calc_Yards_by_Color(True)
                'Set Qty Ordered = Production. For Units and Dzns

                If Val(optUD.Value) = 1 Then
                    Absx1.numFor("PO_QTY_ORD_DZ").Value = Val(Absx1.numFor("YIELD_QTY").Value & "") / UMFEA
                    Absx1.numFor("PO_QTY_ORD").Value = Val(Absx1.numFor("YIELD_QTY").Value & "") * UMF
                Else
                    Absx1.numFor("PO_QTY_ORD_DZ").Value = Val(Absx1.numFor("YIELD_QTY").Value & "")
                    Absx1.numFor("PO_QTY_ORD").Value = Val(Absx1.numFor("YIELD_QTY").Value & "") * UMFEA
                End If
                'Set Open Qty = Ordered - Shipped.  For Units and Dzns.
                Dim PO_QTY_OPN As Int64 = Val(Absx1.numFor("PO_QTY_ORD").Value & "") - Val(Absx1.numFor("PO_QTY_SHP").Value & "")
                If PO_QTY_OPN < 0 Or Absx1.chkFor("LINE_CLOSED").Checked Then PO_QTY_OPN = 0
                Absx1.numFor("PO_QTY_OPN").Value = PO_QTY_OPN
                Dim PO_QTY_OPN_DZ As Int64 = Val(Absx1.numFor("PO_QTY_ORD_DZ").Value & "") - Val(Absx1.numFor("PO_QTY_SHP_DZ").Value & "")
                If PO_QTY_OPN_DZ < 0 Or Absx1.chkFor("LINE_CLOSED").Checked Then PO_QTY_OPN_DZ = 0
                Absx1.numFor("PO_QTY_OPN_DZ").Value = PO_QTY_OPN_DZ

            Case "PO_QTY_ORD_DZ"
                Absx1.numFor("PO_QTY_ORD").Value = Val(Absx1.numFor("YIELD_QTY").Value) * UMF

            Case "PO_COST_OTHER", "PO_COST_COMM", "PO_COST_BUFFER"
                'Who knows what to do here?
            Case "PO_COST_QUOTA"
                'Who knows what to do here?

            Case "PO_COST_VCOST_DZ"
                Absx1.numFor("POTORDR2_LINE.PO_COST_VCOST").Value = System.Math.Round((Val(Absx1.numFor("POTORDR2_LINE.PO_COST_VCOST_DZ").Value) / (12 / SUB_UNIT_PACK_QTY)), 6)

            Case "PO_COST_VCOST"
                Absx1.numFor("POTORDR2_LINE.PO_COST_VCOST_DZ").Value = System.Math.Round((Val(Absx1.numFor("POTORDR2_LINE.PO_COST_VCOST").Value) * (12 / SUB_UNIT_PACK_QTY)), 6)

        End Select

        If ScreenMode AndAlso splPOLine.Visible = True And Absx1.optFor("FOB_CMT").Value = "I" Then
            ReCalculate_PO_Cost()
        End If
    End Sub

    Sub Set_grdPOTORDR2_cols_Visibility()
        Dim FOB_CMT As String = Absx1.optFor("FOB_CMT").Value
        With grdPOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_OPN", "PO_QTY_REC"}
                .Columns(COLUMN_NAME).Hidden = (ssdDZGRD.Value <> 1) Or (COLUMN_NAME <> "PO_QTY_ORD" And (EntryMode = "N" Or (EntryMode = "E" And COLUMN_NAME <> "PO_QTY_OPN")))
                .Columns(COLUMN_NAME & "_DZ").Hidden = (ssdDZGRD.Value = 1) Or (COLUMN_NAME <> "PO_QTY_ORD" And (EntryMode = "N" Or (EntryMode = "E" And COLUMN_NAME <> "PO_QTY_OPN")))
            Next

            .Columns("PO_COST_MATLS").Hidden = Not (ssdDZGRD.Value = 1 And FOB_CMT <> "F" And FOB_CMT <> "B") Or No_Costs
            .Columns("PO_COST_MATLS_DZ").Hidden = Not (ssdDZGRD.Value <> 1 And FOB_CMT <> "F" And FOB_CMT <> "B") Or No_Costs
            .Columns("PO_COST").Hidden = Not (ssdDZGRD.Value = 1) Or No_Costs
            .Columns("PO_COST_DZ").Hidden = (ssdDZGRD.Value = 1) Or No_Costs

            .Columns("PO_FIRST_COST_UN").Hidden = Not (ssdDZGRD.Value = 1) Or No_Costs
            .Columns("PO_FIRST_COST_DZ").Hidden = (ssdDZGRD.Value = 1) Or No_Costs

            .Columns("PO_COST_QUOTA").Hidden = (ssdDZGRD.Value = 1) Or No_Costs 'Or (EntryMode = "N")
            .Columns("PO_COST_OTHER").Hidden = (ssdDZGRD.Value = 1) Or No_Costs 'Or (EntryMode = "N")

            .Columns("PO_COST_COMM").Hidden = No_Costs 'Or (EntryMode = "N")
            .Columns("PO_COST_BUFFER").Hidden = No_Costs 'Or (EntryMode = "N")

            .Columns("PO_COST_QUOTA_UN").Hidden = Not (ssdDZGRD.Value = 1) Or No_Costs 'Or (EntryMode = "N")
            .Columns("PO_COST_OTHER_UN").Hidden = Not (ssdDZGRD.Value = 1) Or No_Costs 'Or (EntryMode = "N")
            .Columns("PO_COST_OTHER_UN").MaskInput = "nnn.nnnnnn"

            '.Columns("PO_COST_COMM_DZ").Hidden = (ssdDZGRD.Value = 1) Or No_Costs 'Or (EntryMode = "N")
            '.Columns("PO_COST_BUFFER_DZ").Hidden = (ssdDZGRD.Value = 1) Or No_Costs 'Or (EntryMode = "N")

            .Columns("PO_COST_VCOST").Hidden = (ssdDZGRD.Value <> 1) Or No_Costs
            .Columns("PO_COST_VCOST").MaskInput = "nnn.nnnnnn"
            .Columns("PO_COST_VCOST_DZ").Hidden = (ssdDZGRD.Value = 1) Or No_Costs
            .Columns("YARDS_CONSUMED").Hidden = No_Costs Or (FOB_CMT <> "C")
            .Columns("FABRIC_COST").Hidden = No_Costs Or (FOB_CMT <> "C")
        End With

        Display_Totals()
    End Sub

#Region "grdPOTORDR3"

    Private Sub grdPOTORDR3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDR3.AfterRowUpdate
        Calc_Totals_POTORDR3(True)
    End Sub

    Private Sub grdPOTORDR3_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTORDR3.BeforeRowUpdate
        Dim FABRIC_NO As Integer = Val(e.Row.Cells("FABRIC_NO").Value & "")
        Dim rowPOTORDR4s() As DataRow = dst.Tables("POTORDR4_LINE").Select("FABRIC_NO = " & CStr(FABRIC_NO)) ' (0)
        If rowPOTORDR4s.Length > 0 Then
            Dim rowPOTORDR4 As DataRow = rowPOTORDR4s(0)
            Dim CONSUMPTION_RATE As Decimal = Val(rowPOTORDR4.Item("CONSUMPTION_RATE") & "")
            Dim YIELD_QTY As Int32 = Val(e.Row.Cells("YIELD_QTY").Value & "")

            If Val(optUD.Value) = 1 Then
                e.Row.Cells("YARDS_CONSUMED").Value = (CONSUMPTION_RATE / 12) * YIELD_QTY
            Else
                e.Row.Cells("YARDS_CONSUMED").Value = (CONSUMPTION_RATE) * YIELD_QTY
            End If
        End If

    End Sub

    Private Sub grdPOTORDR3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDR3.InitializeRow
        e.Row.Cells("FABRIC_NO").Appearance.BackColor = Get_Fabric_Color(Val(e.Row.Cells("FABRIC_NO").Value & ""))
    End Sub
#End Region

    Sub Calc_Totals_POTORDR3(update4 As Boolean)

        If Val(Absx1.numFor("POTORDR2_LINE.SUB_UNIT_PACK_QTY").Value & "") = 0 Then
            Absx1.numFor("POTORDR2_LINE.SUB_UNIT_PACK_QTY").Value = 1
        End If

        Dim YIELD_QTY As Int32 = Val(Absx1.numFor("POTORDR2_LINE.YIELD_QTY").Value & "")

        OOBAL = False
        Dim z As String = "(Balanced)"

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR3_LINE").Select("YIELD_QTY <> 0"), New String() {"FABRIC_NO"}).Rows
            Dim FABRIC_NO As Integer = Val(row.Item("FABRIC_NO") & "")
            Dim YIELD_QTY_FABRIC As Int32 = Val(dst.Tables("POTORDR3_LINE").Compute("SUM(YIELD_QTY)", "FABRIC_NO = " & CStr(FABRIC_NO)) & "")

            If update4 Then
                Dim rowPOTORDR4 As DataRow = dst.Tables("POTORDR4_LINE").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO, FABRIC_NO})
                rowPOTORDR4.Item("YIELD_QTY") = YIELD_QTY_FABRIC
                'For Each rowPOTORDR4 As DataRow In dst.Tables("POTORDR4_LINE").Select("FABRIC_NO = " & CStr(FABRIC_NO))
                '    rowPOTORDR4.Item("YIELD_QTY") = YIELD_QTY_FABRIC
                'Next
            End If
            If YIELD_QTY_FABRIC <> YIELD_QTY * Val(optUD.Value) Then
                z = "(Out/Bal)"
                OOBAL = True
            End If
        Next

        Dim TOTAL_COST As Decimal = Val(dst.Tables("POTORDR3_LINE").Compute("SUM(TOTAL_COST)", "") & "")
        Dim YARDS_CONSUMED As Decimal = Val(dst.Tables("POTORDR3_LINE").Compute("SUM(YARDS_CONSUMED)", "") & "")

        grdPOTORDR3.Text = "By Color " & z

        Dim PO_COST_MATLS_DZ As Decimal
        If YIELD_QTY = 0 Then
            PO_COST_MATLS_DZ = 0
        Else
            If Val(optUD.Value) = 1 Then
                PO_COST_MATLS_DZ = System.Math.Round(TOTAL_COST / (YIELD_QTY / 12), 6)
            Else
                PO_COST_MATLS_DZ = System.Math.Round(TOTAL_COST / (YIELD_QTY), 6)
            End If
        End If
        Absx1.numFor("POTORDR2_LINE.PO_COST_MATLS_DZ").Value = PO_COST_MATLS_DZ
        Absx1.numFor("POTORDR2_LINE.PO_COST_MATLS").Value = System.Math.Round((PO_COST_MATLS_DZ / (12 / Val(Absx1.numFor("POTORDR2_LINE.SUB_UNIT_PACK_QTY").Value & ""))), 6)
    End Sub

    Function Get_Fabric_Color(FABRIC_NO As Integer) As System.Drawing.Color
        Select Case 1 + ((FABRIC_NO - 1) Mod 3)
            Case 1
                Return Drawing.Color.LightGray
            Case 2
                Return Drawing.Color.Fuchsia
            Case 3
                Return Drawing.Color.LightGreen
        End Select
    End Function

#Region "grdPOTORDR4"
    Private Sub grdPOTORDR4_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDR4.AfterRowUpdate
        For Each rowPOTORDR3 As DataRow In dst.Tables("POTORDR3_LINE").Select()
            Dim rowPOTORDR4 As DataRow = dst.Tables("POTORDR4_LINE").Rows.Find(New Object() {rowPOTORDR3.Item("PO_ORDER_NO"),
                                                                                        rowPOTORDR3.Item("PO_ORDER_LNO"),
                                                                                        rowPOTORDR3.Item("FABRIC_NO")})
            Dim CONSUMTION_RATE As Decimal = 0
            If rowPOTORDR4 IsNot Nothing Then
                CONSUMTION_RATE = Val(rowPOTORDR4.Item("CONSUMPTION_RATE") & "")
            End If
            rowPOTORDR3.Item("YARDS_CONSUMED") = Val(rowPOTORDR3.Item("YIELD_QTY") & "") * CONSUMTION_RATE / IIf(optUD.Value = 12, 1, 12)
        Next

        Calc_Yards_by_Color(False)
        ReCalculate_PO_Cost()
    End Sub

    Private Sub grdPOTORDR4_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDR4.InitializeRow
        e.Row.Cells("FABRIC_NO").Appearance.BackColor = Get_Fabric_Color(Val(e.Row.Cells("FABRIC_NO").Value & ""))

        If Val(e.Row.Cells("YIELD_QTY").Value & "") <> Val(Absx1.numFor("YIELD_QTY").Value & "") Then
            e.Row.Cells("YIELD_QTY").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("YIELD_QTY").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub
#End Region

#Region "grdPOTORDR7"
    Private Sub grdPOTORDR7_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR7.AfterCellUpdate
        If e.Cell.Column.Key <> "PPK_CODE" Then
            Check_for_PPK(e.Cell.Row)
            'If Val(e.Cell.Row.Cells("STYLES").Value & "") > 1 _
            'Or e.Cell.Row.Cells("CUSTOM_PPK").Value & "" = "1" _
            'Or e.Cell.Row.Cells("CARTON_COMMENTS").Value & "" <> "" Then
            '    e.Cell.Row.Cells("PPK_CODE").Value = Get_Next_PPK_CODE()
            'Else
            '    e.Cell.Row.Cells("PPK_CODE").Value = ""
            'End If
        End If
        ' e.Cell.Row.Update()
    End Sub

    Private Sub grdPOTORDR7_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDR7.AfterRowActivate
        Setup_grdPOTORDR8()
    End Sub

    Private Sub grdPOTORDR7_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDR7.InitializeRow
        If e.Row.Cells("PPK_CODE").Value & "" <> "" Then
            e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.LightGreen
        Else
            e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub
#End Region

    Sub Setup_grdPOTORDR8()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show All Carton Details"), UltraWinToolbars.StateButtonTool)
        grdPOTORDR8.DisplayLayout.Bands(0).Columns("CARTON_NO").Hidden = Not tlb_sbt.Checked

        If grdPOTORDR7.ActiveRow Is Nothing OrElse Not grdPOTORDR7.ActiveRow.IsDataRow Then
            grdPOTORDR8.Visible = False
        Else

            Dim dvw As DataView = DirectCast(grdPOTORDR8.DataSource, DataTable).DefaultView
            Dim CARTON_NO As Integer = Val(grdPOTORDR7.ActiveRow.Cells("CARTON_NO").Value & "")
            ' Dim PO_SHIPMENT_LNO As Integer = Val(grdPOTORDR7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            If tlb_sbt.Checked Then
                dvw.RowFilter = ("PO_ORDER_NO = '" & PO_ORDER_NO & "'")
                grdPOTORDR8.Text = "Carton Configuration by Style/Color for All Carton Types"
            Else
                dvw.RowFilter = ("PO_ORDER_NO = '" & PO_ORDER_NO & "' and CARTON_NO = " & CStr(CARTON_NO))
                grdPOTORDR8.Text = "Carton Configuration by Style/Color for Carton Type " & CStr(CARTON_NO)
            End If
            grdPOTORDR8.Visible = True
        End If
    End Sub

    Sub Check_for_PPK(grow As UltraWinGrid.UltraGridRow)
        If Val(grow.Cells("STYLES").Value & "") > 1 _
        Or grow.Cells("CUSTOM_PPK").Value & "" = "1" _
        Or grow.Cells("CARTON_COMMENTS").Value & "" <> "" Then
            grow.Cells("PPK_CODE").Value = Get_Next_PPK_CODE()
        Else
            grow.Cells("PPK_CODE").Value = ""
        End If
        If grow.DataChanged Then
            grow.Update()
        End If
    End Sub

    Function Get_Next_PPK_CODE() As String
        PPK_CODE_ctr += 1
        Return "TMP" & Format(PPK_CODE_ctr, "0000000")
    End Function

    Sub Create_POTORDRR(STYLE_CODE As String, COLOR_CODE As String)
        Dim rowPOTORDRR As DataRow = dst.Tables("POTORDRR").NewRow
        rowPOTORDRR.Item("PO_ORDER_NO") = PO_ORDER_NO
        rowPOTORDRR.Item("STYLE_CODE") = STYLE_CODE
        rowPOTORDRR.Item("COLOR_CODE") = COLOR_CODE
        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
        rowPOTORDRR.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC") & ""
        dst.Tables("POTORDRR").Rows.Add(rowPOTORDRR)
    End Sub

#Region "grdPOTORDR8"

    Private Sub grdPOTORDR8_AfterCellActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDR8.AfterCellActivate

    End Sub

    Private Sub grdPOTORDR8_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR8.AfterCellUpdate
        e.Cell.Row.Update()
    End Sub

#End Region

#Region "grdPOTORDRR"

    Private Sub grdPOTORDRR_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRR.InitializeRow
        If Val(e.Row.Cells("QTY_VAR").Value & "") <> 0 Then
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub
#End Region

    Private Sub grdPOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRX.InitializeRow
        If e.Row.Cells("PO_STATUS").Value & "" = "C" Then
            e.Row.Cells("PO_STATUS").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("PO_STATUS").Appearance.ForeColor = Drawing.Color.Empty
        End If

        If e.Row.Cells("FOB_CMT").Value & "" = "F" Then
            e.Row.Cells("FOB_CMT").Appearance.BackColor = Drawing.Color.Empty
        Else
            e.Row.Cells("FOB_CMT").Appearance.BackColor = Drawing.Color.LightGreen
        End If

        Dim PO_LINES As Int64 = Val(e.Row.Cells("PO_LINES").Value & "")
        Dim PO_LINES_CONF As Int64 = Val(e.Row.Cells("PO_LINES_CONF").Value & "")
        If PO_LINES_CONF = 0 Then
            e.Row.Cells("PO_LINES").Appearance.BackColor = Drawing.Color.Pink
            e.Row.Cells("PO_LINES").ToolTipText = "PO not Confirmed"
        ElseIf PO_LINES_CONF <> PO_LINES Then
            e.Row.Cells("PO_LINES").Appearance.BackColor = Drawing.Color.Orange
            e.Row.Cells("PO_LINES").ToolTipText = "PO partially Confirmed"
        Else
            e.Row.Cells("PO_LINES").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("PO_LINES").ToolTipText = ""
        End If

        If e.Row.Cells("PO_APPR_PENDING").Value & "" = "0" And e.Row.Cells("PO_APPR_NOTES").Value & "" <> "" Then
            e.Row.Cells("PO_ORDER_NO").Appearance.BackColor = Drawing.Color.Orange
            e.Row.Cells("PO_ORDER_NO").ToolTipText = "Rejected: " & e.Row.Cells("PO_APPR_NOTES").Value
        End If

        If optStatus.Value = "N" Then
            If e.Row.Cells("PO_APPR_BY").Value & "" = "" Then
                e.Row.Cells("PO_APPR_BY").Appearance.BackColor = Drawing.Color.Red
            End If
            If e.Row.Cells("PO_WEB_VISIBLE").Value & "" = "0" Then
                e.Row.Cells("PO_WEB_VISIBLE").Appearance.BackColor = Drawing.Color.Red
            End If
            If e.Row.Cells("PO_XMIT_IND").Value & "" = "0" Then
                e.Row.Cells("PO_XMIT_IND").Appearance.BackColor = Drawing.Color.Red
            End If
        Else
            e.Row.Cells("PO_APPR_BY").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Display_Totals()
        dst.Tables("POTORDRT").Rows.Clear()
        dst.Tables("POTORDRT").Rows.Add(New Object() {1, "Ordd", Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_ORD)", "") & ""),
                                                                 Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")})
        dst.Tables("POTORDRT").Rows.Add(New Object() {2, "Shpd", Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_SHP)", "") & ""),
                                                                 Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_SHP)", "") & "")})
        dst.Tables("POTORDRT").Rows.Add(New Object() {3, "Recd", Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_REC)", "") & ""),
                                                                 Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_REC)", "") & "")})
        dst.Tables("POTORDRT").Rows.Add(New Object() {4, "Open", Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", "") & ""),
                                                                 Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_OPN)", "") & "")})
        Sort_grdColumns(grdPOTORDRT, "LNO", True)
    End Sub

    Sub Set_POTSHIPX()
        If InquiryMode Or EntryMode = "V" Then
            If grdPOTORDR2.ActiveRow Is Nothing OrElse Not grdPOTORDR2.ActiveRow.IsDataRow Then
                grdPOTSHIPX.Visible = False
            Else
                grdPOTSHIPX.Visible = True
                Dim PO_ORDER_LNO As Integer = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
                Dim dvw As DataView = DirectCast(grdPOTSHIPX.DataSource, DataTable).DefaultView
                dvw.RowFilter = "PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
                grdPOTSHIPX.Text = "PO Shipment Details for Line No " & CStr(PO_ORDER_LNO)
            End If
        End If
    End Sub

    Sub LOAD_CMT_TYPE()

        With grdPOTORDR2.DisplayLayout.Bands(0)

            Select Case Absx1.optFor("FOB_CMT").Value
                Case "F", "B"
                    For Each COLUMN_NAME As String In New String() _
                        {"CMT_NO", "STYLE_NOTES", "YIELD_QTY", "YARDS_CONSUMED", "FABRIC_COST"}
                        .Columns(COLUMN_NAME).Hidden = True
                    Next
                    If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                        With .Columns("STYLE_NOTES")
                            .Hidden = False
                            .Width = 60
                            .Header.SetVisiblePosition(grdPOTORDR2.DisplayLayout.Bands(0).Columns("PO_DATE_ETA").Header.VisiblePosition + 1, False)
                        End With

                    End If
                    ssdDZGRD.Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")
                    .Columns("PO_COST_VCOST_DZ").Header.Caption = "Cost"
                    If Not No_Costs Then
                        .Columns("PO_FIRST_COST_DZ").Hidden = False
                        .Columns("PO_COST_DZ").Hidden = False
                        .Columns("PO_COST_VCOST_DZ").Hidden = False
                        .Columns("PO_COST_MATLS_DZ").Hidden = True
                    End If

                Case "C"
                    .Columns("CMT_NO").Hidden = True
                    .Columns("STYLE_NOTES").Hidden = True
                    .Columns("YIELD_QTY").Hidden = True
                    ssdDZGRD.Visible = True
                    .Columns("PO_COST_VCOST_DZ").Header.Caption = "CMT Cost"
                    If Not No_Costs Then
                        .Columns("PO_FIRST_COST_DZ").Hidden = False
                        .Columns("PO_COST_DZ").Hidden = False
                        .Columns("PO_COST_VCOST_DZ").Hidden = False
                        .Columns("PO_COST_MATLS_DZ").Hidden = False
                        .Columns("PO_COST_MATLS_DZ").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("YARDS_CONSUMED").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("YARDS_CONSUMED").Hidden = False
                        .Columns("FABRIC_COST").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("FABRIC_COST").Hidden = False
                    End If

                Case "I"
                    .Columns("CMT_NO").Hidden = False
                    .Columns("STYLE_NOTES").Hidden = False
                    .Columns("YIELD_QTY").Hidden = False
                    .Columns("YARDS_CONSUMED").Hidden = True
                    .Columns("FABRIC_COST").Hidden = True

                    ssdDZGRD.Visible = True
                    .Columns("PO_COST_VCOST_DZ").Header.Caption = "CMT Cost"
                    If Not No_Costs Then
                        .Columns("PO_FIRST_COST_DZ").Hidden = False
                        .Columns("PO_COST_DZ").Hidden = False
                        .Columns("PO_COST_VCOST_DZ").Hidden = False
                        .Columns("PO_COST_MATLS_DZ").Hidden = False
                        .Columns("PO_COST_MATLS_DZ").CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
            End Select
        End With
    End Sub

    Sub Setup_CMT()
        ASCMAIN1.sql = "Select PO_ORDER_NO, PO_ORDER_LNO" & vbCrLf _
            & " from POTORDR2 where STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & " and CMT_NO = '" & CMT_NO & "'" & vbCrLf _
            & " order by PO_DATE_ETA desc, PO_ORDER_NO DESC"
        ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where ROWNUM < 2"
        Dim row_Most_Recent_Use As DataRow = ASCDATA1.GetDataRow

        dst.Tables("POTORDR3_LINE").Rows.Clear()

        ASCMAIN1.sql = "Select * from ICTCMTM4 where CMT_NO = '" & CMT_NO & "'"
        For Each rowICTCMTM4 As DataRow In ASCDATA1.GetDataTable.Select("", "COLOR_NO")
            Dim rowPOTORDR3_LINE As DataRow = dst.Tables("POTORDR3_LINE").NewRow
            rowPOTORDR3_LINE.Item("PO_ORDER_NO") = PO_ORDER_NO
            rowPOTORDR3_LINE.Item("PO_ORDER_LNO") = PO_ORDER_LNO
            rowPOTORDR3_LINE.Item("COLOR_NO") = rowICTCMTM4.Item("COLOR_NO")
            rowPOTORDR3_LINE.Item("COLOR_DESC") = rowICTCMTM4.Item("COLOR_DESC")
            rowPOTORDR3_LINE.Item("FABRIC_NO") = rowICTCMTM4.Item("FABRIC_NO")
            rowPOTORDR3_LINE.Item("COLOR_CODE") = rowICTCMTM4.Item("COLOR_CODE")
            dst.Tables("POTORDR3_LINE").Rows.Add(rowPOTORDR3_LINE)
        Next

        dst.Tables("POTORDR4_LINE").Rows.Clear()

        If row_Most_Recent_Use IsNot Nothing Then
            ASCMAIN1.sql = "Select ICTCMTM5.FABRIC_NO, ICTCMTM5.FABRIC_DESC, POTORDR4.CONSUMPTION_RATE " & vbCrLf _
                & " from ICTCMTM5,POTORDR4" & vbCrLf _
                & " where POTORDR4.PO_ORDER_NO (+) = '" & row_Most_Recent_Use.Item("PO_ORDER_NO") & "'" & vbCrLf _
                & "   and POTORDR4.PO_ORDER_LNO (+) = " & row_Most_Recent_Use.Item("PO_ORDER_LNO") & vbCrLf _
                & "   and POTORDR4.FABRIC_NO (+) = ICTCMTM5.FABRIC_NO" & vbCrLf _
                & "   and ICTCMTM5.CMT_NO = '" & CMT_NO & "'"
            For Each rowICTCMTM5 As DataRow In ASCDATA1.GetDataTable.Select("", "FABRIC_NO")
                Dim rowPOTORDR4_LINE As DataRow = dst.Tables("POTORDR4_LINE").NewRow
                rowPOTORDR4_LINE.Item("PO_ORDER_NO") = PO_ORDER_NO
                rowPOTORDR4_LINE.Item("PO_ORDER_LNO") = PO_ORDER_LNO
                rowPOTORDR4_LINE.Item("FABRIC_NO") = rowICTCMTM5.Item("FABRIC_NO")
                rowPOTORDR4_LINE.Item("FABRIC_DESC") = rowICTCMTM5.Item("FABRIC_DESC")
                rowPOTORDR4_LINE.Item("CONSUMPTION_RATE") = rowICTCMTM5.Item("CONSUMPTION_RATE")
                dst.Tables("POTORDR4_LINE").Rows.Add(rowPOTORDR4_LINE)
            Next
        Else

            ASCMAIN1.sql = "Select * from ICTCMTM5 where CMT_NO = '" & CMT_NO & "'"
            For Each rowICTCMTM5 As DataRow In ASCDATA1.GetDataTable.Select("", "FABRIC_NO")
                Dim rowPOTORDR4_LINE As DataRow = dst.Tables("POTORDR4_LINE").NewRow
                rowPOTORDR4_LINE.Item("PO_ORDER_NO") = PO_ORDER_NO
                rowPOTORDR4_LINE.Item("PO_ORDER_LNO") = PO_ORDER_LNO
                rowPOTORDR4_LINE.Item("FABRIC_NO") = rowICTCMTM5.Item("FABRIC_NO")
                rowPOTORDR4_LINE.Item("FABRIC_DESC") = rowICTCMTM5.Item("FABRIC_DESC")
                dst.Tables("POTORDR4_LINE").Rows.Add(rowPOTORDR4_LINE)
            Next
        End If

        If dst.Tables("POTORDR4_LINE").Rows.Count > 1 Then
            grdPOTORDR3.DisplayLayout.Bands(0).Columns("YIELD_QTY").Hidden = True
            ' ASK GA HOW TO SPECIFY ALL BLACK NO WHITE IF HE CANNOT MANIPULATE PROD QTY
            grdPOTORDR3.DisplayLayout.Bands(0).Columns("YIELD_QTY").Hidden = False
        Else
            grdPOTORDR3.DisplayLayout.Bands(0).Columns("YIELD_QTY").Hidden = False
        End If

        Dim FC_AVG() As Decimal = Calc_FC_AVG(CMT_NO)

        For Each rowPOTORDR3_LINE As DataRow In dst.Tables("POTORDR3_LINE").Select()
            Dim i As Integer = Val(rowPOTORDR3_LINE.Item("COLOR_NO") & "")
            If FC_AVG.Length > i Then
                rowPOTORDR3_LINE.Item("FABRIC_COST") = FC_AVG(i)
            End If
        Next
    End Sub

    Function POOpen(PO_ORDER_NO As String) As String
        ' this routine is not presently in use

        Dim PO_STATUS As String = "C"

        ASCMAIN1.sql = "" _
            & "Select PO_ORDER_NO, PO_ORDER_LNO, SUM(PO_QTY_REC_PO) PO_QTY_REC_PO, SUM(PO_QTY_REC_SHIP) PO_QTY_REC_SHIP FROM (" & vbCrLf _
            & " Select PO_ORDER_NO, PO_ORDER_LNO, SUM(PO_QTY_REC) PO_QTY_REC_PO, 0 PO_QTY_REC_SHIP" & vbCrLf _
            & "  from POTORDR2" & vbCrLf _
            & "  where PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
            & "  group by PO_ORDER_NO, PO_ORDER_LNO " & vbCrLf _
            & "union" & vbCrLf _
            & " Select PO_ORDER_NO, PO_ORDER_LNO, 0 PO_QTY_REC_PO, SUM(PO_QTY_REC) PO_QTY_REC_SHIP " & vbCrLf _
            & "  from POTSHIP3 " & vbCrLf _
            & "  where PO_ORDER_NO = '" & PO_ORDER_NO & "' " & vbCrLf _
            & "  group by BY PO_ORDER_NO, PO_ORDER_LNO" & vbCrLf _
            & ") group by PO_ORDER_NO, PO_ORDER_LNO"

        BeginTrans()

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            If Val(row.Item("PO_QTY_REC_SHIP")) = 0 Then
                ASCMAIN1.sql = "Update POTORDR2 Set PO_STATUS = 'O'" & vbCrLf _
                    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "" & "'" & vbCrLf _
                    & "   and PO_ORDER_LNO = " & row.Item("PO_ORDER_LNO")
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Update POTORDR1 SET PO_STATUS = 'O'" & vbCrLf _
                    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                ASCDATA1.ExecuteSQL()

                If PO_STATUS = "C" Then PO_STATUS = "O"
            Else
                If PO_STATUS = "O" Then PO_STATUS = "P"
            End If
        Next

        CommitTrans()

        Return PO_STATUS
    End Function

    Sub Print_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing POs")

        Dim REPORTFILE As String = "POROPRT1"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and PO_ORDER_NO in ('" & PO_ORDER_NO & "')"})

        Dim REPORT_NO As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("FORM_TYPE", "P")

            Dim RPT As String = REPORTFILE
            Dim PO_PARM_PO_RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
            If PO_PARM_PO_RPT <> "" Then RPT = PO_PARM_PO_RPT

            REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , , , False)

            .Print_Report_End()
        End With


        If ASCMAIN1.CLIENT = "NYA" Then

            If Absx1.txtFor("PO_SPEC_ORDR_NO").Text = "LOBLAW" Then

                REPORTFILE = "POROPRTL"
                If Not REPORTS.ContainsKey(REPORTFILE) Then
                    REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
                    REPORTS(REPORTFILE).Prepare_dst(False, "")
                End If

                'To fill the report's dataset with data from Oracle, 
                ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

                REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and PO_ORDER_NO in ('" & PO_ORDER_NO & "')"})

                REPORT_NO = ""

                With REPORTS(REPORTFILE).clsASCBASE1
                    .Print_Report_Begin()
                    .CR_params.Add("SUBT", "")

                    Dim RPT As String = REPORTFILE
                    .Generate_Report(RPT, "Packing and Shipping Instructions", , True, , , , , False)
                    .Print_Report_End()
                End With

            End If
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Print_POs(PO_ORDER_NO As String, Optional make_pdf As Boolean = False, Optional FILENAME_body As String = "") As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing POs")

        Dim REPORTFILE As String = "POROPRT1"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and PO_ORDER_NO in ('" & PO_ORDER_NO & "')"})


        Dim REPORT_NO As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("FORM_TYPE", "P")

            Dim RPT As String = REPORTFILE
            Dim PO_PARM_PO_RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
            If PO_PARM_PO_RPT <> "" Then RPT = PO_PARM_PO_RPT

            If make_pdf Then
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , "PDF", FILENAME_body, False)
            Else
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , , , False)
            End If
            .Print_Report_End(make_pdf, make_pdf)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return REPORT_NO
    End Function

    Function Print_PO_Packing_Instructions(PO_ORDER_NO As String, Optional make_pdf As Boolean = False, Optional FILENAME_body As String = "") As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Packing and Shipping Instructions")

        Dim REPORTFILE As String = "POROPRTL"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and PO_ORDER_NO in ('" & PO_ORDER_NO & "')"})

        Dim REPORT_NO As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")

            Dim RPT As String = REPORTFILE
            If make_pdf Then
                REPORT_NO = .Generate_Report(RPT, "Packing and Shipping Instructions", , True, , , "PDF", FILENAME_body, False)
            Else
                REPORT_NO = .Generate_Report(RPT, "Packing and Shipping Instructions", , True, , , , , False)
            End If
            .Print_Report_End(make_pdf, make_pdf)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return REPORT_NO
    End Function

    Sub Transmit_POs(PO_ORDER_NO As String, Optional email_to_myself As Boolean = False)

        If Not InquiryMode Then
            If MsgBox("emailing this PO will lock it," _
                      & vbCrLf & " and if it is changed that change will be considered a Revision." _
                      & vbCrLf & vbCrLf & "Continue with this email?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                Exit Sub
            Else
                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                    Exit Sub
                End If
            End If
        End If

        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
        Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
        Dim VEND_CODE As String = rowPOTORDR1.Item("VEND_CODE")

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

        Dim REPORT_NO As String = Print_POs(PO_ORDER_NO, True, PO_ORDER_NO)
        ATTACHMENTs.Add(PO_ORDER_NO & ".pdf", ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".pdf")

        If ASCMAIN1.CLIENT = "NYA" Then
            If Absx1.txtFor("PO_SPEC_ORDR_NO").Text = "LOBLAW" Then
                Dim REPORT_NO2 As String = Print_PO_Packing_Instructions(PO_ORDER_NO, True, PO_ORDER_NO & "_Packing")
                ATTACHMENTs.Add(PO_ORDER_NO & "_Packing.pdf", ASCMAIN1.Folders("Temp") & PO_ORDER_NO & "_Packing.pdf")
            End If
        End If


        Dim SUBJECT As String = ""
        Dim PFX As String = ""
        If ASCMAIN1.CLIENT = "RGI" Then PFX = "Regency  "
        SUBJECT = PFX & "PO " & PO_ORDER_NO

        'If optStatus.Value = "O" Then
        '    SUBJECT &= " - Re-Transmit"
        'End If

        Dim SEND_CC_to_USER_ID As Boolean = True

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")
        If email_to_myself Then
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
            SEND_CC_to_USER_ID = False
        Else
            EMAIL_ADDRESSs.Add(rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "", rowAPTVEND1.Item("VEND_PURCH_CONTACT") & "")
        End If

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, "PO", False, SEND_CC_to_USER_ID, VEND_CODE, rowAPTVEND1.Item("VEND_NAME"), "Supplier")

        If SEND_NO <> "" Then
            If Not InquiryMode Then
                ASCMAIN1.sql = "Update POTORDR1 " & vbCrLf _
                    & " Set PO_PRINTED_IND = '1', PO_DATE_PRINTED = SYSDATE" & vbCrLf _
                    & ", PO_XMIT_IND = '1', PO_XMIT_BY = '" & ASCMAIN1.USER_ID & "', PO_XMIT_DATE = SYSDATE, PO_XMIT_XNO = '" & XNO & "'" & vbCrLf _
                    & " where (PO_ORDER_NO) in ('" & PO_ORDER_NO & "')"
                ASCDATA1.ExecuteSQL()
                'Dim rowPOTORDRX As DataRow = dst.Tables("POTORDRX").Rows.Find(PO_ORDER_NO)
                'rowPOTORDRX.Delete()
            End If

            '   Dim REVISION_FILENAME As String = ASCMAIN1.Folders("Archive") & "PO\" & PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV) & ".PDF")
            If Not My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Archive") & "PO\" & PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV) & ".PDF") Then
                My.Computer.FileSystem.CopyFile(
                    ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".PDF",
                    ASCMAIN1.Folders("Archive") & "PO\" & PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV) & ".PDF")

                ' RDW Portal notes
                ' 4)	Save a copy of the printed PO in a folder visible by the web _____. 
                ' I'd like to revise my initial instruction to put in printed PO on the S:\ drive. 
                ' If you can also put it in \\192.168.170.103\c$\NYA_Portal\purchaseOrders that would be great. 
                ' It's possible to map the virtual drive to a network share 
                ' but I need to turn off default pass-thru authentication 
                ' and use an account will elevated privileges to access it. 
                ' If the account I use is removed or demoted (this happened @ COSFF) 
                ' access to the virtual directory will be lost. 
                ' Probably safer and more reliable for you to duplex the copy. 

                My.Computer.FileSystem.CopyFile(
                    ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".PDF",
                    ROWs("POTPARM1").Item("PO_PARM_PO_PDF_FOLDER") & PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV) & ".PDF")
            End If


            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'PO-XMIT','PO Transmitted', '" & SEND_NO & "'" _
                & " from POTORDR1 " & vbCrLf _
                & " where (PO_ORDER_NO) in ('" & PO_ORDER_NO & "')"
            ASCDATA1.ExecuteSQL()

            Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
            rowTATEVNT1.Item("TABLE_NAME") = "POTORDR1"
            rowTATEVNT1.Item("TABLE_KEY") = PO_ORDER_NO
            rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
            rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowTATEVNT1.Item("EVENT_TYPE") = "PO-XMIT"
            rowTATEVNT1.Item("EVENT_DESC") = "PO Transmitted"
            rowTATEVNT1.Item("EVENT_KEY") = SEND_NO
            dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)

            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

            If Not InquiryMode Then

                ASCMAIN1.sql = "Select POTORDR1.PO_ORDER_NO, NVL(POTORDR1.PO_HDR_CTR_REV,0) CTR1" & vbCrLf _
                    & " , MAX (NVL(POTORDRH.PO_HDR_CTR_REV,-1)) CTRH" & vbCrLf _
                    & "  from POTORDR1,POTORDRH" & vbCrLf _
                    & " where POTORDRH.PO_ORDER_NO (+) = POTORDR1.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                    & " group by POTORDR1.PO_ORDER_NO, NVL(POTORDR1.PO_HDR_CTR_REV,0)"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing AndAlso Val(row.Item("CTR1") & "") > Val(row.Item("CTRH") & "") Then

                    ASCMAIN1.sql = "Insert into POTORDRZ (PO_ORDER_NO,PO_HDR_CTR_REV,PO_ORDER_LNO" & vbCrLf _
                        & ",STYLE_CODE,COLOR_CODE,PO_QTY_ORD,PO_COST,PO_DATE_SHIP_BY,PO_STATUS,CARTON_PACK_QTY)" & vbCrLf _
                        & " Select POTORDR2.PO_ORDER_NO, NVL(POTORDR1.PO_HDR_CTR_REV,0), POTORDR2.PO_ORDER_LNO" & vbCrLf _
                        & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR2.PO_QTY_ORD" & vbCrLf _
                        & ", POTORDR2.PO_COST, POTORDR2.PO_DATE_SHIP_BY, POTORDR2.PO_STATUS, POTORDR2.CARTON_PACK_QTY" & vbCrLf _
                        & " from POTORDR1,POTORDR2" & vbCrLf _
                        & " where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
                        & "   and POTORDR1.PO_ORDER_NO in ('" & PO_ORDER_NO & "')"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "Insert into POTORDRH (PO_ORDER_NO,PO_HDR_CTR_REV,PO_REVISION_NOTE,INIT_OPER,INIT_DATE,LAST_OPER,LAST_DATE)" & vbCrLf _
                        & " Select PO_ORDER_NO, NVL(PO_HDR_CTR_REV,0), DECODE(NVL(PO_HDR_CTR_REV,0),0,'Original',PO_REVISION_NOTE), LAST_OPER, LAST_DATE, LAST_OPER, LAST_DATE" & vbCrLf _
                        & " from POTORDR1" & vbCrLf _
                        & " where POTORDR1.PO_ORDER_NO in ('" & PO_ORDER_NO & "')"
                    ASCDATA1.ExecuteSQL()

                    Fill_Records("POTORDRH", PO_ORDER_NO)
                    Sort_grdColumns(grdPOTORDRH, "PO_HDR_CTR_REV")

                End If

            End If

            If rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "" = "" Then
                Dim rowTATSEND1 As DataRow = LookUp("TATSEND1", SEND_NO)
                Dim SEND_TO As String = rowTATSEND1.Item("SEND_TO") & ""
                If SEND_TO <> "" And Not SEND_TO.Contains(";") Then
                    ASCMAIN1.sql = "Update APTVEND1 Set VEND_PURCH_EMAIL = :PARM1 where VEND_CODE = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {rowTATSEND1.Item("SEND_TO"), VEND_CODE})
                End If
            End If
        End If

        If Not InquiryMode Then
            ASCMAIN1.MultiTask_Release()
        End If
    End Sub


    Function Check_Shipped(PO_ORDER_NO As String, Optional PO_ORDER_LNO As Integer = 0) As Boolean
        Dim SQLW As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        If PO_ORDER_LNO <> 0 Then
            SQLW &= " and PO_ORDER_LNO = " & PO_ORDER_LNO
        End If
        Return (dst.Tables("POTSHIPX").Select(SQLW).Length > 0)
    End Function

#Region "PO Line Maintenance"

    Sub Duplicate_Line()

        If grdPOTORDR2.ActiveRow IsNot Nothing AndAlso Not grdPOTORDR2.ActiveRow.IsAddRow Then
            With grdPOTORDR2.ActiveRow
                Dim PO_QTY_ORD As Int32 = 0
                Dim PO_QTY_ORD_DZ As Int32 = 0

                Dim LNO As Integer = Val(.Cells("PO_ORDER_LNO").Value & "")
                'If Absx1.optFor("FOB_CMT").Value <> "I" Then

                '    If MsgBox("Would You Like To Deduct From The Original?", MsgBoxStyle.YesNo, "Deduction") = MsgBoxResult.Yes Then
                '        Dim Result As Integer = ASCMAIN1.Get_num_from_User _
                '                                ("New Line Order Qty", "Please Enter The Order Qty For The New Line")
                '        ' what about cancel?
                '        If ssdDZGRD.Value = 1 Then
                '            PO_QTY_ORD = Result
                '            PO_QTY_ORD_DZ = System.Math.Round(Result / 12 * (Val(.Cells("SUB_UNIT_PACK_QTY").Value & "")), 6)
                '        Else
                '            PO_QTY_ORD = System.Math.Round(Result * 12 / (Val(.Cells("SUB_UNIT_PACK_QTY").Value & "")), 6)
                '            PO_QTY_ORD_DZ = Result
                '        End If
                '        If Val(.Cells("PO_QTY_OPN").Value & "") - Val(PO_QTY_ORD) < 0 Then
                '            MsgBox("You Can Not Deduct More Than The Qty Open!", vbCritical, "Deduction Error")
                '            Exit Sub
                '        Else
                '            .Cells("PO_QTY_ORD").Value = Val(.Cells("PO_QTY_ORD").Value & "") - PO_QTY_ORD
                '            .Cells("PO_QTY_ORD_DZ").Value = Val(.Cells("PO_QTY_ORD_DZ").Value & "") - PO_QTY_ORD_DZ
                '            .Cells("PO_QTY_OPN").Value = Val(.Cells("PO_QTY_OPN").Value & "") - PO_QTY_ORD
                '            .Cells("PO_QTY_OPN_DZ").Value = Val(.Cells("PO_QTY_OPN_DZ").Value & "") - PO_QTY_ORD_DZ
                '        End If

                '    Else
                '        PO_QTY_ORD = Val(.Cells("PO_QTY_ORD").Value & "")
                '        PO_QTY_ORD_DZ = Val(.Cells("PO_QTY_ORD_DZ").Value & "")
                '    End If
                'Else
                PO_QTY_ORD = Val(.Cells("PO_QTY_ORD").Value & "")
                PO_QTY_ORD_DZ = Val(.Cells("PO_QTY_ORD_DZ").Value & "")
                'End If


                Dim NEW_LINE As Integer = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", "") & "") + 1

                For Each TABLE_NAME As String In New String() {"POTORDR2", "POTORDR3", "POTORDR4"}
                    Dim sqlw As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(LNO)
                    For Each row_orig As DataRow In dst.Tables(TABLE_NAME).Select(sqlw)
                        Dim row_copy As DataRow = dst.Tables(TABLE_NAME).NewRow
                        For Each DCOL As DataColumn In dst.Tables(TABLE_NAME).Columns
                            If Not DCOL.ReadOnly Then
                                row_copy.Item(DCOL.ColumnName) = row_orig.Item(DCOL.ColumnName)
                            End If
                        Next
                        If TABLE_NAME = "POTORDR2" Then
                            row_copy.Item("PO_QTY_SHP") = 0
                            row_copy.Item("PO_QTY_SHP_DZ") = 0
                            row_copy.Item("PO_QTY_REC") = 0
                            row_copy.Item("PO_QTY_REC_DZ") = 0
                            row_copy.Item("PO_QTY_OPN") = row_copy.Item("PO_QTY_ORD")
                            row_copy.Item("PO_QTY_OPN_DZ") = row_copy.Item("PO_QTY_ORD_DZ")
                        End If
                        row_copy.Item("PO_ORDER_LNO") = NEW_LINE
                        dst.Tables(TABLE_NAME).Rows.Add(row_copy)
                    Next
                Next

                Display_Totals()
            End With
        End If
    End Sub

    Sub Clear_Line()
        EnforceConstraints(False)
        dst.Tables("POTORDR2_LINE").Rows.Clear()
        dst.Tables("POTORDR3_LINE").Rows.Clear()
        dst.Tables("POTORDR4_LINE").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Sub Cancel_Changes()
        Clear_Line()
        Setup_POLine(False)
    End Sub

    Sub Done_with_Line()
        ' UltraExplorerBar1.Groups("Cost Calculation").Visible = False
        'UltraExplorerBar1.Groups("Line Item Commands").Items("Update Line").Visible = True
        'UltraExplorerBar1.Groups("Line Item Commands").Items("Return to Summary").Visible = True
    End Sub

    Sub Setup_POLine(tf As Boolean)

        tabPOTORDR2.Tabs("Cartonization").Visible = Not tf
        tabPOTORDR2.Tabs("Audit Trails").Visible = Not tf And (EntryMode <> "N")
        tabPOTORDR2.Tabs("Revision History").Visible = Not tf And (EntryMode <> "N")
        tabPOTORDR2.Tabs("Line Item Details").Visible = tf
        tabPOTORDR2.Tabs("PO Details").Visible = Not tf

        If tf Then
            tabPOTORDR2.SelectedTab = tabPOTORDR2.Tabs("Line Item Details")
            cmdProrate.Visible = Not (EntryMode = "V" Or InquiryMode)
        Else
            tabPOTORDR2.SelectedTab = tabPOTORDR2.Tabs("PO Details")
        End If

        lblLastCost.Visible = False

        With UltraExplorerBar1.Groups("Line Item Commands")
            .Items("Add Line").Settings.Enabled = IIf(Not tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("View Line").Settings.Enabled = IIf(Not tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("Change Line").Settings.Enabled = IIf(Not tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("Update Line").Settings.Enabled = IIf(tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("Delete Line").Settings.Enabled = IIf(tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("Cancel Changes").Settings.Enabled = IIf(tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("Return to Summary").Settings.Enabled = IIf(tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("Refresh Line").Settings.Enabled = IIf(tf, DefaultableBoolean.True, DefaultableBoolean.False)
            .Items("Duplicate Line").Settings.Enabled = IIf(Not tf, DefaultableBoolean.True, DefaultableBoolean.False)
            '    .Items("Use Calculated Cost").Settings.Enabled = IIf(tf, DefaultableBoolean.True, DefaultableBoolean.False)

            .Items("Add Line").Visible = Not (InquiryMode Or EntryMode = "V")
            .Items("View Line").Visible = (InquiryMode Or EntryMode = "V")
            .Items("Change Line").Visible = Not (InquiryMode Or EntryMode = "V")
            .Items("Update Line").Visible = Not (InquiryMode Or EntryMode = "V")
            .Items("Delete Line").Visible = Not (InquiryMode Or EntryMode = "V")
            .Items("Cancel Changes").Visible = Not (InquiryMode Or EntryMode = "V")
            .Items("Return to Summary").Visible = (InquiryMode Or EntryMode = "V")
            .Items("Refresh Line").Visible = Not (InquiryMode Or EntryMode = "V")
            .Items("Duplicate Line").Visible = Not (InquiryMode Or EntryMode = "V")
            '    .Items("Use Calculated Cost").Visible = Not (InquiryMode Or EntryMode = "V")
        End With

        If tf Then
            tabPOTORDR2.Tabs("Line Item Details").Text = "Line Item Details for PO Line " & CStr(PO_ORDER_LNO)
        End If

        If tf Then
            optUD.Enabled = False
            If Val(Absx1.numFor("POTORDR2_LINE.PO_COST_BUFFER").Value & "") = 0 Then
                cmdCostInc.Text = "2%"
            Else
                cmdCostInc.Text = "0%"
            End If
            If Val(Absx1.numFor("POTORDR2_LINE.PO_COST_COMM").Value & "") = 0 Then
                cmdPercent.Text = "3%"
            Else
                cmdPercent.Text = "0%"
            End If
            Validate_Style(Absx1.txtFor("POTORDR2_LINE.STYLE_CODE").Text, False)
        End If
    End Sub

    Sub Add_Line()

        STYLE_CODE = ""
        CMT_NO = ""
        Clear_Line()

        rowPOTORDR2 = dst.Tables("POTORDR2_LINE").NewRow
        rowPOTORDR2.Item("PO_ORDER_NO") = PO_ORDER_NO
        PO_ORDER_LNO = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", "") & "") + 1
        rowPOTORDR2.Item("PO_ORDER_LNO") = PO_ORDER_LNO
        rowPOTORDR2.Item("PO_STATUS") = "O"

        dst.Tables("POTORDR2_LINE").Rows.Add(rowPOTORDR2)

        'Added only for "AT" for Marilyn 9/13/?? - WR.
        'Now set to 5% regardless of vendor per Anna - 9/16/10
        'If txtCode(0).Text = "AT" Then
        'datPOWORDR2.Recordset.Fields("PO_COST_COMM").Value = 5
        If Absx1.txtFor("VEND_CODE").Text = "AT" Then
            rowPOTORDR2.Item("PO_COST_BUFFER") = 0
            'Changed from 4 to 3 Per Anna 12/31/14 - WR.
            'Changed from 3 to 2.5 Per Anna 1/20/17 - WR.
            'Changed from 2.5 to 2.0 Per Anna 7/7/19 - WR.
            'Changed from 2.0 to 2.5 Per Anna 7/17/19 - WR.
            'Changed from 2.5 to 2.0 Per Anna 1/26/21 - WR.
            'Changed from 2.0 to 0.0 Per Anna 8/16/24 - WR.
            rowPOTORDR2.Item("PO_COST_COMM") = 0.0
        Else
            rowPOTORDR2.Item("PO_COST_BUFFER") = 2
            rowPOTORDR2.Item("PO_COST_COMM") = 0
        End If
        'End If

        'Only on new records for this.
        rowPOTORDR2.Item("PO_DATE_SHIP_BY") = Absx1.dteFor("PO_DATE_SHIP_BY").Value
        rowPOTORDR2.Item("PO_DATE_ETA") = Absx1.dteFor("PO_DATE_ETA").Value
        optUD.Value = Format(Val(ssdDZGRD.Value), "00")

        Absx1.chkFor("LINE_CLOSED").Visible = False
        'Copy_Line_to_LINE_Tables() ' nothing to copy
        Setup_POLine(True)
    End Sub

    Sub Change_Line()
        ssdDZGRD.Value = 1
        optUD.Value = Format(Val(ssdDZGRD.Value), "00")

        With grdPOTORDR2.ActiveRow
            STYLE_CODE = .Cells("STYLE_CODE").Value & ""
            CMT_NO = .Cells("CMT_NO").Value & ""
            PO_ORDER_LNO = Val(.Cells("PO_ORDER_LNO").Value & "")
            Copy_Line_to_LINE_Tables()
            Calc_Totals_POTORDR3(True)
            ReCalculate_PO_Cost()
            'Dim row As DataRow = dst.Tables("POTORDR2_LINE").Rows(0)
            'Absx1.numFor("POTORDR2_LINE.PO_COST_COMM").Value = .Cells("PO_COST_COMM").Value
            'Absx1.numFor("POTORDR2_LINE.PO_COST_BUFFER").Value = .Cells("PO_COST_BUFFER").Value
        End With

        Setup_POLine(True)
    End Sub

    Sub Update_Line()
        Save_LNO()
        Setup_POLine(False)
    End Sub

    Private Sub Refresh_Line()
        Dim YIELD_QTYs As New Dictionary(Of Integer, Integer)
        For Each rowPOTORDR3 As DataRow In dst.Tables("POTORDR3_LINE").Select("", "COLOR_NO")
            Dim COLOR_NO As Int32 = Val(rowPOTORDR3.Item("COLOR_NO") & "")
            YIELD_QTYs.Add(COLOR_NO, Val(rowPOTORDR3.Item("YIELD_QTY") & ""))
        Next
        Setup_CMT()
        Calc_Totals_POTORDR3(True)
        Calc_Yards_by_Color(False)
        For Each rowPOTORDR3 As DataRow In dst.Tables("POTORDR3_LINE").Select("", "COLOR_NO")
            Dim COLOR_NO As Int32 = Val(rowPOTORDR3.Item("COLOR_NO") & "")
            rowPOTORDR3.Item("YIELD_QTY") = YIELD_QTYs(COLOR_NO)
        Next

        Dim UMF As Integer = Val(optUD.Value)
        Dim UMFEA As Integer = 13 - UMF
        If UMF = 12 Then
            If Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "") <> 0 Then
                UMFEA = UMFEA * Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "")
            End If
        End If
        Absx1.numFor("POTORDR2_LINE.PO_QTY_ORD").Value = Val(Absx1.numFor("POTORDR2_LINE.YIELD_QTY").Value & "") * UMF
        Absx1.numFor("POTORDR2_LINE.PO_QTY_OPN").Value = Val(Absx1.numFor("POTORDR2_LINE.PO_QTY_ORD").Value & "") - Val(Absx1.numFor("POTORDR2_LINE.PO_QTY_SHP").Value & "")
        'End If
        ReCalculate_PO_Cost()
    End Sub

    Sub Save_LNO()
        Dim sqlw As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
        For Each TABLE_NAME As String In New String() {"POTORDR2", "POTORDR3", "POTORDR4"}
            If TABLE_NAME = "POTORDR2" Then
                Synch_TABLE_NAME("POTORDR2_LINE")
            Else
                ASCDATA1.DeleteRows(dst.Tables(TABLE_NAME), sqlw)
            End If
            For Each row_LINE As DataRow In dst.Tables(TABLE_NAME & "_LINE").Select()
                Dim row As DataRow
                If TABLE_NAME <> "POTORDR2" Then
                    row = dst.Tables(TABLE_NAME).NewRow
                Else
                    row = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO}) ' rowPOTORDR2
                    If row Is Nothing Then
                        row = dst.Tables(TABLE_NAME).NewRow
                    End If

                    If dst.Tables("POTORDRR").Rows.Find(New String() {PO_ORDER_NO, row_LINE.Item("STYLE_CODE"), row_LINE.Item("COLOR_CODE")}) Is Nothing Then
                        dst.Tables("POTORDRR").Rows.Add(New String() {PO_ORDER_NO, row_LINE.Item("STYLE_CODE"), row_LINE.Item("COLOR_CODE")})
                    End If


                End If
                For i As Integer = 0 To row_LINE.Table.Columns.Count - 1
                    If dst.Tables(TABLE_NAME).Columns(i).ReadOnly Then
                    Else
                        row.Item(i) = row_LINE.Item(i)
                    End If
                Next
                If TABLE_NAME = "POTORDR2" Then
                    ' SEE grdPOTORDR2_AfterCellUpdate
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", row.Item("STYLE_CODE"))
                    row.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                    row.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
                    row.Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
                    row.Item("CASE_CUBE") = rowICTSTYL1.Item("CASE_CUBE")
                End If

                If TABLE_NAME <> "POTORDR2" Or row.RowState = DataRowState.Detached Then dst.Tables(TABLE_NAME).Rows.Add(row)
            Next
        Next
        ReCalculate_PO_Cost()
        Display_Totals()
    End Sub

    Sub Copy_Line_to_LINE_Tables()
        EnforceConstraints(False)
        Dim sqlw As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
        For Each TABLE_NAME As String In New String() {"POTORDR2", "POTORDR3", "POTORDR4"}
            dst.Tables(TABLE_NAME & "_LINE").Rows.Clear()
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlw)
                If TABLE_NAME = "POTORDR2" Then
                    If optUD.Value = 12 Then
                        row.Item("YIELD_QTY") = row.Item("PO_QTY_ORD_DZ")
                    Else
                        row.Item("YIELD_QTY") = row.Item("PO_QTY_ORD")
                    End If
                End If
                Dim row_LINE As DataRow = dst.Tables(TABLE_NAME & "_LINE").NewRow
                For i As Integer = 0 To row_LINE.Table.Columns.Count - 1
                    If dst.Tables(TABLE_NAME).Columns(i).ReadOnly Then
                    Else
                        row_LINE.Item(i) = row.Item(i)
                    End If
                Next

                dst.Tables(TABLE_NAME & "_LINE").Rows.Add(row_LINE)
            Next
        Next
        EnforceConstraints(True)
    End Sub

    Sub ReCalculate_PO_Cost()
        Dim PO_COST_VCOST As Decimal
        Dim PO_COST_MATLS As Decimal
        Dim PO_COST_OTHER As Decimal
        Dim PO_COST_OTHER_UN As Decimal = 0
        Dim PO_COST_QUOTA As Decimal = 0
        Dim PO_COST_QUOTA_UN As Decimal = 0
        Dim PO_COST_COMM As Decimal
        Dim PO_COST_COMM_UN As Decimal = 0
        Dim PO_COST As Decimal
        Dim PO_COST_SUBTOTAL As Decimal = 0

        Dim SUB_UNIT_PACK_QTY As Integer
        Dim UOM As Integer

        If grdPOTORDR2.ActiveRow Is Nothing Then Exit Sub

        ' Note that COMM, QUOTA, OTHER, and BUFFER are all per Dz Pcs

        Select Case Absx1.optFor("FOB_CMT").Value
            Case Is = "C", "F", "I", "B"

                With grdPOTORDR2.ActiveRow

                    SUB_UNIT_PACK_QTY = Val(.Cells("SUB_UNIT_PACK_QTY").Value & "")
                    UOM = Val(ssdDZGRD.Value)

                    If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                        PO_COST_VCOST = System.Math.Round(Val(.Cells("PO_COST_VCOST_DZ").Value & "") / 12 * SUB_UNIT_PACK_QTY, 6)
                        '.Cells("PO_COST_VCOST").Value = PO_COST_VCOST
                        PO_COST_MATLS = System.Math.Round(Val(.Cells("PO_COST_MATLS_DZ").Value & "") / 12 * SUB_UNIT_PACK_QTY, 6)
                        '.Cells("PO_COST_MATLS").Value = PO_COST_MATLS
                    Else
                        If UOM = 1 Then
                            PO_COST_VCOST = System.Math.Round(Val(.Cells("PO_COST_VCOST").Value & ""), 6)
                            PO_COST_MATLS = System.Math.Round(Val(.Cells("PO_COST_MATLS").Value & ""), 6)
                        Else
                            PO_COST_VCOST = System.Math.Round(Val(.Cells("PO_COST_VCOST_DZ").Value & "") / 12 * SUB_UNIT_PACK_QTY, 6)
                            PO_COST_MATLS = System.Math.Round(Val(.Cells("PO_COST_MATLS_DZ").Value & "") / 12 * SUB_UNIT_PACK_QTY, 6)
                        End If
                    End If

                    PO_COST_OTHER_UN = Val(.Cells("PO_COST_OTHER_UN").Value & "")
                    PO_COST_QUOTA_UN = Val(.Cells("PO_COST_QUOTA_UN").Value & "")
                    PO_COST_OTHER_UN = System.Math.Round(Val(.Cells("PO_COST_OTHER").Value & "") / (12 / SUB_UNIT_PACK_QTY), 6)
                    PO_COST_QUOTA_UN = System.Math.Round(Val(.Cells("PO_COST_QUOTA").Value & "") / (12 / SUB_UNIT_PACK_QTY), 6)

                    PO_COST_COMM = Val(.Cells("PO_COST_COMM").Value & "")
                    PO_COST_SUBTOTAL = PO_COST_VCOST + PO_COST_MATLS + PO_COST_OTHER_UN ' + PO_COST_QUOTA_PER_UNIT Per Anna this will not be used in any calculations until after duty is applied - 11/24.
                    PO_COST_COMM_UN = Val(.Cells("PO_COST_COMM_UN").Value & "")
                    PO_COST_COMM_UN = System.Math.Round(PO_COST_SUBTOTAL * Val(.Cells("PO_COST_COMM").Value & "") / 100, 6)

                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        ' UPDATE 09/12/16 - WJZ
                        ' DUTY IS CALCULATED FROM THE COMPONENTS OF THE COST IN POFSHIP1
                        ' IN A LINE OF CODE THAT LOOKS LIKE THIS:
                        ' Dim PO_COST_DUTY As Decimal = System.Math.Round(a * (b + c - PO_COST_COMM_paid_by_supplier), 6)
                        ' SO WHAT WE DO TO PO_COST here SHOULD ONLY IMPACT AP INTEGRATION
                        ' AS PER ANNA EMAIL 09/09, QUOTA COST NEEDS TO BE PULLED INTO THE PO COST BECAUSE MARIA NEEDS IT TO MATCH AGAINST INVOICES

                        .Cells("PO_COST").Value = PO_COST_SUBTOTAL + PO_COST_COMM_UN + PO_COST_QUOTA_UN
                    Else
                        .Cells("PO_COST").Value = PO_COST_SUBTOTAL
                    End If
                    '.Cells("PO_COST").Appearance.ForeColor = Drawing.Color.Black

                    'If Val(UOM) = 1 Then
                    '    .Cells("PO_COST_DZ").Value = Val(.Cells("PO_COST").Value & "") * 12
                    'Else
                    '    .Cells("PO_COST_DZ").Value = System.Math.Round(Val(.Cells("PO_COST").Value & "") / (Val(SUB_UNIT_PACK_QTY) / 12), 6)
                    'End If
                    '.Cells("PO_COST_DZ").Appearance.ForeColor = Drawing.Color.Black
                End With

            Case Is = "Ilater"
                'loads of changes made above - need to be repeated here

                If Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "") = 0 Then
                    SUB_UNIT_PACK_QTY = 1
                Else
                    SUB_UNIT_PACK_QTY = Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "")
                End If

                If Val(optUD.Value) = 0 Then
                    UOM = 12
                Else
                    UOM = Val(optUD.Value)
                End If

                PO_COST_VCOST = System.Math.Round((Val(Absx1.numFor("PO_COST_VCOST_DZ").Value & "") / 12 * Val(SUB_UNIT_PACK_QTY)), 6)
                Absx1.numFor("PO_COST_VCOST").Value = PO_COST_VCOST

                PO_COST_MATLS = System.Math.Round((Val(Absx1.numFor("PO_COST_MATLS_DZ").Value & "") / 12 * Val(SUB_UNIT_PACK_QTY)), 6)
                Absx1.numFor("PO_COST_MATLS").Value = PO_COST_MATLS

                If Val(Absx1.numFor("PO_COST_OTHER").Value & "") <> 0 Then
                    PO_COST_OTHER = System.Math.Round((Val(Absx1.numFor("PO_COST_OTHER").Value & "") / ((SUB_UNIT_PACK_QTY * 12))), 6)
                Else
                    PO_COST_OTHER = 0
                End If

                If Val(Absx1.numFor("PO_COST_QUOTA").Value & "") <> 0 Then
                    PO_COST_QUOTA = System.Math.Round((Val(Absx1.numFor("PO_COST_QUOTA").Value & "") / (SUB_UNIT_PACK_QTY * 12)), 6)
                Else
                    PO_COST_QUOTA = 0
                End If

                PO_COST_COMM = Val(Absx1.numFor("PO_COST_COMM").Value & "")

                PO_COST = PO_COST_VCOST + PO_COST_MATLS + PO_COST_OTHER ' + PO_COST_QUOTA Per Anna this will not be used in any calculations until after duty is applied - 11/24.

                PO_COST_COMM = System.Math.Round(((PO_COST) * (PO_COST_COMM / 100)), 6)

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    PO_COST = PO_COST + PO_COST_COMM
                End If

                Absx1.numFor("PO_COST").Value = PO_COST
                Absx1.numFor("PO_COMM_DZ").Value = PO_COST_COMM * (SUB_UNIT_PACK_QTY * 12)
                Absx1.numFor("PO_COMM_EA").Value = PO_COST_COMM
                Absx1.numFor("PO_OTHER_DZ").Value = PO_COST_OTHER * (12 / SUB_UNIT_PACK_QTY)
                Absx1.numFor("PO_OTHER_EA").Value = PO_COST_OTHER
                Absx1.numFor("PO_QUOTA_DZ").Value = PO_COST_QUOTA * (SUB_UNIT_PACK_QTY * 12)
                Absx1.numFor("PO_QUOTA_EA").Value = PO_COST_QUOTA

                If Val(UOM) = 1 Then
                    Absx1.numFor("PO_COST_DZ").Value = Val(Absx1.numFor("PO_COST").Value & "") * 12
                Else
                    Absx1.numFor("PO_COST_DZ").Value = System.Math.Round(Val(Absx1.numFor("PO_COST").Value & "") / (Val(SUB_UNIT_PACK_QTY) / 12), 6)
                End If
        End Select
    End Sub

    Sub Lock_Line(tf As Boolean)
        'TF - True to Lock the line False to UnLock the line

        'Always lock when in inquiry mode.

        If splPOLine.Visible = True Then 'Detail mode of type "I" order.
            ' If InquiryMode Then tf = True
            tf = Not (EntryMode = "N" Or EntryMode = "E")
            Set_Read_Only(grpPOTORDR2, tf)
            Set_Read_Only(grpPOTORDR2X, tf)
            Set_Read_Only_for_ctl(chkFINISHED, InquiryMode Or EntryMode = "V")

            optUD.Enabled = Not tf

            grdPOTORDR3.DisplayLayout.Bands(0).Columns("YIELD_QTY").CellActivation = IIf(tf, UltraWinGrid.Activation.NoEdit, UltraWinGrid.Activation.AllowEdit)
            grdPOTORDR4.DisplayLayout.Bands(0).Columns("CONSUMPTION_DZ").CellActivation = IIf(tf, UltraWinGrid.Activation.NoEdit, UltraWinGrid.Activation.AllowEdit)

        Else 'Grid mode
            If Absx1.optFor("FOB_CMT").Value <> "I" Then 'Don't screw with the grid when it's a Inventory type

                With grdPOTORDR2.DisplayLayout.Bands(0)
                    ' I THINK THESE COLS SHOULD BE MAINTAINED ONLY IF DOING COSTING
                    If InquiryMode Then
                        tf = True
                        .Columns("LINE_FINISHED").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("DFQUOTA").CellActivation = UltraWinGrid.Activation.NoEdit
                    Else
                        .Columns("LINE_FINISHED").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("DFQUOTA").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If

                    For Each COLUMN_NAME As String In New String() _
                        {"STYLE_CODE", "COLOR_CODE", "YIELD_QTY",
                         "PO_QTY_ORD_DZ", "PO_QTY_ORD", "PO_DATE_SHIP_BY", "PO_DATE_ETA", "CMT_NO",
                         "PO_COST_VCOST_DZ", "PO_COST_VCOST", "PO_COST_MATLS", "PO_COST_OTHER", "PO_COST_COMM",
                         "PO_COST_BUFFER", "LINE_CLOSED", "STYLE_NOTES", "PO_COST_QUOTA"}
                        .Columns(COLUMN_NAME).CellActivation = IIf(tf, UltraWinGrid.Activation.NoEdit, UltraWinGrid.Activation.AllowEdit)
                    Next

                    For Each COLUMN_NAME As String In New String() _
                        {"PO_COST_MATLS_DZ"}
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                    Next
                End With
            End If
        End If
    End Sub

    Sub Calc_Yards_by_Color(Refresh_Yield As Boolean)

        Dim UMF As Integer
        Dim UMFEA As Integer
        Dim ReCalc As Boolean = True

        If Val(dst.Tables("POTORDR3").Compute("SUM(YIELD_QTY)", "PO_ORDER_LNO = 0") & "") <> 0 And Not Refresh_Yield Then
            ReCalc = False
        End If

        'Calc_UMF(UMF, UMFEA)
        UMF = Val(optUD.Value & "")
        UMFEA = 13 - UMF
        If UMF = 12 Then
            If Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "") <> 0 Then
                UMFEA = UMFEA * Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value)
            End If
        End If

        Dim YIELD_QTY As Int32 = Val(Absx1.numFor("YIELD_QTY").Value & "")

        For Each rowPOTORDR4 As DataRow In dst.Tables("POTORDR4").Select("PO_ORDER_LNO = 0")
            Dim SQLW As String = "PO_ORDER_LNO = 0 and FABRIC_NO = " & rowPOTORDR4.Item("FABRIC_NO")
            Dim COLOR_CODE As String = Absx1.txtFor("COLOR_CODE").Text
            If COLOR_CODE <> "AST" Then
                SQLW &= " and (COLOR_CODE is Null or COLOR_CODE = '" & COLOR_CODE & "')"
            End If
            Dim CPF As Integer = Val(dst.Tables("POTORDR3").Compute("COUNT(PO_ORDER_LNO)", SQLW) & "")

            SQLW = "PO_ORDER_LNO = 0 and FABRIC_NO = " & rowPOTORDR4.Item("FABRIC_NO")

            Dim YIELD_QTY_COLOR_TOTAL As Int32 = 0
            Dim i As Integer = 0
            Dim YIELD_QTY_COLOR As Int32 = 0

            For Each rowPOTORDR3 As DataRow In dst.Tables("POTORDR3").Select(SQLW)
                If i = CPF Then
                    YIELD_QTY_COLOR = YIELD_QTY - YIELD_QTY_COLOR_TOTAL
                Else
                    YIELD_QTY_COLOR = YIELD_QTY \ CPF
                    YIELD_QTY_COLOR_TOTAL = YIELD_QTY_COLOR_TOTAL + YIELD_QTY_COLOR
                End If
                If COLOR_CODE <> "AST" And rowPOTORDR3.Item("COLOR_CODE") & "" <> "" And rowPOTORDR3.Item("COLOR_CODE") & "" <> COLOR_CODE Then
                    rowPOTORDR3.Item("YIELD_QTY") = 0
                    rowPOTORDR3.Item("YARDS_CONSUMED") = 0
                Else
                    If ReCalc = True Then
                        rowPOTORDR3.Item("YIELD_QTY") = YIELD_QTY_COLOR
                        'Always in Dz - 8/22/02
                    End If
                    Dim YIELD_QTY_POTORDR3 As Int32 = Val(rowPOTORDR3.Item("YIELD_QTY") & "")
                    Dim CONSUMPTION_RATE As Int32 = Val(rowPOTORDR4.Item("CONSUMPTION_RATE") & "")
                    Dim YARDS_CONSUMED As Decimal
                    If Val(optUD.Value & "") = 1 Then
                        YARDS_CONSUMED = YIELD_QTY_POTORDR3 * (CONSUMPTION_RATE / 12)
                    Else
                        YARDS_CONSUMED = (YIELD_QTY_POTORDR3 / UMFEA) * CONSUMPTION_RATE
                    End If
                    rowPOTORDR3.Item("YARDS_CONSUMED") = System.Math.Round(YARDS_CONSUMED, 6)
                End If
            Next
        Next

        Call ReCalculate_PO_Cost()
        Calc_Totals_POTORDR3(False)
    End Sub

    Function Calc_FC_AVG(CMT_NO As String) As Decimal()
        Dim t(,) As Decimal
        Dim l(,) As Decimal

        ASCMAIN1.sql = "Select Max (COLOR_NO) from ICTCMTM4 WHERE CMT_NO = '" & CMT_NO & "'"
        Dim COLOR_MAX As Integer = Val(ASCDATA1.GetDataValue & "")

        ReDim t(COLOR_MAX, 2)
        ReDim l(COLOR_MAX, 1)

        Dim rowCMTCOST As DataRow

        ' Total Purchases, net of Purchase Adjustments
        't(i, 1) is the yards still available for each color.
        't(i, 2) is the total cost remaining for each color.

        ASCMAIN1.sql = "Select 'X' " & vbCrLf
        For i As Integer = 1 To COLOR_MAX
            ASCMAIN1.sql &= ", SUM (DECODE (ICTCMTM3.COLOR_NO, " & Format(i, "0") &
                ",NVL(ICTCMTM3.YARDS,0),0)) COLOR_" & Format(i, "0") & vbCrLf
        Next i
        For i As Integer = 1 To COLOR_MAX
            ASCMAIN1.sql &= ", SUM (DECODE (ICTCMTM3.COLOR_NO, " & Format(i, "0") &
                ",NVL(ICTCMTM3.YARDS,0) * NVL(ICTCMTM2.FABRIC_COST,0),0)) FABRIC_COST_" & Format(i, "0") & vbCrLf
        Next i
        ASCMAIN1.sql &= " from ICTCMTM2,ICTCMTM3 " & vbCrLf _
        & " Where ICTCMTM2.CMT_NO = ICTCMTM3.CMT_NO " & vbCrLf _
        & "   AND ICTCMTM2.CMT_LNO = ICTCMTM3.CMT_LNO " & vbCrLf _
        & "   AND ICTCMTM2.CMT_NO = '" & CMT_NO & "'" & vbCrLf
        rowCMTCOST = ASCDATA1.GetDataRow
        For i As Integer = 1 To COLOR_MAX
            t(i, 1) = Val(rowCMTCOST.Item(i) & "")
            t(i, 2) = Val(rowCMTCOST.Item(i + COLOR_MAX) & "")
        Next i

        'Fugure out the last cost
        ASCMAIN1.sql = "Select 'X' " & vbCrLf
        For i As Integer = 1 To COLOR_MAX
            ASCMAIN1.sql &= ", DECODE (ICTCMTM3.COLOR_NO, " & Format(i, "0") &
                ",NVL(ICTCMTM3.YARDS,0),0) COLOR_" & Format(i, "0") & vbCrLf
        Next i
        For i As Integer = 1 To COLOR_MAX
            ASCMAIN1.sql &= ", DECODE (ICTCMTM3.COLOR_NO, " & Format(i, "0") &
                ",NVL(ICTCMTM3.YARDS,0) * NVL(ICTCMTM2.FABRIC_COST,0),0) FABRIC_COST_" & Format(i, "0") & vbCrLf
        Next i
        ASCMAIN1.sql &= " from ICTCMTM2,ICTCMTM3 " & vbCrLf _
            & " Where ICTCMTM2.CMT_NO = ICTCMTM3.CMT_NO " & vbCrLf _
            & "   AND ICTCMTM2.CMT_LNO = ICTCMTM3.CMT_LNO " & vbCrLf _
            & "   AND ICTCMTM2.CMT_NO = '" & CMT_NO & "'" & vbCrLf

        For Each rowCMTCOST In ASCDATA1.GetDataTable.Rows
            For i As Integer = 1 To COLOR_MAX
                If Val(rowCMTCOST.Item(i + COLOR_MAX) & "") <> 0 Then
                    If Val(rowCMTCOST.Item(i) & "") / Val(rowCMTCOST.Item(i + COLOR_MAX) & "") <> 0 Then
                        l(i, 1) = Val(rowCMTCOST.Item(i + COLOR_MAX) & "") / Val(rowCMTCOST.Item(i) & "")
                    Else
                        l(i, 1) = l(i, 1)
                    End If
                Else
                    l(i, 1) = l(i, 1)
                End If
            Next i
        Next

        ' Total Consumption from Purchase Orders

        ASCMAIN1.sql = "Select 'X'" & vbCrLf
        For i As Integer = 1 To COLOR_MAX
            ASCMAIN1.sql &= ", Sum (DECODE(POTORDR3.COLOR_NO, " & Format(i, "0") &
                ", NVL(POTORDR3.YARDS_CONSUMED,0),0)) COLOR_" & Format(i, "0") & vbCrLf
        Next i
        For i As Integer = 1 To COLOR_MAX
            ASCMAIN1.sql &= ", Sum (DECODE(POTORDR3.COLOR_NO, " & Format(i, "0") &
                ", NVL(POTORDR3.YARDS_CONSUMED,0) * NVL(POTORDR3.FABRIC_COST,0),0)) FABRIC_COST_" & Format(i, "0") & vbCrLf
        Next i
        ASCMAIN1.sql &= " from POTORDR2,POTORDR3 " & vbCrLf _
            & " where POTORDR2.CMT_NO = '" & CMT_NO & "'" & vbCrLf _
            & " and POTORDR2.PO_ORDER_NO <> '" & PO_ORDER_NO & "'" & vbCrLf _
            & " and POTORDR3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & " and POTORDR3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & " and POTORDR3.FABRIC_COST <> 0" & vbCrLf

        rowCMTCOST = ASCDATA1.GetDataRow
        For i As Integer = 1 To COLOR_MAX
            t(i, 1) = t(i, 1) - Val(rowCMTCOST.Item(i) & "")
            t(i, 2) = t(i, 2) - Val(rowCMTCOST.Item(i + COLOR_MAX) & "")
        Next i

        ' Avg Cost of Balances by Color

        Dim FC_AVG(COLOR_MAX) As Decimal
        For i As Integer = 1 To COLOR_MAX
            If t(i, 1) = 0 Then
                FC_AVG(i) = 0
            Else
                If t(i, 1) < 0 Or t(i, 2) < 0 Then
                    'FC_AVG(i) = 0
                    'Use Last cost lot as value per Gabe.
                    FC_AVG(i) = l(i, 1)
                    lblLastCost.Visible = True
                Else
                    FC_AVG(i) = System.Math.Round(t(i, 2) / t(i, 1), 6)
                End If
            End If
            t(0, 1) = t(0, 1) + t(i, 1)
            t(0, 2) = t(0, 2) + t(i, 2)
        Next i

        Return FC_AVG
    End Function
#End Region

    Private Sub cmdPercent_Click(sender As System.Object, e As System.EventArgs) Handles cmdPercent.Click
        Setup_cmd("Percent")
    End Sub

    Private Sub cmdCostInc_Click(sender As System.Object, e As System.EventArgs) Handles cmdCostInc.Click
        Setup_cmd("CostInc")
    End Sub

    Sub Setup_cmd(cmd As String)
        If cmd = "CostInc" Then
            If Val(Absx1.numFor("POTORDR2_LINE.PO_COST_BUFFER").Value) = 0 Then
                Absx1.numFor("POTORDR2_LINE.PO_COST_BUFFER").Value = "2.00"
                cmdCostInc.Text = "0%"
            Else
                Absx1.numFor("POTORDR2_LINE.PO_COST_BUFFER").Value = "0.00"
                cmdCostInc.Text = "2%"
            End If
        ElseIf cmd = "Percent" Then
            If Val(Absx1.numFor("POTORDR2_LINE.PO_COST_COMM").Value) = 0 Then
                'Changed from 4 to 3 Per Anna 12/30/14 - WR.
                Absx1.numFor("POTORDR2_LINE.PO_COST_COMM").Value = "3.00"
                cmdPercent.Text = "0%"
            Else
                Absx1.numFor("POTORDR2_LINE.PO_COST_COMM").Value = "0.00"
                cmdPercent.Text = "3%"
            End If
        End If
    End Sub

    Private Function Get_Code_SQL(p1 As String) As String
        Throw New NotImplementedException
    End Function

    Private Sub optUD_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optUD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Dependent_Calculations("YIELD_QTY")

        'Dim tf As Boolean = (optUD.Value = "D")

        'With grdPOTORDR3.DisplayLayout.Bands(0)
        '    If Not cost_calc Then
        '        .Columns("PO_QTY_SHP").Hidden = tf
        '        .Columns("NET_OPEN").Hidden = tf
        '        .Columns("PO_QTY_SHP_DZ").Hidden = Not tf
        '        .Columns("NET_OPEN_DZ").Hidden = Not tf
        '        .Columns("PO_QTY_REC").Hidden = tf
        '        .Columns("PO_QTY_REC_DZ").Hidden = Not tf
        '    End If

        '    If cost_calc Then
        '        .Columns("PO_QTY_REC").Hidden = tf
        '        .Columns("PO_COST_VCOST").Hidden = tf
        '        .Columns("PO_COST_MATLS").Hidden = tf
        '        .Columns("PO_COST_OTHER").Hidden = tf
        '        .Columns("PO_COST_QUOTA").Hidden = tf
        '        .Columns("PO_COST_QUOTA_DF").Hidden = tf
        '        .Columns("FIRST_COST_TOTAL").Hidden = tf
        '        .Columns("COMMISION_COST").Hidden = tf

        '        .Columns("PO_QTY_REC_DZ").Hidden = Not tf
        '        .Columns("PO_COST_VCOST_DZ").Hidden = Not tf
        '        .Columns("PO_COST_MATLS_DZ").Hidden = Not tf
        '        .Columns("FIRST_COST_TOTAL_DZ").Hidden = Not tf
        '    End If
        'End With

    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_POTORDRX()
    End Sub

    Private Sub grdPOTORDRT_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRT.InitializeRow
        If e.Row.Cells("LNO").Value & "" = "4" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        End If
    End Sub

    Private Sub ssdDZGRD_ValueChanged(sender As System.Object, e As System.EventArgs) Handles ssdDZGRD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If Not ScreenMode Then
            Exit Sub
        End If

        Set_grdPOTORDR2_cols_Visibility()

    End Sub

    Private Sub tabPO_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabPO.SelectedTabChanged
        Setup_tabPO()
    End Sub

    Sub Setup_tabPO()
        If SELECTION_NO = 0 Then Exit Sub
        ' If (EntryMode = "S") Then Exit Sub
        If tabPO.SelectedTab.Key = "Styles on Open POs" Then
            Refresh_Styles_Open_POs()
        End If
        UltraExplorerBar1.Groups("Styles on Open POs").Visible = Not ScreenMode And (tabPO.SelectedTab.Key = "Styles on Open POs")
        UltraExplorerBar1.Groups("Status Filter").Visible = Not ScreenMode And (tabPO.SelectedTab.Key = "Open POs")

        If tabPO.SelectedTab.Key = "XLS" Then
            Refresh_XLS()
        End If
    End Sub

    Sub Refresh_Styles_Open_POs(Optional PO_SHIPMENT_NO As String = "")
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Creating Open POs by Style")
        ASCMAIN1.sql = sqlPOTORDRS
        If PO_SHIPMENT_NO = "" Then
            txtPO_SHIPMENT_NO.Text = ""
            grdPOTORDRS.Text = "Styles on Open POs"
            grdPOTORDRS.DisplayLayout.Bands(0).Columns("PO_QTY_OPN").Header.Caption = "Qty Opn"
        Else
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "PO_STATUS = 'O'", "(PO_ORDER_NO,PO_ORDER_LNO) in (Select PO_ORDER_NO,PO_ORDER_LNO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "SUM (PO_QTY_OPN)", "SUM (PO_QTY_SHP)")
            grdPOTORDRS.Text = "Styles on PO Shipment " & PO_SHIPMENT_NO
            grdPOTORDRS.DisplayLayout.Bands(0).Columns("PO_QTY_OPN").Header.Caption = "Qty Shp"
        End If

        Fill_Records("POTORDRS", "", True, ASCMAIN1.sql)

        If PO_SHIPMENT_NO <> "" Then
            ASCMAIN1.sql = "Select STYLE_CODE, PO_QTY_PER_CTN, SUM (CARTONS) CARTONS" _
                & " from POTSHIP7 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' AND PPK_CODE IS NULL" _
                & " group by STYLE_CODE, PO_QTY_PER_CTN"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "CARTONS")
                Dim rowPOTORDRS As DataRow = dst.Tables("POTORDRS").Rows.Find(New String() {row.Item("STYLE_CODE")})
                If rowPOTORDRS IsNot Nothing Then
                    If rowPOTORDRS.Item("CARTON_PACK_QTY") & "" = row.Item("PO_QTY_PER_CTN") & "" Then
                        ' do nothing
                    Else
                        rowPOTORDRS.Item("CARTON_PACK_QTY") = row.Item("PO_QTY_PER_CTN")
                    End If

                End If
            Next
        End If

        grdPOTORDRS.Tag = PO_SHIPMENT_NO
        Sort_grdColumns(grdPOTORDRS, "STYLE_CODE")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden = True
        grdPOTORDRS.DisplayLayout.Bands(0).Columns("POS_UPDATED").Hidden = True
        grdPOTORDRS.DisplayLayout.Bands(0).Columns("POS_SKIPPED").Hidden = True
    End Sub

    Private Sub grdPOTORDRS_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRS.InitializeRow
        If e.Row.IsDataRow Then
            If Val(e.Row.Cells("CARTON_PACK_QTY_STYLE").Value & "") <> Val(e.Row.Cells("CARTON_PACK_QTY").Value & "") Then

                e.Row.Cells("CARTON_PACK_QTY").Appearance.BackColor = Drawing.Color.Yellow
                e.Row.Cells("CARTON_PACK_QTY").ToolTipText = "" _
                    & "Style Master (" & e.Row.Cells("CARTON_PACK_QTY_STYLE").Value & ")" _
                    & " does not agree with " & IIf(txtPO_SHIPMENT_NO.Text = "", "PO", "Shipment " & txtPO_SHIPMENT_NO.Text) & " (" & e.Row.Cells("CARTON_PACK_QTY").Value & ")"
            End If

            'For Each COLUMN_NAME As String In New String() {"SUB_UNIT_PACK_QTY", "CARTON_PACK_QTY", "INNER_PACK_QTY"}
            '    With e.Row.Cells(COLUMN_NAME)
            '        If Val(.Value & "") = 0 Then
            '            .Appearance.ForeColor = Drawing.Color.Red
            '            .ToolTipText = "Style Master is 0"
            '        Else

            '            'If Val(.Value & "") <> Val(e.Row.Cells(COLUMN_NAME & "_MIN").Value & "") _
            '            'Or Val(.Value & "") <> Val(e.Row.Cells(COLUMN_NAME & "_MIN").Value & "") Then
            '            '    .Appearance.BackColor = Drawing.Color.Yellow
            '            '    .ToolTipText = "Style Master (" & .Value & ") does not agree with PO (" & e.Row.Cells(COLUMN_NAME & "_MIN").Value & ")"
            '            'End If
            '        End If
            '    End With
            'Next
        End If
    End Sub

    Private Sub cmdGetStyles_Click(sender As System.Object, e As System.EventArgs) Handles cmdGetStyles.Click
        Dim PO_SHIPMENT_NO As String = txtPO_SHIPMENT_NO.Text
        Dim rowPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIPMENT_NO)
        If rowPOTSHIP1 Is Nothing Then
            MsgBox("No Record of PO Shipment " & PO_SHIPMENT_NO,
                   MsgBoxStyle.OkOnly, "Cannot Change Case Pack Qtys for this Shipment")
            'ElseIf rowPOTSHIP1.Item("LP_STATUS") & "" = "1" Then
            '    MsgBox("PO Shipment " & PO_SHIPMENT_NO & " has already been sent to the 3PL", _
            '           MsgBoxStyle.OkOnly, "Cannot Change Case Pack Qtys for this Shipment")
        Else
            Refresh_Styles_Open_POs(PO_SHIPMENT_NO)
        End If
    End Sub

    Function Check_Changed_Fields(Optional clear_before_filling As Boolean = True) As Boolean

        Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
        '  PO_HDR_CTR_REV += 1 ' already done in load record

        Dim LAST_DATE As Date = DATETIME_STAMP
        If EntryMode = "N" Then Stop
        Dim REV_LNO As Integer = 0

        Check_Changed_Fields = False

        If clear_before_filling Then
            dst.Tables("POTORDXR").Rows.Clear()
        End If

        ASCMAIN1.Progress("Logging Header Changes")

        For i As Integer = 0 To rowPOTORDR1.Table.Columns.Count - 1
            Dim COLUMN_NAME As String = dst.Tables("POTORDR1").Columns(i).ColumnName

            If rowPOTORDR1.Item(COLUMN_NAME) & "" _
            <> rowPOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                Check_Changed_Fields = True
                ASCMAIN1.Progress("-", COLUMN_NAME)
                Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
                With rowPOTORDXR
                    .Item("REV_NO") = PO_HDR_CTR_REV
                    REV_LNO += 1
                    .Item("REV_LNO") = REV_LNO
                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("PO_ORDER_LNO") = 0
                    .Item("INIT_DATE") = LAST_DATE
                    .Item("INIT_USER") = ASCMAIN1.USER_ID
                    .Item("COLUMN_NAME") = COLUMN_NAME
                    .Item("OLD_VALUE") = rowPOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original)
                    .Item("NEW_VALUE") = rowPOTORDR1.Item(COLUMN_NAME)
                    .Item("EMODE") = EntryMode
                End With
                dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                Check_Changed_Fields = True
            End If
        Next i

        ASCMAIN1.Progress("Logging Detail Changes")

        ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each rowPOTORDR2_orig As DataRow In dt.Rows
            Dim PO_ORDER_LNO As Int64 = rowPOTORDR2_orig.Item("PO_ORDER_LNO")
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
            If rowPOTORDR2 Is Nothing Then ' Line was Deleted
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowPOTORDR2_orig.Table.Columns(i).ColumnName
                    Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
                    With rowPOTORDXR
                        .Item("REV_NO") = PO_HDR_CTR_REV
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("PO_ORDER_NO") = PO_ORDER_NO
                        .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        .Item("INIT_DATE") = LAST_DATE
                        .Item("INIT_USER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("OLD_VALUE") = rowPOTORDR2_orig.Item(COLUMN_NAME)
                        '.Item("NEW_VALUE") = ""
                        .Item("EMODE") = EntryMode
                    End With
                    dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                Next

                Check_Changed_Fields = True
            Else
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowPOTORDR2_orig.Table.Columns(i).ColumnName
                    If rowPOTORDR2.Item(COLUMN_NAME) & "" <> rowPOTORDR2_orig.Item(COLUMN_NAME) & "" Then
                        ' Value in Column was Changed
                        Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
                        With rowPOTORDXR
                            .Item("REV_NO") = PO_HDR_CTR_REV
                            REV_LNO += 1
                            .Item("REV_LNO") = REV_LNO
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            .Item("INIT_DATE") = LAST_DATE
                            .Item("INIT_USER") = ASCMAIN1.USER_ID
                            .Item("COLUMN_NAME") = COLUMN_NAME
                            .Item("OLD_VALUE") = rowPOTORDR2_orig.Item(COLUMN_NAME)
                            .Item("NEW_VALUE") = rowPOTORDR2.Item(COLUMN_NAME)
                            .Item("EMODE") = EntryMode
                        End With
                        dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                        Check_Changed_Fields = True
                    End If
                Next
            End If
        Next

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("", "", DataViewRowState.Added)
            Dim PO_ORDER_LNO = rowPOTORDR2.Item("PO_ORDER_LNO")
            ' For i As Integer = 0 To dt.Columns.Count - 1
            Dim COLUMN_NAME As String = "" ' dt.Columns(i).ColumnName
            Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
            With rowPOTORDXR
                .Item("REV_NO") = PO_HDR_CTR_REV
                REV_LNO += 1
                .Item("REV_LNO") = REV_LNO
                .Item("PO_ORDER_NO") = PO_ORDER_NO
                .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                .Item("INIT_DATE") = LAST_DATE
                .Item("INIT_USER") = ASCMAIN1.USER_ID
                .Item("COLUMN_NAME") = COLUMN_NAME
                '.Item("OLD_VALUE") = ""
                .Item("NEW_VALUE") = "PO Line Added" ' rowPOTORDR2.Item(COLUMN_NAME)
                .Item("EMODE") = EntryMode
            End With
            dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
            Check_Changed_Fields = True
            'Next
        Next

        ASCMAIN1.Progress("")
        Return Check_Changed_Fields
    End Function

    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String,
                     Optional update_database As Boolean = False,
                     Optional EVENT_KEY As String = "")

        Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
        rowTATEVNT1.Item("TABLE_NAME") = "POTORDR1"
        rowTATEVNT1.Item("TABLE_KEY") = PO_ORDER_NO
        rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
        rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowTATEVNT1.Item("EVENT_TYPE") = EVENT_TYPE
        rowTATEVNT1.Item("EVENT_DESC") = EVENT_DESC
        rowTATEVNT1.Item("EVENT_KEY") = EVENT_KEY
        dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
        If update_database Then Update_Record_TDA("TATEVNT1")

    End Sub

    Private Sub grpDetails_Click(sender As System.Object, e As System.EventArgs) Handles grpDetails.Click

    End Sub

    Private Sub grdPOTORDR2_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTORDR2.InitializeLayout

    End Sub

    Private Sub cmdProrate_Click(sender As System.Object, e As System.EventArgs) Handles cmdProrate.Click
        For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDR3.Rows

        Next
    End Sub

    Function Change_Style_Color_in_Shipment(verify As Boolean) As String
        Dim EMsg As String = ""

        ' wjz 10/17/2012
        ' the following check was put in place after allowing change to color for a po line that has been shipped
        ' the problem is trying to re-cartonize the color change, where the cartonization may be complex
        ' for now, since everyone just wants "this receipt" to be processed so that we can move on,
        ' I am locking down the code so that only a simple change (like we have in "this receipt") can be processed
        ' where the cartonization is simple
        ' if we ever get to a more complicated example, we will have to deal with it then
        ASCMAIN1.sql = "SELECT PO_ORDER_NO, PO_ORDER_LNO, STYLE_CODE, COLOR_CODE, PO_QTY_SHP" & vbCrLf _
            & " from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "' AND PO_QTY_SHP <> 0"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim PO_ORDER_LNO As Int64 = Val(row.Item("PO_ORDER_LNO"))
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})

            If rowPOTORDR2.Item("STYLE_CODE") <> row.Item("STYLE_CODE") _
            Or rowPOTORDR2.Item("COLOR_CODE") <> row.Item("COLOR_CODE") Then
                Dim QTY_TRAN As Int64 = Val(rowPOTORDR2.Item("PO_QTY_SHP") & "")

                Dim POSHPs As String = "Select PO_SHIPMENT_NO, PO_SHIPMENT_LNO" & vbCrLf _
                    & " from POTSHIP3 where PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                    & "  and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)

                If verify Then
                    ASCMAIN1.sql = "Select Sum (QTY) from POTSHIP8" & vbCrLf _
                        & " where STYLE_CODE = '" & row.Item("STYLE_CODE") & "'" & vbCrLf _
                        & "   and COLOR_CODE = '" & row.Item("COLOR_CODE") & "'" & vbCrLf _
                        & "   and (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) in (" & POSHPs & ")"
                    Dim QTY As Int64 = Val(ASCDATA1.GetDataValue)
                    If QTY <> QTY_TRAN Then
                        EMsg &= vbCr & "Problem with Style/Color Change on Line " & CStr(PO_ORDER_LNO)
                        EMsg &= vbCr & " - Cannot Change Carton Packing - Please Contact ABS for Help"
                    End If
                Else
                    ASCMAIN1.sql = "Update POTSHIP8" & vbCrLf _
                        & " Set STYLE_CODE = '" & rowPOTORDR2.Item("STYLE_CODE") & "'" _
                        & "   , COLOR_CODE = '" & rowPOTORDR2.Item("COLOR_CODE") & "'" & vbCrLf _
                        & " where STYLE_CODE = '" & row.Item("STYLE_CODE") & "'" & vbCrLf _
                        & "   and COLOR_CODE = '" & row.Item("COLOR_CODE") & "'" & vbCrLf _
                        & "   and (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) in (" & POSHPs & ")"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "POTSHIP8", "POTSHIP7")
                    ASCDATA1.ExecuteSQL()


                    ASCMAIN1.sql = "Select POTSHIP1.WHSE_CODE, POTSHIP3.PO_QTY_SHP from POTSHIP1,POTSHIP3" & vbCrLf _
                        & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                        & "   and (POTSHIP3.PO_SHIPMENT_NO,POTSHIP3.PO_SHIPMENT_LNO) in (" & POSHPs & ")" & vbCrLf _
                        & "   and POTSHIP3.PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                        & "   and POTSHIP3.PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)

                    For Each rowPOTSHIPX As DataRow In ASCDATA1.GetDataTable.Rows
                        Dim WHSE_CODE As String = rowPOTSHIPX.Item("WHSE_CODE")
                        Dim QTY As Int64 = Val(rowPOTSHIPX.Item("PO_QTY_SHP") & "")

                        TAC.ICCMAIN1.Update_ICTSTAT2(
                            row.Item("STYLE_CODE"),
                            row.Item("COLOR_CODE"),
                            WHSE_CODE, "WHSE_QTY_TRAN", -1 * QTY)

                        TAC.ICCMAIN1.Update_ICTSTAT2(
                            rowPOTORDR2.Item("STYLE_CODE"),
                            rowPOTORDR2.Item("COLOR_CODE"),
                            WHSE_CODE, "WHSE_QTY_TRAN", QTY)
                    Next
                End If
            End If
        Next

        Return EMsg

    End Function


#Region "grdPOTORDRN"

    Private Sub grdPOTORDRN_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDRN.AfterRowUpdate

    End Sub

    Private Sub grdPOTORDRN_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTORDRN.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("INIT_OPER").Value <> ASCMAIN1.USER_ID Then
                MsgBox("You may only delete rows that you have entered", MsgBoxStyle.OkOnly, "Cannot Delete Rows")
                e.Cancel = True
            End If
        Next
    End Sub

    Private Sub grdPOTORDRN_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTORDRN.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("PO_ORDER_CLNO").Value = Val(dst.Tables("POTORDRN").Compute("MAX(PO_ORDER_CLNO)", "") & "") + 1
            e.Row.Cells("PO_ORDER_NO").Value = PO_ORDER_NO
            e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
        Else
            e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        End If
    End Sub
#End Region

    Private Sub grdPOTORDRH_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDRH.AfterRowActivate

        dst.Tables("POTORDRZ").Rows.Clear()
        Dim PO_ORDER_NO As String = grdPOTORDRH.ActiveRow.Cells("PO_ORDER_NO").Value
        Dim PO_HDR_CTR_REV As Int32 = Val(grdPOTORDRH.ActiveRow.Cells("PO_HDR_CTR_REV").Value & "")
        TAC.POCMAIN1.Build_POTORDRZ(dst.Tables("POTORDRZ"),
                                    PO_ORDER_NO,
                                     PO_HDR_CTR_REV - 1,
                                    PO_HDR_CTR_REV)
    End Sub

    Private Sub cmbUpdateSelected_Click(sender As System.Object, e As System.EventArgs) Handles cmbUpdateSelected.Click
        If grdPOTORDR2.Selected.Rows.Count = 0 Then
            MsgBox("No Rows Selected", MsgBoxStyle.OkOnly, "Cannot Update Selected Rows")
            Exit Sub
        End If
        For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDR2.Selected.Rows
            grow.Cells("PO_DATE_SHIP_BY").Value = dtePO_CONF_SHIP_BY.Value
            grow.Cells("PO_CONF_DATE").Value = dtePO_CONF_DATE.Value
            grow.Cells("PO_CONF_NO").Value = txtPO_CONF_NO.Text
            grow.Cells("CONFIRMED").Value = "1"
            grow.Update()
        Next
    End Sub

    Sub Cost_Calculator()

        Dim grow As UltraWinGrid.UltraGridRow = grdPOTORDR2.Selected.Rows(0)
        Dim PO_ORDER_LNO As Int32 = Val(grow.Cells("PO_ORDER_LNO").Value & "")

        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
        If rowPOTORDR2 IsNot Nothing Then
            Using F As New POFORDRM
                F.rowPOTORDR2 = rowPOTORDR2
                F.Text = "PO Line " & rowPOTORDR2.Item("PO_ORDER_LNO") & " - Cost Calculator"
                F.ShowDialog()

                If F.ok2update Then
                    For Each grow In grdPOTORDR2.Selected.Rows
                        PO_ORDER_LNO = Val(grow.Cells("PO_ORDER_LNO").Value & "")
                        grdPOTORDR2.ActiveRow = grow
                        rowPOTORDR2 = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                        rowPOTORDR2.Item("YARDS_CONSUMED") = F.CONSUMPTION
                        rowPOTORDR2.Item("FABRIC_COST") = F.FABRIC_COST
                        rowPOTORDR2.Item("PO_COST_MATLS_DZ") = F.TOTAL_COST
                        grow.Update()
                        ReCalculate_PO_Cost()
                    Next
                End If
            End Using
        End If
    End Sub

    Sub Check_ETA()
        If Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Or Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" = "" Then
            Exit Sub
        Else
            If Format(Absx1.dteFor("PO_DATE_ETA").Value, "yyyyMMdd") < Format(Absx1.dteFor("PO_DATE_SHIP_BY").Value, "yyyyMMdd") Then
                Exit Sub
            End If
        End If

        If Absx1.optFor("FOB_CMT").Value = "B" Then Exit Sub
        Calculate_ETD_to_ETA()

        If ASCMAIN1.CLIENT = "NYA" Then
            ' ask when editing for NYA
        Else
            If EntryMode <> "N" Then Exit Sub
        End If


        If Absx1.dteFor("PO_DATE_ETA").Value <> CDate(Absx1.dteFor("PO_DATE_SHIP_BY").Value).AddDays(ETD_to_ETA) Then
            Dim TT As Integer = CDate(Absx1.dteFor("PO_DATE_ETA").Value).Subtract(CDate(Absx1.dteFor("PO_DATE_SHIP_BY").Value)).TotalDays
            If ASCMAIN1.CLIENT = "NYA" OrElse MsgBox("ETA Date is " & CStr(TT) & " Days later than Ship By Date" _
                      & vbCrLf & "Port to Warehouse transit time is " & ETD_to_ETA & " days" _
                      & vbCrLf & "Reset ETA to " & CDate(Absx1.dteFor("PO_DATE_SHIP_BY").Value).AddDays(ETD_to_ETA) & "?",
                      MsgBoxStyle.YesNo, "ETD to ETA - Option to Reset") = MsgBoxResult.Yes Then
                Dim PO_DATE_ETA_FORMER As Date = Absx1.dteFor("PO_DATE_ETA").Value
                Absx1.dteFor("PO_DATE_ETA").Value = CDate(Absx1.dteFor("PO_DATE_SHIP_BY").Value).AddDays(ETD_to_ETA)
                For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("PO_DATE_ETA = '" & Format(PO_DATE_ETA_FORMER, "MM/dd/yyyy") & "'")
                    rowPOTORDR2.Item("PO_DATE_ETA") = Absx1.dteFor("PO_DATE_ETA").Value
                Next
            End If
        End If

    End Sub

    Sub Seek_Approval(eItemKey As String)

        Dim PO_AMT_ORD As Decimal = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
        Dim PO_AMT_OPN As Decimal = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_OPN)", "") & "")
        Dim PO_QTY_ORD As Int64 = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_ORD)", "") & "")
        Dim PO_QTY_OPN As Int64 = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", "") & "")
        Dim PO_DATE_REQUIRED_MIN As Date = Nothing ' dst.Tables("POTORDR2").Compute("MIN(PO_DATE_REQUIRED)", "")
        Dim PO_DATE_REQUIRED_MAX As Date = Nothing ' dst.Tables("POTORDR2").Compute("MAX(PO_DATE_REQUIRED)", "")
        Dim PO_NINV_AMOUNT As Decimal = 0
        Dim PO_TOTAL_AMT = 0

        If dst.Tables("POTORDR2").Select("", "", DataViewRowState.CurrentRows).Length > 0 Then
            PO_AMT_ORD = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
            PO_AMT_OPN = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_OPN)", "") & "")
            PO_QTY_ORD = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_ORD)", "") & "")
            PO_QTY_OPN = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", "") & "")
            PO_DATE_REQUIRED_MIN = dst.Tables("POTORDR2").Compute("MIN(PO_DATE_ETA)", "")
            PO_DATE_REQUIRED_MAX = dst.Tables("POTORDR2").Compute("MAX(PO_DATE_ETA)", "")


            '    ElseIf dst.Tables("POTORDR5").Select("", "", DataViewRowState.CurrentRows).Length > 0 Then
            'PO_AMT_ORD = Val(dst.Tables("POTORDR5").Compute("SUM(PO_AMT_ORD)", "") & "")
            'PO_AMT_OPN = Val(dst.Tables("POTORDR5").Compute("SUM(PO_AMT_OPN)", "") & "")
            'PO_QTY_ORD = Val(dst.Tables("POTORDR5").Compute("SUM(PO_QTY_ORD)", "") & "")
            'PO_QTY_OPN = Val(dst.Tables("POTORDR5").Compute("SUM(PO_QTY_OPN)", "") & "")
            'PO_DATE_REQUIRED_MIN = dst.Tables("POTORDR2").Compute("MIN(PO_NINV_DATE_REQ)", "")
            'PO_DATE_REQUIRED_MAX = dst.Tables("POTORDR2").Compute("MAX(PO_NINV_DATE_REQ)", "")
            'PO_TOTAL_AMT = PO_AMT_ORD
        End If

        APPR_NOTES = ""
        APPR_DECISION = ""
        APPR_BY = ""

        'PO_NINV_AMOUNT = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMOUNT)", "") & "")
        PO_TOTAL_AMT = PO_AMT_ORD '+ PO_NINV_AMOUNT

        Dim LBL As String = "Total Purchase Order Amount is " & Format(PO_TOTAL_AMT, "$#,##0.00") _
                                  & vbCrLf & "Total Amount Open to Ship is " & Format(PO_AMT_OPN, "$#,##0.00") _
                                  & vbCrLf & vbCrLf & "Terms are " & txtTERM_DESC.Text _
                                  & vbCrLf & vbCrLf & "Total Units Ordered are " & Format(PO_QTY_ORD, "#,##0") _
                                  & vbCrLf & "Total Units Open are " & Format(PO_QTY_OPN, "#,##0") _
                                  & vbCrLf & vbCrLf & "ETA Date Range is " & Format(PO_DATE_REQUIRED_MIN, "MM/dd/yy") & " thru " & Format(PO_DATE_REQUIRED_MAX, "MM/dd/yy") _
                                  & vbCrLf & vbCrLf & "Enter Notes to Record with this " & IIf(eItemKey = "Reject", "Rejection", "Approval")

        APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, "OK To " & IIf(eItemKey = "Reject", "Reject", "Approve") & " this Purchase?", False, 60, IIf(eItemKey = "Reject", "Rejected", "Approved"))
        If APPR_NOTES <> "" Then
            APPR_DECISION = IIf(eItemKey = "Reject", "R", "A")
            APPR_BY = ASCMAIN1.USER_ID
        End If

    End Sub

    Sub Update_Approval(eItemKey As String)

        Dim PO_AMT_ORD As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
        'Dim PO_NINV_AMOUNT As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMOUNT)", "") & "")
        Dim PO_TOTAL_AMT As Decimal = PO_AMT_ORD ' + PO_NINV_AMOUNT

        BeginTrans()

        If APPR_DECISION = "R" Then
            ASCMAIN1.sql = "Update POTORDR1 Set PO_APPR_NOTES = :PARM1, PO_APPR_PENDING = :PARM2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {APPR_NOTES, "0"})
            Record_Event("POAPPR", "Rejected at " & Format(PO_TOTAL_AMT, "$#,##0.00") & "; " & APPR_NOTES, True)
        Else
            ASCMAIN1.sql = "Update POTORDR1 Set PO_APPR_DATE = :PARM1, PO_APPR_BY = :PARM2, PO_APPR_AMOUNT = :PARM3, PO_APPR_NOTES = :PARM4 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, APPR_BY, PO_TOTAL_AMT, APPR_NOTES})
            Record_Event("POAPPR", "Approved for " & Format(PO_TOTAL_AMT, "$#,##0.00") & "; " & APPR_NOTES, True)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        If APPR_DECISION = "R" Then
            CommitTrans("Rejection Complete")
        Else
            CommitTrans("Approval Complete")
        End If
    End Sub

    Private Sub optMyPOs_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optMyPOs.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_POTORDRX()
    End Sub

    Sub Setup_tabPOTORDR2()
        UltraExplorerBar1.Groups("Back-to-Back").Visible = ScreenMode And (tabPOTORDR2.SelectedTab.Key = "Back-to-Back")
        UltraExplorerBar1.Groups("Confirmation").Visible = ScreenMode And (confirm_notes_mode Or (EntryMode = "E")) And (tabPOTORDR2.SelectedTab.Key = "PO Details")
        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If (EntryMode = "E") Then UltraExplorerBar1.Groups("Confirmation").Visible = False
        End If
        UltraExplorerBar1.Groups("XLS Commands").Visible = ScreenMode And (tabPOTORDR2.SelectedTab.Key = "XLS")
    End Sub

    Private Sub tabPOTORDR2_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabPOTORDR2.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabPOTORDR2()
    End Sub

    Sub Calculate_ETD_to_ETA()
        ETD_to_ETA = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETD_TO_ETA") & "")

        If Absx1.optFor("FOB_CMT").Value = "B" Then
            'ETD_TO_ETA = 0
        Else
            If Absx1.txtFor("PORT_CODE_ORIG").Text = "" Or Absx1.txtFor("WHSE_CODE").Text = "" Then
            Else
                Dim rowICTPORT2 As DataRow = LookUp("ICTPORT2", New String() {Absx1.txtFor("PORT_CODE_ORIG").Text, Absx1.txtFor("WHSE_CODE").Text})
                If rowICTPORT2 IsNot Nothing Then
                    ETD_to_ETA = Val(rowICTPORT2.Item("ETD_TO_ETA") & "")
                End If
            End If

        End If

        If Absx1.txtFor("PO_SHIP_VIA").Text <> "" Then
            Dim row As DataRow = LookUp("POTSVIA1", Absx1.txtFor("PO_SHIP_VIA").Text)
            If row IsNot Nothing Then
                If row.Item("PO_SHIP_VIA_ETD_TO_ETA") & "" <> "" Then
                    ETD_to_ETA = Val(row.Item("PO_SHIP_VIA_ETD_TO_ETA") & "")
                End If
            End If
        End If

    End Sub

    Sub Toggle_Customer_Style_Fields(show As Boolean)
        With grdPOTORDR2.DisplayLayout.Bands(0)
            .Columns("CUST_UPC").Hidden = Not show
            .Columns("CUST_SKU").Hidden = Not show
            .Columns("CUST_STYLE_CODE").Hidden = Not show
            .Columns("CUST_COLOR_CODE").Hidden = Not show
            .Columns("CUST_SIZE_CODE").Hidden = Not show
            .Columns("STYLE_RETAIL").Hidden = Not show
        End With
    End Sub

    Sub Toggle_Ship_ETA_from_PO_Detail(show As Boolean)
        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("PO_DATE_SHIP_BY_MIN").Hidden = Not show
            .Columns("PO_DATE_ETA_MIN").Hidden = Not show
            .Columns("PO_DATE_SHIP_BY_MAX").Hidden = Not show
            .Columns("PO_DATE_ETA_MAX").Hidden = Not show
        End With
    End Sub

    Private Sub chkSplitByShipDate_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSplitByShipDate.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_POTORDRX()
        Toggle_Ship_ETA_from_PO_Detail(False)

        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("PO_DATE_SHIP_BY").Hidden = chkSplitByShipDate.Checked
            .Columns("PO_DATE_ETA").Hidden = chkSplitByShipDate.Checked
            .Columns("PO_DATE_SHIP_BY_MIN").Hidden = Not chkSplitByShipDate.Checked
            .Columns("PO_DATE_ETA_MIN").Hidden = Not chkSplitByShipDate.Checked
        End With
    End Sub


    Sub Integrity_Check()
        Dim sqlIC As String = TAC.POCMAIN1.Get_sql_Integrity_Check

        Dim dt As DataTable = ASCDATA1.GetDataTable(sqlIC)
        grdPOTORDRK.DataSource = dt
        Sort_grdColumns(grdPOTORDRK, "PO_ORDER_NO,PO_ORDER_LNO")

        For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_OPN", "PO_QTY_REC", "PS_QTY_SHP", "PS_QTY_REC", "SHPS"}
            grdPOTORDRK.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "#,##0"
        Next

        tabPO.Tabs("Integrity Check").Visible = True
        tabPO.SelectedTab = tabPO.Tabs("Integrity Check")

        If dt.Rows.Count = 0 Then
            MsgBox("No POs out of Balance with Shipments", MsgBoxStyle.OkOnly, "Verification")

        Else
            If MsgBox("OK to Fix PO Detail Qty Shipped?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is" & vbCrLf & sqlIC & ";" _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update POTORDR2 Set PO_QTY_SHP = R1.PS_QTY_SHP" & vbCrLf _
                    & "    where PO_ORDER_NO = R1.PO_ORDER_NO" & vbCrLf _
                    & "      and PO_ORDER_LNO = R1.PO_ORDER_LNO;" & vbCrLf _
                    & "   Update POTORDR2 Set PO_QTY_OPN = GREATEST(0,NVL(PO_QTY_ORD,0) - NVL(PO_QTY_SHP,0))" & vbCrLf _
                    & "    where PO_ORDER_NO = R1.PO_ORDER_NO" & vbCrLf _
                    & "      and PO_ORDER_LNO = R1.PO_ORDER_LNO;" & vbCrLf _
                    & "   Update POTORDR2 Set PO_STATUS = DECODE(PO_QTY_OPN,0,'C','O')" & vbCrLf _
                    & "    where PO_ORDER_NO = R1.PO_ORDER_NO" & vbCrLf _
                    & "      and PO_ORDER_LNO = R1.PO_ORDER_LNO;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                MsgBox("Update Complete", MsgBoxStyle.OkOnly, "Verification")
            End If
        End If
    End Sub

    Private Sub grdPOTORDRK_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRK.InitializeRow
        If (e.Row.Cells("PO_STATUS").Value & "" = "C" And Val(e.Row.Cells("PO_QTY_OPN").Value & "") = 0) _
        Or (e.Row.Cells("PO_STATUS").Value & "" = "O" And Val(e.Row.Cells("PO_QTY_OPN").Value & "") <> 0) Then
            ' ALL IS GOOD
            e.Row.Cells("PO_STATUS").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("PO_STATUS").Appearance.ForeColor = Drawing.Color.Empty
        Else
            e.Row.Cells("PO_STATUS").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("PO_STATUS").Appearance.ForeColor = Drawing.Color.White
        End If
    End Sub

    Sub Clone_PO()
        EnforceConstraints(False)

        For Each COLUMN_NAME As String In New String() {"WHSE_CODE", "FOB_CMT", "PO_CONTACT", "PO_NOTES",
                                                        "PORT_CODE_ORIG", "PORT_CODE_DEST", "COST_CODE",
                                                        "PO_DATE_CANCEL", "PO_FOB_DESC", "PO_SHIP_VIA",
                                                        "PO_CARTON_MARKS", "TERM_CODE", "PO_MESSAGE",
                                                        "LABEL_RESP_CODE",
                                                        "PO_COMM_PAYABLE_TO_BRKR", "PO_COMM_CHGBACK_TO_SUPP",
                                                        "PO_COMM_PCT", "PO_HAS_PPK"}
            rowPOTORDR1.Item(COLUMN_NAME) = rowPOTORDR1_clone(COLUMN_NAME)
        Next

        rowPOTORDR1.Item("PO_STATUS") = "O"
        rowPOTORDR1.Item("PO_HDR_CTR_REV") = 0
        Absx1.dteFor("PO_DATE_SHIP_BY").DateTime = Now.Date
        '  rowPOTORDR1.Item("PO_DATE_SHIP_BY") = Now.Date
        '  rowPOTORDR1.Item("PO_DATE_ETA") = 0
        'Calculate_ETD_to_ETA()

        For Each TABLE_NAME As String In New String() {"POTORDR2", "POTORDR6", "POTORDR7", "POTORDR8", "POTORDRN"}
            CLone_PO_Table(TABLE_NAME)
        Next


        dst.Tables("POTORDRR").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR2"), New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
            Create_POTORDRR(row.Item("STYLE_CODE"), row.Item("COLOR_CODE"))
        Next
        Sort_grdColumns(grdPOTORDRR, "STYLE_CODE,COLOR_CODE")

        EnforceConstraints(True)
    End Sub

    Sub CLone_PO_Table(TABLE_NAME As String)
        ASCMAIN1.sql = "Select * from " & TABLE_NAME & " where PO_ORDER_NO = '" & PO_ORDER_NO_clone & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim row2 As DataRow = dst.Tables(TABLE_NAME).NewRow
            row2.Item("PO_ORDER_NO") = PO_ORDER_NO

            Dim COLUMN_NAMEs() As String = {"*"}

            Select Case TABLE_NAME
                Case "POTORDR2"
                    COLUMN_NAMEs = {"PO_ORDER_LNO", "STYLE_CODE", "COLOR_CODE", "PO_QTY_ORD", "PO_QTY_OPN",
                                    "PO_COST", "PO_QTY_UOM", "PO_COST_VCOST",
                                    "STYLE_NOTES", "SUB_UNIT_PACK_QTY",
                                    "PO_COST_VCOST_DZ", "PO_COST_OTHER", "PO_COST_COMM", "PO_COST_QUOTA",
                                    "DFQUOTA", "PO_COST_BUFFER", "CARTON_PACK_QTY", "INNER_PACK_QTY", "PO_LINE_NOTE_INT"}
                    row2.Item("PO_STATUS") = "O"
                    row2.Item("PO_DATE_SHIP_BY") = Absx1.dteFor("PO_DATE_SHIP_BY").Value
                    row2.Item("PO_DATE_ETA") = Absx1.dteFor("PO_DATE_ETA").Value
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", row.Item("STYLE_CODE"))
                    row2.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                    row2.Item("CASE_CUBE") = rowICTSTYL1.Item("CASE_CUBE")

                Case "POTORDRN"
                    COLUMN_NAMEs = {"PO_ORDER_CLNO", "PO_ORDER_COMMENT"}
            End Select

            If COLUMN_NAMEs.Length = 1 AndAlso COLUMN_NAMEs(0) = "*" Then
                For C As Integer = 1 To row.Table.Columns.Count - 1
                    row2.Item(C) = row.Item(C)
                Next
            Else
                For Each COLUMN_NAME As String In COLUMN_NAMEs
                    row2.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
            End If

            If dst.Tables(TABLE_NAME).Columns.Contains("INIT_DATE") Then
                row2.Item("INIT_DATE") = DATETIME_STAMP
                row2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            End If

            dst.Tables(TABLE_NAME).Rows.Add(row2)
        Next
    End Sub

    Private Sub grdPOTORDRL_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    Private Sub grdPOTORDRL_InitializeLayout_1(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTORDRL.InitializeLayout

    End Sub

    Private Sub grdTATEVNT1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATEVNT1.InitializeLayout

    End Sub

    Private Sub SplitContainer9_Panel1_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer9.Panel1.Paint

    End Sub

    Private Sub UltraButton1_Click(sender As Object, e As EventArgs) Handles cmdBuildURL.Click
        UPDATE_LABEL_URL()

    End Sub
    Sub UPDATE_LABEL_URL()

        Dim DESCHASH As String = ""


        Dim PDF_FN As String = ""


        Dim ACCESS_CODE As String = ""
        Dim LABELURL As String = ROWs("POTPARM1").Item("PO_PARM_LABEL_PRINT_API") & ""


        ACCESS_CODE = Now.Ticks Mod 1000000

        DESCHASH = ACCESS_CODE & PO_ORDER_NO


        '  DESCHASH = Replace(DESCHASH, " ", "")
        '  DESCHASH = Replace(DESCHASH, "/", "")
        '  DESCHASH = Replace(DESCHASH, ".", "")
        '  DESCHASH = Replace(DESCHASH, ",", "")
        '  DESCHASH = Replace(DESCHASH, "&", "")

        Dim strToHash As String = ASCMAIN1.Get_Hash(DESCHASH)

        ' PDF_FN = LABELURL & strToHash
        PDF_FN = strToHash


        rowPOTORDR1.Item("PO_LABEL_URL_ACTIVE") = Absx1.dteFor("PO_DATE_ORDERED").Value & ""
        rowPOTORDR1.Item("PO_LABEL_URL_EXPIRES") = Absx1.dteFor("PO_DATE_SHIP_BY").Value & ""
        rowPOTORDR1.Item("PO_LABEL_URL") = PDF_FN
        rowPOTORDR1.Item("PO_LABEL_ACCESS_CODE") = ACCESS_CODE
        cmdBuildURL.Visible = False


    End Sub


    Private Sub cmdPRINTLABEL_Click(sender As Object, e As EventArgs) Handles cmdPRINTLABEL.Click

        '   My.Computer.Clipboard.SetText(rowPOTORDR1.Item("PO_LABEL_URL"))
        Clipboard.SetText(rowPOTORDR1.Item("PO_LABEL_ACCESS_CODE"))
        'Dim url As String = "http://api.regency-rib.com:8181/purchaseorder/labels/" & rowPOTORDR1.Item("PO_LABEL_URL")
        Dim url As String = ROWs("POTPARM1").Item("PO_PARM_LABEL_PRINT_API") & rowPOTORDR1.Item("PO_LABEL_URL")
        Process.Start(url)
    End Sub


    Sub Load_XLS(PO_XLS_NO As String)
        Dim walmartPO As Boolean = False
        WorkbookView1.GetLock()

        workbook = WorkbookView1.ActiveWorkbook

        dst.Tables("POTXLSF2").Rows.Clear()
        dst.Tables("POTXLSF1").Rows.Clear()

        For Each worksheet As SpreadsheetGear.IWorksheet In workbook.Worksheets

            Dim WORKSHEET_NAME As String = worksheet.Name

            Dim XLS_MAKER As String = ""
            Dim XLS_HANDLER As String = ""
            Dim XLS_DELIVERY As Date = Nothing
            Dim XLS_DELIVERY_DBL As Double = 0
            Dim XLS_DATE As String = ""

            ' validate the worksheet as a valid worksheet to pull PO info from 
            Dim worksheet_valid As Boolean = False
            Dim format As String = ""
            Dim r_start As Integer = 0
            Dim q_col As Integer = 0
            Dim firstCell As String = worksheet.Cells(0, 0).Value & ""
            If firstCell = "PURCHASE CONTRACT" And Len(WORKSHEET_NAME) < 15 And Not WORKSHEET_NAME.ToUpper.Contains("SHIPPING") Then

                If worksheet.Cells(9, 0).Value & "" = "Style No." _
                    And worksheet.Cells(9, 1).Value & "" = "Order #" _
                    And worksheet.Cells(9, 4).Value & "" = "Col#" Then
                    format = "1"
                    r_start = 10
                    q_col = 6

                    XLS_MAKER = worksheet.Cells(4, 1).Value & ""
                    XLS_HANDLER = worksheet.Cells(5, 1).Value & ""
                    XLS_DATE = worksheet.Cells(6, 1).Value & ""
                    If Not IsDate(XLS_DATE) Then
                        XLS_DATE = ""

                        XLS_DELIVERY_DBL = Val(worksheet.Cells(6, 1).Value & "")
                    End If

                ElseIf worksheet.Cells(10, 0).Value & "" = "Style No." _
                And worksheet.Cells(10, 1).Value & "" = "Order #" _
                And worksheet.Cells(10, 4).Value & "" = "Col#" Then
                    format = "1"
                    r_start = 11
                    q_col = 6

                    XLS_MAKER = worksheet.Cells(4, 1).Value & ""
                    XLS_HANDLER = worksheet.Cells(5, 1).Value & ""
                    XLS_DATE = worksheet.Cells(6, 1).Value & ""
                    If Not IsDate(XLS_DATE) Then
                        XLS_DATE = ""

                        XLS_DELIVERY_DBL = Val(worksheet.Cells(6, 1).Value & "")
                    End If
                ElseIf worksheet.Cells(11, 0).Value & "" = "Style No." _
                And worksheet.Cells(11, 1).Value & "" = "Order #" _
                And worksheet.Cells(11, 4).Value & "" = "Col#" Then
                    format = "1"
                    r_start = 12
                    q_col = 6

                    XLS_MAKER = worksheet.Cells(4, 1).Value & ""
                    XLS_HANDLER = worksheet.Cells(5, 1).Value & ""
                    XLS_DATE = Get_Delivery_Date(worksheet.Cells(6, 1).Text & "")
                    If Not IsDate(XLS_DATE) Then
                        XLS_DATE = ""

                        XLS_DELIVERY_DBL = Val(worksheet.Cells(6, 1).Value & "")
                    End If

                ElseIf worksheet.Cells(11, 0).Value & "" = "Order No." _
                    And worksheet.Cells(11, 3).Value & "" = "Style No." _
                    And worksheet.Cells(12, 4).Value & "" = "Col code" Then
                    format = "2"
                    r_start = 12
                    q_col = 6

                    XLS_MAKER = worksheet.Cells(2, 1).Value & ""
                    XLS_HANDLER = worksheet.Cells(3, 1).Value & ""
                    XLS_DELIVERY_DBL = worksheet.Cells(4, 1).Value

                ElseIf worksheet.Cells(11, 0).Value & "" = "Style No." _
                    And worksheet.Cells(11, 1).Value & "" = "Order No." _
                    And worksheet.Cells(14, 4).Value & "" = "Col code" Then
                    format = "3"
                    Dim eMsg As String = ""
                    Dim headerRow As Integer = 11
                    Dim headerRowDetails As Integer = 0
                    Dim poDetailRow As Integer = 0
                    Dim cMap As New Dictionary(Of String, Integer)
                    Dim POs As New Dictionary(Of String, String)

                    'only one order/ sheet
                    'style/orderno row = header row + 2

                    XLS_MAKER = worksheet.Cells(7, 1).Value & ""
                    XLS_HANDLER = worksheet.Cells(8, 1).Value & ""
                    XLS_DATE = Get_Delivery_Date(worksheet.Cells(9, 1).Text & "")
                    If Not IsDate(XLS_DATE) Then
                        XLS_DATE = ""
                        XLS_DELIVERY_DBL = 1
                    Else
                        XLS_DELIVERY = XLS_DATE
                    End If

                    cMap.Add("colPO", 1)
                    cMap.Add("colColors", 4)
                    cMap.Add("colQtyDz", 7)
                    cMap.Add("colCosts", 3)
                    cMap.Add("colSizeScale", 0) ' I DON'T THINK THIS MATTERS FOR THIS FORMAT

                    For j As Integer = headerRow + 1 To 100
                        Dim orderNo As String = CStr(worksheet.Cells(j, cMap("colPO")).Value & "")
                        Dim styleNo As String = CStr(worksheet.Cells(j, cMap("colPO") - 1).Value & "")

                        If orderNo <> "" And styleNo <> "" And orderNo.Length > 2 And styleNo.Length > 2 Then
                            If Mid(orderNo, 1, 2) = WORKSHEET_NAME And Mid(styleNo, Len(styleNo) - 1, 2) = WORKSHEET_NAME _
                            And orderNo.Length <= 20 And styleNo.Length <= 12 Then
                                POs.Add(orderNo, styleNo)
                                headerRowDetails = j + 1
                            End If
                        End If
                    Next
                    If POs.Count < 1 Then
                        eMsg &= "Couldn't find any PO's to process." & vbCrLf
                    End If

                    For Each poData As KeyValuePair(Of String, String) In POs
                        Dim poRef As String = poData.Key
                        Dim styleCode As String = poData.Value
                        Dim XLS_ORDER_LNO As Integer = 0
                        Dim VEND_CODE As String = "AT" ' NEED TO FIGURE OUT HOW TO DETERMINE THE VENDOR CODE
                        Dim FACTORY_CODE As String = ""
                        If XLS_MAKER <> "" Then
                            FACTORY_CODE = Get_Factory_Code(XLS_MAKER)
                        End If
                        Dim rowPOTXLSF1 As DataRow = dst.Tables("POTXLSF1").NewRow
                        rowPOTXLSF1.Item("PO_XLS_NO") = PO_XLS_NO
                        rowPOTXLSF1.Item("XLS_ORDER_NO") = poRef
                        rowPOTXLSF1.Item("XLS_MAKER") = XLS_MAKER
                        rowPOTXLSF1.Item("XLS_HANDLER") = XLS_HANDLER
                        rowPOTXLSF1.Item("XLS_DELIVERY") = XLS_DELIVERY
                        rowPOTXLSF1.Item("VEND_CODE") = VEND_CODE
                        rowPOTXLSF1.Item("FACTORY_CODE") = FACTORY_CODE
                        rowPOTXLSF1.Item("XLS_ORDER_STATUS") = "0"
                        dst.Tables("POTXLSF1").Rows.Add(rowPOTXLSF1)

                        Dim FILENAME As String = grdPOTXLSF0.ActiveRow.Cells("FILENAME").Value & ""
                        Dim FILENAME_NEW As String = Replace(FILENAME, "XLS_New\", "XLS_Archived\" & PO_XLS_NO & "_" & poRef & "_")
                        rowPOTXLSF1.Item("XLS_FILENAME") = FILENAME_NEW

                        If FILENAME <> "" Then
                            Try
                                My.Computer.FileSystem.CopyFile(FILENAME, FILENAME_NEW)
                            Catch ex As Exception

                            End Try
                        End If

                        'Costs
                        Dim startRow As Integer = headerRowDetails + 1
                        Dim costCol As Integer = cMap("colCosts")
                        Dim vCost_dz As Double = worksheet.Cells(startRow, costCol).Value
                        Dim vCost_un As Double = 0
                        Dim poCostMatls_dz As Double = 0
                        Dim poCostMatls_un As Double = 0
                        Dim poCostSubTot_un As Double = 0
                        'Changed From 2.5 to 2.0  Per Anna on 7/7/19 - WR.
                        'Changed From 2.0 to 2.5 Per Anna on 7/17/19 - WR.
                        'Changed From 2.5 to 2.0 Per Anna on 1/26/21 - WR.
                        Dim commPct As Double = 2.0 'currently hard coded for AT
                        Dim poCostComm_un As Double = 0
                        Dim poCost_un As Double = 0
                        Dim poCost_dz As Double = 0
                        Dim mCostItem As Integer = startRow + 1
                        Do While CStr(worksheet.Cells(mCostItem, costCol - 1).Value & "") <> ""
                            poCostMatls_dz += Val(worksheet.Cells(mCostItem, costCol).Value & "")
                            mCostItem += 1
                        Loop

                        'Colors
                        Dim poColors As New Dictionary(Of String, Integer)
                        Dim colorCol As Integer = cMap("colColors")
                        Dim r As Integer = headerRowDetails + 2
                        Do While CStr(worksheet.Cells(r, colorCol).Value & "") <> ""
                            Dim cCode As String = CStr(worksheet.Cells(r, colorCol).Value & "")
                            If InStr(cCode, "#") > 0 Then
                                cCode = Mid(cCode, InStr(cCode, "#") + 1, 3)
                            End If
                            If Validate_Color(cCode) Then
                                poColors.Add(cCode, r)
                            Else
                                eMsg &= "Invalid Color for PO: " & poRef & " (" & cCode & ")" & vbCrLf
                            End If
                            r += 1
                        Loop

                        'Style Colors
                        Dim scRow As Integer = startRow
                        Dim colorsProcessed As Integer = 0
                        For Each poStyleColor As KeyValuePair(Of String, Integer) In poColors
                            Dim cCode As String = poStyleColor.Key
                            Dim cRow As Integer = poStyleColor.Value
                            Dim sc_dzs As Double = worksheet.Cells(cRow, cMap("colQtyDz")).Value
                            Dim sc_uns As Double = sc_dzs * 12

                            colorsProcessed += 1

                            Dim rowICTSTYL1 As DataRow = Validate_Style(styleCode, False)
                            If rowICTSTYL1 IsNot Nothing Then
                                Dim scUnits As Double = sc_uns
                                Dim subUnitPackQty As Integer = IIf(Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") = 0, 1, rowICTSTYL1.Item("SUB_UNIT_PACK_QTY"))
                                vCost_un = vCost_dz / (12 / subUnitPackQty)
                                poCostMatls_un = poCostMatls_dz / (12 / subUnitPackQty)
                                poCostSubTot_un = vCost_un + poCostMatls_un
                                poCostComm_un = poCostSubTot_un * (commPct / 100)
                                poCost_un = poCostSubTot_un + poCostComm_un
                                poCost_dz = poCost_un / (12 / subUnitPackQty)
                                Dim rowPOTXLSF2 As DataRow = dst.Tables("POTXLSF2").NewRow
                                rowPOTXLSF2.Item("PO_XLS_NO") = PO_XLS_NO
                                rowPOTXLSF2.Item("XLS_ORDER_NO") = poRef
                                XLS_ORDER_LNO += 1
                                rowPOTXLSF2.Item("XLS_ORDER_LNO") = XLS_ORDER_LNO
                                rowPOTXLSF2.Item("XLS_STYLE_CODE") = styleCode
                                rowPOTXLSF2.Item("XLS_COLOR_CODE") = cCode
                                rowPOTXLSF2.Item("STYLE_CODE") = rowICTSTYL1.Item("STYLE_CODE")
                                rowPOTXLSF2.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_CODE")
                                rowPOTXLSF2.Item("SUB_UNIT_PACK_QTY") = subUnitPackQty
                                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", cCode)
                                If rowICTCOLR1 IsNot Nothing Then
                                    rowPOTXLSF2.Item("COLOR_CODE") = rowICTCOLR1.Item("COLOR_CODE")
                                    rowPOTXLSF2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
                                End If
                                rowPOTXLSF2.Item("PO_QTY") = scUnits
                                rowPOTXLSF2.Item("PO_COST") = poCostSubTot_un
                                rowPOTXLSF2.Item("PO_OTHER") = poCostMatls_un
                                rowPOTXLSF2.Item("PO_OTHER2") = 0
                                rowPOTXLSF2.Item("PO_COST_DZ") = vCost_dz
                                rowPOTXLSF2.Item("PO_OTHER_DZ") = poCostMatls_dz
                                rowPOTXLSF2.Item("PO_OTHER2_DZ") = 0
                                rowPOTXLSF2.Item("PO_DZS") = scUnits / (12 / subUnitPackQty)
                                dst.Tables("POTXLSF2").Rows.Add(rowPOTXLSF2)
                            Else
                                eMsg &= "Invalid Style: " & styleCode & vbCrLf
                            End If
                        Next
                    Next

                    Update_Record_TDA("POTXLSF1")
                    Update_Record_TDA("POTXLSF2")

                End If

                If format = "1" Then
                    For c As Integer = q_col To 50
                        If CStr(worksheet.Cells(r_start - 1, c).Value & "").ToUpper = "TTL QTY" Then
                            q_col = c
                            Exit For
                        End If
                    Next
                End If

                If format <> "" Then worksheet_valid = True

            ElseIf firstCell = "Buyer" Then 'Walmart
                walmartPO = True
                Dim eMsg As String = ""
                Dim headerRow As Integer = 0
                Dim poDetailRow As Integer = 0
                Dim cMap As New Dictionary(Of String, Integer)
                Dim POs As New Dictionary(Of String, Integer)

                'Header fields & Column Map
                For r As Integer = 1 To 20
                    If CStr(worksheet.Cells(r, 0).Value & "") = "PO#" Then
                        headerRow = r
                    ElseIf CStr(worksheet.Cells(r, 0).Value & "") = "Factory :" Then
                        XLS_MAKER = worksheet.Cells(r, 1).Value & ""
                    ElseIf InStr(CStr(worksheet.Cells(r, 0).Value & ""), "Delivery") > 0 Then
                        'XLS_DATE = CDate(Mid(dDate, 1, InStr(dDate, ",") + 5))
                        XLS_DATE = Get_Delivery_Date(worksheet.Cells(r, 1).Value & "")
                        If Not IsDate(XLS_DATE) Then
                            XLS_DATE = ""
                            XLS_DELIVERY_DBL = 1
                        Else
                            XLS_DELIVERY = XLS_DATE
                        End If
                    End If
                Next

                For c As Integer = 0 To 20
                    Dim trCell As String = worksheet.Cells(0, c).Value & ""
                    Dim hrCell As String = worksheet.Cells(headerRow, c).Value & ""
                    If CStr(hrCell) = "PO#" Then
                        cMap.Add("colPO", c)
                    ElseIf InStr(hrCell, "Color Code") <> 0 Then
                        cMap.Add("colColors", c)
                    ElseIf CStr(hrCell) = "Qtn (in dz.)" Then
                        cMap.Add("colQtyDz", c)
                    ElseIf CStr(hrCell) = "Size Scale" Then
                        cMap.Add("colCosts", c - 1)
                        cMap.Add("colSizeScale", c)
                    ElseIf InStr(trCell, "Maker") > 0 Or (InStr(trCell, "Factory") > 0 And XLS_MAKER = "") Then
                        XLS_MAKER = Trim(Mid(trCell, InStr(trCell, ":") + 1, Len(trCell) - InStr(trCell, ":")))
                    ElseIf InStr(trCell, "Handled by") > 0 Then
                        XLS_HANDLER = Mid(trCell, InStr(trCell, "by") + 3)
                    Else
                    End If
                Next

                'how many POs
                For j As Integer = headerRow + 1 To 100
                    If InStr(CStr(worksheet.Cells(j, cMap("colPO")).Value & ""), "Order") > 0 And
                        (CStr(worksheet.Cells(j - 1, cMap("colPO")).Value & "").Length > 0 AndAlso Mid(worksheet.Cells(j - 1, cMap("colPO")).Value, 1, 2) = "WM") Then
                        POs.Add(CStr(worksheet.Cells(j - 1, cMap("colPO")).Value & ""), j - 1)
                        poDetailRow = j - 1
                    End If
                Next
                If POs.Count < 1 Then
                    eMsg &= "Couldn't find any PO's to process." & vbCrLf
                End If

                For Each poData As KeyValuePair(Of String, Integer) In POs
                    Dim poRef As String = poData.Key
                    Dim XLS_ORDER_LNO As Integer = 0
                    Dim VEND_CODE As String = "AT" ' NEED TO FIGURE OUT HOW TO DETERMINE THE VENDOR CODE
                    Dim FACTORY_CODE As String = ""
                    If XLS_MAKER <> "" Then
                        FACTORY_CODE = Get_Factory_Code(XLS_MAKER)
                    End If
                    Dim rowPOTXLSF1 As DataRow = dst.Tables("POTXLSF1").NewRow
                    rowPOTXLSF1.Item("PO_XLS_NO") = PO_XLS_NO
                    rowPOTXLSF1.Item("XLS_ORDER_NO") = poRef
                    rowPOTXLSF1.Item("XLS_MAKER") = XLS_MAKER
                    rowPOTXLSF1.Item("XLS_HANDLER") = XLS_HANDLER
                    rowPOTXLSF1.Item("XLS_DELIVERY") = XLS_DELIVERY
                    rowPOTXLSF1.Item("VEND_CODE") = VEND_CODE
                    rowPOTXLSF1.Item("FACTORY_CODE") = FACTORY_CODE
                    rowPOTXLSF1.Item("XLS_ORDER_STATUS") = "0"
                    dst.Tables("POTXLSF1").Rows.Add(rowPOTXLSF1)

                    Dim FILENAME As String = grdPOTXLSF0.ActiveRow.Cells("FILENAME").Value & ""
                    Dim FILENAME_NEW As String = Replace(FILENAME, "XLS_New\", "XLS_Archived\" & PO_XLS_NO & "_" & poRef & "_")
                    rowPOTXLSF1.Item("XLS_FILENAME") = FILENAME_NEW

                    If FILENAME <> "" Then
                        Try
                            My.Computer.FileSystem.CopyFile(FILENAME, FILENAME_NEW)
                        Catch ex As Exception

                        End Try
                    End If

                    'Costs
                    Dim startRow As Integer = poData.Value
                    Dim costCol As Integer = cMap("colCosts")
                    Dim vCost_dz As Double = worksheet.Cells(startRow, costCol).Value
                    Dim vCost_un As Double = 0
                    Dim poCostMatls_dz As Double = 0
                    Dim poCostMatls_un As Double = 0
                    Dim poCostSubTot_un As Double = 0
                    'Changed from 2.5 to 2.0 Per Anna on 7/7/19 - WR.
                    ' Changed from 2.0 to 2.5 Per Anna on 7/17/19 - WR.
                    ' Changed from 2.5 to 2.0 Per Anna on 1/26/21 - WR.
                    Dim commPct As Double = 2.0 'currently hard coded for AT
                    Dim poCostComm_un As Double = 0
                    Dim poCost_un As Double = 0
                    Dim poCost_dz As Double = 0
                    Dim mCostItem As Integer = startRow + 1
                    Do While CStr(worksheet.Cells(mCostItem, costCol - 1).Value & "") <> ""
                        poCostMatls_dz += worksheet.Cells(mCostItem, costCol).Value
                        mCostItem += 1
                    Loop

                    'Colors
                    Dim poColors As New Dictionary(Of String, Integer)
                    Dim colorCol As Integer = cMap("colColors")
                    Dim r As Integer = startRow
                    Do While CStr(worksheet.Cells(r, colorCol).Value & "") <> ""
                        Dim cCode As String = CStr(worksheet.Cells(r, colorCol).Value & "")
                        If InStr(cCode, "(") > 0 Then
                            cCode = Trim(Mid(cCode, InStr(cCode, "(") + 1, (InStr(cCode, ")") - InStr(cCode, "(")) - 1))
                        End If
                        If Validate_Color(cCode) Then
                            poColors.Add(cCode, r)
                        Else
                            eMsg &= "Invalid Color for PO: " & poRef & " (" & cCode & ")" & vbCrLf
                        End If
                        r += 1
                    Loop

                    'Style Colors
                    Dim scRow As Integer = startRow
                    Dim colorsProcessed As Integer = 0
                    For Each poStyleColor As KeyValuePair(Of String, Integer) In poColors
                        Dim cCode As String = poStyleColor.Key
                        Dim cRow As Integer = poStyleColor.Value
                        Dim sc_dz_tot As Double = worksheet.Cells(cRow, cMap("colQtyDz")).Value
                        Dim sc_un_tot As Double = worksheet.Cells(cRow, cMap("colQtyDz") - 1).Value
                        Dim sCol As Integer = cMap("colSizeScale") + 1
                        Dim scUnitsCheck As Double = 0
                        Do While InStr(worksheet.Cells(scRow, sCol - 1).Value & "", cCode) < 1 And colorsProcessed <= poColors.Count And scRow < 100
                            scRow += 1
                        Loop
                        colorsProcessed += 1
                        Do While Val(worksheet.Cells(scRow, sCol).Value & "") > 0
                            Dim styleCode As String = CStr(worksheet.Cells(scRow - 1, sCol).Value & "")
                            Dim rowICTSTYL1 As DataRow = Validate_Style(styleCode, False)
                            If rowICTSTYL1 IsNot Nothing Then
                                Dim scUnits As Double = worksheet.Cells(scRow, sCol).Value
                                Dim subUnitPackQty As Integer = IIf(Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") = 0, 1, rowICTSTYL1.Item("SUB_UNIT_PACK_QTY"))
                                vCost_un = vCost_dz / (12 / subUnitPackQty)
                                poCostMatls_un = poCostMatls_dz / (12 / subUnitPackQty)
                                poCostSubTot_un = vCost_un + poCostMatls_un
                                poCostComm_un = poCostSubTot_un * (commPct / 100)
                                poCost_un = poCostSubTot_un + poCostComm_un
                                poCost_dz = poCost_un / (12 / subUnitPackQty)
                                Dim rowPOTXLSF2 As DataRow = dst.Tables("POTXLSF2").NewRow
                                rowPOTXLSF2.Item("PO_XLS_NO") = PO_XLS_NO
                                rowPOTXLSF2.Item("XLS_ORDER_NO") = poRef
                                XLS_ORDER_LNO += 1
                                rowPOTXLSF2.Item("XLS_ORDER_LNO") = XLS_ORDER_LNO
                                rowPOTXLSF2.Item("XLS_STYLE_CODE") = styleCode
                                rowPOTXLSF2.Item("XLS_COLOR_CODE") = cCode
                                rowPOTXLSF2.Item("STYLE_CODE") = rowICTSTYL1.Item("STYLE_CODE")
                                rowPOTXLSF2.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_CODE")
                                rowPOTXLSF2.Item("SUB_UNIT_PACK_QTY") = subUnitPackQty
                                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", cCode)
                                If rowICTCOLR1 IsNot Nothing Then
                                    rowPOTXLSF2.Item("COLOR_CODE") = rowICTCOLR1.Item("COLOR_CODE")
                                    rowPOTXLSF2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
                                End If
                                rowPOTXLSF2.Item("PO_QTY") = scUnits
                                rowPOTXLSF2.Item("PO_COST") = poCostSubTot_un
                                rowPOTXLSF2.Item("PO_OTHER") = poCostMatls_un
                                rowPOTXLSF2.Item("PO_OTHER2") = 0
                                rowPOTXLSF2.Item("PO_COST_DZ") = vCost_dz
                                rowPOTXLSF2.Item("PO_OTHER_DZ") = poCostMatls_dz
                                rowPOTXLSF2.Item("PO_OTHER2_DZ") = 0
                                rowPOTXLSF2.Item("PO_DZS") = scUnits / (12 / subUnitPackQty)
                                dst.Tables("POTXLSF2").Rows.Add(rowPOTXLSF2)
                                scUnitsCheck += scUnits
                            Else
                                eMsg &= "Invalid Style: " & styleCode & vbCrLf
                            End If
                            sCol += 1
                        Loop
                        If scUnitsCheck <> sc_un_tot Then
                            eMsg &= "Units do not foot for PO: " & poRef & ", Color: " & cCode & "." & vbCrLf
                        End If
                    Next
                Next

                If eMsg <> "" Then
                    Dim errorMsg As String = "Absolution encountered the following errors while attempting to import: " & vbCrLf & vbCrLf & eMsg
                    MsgBox(errorMsg, vbOKOnly + MsgBoxStyle.Critical, "Import Error")
                End If

                Update_Record_TDA("POTXLSF1")
                Update_Record_TDA("POTXLSF2")

            End If

            If worksheet_valid And Not (walmartPO Or format = "3") Then

                Dim XLS_ORDER_NO As String = ""

                If XLS_DATE = "" Then
                    Dim dt0 As Date = "01/01/1900"
                    XLS_DELIVERY = dt0.AddDays(Val(XLS_DELIVERY_DBL) - 2)
                Else
                    XLS_DELIVERY = XLS_DATE
                End If


                Dim VEND_CODE As String = "AT" ' NEED TO FIGURE OUT HOW TO DETERMINE THE VENDOR CODE
                Dim FACTORY_CODE As String = ""
                If XLS_MAKER <> "" Then
                    FACTORY_CODE = Get_Factory_Code(XLS_MAKER)
                End If
                Dim XLS_ORDER_LNO As Int32 = 0
                Dim XLS_SC As New Dictionary(Of String, Integer)

                For r As Integer = r_start To 100

                    Dim AX As String = worksheet.Cells(r, 0).Value & ""
                    Dim BX As String = worksheet.Cells(r, 1).Value & ""
                    Dim CX As String = worksheet.Cells(r, 2).Value & ""
                    Dim DX As String = worksheet.Cells(r, 3).Value & ""
                    Dim DX_OTHER As String = worksheet.Cells(r + 1, 3).Value & ""
                    Dim XLSEX As String = worksheet.Cells(r, 4).Value & ""
                    Dim FX As String = worksheet.Cells(r, 5).Value & ""
                    Dim GX As String = worksheet.Cells(r, q_col).Value & ""

                    If CX = "Color & Code / Style #" Or (AX <> "" And Not AX.Contains(" ") And AX.Length <= 12) Then

                        Select Case format
                            Case "1"

                                Dim AX1_STYLE As String = AX
                                Dim BX1_ORDER As String = BX
                                Dim DX1_COST As String = DX
                                Dim EX1_COLOR As String = XLSEX
                                Dim GX1_QTY As String = GX

                                Dim DX1_OTHER As String = DX_OTHER

                                If AX1_STYLE = "Packing" Then
                                    r += 2 ' get past the row with Packing in col A, and then next row, which might be blank
                                    Do While (worksheet.Cells(r, 0).Value & "" <> "" Or worksheet.Cells(r, 5).Value & "" = "") And r < 100
                                        r += 1
                                    Loop
                                    Dim FX_TOTAL_OTHER2 As String = worksheet.Cells(r, 5).Value & ""
                                    If Val(FX_TOTAL_OTHER2) <> 0 Then
                                        Dim sql As String = "PO_XLS_NO = '" & PO_XLS_NO & "' and XLS_ORDER_NO = '" & XLS_ORDER_NO & "'"
                                        For Each row As DataRow In dst.Tables("POTXLSF2").Select("")
                                            row.Item("PO_OTHER2_DZ") = Val(FX_TOTAL_OTHER2)
                                            row.Item("PO_OTHER2") = Val(FX_TOTAL_OTHER2) / 12
                                        Next
                                    End If

                                ElseIf AX1_STYLE <> "" And BX1_ORDER <> "" And EX1_COLOR <> "" And GX1_QTY <> "" Then

                                    If XLS_ORDER_NO <> BX1_ORDER Then
                                        XLS_ORDER_NO = BX1_ORDER
                                        XLS_ORDER_LNO = 0
                                        Dim rowPOTXLSF1 As DataRow = dst.Tables("POTXLSF1").NewRow
                                        rowPOTXLSF1.Item("PO_XLS_NO") = PO_XLS_NO
                                        rowPOTXLSF1.Item("XLS_ORDER_NO") = XLS_ORDER_NO
                                        rowPOTXLSF1.Item("XLS_MAKER") = XLS_MAKER
                                        rowPOTXLSF1.Item("XLS_HANDLER") = XLS_HANDLER
                                        rowPOTXLSF1.Item("XLS_DELIVERY") = XLS_DELIVERY
                                        rowPOTXLSF1.Item("VEND_CODE") = VEND_CODE
                                        rowPOTXLSF1.Item("FACTORY_CODE") = FACTORY_CODE
                                        rowPOTXLSF1.Item("XLS_ORDER_STATUS") = "0"

                                        dst.Tables("POTXLSF1").Rows.Add(rowPOTXLSF1)

                                        Dim FILENAME As String = grdPOTXLSF0.ActiveRow.Cells("FILENAME").Value & ""
                                        Dim FILENAME_NEW As String = Replace(FILENAME, "XLS_New\", "XLS_Archived\" & PO_XLS_NO & "_" & XLS_ORDER_NO & "_")
                                        rowPOTXLSF1.Item("XLS_FILENAME") = FILENAME_NEW

                                        If FILENAME <> "" Then
                                            Try
                                                My.Computer.FileSystem.CopyFile(FILENAME, FILENAME_NEW)
                                            Catch ex As Exception

                                            End Try
                                        End If

                                    End If

                                    Dim another_color_on_next_row As Boolean = False
                                    Dim sameCost As Boolean = False
                                    Dim poCost_sc As Double = 0
                                    Dim poCostOther_sc As Double = 0

                                    Do


                                        Dim rowPOTXLSF2 As DataRow = dst.Tables("POTXLSF2").NewRow
                                        rowPOTXLSF2.Item("PO_XLS_NO") = PO_XLS_NO
                                        rowPOTXLSF2.Item("XLS_ORDER_NO") = XLS_ORDER_NO
                                        XLS_ORDER_LNO += 1
                                        rowPOTXLSF2.Item("XLS_ORDER_LNO") = XLS_ORDER_LNO
                                        rowPOTXLSF2.Item("XLS_STYLE_CODE") = AX1_STYLE
                                        rowPOTXLSF2.Item("XLS_COLOR_CODE") = EX1_COLOR

                                        XLS_SC.Add(EX1_COLOR, XLS_ORDER_LNO)

                                        Dim SUB_UNIT_PACK_QTY As Integer = 1
                                        rowPOTXLSF2.Item("SUB_UNIT_PACK_QTY") = SUB_UNIT_PACK_QTY

                                        dst.Tables("POTXLSF2").Rows.Add(rowPOTXLSF2)


                                        Dim STYLE_CODE As String = ""

                                        Dim row As DataRow = LookUp("ICTSTYL1", AX1_STYLE)
                                        If row IsNot Nothing Then
                                            STYLE_CODE = row.Item("STYLE_CODE")
                                            rowPOTXLSF2.Item("STYLE_CODE") = STYLE_CODE
                                            rowPOTXLSF2.Item("STYLE_DESC") = row.Item("STYLE_DESC")
                                            SUB_UNIT_PACK_QTY = Val(row.Item("SUB_UNIT_PACK_QTY") & "")
                                            If SUB_UNIT_PACK_QTY < 1 Then SUB_UNIT_PACK_QTY = 1
                                            rowPOTXLSF2.Item("SUB_UNIT_PACK_QTY") = SUB_UNIT_PACK_QTY
                                        End If

                                        Dim CC As String = Trim(Replace(EX1_COLOR, "#", ""))

                                        Dim row2 As DataRow = LookUp("ICTCOLR1", CC)
                                        If row2 IsNot Nothing Then
                                            rowPOTXLSF2.Item("COLOR_CODE") = row2.Item("COLOR_CODE")
                                            rowPOTXLSF2.Item("COLOR_DESC") = row2.Item("COLOR_DESC")
                                        End If

                                        Dim F As Integer = 1
                                        GX1_QTY = Trim(GX1_QTY)
                                        If GX1_QTY.Contains("dz") Then
                                            GX1_QTY = Trim(Replace(GX1_QTY, "dz", ""))
                                            F = 12
                                        End If

                                        F = 12 ' always interpret the qty as dz

                                        Dim Q As Int64 = Val(GX1_QTY) * F
                                        If Q <> 0 Then
                                            rowPOTXLSF2.Item("PO_QTY") = Q
                                        End If

                                        Dim PO_COST As Decimal = IIf(Not sameCost, Val(DX1_COST), poCost_sc)



                                        rowPOTXLSF2.Item("PO_COST") = PO_COST / F
                                        Dim PO_OTHER As Decimal = IIf(Not sameCost, Val(DX1_OTHER), poCostOther_sc)
                                        If Not sameCost Then
                                            Dim rr As Integer = 1
                                            Do While worksheet.Cells(r + 1 + rr, 3).Value & "" <> "" And worksheet.Cells(r + 1 + rr, 2).Value & "" <> ""
                                                Dim DX1_ANOTHER As String = worksheet.Cells(r + 1 + rr, 3).Value & ""
                                                PO_OTHER += Val(DX1_ANOTHER)
                                                rr += 1
                                            Loop
                                        End If


                                        Dim PO_OTHER2 As Decimal = 0

                                        rowPOTXLSF2.Item("PO_OTHER") = PO_OTHER / F
                                        rowPOTXLSF2.Item("PO_OTHER2") = PO_OTHER2 / F
                                        rowPOTXLSF2.Item("PO_COST_DZ") = PO_COST
                                        rowPOTXLSF2.Item("PO_OTHER_DZ") = PO_OTHER
                                        rowPOTXLSF2.Item("PO_OTHER2_DZ") = PO_OTHER2

                                        rowPOTXLSF2.Item("PO_DZS") = Q / 12


                                        another_color_on_next_row = False
                                        For check_r As Integer = 1 To 5 ' check the next 5 rows for the next color
                                            Dim AX_next_row As String = worksheet.Cells(r + check_r, 0).Value & ""
                                            Dim DX_next_row As String = worksheet.Cells(r + check_r + 1, 3).Value & ""
                                            Dim EX_next_row As String = worksheet.Cells(r + check_r, 4).Value & ""
                                            If EX_next_row.Length = EX1_COLOR.Length And (check_r = 1 Or AX_next_row = "") Then
                                                another_color_on_next_row = True
                                                If check_r = 1 Then
                                                    sameCost = True
                                                    poCost_sc = PO_COST
                                                    poCostOther_sc = PO_OTHER
                                                End If
                                                r += check_r
                                                XLSEX = worksheet.Cells(r, 4).Value & ""
                                                GX = worksheet.Cells(r, q_col).Value & ""
                                                EX1_COLOR = XLSEX
                                                DX1_OTHER = DX_next_row
                                                GX1_QTY = GX
                                                Exit For
                                            End If
                                        Next

                                    Loop While another_color_on_next_row

                                ElseIf worksheet.Cells(r, 2).Value & "" = "Color & Code / Style #" And Trim(worksheet.Cells(r + 1, 2).Value & "") = "Size" Then

                                    Dim XLS_ORDER_LNO_max As Integer = XLS_ORDER_LNO
                                    Dim c As Integer = 1
                                    Do While worksheet.Cells(r, 2 + c).Value & "" <> "" And worksheet.Cells(r, 2 + c).Value & "" <> "TTL/PCS"
                                        Dim STYLE_SZ As String = worksheet.Cells(r, 2 + c).Value & ""
                                        Dim rsz As Integer = 2
                                        Do While worksheet.Cells(r + rsz, 2 + c).Value & "" <> ""
                                            Dim XLS_COLOR As String = worksheet.Cells(r + rsz, 2).Value & ""
                                            XLS_COLOR = Mid(XLS_COLOR, 1, 5)
                                            If XLS_SC.ContainsKey(XLS_COLOR) Then
                                                Dim row0 As DataRow = dst.Tables("POTXLSF2").Rows.Find(New Object() {PO_XLS_NO, XLS_ORDER_NO, XLS_SC(XLS_COLOR)})
                                                Dim Q As Int64 = Val(worksheet.Cells(r + rsz, 2 + c).Value & "") '  * 12

                                                Dim rowSZ As DataRow = dst.Tables("POTXLSF2").NewRow
                                                rowSZ.ItemArray = row0.ItemArray
                                                XLS_ORDER_LNO += 1
                                                rowSZ.Item("XLS_ORDER_LNO") = XLS_ORDER_LNO
                                                Dim XLS_STYLE_CODE As String = Trim(Replace(STYLE_SZ, "-", ""))
                                                rowSZ.Item("XLS_STYLE_CODE") = XLS_STYLE_CODE

                                                Dim STYLE_CODE As String = ""
                                                Dim STYLE_DESC As String = ""

                                                Dim SUB_UNIT_PACK_QTY As Integer = 1
                                                Dim row As DataRow = LookUp("ICTSTYL1", XLS_STYLE_CODE)
                                                If row IsNot Nothing Then
                                                    STYLE_CODE = row.Item("STYLE_CODE")
                                                    rowSZ.Item("STYLE_CODE") = STYLE_CODE
                                                    STYLE_DESC = row.Item("STYLE_DESC")
                                                    rowSZ.Item("STYLE_DESC") = STYLE_DESC
                                                    SUB_UNIT_PACK_QTY = Val(row.Item("SUB_UNIT_PACK_QTY") & "")
                                                    If SUB_UNIT_PACK_QTY < 1 Then SUB_UNIT_PACK_QTY = 1
                                                    rowSZ.Item("SUB_UNIT_PACK_QTY") = SUB_UNIT_PACK_QTY
                                                End If


                                                rowSZ.Item("PO_QTY") = Q
                                                rowSZ.Item("PO_DZS") = Q / 12
                                                dst.Tables("POTXLSF2").Rows.Add(rowSZ)
                                            End If
                                            rsz += 1

                                        Loop
                                        c += 1
                                    Loop

                                    For XLS_ORDER_LNO_del As Integer = 1 To XLS_ORDER_LNO_max
                                        Dim row0 As DataRow = dst.Tables("POTXLSF2").Rows.Find _
                                                              (New Object() {PO_XLS_NO, XLS_ORDER_NO, XLS_ORDER_LNO_del})
                                        row0.Delete()
                                    Next

                                End If

                            Case "2"

                                Stop

                        End Select
                    End If
                Next

                Update_Record_TDA("POTXLSF1")
                Update_Record_TDA("POTXLSF2")

            End If
        Next

        WorkbookView1.ReleaseLock()
    End Sub

    Function Get_Delivery_Date(dDate As String) As Date
        Dim delDate_str As String = ""
        Dim delMonth As String = Mid(dDate, 1, InStr(dDate, " ") - 1)
        Dim delDay As String = Trim(Mid(dDate, InStr(dDate, ",") - 2, 2))
        Dim delYear As String = ""

        If (Len(dDate) - (InStr(dDate, ",") + 1)) < 4 Then
            delYear = "20" & Mid(dDate, InStr(dDate, ",") + 1, 2)
        Else
            delYear = Trim(Mid(dDate, InStr(dDate, ",") + 1, 5))
        End If

        delDate_str = delMonth & " " & delDay & ", " & delYear

        Return CDate(delDate_str)
    End Function
    Function Get_Factory_Code(maker As String) As String
        Dim fc As String = ""
        ASCMAIN1.sql = "Select * from ICTFACT1 where FACTORY_NAME_XLS = :PARM1"
        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", maker)
        If row IsNot Nothing Then
            fc = row.Item("FACTORY_CODE")
        End If
        Return fc
    End Function

    Private Sub grdPOTXLSF0_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTXLSF0.AfterRowActivate
        If ScreenMode Then Exit Sub

        EnforceConstraints(False)
        dst.Tables("POTXLSF1").Rows.Clear()
        dst.Tables("POTXLSF2").Rows.Clear()
        EnforceConstraints(True)

        Setup_PO_XLS()
    End Sub

    Sub Setup_PO_XLS()
        If grdPOTXLSF0.ActiveRow Is Nothing OrElse Not grdPOTXLSF0.ActiveRow.IsDataRow Then
            WorkbookView1.Visible = False

        Else

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Workbook")

            WorkbookView1.Visible = True
            Dim FILENAME As String = grdPOTXLSF0.ActiveRow.Cells("FILENAME").Value

            WorkbookView1.GetLock()
            WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)

            workbook = WorkbookView1.ActiveWorkbook
            worksheet = workbook.Worksheets(0)

            WorkbookView1.ReleaseLock()


            Dim PO_XLS_NO As String = grdPOTXLSF0.ActiveRow.Cells("PO_XLS_NO").Value

            EnforceConstraints(False)
            Fill_Records("POTXLSF1", PO_XLS_NO)
            Fill_Records("POTXLSF2", PO_XLS_NO)



            For Each rowPOTXLSF2 As DataRow In dst.Tables("POTXLSF2").Select("")

                Dim SUB_UNIT_PACK_QTY As Integer = 1
                rowPOTXLSF2.Item("SUB_UNIT_PACK_QTY") = SUB_UNIT_PACK_QTY

                Dim XLS_STYLE_CODE As String = rowPOTXLSF2.Item("XLS_STYLE_CODE")
                Dim XLS_COLOR_CODE As String = rowPOTXLSF2.Item("XLS_COLOR_CODE")
                Dim STYLE_CODE As String = ""

                Dim row As DataRow = LookUp("ICTSTYL1", XLS_STYLE_CODE)
                If row IsNot Nothing Then
                    STYLE_CODE = row.Item("STYLE_CODE")
                    rowPOTXLSF2.Item("STYLE_CODE") = STYLE_CODE
                    rowPOTXLSF2.Item("STYLE_DESC") = row.Item("STYLE_DESC")
                    SUB_UNIT_PACK_QTY = Val(row.Item("SUB_UNIT_PACK_QTY") & "")
                    If SUB_UNIT_PACK_QTY < 1 Then SUB_UNIT_PACK_QTY = 1
                    rowPOTXLSF2.Item("SUB_UNIT_PACK_QTY") = SUB_UNIT_PACK_QTY
                End If

                Dim CC As String = Trim(Replace(XLS_COLOR_CODE, "#", ""))

                Dim row2 As DataRow = LookUp("ICTCOLR1", CC)
                If row2 IsNot Nothing Then
                    rowPOTXLSF2.Item("COLOR_CODE") = row2.Item("COLOR_CODE")
                    rowPOTXLSF2.Item("COLOR_DESC") = row2.Item("COLOR_DESC")
                End If



                'Dim F As Integer = 12 ' always interpret the qty as dz


                'Dim PO_COST As Decimal = Val(rowPOTXLSF2.Item("PO_COST") & "")
                'Dim PO_OTHER As Decimal = Val(rowPOTXLSF2.Item("PO_OTHER") & "")

                'rowPOTXLSF2.Item("PO_COST_DZ") = PO_COST
                'rowPOTXLSF2.Item("PO_OTHER_DZ") = PO_OTHER
                '' rowPOTXLSF2.Item("PO_OTHER2_DZ") = PO_OTHER2

                'rowPOTXLSF2.Item("PO_DZS") = Q / 12
            Next

            EnforceConstraints(True)

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub


    Sub Refresh_XLS()
        Dim FOLDER As String = PO_PARM_AUTOPO_FOLDER & "XLS_New"
        If ASCMAIN1.Running_in_VS Then
            FOLDER = IIf(ASCMAIN1.USER_ID = "rdw", "E:\", "D:\dmp\VAN\") & "AutoPO\XLS_New\"
        End If

        'dst.Tables("POTXLSF0").Rows.Clear()
        Fill_Records("POTXLSF0")
        Dim PO_XLS_NOs As New List(Of String)

        Try

            For Each FILENAME As String In My.Computer.FileSystem.GetFiles(FOLDER)
                Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                If (fi.Extension.ToUpper = ".XLS" Or fi.Extension.ToUpper = ".XLSX") And Not fi.Name.StartsWith("~") Then
                    Dim PO_XLS_NO As String = ""
                    Dim sqlf As String = "FILENAME = '" & FILENAME & "'"
                    Dim rowPOTXLSF0s() As DataRow = dst.Tables("POTXLSF0").Select(sqlf)
                    If rowPOTXLSF0s.Length = 1 Then
                        PO_XLS_NO = rowPOTXLSF0s(0).Item("PO_XLS_NO")
                    Else
                        PO_XLS_NO = ASCMAIN1.Next_Control_No("POTXLSF0.PO_XLS_NO")
                        dst.Tables("POTXLSF0").Rows.Add(New Object() {PO_XLS_NO, FILENAME, fi.LastWriteTime, fi.Name, "0"})
                    End If
                    PO_XLS_NOs.Add(PO_XLS_NO)

                End If
            Next
        Catch ex As Exception

        End Try


        For Each rowPOTXLSF0 As DataRow In dst.Tables("POTXLSF0").Select("")
            Dim PO_XLS_NO As String = rowPOTXLSF0.Item("PO_XLS_NO")
            If Not PO_XLS_NOs.Contains(PO_XLS_NO) Then
                rowPOTXLSF0.Delete()
            End If
        Next
        Update_Record_TDA("POTXLSF0")

        Setup_PO_XLS()

    End Sub

    Private Sub grdPOTXLSF0_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdPOTXLSF0.AfterRowsDeleted

    End Sub

    Private Sub grdPOTXLSF0_BeforeRowsDeleted(sender As Object, e As UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTXLSF0.BeforeRowsDeleted

    End Sub

    Private Sub grdPOTXLSF0_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdPOTXLSF0.ClickCellButton

        Select Case e.Cell.Column.Key

            Case "DELETE"


                If MsgBox("OK to Delete this Workbook", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                    Dim PO_XLS_NO As String = e.Cell.Row.Cells("PO_XLS_NO").Value & ""
                    dst.Tables("POTXLSF0").AcceptChanges()
                    Dim rowPOTXLSF0 As DataRow = dst.Tables("POTXLSF0").Rows.Find(PO_XLS_NO)
                    rowPOTXLSF0.Item("PO_XLS_STATUS") = "1"
                    Dim FILENAME As String = rowPOTXLSF0.Item("FILENAME")
                    Dim FILENAME_NEW As String = Replace(FILENAME, "XLS_New\", "XLS_Deleted\" & PO_XLS_NO & "_")
                    My.Computer.FileSystem.CopyFile(FILENAME, FILENAME_NEW)
                    My.Computer.FileSystem.DeleteFile(FILENAME)
                    Update_Record_TDA("POTXLSF0")
                    rowPOTXLSF0.Delete()
                    dst.Tables("POTXLSF0").AcceptChanges()
                End If

            Case "Q"

                If dst.Tables("POTXLSF1").Rows.Count <> 0 Then
                    MsgBox("This Workbook was already Loaded into Import Queue", MsgBoxStyle.OkOnly, "Cannot Load a PO that has already been Loaded")
                    Exit Sub
                End If

                'Dim FOLDER As String = PO_PARM_AUTOPO_FOLDER & "XLS_New"
                'If ASCMAIN1.Running_in_VS Then
                '    FOLDER = "D:\dmp\VAN\AutoPO\XLS_New\"
                'End If


                'Dim FILENAME As String = e.Cell.Row.Cells("FILENAME").Value

                'dst.Tables("POTXLSF0").AcceptChanges()


                'Dim rowPOTXLSF0 As DataRow = dst.Tables("POTXLSF0").Rows.Find(PO_XLS_NO)

                ' My.Computer.FileSystem.MoveFile(FILENAME, FOLDER & "Queued\" & PO_XLS_NO & ".xls")

                Dim PO_XLS_NO As String = e.Cell.Row.Cells("PO_XLS_NO").Value
                Load_XLS(PO_XLS_NO)

        End Select

    End Sub

    Private Sub grdPOTXLSF0_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdPOTXLSF0.InitializeRow
        e.Row.Cells("Q").Value = "->"
    End Sub

    Private Sub grdPOTXLSF1_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdPOTXLSF1.AfterCellUpdate
        If grdPOTXLSF1.Tag & "" = "*" Then Exit Sub

        grdPOTXLSF1.Tag = "*"
        Select Case e.Cell.Column.Key
            Case "PO_COST", "PO_OTHER", "PO_OTHER2"
                Dim C As String = e.Cell.Column.Key & "_DZ"
                e.Cell.Row.Cells(C).Value = Val(e.Cell.Value & "") * 12
            Case "PO_COST_DZ", "PO_OTHER_DZ", "PO_OTHER2_DZ"
                Dim C As String = Replace(e.Cell.Column.Key, "_DZ", "")
                e.Cell.Row.Cells(C).Value = Val(e.Cell.Value & "") / 12
        End Select
        grdPOTXLSF1.Tag = ""
    End Sub

    Private Sub grdPOTXLSF1_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdPOTXLSF1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim sql_where As String = "" '  Get_Code_SQL("CMT_NO:0")
                grdClickCellButton(grdPOTXLSF1, sql_where)

            Case "COLOR_CODE"
                Dim sql_where As String = "" '  "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                grdClickCellButton(grdPOTXLSF1, sql_where)

            Case "FACTORY_CODE"
                Dim FACTORY_CODE_old As String = Trim(e.Cell.Value & "")
                Dim XLS_MAKER As String = Trim(e.Cell.Row.Cells("XLS_MAKER").Value & "")
                Dim sql_where As String = "" '  "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                grdClickCellButton(grdPOTXLSF1, sql_where)

                Dim FACTORY_CODE_new As String = e.Cell.Value & ""
                If FACTORY_CODE_new <> "" And FACTORY_CODE_new <> FACTORY_CODE_old And XLS_MAKER <> "" Then
                    Dim rowICTFACT1 As DataRow = Fill_Record("ICTFACT1", FACTORY_CODE_new)
                    If rowICTFACT1 IsNot Nothing Then
                        rowICTFACT1.Item("FACTORY_NAME_XLS") = XLS_MAKER
                        Update_Record_TDA("ICTFACT1")
                    End If
                    Update_Record_TDA("POTXLSF1")
                End If


            Case "MAKE_PO"
                Make_PO()

            Case "DELETE"
                If MsgBox("OK to Delete this PO", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    dst.Tables("POTXLSF1").AcceptChanges()
                    Dim PO_XLS_NO As String = e.Cell.Row.Cells("PO_XLS_NO").Value & ""
                    Dim XLS_ORDER_NO As String = e.Cell.Row.Cells("XLS_ORDER_NO").Value & ""
                    Dim rowPOTXLSF1 As DataRow = dst.Tables("POTXLSF1").Rows.Find(New String() {PO_XLS_NO, XLS_ORDER_NO})
                    rowPOTXLSF1.Item("XLS_ORDER_STATUS") = "D"
                    Update_Record_TDA("POTXLSF1")
                    rowPOTXLSF1.Delete()
                    dst.Tables("POTXLSF1").AcceptChanges()
                    If grdPOTXLSF1.Rows.Count = 0 Then
                        dst.Tables("POTXLSF0").AcceptChanges()
                        Dim rowPOTXLSF0 As DataRow = dst.Tables("POTXLSF0").Rows.Find(PO_XLS_NO)
                        rowPOTXLSF0.Item("PO_XLS_STATUS") = "1"
                        Dim FILENAME As String = rowPOTXLSF0.Item("FILENAME")
                        Dim FILENAME_NEW As String = Replace(FILENAME, "XLS_New\", "XLS_Deleted\" & PO_XLS_NO & "_")
                        My.Computer.FileSystem.CopyFile(FILENAME, FILENAME_NEW)
                        My.Computer.FileSystem.DeleteFile(FILENAME)
                        Update_Record_TDA("POTXLSF0")
                        rowPOTXLSF0.Delete()
                        dst.Tables("POTXLSF0").AcceptChanges()

                    End If
                End If
                'Case Else
                '    Dim sql_where As String = ""
                '    grdClickCellButton(grdPOTXLSF1, sql_where)
        End Select
    End Sub

    Function Validate_Code_PO(Codes() As String, TABLE_NAME As String, COLUMN_CAPTION As String, SOURCE As String) As String
        Dim EMsg As String = ""
        If Codes(0) = "" Then
            EMsg &= vbCr & SOURCE & ": " & COLUMN_CAPTION & " is not Specified"
        Else
            If LookUp(TABLE_NAME, Codes) Is Nothing Then
                EMsg &= vbCr & SOURCE & ": " & COLUMN_CAPTION & " '" & Join(Codes, "/") & "' does not exist"
            Else
                If TABLE_NAME = "APTVEND1" Then
                    If cdr.Item("VEND_TYPE") & "" <> "S" Then
                        EMsg &= vbCr & SOURCE & ": " & COLUMN_CAPTION & " " & Join(Codes, "/") & " is not a Supplier"
                    End If
                End If
            End If
        End If

        Return EMsg
    End Function

    Private Sub grdPOTXLSF1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTXLSF1.InitializeLayout

    End Sub

    Sub Make_PO()

        Dim grow As UltraWinGrid.UltraGridRow = grdPOTXLSF1.ActiveRow

        Dim EMsg As String = ""

        Dim PO_XLS_NO As String = grow.Cells("PO_XLS_NO").Value & ""
        Dim XLS_ORDER_NO As String = grow.Cells("XLS_ORDER_NO").Value & ""
        Dim XLS_HANDLER As String = grow.Cells("XLS_HANDLER").Value & ""
        Dim rowPOTXLSF1 As DataRow = dst.Tables("POTXLSF1").Rows.Find(New String() {PO_XLS_NO, XLS_ORDER_NO})

        Dim FACTORY_CODE As String = rowPOTXLSF1.Item("FACTORY_CODE") & ""
        EMsg &= Validate_Code_PO(New String() {FACTORY_CODE}, "ICTFACT1", "Factory", "Order " & XLS_ORDER_NO)

        Dim VEND_CODE As String = rowPOTXLSF1.Item("VEND_CODE") & ""
        EMsg &= Validate_Code_PO(New String() {VEND_CODE}, "APTVEND1", "Supplier", "Order " & XLS_ORDER_NO)

        For Each rowPOTXLSF2 As DataRow In rowPOTXLSF1.GetChildRows("POTXLSF1_POTXLSF2")
            Dim XLS_ORDER_LNO As Integer = Val(rowPOTXLSF2.Item("XLS_ORDER_LNO") & "")
            Dim STYLE_CODE As String = rowPOTXLSF2.Item("STYLE_CODE") & ""
            EMsg &= Validate_Code_PO(New String() {STYLE_CODE}, "ICTSTYL1", "Style", "Line " & XLS_ORDER_LNO)
            Dim COLOR_CODE As String = rowPOTXLSF2.Item("COLOR_CODE") & ""
            EMsg &= Validate_Code_PO(New String() {COLOR_CODE}, "ICTCOLR1", "Color", "Line " & XLS_ORDER_LNO)
            EMsg &= Validate_Code_PO(New String() {STYLE_CODE, COLOR_CODE}, "ICTSTYC1", "Style/Color", "Line " & XLS_ORDER_LNO)
        Next

        Dim PO_DATE_SHIP_BY As Date = rowPOTXLSF1.Item("XLS_DELIVERY")
        If System.Math.Abs(PO_DATE_SHIP_BY.Subtract(Now.Date).TotalDays) > 180 Then
            EMsg &= vbCr & "Please Check ETD - more than 180 days away from Today"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Import this Order")
        Else

            ASCMAIN1.sql = "Select * from POTORDR1 where VEND_CODE = :PARM1 and PO_REFERENCE = :PARM2 and PO_STATUS in ('O','C')"
            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {VEND_CODE, XLS_ORDER_NO})
            If row IsNot Nothing Then
                If MsgBox("PO " & row.Item("PO_ORDER_NO") & " already on file with this Reference" & vbCrLf & vbCrLf & "OK to Continue?",
                           MsgBoxStyle.OkCancel, "Verification") = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            Absx1.txtFor("VEND_CODE").Text = VEND_CODE
            Absx1.txtFor("PO_REFERENCE").Text = rowPOTXLSF1.Item("XLS_ORDER_NO") & ""
            Absx1.optFor("FOB_CMT").Value = "F"
            Click_Command("New")
            If ScreenMode Then

                ' rowPOTORDR1.Item("PO_DATE_SHIP_BY") = rowPOTXLSF1.Item("XLS_DELIVERY")
                Absx1.dteFor("PO_DATE_SHIP_BY").Value = rowPOTXLSF1.Item("XLS_DELIVERY")
                Absx1.txtFor("FACTORY_CODE").Text = FACTORY_CODE
                Absx1.txtFor("WHSE_CODE").Text = ROWs("POTPARM1").Item("PO_PARM_DEF_WHSE_CODE")

                Absx1.txtFor("PO_CONTACT").Text = XLS_HANDLER
                Me.rowPOTXLSF1 = rowPOTXLSF1
                Record_Event("AUTOPO", "PO Import from XLS", False, PO_XLS_NO)

                For Each rowPOTXLSF2 As DataRow In rowPOTXLSF1.GetChildRows("POTXLSF1_POTXLSF2")
                    Dim XLS_ORDER_LNO As Integer = Val(rowPOTXLSF2.Item("XLS_ORDER_LNO") & "")
                    Dim STYLE_CODE As String = rowPOTXLSF2.Item("STYLE_CODE") & ""
                    Dim COLOR_CODE As String = rowPOTXLSF2.Item("COLOR_CODE") & ""
                    Dim PO_QTY As Int64 = Val(rowPOTXLSF2.Item("PO_QTY") & "")
                    Dim PO_COST As Decimal = Val(rowPOTXLSF2.Item("PO_COST") & "")
                    Dim PO_OTHER As Decimal = Val(rowPOTXLSF2.Item("PO_OTHER") & "")
                    Dim PO_OTHER2 As Decimal = Val(rowPOTXLSF2.Item("PO_OTHER2") & "")
                    Dim PO_COST_DZ As Decimal = Val(rowPOTXLSF2.Item("PO_COST_DZ") & "")
                    Dim PO_OTHER_DZ As Decimal = Val(rowPOTXLSF2.Item("PO_OTHER_DZ") & "")
                    Dim PO_OTHER2_DZ As Decimal = Val(rowPOTXLSF2.Item("PO_OTHER2_DZ") & "")

                    If grdPOTORDR2.ActiveRow IsNot Nothing AndAlso grdPOTORDR2.ActiveRow.IsAddRow Then
                        grdPOTORDR2.ActiveRow = Nothing
                    End If
                    grdPOTORDR2.DisplayLayout.Bands(0).AddNew()
                    With grdPOTORDR2.ActiveRow
                        .Cells("STYLE_CODE").Value = STYLE_CODE
                        .Cells("COLOR_CODE").Value = COLOR_CODE
                        .Cells("PO_QTY_ORD").Value = PO_QTY
                        .Cells("PO_COST").Value = PO_COST
                        .Cells("PO_COST_VCOST_DZ").Value = PO_COST_DZ
                        .Cells("PO_COST_OTHER").Value = PO_OTHER_DZ + PO_OTHER2_DZ

                        If rowPOTORDR1.Item("VEND_CODE") = "AT" And (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
                            ' MAYBE WE SET THESE UP IN APTVEND1 AS OVERRIDES, BUT HOW TO OVERRIDE WITH A 0?
                            'Changed to 3 from 4 on 12/30/14 Per Anna - WR.
                            'Changed from 3 To 2.5 Per Anna 1/20/17 - WR.
                            'Changed from 2.5 to 2.0 Per Anna 7/7/19 - WR.
                            'Changed from 2.0 to 2.5 Per Anna 7/17/19 - WR.
                            'Changed from 2.5 to 2.0 Per Anna 1/26/21 - WR.
                            'Changed from 2.0 to 0.0 Per Anna 8/16/24 - WR.
                            .Cells("PO_COST_COMM").Value = 0.0
                            .Cells("PO_COST_BUFFER").Value = 1
                        Else
                            .Cells("PO_COST_COMM").Value = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_COMM") & "")
                            .Cells("PO_COST_BUFFER").Value = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_BUFFER") & "")
                        End If

                        .Update()
                        If .IsAddRow Then
                            grdPOTORDR2.ActiveRow = Nothing
                        End If
                    End With
                Next
                Sort_grdColumns(grdPOTORDR2, "PO_ORDER_LNO")
                ' WHEN PO IS UPDATED NEED TO DO THE FOLLOWING: rowPOTXLSF1.ITEM("") = "1"  UPDATE_RECORD_TDA("POTXLSF1")
            End If
        End If
    End Sub

    Private Sub grdPOTXLSF1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdPOTXLSF1.InitializeRow
        If e.Row.Band.Index = 0 Then
            Dim PO_REFERENCE As String = e.Row.Cells("XLS_ORDER_NO").Value & ""
            Dim VEND_CODE As String = e.Row.Cells("VEND_CODE").Value & ""
            ASCMAIN1.sql = "Select * from POTORDR1 where VEND_CODE = :PARM1 and PO_REFERENCE = :PARM2 and PO_STATUS in ('O','C')"
            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {VEND_CODE, PO_REFERENCE})
            If row IsNot Nothing Then
                e.Row.Cells("XLS_ORDER_NO").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("XLS_ORDER_NO").ToolTipText = "Duplicate PO " & row.Item("PO_ORDER_NO") & ". Status is " & row.Item("PO_STATUS")
            Else
                e.Row.Cells("XLS_ORDER_NO").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("XLS_ORDER_NO").ToolTipText = ""
            End If

        End If
    End Sub

    Private Sub cmdAutoPO_Click(sender As Object, e As EventArgs) Handles cmdAutoPO.Click

        Dim sqlMain As String = "SELECT ICTSTYL1.VEND_CODE, J.EDI_ARRIVAL_DATE, JF.WORKSHEET," & vbCrLf _
            & "J.STYLE_CODE, J.QTY * J.PACK * J.PPQ ORDER_QUANTITY, " & vbCrLf _
            & "J.NYAG_FIRST_COST_USD, J.STYLE, ICTSTYL1.SEASON_CODE" & vbCrLf _
            & "FROM JF_LINES J, ICTSTYL1, JF" & vbCrLf _
            & "WHERE ICTSTYL1.STYLE_CODE = J.STYLE_CODE AND JF.STYLE_CODE_PLM = ICTSTYL1.STYLE_CODE_PLM"
        ASCMAIN1.sql = "Select EDI_ARRIVAL_DATE, WORKSHEET, VEND_CODE, COUNT (*) LINES, SUM (ORDER_QUANTITY) QTY, SUM (ORDER_QUANTITY * NYAG_FIRST_COST_USD) AMT FROM (" & vbCrLf _
            & sqlMain _
            & ") GROUP BY EDI_ARRIVAL_DATE, WORKSHEET, VEND_CODE"

        Dim tbl As DataTable = ASCDATA1.GetDataTable


        Dim automatic As Boolean = True

        If automatic Then
            For Each row As DataRow In tbl.Select("", "EDI_ARRIVAL_DATE, WORKSHEET, VEND_CODE")
                Absx1.txtFor("VEND_CODE").Text = row.Item("VEND_CODE")
                Absx1.txtFor("PO_REFERENCE").Text = "."
                Dim WORKSHEET As String = row.Item("WORKSHEET")
                Dim EDI_ARRIVAL_DATE As Date = row.Item("EDI_ARRIVAL_DATE")
                Dim VEND_CODE As String = row.Item("VEND_CODE")

                Create_PO(sqlMain, WORKSHEET, EDI_ARRIVAL_DATE, VEND_CODE)

                blnAutomatic = True
                Click_Command("Update")
                blnAutomatic = False

                If ScreenMode Then
                    Stop
                End If
            Next

        Else
            Using frmmsg As New ASFMSGBF
                frmmsg.Show_grd(tbl, Me, "Select a JF Vendor/Division to Start a PO from the XLS file data")
                If frmmsg.grow Is Nothing Then
                    Exit Sub
                End If

                Absx1.txtFor("VEND_CODE").Text = frmmsg.grow.Cells("VEND_CODE").Value
                Absx1.txtFor("PO_REFERENCE").Text = "."
                Dim WORKSHEET As String = frmmsg.grow.Cells("WORKSHEET").Value
                Dim EDI_ARRIVAL_DATE As Date = frmmsg.grow.Cells("EDI_ARRIVAL_DATE").Value
                Dim VEND_CODE As String = frmmsg.grow.Cells("VEND_CODE").Value

                Create_PO(sqlMain, WORKSHEET, EDI_ARRIVAL_DATE, VEND_CODE)
            End Using
        End If

    End Sub

    Sub Create_PO(sqlMain As String, WORKSHEET As String, EDI_ARRIVAL_DATE As Date, VEND_CODE As String)

        Click_Command("New")

        If ScreenMode Then
            Absx1.txtFor("WHSE_CODE").Text = "18"
            Absx1.dteFor("PO_DATE_ORDERED").Value = Now.Date
            Absx1.dteFor("PO_DATE_SHIP_BY").Value = EDI_ARRIVAL_DATE.AddDays(-9 * 7)
            'Absx1.dteFor("PO_DATE_ETA").Value = Now.Date
            Absx1.dteFor("PO_DATE_CANCEL").Value = EDI_ARRIVAL_DATE
            ' Absx1.txtFor("COST_CODE").Text = "?"
            Absx1.txtFor("PO_SPEC_ORDR_NO").Text = "LOBLAW"
            ' Absx1.txtFor("TERM_CODE").Text = "?"
            Absx1.txtFor("PO_SHIP_VIA").Text = "BOAT"
            Absx1.chkFor("PO_WEB_VISIBLE").Checked = True
            Dim PO_FOB_DESC As String = "?"
            Dim rowICTPORT1 As DataRow = LookUp("ICTPORT1", Absx1.txtFor("PORT_CODE_ORIG").Text)
            If rowICTPORT1 IsNot Nothing Then PO_FOB_DESC = rowICTPORT1.Item("PORT_NAME") & ""
            PO_FOB_DESC = Split(PO_FOB_DESC & " (", " (")(0)
            Absx1.txtFor("PO_FOB_DESC").Text = PO_FOB_DESC
            Absx1.txtFor("PO_APPR_QUEUE").Text = "jtawil"
            'Absx1.chkFor("PO_APPR_PENDING").Checked = True
            chkReadyForApproval.Checked = True

            Absx1.txtFor("PO_MESSAGE").Text = "Packing, Ticketing & Labeling instructions to follow from Ian Cooper:" & vbCrLf & "ian@nyagroup.com"

            '   Absx1.txtFor("PO_NOTES").Text = frmmsg.grow.Cells("WORKSHEET").Value
            Absx1.txtFor("PO_NOTES").Text = WORKSHEET

            'PORT_CODE_ORIG,
            'PORT_CODE_DEST,
            'PO_MESSAGE,
            'LABEL_RESP_CODE,

            ASCMAIN1.sql = "Select * from (" & vbCrLf _
               & sqlmain _
               & ") where WORKSHEET = '" & WORKSHEET & "' and VEND_CODE = '" & VEND_CODE & "' AND EDI_ARRIVAL_DATE = '" & Format(EDI_ARRIVAL_DATE, "dd-MMM-yyyy") & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_CODE")
                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                Dim COLOR_CODE As String = "AST"
                Dim PO_QTY_ORD As Int64 = Val(row.Item("ORDER_QUANTITY") & "")
                Dim PO_COST_VCOST As Decimal = Val(row.Item("NYAG_FIRST_COST_USD") & "")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                Dim STYLE_CODE_PLM As String = rowICTSTYL1.Item("STYLE_CODE_PLM") & ""
                Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                Dim DESIGN_STYLE_NO As String = rowICTPLIN2.Item("DESIGN_STYLE_NO")

                Dim UM As Integer = 1

                Dim GROW As UltraWinGrid.UltraGridRow = Add_grdPOTORDR2(STYLE_CODE, COLOR_CODE, UM, PO_QTY_ORD, PO_COST_VCOST)
                GROW.Cells("STYLE_NOTES").Value = DESIGN_STYLE_NO
                GROW.Update()
            Next

        Else
            Stop
        End If

    End Sub

    Private Sub grdPOTORDRA_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRA.InitializeRow
        If e.Row.Cells("STATUS").Value & "" = "W" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Blue
        ElseIf e.Row.Cells("STATUS").Value & "" = "R" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("STATUS").Value & "" = "I" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Purple
        ElseIf e.Row.Cells("STATUS").Value & "" = "A" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Green
        Else
            e.Row.CellAppearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Function Load_Ship_Confirmation(FILENAME As String) As String
        Dim EMsg As String = ""
        Dim R As Integer = 0

        ASCMAIN1.Progress("Now Loading Ship Confirmation XLS")

        grdPOTORDS1.Text = "PO Shipment Confirmation Data Uploaded from " & FILENAME

        Try

            Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)

            Dim range As SpreadsheetGear.IRange = Nothing
            Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
            Dim rangePasteTo As SpreadsheetGear.IRange = Nothing

            EnforceConstraints(False)
            dst.Tables("POTORDS1").Rows.Clear()
            dst.Tables("POTORDS2").Rows.Clear()
            dst.Tables("POTORDS3").Rows.Clear()
            dst.Tables("POTORDS4").Rows.Clear()
            EnforceConstraints(True)

            Dim SHIPPED As Date = Nothing
            Dim VESSEL_NAME As String = ""
            Dim ETD As Date = Nothing
            Dim ETA As Date = Nothing
            Dim PORT_CODE_ORIG As String = ""
            Dim PORT_CODE_DEST As String = ""
            Dim T As String
            Dim I As Integer
            Dim RECORD_NO_ctr As Integer = 0
            Dim LNO As Integer = 0
            Dim LNO2 As Integer = 0
            Dim RECORD_NO As String = ""
            Dim MERCHANDISER_PREV As String = ""

            Dim rowPOTORDS1 As DataRow = Nothing
            Dim rowPOTORDS2 As DataRow = Nothing
            Dim rowPOTORDS3 As DataRow = Nothing
            Dim rowPOTORDS4 As DataRow = Nothing

            For R = 0 To worksheet.UsedRange.Rows.RowCount - 1
                'If R = 102 Then Stop
                If worksheet.Cells(R, 0).Text.Trim.StartsWith("The following") Then
                    T = worksheet.Cells(R, 0).Text.Trim
                    I = InStr(T, " on ")
                    If I <> 0 Then
                        T = Mid(T, I + 4).Trim
                        SHIPPED = CDate(T)
                    End If

                ElseIf worksheet.Cells(R, 1).Text.Trim.StartsWith("MERCHANDISER") Then
                    ' DO NOTHING - THIS IS A HEADING LINE
                ElseIf worksheet.Cells(R, 1).Text.Trim.StartsWith("VESSEL:") Then
                    VESSEL_NAME = ""
                    ETD = Nothing
                    ETA = Nothing
                    PORT_CODE_ORIG = ""
                    PORT_CODE_DEST = ""

                    T = worksheet.Cells(R, 1).Text.Trim.Substring(8).Trim
                    I = InStr(T, "(")
                    If I <> 0 Then
                        VESSEL_NAME = Mid(T, 1, I - 1)
                        T = Trim(Mid(T, I + 1))
                        I = InStr(T, ")")
                        If I <> 0 Then
                            Dim DTS As String = Mid(T, 1, I - 1)
                            T = Trim(Mid(T, I + 1))
                            I = InStr(DTS, "-")
                            If I <> 0 Then
                                Dim ETDX As String = Mid(DTS, 1, I - 1).Trim
                                If ETDX.StartsWith("ETD:") Then
                                    ETDX = Mid(ETDX, 5).Trim
                                End If
                                ETD = CDate(ETDX)
                                Dim ETAX As String = Mid(DTS, I + 1).Trim
                                If ETAX.StartsWith("ETA:") Then
                                    ETAX = Mid(ETAX, 5).Trim
                                End If
                                ETA = CDate(ETAX)
                            End If
                            I = InStr(T, " TO ")
                            If I <> 0 Then
                                PORT_CODE_ORIG = Trim(Mid(T, 1, I - 1))
                                PORT_CODE_DEST = Trim(Mid(T, I + 4))
                            End If
                        End If
                    End If

                    RECORD_NO_ctr += 1
                    RECORD_NO = Format(RECORD_NO_ctr, "000000")
                    LNO = 0

                    rowPOTORDS1 = dst.Tables("POTORDS1").NewRow
                    With rowPOTORDS1
                        .Item("RECORD_NO") = RECORD_NO
                        .Item("SHIPPED") = SHIPPED
                        .Item("VESSEL_NAME") = VESSEL_NAME
                        .Item("ETD") = ETD
                        .Item("ETA") = ETA
                        .Item("PORT_CODE_ORIG") = PORT_CODE_ORIG
                        .Item("PORT_CODE_DEST") = PORT_CODE_DEST
                    End With
                    dst.Tables("POTORDS1").Rows.Add(rowPOTORDS1)
                    'And worksheet.Cells(R, 3).Text <> ""
                ElseIf worksheet.Cells(R, 0).Text = "" _
                    And worksheet.Cells(R, 1).Text <> "" And worksheet.Cells(R, 2).Text <> "" _
                    And ((worksheet.Cells(R, 9).Text <> "" And worksheet.Cells(R, 10).Text <> "") Or (worksheet.Cells(R + 1, 4).Text = "DZ")) Then

                    'If worksheet.Cells(R, 9).Text <> "" And worksheet.Cells(R, 10).Text <> "" Then
                    'Else
                    '    If ASCMAIN1.Running_in_VS Then Stop
                    'End If

                    LNO2 = 0

                    rowPOTORDS2 = dst.Tables("POTORDS2").NewRow
                    With rowPOTORDS2
                        .Item("RECORD_NO") = RECORD_NO
                        LNO += 1
                        .Item("LNO") = LNO
                        .Item("FACTORY_NAME") = worksheet.Cells(R, 1).Text
                        .Item("CONTAINER_SIZE") = worksheet.Cells(R, 2).Text
                        .Item("SHIPPER") = worksheet.Cells(R, 3).Text
                        .Item("INVOICE_NO") = worksheet.Cells(R, 9).Text
                        Dim CONTAINER As String = worksheet.Cells(R, 10).Text
                        If CONTAINER.StartsWith("CTR:") Then
                            CONTAINER = CONTAINER.Substring(4).Trim
                        End If
                        .Item("CONTAINER") = CONTAINER
                    End With
                    dst.Tables("POTORDS2").Rows.Add(rowPOTORDS2)

                ElseIf worksheet.Cells(R, 0).Text <> "" _
                    And worksheet.Cells(R, 1).Text <> "" And worksheet.Cells(R, 2).Text <> "" And worksheet.Cells(R, 3).Text <> "" _
                    And worksheet.Cells(R, 4).Text <> "" And worksheet.Cells(R, 5).Text <> "" Then

                    Dim PO_REFERENCE As String
                    Dim STYLE As String

                    rowPOTORDS3 = dst.Tables("POTORDS3").NewRow
                    With rowPOTORDS3
                        .Item("RECORD_NO") = RECORD_NO
                        .Item("LNO") = LNO
                        LNO2 += 1
                        .Item("LNO2") = LNO2
                        Dim MERCHANDISER As String = worksheet.Cells(R, 0).Text
                        If MERCHANDISER = Chr(34) Then
                            MERCHANDISER = MERCHANDISER_PREV
                        End If
                        MERCHANDISER_PREV = MERCHANDISER
                        .Item("MERCHANDISER") = MERCHANDISER
                        PO_REFERENCE = worksheet.Cells(R, 1).Text
                        .Item("PO_REFERENCE") = PO_REFERENCE
                        STYLE = worksheet.Cells(R, 2).Text
                        .Item("STYLE") = STYLE
                        .Item("QTY") = Val(worksheet.Cells(R, 3).Value & "")
                        .Item("UM") = worksheet.Cells(R, 4).Text
                        .Item("CTN") = Val(worksheet.Cells(R, 5).Value & "")
                    End With
                    dst.Tables("POTORDS3").Rows.Add(rowPOTORDS3)

                    ASCMAIN1.sql = "Select POTORDR2.* from POTORDR1,POTORDR2" & vbCrLf _
                        & " where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
                        & "   and POTORDR2.STYLE_CODE like '%' || :PARM1 || '%'" & vbCrLf _
                        & "   and POTORDR1.PO_REFERENCE = :PARM2"
                    For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", New String() {STYLE, PO_REFERENCE}).Select()
                        Dim SUB_UNIT_PACK_QTY As Integer = Val(row.Item("SUB_UNIT_PACK_QTY") & "")
                        If SUB_UNIT_PACK_QTY = 0 Then SUB_UNIT_PACK_QTY = 1
                        rowPOTORDS4 = dst.Tables("POTORDS4").NewRow
                        With rowPOTORDS4
                            .Item("RECORD_NO") = RECORD_NO
                            .Item("LNO") = LNO
                            .Item("LNO2") = LNO2
                            .Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                            .Item("PO_ORDER_LNO") = row.Item("PO_ORDER_LNO")
                            .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                            .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                            .Item("PO_QTY_OPN") = Val(row.Item("PO_QTY_OPN") & "") * SUB_UNIT_PACK_QTY
                            .Item("SUB_UNIT_PACK_QTY") = SUB_UNIT_PACK_QTY
                            .Item("PO_QTY_OPN_DZ") = Val(row.Item("PO_QTY_OPN") & "") * SUB_UNIT_PACK_QTY / 12
                            .Item("PO_DATE_SHIP_BY") = row.Item("PO_DATE_SHIP_BY")
                            .Item("PO_DATE_ETA") = row.Item("PO_DATE_ETA")
                            .Item("LAST_OPER_SHIP_BY") = row.Item("LAST_OPER_SHIP_BY")
                            .Item("LAST_DATE_SHIP_BY") = row.Item("LAST_DATE_SHIP_BY")
                            .Item("PO_CONF_NO") = row.Item("PO_CONF_NO")
                            .Item("PO_CONF_DATE") = row.Item("PO_CONF_DATE")
                            .Item("PO_LINE_NOTE_INT") = row.Item("PO_LINE_NOTE_INT")
                        End With
                        dst.Tables("POTORDS4").Rows.Add(rowPOTORDS4)
                    Next

                    If worksheet.Cells(R, 8).Text <> "" Then
                        rowPOTORDS2.Item("CBM") = Val(worksheet.Cells(R, 8).Value & "")
                    ElseIf worksheet.Cells(R + 1, 8).Text <> "" Then
                        rowPOTORDS2.Item("CBM") = Val(worksheet.Cells(R + 1, 8).Value & "")
                    End If
                End If
            Next

        Catch ex As Exception
            EMsg = "Row " & CStr(R + 1) & ":" & ex.Message

        End Try

        ASCMAIN1.Progress("")

        Return EMsg

    End Function

    Private Sub grdPOTORDS1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTORDS1.AfterRowActivate
        Dim sql As String = ""
        If grdPOTORDS1.ActiveRow.Band.Index = 0 Then
            sql = "RECORD_NO = '" & grdPOTORDS1.ActiveRow.Cells("RECORD_NO").Value & "'"
        ElseIf grdPOTORDS1.ActiveRow.Band.Index = 1 Then
            sql = "RECORD_NO = '" & grdPOTORDS1.ActiveRow.Cells("RECORD_NO").Value & "' and LNO = " & grdPOTORDS1.ActiveRow.Cells("LNO").Value
        ElseIf grdPOTORDS1.ActiveRow.Band.Index = 2 Then
            sql = "RECORD_NO = '" & grdPOTORDS1.ActiveRow.Cells("RECORD_NO").Value & "' and LNO = " & grdPOTORDS1.ActiveRow.Cells("LNO").Value & "  and LNO2 = " & grdPOTORDS1.ActiveRow.Cells("LNO2").Value
        Else
            sql = ""
        End If

        If sql = "" Then
            grdPOTORDS4.Visible = False
        Else
            DirectCast(grdPOTORDS4.DataSource, DataTable).DefaultView.RowFilter = sql
            Sort_grdColumns(grdPOTORDS4, "RECORD_NO,LNO,LNO2,PO_ORDER_NO,PO_ORDER_LNO")
            grdPOTORDS4.Visible = True
        End If
    End Sub

    Private Sub grdPOTORDS1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDS1.InitializeRow
        If e.Row.Band.Index = 2 Then
            Dim UNITS As Int32 = Val(e.Row.Cells("UNITS").Value & "")
            Dim UNITS_PO As Int32 = Val(e.Row.Cells("UNITS_PO").Value & "")
            If UNITS <> UNITS_PO Then
                e.Row.Cells("UNITS").Appearance.BackColor = Drawing.Color.Yellow
            Else
                e.Row.Cells("UNITS").Appearance.BackColor = Drawing.Color.Empty
            End If
        End If
        If e.Row.Band.Index = 0 Then
            e.Row.Cells("SHIPPED").Appearance.ForeColor = Drawing.Color.Red
        End If
        If e.Row.Band.Index = 3 Then
            Dim PO_DATE_SHIP_BY As String = Format(e.Row.Cells("PO_DATE_SHIP_BY").Value, "yyyyMMdd")
            Dim SHIPPED As String = Format(e.Row.ParentRow.ParentRow.ParentRow.Cells("SHIPPED").Value, "yyyyMMdd")

            If PO_DATE_SHIP_BY > SHIPPED Then
                e.Row.Cells("PO_DATE_SHIP_BY").Appearance.BackColor = Drawing.Color.LightGreen
            ElseIf PO_DATE_SHIP_BY < SHIPPED Then
                e.Row.Cells("PO_DATE_SHIP_BY").Appearance.BackColor = Drawing.Color.LightPink
            Else
                e.Row.Cells("PO_DATE_SHIP_BY").Appearance.BackColor = Drawing.Color.Empty
            End If

            If e.Row.Cells("SEL").Value & "" = "1" Then
                e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Violet
            Else
                e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Empty
            End If

        End If
    End Sub

    Sub Update_Ship_Confirmation()

        dst.Tables("POTORDR1").Rows.Clear()
        dst.Tables("POTORDR2").Rows.Clear()

        ASCMAIN1.Progress("Now Updating Ship Confirmation")
        BeginTrans()

        EnforceConstraints(False)
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDS4").Select("SEL = '1'"), "PO_ORDER_NO").Select("")
            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
            Fill_Record("POTORDR1", PO_ORDER_NO, False, False)

            Fill_Records("POTORDR2", PO_ORDER_NO, False)
        Next

        ' DOING THIS BECAUSE WE HAVE TO HONOR THE CONSTRAINT
        dst.Tables("POTORDRR").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR2"), New String() {"PO_ORDER_NO", "STYLE_CODE", "COLOR_CODE"}).Rows
            Dim rowPOTORDRR As DataRow = dst.Tables("POTORDRR").NewRow
            rowPOTORDRR.Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
            rowPOTORDRR.Item("STYLE_CODE") = row.Item("STYLE_CODE")
            rowPOTORDRR.Item("COLOR_CODE") = row.Item("COLOR_CODE")
            dst.Tables("POTORDRR").Rows.Add(rowPOTORDRR)
        Next

        EnforceConstraints(True)

        Dim PO_ORDER_NOs As New List(Of String)

        For Each row As DataRow In dst.Tables("POTORDS4").Select("SEL = '1'")
            Dim RECORD_NO As String = row.Item("RECORD_NO")
            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(row.Item("PO_ORDER_LNO") & "")
            Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(PO_ORDER_NO)
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})

            If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then
                PO_ORDER_NOs.Add(PO_ORDER_NO)
            End If

            Dim rowPOTORDS1 As DataRow = dst.Tables("POTORDS1").Rows.Find(RECORD_NO)
            Dim SHIPPED As DateTime = rowPOTORDS1.Item("SHIPPED")

            With rowPOTORDR2
                If .Item("PO_ORIG_DATE_SHIP_BY") & "" = "" Then
                    .Item("PO_ORIG_DATE_SHIP_BY") = .Item("PO_DATE_SHIP_BY")
                    .Item("PO_ORIG_DATE_ETA") = .Item("PO_DATE_ETA")
                End If

                Dim PO_DATE_SHIP_BY As DateTime = .Item("PO_DATE_SHIP_BY")
                Dim PO_DATE_ETA As DateTime = .Item("PO_DATE_ETA")
                Dim DYS As Integer = PO_DATE_ETA.Subtract(PO_DATE_SHIP_BY).Days

                .Item("PO_DATE_SHIP_BY") = SHIPPED
                .Item("PO_DATE_ETA") = SHIPPED.AddDays(DYS)

                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP

                .Item("LAST_OPER_SHIP_BY") = ASCMAIN1.USER_ID
                .Item("LAST_DATE_SHIP_BY") = DATETIME_STAMP

                .Item("PO_CONF_NO") = "XLS"
                .Item("PO_CONF_DATE") = DATETIME_STAMP
            End With
        Next

        For Each PO_ORDER_NO In PO_ORDER_NOs
            rowPOTORDR1 = dst.Tables("POTORDR1").Rows.Find(PO_ORDER_NO)
            If Check_Changed_Fields(False) Then ' this function records audit trail, so it is important to call even if we do nothing with the results returned

            End If
            Dim sqlx As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            INIT_LAST("POTORDR1", False, sqlx, True)

            Record_Event("UPDETD", "PO Conf Dates Updated", True)
        Next

        Update_Record_TDA("POTORDR1")
        Update_Record_TDA("POTORDR2")
        Update_Record_TDA("POTORDXR")

        CommitTrans("Update Complete")
        ASCMAIN1.Progress("")

    End Sub

    Sub Toggle_SC()

        With UltraExplorerBar1
            .Groups("Screen Control").Visible = Not ScreenMode
            .Groups("Ship Confirmation").Visible = ScreenMode
            .Groups("Status Filter").Visible = Not ScreenMode
        End With

        grdPOTORDS1.Visible = ScreenMode
        tabPO.Visible = Not ScreenMode
        spl.Panel1Collapsed = ScreenMode

    End Sub

    Private Sub grdPOTORDS1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTORDS1.InitializeLayout

    End Sub

    Private Sub grdPOTORDS4_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTORDS4.InitializeLayout

    End Sub

    Private Sub grdPOTORDS4_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDS4.InitializeRow

        If Not e.Row.IsDataRow Then Exit Sub

        Dim PO_DATE_SHIP_BY As String = Format(e.Row.Cells("PO_DATE_SHIP_BY").Value, "yyyyMMdd")

        Try
            Dim RECORD_NO As String = e.Row.Cells("RECORD_NO").Value
            Dim LNO As Integer = Val(e.Row.Cells("LNO").Value & "")
            Dim LNO2 As Integer = Val(e.Row.Cells("LNO2").Value & "")

            Dim PO_ORDER_NO As String = e.Row.Cells("PO_ORDER_NO").Value
            Dim PO_ORDER_LNO As Integer = Val(e.Row.Cells("PO_ORDER_LNO").Value & "")
            Dim rowPOTORDS4 As DataRow = dst.Tables("POTORDS4").Rows.Find(New Object() {RECORD_NO, LNO, LNO2, PO_ORDER_NO, PO_ORDER_LNO})
            Dim rowPOTORDS1 As DataRow = dst.Tables("POTORDS1").Rows.Find(RECORD_NO) ' rowPOTORDS4.GetParentRow("POTORDS3_POTORDS4").GetParentRow("POTORDS2_POTORDS3").GetParentRow("POTORDS1_POTORDS2")

            Dim SHIPPED As String = Format(rowPOTORDS1.Item("SHIPPED"), "yyyyMMdd")

            If PO_DATE_SHIP_BY > SHIPPED Then
                e.Row.Cells("PO_DATE_SHIP_BY").Appearance.BackColor = Drawing.Color.LightGreen
            ElseIf PO_DATE_SHIP_BY < SHIPPED Then
                e.Row.Cells("PO_DATE_SHIP_BY").Appearance.BackColor = Drawing.Color.LightPink
            Else
                e.Row.Cells("PO_DATE_SHIP_BY").Appearance.BackColor = Drawing.Color.Empty
            End If

            If e.Row.Cells("SEL").Value & "" = "1" Then
                e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Violet
            Else
                e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Empty
            End If
        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then Stop

        End Try
    End Sub

    Private Sub txtCUST_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtCUST_CODE.ValueChanged
        Dim blnShowYintak As Boolean = ScreenMode And ASCMAIN1.CLIENT = "VAN" And (Absx1.txtFor("VEND_CODE").Text = "YINTAK" Or Absx1.txtFor("VEND_CODE").Text = "CIVIC")
        Set_Visible_CARTON_COUNT(blnShowYintak)
    End Sub

    Sub Set_Visible_CARTON_COUNT(blnShowYintak As Boolean)
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim rowPOTPACKC As DataRow = LookUp("POTPACKC", CUST_CODE)
        'lblCARTON_COUNT.Visible = blnShowYintak And txtCUST_CODE.Text = "WALMART"
        'txtCARTON_COUNT.Visible = blnShowYintak And txtCUST_CODE.Text = "WALMART"
        lblCARTON_COUNT.Visible = ScreenMode AndAlso blnShowYintak AndAlso rowPOTPACKC IsNot Nothing AndAlso rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "0"
        txtCARTON_COUNT.Visible = ScreenMode AndAlso blnShowYintak AndAlso rowPOTPACKC IsNot Nothing AndAlso rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "0"
    End Sub
End Class