Imports Infragistics.Win.UltraWinGrid

' ALTER TABLE POTVBKG1 DROP COLUMN PO_SHIPMENT_NO;
' ALTER TABLE POTVBKG1 DROP COLUMN PO_SHIPMENT_LNO;

Public Class POFSHIP1
    ' warehouse receipt currently piggy backs on code flow around LP_STATUS, and not all of the code is proper - ie nadine should not be able to recall a shipment that has been "sent" to the warehouse

    Dim automated_cost_complete As Boolean = False

    Dim PO_SHIPMENT_NO As String
    Dim rowPOTSHIP1 As DataRow
    Dim PPK_CODE_ctr As Int32

    Dim WHSE_CODE As String
    Dim WHSE_TYPE As String
    Dim LP_CODE As String
    Dim WHSE_LOCATOR As String
    Dim WHSE_CTN_CTL As String

    Dim WHTSTYLX As String
    Dim sqlPOTSHIPX As String
    Dim select_from_3PL_list As Boolean = False

    Dim Select_from_Whse_Receipt As Boolean = False

    Dim ASW As New Dictionary(Of String, String)
    Dim Checking_for_PPK As Boolean = False

    Dim sql_POTSHIPI As String
    Dim POTSHIPI As String
    Dim POTSHIPX As String

    Dim CARTON_DIMS_by_PO As New Dictionary(Of String, String)

    Dim EDI_DOC_SEQ_NO As String = ""

    Dim ship_entry As Boolean
    Dim receipt_mode As Boolean
    Dim cost_calc As Boolean
    Dim cost_ind As Boolean
    Dim XYP As String
    Dim STYLE_CODEs_No_Duty As New List(Of String)
    Dim POTSHIPP As String
    Dim sqlPOTSHIPP As String = ""

    Dim dicPOTORDR1 As New Dictionary(Of String, DataRow)
    Dim dicPOTORDR2 As New Dictionary(Of String, DataTable)
    Dim packingListPOs As New List(Of String)
    Dim dicStyleDesc As New Dictionary(Of String, String)
    Dim dicColorDesc As New Dictionary(Of String, String)
    Dim POTORDR1_added As New List(Of String)

    Dim dicWorksheetPOs As New Dictionary(Of String, String)
    Dim packingFromXLS As Boolean = False
    Dim packingFromBooking As Boolean = False
    Dim dicWorkbooks As New List(Of String)

    'Dim fix_ICTSTYL1_packs As Boolean = False

    Dim AT_Packing As Boolean = False
    Dim AT_Packing_Errors As String = ""
    Dim rowATSHIPS As DataRow
    Dim QTY_PACKED As New Dictionary(Of String, Int64)
    Dim WORKBOOK_COUNTER As Integer = 0
    Dim loading_AT As Boolean = False
    Dim WH_REC_NOsInProcess As New List(Of String)
    Dim sqlPOTVBKGX As String
    Dim eMsg_Booking As String
    Dim POTVBKG2_RECORDS As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")
        Get_PARM("WHTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        InquiryMode = (MENU_ITEM_OBJECT = "POFSHIPI")
        receipt_mode = (MENU_ITEM_OBJECT = "POFSHIPR")
        cost_calc = (MENU_ITEM_OBJECT = "POFSHIPC")
        ship_entry = (MENU_ITEM_OBJECT = "POFSHIP1")

        If ship_entry And (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
            WHTSTYLX = TAC.WHCMAIN1.Prepare_WHTSTYLX("", "", True)
        End If

        sql_POTSHIPI = "" _
            & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ", POTSHIP3.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.WHSE_CODE" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
            & ", POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & ", POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_REC" & vbCrLf _
            & "     from POTSHIP3, POTSHIP2, POTORDR2, POTSHIP1" & vbCrLf _
            & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIP_STATUS = 'C'" & vbCrLf _
            & "   and POTSHIP1.COST_COMPLETE = '0'" & vbCrLf
        POTSHIPI = ASCMAIN1.Temp_Table(sql_POTSHIPI & " and ROWNUM < 1")

        With dst

            If ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "RGI" Then
                Create_TDA(.Tables.Add, "ICTSTYL1", "*")
            End If

            sqlPOTSHIPX = "SELECT POTSHIP1.*" & vbCrLf _
                & ", ICTWHSE1.LP_CODE, WHTLPXN1.INIT_OPER LP_XNO_INIT_OPER, WHTLPXN1.INIT_DATE LP_XNO_INIT_DATE" & vbCrLf _
                & " from POTSHIP1,WHTLPXN1,ICTWHSE1" & vbCrLf _
                & " where WHTLPXN1.LP_XNO (+) = POTSHIP1.LP_XNO" & vbCrLf _
                & "   and ICTWHSE1.WHSE_CODE (+) = POTSHIP1.WHSE_CODE"
            ASCMAIN1.sql = sqlPOTSHIPX & " and ROWNUM <1"
            POTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & POTSHIPX & " Add Primary Key (PO_SHIPMENT_NO)")

            ASCMAIN1.sql = "Select POTSHIPX.*" & vbCrLf _
                & ", X.CONTAINER_NO, X.BOL_NO, X.COMM_INV_NO, X.PO_SHIP_CTNS, X.PO_DATE_RECEIVED_MIN, X.PO_DATE_RECEIVED_MAX, X.ORDR_NO" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from " & POTSHIPX & " POTSHIPX, SOTORDR1" & vbCrLf _
                & ", (Select PO_SHIPMENT_NO" & vbCrLf _
                & ", Min (CONTAINER_NO) CONTAINER_NO" & vbCrLf _
                & ", Min (BOL_NO) BOL_NO" & vbCrLf _
                & ", Min (COMM_INV_NO) COMM_INV_NO" & vbCrLf _
                & ", Min (ORDR_NO) ORDR_NO" & vbCrLf _
                & ", Sum (PO_SHIP_CTNS) PO_SHIP_CTNS" & vbCrLf _
                & ", Min (PO_DATE_RECEIVED) PO_DATE_RECEIVED_MIN" & vbCrLf _
                & ", Max (PO_DATE_RECEIVED) PO_DATE_RECEIVED_MAX" & vbCrLf _
                & " from POTSHIP2 where PO_SHIPMENT_NO " & vbCrLf _
                & " in (Select PO_SHIPMENT_NO from " & POTSHIPX & ") group by PO_SHIPMENT_NO) X" & vbCrLf _
                & " where X.PO_SHIPMENT_NO (+) = POTSHIPX.PO_SHIPMENT_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO (+) = X.ORDR_NO"
            Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "", 1)
            .Tables("POTSHIPX").Columns.Add("LINES", GetType(System.Int64))
            .Tables("POTSHIPX").Columns.Add("LINES_REC", GetType(System.Int64))
            .Tables("POTSHIPX").Columns.Add("WH_LINES_REC", GetType(System.Int64))
            .Tables("POTSHIPX").Columns("PO_SHIP_CTNS").DataType = GetType(System.Int64)

            ASCMAIN1.sql = "Select * from " & POTSHIPI
            Create_TDA(.Tables.Add, "POTSHIPI", "**", 0, False, "", 0)
            With .Tables("POTSHIPI")
                .Columns.Add("ORDR_QTY_SHIP", GetType(System.Int64))
            End With

            Create_TDA(.Tables.Add, "POTSHIP1", "*")
            With .Tables("POTSHIP1")
                .Columns.Add("PO_DATE_RECEIVED", GetType(System.DateTime))
                .Columns.Add("PO_SOURCE_DOC")
                .Columns.Add("CUSTOMS_DUTY_AMT_DIST", GetType(System.Decimal))
                .Columns.Add("CUSTOMS_DUTY_AMT_NOT_DIST", GetType(System.Decimal), "ISNULL(CUSTOMS_DUTY_AMT,0) - ISNULL(CUSTOMS_DUTY_AMT_DIST,0)")
            End With

            Create_TDA(.Tables.Add, "POTSHIP2", "*", 1)
            With .Tables("POTSHIP2")
                .Columns.Add("CLOSE")
                .Columns.Add("CONTAINER_SEAL_NO")
                .Columns.Add("WORKBOOK_COUNTER", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select PO_ORDER_NO, PO_ORDER_LNO, PO_QTY_OPN from POTORDR2" _
                & " where (PO_ORDER_NO, PO_ORDER_LNO) in " _
                & " (Select Distinct PO_ORDER_NO, PO_ORDER_LNO from POTSHIP3 where PO_SHIPMENT_NO = :PARM1)"
            Create_TDA(.Tables.Add, "POTORDRO", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select POTSHIP3.* " & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.PO_COMM_PAYABLE_TO_BRKR, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE " & vbCrLf _
                & ", POTORDR2.PO_QTY_OPN, POTORDR2.PO_QTY_UOM, POTORDR2.PO_COST ORDR2_COST" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_BODY_CODE, POTORDR2.SUB_UNIT_PACK_QTY, POTORDR2.CARTON_PACK_QTY" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
                & ", POTSHIP3.PO_QTY_REC PO_QTY_REC_OLD, ICTSTYL1.CASE_CUBE" & vbCrLf _
                & " from POTSHIP3,POTORDR2,ICTSTYL1,POTORDR1 " & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTSHIP3", "**", 0, True, "V", 4)
            '.Tables("POTSHIP3").Columns("PO_QTY_SHP").DataType = GetType(System.Decimal)
            '.Tables("POTSHIP3").Columns("PO_QTY_REC").DataType = GetType(System.Decimal)

            Create_Relation("POTORDRO", "POTSHIP3", "PO_ORDER_NO,PO_ORDER_LNO")
            .Tables("POTORDRO").Columns.Add("PO_QTY_SHP", GetType(System.Int32), "SUM(CHILD(POTORDRO_POTSHIP3).PO_QTY_SHP)")
            .Tables("POTORDRO").Columns.Add("PO_QTY_OPN_PRE", GetType(System.Int32))

            Create_Relation("POTSHIP2", "POTSHIP3", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")
            With .Tables("POTSHIP3").Columns
                .Add("PO_QTY_VAR", GetType(System.Int32), "IIF(PARENT(POTSHIP2_POTSHIP3).PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_REC,0) - ISNULL(PO_QTY_SHP,0) = 0,NULL,ISNULL(PO_QTY_REC,0) - ISNULL(PO_QTY_SHP,0)))")
                .Add("PO_QTY_OPN_PRE", GetType(System.Int32), "PARENT(POTORDRO_POTSHIP3).PO_QTY_OPN_PRE")
                .Add("PO_QTY_SHP_DZ", GetType(System.Decimal), "PO_QTY_SHP / (12 / SUB_UNIT_PACK_QTY)")
                .Add("PO_QTY_REC_DZ", GetType(System.Decimal), "PO_QTY_REC / (12 / SUB_UNIT_PACK_QTY)")
                .Add("PO_SHIP_STATUS", GetType(System.String), "PARENT(POTSHIP2_POTSHIP3).PO_SHIP_STATUS")
                If cost_calc Or receipt_mode Then
                    .Add("PO_QTY_SR", GetType(System.Decimal), "ISNULL(PO_QTY_SHP,0)")
                Else
                    .Add("PO_QTY_SR", GetType(System.Decimal), "IIF(PO_SHIP_STATUS='C',ISNULL(PO_QTY_REC,0),ISNULL(PO_QTY_SHP,0))")
                End If
                .Add("PO_QTY_SR_DZ", GetType(System.Decimal), "IIF(PO_SHIP_STATUS='C',ISNULL(PO_QTY_REC_DZ,0),ISNULL(PO_QTY_SHP_DZ,0))")
                .Add("TOTAL_DUTY", GetType(System.Decimal), "PO_QTY_SR * ISNULL(PO_COST_DUTY,0)")
                .Add("CONTAINER_NO", GetType(System.String), "PARENT(POTSHIP2_POTSHIP3).CONTAINER_NO")

                If receipt_mode Then
                    .Add("NET_OPEN", GetType(System.Decimal), "PO_QTY_OPN")
                Else
                    .Add("NET_OPEN", GetType(System.Decimal), "PO_QTY_OPN_PRE")
                End If
                .Add("NET_OPEN_DZ", GetType(System.Decimal), "NET_OPEN / (12 / SUB_UNIT_PACK_QTY)")
                '.Add("PO_AMT_REC", GetType(System.Decimal), "ISNULL(PO_COST_VCOST,0) * ISNULL(PO_QTY_REC,0)")

                'If ASCMAIN1.CLIENT = "VAN" Then
                ' MAKING THIS CHANGE ON 09/06/19 
                ' BECAUSE I NOTICED THAT ICTIREC1 (WHICH GETS AMT_REC FROM THE SUM OF THIS FIELD) 
                ' DOES NOT AGREE WITH THE COST VALUE FROM SUM OF THE DETAILS (WHICH EXTENDS QTY REC BY PO_COST)
                ' PROBABLY SHOULD DO THIS AT RGI AND NYA AS WELL
                ' WHEN AP INVOICING A SHIPMENT, PO_COST IS USED FROM POTSHIP3
                ' BUT WHEN INVOICING A RECEIPT, ICTIREC1.AMT_REC IS DISPLAYED IN THE LOWER GRID, 
                '  AND THEN IT GETS CHANGED TO THE PO_COST EXTENDED VALUE WHEN THE RECEIPT IS SELECTED. CAUSING CONFUSION (MARIA TALAN)
                .Add("PO_AMT_REC", GetType(System.Decimal), "ISNULL(PO_COST,0) * ISNULL(PO_QTY_REC,0)")
                ' AT VANDALE, PO_COST = PO_COST_VCOST + PO_COST_OTHER + PO_COST_QUOTA_DF
                ' End If

                ' ON 09/06/2019 WJZ CORRECTED VAN.ICTIREC1.AMT_REC TO THE SUM OF ICTIREC2.QTY_REC * ICTIREC2.PO_COST FOR YP >= 201301

                ' ALSO - IT APPEARS THAT WHEN A 2ND RECEIPT WAS ENTERED THE REC QTY DIFFERENTIAL GOT RECORDED IN ICTIREC2, RATHER THAN THE REVISED RECEIPT QTY
                ' 2018 - HAPPENED ONCE - 084262 166685                084264 154           
                ' 2019 - HAPPENED ONCE - 084805 0                     084797 1626                  


                .Add("FIRST_COST_TOTAL", GetType(System.Decimal), "ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0)")
                .Add("FIRST_COST_TOTAL_DZ", GetType(System.Decimal), "(PO_COST_VCOST_DZ + PO_COST_MATLS_DZ + PO_COST_OTHER_DZ)")
                'If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                '    .Add("COMMISSION_COST", GetType(System.Decimal), "(((ISNULL(PO_COST_COMM,0)+ISNULL(PO_COST_BUFFER,0)) / 100) * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_QUOTA,0)))")
                '    .Add("COMMISSION_COST_DZ", GetType(System.Decimal), "(((ISNULL(PO_COST_COMM,0)+ISNULL(PO_COST_BUFFER,0)) / 100) * (ISNULL(PO_COST_VCOST_DZ,0) + ISNULL(PO_COST_MATLS_DZ,0) + ISNULL(PO_COST_QUOTA_DZ,0)))")
                'Else
                .Add("COMMISSION_COST", GetType(System.Decimal), "(((ISNULL(PO_COST_COMM,0)+ISNULL(PO_COST_BUFFER,0)) / 100) * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0) + ISNULL(PO_COST_QUOTA,0)))")
                .Add("COMMISSION_COST_DZ", GetType(System.Decimal), "(((ISNULL(PO_COST_COMM,0)+ISNULL(PO_COST_BUFFER,0)) / 100) * (ISNULL(PO_COST_VCOST_DZ,0) + ISNULL(PO_COST_MATLS_DZ,0) + ISNULL(PO_COST_OTHER_DZ,0) + ISNULL(PO_COST_QUOTA_DZ,0)))")
                'End If

                .Add("EXT_WEIGHT_FACTOR", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(WEIGHT_FACTOR,0)")
                .Add("EXT_VCOST", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_VCOST,0)")
                .Add("EXT_MATLS", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_MATLS,0)")
                .Add("EXT_OTHER", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_OTHER,0)")
                .Add("EXT_FIRST", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(FIRST_COST_TOTAL,0)")
                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    .Add("EXT_COMM", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_COMM,0)/100 * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_QUOTA,0))")
                Else
                    .Add("EXT_COMM", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_COMM,0)/100 * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0) + ISNULL(PO_COST_QUOTA,0))")
                End If
                .Add("EXT_BUFFER", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_BUFFER,0)/100 * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0) + ISNULL(PO_COST_QUOTA,0))")
                .Add("EXT_QUOTA", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_QUOTA,0)")
                .Add("EXT_QUOTA_DF", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_QUOTA_DF,0)")
                .Add("EXT_FREIGHT", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_FREIGHT_IN,0)")
                .Add("EXT_CUSTOMS", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_CUSTOMS,0)")
                .Add("EXT_DUTY", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_DUTY,0)")
                .Add("EXT_TRUCKING", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_TRUCKING,0)")
                .Add("EXT_MISC", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_MISC,0)")
                .Add("EXT_LANDED", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_LANDED,0)")
                .Add("EXT_FIRST_CALC", GetType(System.Decimal), "EXT_VCOST + EXT_MATLS + EXT_OTHER")
                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    ' note - landed for NYA does NOT include Commission
                    .Add("EXT_LANDED_CALC", GetType(System.Decimal), "EXT_FIRST + EXT_BUFFER + EXT_QUOTA + EXT_QUOTA_DF + EXT_FREIGHT + EXT_CUSTOMS + EXT_DUTY + EXT_TRUCKING + EXT_MISC")
                Else
                    .Add("EXT_LANDED_CALC", GetType(System.Decimal), "EXT_FIRST + EXT_COMM + EXT_BUFFER + EXT_QUOTA + EXT_QUOTA_DF + EXT_FREIGHT + EXT_CUSTOMS + EXT_DUTY + EXT_TRUCKING + EXT_MISC")
                End If

                If Not cost_calc And Not receipt_mode Then
                    For Each COLUMN_NAME As String In New String() _
                        {"EXT_WEIGHT_FACTOR", "EXT_VCOST", "EXT_MATLS", "EXT_OTHER", "EXT_FIRST",
                         "EXT_COMM", "EXT_BUFFER", "EXT_QUOTA", "EXT_QUOTA_DF", "EXT_FREIGHT",
                         "EXT_CUSTOMS", "EXT_DUTY", "EXT_TRUCKING", "EXT_MISC", "EXT_LANDED"}
                        dst.Tables("POTSHIP3").Columns(COLUMN_NAME).Expression = ""
                    Next
                End If

                .Add("LINE_EXACT", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_SHP,0) = ISNULL(PO_QTY_REC,0),1,0))")
                .Add("LINE_OVER", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_SHP,0) < ISNULL(PO_QTY_REC,0),1,0))")
                .Add("LINE_SHORT", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_SHP,0) > ISNULL(PO_QTY_REC,0),1,0))")
                .Add("LINE_ZERO", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_REC,0) = 0,1,0))")

                .Add("PO_QTY_PCK", GetType(System.Decimal))

            End With

            .Tables("POTSHIP3").Columns("PO_QTY_PCK").DefaultValue = 0

            With .Tables("POTSHIP2")
                .Columns.Add("TOTAL_WEIGHT_FACTOR", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_WEIGHT_FACTOR)")
                .Columns.Add("TOTAL_VCOST", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_VCOST)")
                .Columns.Add("TOTAL_MATLS", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_MATLS)")
                .Columns.Add("TOTAL_OTHER", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_OTHER)")
                .Columns.Add("TOTAL_FIRST", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_FIRST)")
                .Columns.Add("TOTAL_COMM", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_COMM)")
                .Columns.Add("TOTAL_BUFFER", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_BUFFER)")
                .Columns.Add("TOTAL_QUOTA", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_QUOTA)")
                .Columns.Add("TOTAL_QUOTA_DF", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_QUOTA_DF)")
                .Columns.Add("TOTAL_FREIGHT", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_FREIGHT)")
                .Columns.Add("TOTAL_CUSTOMS", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_CUSTOMS)")
                .Columns.Add("TOTAL_DUTY", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_DUTY)")
                .Columns.Add("TOTAL_TRUCKING", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_TRUCKING)")
                .Columns.Add("TOTAL_MISC", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_MISC)")
                .Columns.Add("TOTAL_LANDED", GetType(System.Decimal), "SUM(CHILD(POTSHIP2_POTSHIP3).EXT_LANDED)")
                .Columns.Add("TOTAL_FIRST_CALC", GetType(System.Decimal), "TOTAL_VCOST + TOTAL_MATLS + TOTAL_OTHER")
                If ASCMAIN1.CLIENT = "NYA" Then
                    .Columns.Add("TOTAL_LANDED_CALC", GetType(System.Decimal), "TOTAL_FIRST + TOTAL_BUFFER + TOTAL_QUOTA + TOTAL_QUOTA_DF + TOTAL_FREIGHT + TOTAL_CUSTOMS + TOTAL_DUTY + TOTAL_TRUCKING + TOTAL_MISC")
                Else
                    .Columns.Add("TOTAL_LANDED_CALC", GetType(System.Decimal), "TOTAL_FIRST + TOTAL_COMM + TOTAL_BUFFER + TOTAL_QUOTA + TOTAL_QUOTA_DF + TOTAL_FREIGHT + TOTAL_CUSTOMS + TOTAL_DUTY + TOTAL_TRUCKING + TOTAL_MISC")
                End If

                If Not cost_calc And Not receipt_mode Then
                    For Each COLUMN_NAME As String In New String() _
                        {"TOTAL_WEIGHT_FACTOR", "TOTAL_VCOST", "TOTAL_MATLS", "TOTAL_OTHER", "TOTAL_FIRST",
                         "TOTAL_COMM", "TOTAL_BUFFER", "TOTAL_QUOTA", "TOTAL_QUOTA_DF", "TOTAL_FREIGHT",
                         "TOTAL_CUSTOMS", "TOTAL_DUTY", "TOTAL_TRUCKING", "TOTAL_MISC", "TOTAL_LANDED"}
                        .Columns(COLUMN_NAME).Expression = ""
                    Next
                End If

                .Columns.Add("PO_QTY_SHP", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).PO_QTY_SHP)")
                .Columns.Add("PO_QTY_REC", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).PO_QTY_REC)")
                .Columns.Add("PO_QTY_VAR", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).PO_QTY_VAR)")
                .Columns.Add("LINES", GetType(System.Int64), "COUNT(CHILD(POTSHIP2_POTSHIP3).PO_ORDER_LNO)")
                .Columns.Add("LINES_EXACT", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).LINE_EXACT)")
                .Columns.Add("LINES_OVER", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).LINE_OVER)")
                .Columns.Add("LINES_SHORT", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).LINE_SHORT)")
                .Columns.Add("LINES_ZERO", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).LINE_ZERO)")

            End With

            ASCMAIN1.sql = "Select * from ICTDUTY4"
            Create_TDA(.Tables.Add, "ICTDUTY4", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "POTSHIP4", "*", 1)
            Create_TDA(.Tables.Add, "POTSHIP7", "*", 1)
            Create_TDA(.Tables.Add, "POTSHIP8", "*", 1)
            .Tables("POTSHIP8").Columns.Add("UNITS", GetType(System.Int32), "QTY*IIF(ISNULL(DOZENS,'0')='1',12,1)")

            Create_Relation("POTSHIP2", "POTSHIP7", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")

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
            .Tables("POTSHIP7").Columns.Add("TOTAL_WEIGHT", GetType(System.Decimal), "ISNULL(CARTONS,0) * ISNULL(CARTON_WEIGHT,0)")
            .Tables("POTSHIP8").Columns.Add("CBM", GetType(System.Decimal), "IIF(ISNULL(PARENT(POTSHIP7_POTSHIP8).TOTAL_UNITS,0) = 0, 0, ISNULL(TOTAL_UNITS,0) * ISNULL(PARENT(POTSHIP7_POTSHIP8).CBM,0) / ISNULL(PARENT(POTSHIP7_POTSHIP8).TOTAL_UNITS,0))")


            With .Tables.Add("POTSHIPQ")
                .Columns.Add("CTN_NO", GetType(System.Int32))
                .Columns.Add("PACK", GetType(System.Int32))
                .Columns.Add("CTNS", GetType(System.Int32))
                .Columns.Add("NOTE")
                .PrimaryKey = New DataColumn() { .Columns("CTN_NO")}
            End With

            With .Tables.Add("POTSHIPR")
                .Columns.Add("PO_SHIPMENT_NO")
                .Columns.Add("PO_SHIPMENT_LNO", GetType(System.Int64))
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("QTY_SHP", GetType(System.Int32))
                .Columns.Add("QTY_CTN", GetType(System.Int32))
                .Columns.Add("QTY_VAR", GetType(System.Int32), "ISNULL(QTY_SHP,0) - ISNULL(QTY_CTN,0)")
                .Columns.Add("COLOR_DESC")
                .PrimaryKey = New DataColumn() { .Columns("PO_SHIPMENT_NO"), .Columns("PO_SHIPMENT_LNO"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With

            With .Tables.Add("POTSHPWB")
                .Columns.Add("WORKBOOK")
                .Columns.Add("WORKSHEET")
                .Columns.Add("PO_SHIPMENT_NO")
                .Columns.Add("PO_SHIPMENT_LNO", GetType(System.Int64))
                .Columns.Add("CONTAINER_NO")
                .Columns.Add("IS_PPK")
                .Columns.Add("NW", GetType(System.Decimal))
                .Columns.Add("MEAS")
                .Columns.Add("TOTAL_CTNS")
                .Columns.Add("IMPORTED")
                .PrimaryKey = New DataColumn() { .Columns("WORKBOOK"), .Columns("WORKSHEET"), .Columns("PO_SHIPMENT_NO"), .Columns("PO_SHIPMENT_LNO")}
            End With

            With .Tables.Add("POTSHPIE")
                .Columns.Add("WORKBOOK")
                .Columns.Add("WORKSHEET")
                .Columns.Add("IE_LNO", GetType(System.Int64))
                .Columns.Add("ERROR_MSG")
                .Columns.Add("XLS_REF")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("QTY", GetType(System.Int64))
                .Columns.Add("PO_ORDER_NO")
                .Columns.Add("PO_REFERENCE")
                .Columns.Add("CONTAINER_NO")
                .Columns.Add("COMM_INV_NO")
                .Columns.Add("BOL_NO")
                .Columns.Add("PO_SHIPMENT_LNO", GetType(System.Int32))
                .Columns.Add("QTY_PACK", GetType(System.Int64))
                .Columns.Add("QTY_OPEN_THIS_PO", GetType(System.Int64))
                .Columns.Add("QTY_NEEDED", GetType(System.Int64), "ISNULL(QTY_PACK,0) - ISNULL(QTY_OPEN_THIS_PO,0)")
                .Columns.Add("QTY_OPEN_OTHER_POS", GetType(System.Int64))
                .Columns.Add("PO_REFERENCE1")
                .Columns.Add("PO_ORDER_NO1")
                .Columns.Add("QTY_OPEN_PO_ORDER_NO1", GetType(System.Int64))
                .Columns.Add("PO_REFERENCE2")
                .Columns.Add("PO_ORDER_NO2")
                .Columns.Add("QTY_OPEN_PO_ORDER_NO2", GetType(System.Int64))
            End With

            With .Tables.Add("POTSHPXL")
                .Columns.Add("PO_SHIPMENT_NO")
                .Columns.Add("PO_SHIPMENT_LNO", GetType(System.Int64))
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("PACKING_LNO", GetType(System.Int64))
                .Columns.Add("CTN_NO_START", GetType(System.Int64))
                .Columns.Add("CTN_NO_END", GetType(System.Int64))
                .Columns.Add("PO_ORDER_NO")
                .Columns.Add("PO_ORDER_LNO", GetType(System.Int64))
                .Columns.Add("SIZE")
                .Columns.Add("TOTAL_CTN", GetType(System.Int64))
                .Columns.Add("PER_CTN", GetType(System.Int64))
                .Columns.Add("TOTAL_PCS", GetType(System.Int64))
                .Columns.Add("GW", GetType(System.Decimal))
                .Columns.Add("NW", GetType(System.Decimal))
                .Columns.Add("TTL_GW", GetType(System.Decimal))
                .Columns.Add("TTL_NW", GetType(System.Decimal))
                .Columns.Add("MEAS")
                .Columns.Add("IS_SPLIT")
                .Columns.Add("WORKBOOK")
                .Columns.Add("WORKSHEET")
                .Columns.Add("CONTAINER_NO")
                .PrimaryKey = New DataColumn() { .Columns("PO_SHIPMENT_NO"), .Columns("PO_SHIPMENT_LNO"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE"), .Columns("PACKING_LNO")}
            End With

            Create_Relation("POTSHIPR", "POTSHIP3", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO,STYLE_CODE,COLOR_CODE")
            .Tables("POTSHIPR").Columns("QTY_SHP").Expression = "SUM(CHILD(POTSHIPR_POTSHIP3).PO_QTY_SHP)"
            Create_Relation("POTSHIPR", "POTSHIP8", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO,STYLE_CODE,COLOR_CODE")
            .Tables("POTSHIPR").Columns("QTY_CTN").Expression = "SUM(CHILD(POTSHIPR_POTSHIP8).TOTAL_UNITS)"
            .Tables("POTSHIPR").Columns.Add("CBM", GetType(System.Decimal), "ISNULL(SUM(CHILD(POTSHIPR_POTSHIP8).CBM),0)")

            .Tables("POTSHIP3").Columns.Add("CBM", GetType(System.Decimal)) ', "ISNULL(PARENT(POTSHIPR_POTSHIP3).CBM,0) * ISNULL(PO_QTY_SHP,0) / ISNULL(PARENT(POTSHIPR_POTSHIP3).QTY_SHP,0)")
            .Tables("POTSHIP2").Columns.Add("TOTAL_CBM", GetType(System.Decimal))
            ' If cost_calc Or receipt_mode Then .Tables("POTSHIP2").Columns("TOTAL_CBM").Expression = "ISNULL(SUM(CHILD(POTSHIP2_POTSHIP3).CBM),0)"
            .Tables("POTSHIP2").Columns("TOTAL_CBM").Expression = "ISNULL(SUM(CHILD(POTSHIP2_POTSHIP3).CBM),0)"

            Create_Relation("POTSHIP1", "POTSHIP2", "PO_SHIPMENT_NO")

            With .Tables("POTSHIP1")
                .Columns.Add("TOTAL_DUTY", GetType(System.Decimal), "SUM(CHILD(POTSHIP1_POTSHIP2).TOTAL_DUTY)")
                .Columns.Add("TOTAL_WEIGHT_FACTOR", GetType(System.Decimal), "SUM(CHILD(POTSHIP1_POTSHIP2).TOTAL_WEIGHT_FACTOR)")
                .Columns.Add("TOTAL_CBM", GetType(System.Decimal), "SUM(CHILD(POTSHIP1_POTSHIP2).TOTAL_CBM)")
            End With

            ASCMAIN1.sql = "Select POTSHIP5.*, POTCATG1.COST_CATGY_DESC" _
                & " from POTSHIP5,POTCATG1" _
                & " where POTCATG1.COST_CATGY_CODE (+) = POTSHIP5.COST_CATGY_CODE"
            Create_TDA(.Tables.Add, "POTSHIP5", "**", 1)
            With .Tables("POTSHIP5").Columns
                .Add("CHARGEBACK_AMT", GetType(System.Decimal))
                .Add("NET_COST", GetType(System.Decimal), "ISNULL(LANDING_COST_AMT,0) - ISNULL(CHARGEBACK_AMT,0)")
                .Add("LANDING_COST_W", GetType(System.Decimal), "IIF(LANDING_COST_DIST='W',NET_COST,0)")
                .Add("LANDING_COST_D", GetType(System.Decimal), "IIF(LANDING_COST_DIST='D',NET_COST,0)")
                .Add("LANDING_COST_T", GetType(System.Decimal), "IIF(LANDING_COST_DIST='T',NET_COST,0)")
                .Add("LANDING_COST_M", GetType(System.Decimal), "IIF(LANDING_COST_DIST='M',NET_COST,0)")
                .Add("LANDING_COST_F", GetType(System.Decimal), "IIF(LANDING_COST_DIST='F',NET_COST,0)")
                .Add("VOUCHER_NO")
                .Add("VEND_CODE")
                .Add("PO_SHIPMENT_LNO_DIST", GetType(System.Int32))
            End With

            For Each TABLE_NAME As String In New String() {"POTORDR1", "POTORDR2"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*")
                .Tables.Add(TABLE_NAME & "_SPLIT")
                .Tables(TABLE_NAME & "_SPLIT").Merge(.Tables(TABLE_NAME))
            Next
            With .Tables("POTORDR2_SPLIT").Columns
                .Add("QTY_PACKED", GetType(System.Decimal))
                .Add("PO_ORDER_LNO_ORIG", GetType(System.Int64))
                .Add("PO_SHIPMENT_LNO", GetType(System.Int64))
                .Add("STYLE_DESC", GetType(System.String))
                .Add("COLOR_DESC", GetType(System.String))
            End With

            Create_TDA(.Tables.Add, "ICTIREC1", "*")
            Create_TDA(.Tables.Add, "ICTIREC2", "*")

            '  Dim sql_cols As String = "PO_COST_VCOST, PO_COST_MATLS, PO_COST_VCOST_DZ, PO_COST_MATLS_DZ, PO_COST_OTHER, PO_COST_COMM, PO_COST_QUOTA, PO_COST_QUOTA_DF, SHIP_COST_CHANGE_DATE, SHIP_COST_CHANGE_USER"
            ASCMAIN1.sql = "Select Distinct POTORDR2.* from POTORDR2,POTSHIP3" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1"
            Create_TDA(.Tables.Add("POTORDR2_COSTS"), "POTORDR2", "**", 0, False, "V", 2)

            If ASCMAIN1.CLIENT = "NYA - NOT" Then

                ASCMAIN1.sql = "Select Distinct POTORDR2.STYLE_CODE" & vbCrLf _
                    & ", DECODE(:PARM1,NULL,ICTSTYL1.DUTY_RATE_CODE,ICTSTYLC.DUTY_RATE_CODE) DUTY_RATE_CODE" & vbCrLf _
                    & " from ICTSTYL1,POTORDR2,POTSHIP3,ICTSTYLC" & vbCrLf _
                    & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM2" & vbCrLf _
                    & "   and ICTSTYLC.STYLE_CODE (+) = POTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and ICTSTYLC.COUNTRY_CODE (+) = :PARM3" & vbCrLf

                ASCMAIN1.sql = "Select X.STYLE_CODE, X.DUTY_RATE_CODE, NVL(ICTDUTY3.DUTY_RATE,ICTDUTY1.DUTY_RATE) DUTY_RATE" & vbCrLf _
                    & " from (" & ASCMAIN1.sql & ") X, ICTDUTY3, ICTDUTY1" & vbCrLf _
                    & " where ICTDUTY1.DUTY_RATE_CODE = X.DUTY_RATE_CODE" & vbCrLf _
                    & "   and ICTDUTY3.DUTY_RATE_CODE (+) = X.DUTY_RATE_CODE" & vbCrLf _
                    & "   and ICTDUTY3.OPS_YYYY (+) = :PARM4" & vbCrLf

                Create_TDA(.Tables.Add, "ICTSTYLD", "**", 0, False, "VVVV", 1)

            Else

                ASCMAIN1.sql = "Select Distinct POTORDR2.STYLE_CODE, ICTDUTY1.DUTY_RATE_CODE" & vbCrLf _
                    & ", NVL(ICTDUTY3.DUTY_RATE,ICTDUTY1.DUTY_RATE) DUTY_RATE" & vbCrLf _
                    & " from ICTDUTY3,ICTSTYL1,POTORDR2,POTSHIP3,ICTDUTY1" & vbCrLf _
                    & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1" & vbCrLf _
                    & "   and ICTDUTY1.DUTY_RATE_CODE = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                    & "   and ICTDUTY3.DUTY_RATE_CODE (+) = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                    & "   and ICTDUTY3.OPS_YYYY (+) = :PARM2"
                Create_TDA(.Tables.Add, "ICTSTYLD", "**", 0, False, "VV", 1)

            End If




            ASCMAIN1.sql = "Select Distinct POTORDR2.STYLE_CODE, ICTSTYL1.WEIGHT_CODE, ICTWGHT1.WEIGHT_FACTOR" & vbCrLf _
                & " from ICTWGHT1,ICTSTYL1,POTORDR2,POTSHIP3" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1" & vbCrLf _
                & "   and ICTWGHT1.WEIGHT_CODE (+) = ICTSTYL1.WEIGHT_CODE"
            Create_TDA(.Tables.Add, "ICTSTYLW", "**", 0, False, "V", 1)

            ' TURN OFF UPDATEABILITY IF NOT NEEDED FOR ALL OF THESE TABLES

            Create_TDA(.Tables.Add, "POTCNTT1", "*")

            ASCMAIN1.sql = "Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO " & vbCrLf _
            & ", POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO, POTORDR1.PO_REFERENCE" & vbCrLf _
            & ", POTORDR2.PO_COST SHIP3_VCOST, POTORDR2.PO_COST ORDR2_VCOST" & vbCrLf _
            & ", POTORDR2.PO_COST SHIP3_VCOST_DZ, POTORDR2.PO_COST ORDR2_VCOST_DZ" & vbCrLf _
            & ", POTORDR2.PO_COST SHIP3_MATLS, POTORDR2.PO_COST ORDR2_MATLS" & vbCrLf _
            & ", POTORDR2.PO_COST SHIP3_MATLS_DZ, POTORDR2.PO_COST ORDR2_MATLS_DZ" & vbCrLf _
            & ", POTORDR2.PO_COST SHIP3_OTHER, POTORDR2.PO_COST ORDR2_OTHER" & vbCrLf _
            & ", POTORDR2.PO_COST SHIP3_COMM, POTORDR2.PO_COST ORDR2_COMM" & vbCrLf _
            & ", POTORDR2.PO_COST SHIP3_QUOTA, POTORDR2.PO_COST ORDR2_QUOTA" & vbCrLf _
            & " from POTSHIP3,POTORDR2,POTORDR1 " & vbCrLf _
            & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTCOSTV", "**", 0, False, "V", 4)

            ASCMAIN1.sql = "Select POTORDR1.PO_REFERENCE" & vbCrLf _
            & ", POTORDR2.PO_COST_VCOST, POTORDR2.PO_COST_MATLS" & vbCrLf _
            & ", POTORDR2.PO_COST_VCOST_DZ, POTORDR2.PO_COST_MATLS_DZ" & vbCrLf _
            & ", POTORDR2.PO_COST_OTHER, POTORDR2.PO_COST_COMM, POTORDR2.PO_COST_QUOTA" & vbCrLf _
            & " from POTORDR1, POTORDR2" & vbCrLf _
            & " where POTORDR2.PO_ORDER_NO = :PARM1 and POTORDR2.PO_ORDER_LNO = :PARM2" & vbCrLf _
            & " and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTCOSTF", "**", 0, False, "VN", 4)

            Create_TDA(.Tables.Add, "ICTTRAN1", "*")
            Create_TDA(.Tables.Add, "ICTTRAN2", "*")

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then

                Dim sqlstuff As String = ""

                ASCMAIN1.sql = "SELECT COST_CATGY_CODE FROM POTCATG1 WHERE COST_CODE_REQUIRED = '1' "
                For Each rowPOTCATG1 As DataRow In ASCDATA1.GetDataTable.Select("", "COST_CATGY_CODE")
                    sqlstuff = sqlstuff & ",SUM(DECODE (POTLCST1.COST_CATGY_CODE, '" & rowPOTCATG1.Item("COST_CATGY_CODE") & "' , NVL(POTLCST2.COST_ACT_PO,0))) AS " & rowPOTCATG1.Item("COST_CATGY_CODE") & vbCrLf
                Next

                ASCMAIN1.sql = "SELECT POTSHIP1.PO_SHIPMENT_NO" & vbCrLf _
                & sqlstuff _
                & "FROM POTLCST2, POTLCST1, POTSHIP1 " & vbCrLf _
                & "WHERE POTLCST1.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO " & vbCrLf _
                & "AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" & vbCrLf _
                & "AND  nvl(POTSHIP1.COST_COMPLETE,'0') <> '1'" & vbCrLf _
                & "AND POTLCST1.COST_CATGY_CODE IN (SELECT COST_CATGY_CODE FROM POTCATG1 WHERE NVL(COST_CODE_REQUIRED, ' ') = '1') " & vbCrLf _
                & "GROUP BY POTSHIP1.PO_SHIPMENT_NO " & vbCrLf _
                & "ORDER BY POTSHIP1.PO_SHIPMENT_NO"
                Create_TDA(.Tables.Add, "POTSHIPS", "**", 0, False, "", 1)

                'New Receiving Notes Table for RGI
                ASCMAIN1.sql = "Select  POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO, POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO " & vbCrLf _
                & ", POTSHIP3.PO_QTY_SHP, WHTPREC3.PO_QTY_REC, WHTPREC3.PO_REC_NOTE, WHTPREC3.LOCATION_CODE " & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_BODY_CODE, POTORDR2.SUB_UNIT_PACK_QTY, POTORDR2.CARTON_PACK_QTY" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
                & " from WHTPREC3, POTSHIP3, POTORDR2, ICTSTYL1, POTORDR1 " & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and WHTPREC3.PO_SHIPMENT_NO(+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and WHTPREC3.PO_SHIPMENT_LNO(+) = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and WHTPREC3.PO_ORDER_NO(+) = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and WHTPREC3.PO_ORDER_LNO(+) = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1"
                Create_TDA(.Tables.Add, "WHTPREC3", "**", 0, True, "V", 4)
                With .Tables("WHTPREC3")
                    .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PO_QTY_SHP,0)-ISNULL(PO_QTY_REC,0)")
                    .Columns.Add("REC_LOC_QTY", GetType(System.Int64))
                End With
            End If

            If ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "Select WHTWREC1.*, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
                    & " from WHTWREC1, POTSHIP1 where POTSHIP1.PO_SHIPMENT_NO = WHTWREC1.PO_SHIPMENT_NO" & vbCrLf _
                    & " and WHTWREC1.WH_REC_STATUS = 'C'"
                Create_TDA(.Tables.Add, "WHTWRECX", "**", 0, False, "", 1)

                ASCMAIN1.sql = " Select C7.PO_SHIPMENT_NO, C7.PO_SHIPMENT_LNO," & vbCrLf _
                    & " C7.STYLE_CODE, C7.COLOR_CODE,  C7.WH_REC_NO," & vbCrLf _
                    & " Sum(Nvl(C7.CARTONS_RECEIVED,0) *  QTY) As UNITS_REC" & vbCrLf _
                    & " from WHTWREC1 C1,WHTWREC7 C7, WHTWREC8 C8" & vbCrLf _
                    & " Where C1.WH_REC_NO = C7.WH_REC_NO" & vbCrLf _
                    & " And C7.WH_REC_NO = C8.WH_REC_NO" & vbCrLf _
                    & " And C7.PO_SHIPMENT_NO = C8.PO_SHIPMENT_NO" & vbCrLf _
                    & " And C7.PO_SHIPMENT_LNO = C8.PO_SHIPMENT_LNO" & vbCrLf _
                    & " And C7.CARTON_NO = C8.CARTON_NO" & vbCrLf _
                    & " Group By C7.WH_REC_NO, C7.PO_SHIPMENT_NO, C7.PO_SHIPMENT_LNO," & vbCrLf _
                    & " C7.STYLE_CODE, C7.COLOR_CODE, " & vbCrLf _
                    & " C1.CONTAINER_NO"
                If ASCMAIN1.CLIENT = "RGI" Then
                    Create_TDA(.Tables.Add, "WHTWRECD", "**", 0, False, "", 5)
                Else
                    Create_TDA(.Tables.Add, "WHTWRECD", "**", 0, False, "", 4)
                End If
            End If

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                Create_TDA(.Tables.Add, "WHTPPKM1", "*")
                Create_TDA(.Tables.Add, "WHTPPKM2", "*")

                Create_TDA(.Tables.Add, "WHTWREC7", "*")
                Create_TDA(.Tables.Add, "WHTWREC8", "*")

                For Each TABLE_NAME As String In New String() {"WHT3PLR1", "WHT3PLR2", "WHT3PLR3"}
                    ASW.Add(TABLE_NAME, ASCMAIN1.Temp_Table("Select * from " & TABLE_NAME & " where ROWNUM < 1"))
                Next

                ASCMAIN1.sql = "Select WHT3PLR1.*, POTSHIP1.PO_SHIP_VESSEL" _
                    & " from " & IIf(ASCMAIN1.USER_ID = "gcv", "WHT3PLR1", ASW("WHT3PLR1")) & " WHT3PLR1,POTSHIP1" _
                    & " where POTSHIP1.PO_SHIPMENT_NO (+) = WHT3PLR1.PO_SHIPMENT_NO"
                IIf(ASCMAIN1.USER_ID = "gcv", "", ASCMAIN1.sql)
                Create_TDA(.Tables.Add, "WHT3PLR1", "**", 0)

                ASCMAIN1.sql = "Select WHT3PLR2.* " _
                    & " from " & IIf(ASCMAIN1.USER_ID = "gcv", "WHT3PLR2", ASW("WHT3PLR2")) & " WHT3PLR2"
                Create_TDA(.Tables.Add, "WHT3PLR2", "**", 0)
                ASCMAIN1.sql = "Select WHT3PLR3.* " _
                    & " from " & IIf(ASCMAIN1.USER_ID = "gcv", "WHT3PLR3", ASW("WHT3PLR3")) & " WHT3PLR3"
                Create_TDA(.Tables.Add, "WHT3PLR3", "**", 0)

            ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                ASCMAIN1.sql = "Select EDT944T1.*" _
                    & " from EDT944T1" _
                    & " where EDT944T1.EDI_PROCESS_IND = '0'"
                Create_TDA(.Tables.Add, "EDT944T1", "**", 0, False)

                ASCMAIN1.sql = "Select EDT944T2.* " _
                    & " from EDT944T2"
                Create_TDA(.Tables.Add, "WHT3PLR2", "**", 0, False)
                ASCMAIN1.sql = "Select EDT944T3.* " _
                    & " from EDT944T3"
                Create_TDA(.Tables.Add, "WHT3PLR3", "**", 0, False)
            End If

            ASCMAIN1.sql = "Select ICTIREC1.*, POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_ADV_DATE, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
                & ", POTSHIP2.CONTAINER_NO, POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.PO_SHIP_CTNS" & vbCrLf _
                & " from ICTIREC1,POTSHIP1,POTSHIP2" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = ICTIREC1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO = ICTIREC1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO = ICTIREC1.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and ICTIREC1.OPS_YYYYPP >= :PARM1" & vbCrLf _
                & "   and ICTIREC1.OPS_YYYYPP <= :PARM2"

            If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                ASCMAIN1.sql &= " and ICTIREC1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")"
            End If

            Create_TDA(.Tables.Add, "ICTIRECX", "**", 0, False, "VV", 1)

            'ASCMAIN1.sql = "Select ICTIREC2.*, POTORDR2.PO_QTY_SHP" & vbCrLf _
            '    & " from ICTIREC2,POTORDR2" & vbCrLf _
            '    & " where POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
            '    & "   and POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
            '    & "   and ICTIREC2.RECEIPT_NO = :PARM1"
            'Create_TDA(.Tables.Add, "ICTIREC2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select POTSHIP1.PO_SHIPMENT_NO" & vbCrLf _
                & ", COUNT (*) CONTAINER_COUNT, SUM (FREIGHT_AMT) FREIGHT_AMT" & vbCrLf _
                & " from POTSHIP1,POTSHIP4" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO in (" & vbCrLf _
                & "Select DISTINCT POTSHIP3.PO_SHIPMENT_NO from POTSHIP2,POTSHIP3,POTORDR1" & vbCrLf _
                & " where POTSHIP2.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & ") and POTSHIP4.PO_SHIPMENT_NO (+) = POTSHIP1.PO_SHIPMENT_NO" & vbCrLf _
                & " group by POTSHIP1.PO_SHIPMENT_NO"
            Create_TDA(.Tables.Add, "POTSHIPF", "**", 0, False, "V", 1)
            .Tables("POTSHIPF").Columns("CONTAINER_COUNT").DataType = GetType(System.Int64)

            ASCMAIN1.sql = "Select POTORDR1.VEND_CODE" & vbCrLf _
                & ", NVL(POTSHIP3.PO_COST_COMM,0) PO_COST_COMM, NVL(POTSHIP3.PO_COST_BUFFER,0) PO_COST_BUFFER, COUNT (*) RECORDS" & vbCrLf _
                & " from POTSHIP2,POTSHIP3,POTORDR1" _
                & " where POTSHIP2.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & " group by POTORDR1.VEND_CODE, NVL(POTSHIP3.PO_COST_COMM,0), NVL(POTSHIP3.PO_COST_BUFFER,0)"
            Create_TDA(.Tables.Add, "POTSHIPC", "**", 0, False, "V", 3)
            .Tables("POTSHIPC").Columns("RECORDS").DataType = GetType(System.Int64)
            ASCMAIN1.sql = "Select Distinct POTORDR1.VEND_CODE" & vbCrLf _
                & ", NVL(POTSHIP3.PO_COST_COMM,0) PO_COST_COMM, NVL(POTSHIP3.PO_COST_BUFFER,0) PO_COST_BUFFER, POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & " from POTSHIP2,POTSHIP3,POTORDR1" _
                & " where POTSHIP2.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTSHIPC2", "**", 0, False, "V", 4)
            Create_Relation("POTSHIPC", "POTSHIPC2", "VEND_CODE,PO_COST_COMM,PO_COST_BUFFER")


            ASCMAIN1.sql = "Select * from SOTORDP1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDP1", "**", 0, True, "V")
            .Tables("SOTORDP1").Columns.Add("INV_NO_PREV")

            ASCMAIN1.sql = "Select * from SOTORDP2 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDP2", "**", 0, True, "V")

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1, , , , "ORDR_PICK_SEQ,ORDR_STATUS") ' ORDR_DATE_CLOSED,ORDR_YYYYPP_CLOSED

            ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", , "ORDR_QTY_OPEN,ORDR_QTY_SHIP,ORDR_QTY_CANC,ORDR_STATUS")

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

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "SOTPICK1", "*")
            Create_TDA(.Tables.Add, "SOTPICK2", "*")
            Create_TDA(.Tables.Add, "SOTSHIP1", "*")

            ASCMAIN1.sql = "Select APTINVH1.*" _
               & " from APTINVH1" _
               & " where VOUCHER_NO IN (SELECT DISTINCT VOUCHER_NO FROM APTINVH5 WHERE PO_SHIPMENT_NO = :PARM1)"
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME from APTVEND1"
            Create_TDA(.Tables.Add, "APTCHCKV", "**", 0, False, "", 1)
            With .Tables("APTCHCKV").Columns
                .Add("SEL")
                .Add("AMT_SHP", GetType(System.Decimal))
                .Add("AMT_REC", GetType(System.Decimal))
                .Add("AMT_INV", GetType(System.Decimal))
                .Add("AMT_ADV", GetType(System.Decimal))
                .Add("AMT_PMT", GetType(System.Decimal))
            End With
            .Tables("APTCHCKV").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE, APTCHCK2.VOUCHER_NO" & vbCrLf _
                & ", APTINVH1.INV_TYPE, APTINVH1.INV_NUM, APTCHCK1.CHECK_AMT AMT, APTVEND1.VEND_NAME" & vbCrLf _
                 & " from APTVEND1,APTCHCK1,APTCHCK2,APTINVH1" & vbCrLf _
                 & " where APTVEND1.VEND_CODE = APTCHCK1.VEND_CODE and APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM"
            Create_TDA(.Tables.Add, "APTCHCKP", "**", 0, False, "", 0)

            Create_Relation("APTCHCKV", "APTCHCKP", "VEND_CODE")
            With .Tables("APTCHCKP").Columns
                .Add("SEL", GetType(System.String), "PARENT.SEL")
            End With


            ASCMAIN1.sql = "Select APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_NUM, APTINVH1.VEND_CODE, APTCHCK2.VOUCHER_NO" _
                & ", APTINVH1.INV_NUM, APTINVH1.INV_TYPE, APTINVH5.PO_SHIPMENT_NO, POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" _
                & ", APTCHCK1.CHECK_AMT AMT, APTVEND1.VEND_NAME" _
                & " from APTCHCK1,APTCHCK2, APTINVH5, POTSHIP1, APTVEND1, APTINVH1"
            Create_TDA(.Tables.Add, "APTCHCKQ", "**", 0, False, "", 0)

            Create_Relation("APTCHCKV", "APTCHCKQ", "VEND_CODE")
            With .Tables("APTCHCKQ").Columns
                .Add("SEL", GetType(System.String), "PARENT.SEL")
            End With




            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "Select " & vbCrLf _
                    & "POTSHIP2.PO_SHIPMENT_NO," & vbCrLf _
                    & "POTSHIP2.PO_SHIPMENT_LNO," & vbCrLf _
                    & "APTINVH1.VEND_CODE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_STATUS, APTINVH1.CHECK_NUM, APTINVH1.CHECK_DATE," & vbCrLf _
                    & "POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_ADV_DATE, POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.WHSE_CODE, " & vbCrLf _
                    & "POTSHIP2.CONTAINER_NO," & vbCrLf _
                    & "POTSHIP2.BOL_NO," & vbCrLf _
                    & "POTSHIP2.PO_SHIP_CTNS," & vbCrLf _
                    & "POTSHIP2.PO_SHIP_STATUS," & vbCrLf _
                    & "POTSHIP2.TRAN_NO," & vbCrLf _
                    & "POTSHIP2.OPS_YYYYPP," & vbCrLf _
                    & "POTSHIP2.PO_DATE_RECEIVED," & vbCrLf _
                    & "POTSHIP2.CONTAINER_SIZE," & vbCrLf _
                    & "POTSHIP2.COMM_INV_NO," & vbCrLf _
                    & "POTSHIP2.ACCRUAL_STATUS," & vbCrLf _
                    & "POTSHIP2.VOUCHER_NO" & vbCrLf _
                    & " from POTSHIP1,APTINVH1,(" & vbCrLf _
                    & "Select * FROM POTSHIP2 WHERE OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                    & " union " & vbCrLf _
                    & "Select * FROM POTSHIP2 WHERE PO_SHIP_STATUS = 'O'" & vbCrLf _
                    & " union " & vbCrLf _
                    & "Select * FROM POTSHIP2 WHERE (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) " & vbCrLf _
                    & " in (Select PO_SHIPMENT_NO, PO_SHIPMENT_LNO from APTINVH5,APTINVH1" & vbCrLf _
                    & "      where APTINVH5.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
                    & "        and APTINVH1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "')" & vbCrLf _
                    & ") POTSHIP2 where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                    & "  and APTINVH1.VOUCHER_NO (+) = POTSHIP2.VOUCHER_NO"
                sqlPOTSHIPP = ASCMAIN1.sql
                ASCMAIN1.sql &= " and rownum < 1"
                POTSHIPP = ASCMAIN1.Temp_Table()

                ASCDATA1.ExecuteSQL("Alter Table " & POTSHIPP & " Add SHP NUMBER (13,2)")
                ASCDATA1.ExecuteSQL("Alter Table " & POTSHIPP & " Add REC NUMBER (13,2)")
                ASCDATA1.ExecuteSQL("Alter Table " & POTSHIPP & " Add ACC NUMBER (13,2)")
                ASCDATA1.ExecuteSQL("Alter Table " & POTSHIPP & " Add OPN NUMBER (13,2)")
                ASCDATA1.ExecuteSQL("Alter Table " & POTSHIPP & " Add INV NUMBER (13,2)")

                ASCMAIN1.sql = "Select * from " & POTSHIPP
                Create_TDA(.Tables.Add, "POTSHIPP", "**", 0, False)

                Create_Relation("APTCHCKV", "POTSHIPP", "VEND_CODE")
                With .Tables("POTSHIPP").Columns
                    .Add("SEL", GetType(System.String), "PARENT.SEL")
                End With

            End If


            If ASCMAIN1.CLIENT = "VAN" Then
                If MENU_ITEM_OBJECT = "POFSHIP1" Then
                    tab0.Tabs("AT Shipments").Visible = True

                    ASCMAIN1.sql = "SELECT I.`ShipDate`, I.`Carrier`, Count (*) INVS" & vbCrLf _
                        & " from AT.`invhdr` I,POTIHDRA P" & vbCrLf _
                        & " where I.VAN_REF = P.VAN_REF AND P.STATUS = 'W'" & vbCrLf _
                        & " group by I.`ShipDate`, I.`Carrier`"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    Create_TDA(.Tables.Add, "ATSHIPS", "**", 0, False)

                    ASCMAIN1.sql = "Select I.* from AT.`invhdr` I, POTIHDRA P" & vbCrLf _
                        & " where I.VAN_REF = P.VAN_REF and P.STATUS = 'W'"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    Create_TDA(.Tables.Add, "ATINVHDR", "**", 0, False)

                    Create_Relation("ATSHIPS", "ATINVHDR", "ShipDate,Carrier")

                    ASCMAIN1.sql = "Select H.* from AT.`invhdr` I, AT.`packhdr` H, POTIHDRA P, POTPACKA A" & vbCrLf _
                        & " where I.VAN_REF = P.VAN_REF and P.STATUS = 'W'" & vbCrLf _
                        & "   and H.`invhdrkey` = P.INVOICE_HDR_KEY" & vbCrLf _
                        & "   and A.INVHDRKEY = H.`invhdrkey` and A.STATUS = 'W' and A.VAN_REF = H.VAN_REF"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    Create_TDA(.Tables.Add, "ATPACKHDR", "**", 0, False)

                    Create_Relation("ATINVHDR", "ATPACKHDR", "InvHdrKey")

                    ASCMAIN1.sql = "Select X.* from AT.`invhdr` I, AT.`packhdr` H, AT.`packbag` X, POTIHDRA P, POTPACKA A" & vbCrLf _
                        & " where I.VAN_REF = P.VAN_REF and P.STATUS = 'W'" & vbCrLf _
                        & "   and H.`invhdrkey` = P.INVOICE_HDR_KEY" & vbCrLf _
                        & "   and X.VAN_REF = H.VAN_REF" & vbCrLf _
                        & "   and X.`packhdrkey` = H.`packhdrkey`" & vbCrLf _
                        & "   and A.INVHDRKEY = H.`invhdrkey` and A.STATUS = 'W' and A.VAN_REF = H.VAN_REF"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    Create_TDA(.Tables.Add, "ATPACKBAG", "**", 0, False)

                    Create_Relation("ATPACKHDR", "ATPACKBAG", "VAN_REF,packhdrkey")

                    ASCMAIN1.sql = "Select X.* from AT.`invhdr` I, AT.`packhdr` H, AT.`packpo` X, POTIHDRA P, POTPACKA A" & vbCrLf _
                        & " where I.VAN_REF = P.VAN_REF and P.STATUS = 'W'" & vbCrLf _
                        & "   and H.`invhdrkey` = P.INVOICE_HDR_KEY" & vbCrLf _
                        & "   and X.VAN_REF = H.VAN_REF" & vbCrLf _
                        & "   and X.`packhdrkey` = H.`packhdrkey`" & vbCrLf _
                        & "   and A.INVHDRKEY = H.`invhdrkey` and A.STATUS = 'W' and A.VAN_REF = H.VAN_REF"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    Create_TDA(.Tables.Add, "ATPACKPO", "**", 0, False)

                    Create_Relation("ATPACKHDR", "ATPACKPO", "VAN_REF,packhdrkey")

                    ASCMAIN1.sql = "Select X.* from AT.`invhdr` I, AT.`packhdr` H, AT.`packcarton` X, POTIHDRA P, POTPACKA A" & vbCrLf _
                        & " where I.VAN_REF = P.VAN_REF and P.STATUS = 'W'" & vbCrLf _
                        & "   and H.`invhdrkey` = P.INVOICE_HDR_KEY" & vbCrLf _
                        & "   and X.VAN_REF = H.VAN_REF" & vbCrLf _
                        & "   and X.`packhdrkey` = H.`packhdrkey`" & vbCrLf _
                        & "   and A.INVHDRKEY = H.`invhdrkey` and A.STATUS = 'W' and A.VAN_REF = H.VAN_REF"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    Create_TDA(.Tables.Add, "ATPACKCARTON", "**", 0, False)

                    Create_Relation("ATPACKHDR", "ATPACKCARTON", "VAN_REF,packhdrkey")

                Else
                    tab0.Tabs("AT Shipments").Visible = False
                End If

            Else
                tab0.Tabs("AT Shipments").Visible = False
            End If


            If ASCMAIN1.CLIENT = "VAN" Then
                If MENU_ITEM_OBJECT = "POFSHIPC" Then
                    ASCMAIN1.sql = "Select POTSHIP5.*, POTSHIP1.COST_COMPLETE_OPS_YYYYPP" & vbCrLf _
                        & " from POTSHIP5,POTSHIP1" & vbCrLf _
                        & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP5.PO_SHIPMENT_NO" & vbCrLf _
                        & "   and POTSHIP5.COST_CATGY_CODE = 'TARIFF'"

                    Create_TDA(.Tables.Add, "POTSHIP5_ALL", "**", 0, False)
                Else
                    tab0.Tabs("Tariffs").Visible = False
                End If

            Else
                tab0.Tabs("Tariffs").Visible = False
            End If

            With .Tables.Add("POTPACKR")
                .Columns.Add("PO_SHIPMENT_NO")
                .Columns.Add("PO_SHIPMENT_LNO", GetType(System.Int64))
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("QTY_PCK", GetType(System.Int32))
                .Columns.Add("QTY_SHP", GetType(System.Int32))
                .Columns.Add("QTY_CTN", GetType(System.Int32))
                .Columns.Add("QTY_VAR", GetType(System.Int32), "ISNULL(QTY_SHP,0) - ISNULL(QTY_CTN,0)")
                .Columns.Add("COLOR_DESC")
                .Columns.Add("CONTAINER_NO")
                .Columns.Add("COMM_INV_NO")
                .PrimaryKey = New DataColumn() { .Columns("PO_SHIPMENT_NO"), .Columns("PO_SHIPMENT_LNO"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With

            ASCMAIN1.sql = "Select " & vbCrLf _
                    & "POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                    & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                    & ", POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & ", POTORDR2.PO_QTY_OPN, POTORDR2.PO_QTY_SHP" & vbCrLf _
                    & ", POTORDR1.PO_REFERENCE, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
                    & " from POTSHIP3,POTORDR2,ICTSTYL1,POTORDR1 " & vbCrLf _
                    & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and ROWNUM < 1"
            Create_TDA(.Tables.Add, "POTPACKD", "**", 0, False)

            Create_Relation("POTPACKR", "POTPACKD", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO,STYLE_CODE,COLOR_CODE")

            If ASCMAIN1.CLIENT = "VAN" Then
                sqlPOTVBKGX = "Select POTVBKG1.*,APTVEND1.VEND_NAME" & vbCrLf _
                    & " from POTVBKG1,APTVEND1" & vbCrLf _
                    & " where APTVEND1.VEND_CODE = POTVBKG1.VEND_CODE"
                ASCMAIN1.sql = sqlPOTVBKGX ' & "  and POTPACK1.OPS_YYYYPP = :PARM1"
                Create_TDA(.Tables.Add, "POTVBKGX", "**", 0, False, "")

                Create_TDA(.Tables.Add, "POTVBKG1", "*")
                Create_TDA(.Tables.Add, "POTVBKG2", "*", 1)
                Create_TDA(.Tables.Add, "POTVBKG3", "*", 1)

                Create_TDA(.Tables.Add, "POTPACK2", "*", 1)
                Create_TDA(.Tables.Add, "POTPACK3", "*", 1)

                tab0.Tabs("Bookings").Visible = ship_entry
            Else
                tab0.Tabs("Bookings").Visible = False
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                'finish changes for pack_qty after done with VAN
                ASCMAIN1.sql = "Select POTSHIP2.PO_SHIPMENT_NO,  POTSHIP2.PO_SHIPMENT_LNO, POTSHIP1.WHSE_CODE" & vbCrLf _
                    & ", POTSHIP1.PO_SHIP_VESSEL, POTSHIP2.CONTAINER_NO, POTSHIP4.CONTAINER_DATE_REC, POTSHIP2.ORDR_NO, POTSHIP2.PO_SHIP_CTNS" & vbCrLf _
                    & ", SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO, POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_PACK" & vbCrLf _
                    & " From POTSHIP1, POTSHIP2, SOTORDR1, POTSHIP4," & vbCrLf _
                    & " (Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                    & ", MIN(POTSHIP3.PO_ORDER_NO) PO_ORDER_NO, sum(POTSHIP3.PO_QTY_SHP)  PO_QTY_SHP, 0 PO_QTY_PACK" & vbCrLf _
                    & " From POTSHIP3" & vbCrLf _
                    & " GROUP BY POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO) POTSHIP3" & vbCrLf _
                    & " Where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                    & " And POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                    & " And POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                    & " And POTSHIP2.PO_SHIPMENT_NO = POTSHIP4.PO_SHIPMENT_NO" & vbCrLf _
                    & " And POTSHIP2.CONTAINER_NO = POTSHIP4.CONTAINER_NO" & vbCrLf _
                    & " And POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                    & " And POTSHIP2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                    & " And POTSHIP1.WHSE_CODE = 'NC'"
                Create_TDA(.Tables.Add, "POTPACKG", "**", 0, False, "")
                With .Tables("POTPACKG")
                    .Columns("PO_QTY_SHP").DataType = GetType(System.Int32)
                    .Columns("PO_QTY_PACK").DataType = GetType(System.Int32)
                    .Columns.Add("PO_QTY_BAL", GetType(System.Int32), "ISNULL(PO_QTY_SHP,0) - ISNULL(PO_QTY_PACK,0)")
                    .PrimaryKey = New DataColumn() { .Columns("PO_SHIPMENT_NO"), .Columns("PO_SHIPMENT_LNO")}
                End With

                ASCMAIN1.sql = "Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO, POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP3.PO_QTY_SHP, nvl(POTPCKS2.PO_QTY_PACK,0) PO_QTY_PACK" & vbCrLf _
                    & ", POTORDR2.ORDR_NO, POTORDR2.ORDR_LNO, SOTORDR2.CUST_SKU, SOTORDR2.CARTON_PACK_QTY, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                    & " From POTSHIP3, POTORDR2, POTSHIP2, POTSHIP1,SOTORDR1, SOTORDR2, " & vbCrLf _
                    & " (select POTPCKS2.PO_SHIPMENT_NO, POTPCKS2.PO_SHIPMENT_LNO, POTPCKS2.PO_ORDER_NO, POTPCKS2.PO_ORDER_LNO" & vbCrLf _
                    & ", sum(nvl(POTPCKS2.PO_QTY_PACK,0))PO_QTY_PACK" & vbCrLf _
                    & " from POTPCKS2" & vbCrLf _
                    & " group by POTPCKS2.PO_SHIPMENT_NO, POTPCKS2.PO_SHIPMENT_LNO" & vbCrLf _
                    & ", POTPCKS2.PO_ORDER_NO, POTPCKS2.PO_ORDER_LNO) POTPCKS2" & vbCrLf _
                    & " Where POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & " And POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                    & " And POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                    & " And POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                    & " And POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                    & " And POTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                    & " And POTORDR2.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                    & " And POTORDR2.ORDR_LNO = SOTORDR2.ORDR_LNO" & vbCrLf _
                    & " And POTSHIP3.PO_ORDER_NO = POTPCKS2.PO_ORDER_NO(+)" & vbCrLf _
                    & " AND POTSHIP3.PO_ORDER_LNO = POTPCKS2.PO_ORDER_LNO(+)" & vbCrLf _
                    & " AND POTSHIP3.PO_SHIPMENT_NO = POTPCKS2.PO_SHIPMENT_NO(+)" & vbCrLf _
                    & " AND POTSHIP3.PO_SHIPMENT_LNO = POTPCKS2.PO_SHIPMENT_LNO(+)" & vbCrLf _
                    & " And POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                    & " And POTSHIP2.ORDR_NO IS NOT NULL" & vbCrLf _
                    & " And POTSHIP1.WHSE_CODE = 'NC'"
                Create_TDA(.Tables.Add, "POTPACKH", "**", 0, False, "")
                With .Tables("POTPACKH")
                    .Columns("PO_QTY_SHP").DataType = GetType(System.Int32)
                    .Columns("PO_QTY_PACK").DataType = GetType(System.Int32)
                    .Columns.Add("PO_QTY_BAL", GetType(System.Int32), "ISNULL(PO_QTY_SHP,0) - ISNULL(PO_QTY_PACK,0)")
                    .Columns.Add("PO_PACK_CTNS", GetType(System.Int32), "ISNULL(PO_QTY_PACK,0) / ISNULL(CARTON_PACK_QTY,0)")
                    .PrimaryKey = New DataColumn() { .Columns("PO_SHIPMENT_NO"), .Columns("PO_SHIPMENT_LNO"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
                End With
                Create_Relation("POTPACKG", "POTPACKH", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")

                ASCMAIN1.sql = "select POTPCKS1.PACK_SLIP_NO ,POTPCKS1.PACK_SLIP_DATE ,POTPCKS1.WHSE_CODE ,POTPCKS1.CUST_STORE_NO, POTPCKS1.TRAILER_NO" & vbCrLf _
                    & ", ARTCUST2.CUST_NAME, ARTCUST2.CUST_ADDR1, ARTCUST2.CUST_ADDR2, ARTCUST2.CUST_CITY,  ARTCUST2.CUST_STATE, ARTCUST2.CUST_COUNTRY" & vbCrLf _
                    & ", ARTCUST2.CUST_ZIP_CODE ,POTPCKS1.INIT_OPER ,POTPCKS1.LAST_OPER ,POTPCKS1.INIT_DATE ,POTPCKS1.LAST_DATE, 'Z' AR_PARM_KEY" & vbCrLf _
                    & " from POTPCKS1, ARTCUST2" & vbCrLf _
                    & " where ARTCUST2.CUST_CODE = '171659'" & vbCrLf _
                    & " and ARTCUST2.CUST_ADDR_TYPE = 'MK'" & vbCrLf _
                    & " and ARTCUST2.CUST_ADDR_STATUS = 'A'" & vbCrLf _
                    & " and ARTCUST2.cust_addr_code = POTPCKS1.CUST_STORE_NO"
                Create_TDA(.Tables.Add, "POTPCKS1", "**", 0)

                ASCMAIN1.sql = "select POTPCKS2.PACK_SLIP_NO, POTPCKS2.PO_SHIPMENT_NO ,POTPCKS2.PO_SHIPMENT_LNO ,POTPCKS2.PO_ORDER_NO" & vbCrLf _
                    & ", POTPCKS2.PO_ORDER_LNO ,POTPCKS2.STYLE_CODE ,POTPCKS2.COLOR_CODE ,POTPCKS2.PO_QTY_PACK, POTORDR2.CUST_SKU" & vbCrLf _
                    & ", POTORDR2.CUST_STYLE_CODE, POTORDR2.CUST_COLOR_CODE, POTORDR2.STYLE_UOM, POTORDR2.CARTON_PACK_QTY, POTORDR2.CASE_CUBE" & vbCrLf _
                    & ", POTPCKS2.LOAD_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO, POTSHIP2.CONTAINER_NO" & vbCrLf _
                    & " from POTPCKS2, POTSHIP2, SOTORDR1," & vbCrLf _
                    & " (select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO, SOTORDR2.CUST_SKU, SOTORDR2.CUST_STYLE_CODE" & vbCrLf _
                    & ", SOTORDR2.CUST_COLOR_CODE, SOTORDR2.STYLE_UOM, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                    & "  from POTORDR2, SOTORDR2, ICTSTYL1" & vbCrLf _
                    & "  where POTORDR2.ORDR_NO =  SOTORDR2.ORDR_NO" & vbCrLf _
                    & "  and POTORDR2.ORDR_LNO =  SOTORDR2.ORDR_LNO" & vbCrLf _
                    & "  and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE) POTORDR2" & vbCrLf _
                    & " where POTPCKS2.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & " and POTPCKS2.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                    & " and POTSHIP2.PO_SHIPMENT_NO = POTPCKS2.PO_SHIPMENT_NO" & vbCrLf _
                    & " and POTSHIP2.PO_SHIPMENT_LNO = POTPCKS2.PO_SHIPMENT_LNO" & vbCrLf _
                    & " and SOTORDR1.ORDR_NO = POTSHIP2.ORDR_NO"
                Create_TDA(.Tables.Add, "POTPCKS2", "**", 0)
                With .Tables("POTPCKS2")
                    .Columns.Add("PO_QTY_BAL", GetType(System.Int32))
                    .Columns.Add("CARTONS", GetType(System.Int32), "ISNULL(PO_QTY_PACK,0) / ISNULL(CARTON_PACK_QTY,1)")
                    .Columns.Add("VOLUME", GetType(System.Int32), "ISNULL(PO_QTY_PACK,0) / ISNULL(CARTON_PACK_QTY,1) * ISNULL(CASE_CUBE,0)")
                    .Columns.Add("IN_ERR")
                End With

                Create_Relation("POTPCKS1", "POTPCKS2", "PACK_SLIP_NO")
                With .Tables("POTPCKS1").Columns
                    .Add("TOTAL_UNITS", GetType(System.Decimal), "SUM(CHILD(POTPCKS1_POTPCKS2).PO_QTY_PACK)")
                    .Add("TOTAL_CARTONS", GetType(System.Decimal), "SUM(CHILD(POTPCKS1_POTPCKS2).CARTONS)")
                    .Add("TOTAL_VOLUME", GetType(System.Decimal), "SUM(CHILD(POTPCKS1_POTPCKS2).VOLUME)")
                End With


                With .Tables.Add("SOTINVP0")
                    .Columns.Add("AR_PARM_KEY")
                    .Columns.Add("REMIT0")
                    .Columns.Add("REMIT1")
                    .Columns.Add("REMIT2")
                    .Columns.Add("REMIT3")
                    .Columns.Add("AR_PARM_REMIT_MESSAGE")
                    .Columns.Add("AR_PARM_DUNS_NO")
                    .Columns.Add("ADDRESS_LINE")
                    .Columns.Add("LOGO", GetType(System.Byte()))
                    .PrimaryKey = New DataColumn() { .Columns("AR_PARM_KEY")}
                End With

                Dim rowSOTINVP0 As DataRow = dst.Tables("SOTINVP0").NewRow
                With ROWs("ARTPARM1")
                    rowSOTINVP0.Item("AR_PARM_KEY") = "Z"
                    rowSOTINVP0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
                    rowSOTINVP0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
                    rowSOTINVP0.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                            & .Item("AR_PARM_REMIT_STATE") & " " _
                            & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                            & .Item("AR_PARM_REMIT_COUNTRY")
                    If .Item("AR_PARM_REMIT_PHONE") & "" <> "" And .Item("AR_PARM_REMIT_FAX") & "" <> "" Then
                        rowSOTINVP0.Item("REMIT3") = "" _
                            & "  Tel " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_PHONE")) _
                            & ", Fax " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_FAX"))
                    End If
                    rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
                    If 1 = 1 Then
                        rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") & vbCrLf & .Item("AR_PARM_REMIT_MESSAGE_EXPORT")
                    End If
                    rowSOTINVP0.Item("AR_PARM_DUNS_NO") = .Item("AR_PARM_DUNS_NO") & ""
                    rowSOTINVP0.Item("ADDRESS_LINE") = "" _
                        & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ADDR1") _
                        & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_CITY") _
                        & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_STATE") _
                        & " " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE") _
                        & IIf(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "" <> "" _
                          And ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & "" <> "", "" _
                              & ", Tel " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "") _
                              & ", Fax " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & ""), "")
                End With
                rowSOTINVP0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
                dst.Tables("SOTINVP0").Rows.Add(rowSOTINVP0)

                Create_TDA(.Tables.Add, "SOTSREP1", "*", 0, False)
                Create_TDA(.Tables.Add, "SOTSVIA1", "*", 0, False)
                ASCMAIN1.sql = "Select * from POTSHIP2"
                Create_TDA(.Tables.Add, "POTSHIP2_R", "**", 0, False)

            End If

        End With



        Bind_Controls(grpSOTORDR1, "SOTORDR1")
        '   Bind_Controls(grpCustomsDuty, "POTSHIP1")

        'Set_Read_Only(grpSOTORDR1, True)

        splCartonQ.Panel2Collapsed = Not (ASCMAIN1.CLIENT = "VAN")
        grdPOTSHIPQ.DataSource = dst.Tables("POTSHIPQ")
        Sort_grdColumns(grdPOTSHIPQ, "CTN_NO", True)

        grdPOTSHIPX.DataSource = dst.Tables("POTSHIPX")
        grdPOTSHIP2.DataSource = dst.Tables("POTSHIP2")
        grdPOTSHIP3.DataSource = dst.Tables("POTSHIP3")
        grdPOTSHIP4.DataSource = dst.Tables("POTSHIP4")
        grdPOTSHIP5.DataSource = dst.Tables("POTSHIP5")
        grdPOTSHIP7.DataSource = dst.Tables("POTSHIP7")
        grdPOTSHIP8.DataSource = dst.Tables("POTSHIP8")
        grdPOTSHIPR.DataSource = dst.Tables("POTSHIPR")
        grdSOTORDP1.DataSource = dst.Tables("SOTORDP1")
        grdSOTORDP2.DataSource = dst.Tables("SOTORDP2")
        grdPOTSHIPS.DataSource = dst.Tables("POTSHIPS")
        grdAPTINVH1.DataSource = dst.Tables("APTINVH1")
        grdPOTSHIPP.DataSource = dst.Tables("POTSHIPP")

        grdPOTSHPIE.DataSource = dst.Tables("POTSHPIE")
        grdPOTPACKR.DataSource = dst.Tables("POTPACKR")

        If (ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI") And receipt_mode Then
            tab0.Tabs("Warehouse Receipts").Visible = True
        Else
            tab0.Tabs("Warehouse Receipts").Visible = False
        End If

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            grdWHT3PLR1.DataSource = dst.Tables("WHT3PLR1")
            grdWHTWRECX.DataSource = dst.Tables("WHTWRECX")
            grdEDT944T1.Visible = False
        ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            grdEDT944T1.DataSource = dst.Tables("EDT944T1")
            grdWHT3PLR1.Visible = False
            grdEDT944T1.Visible = True
            grdWHT3PLR2.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grdWHT3PLR3.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        ElseIf ASCMAIN1.CLIENT = "RGI" Then
            grdWHTPREC3.DataSource = dst.Tables("WHTPREC3")
            grdWHTWRECX.DataSource = dst.Tables("WHTWRECX")
        End If

        grdWHT3PLR2.DataSource = dst.Tables("WHT3PLR2")
        grdWHT3PLR3.DataSource = dst.Tables("WHT3PLR3")

        grdICTIRECX.DataSource = dst.Tables("ICTIRECX")
        grdICTIREC2.DataSource = dst.Tables("ICTIREC2")
        grdPOTSHIPI.DataSource = dst.Tables("POTSHIPI")
        grdPOTSHIPF.DataSource = dst.Tables("POTSHIPF")
        grdPOTSHIPC.DataSource = dst.Tables("POTSHIPC")

        grdAPTCHCKV.DataSource = dst.Tables("APTCHCKV")
        grdAPTCHCKP.DataSource = dst.Tables("APTCHCKP")
        grdAPTCHCKQ.DataSource = dst.Tables("APTCHCKQ")

        If ASCMAIN1.CLIENT = "VAN" Then
            grdPOTVBKGX.DataSource = dst.Tables("POTVBKGX")
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            grdPOTPACKG.DataSource = dst.Tables("POTPACKG")
            grdPOTPCKS1.DataSource = dst.Tables("POTPCKS1")
            grdPOTPCKS2.DataSource = dst.Tables("POTPCKS2")
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            If MENU_ITEM_OBJECT = "POFSHIP1" Then
                grdATSHIPS.DataSource = dst.Tables("ATSHIPS")
            End If
            If MENU_ITEM_OBJECT = "POFSHIPC" Then
                grdPOTSHIP5_ALL.DataSource = dst.Tables("POTSHIP5_ALL")
                Create_Summary(grdPOTSHIP5_ALL, "LANDING_COST_AMT")
            End If
        End If

        If Val(ROWs("POTPARM1").Item("PO_PARM_DEF_PO_UM") & "") = 12 And Not receipt_mode And Not ship_entry Then
            optUD.Value = "D"
        Else
            optUD.Value = "U"
        End If

        If ASCMAIN1.CLIENT = "NYA" Then
            With grdPOTSHIP7.DisplayLayout.Bands(0)
                .Columns("CARTON_DIMS").Header.Caption = "Ctn Dims (cm)"
                .Columns("CARTON_WEIGHT").Header.Caption = "Ctn Wgt (kg)"
                .Columns("TOTAL_WEIGHT").Header.Caption = "Tot Wgt (kg)"

            End With
        End If

        Create_Summary(grdPOTSHIPX, "PO_SHIPMENT_NO", "Count")
        Create_Summary(grdPOTSHIPX, New String() {"LINES", "LINES_REC"})

        Create_Summary(grdPOTSHIP2, "PO_SHIPMENT_LNO", "Count")
        Create_Summary(grdPOTSHIP2, New String() {"PO_SHIP_CTNS", "CBM", "BOL_FEE", "TRUCKING", "TOTAL_WEIGHT", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_VAR"})
        If cost_calc Then
            Create_Summary(grdPOTSHIP2, New String() {
                           "TOTAL_WEIGHT_FACTOR", "TOTAL_CBM", "TOTAL_VCOST", "TOTAL_MATLS", "TOTAL_OTHER", "TOTAL_FIRST",
                         "TOTAL_COMM", "TOTAL_BUFFER", "TOTAL_QUOTA", "TOTAL_QUOTA_DF", "TOTAL_FREIGHT",
                         "TOTAL_CUSTOMS", "TOTAL_DUTY", "TOTAL_TRUCKING", "TOTAL_MISC", "TOTAL_LANDED"})
        Else
            Create_Summary(grdPOTSHIP2, New String() {
                           "TOTAL_CBM"})
        End If

        Create_Summary(grdPOTSHIP3, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTSHIP3, New String() {"PO_QTY_OPN", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_SHP_DZ", "PO_QTY_REC_DZ", "TOTAL_DUTY", "PO_QTY_VAR"})

        If cost_calc Then
            Create_Summary(grdPOTSHIP3, New String() {"PO_COST_VCOST_DZ", "PO_COST_MATLS_DZ", "PO_COST_OTHER_DZ", "FIRST_COST_TOTAL_DZ",
                                                      "PO_COST_QUOTA_DZ", "PO_COST_QUOTA_DF_DZ", "COMMISSION_COST_DZ",
                                                      "PO_COST_VCOST_UM", "PO_COST_MATLS_UM", "PO_COST_OTHER", "FIRST_COST_TOTAL",
                                                      "PO_COST_QUOTA", "PO_COST_QUOTA_DF", "COMMISSION_COST",
                                                      "PO_COST_CUSTOMS", "PO_COST_DUTY", "PO_COST_FREIGHT_IN", "PO_COST_LANDED",
                                                      "PO_COST_TRUCKING", "PO_COST_MISC"}, "Custom")

            grdPOTSHIP2.DisplayLayout.Bands(0).Columns("PO_DATE_RECEIVED_PORT").Hidden = Not (ASCMAIN1.CLIENT = "NYA")
            'lblPO_DATE_RECEIVED_PORT.Visible = (ASCMAIN1.CLIENT = "NYA")
            'dtePO_DATE_RECEIVED_PORT.Visible = (ASCMAIN1.CLIENT = "NYA")
        Else
            grdPOTSHIP2.DisplayLayout.Bands(0).Columns("PO_DATE_RECEIVED_PORT").Hidden = True
            'lblPO_DATE_RECEIVED_PORT.Visible = False
            'dtePO_DATE_RECEIVED_PORT.Visible = False

        End If

        Create_Summary(grdPOTSHIP4, "PO_SHIPMENT_LNO", "Count")
        Create_Summary(grdPOTSHIP4, New String() {"PO_SHIP_CTNS", "TOTAL_WEIGHT", "CBM", "TRUCKING", "FREIGHT_AMT"})

        Create_Summary(grdPOTSHIP5, "PO_SHIPMENT_LNO", "Count")
        Create_Summary(grdPOTSHIP5, New String() {"LANDING_COST_AMT", "LANDING_COST_W", "LANDING_COST_D", "LANDING_COST_T", "LANDING_COST_M", "LANDING_COST_F"})

        Create_Summary(grdPOTSHIPR, "STYLE_CODE", "Count")
        Create_Summary(grdPOTSHIPR, New String() {"QTY_SHP", "QTY_CTN", "QTY_VAR"})

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Create_Summary(grdPOTSHIPS, "PO_SHIPMENT_NO", "Count")
        End If

        Create_Summary(grdPOTSHIP7, "CARTON_NO", "Count")
        Create_Summary(grdPOTSHIP7, New String() {"CARTONS", "UNITS", "TOTAL_UNITS", "CARTON_VOLUME", "CBM", "TOTAL_WEIGHT"})

        Create_Summary(grdPOTSHIP8, "STYLE_CODE", "Count")
        Create_Summary(grdPOTSHIP8, New String() {"QTY", "UNITS", "TOTAL_UNITS", "CBM"})

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then

            Create_Summary(grdWHT3PLR1, "TRANS_SEQ", "Count")

            Create_Summary(grdWHT3PLR2, "ITEM_CODE", "Count")
            Create_Summary(grdWHT3PLR2, New String() {"RCVQTY"})

            Create_Summary(grdWHT3PLR3, "SUBNUM", "Count")
            Create_Summary(grdWHT3PLR3, New String() {"RCVQTY"})

            Create_Summary(grdWHTWRECX, "WH_REC_NO", "Count")
        ElseIf ASCMAIN1.CLIENT = "RGI" Then

            Create_Summary(grdWHTWRECX, "WH_REC_NO", "Count")
        End If


        Create_Summary(grdICTIRECX, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIRECX, New String() {"QTY_REC", "AMT_REC"})

        Create_Summary(grdICTIREC2, "RECEIPT_LNO", "Count")


        Create_Summary(grdPOTSHIPI, "STYLE_CODE", "Count")

        Create_Summary(grdPOTSHIPF, "PO_SHIPMENT_NO", "Count")
        Create_Summary(grdPOTSHIPF, New String() {"CONTAINER_COUNT", "FREIGHT_AMT"})

        Create_Summary(grdPOTSHIPC, "VEND_CODE", "Count")
        Create_Summary(grdPOTSHIPC, New String() {"RECORDS"}, , "POTSHIPC")

        Create_Summary(grdSOTORDP1, "INV_NO", "Count")
        Create_Summary(grdSOTORDP1, New String() {"INV_TOTAL_AMOUNT"})

        Create_Summary(grdSOTORDP2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDP2, New String() {"ORDR_QTY_SHIP", "ORDR_AMT_SHIP"})

        Create_Summary(grdAPTCHCKV, "SEL", "Count")
        Create_Summary(grdAPTCHCKV, "VEND_CODE", "Count")
        Create_Summary(grdAPTCHCKV, New String() {"AMT_SHP", "AMT_REC", "AMT_INV", "AMT_ADV", "AMT_PMT"})

        Create_Summary(grdAPTCHCKP, "VEND_CODE", "Count")
        Create_Summary(grdAPTCHCKP, New String() {"AMT"})

        Create_Summary(grdAPTCHCKQ, "VEND_CODE", "Count")
        Create_Summary(grdAPTCHCKQ, New String() {"AMT"})

        If ASCMAIN1.CLIENT = "VAN" Then
            Create_Summary(grdPOTVBKGX, "VBKG_NO", "Count")
            'Create_Summary(grdPOTVBKGX, New String() {"AMT"})
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            Create_Summary(grdWHTPREC3, "PO_ORDER_NO", "Count")
            Create_Summary(grdWHTPREC3, New String() {"PO_QTY_SHP", "PO_QTY_REC", "REC_LOC_QTY", "VARIANCE"})
        End If
        Show_Filter(grdWHTPREC3, True)



        With grdPOTSHPIE.DisplayLayout.Bands(0)
            '.Columns("PO_SHIPMENT_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTSHPIE.DisplayLayout.Bands(0).Columns

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray

                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "PO_REFERENCE1" Or gcol.Key = "PO_ORDER_NO1" Or gcol.Key = "QTY_OPEN_PO_ORDER_NO1" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf gcol.Key = "PO_REFERENCE2" Or gcol.Key = "PO_ORDER_NO2" Or gcol.Key = "QTY_OPEN_PO_ORDER_NO2" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else

                End If
            Next
        End With



        With grdPOTSHIPX.DisplayLayout.Bands(0)
            .Columns("PO_SHIPMENT_NO").Header.Fixed = True
            .Columns("PO_SHIP_VESSEL").Header.Fixed = True
            .Columns("PO_SHIP_ETA").Header.Fixed = True

            .Columns("LINES").Hidden = Not cost_calc
            .Columns("LINES_REC").Hidden = Not cost_calc
            .Columns("PO_DATE_RECEIVED_MIN").Hidden = Not cost_calc
            .Columns("WH_LINES_REC").Hidden = Not ASCMAIN1.CLIENT = "RGI"
        End With
        grdPOTSHIPX.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grdPOTSHIPX, True)

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            With grdPOTSHIPS.DisplayLayout.Bands(0)
                With .Columns("PO_SHIPMENT_NO")
                    .Header.Fixed = True
                    .Header.Caption = "Shipment"
                    .Width = 100
                End With
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If gcol.Key <> "PO_SHIPMENT_NO" Then
                        gcol.Width = 90
                    End If
                Next
            End With
        End If

        With grdICTIRECX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"RECEIPT_NO", "RECEIPT_DATE", "PO_SHIPMENT_NO", "VEND_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Fixed = True
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        With grdPOTSHIPX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_SHIP_VESSEL", "PO_SHIP_ETA", "PO_SHIP_REF_NO", "COST_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"PO_DATE_SHIPPED", "PORT_CODE", "WHSE_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Fuchsia
                ElseIf New String() {"LINES", "LINES_REC", "WH_LINES_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"LP_STATUS", "LP_XNO", "LP_XNO_INIT_DATE", "LP_XNO_INIT_OPER", "LP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                    If gcol.Key <> "LP_CODE" Then gcol.CellAppearance.ForeColor = Drawing.Color.Green
                ElseIf New String() {"INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE", "PO_SHIPMENT_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        With grdPOTSHIP2.DisplayLayout.Bands(0)
            .Columns("PO_SHIPMENT_LNO").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("BAR_CODE").Hidden = True
            .Columns("TOTAL_FIRST_CALC").Hidden = True
            .Columns("TOTAL_LANDED_CALC").Hidden = True
            .Columns("CONTAINER_SEAL_NO").Hidden = Not (ASCMAIN1.CLIENT = "NYA") Or Not receipt_mode

            For Each COLUMN_NAME As String In New String() _
                {"TOTAL_WEIGHT_FACTOR", "TOTAL_CBM", "TOTAL_VCOST", "TOTAL_MATLS", "TOTAL_OTHER", "TOTAL_FIRST",
                 "TOTAL_COMM", "TOTAL_BUFFER", "TOTAL_QUOTA", "TOTAL_QUOTA_DF", "TOTAL_FREIGHT",
                 "TOTAL_CUSTOMS", "TOTAL_DUTY", "TOTAL_TRUCKING", "TOTAL_MISC", "TOTAL_LANDED"}
                .Columns(COLUMN_NAME).Hidden = Not cost_calc
                .Columns(COLUMN_NAME).Width = 90
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Yellow
            Next
            .Columns("TOTAL_CBM").Hidden = False

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Else
                For Each COLUMN_NAME As String In New String() _
                     {"TOTAL_MATLS", "TOTAL_BUFFER", "TOTAL_QUOTA", "TOTAL_QUOTA_DF"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End If

            .Columns("TOTAL_FIRST").CellAppearance.BackColor = Drawing.Color.LightBlue
            .Columns("TOTAL_DUTY").CellAppearance.BackColor = Drawing.Color.LightGreen
            .Columns("TOTAL_LANDED").CellAppearance.BackColor = Drawing.Color.Lavender

            For Each COLUMN_NAME In New String() {"PO_SHIPMENT_LNO", "CONTAINER_NO", "BOL_NO", "COMM_INV_NO"} ' {"PO_SHIPMENT_LNO", "CONTAINER_NO", "BOL_NO", "COMM_INV_NO", "PO_SHIP_CTNS", "PO_SHIP_STATUS"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_SHIPMENT_LNO", "CONTAINER_NO", "CONTAINER_SEAL_NO", "BOL_NO", "COMM_INV_NO", "PO_SHIP_CTNS", "PO_SHIP_STATUS", "CONTAINER_SIZE", "ACTION"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"CBM_RATE", "CBM", "BOL_FEE", "TRUCKING", "TOTAL_WEIGHT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"TRAN_NO", "OPS_YYYYPP", "PO_SOURCE_DOC", "PO_DATE_RECEIVED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_VAR", "LINES", "LINES_EXACT", "LINES_OVER", "LINES_SHORT", "LINES_ZERO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            'For Each COLUMN_NAME As String In New String() {"ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_LNO", "STYLE_DESC", "COLOR_DESC"}
            '    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            'Next
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                .Columns("CONTAINER_SIZE").Hidden = True
                .Columns("TOTAL_WEIGHT").Hidden = True
            End If
            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                .Columns("CONTAINER_NO").Width = 120
                .Columns("BOL_NO").Width = 140
                .Columns("COMM_INV_NO").Width = 120
            End If
        End With

        grdPOTSHIP3.DisplayLayout.UseFixedHeaders = True
        With grdPOTSHIP3.DisplayLayout.Bands(0)
            .Columns("PO_SHIPMENT_LNO").CellActivation = UltraWinGrid.Activation.NoEdit

            For Each COLUMN_NAME As String In New String() _
                    {"EXT_WEIGHT_FACTOR", "CBM", "COST_CHANGED", "EXT_VCOST", "EXT_MATLS", "EXT_OTHER", "EXT_FIRST",
                     "EXT_COMM", "EXT_BUFFER", "EXT_QUOTA", "EXT_QUOTA_DF", "EXT_FREIGHT",
                     "EXT_CUSTOMS", "EXT_DUTY", "EXT_TRUCKING", "EXT_MISC", "EXT_LANDED", "EXT_FIRST_CALC", "EXT_LANDED_CALC"}
                .Columns(COLUMN_NAME).Hidden = True
            Next

            For Each COLUMN_NAME In New String() {"PO_ORDER_NO", "PO_ORDER_LNO", "PO_REFERENCE", "PO_DATE_SHIP_BY"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            If cost_calc Then
                For Each COLUMN_NAME In New String() {"STYLE_CODE", "COLOR_CODE"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
            End If

            For Each COLUMN_NAME In New String() {"LINE_EXACT", "LINE_OVER", "LINE_SHORT", "LINE_ZERO"}
                .Columns(COLUMN_NAME).Hidden = True
            Next

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {IIf(receipt_mode, "PO_QTY_REC", "PO_QTY_SHP"), "CLOSE_PO"}.Contains(gcol.Key) And Not cost_calc _
                Or New String() {"PO_COST_COMM", "PO_COST_BUFFER"}.Contains(gcol.Key) And cost_calc Then
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If

                If New String() {"PO_ORDER_NO", "PO_ORDER_LNO", "PO_REFERENCE", "PO_DATE_SHIP_BY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PO_QTY_SHP", "PO_QTY_REC", "NET_OPEN", "PO_QTY_VAR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                ElseIf New String() {"PO_QTY_SHP_DZ", "PO_QTY_REC_DZ", "NET_OPEN_DZ"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightSalmon
                ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"PO_UOM", "SUB_UNIT_PACK_QTY", "CARTON_PACK_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"PO_COST_QUOTA", "PO_COST_QUOTA_DF", "PO_COST_QUOTA_DZ", "PO_COST_QUOTA_DF_DZ",
                                     "PO_COST_BUFFER", "PO_COST_COMM", "COMMISSION_COST", "COMMISSION_COST_DZ",
                                     "PO_COST_VCOST_UM", "PO_COST_VCOST_DZ", "PO_COST_MATLS_UM", "PO_COST_MATLS_DZ",
                                     "PO_COST_OTHER", "PO_COST_OTHER_DZ", "FIRST_COST_TOTAL", "FIRST_COST_TOTAL_DZ",
                                     "PO_COST_FREIGHT_IN", "PO_COST_CUSTOMS", "PO_COST_DUTY", "PO_COST_TRUCKING", "PO_COST_MISC", "PO_COST_LANDED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                .Columns("DUTY_RATE_CODE").Style = UltraWinGrid.ColumnStyle.Edit
            Else
                ' .Columns("DUTY_RATE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If

            .Columns("FIRST_COST_TOTAL").CellAppearance.BackColor = grdPOTSHIP2.DisplayLayout.Bands(0).Columns("TOTAL_FIRST").CellAppearance.BackColor
            .Columns("FIRST_COST_TOTAL_DZ").CellAppearance.BackColor = grdPOTSHIP2.DisplayLayout.Bands(0).Columns("TOTAL_FIRST").CellAppearance.BackColor
            .Columns("PO_COST_DUTY").CellAppearance.BackColor = grdPOTSHIP2.DisplayLayout.Bands(0).Columns("TOTAL_DUTY").CellAppearance.BackColor
            .Columns("PO_COST_LANDED").CellAppearance.BackColor = grdPOTSHIP2.DisplayLayout.Bands(0).Columns("TOTAL_LANDED").CellAppearance.BackColor

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If gcol.Format = "#.000000" Then
                        gcol.Format = "#.0000"
                    End If
                Next
            End If
        End With


        grdPOTSHIP4.DisplayLayout.UseFixedHeaders = True
        With grdPOTSHIP4.DisplayLayout.Bands(0)
            .Columns("PO_SHIPMENT_LNO").CellActivation = UltraWinGrid.Activation.NoEdit
            For Each COLUMN_NAME In New String() {"PO_SHIPMENT_LNO"} ' , "CONTAINER_NO", "CONTAINER_TYPE_CODE", "PO_SHIP_CTNS", "PO_SHIP_STATUS"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"TOTAL_WEIGHT", "CBM", "TRUCKING", "FREIGHT_AMT", "CONTAINER_DATE_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                    If gcol.Key = "CONTAINER_DATE_REC" Then gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                ElseIf New String() {"PO_SHIPMENT_LNO", "CONTAINER_NO", "CONTAINER_TYPE_CODE", "PO_SHIP_CTNS", "PO_SHIP_STATUS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                ElseIf New String() {"INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        With grdPOTSHIP7.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"PO_SHIPMENT_LNO", "CARTON_NO", "PPK_CODE", "STYLES", "UNITS", "TOTAL_UNITS", "PPK_INNER_QTY_CALC"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
            .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        With grdPOTSHIP7.DisplayLayout.Bands(0)
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

        With grdPOTSHIP8.DisplayLayout.Bands(0)
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
            For Each COLUMN_NAME In New String() {"PO_SHIPMENT_LNO", "CARTON_NO", "STYLE_CODE", "COLOR_CODE", "UNITS", "TOTAL_UNITS"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With


        With grdPOTSHIPR.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"QTY_CTN", "QTY_SHP", "QTY_VAR"}.Contains(gcol.Key) Then
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

        With grdPOTSHIP5.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"COST_CATGY_CODE", "LANDING_COST_AMT", "LANDING_COST_DIST", "LANDING_COST_COMMENT"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If New String() {"PO_SHIPMENT_LNO", "COST_CATGY_CODE", "COST_CATGY_DESC", "LANDING_COST_AMT", "CHARGEBACK_AMT", "NET_COST", "LANDING_COST_DIST", "LANDING_COST_COMMENT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"LANDING_COST_W", "LANDING_COST_D", "LANDING_COST_T", "LANDING_COST_M", "LANDING_COST_F"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        With grdSOTORDP1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "INV_REF" Or gcol.Key = "INV_DATE" Or gcol.Key = "INV_NO_PREV" Or gcol.Key = "INV_COMMENT" Then
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
                If gcol.Key = "" Then
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

        If ASCMAIN1.CLIENT = "VAN" Then
            With grdPOTVBKGX.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"???"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                    ElseIf New String() {"???"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                    Else
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next
            End With
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            grdWHTPREC3.DisplayLayout.UseFixedHeaders = True
            With grdWHTPREC3.DisplayLayout.Bands(0)
                .Columns("PO_SHIPMENT_LNO").CellActivation = UltraWinGrid.Activation.NoEdit

                For Each COLUMN_NAME In New String() {"PO_ORDER_NO", "PO_ORDER_LNO", "STYLE_CODE"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next

                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"PO_REC_NOTE", "LOCATION_CODE"}.Contains(gcol.Key) Then
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    End If

                    If New String() {"PO_ORDER_NO", "PO_ORDER_LNO", "PO_REFERENCE", "PO_DATE_SHIP_BY"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    ElseIf New String() {"PO_QTY_SHP", "PO_QTY_REC", "NET_OPEN", "PO_REC_NOTE", "LOCATION_CODE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    ElseIf New String() {"VARIANCE", "REC_LOC_QTY"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Red
                    End If
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next

            End With
            grdPOTPACKG.DisplayLayout.UseFixedHeaders = True
            With grdPOTPACKG.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With
            grdPOTPACKG.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Synchronized
            For band As Integer = 0 To 1
                With grdPOTPACKG.DisplayLayout.Bands(band)
                    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        'gcol.CellAppearance.BackColor = Drawing.Color.Beige

                        If New String() {"CONTAINER_DATE_REC"}.Contains(gcol.Key) Then
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                            If gcol.Key = "CONTAINER_DATE_REC" Then gcol.Format = "MM/dd/yy"
                        ElseIf New String() {"PO_QTY_SHP", "PO_QTY_PACK", "PO_QTY_BAL"}.Contains(gcol.Key) Then
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                        ElseIf New String() {"PO_PACK_CTNS", "PO_SHIP_CTNS"}.Contains(gcol.Key) Then
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        End If
                        gcol.Header.Appearance.BackColor = Drawing.Color.White
                        gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    Next
                End With
            Next
            grdPOTPCKS2.DisplayLayout.UseFixedHeaders = True
            With grdPOTPCKS2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.True
                .AllowUpdate = DefaultableBoolean.True
            End With
            With grdPOTPCKS2.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"PO_QTY_PACK", "LOAD_NO"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    End If

                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next
            End With
        End If

        With grdAPTCHCKV.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"AMT_SHP", "AMT_REC", "AMT_INV", "AMT_ADV", "AMT_PMT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                    gcol.Width = 110

                ElseIf New String() {"SEL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Fixed = True
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Fixed = True
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next

            .Columns("AMT_SHP").Hidden = True
            .Columns("AMT_REC").Hidden = True
        End With
        grdAPTCHCKV.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True


        With grdAPTCHCKP.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"AMT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                    gcol.Width = 110
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                    ' gcol.Header.Fixed = True
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With


        With grdAPTCHCKQ.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"AMT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                    gcol.Width = 110
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                    ' gcol.Header.Fixed = True
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        If ASCMAIN1.CLIENT = "VAN" Then


            Show_Filter(grdPOTSHIPP)
            grdPOTSHIPP.DisplayLayout.GroupByBox.Hidden = False
            With grdPOTSHIPP.DisplayLayout.Bands(0)
                .Columns("PO_SHIPMENT_NO").Header.Fixed = True
                .Columns("PO_SHIP_VESSEL").Header.Fixed = True
                .Columns("VEND_CODE").Header.Fixed = True
                .Columns("COMM_INV_NO").Header.Fixed = True
                .Columns("CONTAINER_NO").Header.Fixed = True
                .Columns("BOL_NO").Header.Fixed = True
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"PO_SHIPMENT_NO", "PO_SHIP_VESSEL", "VEND_CODE", "COMM_INV_NO", "CONTAINER_NO", "BOL_NO"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    ElseIf New String() {"PO_DATE_SHIPPED", "PO_SHIP_ETA", "PORT_CODE", "WHSE_CODE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Fuchsia
                    ElseIf New String() {"SHP", "REC", "ACC", "OPN", "INV"}.Contains(gcol.Key) Then
                        Create_Summary(grdPOTSHIPP, gcol.Key)
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    ElseIf New String() {"INV_NUM", "INV_DATE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    ElseIf New String() {"CHECK_NUM", "CHECK_DATE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                    Else
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next
            End With

            splAPTCHCKV.Panel1Collapsed = True ' SHOW Q AND NOT P
        End If

        cbeReceipts.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeReceipts.SelectedItem = cbeReceipts.Items(0)
        cbeReceipts2.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeReceipts2.SelectedItem = cbeReceipts2.Items(0)

        Bind_Controls(grpShipment, "POTSHIP1")

        'Show_Filter(grdWHT3PLR2, True)
        'Show_Filter(grdWHT3PLR3, True)

        ASCMAIN1.Add_Value_List(grdPOTSHIP2, "PO_SHIP_STATUS", Nothing, New String() {":", "O:In Transit", "X:Receive Now", "R:Reverse Now", "C:Received"}) ' R = RECEIVED
        ASCMAIN1.Add_Value_List(grdPOTSHIP4, "CONTAINER_TYPE_CODE", "Select CONTAINER_TYPE_CODE, CONTAINER_TYPE_DESC from POTCNTT1")
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.Add_Value_List(grdPOTSHIP5, "LANDING_COST_DIST", Nothing, New String() {":", "T:Truck", "M:Misc", "W:Customs", "D:Duty"})
        Else
            ASCMAIN1.Add_Value_List(grdPOTSHIP5, "LANDING_COST_DIST", Nothing, New String() {":", "D:Duty", "W:Customs", "T:Truck", "M:Misc", "F:Freight"})
        End If
        ASCMAIN1.Add_Value_List(grdPOTSHIPX, "LP_STATUS", Nothing, New String() {":", "1:Transmitted", "0:Not Transmitted"})

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.Add_Value_List(grdPOTVBKGX, "VBKG_STATUS", Nothing, New String() {":", "O:Open", "F:Finalized"})
        End If

        Check_InquiryMode()


        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            'grdPOTSHIP2.DisplayLayout.Bands(0).Columns("COMM_INV_NO").Hidden = True
            grdPOTSHIPX.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = True
            grdPOTSHIPX.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
            grdPOTSHIPX.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
            grdPOTSHIPX.DisplayLayout.Bands(0).Columns("ORDR_CUST_PO").Hidden = True

            With grdPOTSHIP5.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"LANDING_COST_F", "CTL_NO", "VOUCHER_NO", "VEND_CODE", "PO_SHIPMENT_LNO_DIST"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End With
        Else
            grdPOTSHIP7.DisplayLayout.Bands(0).Columns("CUSTOM_PPK").Hidden = True
            grdPOTSHIP7.DisplayLayout.Bands(0).Columns("CARTON_COMMENTS").Hidden = True
            grdPOTSHIPC.DisplayLayout.Bands(0).Columns("PO_COST_BUFFER").Hidden = True
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            grdPOTSHIP2.DisplayLayout.Bands(0).Columns("BOL_NO").Header.Caption = "FCR No"
            lblPO_SHIP_REF_NO.Text = "Ship BL#"
            grdPOTSHIP2.Text = "Commercial Invoice / FCR"
        End If

        If ROWs("POTPARM1").Item("PO_PARM_PO_SELECT_SHIP") = "S" Then optPOLines.Value = "S"

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            With grdPOTSHIP7.DisplayLayout.Bands(0)
                .Columns("PPK_INNER_QTY").Hidden = True
                .Columns("STYLES").Hidden = True
                .Columns("PPK_INNER_QTY_CALC").Hidden = True
            End With
            With grdPOTSHIP8.DisplayLayout.Bands(0)
                .Columns("PPK_INNER_QTY").Hidden = True
                .Columns("DOZENS").Hidden = True
                .Columns("UNITS").Hidden = True
            End With
        End If

        tabBOL.Tabs("AP Invoices").Visible = InquiryMode


        If ASCMAIN1.CLIENT = "RGI" Then
            Dim rows() As DataRow = ASCDATA1.GetDataTable("SELECT *  FROM WHTLPRT1").Select("")
            For Each row As DataRow In rows
                cbxLabelPrinter.Items.Add(row.Item("LABEL_PRINTER_ID"))
            Next
            cbxLabelPrinter.SelectedIndex = 0
            If Not InquiryMode Then
                tabBOL.Tabs("Receiving").Visible = False
                UltraExplorerBar1.Groups("Print Labels").Visible = False
            End If
        Else
            tabBOL.Tabs("Receiving").Visible = False
            UltraExplorerBar1.Groups("Print Labels").Visible = False
        End If

        cbeYPSFrom.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -120) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYPSFrom.SelectedItem = cbeYPSFrom.Items(0) ' cbeYPSFrom.Items(Val(Mid(ASCMAIN1.CYP, 5, 2)) - 1)
        cbeYPSTo.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -120) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYPSTo.SelectedItem = cbeYPSTo.Items(0)

        tab0.Tabs("Payments").Visible = (ASCMAIN1.CLIENT = "VAN") And MENU_ITEM_OBJECT = "POFSHIPI"
        tab0.Tabs("Open && Paid").Visible = (ASCMAIN1.CLIENT = "VAN") And MENU_ITEM_OBJECT = "POFSHIPI"
        tab0.Tabs("Glen Raven").Visible = (ASCMAIN1.CLIENT = "RGI") And ship_entry

        MakeTransparent(chkFixCasePacks)


        lblPlusDuty.Visible = (ASCMAIN1.CLIENT = "VAN") And MENU_ITEM_OBJECT = "POFSHIPC"
        numPlusDuty.Visible = (ASCMAIN1.CLIENT = "VAN") And MENU_ITEM_OBJECT = "POFSHIPC"
        btnPlusDuty.Visible = (ASCMAIN1.CLIENT = "VAN") And MENU_ITEM_OBJECT = "POFSHIPC"
        'Set_Read_Only(grpPackSlipHdr, True)
        Set_Read_Only(Absx1.CtlFor("PACK_SLIP_NO"), True)
        Set_Read_Only(Absx1.CtlFor("PACK_WHSE_CODE"), True)
        Set_Read_Only(Absx1.CtlFor("ADDRESS"), True)


    End Sub

    Sub Check_InquiryMode()

        tab0.Tabs("Shipments Cost Summary").Visible = cost_calc And (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
        tab0.Tabs("Costing Summaries").Visible = cost_calc
        tab0.Tabs("3PL Receipts").Visible = receipt_mode

        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
            tab0.Tabs("3PL Receipts").Visible = False
        End If

        tab0.Tabs("Receipts History").Visible = InquiryMode
        With UltraExplorerBar1.Groups("Screen Control")
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
            .Items("Delete").Visible = Not InquiryMode
            .Items("Select").Visible = receipt_mode
        End With

        tabBOL.Tabs("Container Summary").Visible = cost_calc Or InquiryMode

        With grdPOTSHIP3.DisplayLayout.Bands(0)
            .Columns("PO_QTY_VAR").Hidden = Not InquiryMode
            For Each COLUMN_NAME As String In New String() {"PO_QTY_SHP", "PO_QTY_REC", "NET_OPEN"}
                .Columns(COLUMN_NAME).Hidden = True
                .Columns(COLUMN_NAME & "_DZ").Hidden = True
            Next
            For Each COLUMN_NAME As String In New String() _
                {"PO_COST_FREIGHT_IN", "PO_COST_DUTY", "PO_COST_CUSTOMS", "WEIGHT_FACTOR", "CBM", "PO_COST_LANDED", "PO_COST_TRUCKING", "PO_COST_MISC"}
                .Columns(COLUMN_NAME).Hidden = True
            Next
            For Each COLUMN_NAME As String In New String() _
                {"PO_COST_VCOST", "PO_COST_MATLS", "FIRST_COST_TOTAL", "PO_COST_OTHER", "PO_COST_QUOTA", "PO_COST_QUOTA_DF", "COMMISSION_COST"}
                .Columns(COLUMN_NAME).Hidden = True
                .Columns(COLUMN_NAME & "_DZ").Hidden = True
            Next
            For Each COLUMN_NAME As String In New String() {"DUTY_RATE", "CLOSE_PO", "DUTY_RATE_CODE", "SUB_BODY_CODE"}
                .Columns(COLUMN_NAME).Hidden = True
            Next

            .Columns("PO_COMM_PAYABLE_TO_BRKR").Hidden = Not (ASCMAIN1.CLIENT = "NYA")

            If ship_entry Then
                For Each COLUMN_NAME As String In New String() {"FOB_CMT", "PO_COST_VCOST_UM", "PO_COST_MATLS_UM",
                                                                "PO_COST_OTHER", "PO_COST_COMM", "PO_COST_BUFFER", "PO_QTY_UOM"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
                If ASCMAIN1.CLIENT = "NYA" Then
                    .Columns("PO_COST_COMM").Hidden = False
                    .Columns("PO_COMM_PAYABLE_TO_BRKR").Hidden = False
                End If
            End If

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Else
                For Each COLUMN_NAME As String In New String() {"PO_COST_MATLS_UM", "PO_COST_QUOTA", "PO_COST_QUOTA_DF", "PO_COST_BUFFER", "FOB_CMT"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End If
        End With

        'With grdPOTSHIP4.DisplayLayout.Bands(0)
        '    .Columns("CONTAINER_TYPE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        'End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTSHIP2, grdPOTSHIP3, grdPOTSHIP4, grdPOTSHIP5, grdWHTPREC3}
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                If grd.Name = "grdWHTPREC3" Then
                Else
                    .AllowUpdate = DefaultableBoolean.False
                End If
            End With
        Next

        'frmCodes.Enabled = True
        'cmdRefreshFC.Visible = False
        'frmFrtOpt.Visible = False

        With grdPOTSHIP2.DisplayLayout.Bands(0)
            .Columns("ACTION").Hidden = Not ship_entry And Not receipt_mode
        End With

        If ship_entry Then
            grdPOTSHIP4.Parent = splBOL.Panel2

            With grdPOTSHIP2.DisplayLayout.Bands(0)
                .Columns("ACTION").Header.Caption = "POs"
                .Columns("ACTION").Width = 40
                .Columns("ACTION").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("PO_SHIP_STATUS").CellActivation = UltraWinGrid.Activation.NoEdit
            End With
            With grdPOTSHIP3.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"COST_CHANGED"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
                For Each COLUMN_NAME As String In New String() {"PO_QTY_SHP", "NET_OPEN", "CLOSE_PO"}
                    .Columns(COLUMN_NAME).Hidden = False
                Next

                .Columns("CLOSE_PO").Hidden = True

            End With
            With grdPOTSHIP4.DisplayLayout.Bands(0)
                .Columns("PO_SHIP_STATUS").CellActivation = UltraWinGrid.Activation.NoEdit
            End With

            chkFlag.Visible = True
            optUD.Visible = False
            'frmFrtOpt.Enabled = False

            With grdPOTSHIP3.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PO_QTY_REC", "NET_OPEN"}
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns(COLUMN_NAME & "_DZ").CellActivation = UltraWinGrid.Activation.NoEdit
                Next
            End With

            ' tabPOTSHIP1.Tabs("Containers").Visible = False
            tabBOL.Tabs("Customs / Other").Visible = False
        End If

        If receipt_mode Then
            With grdPOTSHIP2.DisplayLayout.Bands(0).Columns("ACTION")
                .Header.Caption = "Receiving Options"
                .Width = 150
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .Style = UltraWinGrid.ColumnStyle.EditButton
                .CellButtonAppearance.ImageHAlign = HAlign.Left
            End With
            With grdPOTSHIP2.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PO_SHIPMENT_LNO", "CONTAINER_NO", "BOL_NO", "COMM_INV_NO", "PO_SHIP_CTNS"} ', "STATUS"
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Next
            End With
            With grdPOTSHIP3.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PO_QTY_SHP", "PO_QTY_SHP_DZ", "NET_OPEN_DZ"}
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Next
            End With
            'frmCodes.Enabled = False
            'frmReceipts.Enabled = True
            ' tabPOTSHIP1.Tabs("Containers").Visible = False
            tabBOL.Tabs("Customs / Other").Visible = False
            tabBOL.Tabs("Cartons").Visible = False
            tabBOL.Tabs("Container Summary").Visible = False
        End If

        If InquiryMode Then
            'frmFrtOpt.Enabled = False
            chkFlag.Visible = True
            chkFlag.Enabled = False
            chkAir.Enabled = False
            chkCostComplete.Visible = True
            chkCostComplete.Enabled = False
        End If

        If cost_calc Then
            chkCostComplete.Visible = True

            With grdPOTSHIP3.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PO_COST_FREIGHT_IN", "PO_COST_TRUCKING", "PO_COST_MISC", "PO_COST_DUTY", "PO_COST_CUSTOMS", "PO_COST_LANDED", "WEIGHT_FACTOR", "CBM", "DUTY_RATE"}
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Next
            End With

            With grdPOTSHIP3.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() _
                    {"PO_COST_FREIGHT_IN",
                     "PO_COST_DUTY", "PO_COST_CUSTOMS", "WEIGHT_FACTOR", "CBM", "PO_COST_LANDED", "PO_COST_TRUCKING", "PO_COST_MISC"}
                    .Columns(COLUMN_NAME).Hidden = False
                    .Columns(COLUMN_NAME).Format = "#.0000"
                    .Columns(COLUMN_NAME).Width = 80
                Next

                For Each COLUMN_NAME As String In New String() _
                   {"DUTY_RATE", "DUTY_RATE_CODE", "SUB_BODY_CODE"}
                    .Columns(COLUMN_NAME).Hidden = False
                Next

                For Each COLUMN_NAME As String In New String() _
                    {"PO_COST_VCOST_DZ", "PO_COST_MATLS_DZ", "PO_COST_OTHER_DZ", "PO_COST_QUOTA_DZ",
                     "PO_COST_QUOTA_DF_DZ", "FIRST_COST_TOTAL_DZ", "COMMISSION_COST_DZ",
                     "PO_COST_VCOST_UM", "PO_COST_MATLS_UM", "PO_COST_OTHER", "PO_COST_QUOTA",
                     "PO_COST_QUOTA_DF", "FIRST_COST_TOTAL", "COMMISSION_COST"}
                    .Columns(COLUMN_NAME).Hidden = False
                    If COLUMN_NAME.EndsWith("DZ") Then
                        .Columns(COLUMN_NAME).Format = "#.00"
                    Else
                        .Columns(COLUMN_NAME).Format = "#.0000"
                    End If
                    .Columns(COLUMN_NAME).Width = 80
                Next
                For Each COLUMN_NAME As String In New String() {"STYLE_DESC", "CLOSE_PO", "PO_ORDER_LNO", "PO_DATE_SHIP_BY", "CARTON_PACK_QTY"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
                ' "PO_COST_VCOST_UN", "PO_COST_MATLS_UN",
                For Each COLUMN_NAME As String In New String() {"PO_COST_VCOST", "PO_COST_MATLS", "PO_COST", "NET_OPEN", "NET_OPEN_DZ", "COST_CHANGED"}
                    .Columns(COLUMN_NAME).Hidden = True
                    '  .Columns(COLUMN_NAME).Width = 80
                Next
                For Each COLUMN_NAME As String In New String() {"PO_QTY_SHP", "PO_QTY_SHP_DZ", "PO_QTY_REC", "PO_QTY_REC_DZ"}
                    .Columns(COLUMN_NAME).Hidden = False
                    .Columns(COLUMN_NAME).Width = 80
                Next
            End With

            With grdPOTSHIP2.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PO_SHIPMENT_LNO", "CONTAINER_NO", "BOL_NO", "COMM_INV_NO", "PO_SHIP_CTNS", "PO_SHIP_STATUS"}
                    If (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") And (COLUMN_NAME = "COMM_INV_NO" Or COLUMN_NAME = "BOL_NO") Then
                    Else
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
                    End If
                Next
            End With
            With grdPOTSHIP4.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"TOTAL_WEIGHT", "CBM", "TRUCKING", "FREIGHT_AMT"}.Contains(gcol.Key) Then
                    Else

                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    End If
                Next
            End With
            tabBOL.Tabs("Cartons").Visible = True
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            optFreight.ValueList.ValueListItems.Remove(2)
        Else
            'optFreight.ValueList.ValueListItems.Remove(1)
            'optFreight.ValueList.ValueListItems.Remove(0)
            chkFlag.Visible = False
            optUD.Visible = False

            With grdPOTSHIP3.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"SUB_BODY_CODE", "SUB_UNIT_PACK_QTY", "PO_QTY_UOM"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End With
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                ' Validate_Code("WHSE_CODE")


            Case "View"
                Validate_Code("PO_SHIPMENT_NO")

                If EMsg = "" Then

                    If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                        PO_SHIPMENT_NO = Absx1.txtFor("PO_SHIPMENT_NO").Text
                        rowPOTSHIP1 = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                        If Not TAC.TACMAIN1.NyaCanadaWhseList.Contains(rowPOTSHIP1.Item("WHSE_CODE") & "") Then '  <> "18" Then  
                            MsgBox("Invalid Selection")
                            Exit Sub
                        End If
                    End If

                End If

            Case "Edit", "Select"
                Validate_Code("PO_SHIPMENT_NO")

                PO_SHIPMENT_NO = Absx1.txtFor("PO_SHIPMENT_NO").Text

                ' If eItemKey = "Edit" Then
                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then Exit Sub
                End If

                'Lock all PO's on this shipment for exclusive use.
                Dim PO_List As String = ""
                If EMsg = "" Then
                    ASCMAIN1.sql = " Select Distinct PO_ORDER_NO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                        Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                            PO_List &= vbCrLf & PO_ORDER_NO
                            'Else
                            '    PO_locked &= PO_ORDER_NO & vbTab
                        End If

                        If PO_List <> "" Then
                            MsgBox(Mid(PO_List, 3), MsgBoxStyle.OkOnly,
                                   "Selected Shipment includes these PO's which are currently being edited somewhere")
                            Exit Sub
                        End If
                    Next
                End If
                'End If

                If EMsg = "" Then
                    If cost_calc Then
                        If rowPOTSHIP1.Item("COST_COMPLETE") & "" = "1" Then
                            EMsg &= vbCr & "Shipment " & PO_SHIPMENT_NO & " was Costed on " _
                                & Format(rowPOTSHIP1.Item("COST_COMPLETE_INIT_DATE"), "MM/dd/yy") _
                                & " by " & rowPOTSHIP1.Item("COST_COMPLETE_INIT_OPER")
                            Dim warning As String = ""
                            If rowPOTSHIP1.Item("COST_COMPLETE_OPS_YYYYPP") <> ASCMAIN1.CYP Then
                                warning = vbCrLf & vbCrLf & "* Warning *" _
                                    & vbCrLf & "   If you change the costs of this shipment," _
                                    & vbCrLf & "    the CGS of sales reported in prior periods may be altered" & vbCrLf
                            End If
                            'If rowPOTSHIP1.Item("COST_COMPLETE_OPS_YYYYPP") = ASCMAIN1.CYP Then
                            If MsgBox(EMsg & warning & vbCrLf & vbCrLf & "Would you like to Re-Open this Shipment for Costing?", MsgBoxStyle.YesNo, "Option to Re-Open Shipment for Further Costing") = MsgBoxResult.Yes Then
                                ASCMAIN1.sql = "Update POTSHIP1 Set COST_COMPLETE = NULL" _
                                    & ", COST_COMPLETE_INIT_DATE = NULL" _
                                    & ", COST_COMPLETE_INIT_OPER = NULL" _
                                    & ", COST_COMPLETE_OPS_YYYYPP = NULL" _
                                    & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
                                ASCDATA1.ExecuteSQL()
                                MsgBox("Shipment " & "" & " has been Re-Opened for Costing", MsgBoxStyle.OkOnly, "Verification")
                                EMsg = ""
                                EnforceConstraints(False)
                                rowPOTSHIP1 = Fill_Record("POTSHIP1", PO_SHIPMENT_NO)
                                EnforceConstraints(True)
                            End If
                            'Else
                            '    EMsg &= " - no further changes permitted"
                            'End If
                        End If
                    Else
                        Dim rowPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowPOTSHIP1.Item("WHSE_CODE"))


                        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                            If Not TAC.TACMAIN1.NyaCanadaWhseList.Contains(rowPOTSHIP1.Item("WHSE_CODE") & "") Then ' <> "18" Then
                                MsgBox("Invalid Selection")
                                Exit Sub
                            End If
                        End If

                        If eItemKey = "Select" Then

                            WHSE_TYPE = rowICTWHSE1.Item("WHSE_TYPE") & ""

                            If Not select_from_3PL_list And Not Select_from_Whse_Receipt Then
                                If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                                    MsgBox("You cannot manually enter a receipt against this shipment",
                                           vbOKOnly, "Selected Shipment is associated with a 3PL Warehouse")
                                    ASCMAIN1.sql = "Select Count (*) from POTSHIP2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and TRAN_NO is Not Null"
                                    If Val(ASCDATA1.GetDataValue) = 0 Then
                                        Exit Sub
                                    Else

                                        ASCMAIN1.sql = "Select Count (*) from POTSHIP2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and TRAN_NO is Not Null and PO_SHIP_STATUS = 'C'"
                                        Dim partially_received As Boolean = (Val(ASCDATA1.GetDataValue) > 0)

                                        If MsgBox("Do you want to Proceed so that you may De-Receive this Shipment" _
                                                  & IIf(partially_received, vbCrLf & " or manually enter a partial container receipt from an EDI 3PL", "") & "?",
                                              MsgBoxStyle.YesNo,
                                              "NOTE: Using this option, All Receipts on the Shipment must be De-Received") = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                End If

                                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

                                Else
                                    If rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1" Then
                                        If rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C" Then
                                            MsgBox("The Warehouse listed on Receipt" _
                                               & vbCrLf & " is set up with Locator and Carton Control",
                                               vbOKOnly, "Selected Shipment is NOT ok to receive Manually")
                                            Exit Sub
                                        Else
                                            If ASCMAIN1.CLIENT = "RGI" And Not Select_from_Whse_Receipt Then
                                                MsgBox("The Warehouse listed on Receipt" _
                                               & vbCrLf & " is set up with Locator controls.",
                                               vbOKOnly, "Selected Shipment is NOT ok to receive Manually")
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If Select_from_Whse_Receipt Then
                                    If rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
                                        MsgBox("The Warehouse listed on Receipt" _
                                               & vbCrLf & " is NOT set up as a Locatable Warehouse",
                                               vbOKOnly, "Selected Shipment is NOT associated with a Locatable Warehouse")
                                        Exit Sub
                                    End If
                                Else
                                    If rowICTWHSE1.Item("LP_CODE") & "" = "" Then
                                        MsgBox("The Warehouse listed on Receipt" _
                                               & vbCrLf & " is NOT set up as a 3PL Warehouse",
                                               vbOKOnly, "Selected Shipment is NOT associated with a 3PL Warehouse")
                                        Exit Sub
                                    End If
                                End If

                            End If
                        Else
                            Dim sqlw As String = " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
                            If rowPOTSHIP1.Item("LP_STATUS") & "" = "1" Then

                                Dim blnSkip3PL As Boolean = False

                                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                                    ASCMAIN1.sql = "SELECT COUNT (*) FROM POTSHIP2 WHERE PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' AND PO_SHIP_STATUS = 'O'"
                                    Dim openPOTSHIP2 As Integer = ASCDATA1.GetDataValue
                                    If openPOTSHIP2 = 0 Then

                                        If MsgBox("There are no more Open Containers on this shipment," _
                                                  & vbCrLf & " yet you are calling it up for Editing." _
                                               & vbCrLf & vbCrLf & "Any changes made must keep this shipment in balance with how it was Received.",
                                               MsgBoxStyle.OkCancel,
                                               "Verification - IMPORTANT") = MsgBoxResult.Cancel Then
                                            Exit Sub
                                        Else
                                            blnSkip3PL = True
                                        End If

                                    End If
                                End If

                                If rowICTWHSE1.Item("LP_CODE") & "" <> "" And Not blnSkip3PL Then

                                    If MsgBox("Would you like to try to retrieve this Shipment from the 3PL?" _
                                         & vbCrLf & vbCrLf & "Please Note: If this process is successful, you will need to Re-Transmit to 3PL",
                                      MsgBoxStyle.YesNo,
                                      "Shipment " & PO_SHIPMENT_NO & " has been Transmitted to 3PL; no further modifications permitted") = MsgBoxResult.Yes Then

                                        BeginTrans()

                                        Dim i As Integer = 1

                                        'If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                                        '    i = 1
                                        'ElseIf ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                                        '    i = 1
                                        'ElseIf ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                        '    i = 1
                                        'End If
                                        If i = 1 Then
                                            ASCDATA1.ExecuteSQL("UPDATE POTSHIP1 SET LP_STATUS = '0'" & sqlw)
                                        End If

                                        CommitTrans()
                                        If i = 1 Then
                                            MsgBox("Attempt to Retrieve Shipment " & PO_SHIPMENT_NO & " has Succeeded" _
                                                   & vbCrLf & vbCrLf & "Now proceeding to call up this Shipment for Editing",
                                                    MsgBoxStyle.OkOnly, "Verification")
                                        Else
                                            MsgBox("Attempt to Retrieve Shipment " & PO_SHIPMENT_NO & " has Failed" _
                                                   & vbCrLf & "" _
                                                   & vbCrLf & "You may make limited changes to this Shipment",
                                                   MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                                        End If
                                    Else
                                        Exit Sub
                                    End If

                                Else

                                    ' WAREHOUSE IS A LOCATED WAREHOUSE WITH CARTON CONTROL (THAT IS WHY WE ARE USING LP_STATUS)

                                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

                                        ASCMAIN1.sql = "Select * From POTSHIP2" & sqlw & " and WH_REC_NO is Not Null"
                                        Dim row As DataRow = ASCDATA1.GetDataRow
                                        If row Is Nothing Then
                                            ASCDATA1.ExecuteSQL("Update POTSHIP1 Set LP_STATUS = '0'" & sqlw)
                                        Else
                                            If ship_entry And ASCMAIN1.Running_in_VS Then
                                                Stop
                                                ' BEFORE THINKING ABOUT RELEASING THIS TO THE WILD
                                                ' still need to lock down things like changes and deleting POTSHIP7/8, changes to POTSHIP3
                                                ' STILL NEED TO INCORPORATE SWAP PO FEATURE
                                                ' STILL NEED TO PREVENT CHANGING A POTSHIP2 THAT HAS HAD AP RECORDED AGAINST IT   LIKE 017775 LINE 9
                                                ASCMAIN1.sql = "SELECT WH_REC_NO FROM WHTWREC1 WHERE WH_REC_NO IN (SELECT DISTINCT WH_REC_NO FROM POTSHIP2 " & sqlw & ") AND WH_REC_STATUS = 'P'"
                                                WH_REC_NOsInProcess.Clear()
                                                For Each rowW As DataRow In ASCDATA1.GetDataTable.Select("")
                                                    WH_REC_NOsInProcess.Add(rowW.Item("WH_REC_NO"))
                                                Next

                                                If WH_REC_NOsInProcess.Count = 0 Then
                                                    EMsg &= "Shipment " & PO_SHIPMENT_NO & " has been completely received by Whse"
                                                Else
                                                    MsgBox("Shipment " & PO_SHIPMENT_NO & " has been partially received by Whse" _
                                                        & vbCrLf & "Editing is Restricted", MsgBoxStyle.OkOnly, "Notification")
                                                End If
                                            Else
                                                EMsg &= "Shipment " & PO_SHIPMENT_NO & " has been (either partially or completely) received by Whse"
                                            End If
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If





            Case "Update"

                If receipt_mode Then
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowPOTSHIP1.Item("WHSE_CODE"))
                    If rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1" Then
                        If Not Select_from_Whse_Receipt Then
                            EMsg &= vbCr & "You cannot receive into a Locatable Warehouse from the Shipment"
                            EMsg &= vbCr & "You must select the Container from the Warehouse Receipts tab"
                        End If
                    End If
                    If Not select_from_3PL_list And Not Select_from_Whse_Receipt Then
                        If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                            If dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'X'").Length <> 0 Then
                                If dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'C'").Length <> 0 Then
                                    If MsgBox("Are you entering a Partial Receipt" _
                                              & vbCrLf & " (using a Shipment Line created by the Variance Qtys" _
                                              & vbCrLf & "  following the Initial Receipt)?" _
                                              & vbCrLf & vbCrLf & "(If not, you should respond with No)",
                                              MsgBoxStyle.YesNo, "This Shipment Appears to be a Partial Receiving") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    EMsg &= vbCr & "Selected Shipment is associated with a 3PL Warehouse"
                                    EMsg &= vbCr & "You cannot manually enter a receipt against this shipment"
                                End If
                            End If
                        End If
                    End If

                    Dim rowsbad() As DataRow = dst.Tables("SOTORDP1").Select("INV_DATE IS NULL")
                    If rowsbad.Length <> 0 Then
                        EMsg &= vbCr & "Invalid or empty Invoice Date for a Pro-Forma Invoice " & rowsbad(0).Item("INV_NO") & " on a Back-to-Back Order " & rowsbad(0).Item("ORDR_NO")
                    End If

                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        If Absx1.dteFor("PO_DATE_RECEIVED").Value & "" = "" Then
                            ' this will be trapped below
                        Else
                            If dst.Tables("SOTORDP1").Select("INV_DATE <> '" & Format(Absx1.dteFor("PO_DATE_RECEIVED").Value, "MM/dd/yyyy") & "'").Length <> 0 Then
                                If MsgBox("Invoice Date for Pro-Forma Invoice not the same as Date Received." _
                                 & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            End If

                            For Each row As DataRow In dst.Tables("SOTORDP1").Select("")
                                If row.Item("INV_DATE") & "" = "" OrElse
                                    Format(row.Item("INV_DATE"), "yyyyMM") <> ASCMAIN1.CYM Then
                                    EMsg &= vbCr & "Cannot Receive (or BTB Invoice) with a Date that is NOT in the Current Operations Period (" & ASCMAIN1.CYM & ")"
                                End If
                            Next

                        End If
                    End If

                    For Each rowSOTORDP1 As DataRow In dst.Tables("SOTORDP1").Select("ISNULL(INV_NO_PREV,'') <> ''")
                        Dim ORDR_NO As String = rowSOTORDP1.Item("ORDR_NO")
                        Dim INV_NO_PREV As String = rowSOTORDP1.Item("INV_NO_PREV")
                        If Not Check_INV_NO_PREV(ORDR_NO, INV_NO_PREV) Then
                            EMsg &= vbCr & "Previously Generated Invoice " & INV_NO_PREV & " does not match Receipt"
                        End If
                    Next

                    'Dim ZYP As String = Set_Period()
                    'If ZYP <> ASCMAIN1.CYP Then
                    '    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    '    Else
                    '        EMsg &= vbCr & "Cannot Receive into a Future Period (" & ZYP & ")"
                    '    End If
                    'End If


                    'If ASCMAIN1.CLIENT = "NYA" Then
                    '    ' SHOULD WE SKIP THIS IF WE ARE NOT INTERESTED IN CALCULATING DUTY?
                    '    Dim WHSE_COUNTRY As String = rowICTWHSE1.Item("WHSE_COUNTRY") & ""
                    '    Dim COUNTRY_CODE As String = ""
                    '    If WHSE_COUNTRY <> "" And WHSE_COUNTRY <> "USA" Then
                    '        COUNTRY_CODE = WHSE_COUNTRY
                    '    End If

                    '    For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'X'")
                    '        Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
                    '        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO))
                    '            Dim DUTY_RATE_CODE As String = rowPOTSHIP3.Item("DUTY_RATE_CODE") & ""
                    '            Dim STYLE_CODE As String = rowPOTSHIP3.Item("STYLE_CODE") & ""
                    '            If COUNTRY_CODE = "" Then
                    '                If DUTY_RATE_CODE.Length <> 12 Then
                    '                    EMsg &= vbCr & "Incorrect Duty Rate Code for Style " & STYLE_CODE & ":" & DUTY_RATE_CODE
                    '                End If
                    '            Else
                    '                If DUTY_RATE_CODE.Length <> 16 Or Not DUTY_RATE_CODE.EndsWith("-" & COUNTRY_CODE) Then
                    '                    EMsg &= vbCr & "Incorrect Duty Rate Code for Style " & STYLE_CODE & ":" & DUTY_RATE_CODE
                    '                End If
                    '            End If
                    '        Next
                    '    Next
                    'End If

                    If ASCMAIN1.CLIENT = "VAN" Then
                        If EMsg = "" Then
                            Dim WH_PARM_REC_VAR_WARNING_UNITS As Integer = Val(ROWs("WHTPARM1").Item("WH_PARM_REC_VAR_WARNING_UNITS") & "")
                            Dim WH_PARM_REC_VAR_WARNING_PWD As String = ROWs("WHTPARM1").Item("WH_PARM_REC_VAR_WARNING_PWD") & ""
                            Dim sqlw As String = String.Format("ISNULL(PO_QTY_VAR,0) > {0} OR ISNULL(PO_QTY_VAR,0) < -{0}", CStr(WH_PARM_REC_VAR_WARNING_UNITS))
                            If dst.Tables("POTSHIP2").Select(sqlw).Length > 0 Then
                                If MsgBox("There is a large Receiving Variance on this Receipt." & vbCrLf & vbCrLf & "Continue with Update anyway?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                    EMsg &= vbCr & "Large Receiving Variance - please check"
                                End If
                            End If
                        End If
                    End If
                End If


                If cost_calc Then
                    ' Get_Duty()
                    If STYLE_CODEs_No_Duty.Count <> 0 Then
                        If chkNoDuty.Checked Then
                        Else
                            If automated_cost_complete Then
                                ASCDATA1.ExecuteSQL("Insert into POTSHIP1_EMSG values ('" & PO_SHIPMENT_NO & "','No Containers Entered')")
                            Else
                                EMsg &= vbCr & "Some Styles have no Duty:" & Join(STYLE_CODEs_No_Duty.ToArray, ",")
                            End If
                        End If
                    End If

                    If ASCMAIN1.CLIENT = "NYA" Then
                        If Not chkNoDuty.Checked Then ' SHOULD WE SKIP THIS IF WE ARE NOT INTERESTED IN CALCULATING DUTY?
                            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowPOTSHIP1.Item("WHSE_CODE"))
                            Dim WHSE_COUNTRY As String = rowICTWHSE1.Item("WHSE_COUNTRY") & ""
                            Dim COUNTRY_CODE As String = ""
                            If WHSE_COUNTRY <> "" And WHSE_COUNTRY <> "USA" Then
                                COUNTRY_CODE = WHSE_COUNTRY
                            End If

                            For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'X'")
                                Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
                                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO))
                                    Dim DUTY_RATE_CODE As String = rowPOTSHIP3.Item("DUTY_RATE_CODE") & ""
                                    Dim STYLE_CODE As String = rowPOTSHIP3.Item("STYLE_CODE") & ""
                                    If COUNTRY_CODE = "" Then
                                        If DUTY_RATE_CODE.Length <> 12 Then
                                            EMsg &= vbCr & "Incorrect Duty Rate Code for Style " & STYLE_CODE & ":" & DUTY_RATE_CODE
                                        End If
                                    Else
                                        If DUTY_RATE_CODE.Length <> 16 Or Not DUTY_RATE_CODE.EndsWith("-" & COUNTRY_CODE) Then
                                            EMsg &= vbCr & "Incorrect Duty Rate Code for Style " & STYLE_CODE & ":" & DUTY_RATE_CODE
                                        End If
                                    End If
                                Next
                            Next
                        End If

                    End If


                    Dim FRT_TERMS As String = ""
                    If grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & "" <> "" Then
                        Dim ORDR_NO As String = grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & ""
                        Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO)
                        FRT_TERMS = rowSOTORDR1.Item("FRT_TERMS")
                    End If

                    If automated_cost_complete Then
                    Else
                        If rowPOTSHIP1.Item("WHSE_CODE") & "" = "FE" And FRT_TERMS = "PPA" Then
                            For Each ROW As DataRow In dst.Tables("POTSHIP5").Select("ISNULL(NET_COST,0) <> 0 AND ISNULL(CTL_NO,'') <> ''")
                                Dim CTL_NO As String = ROW.Item("CTL_NO") & ""
                                Dim rowPOTLCST1 As DataRow = LookUp("POTLCST1", CTL_NO)
                                If rowPOTLCST1 IsNot Nothing Then
                                    Dim COST_CATGY_CODE As String = rowPOTLCST1.Item("COST_CATGY_CODE") & ""

                                    ' Dim rowPOTCATG1 As DataRow = dst.Tables("POTCATG1").Rows.Find(COST_CATGY_CODE)
                                    Dim rowPOTCATG1 As DataRow = LookUp("POTCATG1", COST_CATGY_CODE)

                                    If rowPOTCATG1.Item("CHARGEBACK_IND") & "" = "1" Then
                                        If MsgBox("There are costs on this shipment that have not been charged back" _
                                                  & vbCrLf & "Continue Anyway?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Else

                    If ASCMAIN1.CLIENT = "NYA" Then
                        TAC.TACMAIN1.Check_Division_MixMatch(Me, EMsg, "POTSHIP3", "", rowPOTSHIP1.Item("WHSE_CODE"))
                    End If

                    If Not ship_entry Then
                        Get_Duty()
                        Get_Weight_Factor()

                        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3").Select _
                                                    ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"),
                                                    New String() {"STYLE_CODE", "DUTY_RATE", "WEIGHT_FACTOR"}).Rows
                                If rowPOTSHIP1.Item("COST_NO_DUTY") & "" <> "1" Then
                                    If Val(row.Item("DUTY_RATE") & "") = 0 Then
                                        EMsg &= vbCr & "Style " & row.Item("STYLE_CODE") & " Does not have a Duty Rate"
                                    End If
                                End If
                                If Val(row.Item("WEIGHT_FACTOR") & "") = 0 Then
                                    EMsg &= vbCr & "Style " & row.Item("STYLE_CODE") & " Does not have a Weight Factor"
                                End If
                            Next
                        Else

                        End If

                    End If
                End If

                If receipt_mode Then
                    'If Absx1.chkFor("REVIEW").Checked Then
                    '    EMsg &= vbCr & "This Shipment has been flagged as needing review. See Comments."
                    'End If

                    Dim STATUS_X As Integer = dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'X'").Length
                    Dim STATUS_R As Integer = dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'R'").Length
                    Dim STATUS_R_BTB As Integer = dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'R' AND ORDR_NO IS NOT NULL").Length
                    If STATUS_X <> 0 And STATUS_R <> 0 Then
                        EMsg &= vbCr & "You cannot mix Receipts with De-Receipts in the same Update - please process separately"
                    ElseIf STATUS_X = 0 And STATUS_R = 0 Then
                        EMsg &= vbCr & "No BOL/Containers Received"
                    ElseIf STATUS_R_BTB <> 0 And 1 <> 1 Then
                        'EMsg &= vbCr & "You may not Reverse the Receipt of a BTB Shipment"
                    Else
                        If STATUS_X <> 0 Then
                            Dim XYP As String = Set_Period()
                            If XYP <> ASCMAIN1.CYP Then
                                If MsgBox("You Are Receiving Into The Next Fiscal Month" _
                                          & vbCrLf & "Are You Sure This Is What You Want To Do?",
                                          MsgBoxStyle.YesNo, "Next Month Posting") = MsgBoxResult.No Then
                                    EMsg &= vbCr & "Next Month Posting Cancelled"
                                End If
                            End If
                            Dim LINES_REC As Int64 = 0
                            For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'X' or PO_SHIP_STATUS = 'R'")
                                Dim LINES_3 As Int32 = dst.Tables("POTSHIP3").Select("PO_SHIPMENT_NO = '" & rowPOTSHIP2.Item("PO_SHIPMENT_NO") & "' and PO_SHIPMENT_LNO = " & rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & " and ISNULL(PO_QTY_REC,0) <> 0").Length
                                LINES_REC += LINES_3
                                If ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "NC" And rowPOTSHIP2.Item("PO_SHIP_STATUS") = "X" And EMsg = "" Then
                                    Dim PO_QTY_PACK As Int64 = Val(ASCDATA1.GetDataValue("Select Sum(PO_QTY_PACK) from POTPCKS2 where PO_SHIPMENT_NO =:PARM1 and PO_SHIPMENT_LNO = :PARM2", "VV", New Object() {rowPOTSHIP2.Item("PO_SHIPMENT_NO"), rowPOTSHIP2.Item("PO_SHIPMENT_LNO")}))
                                    If PO_QTY_PACK <> Val(rowPOTSHIP2.Item("PO_QTY_SHP") & "") Then
                                        EMsg &= vbCr & "Glen Raven Receipts must be slpit when there is an unpacked balance."
                                    End If
                                End If
                            Next
                            If LINES_REC = 0 Then
                                EMsg &= vbCr & "No Lines with Qty Received"
                            End If
                            If Absx1.txtFor("PO_SOURCE_DOC").Text = "" Then
                                EMsg &= vbCr & "PO Source Document Required"
                            End If
                            If Absx1.dteFor("PO_DATE_RECEIVED").Value & "" = "" Then
                                EMsg &= vbCr & "Value for Date Received Required"
                            End If

                        End If
                    End If
                Else
                    If dst.Tables("POTSHIP2").Select.Length = 0 Then EMsg &= vbCr & "No BOLs Entered"
                    If dst.Tables("POTSHIP3").Select.Length = 0 Then EMsg &= vbCr & "No BOL Details Entered"

                    If dst.Tables("POTSHIP4").Select.Length = 0 Then
                        If EntryMode = "N" Then
                            Create_Containers_from_BOL()
                        Else
                            If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") And cost_calc Then
                            Else
                                If automated_cost_complete And ASCMAIN1.Running_in_VS Then
                                    ASCDATA1.ExecuteSQL("Insert into POTSHIP1_EMSG values ('" & PO_SHIPMENT_NO & "','No Containers Entered')")
                                Else
                                    EMsg &= vbCr & "No Containers Entered"
                                End If
                            End If

                        End If
                    End If

                    If Absx1.dteFor("PO_DATE_SHIPPED").Value & "" = "" _
                    Or Absx1.dteFor("PO_SHIP_ETA").Value & "" = "" Then
                        EMsg &= vbCr & "Shipped and ETA Dates are Required for all Shipments"
                    Else
                        If Format(Absx1.dteFor("PO_DATE_SHIPPED").Value, "yyyyMMdd") _
                        > Format(Absx1.dteFor("PO_SHIP_ETA").Value, "yyyyMMdd") Then
                            EMsg &= vbCr & "ETA Date cannot be earlier than Date Shipped"
                        End If
                    End If

                    If LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Destination Warehouse Specified"
                    End If

                End If
                If cost_calc Then
                    If Not cost_ind Then

                        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                            EMsg &= vbCr & "Click Calculate Costs and Verify before Updating"
                        Else
                            MsgBox("Costing data has changed and this update will not reflect those changes.",
                                    MsgBoxStyle.OkOnly, "Costing changed")
                        End If
                    End If

                End If
                If ship_entry Then
                    If ASCMAIN1.CLIENT = "VAN" Then
                        If eMsg_Booking <> "" Then
                            EMsg &= vbCr & "Errors During Booking Import"
                        End If
                    End If
                    If ASCMAIN1.CLIENT = "NYA" Then
                        If Absx1.txtFor("PORT_CODE_ORIG").Text = "" Then
                            EMsg &= vbCr & "Origination Port is Mandatory"
                        Else
                            If LookUp("ICTPORT1", Absx1.txtFor("PORT_CODE_ORIG").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid value specified for Origination Port"
                            End If
                        End If

                        If Absx1.txtFor("PORT_CODE_DEST").Text = "" Then
                            EMsg &= vbCr & "Destination Port is Mandatory"
                        Else
                            If LookUp("ICTPORT1", Absx1.txtFor("PORT_CODE_DEST").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid value specified for Destination Port"
                            End If
                        End If

                        If Absx1.txtFor("COST_CODE").Text = "" Then
                            EMsg &= vbCr & "Freight Cost Terms is Mandatory"
                        Else
                            If LookUp("ICTCOSTB", Absx1.txtFor("COST_CODE").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid value specified for Freight Cost Terms"
                            End If
                        End If
                    End If


                    ' check that there is only 1 vendor on each of the pos within a shipment lno
                    For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
                        Dim CONTAINER_NO As String = rowPOTSHIP2.Item("CONTAINER_NO") & ""
                        If CONTAINER_NO <> Trim(Replace(CONTAINER_NO, " ", "")) Then
                            EMsg &= vbCr & "Embedded Spaces in Container [" & CONTAINER_NO & "]. See Shipment Line " & rowPOTSHIP2.Item("PO_SHIPMENT_LNO")
                        End If
                        Dim sqlw As String = "PO_SHIPMENT_NO = '" & rowPOTSHIP2.Item("PO_SHIPMENT_NO") & "' and PO_SHIPMENT_LNO = " & rowPOTSHIP2.Item("PO_SHIPMENT_LNO")
                        Dim VEND_CODE As String = ""
                        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3").Select(sqlw), New String() {"PO_ORDER_NO"}).Rows
                            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                            Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                            If VEND_CODE = "" Then
                                VEND_CODE = rowPOTORDR1.Item("VEND_CODE")
                            Else
                                If VEND_CODE <> rowPOTORDR1.Item("VEND_CODE") Then
                                    EMsg &= vbCr & "Multiple Vendors referenced (" & VEND_CODE & "," & rowPOTORDR1.Item("VEND_CODE") & ")" & vbCr & " in POs combined into Shipment Line " & rowPOTSHIP2.Item("PO_SHIPMENT_LNO")
                                    Exit For
                                End If
                            End If

                        Next
                    Next

                    Dim rows0() As DataRow = dst.Tables("POTSHIP7").Select("STYLES = 0 OR TOTAL_UNITS = 0")

                    If rows0.Length <> 0 Then
                        EMsg &= vbCr & "There are Carton types defined with no styles or 0 units" _
                            & vbCr & " - you must delete carton types with no styles or 0 units" _
                            & vbCr & " - see PO Shipment Line " & rows0(0).Item("PO_SHIPMENT_LNO") & ", Carton Type " & rows0(0).Item("CARTON_NO")
                    End If

                    Dim rows_Cartonized_more_than_Shipped() As DataRow = dst.Tables("POTSHIPR").Select("ISNULL(QTY_CTN,0) > ISNULL(QTY_SHP,0)")
                    If rows_Cartonized_more_than_Shipped.Length <> 0 Then
                        EMsg &= vbCr & "Qty Cartonized Cannot be Greater than Qty Shipped (See Line " & rows_Cartonized_more_than_Shipped(0).Item("PO_SHIPMENT_LNO") & ")"
                    End If

                    Dim rows7() As DataRow = dst.Tables("POTSHIP7").Select("ISNULL(STYLES,0) = 0")
                    If rows7.Length <> 0 Then
                        EMsg &= vbCr & "There is a least 1 Carton Type defined with no Styles (See Line " & rows7(0).Item("PO_SHIPMENT_LNO") & ")"
                    End If

                    If Absx1.txtFor("PO_SHIP_VESSEL").Text = "" Then
                        EMsg &= vbCr & "Vessel is a required field for all Shipments"
                    End If

                    If ASCMAIN1.CLIENT = "VAN" And (packingFromXLS Or packingFromBooking) Then
                        If dst.Tables("POTSHPIE").Rows.Count <> 0 Then
                            EMsg &= vbCr & "Cannot Update with Import Errors"
                        End If
                    End If

                End If

                If ASCMAIN1.CLIENT = "VAN" Then
                    '            CHECK TO MAKE SURE CARTONS HAVE NOT BEEN CREATED OR EXISTS.
                    For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'R'")
                        Dim PO_SHIPMENT_LNO As Integer = rowPOTSHIP2.Item("PO_SHIPMENT_LNO")
                        ASCMAIN1.sql = " SELECT WHTBARC1.BAR_CODE" & vbCrLf _
                        & "  from WHTBARC1" & vbCrLf _
                        & "  where WHTBARC1.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                        & "    and WHTBARC1.PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & vbCrLf _
                        & "    and WHTBARC1.STATUS_CODE <> 'R'" & vbCrLf _
                        & "  group by WHTBARC1.BAR_CODE" & vbCrLf
                        Dim row As DataRow = ASCDATA1.GetDataRow
                        If row IsNot Nothing Then
                            If ASCMAIN1.Running_in_VS Then Stop
                            EMsg &= vbCr & "Cartons On This Shipment Have Been Cartonized"
                        Else
                            ASCMAIN1.sql = "Select WHTBARC1.BAR_CODE" & vbCrLf _
                            & "  from WHTBARC1" & vbCrLf _
                            & "  where WHTBARC1.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                            & "    and WHTBARC1.PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & vbCrLf _
                            & "    and WHTBARC1.STATUS_CODE = 'R'" & vbCrLf _
                            & "  group by WHTBARC1.BAR_CODE" & vbCrLf
                            row = ASCDATA1.GetDataRow
                            ' SHOULD WE CARE ABOUT THIS ON A REVERSE RECEIPT?
                            ' 02/15/2018 WJZ I DON'T THINK SO
                            ' I NEED TO REVERSE A RECEIPT IN TO WHSE_CODE = CHINA TODAY - A NON-LOCATOR WHSE
                            ' SO I AM MANUALLY SKIPPING AROUND THE CHECK
                            ' PROBABLY SHOULD CODE IT SO THAT THIS CHECK APPLIES ONLY IF VAN-LOCATOR-WHSE

                            Dim ok_if_no_cartons As Boolean = False

                            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                            If rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C" Or rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1" Then
                                ' NOT OK IF NO CARTONS
                            Else
                                ok_if_no_cartons = True
                            End If
                            If row Is Nothing And Not ok_if_no_cartons Then 'If row IsNot Nothing Then -- was riginally this but I think you wanted to check if NO records existed, Not if records rows existed
                                EMsg &= vbCr & "No Cartons Have Ever Been Created For This Shipment"
                            End If
                        End If
                    Next
                End If

                If ship_entry And EMsg = "" Then

                    'Check to Make sure there are not mutiple Warehouses between PO and Shipment.

                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    Dim whse_msg As String = ""
                    Dim whse_msg_count As Integer = 0
                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                            (dst.Tables("POTSHIP3").Select, New String() {"PO_ORDER_NO"}).Rows
                        Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                        If rowPOTORDR1.Item("WHSE_CODE") & "" <> WHSE_CODE Then
                            If whse_msg_count < 10 Then
                                whse_msg_count += 1
                                whse_msg &= vbCrLf & " PO:" & PO_ORDER_NO & " was destined for Whse:" & rowPOTORDR1.Item("WHSE_CODE")
                            Else
                                whse_msg &= vbCrLf & " ... (and more) ..."
                                Exit For
                            End If
                        End If
                    Next

                    If whse_msg <> "" Then
                        If MsgBox("Shipment is destined for Whse:" & WHSE_CODE & " and " _
                                  & vbCrLf _
                                  & whse_msg _
                                  & vbCrLf & vbCrLf _
                                  & "OK to Proceed?",
                            MsgBoxStyle.YesNo, "Destination Warehouse has Changed for PO") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If



                    'fix_ICTSTYL1_packs = False

                    'If ASCMAIN1.CLIENT = "NYA" Then
                    '    ASCMAIN1.sql = "Select PO_ORDER_NO, PO_ORDER_LNO, STYLE_CODE" & vbCrLf _
                    '        & ", CARTON_PACK_QTY, CARTON_PACK_QTY CARTON_PACK_QTY_STYLE" & vbCrLf _
                    '        & ", INNER_PACK_QTY, INNER_PACK_QTY INNER_PACK_QTY_STYLE" & vbCrLf _
                    '        & " from POTORDR2 where ROWNUM < 1"
                    '    Dim tblPack As DataTable = ASCDATA1.GetDataTable
                    '    '    For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                    '    '        Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE")
                    '    '        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    '    '        Dim CARTON_PACK_QTY_po As Integer = Val(rowPOTORDR2.Item("CARTON_PACK_QTY") & "")
                    '    '        Dim CARTON_PACK_QTY_style As Integer = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
                    '    '        Dim INNER_PACK_QTY_po As Integer = Val(rowPOTORDR2.Item("INNER_PACK_QTY") & "")
                    '    '        Dim INNER_PACK_QTY_style As Integer = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")

                    '    '        If (CARTON_PACK_QTY_po <> 0 And CARTON_PACK_QTY_po <> CARTON_PACK_QTY_style And CARTON_PACK_QTY_style <> 0) _
                    '    '        Or (INNER_PACK_QTY_po <> INNER_PACK_QTY_style And INNER_PACK_QTY_style <> 0) Then
                    '    '            tblPack.Rows.Add(New Object() {rowPOTORDR2.Item("PO_ORDER_NO"), rowPOTORDR2.Item("PO_ORDER_LNO"), _
                    '    '                                           STYLE_CODE, CARTON_PACK_QTY_po, CARTON_PACK_QTY_style, INNER_PACK_QTY_po, INNER_PACK_QTY_style})
                    '    '        End If
                    '    '    Next

                    '    If tblPack.Rows.Count <> 0 Then
                    '        Using frmmsg As New ASFMSGBF
                    '            frmmsg.Show_grd(tblPack, Me, "Styles on this Shipment with Carton Pack or Inner Pack Qty at odds with Style Table")

                    '            Dim msg_answer As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Do you want to correct the Style Master Qtys from this Shipment", _
                    '                                                                          MsgBoxStyle.YesNoCancel, _
                    '                                                                          "Option to Correct Carton and Inner Pack Qtys in Style Table")
                    '            If msg_answer = MsgBoxResult.Cancel Then
                    '                Exit Sub
                    '            ElseIf msg_answer = MsgBoxResult.Yes Then
                    '                fix_ICTSTYL1_packs = True
                    '            End If
                    '        End Using
                    '    End If
                    'End If

                    Dim EMsg_Cartons As String = Check_Cartons_in_Balance(Not chkFinalize.Checked)

                    If Not chkFinalize.Checked Then
                        If EMsg_Cartons <> "" Then
                            If MsgBox(Mid(EMsg_Cartons, 2) & vbCrLf & vbCrLf & "Continue with Update?",
                                      MsgBoxStyle.YesNo, "Verification to Continue with Update") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If

                    Else
                        EMsg &= EMsg_Cartons


                        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                            If Val(dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'C'").Length) > 0 Then
                                ' this became necessary when we started doing partial receipts based on receiving shortages - don't want to resend the 943
                                ' wjz re-enabled this when i emailed leslie that I changed the send routines to send only in transit containers
                                'EMsg &= vbCr & "You cannot send a shipment to the 3PL which has been partially received"
                            End If
                        End If


                        If EMsg = "" Then
                            If MsgBox("You have elected to Finalize this Shipment by Sending it to the 3PL" _
                                       & vbCrLf _
                                      & vbCrLf & "No Changes are Permitted to the Shipment once it is sent to the 3PL" _
                                      & vbCrLf & " without getting the 3PL to Void this Record in their System" _
                                      & vbCrLf _
                                      & vbCrLf & "OK To Proceed?",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                If ship_entry And EMsg = "" Then

                    Dim PPK_WARNING As String = ""

                    For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
                        Dim ORDR_NO As String = ""
                        Dim PO_ORDER_NOs As New List(Of String)
                        For Each rowPOTSHIP3 As DataRow In rowPOTSHIP2.GetChildRows("POTSHIP2_POTSHIP3")
                            Dim PO_ORDER_NO As String = rowPOTSHIP3.Item("PO_ORDER_NO")
                            If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then PO_ORDER_NOs.Add(PO_ORDER_NO)
                        Next
                        Dim first As Boolean = True
                        For Each PO_ORDER_NO As String In PO_ORDER_NOs
                            Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                            If first Then
                                ORDR_NO = rowPOTORDR1.Item("ORDR_NO") & ""
                            Else
                                If ORDR_NO <> rowPOTORDR1.Item("ORDR_NO") & "" Then
                                    EMsg &= vbCr & "Cannot Mix BTB POs with other BTB POs" _
                                        & vbCr & " or with other non-BTB POs on a single Shipment Line" _
                                        & vbCr & " (see POs listed on Line " & rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & ")"
                                    Exit For
                                End If
                            End If
                            If rowPOTORDR1.Item("PO_HAS_PPK") & "" = "1" Then
                                Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
                                Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) _
                                                     & " and ISNULL(STYLES,0) < 2 and ISNULL(CUSTOM_PPK,'0') <> '1'"
                                If dst.Tables("POTSHIP7").Select(sqlw).Length <> 0 Then
                                    PPK_WARNING &= "," & PO_ORDER_NO
                                End If
                            End If
                            first = False
                        Next
                        rowPOTSHIP2.Item("ORDR_NO") = ORDR_NO
                    Next

                    If PPK_WARNING <> "" Then
                        If MsgBox("Some POs in this shipment have been indicated" _
                                  & vbCrLf & " to have Styles which are Pre-Packed" _
                                  & vbCrLf & " and Some Cartons in this Shipment" _
                                  & vbCrLf & " do not appear to be Pre-Packs." _
                                  & vbCrLf & vbCrLf & "Do you wish to Proceed with this Update?",
                                  MsgBoxStyle.YesNo,
                                  "Verification - Pre-Packs may need to be Specified") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    Dim TBL As New DataTable
                    For Each DC As DataColumn In dst.Tables("POTORDRO").Columns
                        TBL.Columns.Add(DC.ColumnName, DC.DataType)
                    Next
                    TBL.Columns.Add("CONDITION")

                    For Each rowPOTORDRO As DataRow In dst.Tables("POTORDRO").Select("")
                        If rowPOTORDRO.GetChildRows("POTORDRO_POTSHIP3").Length > 1 Then
                            Dim rows_with_non0_qty_shipped As Integer = 0
                            For Each row2 As DataRow In rowPOTORDRO.GetChildRows("POTORDRO_POTSHIP3")
                                If Val(row2.Item("PO_QTY_SHP") & "") <> 0 Then
                                    rows_with_non0_qty_shipped += 1
                                End If
                            Next

                            If rows_with_non0_qty_shipped > 1 Then
                                Dim row As DataRow = TBL.NewRow
                                row.ItemArray = rowPOTORDRO.ItemArray
                                row.Item("CONDITION") = "Multiple Shipment Lines referencing this PO Detail"
                                TBL.Rows.Add(row)
                            End If
                        End If
                    Next
                    For Each rowPOTORDRO As DataRow In dst.Tables("POTORDRO").Select("ISNULL(PO_QTY_SHP,0) > ISNULL(PO_QTY_OPN_PRE,0)")
                        Dim row As DataRow = TBL.NewRow
                        row.ItemArray = rowPOTORDRO.ItemArray
                        row.Item("CONDITION") = "Qty Shipped > Qty on Open PO"
                        TBL.Rows.Add(row)
                    Next
                    If TBL.Rows.Count > 0 Then
                        Using f As New ASFMSGBF
                            f.Show_grd(TBL, Me, "Please acknowledge the following conditions with respect to the PO Details on this Shipment")
                            If f.user_option = -1 Then ' Cancel clicked
                                EMsg &= vbCr & "Returning to Edit Mode"
                            End If
                        End Using
                    End If

                    If ASCMAIN1.CLIENT = "VAN" Then
                        If AT_Packing Then
                            If AT_Packing_Errors <> "" Then
                                EMsg &= vbCr & "Errors in AT Packing" & AT_Packing_Errors
                            Else
                                ' EMsg &= vbCr & "No Updates Permitted for AT Packing (yet) - still testing"
                            End If
                        End If
                    End If
                End If


                If cost_calc And EMsg = "" Then
                    If chkCostComplete.Checked Then
                        If Not cost_ind Then
                            EMsg &= vbCr & "Click Calculate Costs and Verify before Updating"
                        End If

                        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        Else
                            If Val(Absx1.numFor("CUSTOMS_DUTY_AMT").Value & "") <> 0 Then
                                If Absx1.txtFor("CUSTOMS_ENTRY_NO").Text = "" Then
                                    EMsg &= vbCr & "Missing a Customs Entry Number"
                                End If
                            End If

                            Dim TOLERANCE As Decimal = Val(ROWs("POTPARM1").Item("PO_PARM_DUTY_TOLERANCE") & "") ' 0.1

                            ' spreading the duty across containers and styles sometimes results in small OOBALs - defining an arbitrary threshold at .02
                            If System.Math.Abs(Val(numDutyNotDist.Value & "")) > TOLERANCE Then
                                EMsg &= vbCr & "Duty Distribution is Out of Balance (by more than " & Format(TOLERANCE, "#.00") & ")"
                            End If

                            Dim TOTAL_TRUCKING As Decimal = Val(dst.Tables("POTSHIP2").Compute("SUM(TOTAL_TRUCKING)", "") & "")
                            Dim TOTAL_MISC As Decimal = Val(dst.Tables("POTSHIP2").Compute("SUM(TOTAL_MISC)", "") & "")
                            Dim TOTAL_FREIGHT As Decimal = Val(dst.Tables("POTSHIP2").Compute("SUM(TOTAL_FREIGHT)", "") & "")
                            Dim TOTAL_CUSTOMS As Decimal = Val(dst.Tables("POTSHIP2").Compute("SUM(TOTAL_CUSTOMS)", "") & "")
                            Dim LANDING_COST_T As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_T)", "") & "")
                            Dim LANDING_COST_M As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_M)", "") & "")
                            Dim LANDING_COST_F As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_F)", "") & "")
                            Dim LANDING_COST_W As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_W)", "") & "")

                            If System.Math.Abs(TOTAL_TRUCKING - LANDING_COST_T) > TOLERANCE Then
                                EMsg &= vbCr & "Trucking Costs Distribution is Out of Balance (by more than " & Format(TOLERANCE, "#.00") & ")"
                            End If
                            If System.Math.Abs(TOTAL_MISC - LANDING_COST_M) > TOLERANCE Then
                                EMsg &= vbCr & "Misc Costs Distribution is Out of Balance (by more than " & Format(TOLERANCE, "#.00") & ")"
                            End If
                            If System.Math.Abs(TOTAL_FREIGHT - LANDING_COST_F) > TOLERANCE Then
                                EMsg &= vbCr & "Freight Costs Distribution is Out of Balance (by more than " & Format(TOLERANCE, "#.00") & ")"
                            End If
                            If System.Math.Abs(TOTAL_CUSTOMS - LANDING_COST_W) > TOLERANCE Then
                                EMsg &= vbCr & "Customs Costs Distribution is Out of Balance (by more than " & Format(TOLERANCE, "#.00") & ")"
                            End If

                            If EMsg = "" Then
                                If automated_cost_complete Then
                                Else
                                    If MsgBox("By Indicating the Costing is Complete," _
                                      & vbCrLf & " you are locking in the Landed Cost Values for Styles in this Shipment" _
                                      & vbCrLf & " and Journal Entries will be created based on these values." _
                                      & vbCrLf & vbCrLf & "This process is not reversible." _
                                      & vbCrLf & vbCrLf & "OK to Continue with this Update?",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                    End If
                End If

            Case "Delete"
                ASCMAIN1.sql = "Select * from ICTTRAN1 where TRAN_TYPE_ORIG = 'P' and TRAN_NO_ORIG = '" & PO_SHIPMENT_NO & "'"
                If ASCDATA1.GetDataRow IsNot Nothing Then
                    EMsg &= vbCr & "You May Not Delete a Shipment which has been Received"
                End If

                If EMsg = "" Then
                    ASCMAIN1.sql = "Select * from POTSHIP2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
                    & " AND PO_SHIP_STATUS = 'C'"
                    ' NEED TO CHECK POTSHIP2 TO PROTECT AGAINST DELETING SHIPMENTS CONVERTED AS RECEIVED
                    If ASCDATA1.GetDataRow IsNot Nothing Then
                        EMsg &= vbCr & "You May Not Delete a Shipment which has been Received"
                    End If
                End If

                If EMsg = "" Then
                    If MsgBox("THIS WILL PERMANANTLY DELETE THIS SHIPMENT!!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "WARNING!") = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If EntryMode = "X" Then
                    ' RECEIVING DISCREPANCIES - JUST CANCEL
                Else
                    If automated_cost_complete Or loading_AT Or packingFromXLS Or packingFromBooking Then
                    Else
                        If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                            "You may have made Changes") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Receive this BOL"
                If Absx1.dteFor("PO_DATE_RECEIVED").Value & "" = "" Then
                    EMsg &= vbCr & "You Must 1st Specify A Receipt Date"
                End If

                If grdPOTSHIP2.ActiveRow Is Nothing Or Not grdPOTSHIP2.ActiveRow.IsDataRow Then
                    EMsg &= vbCr & "Cannot Determine BOL to Receive - Select a Row from the BOL Grid"
                Else
                    Dim PO_SHIPMENT_LNO As Integer = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                    Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
                    Dim PO_QTY_SHP As Integer = Val(dst.Tables("POTSHIP3").Compute("Sum(PO_QTY_SHP)", sqlw) & "")
                    If PO_QTY_SHP = 0 Then
                        EMsg &= vbCr & "There Are No Receiving Details On This Container"
                    End If
                End If

            Case "Get 1st Cost from PO"
                If MsgBox("Do you really want to Get Fresh Values for 1st Cost from the POs?",
                           MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Verification.") = MsgBoxResult.Cancel Then
                    Exit Sub
                End If

            Case "Import Bookings"

                If grdPOTVBKGX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You must Select 1 or more Bookings to combine into a Shipment"
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grdPOTVBKGX.Selected.Rows
                        Dim VBKG_NO As String = grow.Cells("VBKG_NO").Value
                        If Not ASCMAIN1.Logical_Lock("POTVBKG1", VBKG_NO) Then
                            Exit Sub
                        End If
                        Dim rowPOTVBKG1 As DataRow = LookUp("POTVBKG1", VBKG_NO)
                        If rowPOTVBKG1.Item("VBKG_STATUS") & "" <> "F" Then
                            EMsg &= vbCr & $"Booking {VBKG_NO} is not Finalized"
                        End If
                        If rowPOTVBKG1.Item("PO_SHIPMENT_NO") & "" <> "" Then
                            EMsg &= vbCr & $"Booking {VBKG_NO} has alread been imported into Shipment {rowPOTVBKG1.Item("PO_SHIPMENT_NO")}"
                        End If
                    Next
                End If

                If EMsg = "" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdPOTVBKGX.Selected.Rows
                        Dim VBKG_NO As String = grow.Cells("VBKG_NO").Value & ""
                        ' get all records in BKG2 file
                        ASCMAIN1.sql = $"SELECT * FROM POTVBKG2 WHERE VBKG_NO = '{VBKG_NO}'"
                        Dim TBL As DataTable = ASCDATA1.GetDataTable()
                        For Each row As DataRow In TBL.Select()
                            Dim rowPOTPACK1 As DataRow = LookUp("POTPACK1", row.Item("PACK_LIST_NO"))
                            Dim PO_ORDER_NO As String = rowPOTPACK1.Item("PO_ORDER_NO")
                            packingListPOs.Add(PO_ORDER_NO)
                            Dim PO_ORDER_NO2 As String = rowPOTPACK1.Item("PO_ORDER_NO2") & ""
                            If PO_ORDER_NO2 & "" <> "" Then
                                packingListPOs.Add(PO_ORDER_NO2)
                            End If
                            Dim PO_ORDER_NO3 As String = rowPOTPACK1.Item("PO_ORDER_NO3") & ""
                            If PO_ORDER_NO3 & "" <> "" Then
                                packingListPOs.Add(PO_ORDER_NO3)
                            End If
                        Next
                        'Book2ShiP(VBKG_NO, PO_SHIPMENT_NO)
                    Next
                    Got_PO_Locks(EMsg)


                    If MsgBox($"You have selected {grdPOTVBKGX.Selected.Rows.Count} Bookings to be imported into a Shipment." & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                Else
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "New Packing Slip"
                If grdPOTPACKG.Selected.Rows.Count < 1 Then
                    EMsg = "Please select an Open Shipment for New Packing Slip"
                Else
                    Dim grow As UltraWinGrid.UltraGridRow = grdPOTPACKG.Selected.Rows(0)
                    If grow.Band.Key <> "POTPACKG" Then
                        EMsg = "Please select one or more Packing Slip rows, not the styles"
                    Else
                        For Each grow In grdPOTPACKG.Selected.Rows
                            If grow.Cells("PO_QTY_BAL").Value = 0 Then
                                EMsg = "This Shipment record has been completely allocated"
                                Exit For
                            End If
                            Dim PO_SHIPMENT_NO As String = grow.Cells("PO_SHIPMENT_NO").Value & ""
                            Dim PO_ORDER_NO As String = grow.Cells("PO_ORDER_NO").Value & ""
                            Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value & ""
                            Dim PO_SHIPMENT_NOs As String = ""
                            Dim PO_ORDER_NOs As String = ""
                            Dim ORDR_NOs As String = ""
                            If Not PO_SHIPMENT_NOs.Contains(PO_SHIPMENT_NO) Then
                                If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then Exit Sub
                                PO_SHIPMENT_NOs &= "," & PO_SHIPMENT_NO
                            End If
                            If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then
                                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                                PO_ORDER_NOs &= "," & PO_ORDER_NO
                            End If
                            If Not ORDR_NOs.Contains(ORDR_NO) Then
                                If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub
                                ORDR_NOs &= "," & ORDR_NO
                            End If
                            For Each grw2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                                PO_ORDER_NO = grw2.Cells("PO_ORDER_NO").Value & ""
                                If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then
                                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                                    PO_ORDER_NOs &= "," & PO_ORDER_NO
                                End If
                            Next
                        Next
                    End If
                End If

            Case "Edit Packing Slip"
                If grdPOTPCKS1.Selected.Rows.Count <> 1 Then
                    EMsg = "Please select one Packing Slip to Edit"
                Else
                    Dim grow As UltraWinGrid.UltraGridRow = grdPOTPCKS1.Selected.Rows(0)
                    If grow.Band.Key <> "POTPCKS1" Then
                        EMsg = "Please select one or more Packing Slip rows, not the styles"
                    Else
                        For Each grow In grdPOTPCKS1.Selected.Rows
                            For Each grw2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                                Dim PO_SHIPMENT_NO As String = grw2.Cells("PO_SHIPMENT_NO").Value & ""
                                Dim PO_ORDER_NO As String = grw2.Cells("PO_ORDER_NO").Value & ""
                                Dim ORDR_NO As String = grw2.Cells("ORDR_NO").Value & ""
                                Dim PO_SHIPMENT_NOs As String = ""
                                Dim PO_ORDER_NOs As String = ""
                                Dim ORDR_NOs As String = ""
                                If Not PO_SHIPMENT_NOs.Contains(PO_SHIPMENT_NO) Then
                                    If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then Exit Sub
                                    PO_SHIPMENT_NOs &= "," & PO_SHIPMENT_NO
                                End If
                                If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then
                                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                                    PO_ORDER_NOs &= "," & PO_ORDER_NO
                                End If
                                If Not ORDR_NOs.Contains(ORDR_NO) Then
                                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub
                                    ORDR_NOs &= "," & ORDR_NO
                                End If
                            Next
                        Next
                    End If
                End If


            Case "Update Packing Slip"
                Dim QtyErr As Integer = dst.Tables("POTPCKS2").Compute("count(IN_ERR)", "IN_ERR = '1'")
                If QtyErr > 0 Then
                    If MsgBox("Do you wish to Update lines with quantites not in case packs?",
                           MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Verification.") = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                End If
                'Case "Import Bookings"
                '    ' CREATE PACKING LIST POS FOR THE APPRORIATE DICTIONARY
                '    For Each grow As UltraWinGrid.UltraGridRow In grdPOTVBKGX.Selected.Rows
                '        Dim VBKG_NO As String = grow.Cells("VBKG_NO").Value & ""
                '        ' get all records in BKG2 file
                '        ASCMAIN1.sql = $"SELECT * FROM POTVBKG2 WHERE VBKG_NO = '{VBKG_NO}'"
                '        Dim TBL As DataTable = ASCDATA1.GetDataTable()
                '        For Each row As DataRow In TBL.Select()
                '            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                '            packingListPOs.Add(PO_ORDER_NO)
                '        Next
                '        'Book2ShiP(VBKG_NO, PO_SHIPMENT_NO)
                '    Next
                '    Got_PO_Locks(EMsg)

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

            Case "Edit", "Select"

                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Delete"
                EntryMode = "D"
                Delete_Record()
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
                If automated_cost_complete Then
                Else
                    TAC.POCMAIN1.Check_Status(Me)
                End If


                ASCMAIN1.sql = "Select X.*" & vbCrLf _
                    & ", POTORDR1.PO_STATUS, POTORDR1.PO_DATE_ETA, POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE" & vbCrLf _
                    & " from POTORDR1, (" & vbCrLf _
                    & "Select PO_ORDER_NO, PO_STATUS, COUNT (*) NEGLINES, SUM (PO_QTY_SHP) SHP" & vbCrLf _
                    & " from POTORDR2 WHERE PO_QTY_SHP < 0" & vbCrLf _
                    & " group by PO_ORDER_NO, PO_STATUS" & vbCrLf _
                    & ") X where POTORDR1.PO_ORDER_NO = X.PO_ORDER_NO"

                Dim tbl As DataTable = ASCDATA1.GetDataTable
                If tbl.Rows.Count <> 0 Then
                    Using frm As New ASFMSGBF
                        frm.Show_grd(tbl, Me, "There are POs With Negative Qtys Shipped - Please Take Screenshot And email To ABS")
                    End Using
                End If


            Case "Cancel", "Done"
                If UltraExplorerBar1.Groups("Packing Slips").Visible = True Then
                    PackSlipCancel()
                End If
                Mode_Settings(False)

            Case "Get Duty"
                Get_Duty()

            Case "Get Weight Factor"
                Get_Weight_Factor()

            Case "Get 1st Cost from PO"
                Get_1st_Cost_from_PO()

            Case "Receive this BOL"
                Receive_BOL()

            Case "Variance"
                Calc_Cost_Variance()

            Case "Calculate Costs"
                Calculate_Landed_Cost()

            Case "Print"
                Print_Record()

            Case "Show Invoice"
                Show_Invoice()

            Case "email Invoice"
                EmailInvoice()

            Case "Import Bookings"
                Import_Bookings()
                Mode_Settings(True)

            Case "New Packing Slip"
                EntryMode = "N"
                CreatePackingSlip()
                Set_ScreenMode_Base(True)

            Case "Edit Packing Slip"
                EntryMode = "E"
                EditPackingList()
                Set_ScreenMode_Base(True)

            Case "View Packing Slip"
                EntryMode = "V"
                EditPackingList()
                Set_ScreenMode_Base(False)
                EntryMode = "X"

            Case "Print Packing Slip"
                Print_PackingSlip(False)

            Case "Email Packing Slip"
                Print_PackingSlip(True)

                'Click_Command("View")
            Case "Update Packing Slip"
                PackSlipUpdt()
                Set_ScreenMode_Base(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Select").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                        '.Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Show Invoice").Visible = ScreenMode And receipt_mode AndAlso (optReceiptType.Value = "BTB")
                    .Items("email Invoice").Visible = ScreenMode And receipt_mode AndAlso (optReceiptType.Value = "BTB")
                    .Items("Show Invoice").Visible = False
                    .Items("email Invoice").Visible = False

                    .Items("View").Visible = (EntryMode = "V" Or Not ScreenMode) And Not receipt_mode
                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                    .Items("Print").Visible = ScreenMode And (EntryMode <> "N")

                    .Items("New").Visible = (EntryMode <> "V" Or Not ScreenMode) And Not receipt_mode And Not InquiryMode And Not cost_calc
                    '.Items("Edit").Visible = (EntryMode <> "V" Or Not ScreenMode)

                    .Items("Edit").Visible = Not (MENU_ITEM_OBJECT = "POFSHIPR") And Not InquiryMode
                    .Items("Update").Visible = (Not (EntryMode = "V") And ScreenMode)
                    .Items("Delete").Visible = (EntryMode = "E" And ScreenMode) And Not cost_calc And Not receipt_mode
                    If ship_entry And EntryMode = "E" AndAlso rowPOTSHIP1.Item("LP_STATUS") & "" = "1" Then
                        .Items("Delete").Visible = False
                    End If
                    .Items("Cancel").Visible = (Not (EntryMode = "V") And ScreenMode)

                    .Items("Variance").Visible = EntryMode = "E" And cost_calc And ScreenMode And (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
                    .Items("Calculate Costs").Visible = EntryMode = "E" And cost_calc And ScreenMode

                    .Items("Import Bookings").Visible = False

                End With
                .Groups("BOL Data").Visible = ScreenMode And Not ship_entry And Not receipt_mode
                .Groups("Customs/Duty").Visible = ScreenMode And Not ship_entry And Not receipt_mode
                .Groups("Receipts").Visible = ScreenMode And receipt_mode
                If Not ScreenMode Then
                    .Groups("Back-to-Back").Visible = False ' ScreenMode And receipt_mode AndAlso (rowPOTSHIP2.Item("ORDR_NO") & "" <> "")
                End If
                .Groups("Receipt Type").Visible = Not ScreenMode And receipt_mode
                .Groups("Packing Slips").Visible = False
            End With

            With UltraExplorerBar1.Groups("Cost Options")
                .Items("Get Duty").Visible = (Not (EntryMode = "V") And ScreenMode) And cost_calc
                .Items("Get Weight Factor").Visible = (Not (EntryMode = "V") And ScreenMode) And cost_calc
            End With
        End If

        btnABSonly.Visible = ASCMAIN1.Running_in_VS And Not ScreenMode

        chkNoDuty.Visible = ScreenMode And cost_calc
        ' And _
        '((ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") Or _
        ' (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA"))

        chkFixCasePacks.Visible = ScreenMode And ship_entry _
            And (EntryMode = "N" Or EntryMode = "E") _
            And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Set_Read_Only(grpHeaderData, ScreenMode And (EntryMode = "V" Or cost_calc))
        Set_Read_Only(grpShipment, ScreenMode And (EntryMode = "V"))
        Set_Read_Only(grpCustomsDuty, ScreenMode And (EntryMode = "V"))
        If cost_calc Or InquiryMode Then
            Set_Read_Only_for_ctl(optUD, False)
        End If
        Set_Read_Only(grpPOTSHIP1, Not ScreenMode Or (EntryMode = "V") Or receipt_mode Or cost_calc)
        Set_Read_Only_for_ctl(Absx1.txtFor("PO_SHIPMENT_NO"), ScreenMode)

        If cost_calc And ScreenMode And EntryMode = "E" Then
            Set_Read_Only_for_ctl(chkNoDuty, False)
            Set_Read_Only_for_ctl(txtPO_NOTES, False)
        End If
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else
            Set_Read_Only_for_ctl(Absx1.optFor("FREIGHT_ENTERED_BY"), True)
        End If
        cmdDutyBalance.Visible = cost_calc And (EntryMode = "E")
        SETUP_tabPOTSHIP1()


        'grdPOTSHIPX.Visible = Not tf
        tab0.Visible = Not tf
        Setup_tabBOL()
        If ScreenMode Then

            Toggle_chkFinalize()


            If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), True)
            End If

            'If receipt_mode Then
            '    With grdPOTSHIP2.DisplayLayout.Bands(0)
            '        .Columns("ACTION").Hidden = select_from_3PL_list
            '    End With
            'End If


            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdPOTSHIP2, grdPOTSHIP3, grdPOTSHIP4, grdPOTSHIP5, grdPOTSHIP7, grdPOTSHIP8, grdSOTORDP1}
                With grd.DisplayLayout.Override
                    If ScreenMode And receipt_mode And grd.Name = "grdPOTSHIP2" Then
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    Else
                        If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                            If cost_calc And (grd.Name = "grdPOTSHIP7" Or grd.Name = "grdPOTSHIP8") Then
                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                .AllowUpdate = DefaultableBoolean.False
                                .AllowDelete = DefaultableBoolean.False
                                If grd.Name = "grdPOTSHIP7" Then
                                    .AllowUpdate = DefaultableBoolean.True
                                    For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                                        If gcol.Key = "CARTON_DIMS" Then
                                            gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                                            gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                                        Else
                                            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                                        End If
                                    Next
                                End If
                            Else
                                If cost_calc And (grd.Name = "grdPOTSHIP2" Or grd.Name = "grdPOTSHIP4") Then
                                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                Else
                                    If grd.Name = "grdPOTSHIP7" _
                                    Or grd.Name = "grdPOTSHIP8" _
                                    Or grd.Name = "grdPOTSHIP3" _
                                    Or grd.Name = "grdSOTORDP1" Then
                                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                    Else
                                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                                    End If
                                End If
                                .AllowUpdate = DefaultableBoolean.True
                                .AllowDelete = DefaultableBoolean.True
                            End If
                            Set_Entry_Mode_Controls(False)
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowUpdate = DefaultableBoolean.False
                            .AllowDelete = DefaultableBoolean.False
                        End If
                    End If
                End With
            Next

            'If (EntryMode = “E” Or EntryMode = “N”) And ASCMAIN1.CLIENT = "VAN" And ship_entry And POTVBKG2_RECORDS Then
            '    grdPOTSHIP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            '    grdPOTSHIP3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            '    grdPOTSHIP4.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            'Else
            '    grdPOTSHIP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            '    grdPOTSHIP3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            '    grdPOTSHIP4.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            'End If

            '  Setup_tabBOL()
            Toggle_Columns()

            If select_from_3PL_list Then
                grdWHT3PLR1.Parent = spl3PL.Panel1
                grdEDT944T1.Parent = spl3PL.Panel1
                tabBOL.Tabs("3PL Raw Data").Visible = True
            End If
        Else
            Clear_Record()
            grdWHT3PLR1.Parent = tab0.Tabs("3PL Receipts").TabPage
            grdEDT944T1.Parent = tab0.Tabs("3PL Receipts").TabPage
            tabBOL.Tabs("3PL Raw Data").Visible = False
            Setup_tab0()
            Set_Landed_Cost_Needs_to_be_Calculated_Indicator(True)
            chkFinalize.Visible = False
        End If
        '   Setup_tab0()

        If ship_entry And EntryMode = "E" AndAlso rowPOTSHIP1.Item("LP_STATUS") & "" = "1" Then
            ' CHG MADE TO A SHIPMENT AFTER TRANSMIT - LOCKING THINGS DOWN UNTIL I KNOW THAT THEY ARE NECESSARY
            grdPOTSHIP2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdPOTSHIP4.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Set_Read_Only(grpHeaderData, True)
        End If

        SETUP_tabPOTSHIP1()

        tabBOL.Tabs("BTB Invoices").Visible = False
        If receipt_mode And ScreenMode Then
            If dst.Tables("POTSHIP2").Select("ORDR_NO Is Not NULL").Length > 0 And (WHSE_TYPE = "P" Or (ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "NC")) Then
                tabBOL.Tabs("BTB Invoices").Visible = True
            End If
        End If

        'If T = False Then
        '    SSDateCombo1(3).Enabled = True
        '    cmdSelPO.Visible = False
        '    Call Set_SHIP1(2, True, True)
        'End If

        'frmReceipts.Visible = receipt_mode And T

        If receipt_mode Then
            grdSOTORDP1.DisplayLayout.Bands(0).Columns("INV_NO").Header.Caption = "Order No"
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        dst.Tables("POTSHIP3").Columns("CBM").Expression = ""

        For Each TABLE_NAME As String In New String() _
                {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP5", "POTSHIP7", "POTSHIP8", "POTSHIPR", "POTSHIPQ",
                 "POTORDR2", "POTORDR1_SPLIT", "POTORDR2_SPLIT", "POTSHPXL", "SOTORDP1", "SOTORDP2", "APTINVH1",
                 "SOTINVH1", "SOTINVH2", "ARTOPEN1", "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTORDR1", "SOTORDR2",
                 "POTSHPWB", "POTSHPIE", "POTPACKR", "POTPACKD"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each TABLE_NAME As String In New String() _
        {"WHTPPKM1", "WHTPPKM2", "ICTTRAN1", "ICTTRAN2", "WHTWREC7", "WHTWREC8"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        End If
        EnforceConstraints(True)

        ' NEW 3/25/23 DGJ
        POTORDR1_added.Clear()

        STYLE_CODEs_No_Duty.Clear()
        select_from_3PL_list = False
        Select_from_Whse_Receipt = False
        chkFinalize.Checked = False
        Absx1.txtFor("PO_SHIPMENT_NO").Focus()
        lblMissingCBMs.Visible = False

        chkFixCasePacks.Checked = False
        chkFixCasePacks.Visible = False
        packingFromXLS = False
        packingFromBooking = False
        dicWorksheetPOs.Clear()
        dicWorkbooks.Clear()

        '   fix_ICTSTYL1_packs = False
        WH_REC_NOsInProcess.Clear()


        splCartonQ.Panel2Collapsed = True

        CARTON_DIMS_by_PO.Clear()

        If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Then
            Load_WHTWRECX()
        End If

        Load_POTSHIPX()

        If ASCMAIN1.CLIENT = "RGI" Then
            Fill_Records("POTSHIPS")
        End If

        If cost_calc Then
            If tab0.SelectedTab.Key = "Costing Summaries" Then
                Load_POTSHIPI()
            End If
        End If
        If InquiryMode Then
            If grdICTIRECX.Rows.Count = 0 Then Load_ICTIRECX()
        End If

        If (tab0.SelectedTab.Key = "AT Shipments") Then
            Setup_AT_Shipments()
        End If

        AT_Packing = False
        AT_Packing_Errors = ""

        If ASCMAIN1.CLIENT = "VAN" Then
            dicPOTORDR1.Clear()
            dicPOTORDR2.Clear()
        End If

        tabBOL.Tabs("Import Errors").Visible = False
        tabBOL.Tabs("Packing Discrepancies").Visible = False


        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = sqlPOTVBKGX & " And VBKG_STATUS = 'F' and PO_SHIPMENT_NO is Null"
            Fill_Records("POTVBKGX", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdPOTVBKGX, "VBKG_NO".ToLower)
        End If
        POTVBKG2_RECORDS = False

        WORKBOOK_COUNTER = 0
        QTY_PACKED.Clear()
    End Sub

    Sub Set_Entry_Mode_Controls(importOnly As Boolean)
        If importOnly Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTSHIP2}
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False
                End With
                If grd.Name = "grdPOTSHIP2" Then
                    grd.DisplayLayout.Bands(0).Columns("ACTION").Hidden = True

                    'For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                    '    If gcol.Key = "COMM_INV_NO" Or gcol.Key = "BOL_NO" Or gcol.Key = "CONTAINER_NO" Then
                    '        gcol.CellActivation = Activation.AllowEdit
                    '    Else
                    '        gcol.CellActivation = Activation.NoEdit
                    '    End If
                    'Next
                End If
            Next
        Else
            grdPOTSHIP2.DisplayLayout.Bands(0).Columns("ACTION").Hidden = False
        End If
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Shipment")

        Save_Header_Fields(UltraGroupBox1)

        Dim chk_cost_ind As String
        '  TAC.POCMAIN1.POSHIPCHK()

        If EntryMode = "N" Then
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                PO_SHIPMENT_NO = ASCMAIN1.Next_Control_No("PO_SHIPMENT_NO")
            Else
                PO_SHIPMENT_NO = ASCMAIN1.Next_Control_No("POTSHIP1.PO_SHIPMENT_NO")
            End If
        Else
            PO_SHIPMENT_NO = Absx1.txtFor("PO_SHIPMENT_NO").Text
        End If

        CARTON_DIMS_by_PO.Clear()

        EnforceConstraints(False)

        'Dim rowPOTSHIP1 As DataRow
        If EntryMode = "N" Then
            rowPOTSHIP1 = dst.Tables("POTSHIP1").NewRow
            rowPOTSHIP1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP1.Item("PO_SHIP_LANDING_LEAD_DAYS") = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR") & "")
            rowPOTSHIP1.Item("WHSE_CODE") = ROWs("POTPARM1").Item("PO_PARM_DEF_WHSE_CODE") & ""
            rowPOTSHIP1.Item("REVIEW") = 0

            rowPOTSHIP1.Item("FREIGHT_ENTERED_BY") = ROWs("POTPARM1").Item("PO_PARM_FREIGHT_ENTERED_BY")
            rowPOTSHIP1.Item("COST_FRT_METHOD") = ROWs("POTPARM1").Item("PO_PARM_COST_FRT_METHOD")

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                rowPOTSHIP1.Item("PO_SHIP_REF_NO") = Mid(PO_SHIPMENT_NO, 2, 5)
            End If
            rowPOTSHIP1.Item("PO_SHIP_ADV_DATE") = DATETIME_STAMP.Date
            dst.Tables("POTSHIP1").Rows.Add(rowPOTSHIP1)
        Else
            rowPOTSHIP1 = Fill_Record("POTSHIP1", PO_SHIPMENT_NO)
        End If

        If EntryMode = "E" And ASCMAIN1.CLIENT = "VAN" And ship_entry Then
            ' FILL RECORDS WITH SQL STATMENT
            ASCMAIN1.sql = "SELECT * FROM POTVBKG2 WHERE PO_SHIPMENT_NO ='" & PO_SHIPMENT_NO & "'"
            Fill_Records("POTVBKG2", "", True, ASCMAIN1.sql)
            If dst.Tables(“POTVBKG2”).Rows.Count > 0 Then
                POTVBKG2_RECORDS = True
            End If


        End If


        dst.Tables("POTSHIP3").Columns("CBM").Expression = ""

        Fill_Records("POTSHIP2", PO_SHIPMENT_NO)
        Fill_Records("POTSHIP3", New Object() {PO_SHIPMENT_NO})
        If ASCMAIN1.CLIENT = "RGI" And InquiryMode Then
            Fill_Records("WHTPREC3", New Object() {PO_SHIPMENT_NO})
            LoadWhseReceipts()
        End If

        Fill_Records("POTSHIP4", PO_SHIPMENT_NO)

        If receipt_mode And (ASCMAIN1.CLIENT = "NYA") Then
            For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
                Dim CONTAINER_NO As String = rowPOTSHIP2.Item("CONTAINER_NO") & ""
                Dim rowPOTSHIP4s() As DataRow = dst.Tables("POTSHIP4").Select("CONTAINER_NO = '" & CONTAINER_NO & "'")

                If rowPOTSHIP4s.Length > 0 Then
                    rowPOTSHIP2.Item("CONTAINER_SEAL_NO") = rowPOTSHIP4s(0).Item("CONTAINER_SEAL_NO")
                End If
            Next

        End If

        Fill_Records("POTSHIP5", PO_SHIPMENT_NO)

        ASCMAIN1.sql = "" _
            & "Select POTLCST1.* from POTLCST1,APTINVH1" & vbCrLf _
            & " where POTLCST1.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
            & "   and APTINVH1.VOUCHER_NO = POTLCST1.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.INV_STATUS <> 'D'" & vbCrLf _
            & " union " & vbCrLf _
            & "Select POTLCST1.* from POTLCST1,APTINVH1" & vbCrLf _
            & " where POTLCST1.PO_ORDER_NO in (Select Distinct PO_ORDER_NO from POTSHIP3" & vbCrLf _
            & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')" & vbCrLf _
            & "   and APTINVH1.VOUCHER_NO = POTLCST1.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.INV_STATUS <> 'D'"

        For Each rowPOTLCST1 As DataRow In ASCDATA1.GetDataTable.Select("")

            Dim CTL_NO As String = rowPOTLCST1.Item("CTL_NO")
            Dim rows() As DataRow = dst.Tables("POTSHIP5").Select("CTL_NO = '" & CTL_NO & "'")

            Dim rowPOTSHIP5 As DataRow

            If rows.Length <> 0 Then
                rowPOTSHIP5 = rows(0)

                With rowPOTSHIP5
                    .Item("LANDING_COST_COMMENT") = "Vendor Invoice"
                    .Item("VOUCHER_NO") = rowPOTLCST1.Item("VOUCHER_NO")
                    .Item("VEND_CODE") = rowPOTLCST1.Item("VEND_CODE")
                    .Item("PO_SHIPMENT_LNO_DIST") = rowPOTLCST1.Item("PO_SHIPMENT_LNO")
                End With
            Else
                If cost_calc And rowPOTSHIP1.Item("COST_COMPLETE") & "" <> "1" Then

                    Dim COST_FACTOR As Decimal = 1
                    If rowPOTLCST1.Item("PO_SHIPMENT_NO") & "" = "" Then
                        Dim PO_ORDER_NO As String = rowPOTLCST1.Item("PO_ORDER_NO") ' THERE BETTER BE A PO IF THERE IS NO SHIPMENT
                        ASCMAIN1.sql = "" _
                            & "Select CASE WHEN SHP = 0 THEN 0 ELSE TRUNC(100 * 100 * SHP/ORD) / 100  END PCT from (" & vbCrLf _
                            & "Select POTSHIP3.PO_ORDER_NO" & vbCrLf _
                            & ", SUM (POTORDR2.PO_COST * POTORDR2.PO_QTY_ORD) ORD" & vbCrLf _
                            & ", SUM (POTORDR2.PO_COST * POTSHIP3.PO_QTY_SHP) SHP" & vbCrLf _
                            & " from POTSHIP3,POTORDR2" & vbCrLf _
                            & " where POTSHIP3.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and POTSHIP3.PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                            & "AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                            & "AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                            & "group by POTSHIP3.PO_ORDER_NO)"
                        COST_FACTOR = Val(ASCDATA1.GetDataValue) / 100
                    End If

                    rowPOTSHIP5 = dst.Tables("POTSHIP5").NewRow
                    With rowPOTSHIP5
                        .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        Dim PO_SHIPMENT_LNO As Integer = Val(dst.Tables("POTSHIP5").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
                        .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO

                        Dim rowPOTCATG1 As DataRow = LookUp("POTCATG1", rowPOTLCST1.Item("COST_CATGY_CODE"))
                        .Item("COST_CATGY_DESC") = rowPOTCATG1.Item("COST_CATGY_DESC")
                        .Item("LANDING_COST_AMT") = Val(rowPOTLCST1.Item("COST_ACT") & "") * COST_FACTOR
                        .Item("LANDING_COST_DIST") = rowPOTCATG1.Item("LANDING_COST_DIST")
                        .Item("LANDING_COST_COMMENT") = "Vendor Invoice"
                        .Item("CTL_NO") = rowPOTLCST1.Item("CTL_NO")
                        .Item("VOUCHER_NO") = rowPOTLCST1.Item("VOUCHER_NO")
                        .Item("VEND_CODE") = rowPOTLCST1.Item("VEND_CODE")
                        .Item("PO_SHIPMENT_LNO_DIST") = rowPOTLCST1.Item("PO_SHIPMENT_LNO")

                    End With
                    dst.Tables("POTSHIP5").Rows.Add(rowPOTSHIP5)
                End If
            End If
        Next


        For Each row As DataRow In dst.Tables("POTSHIP5").Select("ISNULL(CTL_NO,'')<> ''")
            Dim CTL_NO As String = row.Item("CTL_NO")
            ASCMAIN1.sql = "Select Sum (COST_ACT_PO) from POTLCST2 where CTL_NO = '" & CTL_NO & "' and CHARGEBACK_STATUS = '2'"
            Dim COST_ACT_PO As Decimal = Val(ASCDATA1.GetDataValue)
            row.Item("CHARGEBACK_AMT") = COST_ACT_PO
        Next


        If cost_calc And EntryMode = "E" Then
            ASCDATA1.DeleteRows("POTSHIP5", "CTL_NO IS NOT NULL AND VOUCHER_NO IS NULL")
        End If

        Fill_Records("POTSHIP7", PO_SHIPMENT_NO)
        Fill_Records("POTSHIP8", PO_SHIPMENT_NO)

        dst.Tables("POTSHIPR").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3"), New String() {"PO_SHIPMENT_LNO", "STYLE_CODE", "COLOR_CODE"}).Rows
            Create_POTSHIPR(row.Item("PO_SHIPMENT_LNO"), row.Item("STYLE_CODE"), row.Item("COLOR_CODE"))
        Next
        Sort_grdColumns(grdPOTSHIPR, "PO_SHIPMENT_LNO,STYLE_CODE,COLOR_CODE")



        Fill_Records("POTORDRO", PO_SHIPMENT_NO)
        For Each rowPOTORDRO As DataRow In dst.Tables("POTORDRO").Select("")
            Dim PO_ORDER_NO As String = rowPOTORDRO.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Int32 = rowPOTORDRO.Item("PO_ORDER_LNO")
            'ASCMAIN1.sql = "Select Sum (PO_QTY_SHP) from POTSHIP3" _
            '    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO) _
            '    & " and "
            'Dim PO_QTY_SHP_PRE As Int64 = ASCDATA1.GetDataValue
            Dim sqlw As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            Dim PO_QTY_SHP_PRE As Int64 = Val(dst.Tables("POTSHIP3").Compute("SUM(PO_QTY_SHP)", sqlw) & "")
            If EntryMode = "N" Or EntryMode = "E" Or EntryMode = "V" Then
                rowPOTORDRO.Item("PO_QTY_OPN_PRE") = Val(rowPOTORDRO.Item("PO_QTY_OPN") & "") + PO_QTY_SHP_PRE
            Else
                rowPOTORDRO.Item("PO_QTY_OPN_PRE") = Val(rowPOTORDRO.Item("PO_QTY_OPN") & "")
            End If

        Next

        dst.Tables("POTSHIP3").Columns("CBM").Expression = "IIF(ISNULL(PARENT(POTSHIPR_POTSHIP3).QTY_SHP,0) = 0, 0, ISNULL(PARENT(POTSHIPR_POTSHIP3).CBM,0) * ISNULL(PO_QTY_SHP,0) / ISNULL(PARENT(POTSHIPR_POTSHIP3).QTY_SHP,0))"

        ' THIS BLOCK PROBABLY ONLY MAKES SENSE WHEN WE DEAL WITH RECEIPTS INQ - IF WHSE_TYPE = 'P'
        'If ASCMAIN1.CLIENT = "RGI" And WHSE_TYPE = "P" Then
        '    For Each row As DataRow In dst.Tables("POTSHIP2").Select("ORDR_NO IS NOT NULL")
        '        Dim ORDR_NO As String = row.Item("ORDR_NO")
        '        Fill_Records("SOTORDP1", ORDR_NO)
        '        Fill_Records("SOTORDP2", ORDR_NO)
        '    Next
        'End If
        'If receipt_mode Then
        ASCMAIN1.sql = "Select POTORDR1.* from POTORDR1 where PO_ORDER_NO in " _
            & " (Select Distinct PO_ORDER_NO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')"
        Fill_Records("POTORDR1", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select POTORDR2.* from POTORDR2 where (PO_ORDER_NO, PO_ORDER_LNO) in " _
            & " (Select Distinct PO_ORDER_NO, PO_ORDER_LNO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')"
        Fill_Records("POTORDR2", "", True, ASCMAIN1.sql)
        'End If

        EnforceConstraints(True)


        Sort_grdColumns(grdPOTSHIP2, "PO_SHIPMENT_LNO")
        ' NEED TO RETHINKG THRU THE EVENT LOGIC FROM THIS POINT DOWN

        Set_Container()

        If cost_calc Then
            chk_cost_ind = rowPOTSHIP1.Item("COST_IND") & ""
            If chk_cost_ind = "1" Then
                Set_Landed_Cost_Needs_to_be_Calculated_Indicator(True)

            Else
                If rowPOTSHIP1.Item("COST_COMPLETE") & "" <> "1" Then
                    Get_Duty()
                    Get_Weight_Factor()
                End If
            End If
        End If

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Setup_Warehouse_Attributes()

        Sort_grdColumns(grdPOTSHIP2, "PO_SHIPMENT_LNO")
        Setup_grdPOTSHIP2_ActiveRow()
        Toggle_UD()
        PPK_CODE_ctr = 0

        'If rowPOTSHIP1.Item("ORDR_NO") & "" <> "" Then AND WHSE_TYPE = 'P'
        '    Dim ORDR_NO As String = rowPOTSHIP1.Item("ORDR_NO") & ""
        '    Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO)
        'End If

        If automated_cost_complete Then
        Else
            Verify_Integrity()
        End If

        If InquiryMode Then
            Fill_Records("APTINVH1", PO_SHIPMENT_NO)
            Sort_grdColumns(grdAPTINVH1, "VOUCHER_NO")
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")

    End Sub

    Sub Verify_Integrity()

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            If EntryMode <> "N" Then
                ASCMAIN1.sql = "Select * FROM POTORDR2 " _
                    & " where PO_QTY_SHP <> PO_QTY_REC " _
                    & "   and PO_QTY_REC <> 0 AND PO_QTY_OPN <> 0 " _
                    & "   and INIT_OPER <> 'conv' " _
                    & "   and PO_QTY_OPN = PO_QTY_ORD"
                Dim tbl As DataTable = ASCDATA1.GetDataTable
                If tbl.Rows.Count <> 0 Then
                    Using frm As New ASFMSGBF
                        frm.Show_grd(tbl, Me, "Shipment " & PO_SHIPMENT_NO & " - Please Take Screenshot and email to ABS")
                    End Using
                End If
            End If
        End If

    End Sub

    Sub Delete_Record()
        BeginTrans()
        If Not cost_calc Then
            Dependent_Updates(-1, PO_SHIPMENT_NO)
        End If
        For Each TABLE_NAME In New String() {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP5", "POTSHIP7", "POTSHIP8"}
            ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'")
        Next
        If ASCMAIN1.CLIENT = "VAN" AndAlso POTVBKG2_RECORDS Then
            BOOKING_INTEGRITY(PO_SHIPMENT_NO)
        End If
        CommitTrans("Shipment " & PO_SHIPMENT_NO & " has been Deleted")
    End Sub

    Sub Dependent_Updates(S As Integer, PO_SHIPMENT_NO As String)

        ' NEXT 3 LINES MAYBE TEMP - TO ALLOW EVA TO DESHIP AND THEN RESHIP FULL CONTAINERS
        'If dynPOTORDR2.item("PO_STATUS") & "" = "C" And QTY_OPN_NEW > 0 Then
        '    dynPOTORDR2.item("PO_STATUS") = "O"
        'End If

        '            & "      If R1.CLOSE_PO = '1' and " & CStr(S) & " = 1 Then QTY_OPN_NEW := 0; END IF;" & vbCrLf _

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare" & vbCrLf _
            & "     Cursor C1 is Select * from POTSHIP3  where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' for Update;" & vbCrLf _
            & "     QTY NUMBER(8,0);" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Begin" & vbCrLf _
            & "    Declare " & vbCrLf _
            & "     Cursor C2 is Select * from POTORDR2 where PO_ORDER_NO = R1.PO_ORDER_NO and PO_ORDER_LNO = R1.PO_ORDER_LNO for Update;" & vbCrLf _
            & "     WHSE_CODE_PO VARCHAR2(6);" & vbCrLf _
            & "     QTY_OPN_NEW NUMBER(8,0);" & vbCrLf _
            & "     QTY_OPN_OLD NUMBER(8,0);" & vbCrLf _
            & "     PO_QTY_SHP_NEW NUMBER(8,0);" & vbCrLf _
            & "     PO_STATUS_NEW VARCHAR2(1);" & vbCrLf _
            & "    Begin" & vbCrLf _
            & "     For R2 in C2 Loop" & vbCrLf _
            & "      PO_QTY_SHP_NEW := NVL(R2.PO_QTY_SHP,0) + " & CStr(S) & " * NVL(R1.PO_QTY_SHP,0);" & vbCrLf _
            & "      QTY_OPN_OLD := NVL(R2.PO_QTY_OPN,0);" & vbCrLf _
            & "      QTY_OPN_NEW := NVL(R2.PO_QTY_ORD,0) - PO_QTY_SHP_NEW;" & vbCrLf _
            & "      If R1.CLOSE_PO = '1' Then QTY_OPN_NEW := 0; END IF;" & vbCrLf _
            & "      PO_STATUS_NEW := R2.PO_STATUS;" & vbCrLf _
            & "      If QTY_OPN_NEW <=0 Then PO_STATUS_NEW := 'C'; End If;" & vbCrLf _
            & "      If PO_STATUS_NEW = 'C' and QTY_OPN_NEW >0 Then PO_STATUS_NEW := 'O'; End If;" & vbCrLf _
            & "      If QTY_OPN_NEW <=0 or PO_STATUS_NEW = 'C' Then QTY_OPN_NEW := 0; End If;" & vbCrLf _
            & "      Update POTORDR2 Set PO_QTY_SHP = PO_QTY_SHP_NEW, PO_QTY_OPN = QTY_OPN_NEW, PO_STATUS = PO_STATUS_NEW where Current of C2;" & vbCrLf _
            & "      Select WHSE_CODE into WHSE_CODE_PO from POTORDR1 where PO_ORDER_NO = R1.PO_ORDER_NO;" & vbCrLf _
            & "      QTY := (QTY_OPN_NEW - QTY_OPN_OLD);" & vbCrLf _
            & "      Update ICTSTAT2 Set WHSE_QTY_ON_ORDER = NVL(WHSE_QTY_ON_ORDER,0) + QTY" & vbCrLf _
            & "       where STYLE_CODE = R2.STYLE_CODE and COLOR_CODE = R2.COLOR_CODE and WHSE_CODE = WHSE_CODE_PO;" & vbCrLf _
            & "      If SQL%NOTFOUND then" & vbCrLf _
            & "       Insert into ICTSTAT2 (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_ON_ORDER)" & vbCrLf _
            & "        values (R2.STYLE_CODE, R2.COLOR_CODE, WHSE_CODE_PO, QTY);" & vbCrLf _
            & "      End If;" & vbCrLf _
            & "     Select WHSE_CODE into WHSE_CODE_PO from POTSHIP1 where PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO;" & vbCrLf _
            & "      QTY := " & CStr(S) & " * NVL(R1.PO_QTY_SHP,0);" & vbCrLf _
            & "      Update ICTSTAT2 Set WHSE_QTY_TRAN = NVL(WHSE_QTY_TRAN,0) + QTY" & vbCrLf _
            & "       where STYLE_CODE = R2.STYLE_CODE and COLOR_CODE = R2.COLOR_CODE and WHSE_CODE = WHSE_CODE_PO;" & vbCrLf _
            & "      If SQL%NOTFOUND then" & vbCrLf _
            & "       Insert into ICTSTAT2 (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_TRAN)" & vbCrLf _
            & "        values (R2.STYLE_CODE, R2.COLOR_CODE, WHSE_CODE_PO, QTY);" & vbCrLf _
            & "      End If;" & vbCrLf _
            & "     End Loop;" & vbCrLf _
            & "    End;" & vbCrLf _
            & "   End;" & vbCrLf _
            & "   Select Count (*) into QTY from POTORDR2 where PO_ORDER_NO = R1.PO_ORDER_NO and PO_STATUS = 'O';" & vbCrLf _
            & "   Update POTORDR1 Set PO_STATUS = (CASE WHEN QTY > 0 THEN 'O' ELSE 'C' END) where PO_ORDER_NO = R1.PO_ORDER_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        If ship_entry And (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
            ASCDATA1.ExecuteSQL("Truncate Table " & WHTSTYLX)
        End If

        If Not cost_calc Then
            For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
                rowPOTSHIP3.Item("PO_QTY_SHP") = Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "")
                rowPOTSHIP3.Item("PO_QTY_REC") = Val(rowPOTSHIP3.Item("PO_QTY_REC") & "")
            Next
            ASCDATA1.DeleteRows(dst.Tables("POTSHIP3"), "PO_QTY_SHP = 0 and PO_QTY_REC = 0")
            'ASCDATA1.DeleteRows(dst.Tables("POTSHIP7"), "STYLES = 0 OR TOTAL_UNITS = 0") ' NERVOUS ABOUT THIS - MAKE THEM DO IT MANUALLY
        End If

        BeginTrans()

        If cost_calc Then
            Update_Costs()
        Else
            If receipt_mode Then
                Update_Receipt()
            Else
                Update_Shipment()
            End If
        End If

        If automated_cost_complete Then
            CommitTrans("")
        Else
            CommitTrans("Update Complete")
            Verify_Integrity()

            If ASCMAIN1.Running_in_VS Or ASCMAIN1.USER_ID = "rgomez" Then
                Dim sqlIC = TAC.POCMAIN1.Get_sql_Integrity_Check
                Dim tbl As DataTable = ASCDATA1.GetDataTable(sqlIC)
                If tbl.Rows.Count <> 0 Then
                    If Format(Now, "MM/dd/yy") = "05/11/17" Then
                    Else
                        MsgBox("Please email a Screenshot to Walter, and describe your work on Shipment " & PO_SHIPMENT_NO, MsgBoxStyle.OkOnly, "PO Shipments are Out of Balance")
                    End If
                End If
            End If
        End If

    End Sub

    Sub Update_Costs()
        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'")

            Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP3.Item("PO_SHIPMENT_LNO") & "")
            Dim PO_ORDER_NO As String = rowPOTSHIP3.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(rowPOTSHIP3.Item("PO_ORDER_LNO") & "")

            Dim sqlw As String = "PO_SHIPMENT_NO = {0} and PO_SHIPMENT_LNO = {1} and PO_ORDER_NO = {2} and PO_ORDER_LNO = {3}"
            sqlw = String.Format(sqlw, PO_SHIPMENT_NO, PO_SHIPMENT_LNO, PO_ORDER_NO, PO_ORDER_LNO)
            Dim rowPOTSHIP3s_orig() As DataRow = dst.Tables("POTSHIP3").Select(sqlw, "", DataViewRowState.OriginalRows)
            Dim rowPOTSHIP3_orig As DataRow = Nothing
            If rowPOTSHIP3s_orig.Length = 1 Then
                rowPOTSHIP3_orig = rowPOTSHIP3s_orig(0)

                For Each COLUMN_NAME As String In New String() _
                    {"PO_COST_VCOST", "PO_COST_MATLS", "PO_COST_COMM", "PO_COST_OTHER", "PO_COST_BUFFER", "PO_COST_QUOTA", "PO_COST_QUOTA_DF"}
                    If Val(rowPOTSHIP3.Item(COLUMN_NAME) & "") <>
                        Val(rowPOTSHIP3_orig.Item(COLUMN_NAME) & "") Then
                        rowPOTSHIP3.Item("COST_CHANGED") = "1"
                        Exit For
                    End If
                Next
            Else
                rowPOTSHIP3.Item("COST_CHANGED") = "1"
            End If

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                Dim WHSE_CODE As String = rowPOTSHIP1.Item("WHSE_CODE") & ""
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                Dim COUNTRY_CODE As String = rowICTWHSE1.Item("WHSE_COUNTRY") & ""

                ' REALLY SHOULD BE TEMP - NO AUDIT TRAIL, AND WE REALLY NEED THE DUTY RATE CODE PRIOR TO PLACING THE PO
                Dim STYLE_CODE As String = rowPOTSHIP3.Item("STYLE_CODE")
                Dim DUTY_RATE_CODE As String = rowPOTSHIP3.Item("DUTY_RATE_CODE") & ""
                If DUTY_RATE_CODE <> "" Then
                    If COUNTRY_CODE = "USA" Or COUNTRY_CODE = "" Then
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        If rowICTSTYL1.Item("DUTY_RATE_CODE") & "" <> DUTY_RATE_CODE Then
                            ASCMAIN1.sql = "Update ICTSTYL1 Set DUTY_RATE_CODE = '" & DUTY_RATE_CODE & "' where STYLE_CODE = '" & STYLE_CODE & "'"
                            ASCDATA1.ExecuteSQL()
                        End If
                    Else
                        If LookUp("ICTSTYLC", New String() {STYLE_CODE, COUNTRY_CODE}) Is Nothing Then
                            ASCMAIN1.sql = "Insert into ICTSTYLC (STYLE_CODE, COUNTRY_CODE, DUTY_RATE_CODE) values (:PARM1,:PARM2,:PARM3)"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New String() {STYLE_CODE, COUNTRY_CODE, DUTY_RATE_CODE})
                        End If
                    End If
                End If
            End If
        Next

        For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
            Dim OPS_YYYYPP_FIFO As String = rowPOTSHIP2.Item("OPS_YYYYPP_FIFO") & ""
            Dim OPS_YYYYPP_FIFO_ORIG As String = rowPOTSHIP2.Item("OPS_YYYYPP_FIFO", DataRowVersion.Original) & ""
            If OPS_YYYYPP_FIFO <> OPS_YYYYPP_FIFO_ORIG Then
                Write_Audit_Trail(rowPOTSHIP2)
            End If
        Next

        If rowPOTSHIP1.Item("COST_COMPLETE") & "" = "1" Then
            rowPOTSHIP1.Item("COST_COMPLETE_OPS_YYYYPP") = ASCMAIN1.CYP
            rowPOTSHIP1.Item("COST_COMPLETE_INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTSHIP1.Item("COST_COMPLETE_INIT_DATE") = DATETIME_STAMP
        End If

        '   ASCDATA1.DeleteRows("POTSHIP5", "ISNULL(CTL_NO,'') <> ''")

        Dim sqlx As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
        Update_Record_TDA("POTSHIP1", sqlx)
        Update_Record_TDA("POTSHIP2", sqlx)
        Update_Record_TDA("POTSHIP3", sqlx)
        Update_Record_TDA("POTSHIP4", sqlx)
        Update_Record_TDA("POTSHIP5", sqlx)
        Update_Record_TDA("POTSHIP7")

        ' NOTE - PROBLEM IF SAME PO & LNO IS USED MORE THAN ONCE ON THE SAME POTSHIP2 RECORD
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select * from POTSHIP2 where PO_SHIPMENT_NO ='" & PO_SHIPMENT_NO & "';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Begin Declare Cursor C2 is" & vbCrLf _
            & "    Select * from POTSHIP3 where PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO and PO_SHIPMENT_LNO = R1.PO_SHIPMENT_LNO;" & vbCrLf _
            & "    Begin" & vbCrLf _
            & "     For R2 in C2 Loop" & vbCrLf _
            & "      Update ICTTRAN2 set STYLE_COST = R2.PO_COST_LANDED" & vbCrLf _
            & "       where ICTTRAN2.OPS_YYYYPP = R1.OPS_YYYYPP" & vbCrLf _
            & "         and ICTTRAN2.TRAN_TYPE = 'R'" & vbCrLf _
            & "         and ICTTRAN2.TRAN_NO = R1.TRAN_NO" & vbCrLf _
            & "         and ICTTRAN2.PO_ORDER_NO = R2.PO_ORDER_NO" & vbCrLf _
            & "         and ICTTRAN2.PO_ORDER_LNO = R2.PO_ORDER_LNO;" & vbCrLf _
            & "      Update ICTIREC2 set STYLE_COST = R2.PO_COST_LANDED" & vbCrLf _
            & "       where ICTIREC2.RECEIPT_NO = R1.TRAN_NO" & vbCrLf _
            & "         and ICTIREC2.PO_ORDER_NO = R2.PO_ORDER_NO" & vbCrLf _
            & "         and ICTIREC2.PO_ORDER_LNO = R2.PO_ORDER_LNO;" & vbCrLf _
            & "     End Loop;" & vbCrLf _
            & "    End;" & vbCrLf _
            & "   End;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"

        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Update_Receipt()

        Dim WHSE_CODE As String = rowPOTSHIP1.Item("WHSE_CODE")
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim LOCATION_CODE As String = rowICTWHSE1.Item("WHSE_LOC_REC") & ""

        Dim RECEIPT_NO_REVERSED As New List(Of String)
        Dim RECEIPT_NO_UPDATED As New List(Of String)

        Dim XYP As String = Set_Period()

        Dim ORDR_NO_BOLs As New Dictionary(Of String, String)


        For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'X' or PO_SHIP_STATUS = 'R'")
            Dim WHSE_CODE_REC As String = WHSE_CODE
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                If rowPOTSHIP2.Item("WH_REC_NO") & "" <> "" Then
                    ASCMAIN1.sql = "Select * from WHTWREC1 " _
                    & " where WH_REC_NO = '" & rowPOTSHIP2.Item("WH_REC_NO") & "'"
                    Dim rowWHTWREC1 As DataRow = ASCDATA1.GetDataRow
                    If rowWHTWREC1 IsNot Nothing Then
                        WHSE_CODE_REC = rowWHTWREC1.Item("WHSE_CODE")
                    End If
                End If
            End If

            ' TRACK THE (1ST) BOL_NO USED FOR AN ORDER WHEN RECEIVING A BTB
            Dim ORDR_NO As String = rowPOTSHIP2.Item("ORDR_NO") & ""
            Dim BOL_NO As String = rowPOTSHIP2.Item("BOL_NO") & ""
            If ORDR_NO <> "" Then
                If Not ORDR_NO_BOLs.ContainsKey(ORDR_NO) Then
                    ORDR_NO_BOLs.Add(ORDR_NO, BOL_NO)
                End If
            End If

            Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
            Dim S As Integer = 1
            Dim TRAN_NO As String = ""
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                TRAN_NO = ASCMAIN1.Next_Control_No("TRAN_NO_R")
            Else
                TRAN_NO = ASCMAIN1.Next_Control_No("ICTIREC1.RECEIPT_NO")
            End If

            RECEIPT_NO_UPDATED.Add(TRAN_NO)

            Dim TRAN_NO_original As String = rowPOTSHIP2.Item("TRAN_NO") & ""

            '   Dim rowICTIREC1x As DataRow = dst.Tables("ICTIREC1").Rows.Find({TRAN_NO_original})

            If rowPOTSHIP2.Item("PO_SHIP_STATUS") = "R" Then
                Dim ACCRUAL_STATUS_REV As String = ""
                S = -1
                RECEIPT_NO_REVERSED.Add(TRAN_NO_original)
                ' PHASE THIS TABLE OUT
                ASCMAIN1.sql = "Update ICTTRAN1 Set TRAN_STATUS_UPD = 'R'" _
                    & " where TRAN_NO = :PARM1" _
                    & "   and TRAN_TYPE = 'R'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {TRAN_NO_original})

                ASCMAIN1.sql = "Update ICTIREC1 Set REVERSED_BY_RECEIPT_NO = '" & TRAN_NO & "'" _
                    & " where RECEIPT_NO = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {TRAN_NO_original})
            End If

            Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
            Dim rowPOTSHIP3s() As DataRow = dst.Tables("POTSHIP3").Select(sqlw)
            Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(rowPOTSHIP3s(0).Item("PO_ORDER_NO"))

            Dim rowICTIREC1_original As DataRow = Nothing
            If S = -1 Then
                rowICTIREC1_original = LookUp("ICTIREC1", TRAN_NO_original)
            End If
            Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").NewRow
            With rowICTIREC1
                .Item("RECEIPT_NO") = TRAN_NO
                If S = -1 Then
                    .Item("RECEIPT_DATE") = rowICTIREC1_original.Item("RECEIPT_DATE")
                    .Item("SOURCE_DOC_NO") = rowICTIREC1_original.Item("SOURCE_DOC_NO")
                    .Item("QTY_REC") = -1 * Val(rowICTIREC1_original.Item("QTY_REC") & "")
                    .Item("AMT_REC") = -1 * Val(rowICTIREC1_original.Item("AMT_REC") & "")
                Else
                    .Item("RECEIPT_DATE") = Absx1.dteFor("PO_DATE_RECEIVED").Value
                    .Item("SOURCE_DOC_NO") = Absx1.txtFor("PO_SOURCE_DOC").Text
                    .Item("QTY_REC") = Val(dst.Tables("POTSHIP3").Compute("SUM(PO_QTY_REC)", sqlw) & "")
                    .Item("AMT_REC") = Val(dst.Tables("POTSHIP3").Compute("SUM(PO_AMT_REC)", sqlw) & "")
                End If

                .Item("VEND_CODE") = rowPOTORDR1.Item("VEND_CODE")
                '.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("OPS_YYYYPP") = XYP ' Set_Period() ' NEED TO SET PERIOD TO CORRESPOND TO RECEIPT DATE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("REGISTER_IND") = "0"
                .Item("WHSE_CODE") = WHSE_CODE_REC 'WHSE_CODE - change to handle 1 shipment being received into multiple whses

                .Item("ACCRUAL_STATUS") = "0"
                If ASCMAIN1.CLIENT = "VAN" Then
                    If rowPOTSHIP2.Item("ACCRUAL_STATUS") & "" = "1" Then
                        .Item("ACCRUAL_STATUS") = "1"
                        '                       MsgBox("Please alert ABS to check Accrual Status on Receipt " & TRAN_NO)
                    End If
                End If

                If S = -1 Then .Item("REVERSES_RECEIPT_NO") = TRAN_NO_original
                .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            End With
            dst.Tables("ICTIREC1").Rows.Add(rowICTIREC1)

            If rowPOTSHIP2.Item("PO_SHIP_STATUS") = "X" Then
                rowPOTSHIP2.Item("PO_SHIP_STATUS") = "C"

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    ' PHASE THIS TABLE OUT
                    Dim rowICTTRAN1 As DataRow = dst.Tables("ICTTRAN1").NewRow
                    With rowICTTRAN1
                        .Item("OPS_YYYYPP") = XYP ' ASCMAIN1.CYP
                        .Item("TRAN_TYPE") = "R"
                        .Item("TRAN_NO") = TRAN_NO

                        .Item("TRAN_SOURCE_DOCUMENT") = Absx1.txtFor("PO_SOURCE_DOC").Text
                        .Item("TRAN_DATE") = Absx1.dteFor("PO_DATE_RECEIVED").Value
                        .Item("TRAN_WHSE_CODE") = WHSE_CODE
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("TRAN_STATUS_UPD") = "U"

                        .Item("TRAN_CCVRW_DESC") = rowICTWHSE1.Item("WHSE_DESC") & ""
                        .Item("TRAN_ORIGINATE") = "P"
                        .Item("TRAN_TYPE_ORIG") = "P"
                        .Item("TRAN_NO_ORIG") = PO_SHIPMENT_NO
                    End With
                    dst.Tables("ICTTRAN1").Rows.Add(rowICTTRAN1)

                End If

                rowPOTSHIP2.Item("TRAN_NO") = TRAN_NO
                rowPOTSHIP2.Item("OPS_YYYYPP") = XYP ' ASCMAIN1.CYP
                rowPOTSHIP2.Item("PO_SOURCE_DOC") = Absx1.txtFor("PO_SOURCE_DOC").Text
                rowPOTSHIP2.Item("PO_DATE_RECEIVED") = Absx1.dteFor("PO_DATE_RECEIVED").Value
                rowPOTSHIP2.Item("PO_DATE_RECEIVED_PORT") = rowPOTSHIP1.Item("PO_SHIP_ETA") '  rowPOTSHIP2.Item("PO_DATE_RECEIVED")
                rowPOTSHIP2.Item("PO_DATE_RECEIVED_WHSE") = rowPOTSHIP2.Item("PO_DATE_RECEIVED")

            ElseIf rowPOTSHIP2.Item("PO_SHIP_STATUS") = "R" Then
                rowPOTSHIP2.Item("PO_SHIP_STATUS") = "O"
                rowPOTSHIP2.Item("TRAN_NO") = DBNull.Value
                rowPOTSHIP2.Item("OPS_YYYYPP") = DBNull.Value
                rowPOTSHIP2.Item("PO_SOURCE_DOC") = DBNull.Value
                rowPOTSHIP2.Item("PO_DATE_RECEIVED") = DBNull.Value
            End If

            Dim TRAN_LNO As Integer = 0
            Dim RECEIPT_LNO As Integer = 0

            Dim msg_shown As Boolean = False

            For Each rowPOTSHIP3 As DataRow In rowPOTSHIP3s
                Dim PO_ORDER_NO As String = rowPOTSHIP3.Item("PO_ORDER_NO")
                Dim PO_ORDER_LNO As Integer = Val(rowPOTSHIP3.Item("PO_ORDER_LNO") & "")

                Dim PO_QTY_SHP As Int64 = Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "")
                Dim PO_QTY_REC_OLD As Int64 = Val(rowPOTSHIP3.Item("PO_QTY_REC_OLD") & "")
                Dim PO_QTY_REC As Int64 = Val(rowPOTSHIP3.Item("PO_QTY_REC") & "")
                If S = -1 Then
                    PO_QTY_REC = 0
                    If PO_QTY_REC_OLD > PO_QTY_SHP And ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "rick" Then
                        Stop
                        PO_QTY_REC_OLD = PO_QTY_SHP
                    End If
                Else
                    If PO_QTY_REC_OLD <> 0 And ASCMAIN1.CLIENT = "RGI" And Not msg_shown Then
                        msg_shown = True
                        MsgBox("Possible Data Correlation issue - please contact ABS - please do NOT click OK", MsgBoxStyle.OkOnly, "Warning")
                    End If
                End If

                rowPOTSHIP3.Item("PO_QTY_REC") = PO_QTY_REC

                Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})

                Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowPOTORDR2.Item("COLOR_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)

                rowPOTORDR2.Item("PO_QTY_REC") = Val(rowPOTORDR2.Item("PO_QTY_REC") & "") + PO_QTY_REC - PO_QTY_REC_OLD

                ' Inventory Status Updates (ICTSTAT2): TRAN, ON_HAND
                Dim QTY As Int64
                QTY = -1 * S * PO_QTY_SHP
                Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_TRAN", QTY)
                QTY = (PO_QTY_REC - PO_QTY_REC_OLD)
                Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE_REC, "WHSE_QTY_ON_HAND", QTY, XYP)
                'Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_ON_HAND", QTY, XYP) '  - change to handle 1 shipment being received into multiple whses


                ' Inventory Transaction File

                Dim rowICTIREC2 As DataRow = dst.Tables("ICTIREC2").NewRow
                With rowICTIREC2
                    .Item("RECEIPT_NO") = TRAN_NO
                    RECEIPT_LNO = RECEIPT_LNO + 1
                    .Item("RECEIPT_LNO") = RECEIPT_LNO
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("QTY_REC") = (PO_QTY_REC - PO_QTY_REC_OLD)
                    .Item("STYLE_COST") = rowPOTSHIP3.Item("PO_COST_LANDED")
                    .Item("PO_COST") = rowPOTSHIP3.Item("PO_COST")
                    '.Item("AP_COST") = Val(rowPOTSHIP3.Item("PO_COST_VCOST") & "") + Val(rowPOTSHIP3.Item("PO_COST_OTHER") & "") _
                    '    - Val(rowPOTSHIP3.Item("PO_COST_VCOST") & "") * Val(rowPOTSHIP3.Item("PO_COST_COMM") & "") / 100

                    ' NOTE - THE FOLLOWING IS CODED TO NOT DISTURB NYA, AND RGI.  
                    ' BUT I THINK THAT THIS Is HOW IT SHOULD BE FOR EVERYONE EXCEPT NYA 
                    ' PROB WONT MATTER TO RGI WHO DOES NOT HAVE MATLS OR COMMISSION
                    ' NYA's COMMISSION (IT SEEMS) IS BURIED IN THE PO COST, 
                    ' And THEY DON'T WANT TO PAY THE COMMISSION WITH THE AP 
                    ' - NOT SURE WHY BUT I REMEMBER WRANGLING ABOUT THIS WITH LESLIE

                    If ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "RGI" Then
                        .Item("AP_COST") = Val(rowPOTSHIP3.Item("PO_COST_VCOST") & "") + Val(rowPOTSHIP3.Item("PO_COST_OTHER") & "") _
                        - Val(rowPOTSHIP3.Item("PO_COST_VCOST") & "") * Val(rowPOTSHIP3.Item("PO_COST_COMM") & "") / 100
                    Else
                        .Item("AP_COST") = Val(rowPOTSHIP3.Item("PO_COST") & "") - Val(rowPOTSHIP3.Item("PO_COST_MATLS") & "")
                    End If

                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                    .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                    .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                    .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
                    .Item("OPS_YYYYPP") = XYP '  ASCMAIN1.CYP
                    .Item("QTY_SHP") = PO_QTY_SHP
                End With
                dst.Tables("ICTIREC2").Rows.Add(rowICTIREC2)

                If S = 1 Then
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        ' ELIMINATE THIS TABLE
                        Dim rowICTTRAN2 As DataRow = dst.Tables("ICTTRAN2").NewRow
                        With rowICTTRAN2
                            .Item("OPS_YYYYPP") = XYP ' ASCMAIN1.CYP
                            .Item("TRAN_TYPE") = "R"
                            .Item("TRAN_NO") = TRAN_NO
                            TRAN_LNO = TRAN_LNO + 1
                            .Item("TRAN_LNO") = TRAN_LNO
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                            .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
                            .Item("STYLE_COST") = rowPOTSHIP3.Item("PO_COST_LANDED")
                            .Item("TRAN_QTY") = (PO_QTY_REC - PO_QTY_REC_OLD)
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                        End With
                        dst.Tables("ICTTRAN2").Rows.Add(rowICTTRAN2)
                    End If
                End If
            Next

            Update_Record_TDA("POTSHIP2")
            Update_Record_TDA("POTSHIP3")
            Update_Record_TDA("POTORDR2")
            Update_Record_TDA("ICTIREC1")
            Update_Record_TDA("ICTIREC2")

            If S = 1 Then
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    Update_Record_TDA("ICTTRAN1") ' ELIMINATE
                    Update_Record_TDA("ICTTRAN2") ' ELIMINATE
                End If
            End If

            If rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1" Then
                ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                       New Object() {"R", TRAN_NO, ASCMAIN1.SESSION_NO},
                       New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
                'want to bring receipt back to life here if s = -1
            End If
        Next

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else
            If (WHSE_TYPE = "P" Or (ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "NC")) Then Update_BTB_Invoices(ORDR_NO_BOLs)
        End If

        If select_from_3PL_list Then

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

            ElseIf ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then

                ASCMAIN1.sql = "Update EDT944T1" _
                    & " Set EDI_PROCESS_IND = '1',LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                    & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_PROCESS_IND = '0'"
                ASCDATA1.ExecuteSQL()

                For Each COLUMN_NAME As String In New String() {"CARTON_PACK_QTY", "CASE_WEIGHT_GRS", "CASE_CUBE"}
                    Dim EXP As String = ""
                    Select Case COLUMN_NAME
                        Case "CARTON_PACK_QTY"
                            EXP = "EDI_PACK_QTY"
                        Case "CASE_WEIGHT_GRS"
                            EXP = "EDI_WEIGHT"
                        Case "CASE_CUBE"
                            EXP = "EDI_VOLUME"
                    End Select

                    ASCMAIN1.sql = "Insert into ASTAUDT1 (" & vbCrLf _
                        & "TABLE_NAME,KEY_VALUE,COLUMN_NAME,USER_ID,INIT_DATE" & vbCrLf _
                        & ",OLD_VALUE,NEW_VALUE" & vbCrLf _
                        & ",FM_MODE,NOTES,KEY_VALUE2,KEY_LNO,SESSION_NO,SELECTION_NO,XNO)" & vbCrLf _
                        & "Select" & vbCrLf _
                        & "'ICTSTYL1' TABLE_NAME, ICTSTYL1.STYLE_CODE KEY_VALUE,'" & COLUMN_NAME & "' COLUMN_NAME" & vbCrLf _
                        & ",'" & ASCMAIN1.USER_ID & "' USER_ID,SYSDATE INIT_DATE" & vbCrLf _
                        & ",ICTSTYL1." & COLUMN_NAME & " OLD_VALUE, EDT944T2." & EXP & vbCrLf _
                        & ",'A' FM_MODE,NULL NOTES,EDT944T2.EDI_DOC_SEQ_NO KEY_VALUE2,EDT944T2.EDI_DOC_LNO KEY_LNO,'" & ASCMAIN1.SESSION_NO & "' SESSION_NO," & CStr(Me.SELECTION_NO) & " SELECTION_NO,'" & Me.XNO & "' XNO" & vbCrLf _
                        & " from EDT944T2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = EDT944T2.EDI_STYLE_NO" & vbCrLf _
                        & "   and EDT944T2.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                        & "   and NVL(ICTSTYL1." & COLUMN_NAME & ",0) <> EDT944T2." & EXP
                    ASCDATA1.ExecuteSQL()
                Next

                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is Select * from EDT944T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "';" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update ICTSTYL1 " & vbCrLf _
                    & "    Set CASE_WEIGHT_GRS = R1.EDI_WEIGHT" & vbCrLf _
                    & "      , CASE_CUBE = R1.EDI_VOLUME" & vbCrLf _
                    & "      , CARTON_PACK_QTY = R1.EDI_PACK_QTY" & vbCrLf _
                    & "    where STYLE_CODE = R1.EDI_STYLE_NO;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()
            End If
        ElseIf Select_from_Whse_Receipt Or ASCMAIN1.CLIENT = "RGI" Then

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.CLIENT = "RGI" Then
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("WHTWRECD").Select(""), "WH_REC_NO").Rows
                    ASCMAIN1.sql = "Update WHTWREC1 Set WH_REC_STATUS = 'R'" _
                    & " where WH_REC_NO = '" & row.Item("WH_REC_NO") & "'"
                    ASCDATA1.ExecuteSQL()
                Next
            End If
        Else

            If RECEIPT_NO_REVERSED.Count <> 0 And rowICTWHSE1.Item("LP_CODE") & "" <> "" Then

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

                ElseIf ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    ASCMAIN1.sql = "Update EDT944T1" _
                        & " Set EDI_PROCESS_IND = '0',LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                        & " where EDI_PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and EDI_PROCESS_IND = '1'"
                    ASCDATA1.ExecuteSQL()
                End If
            End If
        End If
    End Sub

    Sub Update_BTB_Invoices(ORDR_NO_BOLs As Dictionary(Of String, String))

        dst.Tables("SOTORDR1").Rows.Clear()

        Dim ORDR_GROUP_NOs As New List(Of String)
        Dim INV_NOs As New List(Of String)

        For Each rowSOTORDP1 As DataRow In dst.Tables("SOTORDP1").Select("")
            Dim ORDR_NO As String = rowSOTORDP1.Item("ORDR_NO")
            Dim INV_NO_PREV As String = rowSOTORDP1.Item("INV_NO_PREV") & ""
            Dim INV_NO As String = INV_NO_PREV
            If INV_NO = "" Then
                INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            Else
                Dim sqlwP As String = " where ORDR_NO = '" & ORDR_NO & "' and INV_NO = '" & INV_NO_PREV & "'"
                ASCDATA1.ExecuteSQL("Delete from SOTORDP1" & sqlwP)
                ASCDATA1.ExecuteSQL("Delete from SOTORDP2" & sqlwP)
            End If

            rowSOTORDP1.Item("INV_STATUS") = "1"
            rowSOTORDP1.Item("INV_NO") = INV_NO

            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "'"
            Dim SHIP_CNT_CARTONS As Int64 = Val(dst.Tables("POTSHIP2").Compute("SUM(PO_SHIP_CTNS)", sqlw) & "")
            Dim SHIP_TOTAL_WGT As Decimal = Val(dst.Tables("POTSHIP2").Compute("SUM(TOTAL_WEIGHT)", sqlw) & "")

            ' Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
            Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO, , False)
            Dim ORDR_PICK_SEQ As Integer = Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") + 1

            Dim PICK_NO As String = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")
            Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
            Dim ORDR_GROUP_NO As String = rowSOTORDR1.Item("ORDR_GROUP_NO")
            If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)

            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = "I"
                .Item("INV_NO") = INV_NO
                .Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
                .Item("CUST_STORE_NO") = rowSOTORDR1.Item("CUST_STORE_NO")
                .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                .Item("ORDR_NO") = ORDR_NO

                .Item("CUST_FACTOR_IND") = rowSOTORDR1.Item("CUST_FACTOR_IND")

                .Item("WHSE_CODE") = rowPOTSHIP1.Item("WHSE_CODE")
                .Item("INV_SALES") = rowSOTORDP1.Item("INV_TOTAL_AMOUNT")
                .Item("INV_COGS") = 0
                .Item("INV_FREIGHT") = 0
                .Item("INV_MISC_CHG") = 0
                .Item("INV_TOTAL_AMOUNT") = rowSOTORDP1.Item("INV_TOTAL_AMOUNT")

                .Item("INV_DATE") = rowSOTORDP1.Item("INV_DATE")
                .Item("ORDR_DATE_UPDATED") = DATETIME_STAMP
                .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP

                ' ELIMINATE THIS
                .Item("ORDR_BILL_TO_CUST") = rowSOTORDR1.Item("CUST_BILL_TO_CUST")

                .Item("POST_CODE") = rowSOTORDR1.Item("POST_CODE")

                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("SALES_DIVISION_CODE") = rowSOTORDR1.Item("SALES_DIVISION_CODE")

                .Item("INV_PRINTED") = DBNull.Value

                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    .Item("INV_PRINTED") = DATETIME_STAMP

                    ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                    & " values ('SOTORDR1','" & ORDR_NO & "', SYSDATE, '" & ASCMAIN1.USER_ID & "', 'INVPRTB2B','Invoice Marked as Printed', '" & INV_NO & "')"
                    ASCDATA1.ExecuteSQL()
                End If

                .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("PICK_NO") = PICK_NO
                .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                .Item("INV_COMMENT") = rowSOTORDP1.Item("INV_COMMENT")
                .Item("SREP2_CODE") = rowSOTORDR1.Item("SREP_CODE")

                .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT")

                .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                .Item("CURR_EXCH_RATE") = 1

                .Item("INV_SALES_CURR") = .Item("INV_SALES")
                .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
                .Item("INV_MISC_CHG_CURR") = .Item("INV_MISC_CHG")

                ' ELIMINATE THIS
                .Item("INV_TOTAL_AMT_CURR") = .Item("INV_TOTAL_AMOUNT")

                .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")

                .Item("ORDR_TYPE_CODE") = "BTB"
                .Item("CUST_BILL_TO_CUST") = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
            End With
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)



            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
            With rowSOTPICK1
                .Item("PICK_NO") = PICK_NO
                .Item("ORDR_NO") = ORDR_NO
                .Item("PICK_FREIGHT") = 0
                .Item("PICK_PICKER") = ASCMAIN1.USER_ID
                .Item("ORDR_PICK_SEQ") = ORDR_PICK_SEQ
                .Item("PICK_STATUS") = "F"
                .Item("PICK_RELEASED") = DATETIME_STAMP
                .Item("PICK_PRINTED") = DATETIME_STAMP
                .Item("PICK_PACKED") = DATETIME_STAMP
                .Item("PICK_SHIPPED") = DATETIME_STAMP
                .Item("PICK_BATCH_NO") = "000000"
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("INV_NO") = INV_NO
                .Item("PICK_CNT_CARTONS") = SHIP_CNT_CARTONS
                .Item("PICK_TOTAL_WGT") = SHIP_TOTAL_WGT
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
            End With
            dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)


            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
            With rowSOTSHIP1
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("SHIP_DATE_SHIPPED") = rowPOTSHIP1.Item("PO_DATE_SHIPPED")
                .Item("SHIP_VIA_CODE") = "BTB"
                .Item("SHIP_REF") = rowSOTORDP1.Item("INV_REF")
                .Item("SHIP_TOTAL_WGT") = SHIP_TOTAL_WGT
                .Item("SHIP_CNT_CARTONS") = SHIP_CNT_CARTONS
                .Item("SHIP_ADDR_TYPE") = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
                .Item("SHIP_ADDR_CODE") = IIf(rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") = "MK", rowSOTORDR1.Item("CUST_STORE_NO"), rowSOTORDR1.Item("CUST_DC_NO"))
                .Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO")
                .Item("SHIP_PICK_PRINTED") = DATETIME_STAMP
                .Item("PICK_BATCH_NO") = "000000"
                .Item("SHIP_STATUS") = "F"
                .Item("FRT_TERMS") = rowSOTORDR1.Item("FRT_TERMS")
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("INV_DATE") = rowSOTORDP1.Item("INV_DATE")
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("BILL_OF_LADING_NO") = rowPOTSHIP1.Item("PO_SHIP_REF_NO")

                If ASCMAIN1.CLIENT = "NYA" Then
                    ' NOTE - NYAG DESIGN IS 1:1 BETWEEN POTSHIP2 AND SOTINVH1 - RGI IS NOT - SEE RGI ORDR_NO 0000420306
                    ' IF NYAG EVER GENERATED MORE THAN 1 INVOICE ON A SHIPMENT,
                    '  THEN THIS BOL_NO DETERMINATION WOULD FAIL - AND SO WOULD CTNS/WGT CALCS ABOVE
                    ' Dim BOL_NO As String = dst.Tables("POTSHIP2").Compute("MAX(BOL_NO)", sqlw) & ""
                    If ORDR_NO_BOLs.ContainsKey(ORDR_NO) Then
                        Dim BOL_NO As String = ORDR_NO_BOLs(ORDR_NO)
                        '  .Item("BILL_OF_LADING_NO") = BOL_NO
                        .Item("BTB_BOL_NO") = BOL_NO
                    End If
                End If

                .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
                .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                .Item("SHIP_NOTES") = rowPOTSHIP1.Item("PO_NOTES")
                .Item("SHIPPED_ACTUAL") = rowPOTSHIP1.Item("PO_DATE_SHIPPED")
                .Item("CUST_FACTOR_TRANS_IND") = rowSOTORDR1.Item("CUST_FACTOR_IND")
                .Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE")
                .Item("SHIP_SPEC_INST") = rowSOTORDR1.Item("ORDR_SHIP_INSTR")
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP

                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP

                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
            End With
            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)

            ' Create_TDA(.Tables.Add, "SOTORDR1", "*", 1, , , , "ORDR_PICK_SEQ,ORDR_STATUS") ' ORDR_DATE_CLOSED,ORDR_YYYYPP_CLOSED
            ' Create_TDA(.Tables.Add, "SOTORDR2", "*", 1, , , , "ORDR_QTY_OPEN,ORDR_QTY_SHIP,ORDR_QTY_CANC,ORDR_STATUS")

            sqlw = "ORDR_NO = '" & ORDR_NO & "' and INV_NO = '" & INV_NO & "'"
            For Each rowSOTORDP2 As DataRow In dst.Tables("SOTORDP2").Select(sqlw, "ORDR_LNO")
                Dim ORDR_LNO As Integer = Val(rowSOTORDP2.Item("ORDR_LNO") & "")
                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTORDP2.Item("ORDR_QTY_SHIP") & "")

                Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").NewRow
                With rowSOTINVH2
                    .Item("INV_TYPE") = "I"
                    .Item("INV_NO") = INV_NO
                    .Item("INV_LNO") = ORDR_LNO
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("ORDR_UNIT_COST") = rowICTSTYL1.Item("STYLE_COST")
                    .Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                    .Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")

                    .Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
                    .Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
                    .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP ' MAYBE SHOULD USE SAME PERIOD AS RECEIPT

                    .Item("STYLE_CUST_CODE") = rowICTSTYL1.Item("CUST_CODE")
                    .Item("ORDR_PRICE_SOURCE") = rowSOTORDR2.Item("ORDR_PRICE_SOURCE")
                    .Item("COMM_RATE") = rowSOTORDR2.Item("COMM_RATE")

                End With
                dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)
                'ORDR_YYYYPP_UPDATED

                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
                With rowSOTPICK2
                    .Item("PICK_NO") = PICK_NO
                    .Item("PICK_LNO") = ORDR_LNO
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = ORDR_LNO
                    .Item("PICK_QTY") = ORDR_QTY_SHIP
                    .Item("PICK_QTY_CONF") = ORDR_QTY_SHIP
                    .Item("PICK_QTY_CANC") = 0
                    .Item("PICK_QTY_BACK") = 0
                    .Item("PICK_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                    .Item("PICK_QTY_CANC_REL") = 0
                    .Item("PICK_QTY_BACK_REL") = 0
                End With
                dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)

                Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                ORDR_QTY_OPEN = ORDR_QTY_OPEN - ORDR_QTY_SHIP
                If ORDR_QTY_OPEN < 0 Then ORDR_QTY_OPEN = 0
                If ORDR_QTY_OPEN = 0 Then rowSOTORDR2.Item("ORDR_STATUS") = "F"
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                rowSOTORDR2.Item("ORDR_QTY_SHIP") = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & "") + ORDR_QTY_SHIP

            Next

            rowSOTORDR1.Item("ORDR_PICK_SEQ") = ORDR_PICK_SEQ

            Dim ORDR_STATUS As String = "O"
            If Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", "ORDR_NO = '" & ORDR_NO & "'") & "") = 0 Then ORDR_STATUS = "F"
            rowSOTORDR1.Item("ORDR_STATUS") = ORDR_STATUS


            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
            With rowARTOPEN1

                .Item("CUST_CODE") = rowSOTINVH1.Item("CUST_CODE")
                .Item("INV_TYPE") = rowSOTINVH1.Item("INV_TYPE")
                .Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")
                .Item("INV_DATE") = rowSOTINVH1.Item("INV_DATE")
                .Item("CUST_STORE_NO") = rowSOTINVH1.Item("CUST_STORE_NO")
                .Item("POST_CODE") = rowSOTINVH1.Item("POST_CODE")
                .Item("TERM_CODE") = rowSOTINVH1.Item("TERM_CODE")

                Dim INV_DATE As Date = rowSOTINVH1.Item("INV_DATE")
                Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, rowSOTINVH1.Item("TERM_CODE"), Nothing, INV_DATE)

                .Item("INV_DUE_DATE") = INV_DUE_DATE

                '.Item("INV_DISC_DATE") = "?"

                .Item("SREP_CODE") = rowSOTINVH1.Item("SREP_CODE")
                '.Item("STAX_CODE") = rowSOTINVH1.Item("STAX_CODE")
                '.Item("APPLY_TO_INV_NUM")
                '.Item("APPLY_TO_INV_TYPE")
                .Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
                .Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO")

                .Item("INV_SALES") = rowSOTINVH1.Item("INV_SALES")
                .Item("INV_DISC") = 0
                .Item("INV_FREIGHT") = rowSOTINVH1.Item("INV_FREIGHT")
                '.Item("INV_STAX") = rowSOTINVH1.Item("INV_STAX")
                .Item("INV_MISC_CHG") = rowSOTINVH1.Item("INV_MISC_CHG")
                .Item("INV_TOTAL_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")

                ' rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
                ' If it is a Factored Invoice set Balance = 0 if NYA
                If rowSOTINVH1.Item("CUST_FACTOR_IND") & String.Empty = "1" AndAlso (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
                    .Item("INV_BALANCE") = 0
                Else
                    .Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
                End If

                '.Item("REASON_CODE")
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("SALES_DIVISION_CODE") = rowSOTINVH1.Item("SALES_DIVISION_CODE")
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("SREP2_CODE") = rowSOTINVH1.Item("SREP2_CODE")
                .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                .Item("CURR_EXCH_RATE") = 1
                .Item("INV_SALES_CURR") = .Item("INV_SALES")
                .Item("INV_DISC_CURR") = 0
                .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
                '.Item("INV_STAX_CURR") = .Item("INV_STAX")
                .Item("INV_MISC_CHG_CURR") = .Item("INV_MISC_CHG")
                .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT")
                .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
                .Item("INV_NOTES") = rowSOTINVH1.Item("INV_COMMENT")
                .Item("ORDR_TYPE_CODE") = rowSOTINVH1.Item("ORDR_TYPE_CODE")
                '.Item("INV_REF") = rowSOTINVH1.Item("INV_REF")
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            End With

            dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

            INV_NOs.Add(INV_NO)
        Next

        Update_Record_TDA("SOTORDR1")
        Update_Record_TDA("SOTORDR2")

        Update_Record_TDA("SOTORDP1")
        Update_Record_TDA("SOTORDP2")

        Update_Record_TDA("SOTINVH1")
        Update_Record_TDA("SOTINVH2")
        Update_Record_TDA("SOTPICK1")
        Update_Record_TDA("SOTPICK2")
        Update_Record_TDA("SOTSHIP1")
        Update_Record_TDA("ARTOPEN1")

        For Each INV_NO As String In INV_NOs
            ASCMAIN1.sql = "BEGIN SOPSTAT1('I','" & INV_NO & "'); END;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "BEGIN SOPSTAT2_OH_ONLY('I','" & INV_NO & "'); END;"
            ASCDATA1.ExecuteSQL()

            ' When should this be called - Only whses that use Locationss
            'If WHSE_LOCATOR Then
            '    TAC.ICCMAIN1.Update_WHTLOCBX("S", INV_NO)
            'End If

            ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV",
               New Object() {"I", INV_NO},
               New String() {"INV_TYPE_IN", "INV_NO_IN"})
        Next

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
            ASCDATA1.ExecuteSQL()
        Next

        ' Email Invoioces - Put in try catch so if email causes an error the Update still occurs
        Try
            EmailInvoice()
        Catch ex As Exception
            MessageBox.Show("Emailing Invoices caused an error. Data will be updated. Error: " & ex.Message, "Email Ivoices", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Create Web Invoices
        Try
            ASCMAIN1.Progress("Creating Web Invoice", "")
            For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                TAC.SOCMAIN1.CreateWebInvoice(Me, row.Item("INV_TYPE"), row.Item("INV_NO"))
            Next
        Catch ex As Exception

        End Try

    End Sub

    Sub Update_Shipment()

        If ASCMAIN1.CLIENT = "NYA" Then ' PROBABLY SHOULD DO THIS FOR VAN AND RGI TOO
            For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
                rowPOTSHIP2.Item("CBM") = rowPOTSHIP2.Item("TOTAL_CBM")
            Next
        End If

        Dim sqlPOs As String = ""
        If packingFromXLS Or packingFromBooking Then
            dst.Tables("POTORDR1").Clear()
            dst.Tables("POTORDR2").Clear()
            For Each rowPOTORDR1_SPLIT As DataRow In dst.Tables("POTORDR1_SPLIT").Select()
                Dim PO_ORDER_NO As String = rowPOTORDR1_SPLIT.Item("PO_ORDER_NO")
                'If PO_ORDER_NO = "152172" Or PO_ORDER_NO = "151621" Then
                '    Stop
                'End If
                sqlPOs &= "'" & PO_ORDER_NO & "',"
                Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").NewRow
                For i As Int16 = 0 To rowPOTORDR1.ItemArray.Length - 1
                    rowPOTORDR1.Item(i) = rowPOTORDR1_SPLIT.Item(i)
                Next
                rowPOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowPOTORDR1.Item("LAST_DATE") = DATETIME_STAMP
                dst.Tables("POTORDR1").Rows.Add(rowPOTORDR1)

                Dim tblPOTORDR2 As DataTable = dicPOTORDR2(PO_ORDER_NO)
                For Each rowPOTORDR2_SPLIT As DataRow In dst.Tables("POTORDR2_SPLIT").Select("PO_ORDER_NO = '" & PO_ORDER_NO & "'")
                    'If rowPOTORDR2_SPLIT.Item("STYLE_CODE") & "" = "WN2310901" Then
                    '    Stop
                    'End If

                    Dim PO_ORDER_LNO As Integer = Val(rowPOTORDR2_SPLIT.Item("PO_ORDER_LNO") & "")
                    Dim rowOrig As DataRow() = tblPOTORDR2.Select("PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & PO_ORDER_LNO)
                    Dim PO_QTY_SHP_ORIG As Integer = 0
                    If rowOrig.Length = 1 Then
                        PO_QTY_SHP_ORIG = Val(rowOrig(0).Item("PO_QTY_SHP") & "")
                    End If
                    'If PO_ORDER_NO = "152174" And PO_ORDER_LNO = 57 Then
                    '    Stop
                    'End If


                    'Dim WJZSTYLE As String = rowPOTORDR2_SPLIT.Item("STYLE_CODE")
                    'If ASCMAIN1.Running_in_VS And WJZSTYLE = "WM199004" Then Stop

                    Dim PO_ORDER_LNO_ORIG As Integer = Val(rowPOTORDR2_SPLIT.Item("PO_ORDER_LNO_ORIG") & "")
                    Dim sqlPOTORDR2_split As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & PO_ORDER_LNO
                    Dim PO_QTY_SHP As Integer = IIf(PO_ORDER_LNO_ORIG > 0, 0, PO_QTY_SHP_ORIG)
                    Dim PO_QTY_ORD As Integer = Val(rowPOTORDR2_SPLIT.Item("PO_QTY_ORD") & "")
                    Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").NewRow
                    For i As Int16 = 0 To rowPOTORDR2.ItemArray.Length - 1
                        rowPOTORDR2.Item(i) = rowPOTORDR2_SPLIT.Item(i)
                    Next
                    rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_ORD
                    rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP
                    If packingFromXLS Then Throw New Exception("ABS needs to check Update to POTORDR2 if packingFromXLS - which shouldn't be used any more")
                    If packingFromBooking Then ' I THINK WE NEED TO USE 0 EVEN WHEN packingfromXLS - but we should never need to pack from xls any more
                        rowPOTORDR2.Item("PO_QTY_SHP") = 0
                    End If
                    rowPOTORDR2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    rowPOTORDR2.Item("LAST_DATE") = DATETIME_STAMP
                    dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)

                    ' WM199004 & 07 HAVE >0 VALUES IN PO_ORDER_LNO_ORIG, 06 DOES NOT
                    If PO_ORDER_LNO_ORIG > 0 Then
                        If dst.Tables("POTORDR2_SPLIT").Select("PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & PO_ORDER_LNO_ORIG).Length = 0 Then
                            Dim rowPOTORDR2_orig As DataRow = tblPOTORDR2.Select("PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & PO_ORDER_LNO_ORIG)(0)
                            PO_QTY_ORD = Val(rowPOTORDR2_orig.Item("PO_QTY_ORD") & "") - Val(rowPOTORDR2_SPLIT.Item("PO_QTY_ORD") & "")
                            Dim PO_QTY_OPN As Integer = Val(rowPOTORDR2_orig.Item("PO_QTY_OPN") & "") - Val(rowPOTORDR2_SPLIT.Item("PO_QTY_ORD") & "")
                            'IN PPK SHIPMENT BLOWS UP
                            'rowPOTORDR2 = dst.Tables("POTORDR2").NewRow
                            'SO
                            rowPOTORDR2 = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO_ORIG})
                            If rowPOTORDR2 Is Nothing Then
                                rowPOTORDR2 = dst.Tables("POTORDR2").NewRow ' SHOULD NEVER HAPPEN
                                For i As Int16 = 0 To rowPOTORDR2.ItemArray.Length - 1
                                    rowPOTORDR2.Item(i) = rowPOTORDR2_orig.Item(i)
                                Next
                                If packingFromBooking Then
                                    rowPOTORDR2.Item("PO_QTY_SHP") = 0
                                End If
                                dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)
                            End If
                            rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_ORD
                            rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_OPN
                            rowPOTORDR2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            rowPOTORDR2.Item("LAST_DATE") = DATETIME_STAMP
                        End If


                    End If
                Next

                For Each rowPOTORDR2_orig As DataRow In tblPOTORDR2.Select("PO_ORDER_NO = '" & PO_ORDER_NO & "'")
                    Dim PO_ORDER_LNO As Integer = Val(rowPOTORDR2_orig.Item("PO_ORDER_LNO") & "")
                    Dim rowPOTORDR2_split As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                    If rowPOTORDR2_split Is Nothing Then
                        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").NewRow
                        For i As Int16 = 0 To rowPOTORDR2.ItemArray.Length - 1
                            rowPOTORDR2.Item(i) = rowPOTORDR2_orig.Item(i)
                        Next
                        '   "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
                        '    Stop ' look for line in ship3 and then 0 out shipment
                        ' LOOK AT WITH WALT
                        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_ORDER_NO = " & CStr(PO_ORDER_NO) & " AND PO_ORDER_LNO = " & CStr(PO_ORDER_LNO))
                            If rowPOTSHIP3.Item("PO_QTY_SHP") & "" = rowPOTORDR2.Item("PO_QTY_SHP") & "" Then
                                rowPOTORDR2.Item("PO_QTY_OPN") = Val(rowPOTORDR2.Item("PO_QTY_SHP") & "")
                                rowPOTORDR2.Item("PO_QTY_SHP") = 0
                            End If
                        Next

                        'If packingFromBooking Then
                        '    If Val(rowPOTORDR2.Item("PO_QTY_SHP") & "") <> 0 Then
                        '        rowPOTORDR2.Item("PO_QTY_OPN") = Val(rowPOTORDR2.Item("PO_QTY_SHP") & "")
                        '        rowPOTORDR2.Item("PO_QTY_SHP") = 0
                        '    Else
                        '        'Stop
                        '    End If
                        'End If
                        dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)
                    End If
                Next

            Next

            If sqlPOs <> "" Then
                sqlPOs = "PO_ORDER_NO In (" & sqlPOs.TrimEnd(",") & ")"

                Update_Record_TDA("POTORDR1", sqlPOs)
                Update_Record_TDA("POTORDR2", sqlPOs)

            End If
        End If


        ' If Importing XLS packing lists (which means we are in New not Edit), set up pre-packs
        If packingFromXLS Or packingFromBooking Then
            For Each rowPOTSHIP7 As DataRow In dst.Tables("POTSHIP7").Select("CUSTOM_PPK = '1'")
                ' DGJ SELECT ? Select("STYLES>1") OR ("CUSTOM_PPK = '1'" AND CUST_CODE = 'COSCOUS') ' WOULD HAVE TO GET CUST_CODE IN MEMORY
                ' For Each rowPOTSHIP7 As DataRow In dst.Tables("POTSHIP7").Select("STYLES>1")
                ' remove next 3 lines after testing prepacks update
                'If packingFromBooking Then
                '    'Throw New Exception("prepacks not tested yet for bookings")
                'End If
                rowPOTSHIP7.Item("PPK_CODE") = Get_Next_PPK_CODE()
            Next
        End If

        'dst.Tables("WHTPPKM1").Rows.Clear()
        'dst.Tables("WHTPPKM2").Rows.Clear()

        ' Check to see if we need to re-assign PPK Codes

        For Each rowPOTSHIP7 As DataRow In dst.Tables("POTSHIP7").Select("")
            rowPOTSHIP7.Item("PO_QTY_PER_CTN") = Val(rowPOTSHIP7.Item("UNITS") & "")
            rowPOTSHIP7.Item("PPK_INNER_QTY") = Val(rowPOTSHIP7.Item("PPK_INNER_QTY_CALC") & "")

            If rowPOTSHIP7.Item("PPK_CODE") & "" <> "" Then

                Dim PPK_CODE As String = rowPOTSHIP7.Item("PPK_CODE")
                Dim rowWHTPPKM1 As DataRow = Nothing

                ' IF YOU ADD A STYLE TO A CARTON WHICH ALREADY HAD A REAL NON-TMP PPK ASSIGNED, THE PPK CODE IS CHANGED TO TMP SO THAT A NEW WHTPPKM1 RECORD IS CREATED
                ' BUT IF YOU DELETE A STYLE FROM A PPK, THIS DOES NOT HAPPEN
                ' AND IF YOU REVISED THE QTY IN A PPK, THIS DOES NOT HAPPEN
                ' NOT SURE WHY AND WHEN WE NEED TO REGENERATE THE PPK DEFINITION
                ' TOO MANY PPKS WILL CONFUSE SOE WHEN THEY DO PPK SELECTION
                ' PROBABLY NEED TO DO THE TMP THING ONLY IF THE PPK WAS ALREADY SENT
                ' FOR NOW - THE BEHAVIOR IS AS DESCRIBED ABOVE
                ' VAN DOES NOT WANT EXTENSIVE WORK DONE AT THIS TIME - AUG'13 - SO LEAVING THIS ALONE FOR NOW

                Dim PPK_ADO As String = ""


                If PPK_CODE.StartsWith("TMP") Then
                    PPK_CODE = ASCMAIN1.Next_Control_No("PPK_CODE") & "PPK"
                    PPK_CODE = Mid(PPK_CODE, 2)
                    rowPOTSHIP7.Item("PPK_CODE") = PPK_CODE

                    rowWHTPPKM1 = dst.Tables("WHTPPKM1").NewRow
                    rowWHTPPKM1.Item("PPK_CODE") = PPK_CODE
                    rowWHTPPKM1.Item("INIT_DATE") = DATETIME_STAMP
                    rowWHTPPKM1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowWHTPPKM1.Item("PPK_QTY_TOTAL") = 0 ' NULLS NOT ALLOWED
                    dst.Tables("WHTPPKM1").Rows.Add(rowWHTPPKM1)
                Else
                    ' FIX PPK DEFINITION
                    rowWHTPPKM1 = Fill_Record("WHTPPKM1", PPK_CODE, False, False)

                    ASCMAIN1.sql = "Delete from WHTPPKM2 where PPK_CODE = '" & PPK_CODE & "'"
                    ASCDATA1.ExecuteSQL()
                    '  Fill_Records("WHTPPKM2", PPK_CODE, False)
                End If

                For Each rowPOTSHIP8 As DataRow In rowPOTSHIP7.GetChildRows("POTSHIP7_POTSHIP8")
                    Dim rowWHTPPKM2 As DataRow = dst.Tables("WHTPPKM2").NewRow
                    rowWHTPPKM2.Item("PPK_CODE") = PPK_CODE
                    rowWHTPPKM2.Item("STYLE_CODE") = rowPOTSHIP8.Item("STYLE_CODE")
                    rowWHTPPKM2.Item("COLOR_CODE") = rowPOTSHIP8.Item("COLOR_CODE")
                    rowWHTPPKM2.Item("PPK_QTY") = Val(rowPOTSHIP8.Item("QTY") & "") * IIf(rowPOTSHIP8.Item("DOZENS") & "" = "1", 12, 1)
                    dst.Tables("WHTPPKM2").Rows.Add(rowWHTPPKM2)
                Next

                rowWHTPPKM1.Item("PPK_DESC") = rowPOTSHIP7.Item("CARTON_COMMENTS")
                rowWHTPPKM1.Item("LAST_DATE") = DATETIME_STAMP
                rowWHTPPKM1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowWHTPPKM1.Item("CUSTOM_PPK") = rowPOTSHIP7.Item("CUSTOM_PPK")
                rowWHTPPKM1.Item("PPK_QTY_TOTAL") = dst.Tables("WHTPPKM2").Compute("SUM(PPK_QTY)", "PPK_CODE = '" & PPK_CODE & "'")

                rowPOTSHIP7.Item("STYLE_CODE") = ""
                rowPOTSHIP7.Item("COLOR_CODE") = ""
            Else
                rowPOTSHIP7.Item("STYLE_CODE") = rowPOTSHIP7.Item("STYLE_CODE_1")
                rowPOTSHIP7.Item("COLOR_CODE") = rowPOTSHIP7.Item("COLOR_CODE_1")
            End If

            If WH_REC_NOsInProcess.Count > 0 Then
                If rowPOTSHIP7.RowState = DataRowState.Added Then
                    Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP7.Item("PO_SHIPMENT_LNO") & "")
                    Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
                    Dim WH_REC_NO As String = rowPOTSHIP2.Item("WH_REC_NO") & ""
                    Add_WHTWREC7(rowPOTSHIP7, WH_REC_NO, PO_SHIPMENT_LNO)
                    For Each rowPOTSHIP8 As DataRow In rowPOTSHIP7.GetChildRows("POTSHIP7_POTSHIP8")
                        Add_WHTWREC8(rowPOTSHIP8, WH_REC_NO, PO_SHIPMENT_LNO)
                    Next
                End If
            End If

        Next

        If WH_REC_NOsInProcess.Count > 0 Then
            Update_Record_TDA("WHTWREC7")
            Update_Record_TDA("WHTWREC8")
        End If

        Update_Record_TDA("WHTPPKM1")
        Update_Record_TDA("WHTPPKM2")

        If EntryMode <> "N" Then
            Dependent_Updates(-1, PO_SHIPMENT_NO)
        End If


        Dim sqlx As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
        INIT_LAST("POTSHIP1", False, sqlx, True)
        INIT_LAST("POTSHIP2", False, sqlx, True)
        INIT_LAST("POTSHIP3", False, sqlx, True)
        INIT_LAST("POTSHIP4", False, sqlx, True)

        Update_Record_TDA("POTSHIP1", sqlx)
        Update_Record_TDA("POTSHIP2", sqlx)
        Update_Record_TDA("POTSHIP3", sqlx)
        Update_Record_TDA("POTSHIP4", sqlx)
        ' POTSHIP5?
        Update_Record_TDA("POTSHIP7", sqlx)
        Update_Record_TDA("POTSHIP8", sqlx)

        Dependent_Updates(1, PO_SHIPMENT_NO)

        If packingFromBooking Then
            Update_Record_TDA("POTVBKG1")
            Update_Record_TDA("POTVBKG2")
            Update_Record_TDA("POTPACK2")
            Update_Record_TDA("POTPACK3")

            For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select("", "")
                Dim VBKG_NO As String = rowPOTVBKG2.Item("VBKG_NO")
                Dim PACK_LIST_NO As String = rowPOTVBKG2.Item("PACK_LIST_NO")
                ASCMAIN1.sql = "Update POTLPNL1 Set PO_SHIPMENT_NO = :PARM1, PO_SHIPMENT_LNO = " & vbCrLf _
                    & " (Select PO_SHIPMENT_LNO from POTVBKG2 where VBKG_NO = :PARM2 and PACK_LIST_NO = :PARM3)" & vbCrLf _
                    & " where BARCODE_STATUS = 'A' and PACK_LIST_NO = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VNV", New Object() {PO_SHIPMENT_NO, VBKG_NO, PACK_LIST_NO})

                Dim rowPOTPACK1 As DataRow = LookUp("POTPACK1", PACK_LIST_NO)
                Dim INITIAL_ORDER As String = rowPOTPACK1.Item("INITIAL_ORDER") & ""

                ASCMAIN1.sql = "Update POTLPNL1 Set CARTON_NO = " & vbCrLf _
                        & " (Select CARTON_NO from " & IIf(INITIAL_ORDER = "1", "POTPACK2", "POTPACK3") & vbCrLf _
                        & " where PACK_LIST_NO = :PARM1" & vbCrLf _
                        & " and PACK_LIST_SHEET_NO = POTLPNL1.PACK_LIST_SHEET_NO" & vbCrLf _
                        & IIf(INITIAL_ORDER = "1", "", " and PACK_LIST_SHEET_LNO = POTLPNL1.PACK_LIST_SHEET_LNO" & vbCrLf) _
                        & ")" & vbCrLf _
                        & " where BARCODE_STATUS = 'A' and PACK_LIST_NO = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {PACK_LIST_NO})

            Next
        End If

        If chkFixCasePacks.Checked Then
            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Q NUMBER(8,0); Cursor C1 is" & vbCrLf _
                & "  Select X.*, ICTSTYL1.CARTON_PACK_QTY from (Select PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, PO_QTY_PER_CTN, SUM (CARTONS) CARTONS" & vbCrLf _
                & "   from POTSHIP7" & vbCrLf _
                & "   where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                & "     and PPK_CODE IS NULL" & vbCrLf _
                & "   group by PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, PO_QTY_PER_CTN) X,ICTSTYL1" & vbCrLf _
                & "   where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   order by CARTONS;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Select CARTON_PACK_QTY into Q from ICTSTYL1 where STYLE_CODE = R1.STYLE_CODE;" & vbCrLf _
                & "   If NVL(Q,0) <> NVL(R1.PO_QTY_PER_CTN,0) Then " & vbCrLf _
                & "    Update ICTSTYL1 Set CARTON_PACK_QTY = R1.PO_QTY_PER_CTN" & vbCrLf _
                & "     where STYLE_CODE = R1.STYLE_CODE;" & vbCrLf _
                & "    Insert into ASTAUDT1" & vbCrLf _
                & "     (TABLE_NAME,KEY_VALUE,COLUMN_NAME,USER_ID,INIT_DATE,OLD_VALUE,NEW_VALUE,FM_MODE,NOTES,KEY_VALUE2,KEY_LNO,SESSION_NO,SELECTION_NO,XNO)" & vbCrLf _
                & "     Values " & vbCrLf _
                & "     ('ICTSTYL1',R1.STYLE_CODE,'CARTON_PACK_QTY','" & ASCMAIN1.USER_ID & "',SYSDATE,Q,R1.PO_QTY_PER_CTN,'E','Shipment Fix CARTON_PACK_QTY'" & vbCrLf _
                & "      ,R1.PO_SHIPMENT_NO,R1.PO_SHIPMENT_LNO,'" & ASCMAIN1.SESSION_NO & "'," & CStr(Me.SELECTION_NO) & ",'" & Me.XNO & "');" & vbCrLf _
                & "   End If;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

        End If

        If ASCMAIN1.CLIENT = "NYA" AndAlso rowPOTSHIP1.Item("WHSE_CODE") & "" = "18" Then

            ASCMAIN1.sql = "Select POTSHIP7.STYLE_CODE, POTSHIP7.CARTON_VOLUME, POTSHIP7.CARTON_WEIGHT, POTSHIP7.PO_QTY_PER_CTN" & vbCrLf _
                & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_WEIGHT_GRS, ICTSTYL1.CASE_CUBE" & vbCrLf _
                & " from POTSHIP7, ICTSTYL1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = POTSHIP7.STYLE_CODE" & vbCrLf _
                & "   and POTSHIP7.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = Fill_Record("ICTSTYL1", STYLE_CODE)

                'Dim PO_QTY_PER_CTN As Int32 = Val(row.Item("PO_QTY_PER_CTN") & "")
                'If Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "") <> PO_QTY_PER_CTN And PO_QTY_PER_CTN <> 0 Then
                '    rowICTSTYL1.Item("CARTON_PACK_QTY") = PO_QTY_PER_CTN
                'End If

                Dim CARTON_WEIGHT As Decimal = System.Math.Round(Val(row.Item("CARTON_WEIGHT") & "") * 2.2, 1)
                ' KG TO LBS
                If Val(rowICTSTYL1.Item("CASE_WEIGHT_GRS") & "") <> CARTON_WEIGHT And CARTON_WEIGHT <> 0 Then
                    rowICTSTYL1.Item("CASE_WEIGHT_GRS") = CARTON_WEIGHT
                End If

                Dim CARTON_VOLUME As Decimal = System.Math.Round(Val(row.Item("CARTON_VOLUME") & "") * 0.0610237, 1)
                ' CBM TO CUIN
                If Val(rowICTSTYL1.Item("CASE_CUBE") & "") <> CARTON_VOLUME And CARTON_VOLUME <> 0 Then
                    rowICTSTYL1.Item("CASE_CUBE") = CARTON_VOLUME
                End If

                If rowICTSTYL1.RowState = DataRowState.Modified Then
                    Write_Audit_Trail(rowICTSTYL1, "E")
                    Update_Record_TDA("ICTSTYL1")
                End If

            Next
        End If

        If chkFinalize.Checked Then
            Release_Shipment_Send_3PL() ' OR JUST RELEASE TO WAREHOUSE FOR RECEIPT IF CASE CONTROL IS SET
        End If

        If AT_Packing Then
            For Each rowINVHDR As DataRow In rowATSHIPS.GetChildRows("ATSHIPS_ATINVHDR")
                Dim VAN_REF As String = rowINVHDR.Item("VAN_REF")
                ASCMAIN1.sql = "Update POTIHDRA Set STATUS = 'U' where VAN_REF = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {VAN_REF})
            Next

        End If

        If ASCMAIN1.CLIENT = "VAN" AndAlso POTVBKG2_RECORDS Then
            BOOKING_INTEGRITY(PO_SHIPMENT_NO)
        End If

    End Sub

    Sub Add_WHTWREC7(row As DataRow, WH_REC_NO As String, PO_SHIPMENT_LNO As Integer)
        ' ALMOST identical copy of this method exists in POFSHIP1 - probably should refactor both to POCMAIN1
        Dim rowWHTWREC7 As DataRow = dst.Tables("WHTWREC7").NewRow
        With rowWHTWREC7
            .Item("WH_REC_NO") = WH_REC_NO
            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            .Item("CARTON_NO") = row.Item("CARTON_NO")
            .Item("CARTONS") = Val(row.Item("CARTONS") & "")
            .Item("CARTON_COMMENTS") = row.Item("CARTON_COMMENTS") & ""
            .Item("CUSTOM_PPK") = row.Item("CUSTOM_PPK") & ""
            .Item("PPK_CODE") = row.Item("PPK_CODE") & ""
            .Item("PO_QTY_PER_CTN") = Val(row.Item("PO_QTY_PER_CTN") & "")
            .Item("STYLE_CODE") = row.Item("STYLE_CODE") & ""
            .Item("COLOR_CODE") = row.Item("COLOR_CODE") & ""
            .Item("PPK_INNER_QTY") = Val(row.Item("PPK_INNER_QTY") & "")
            .Item("CARTON_DIMS") = row.Item("CARTON_DIMS") & ""
            .Item("CARTON_VOLUME") = Val(row.Item("CARTON_VOLUME") & "")
            .Item("CARTON_WEIGHT") = Val(row.Item("CARTON_WEIGHT") & "")
            .Item("CARTON_NO_CLONED_FROM") = DBNull.Value
        End With
        dst.Tables("WHTWREC7").Rows.Add(rowWHTWREC7)
    End Sub

    Sub Add_WHTWREC8(row As DataRow, WH_REC_NO As String, PO_SHIPMENT_LNO As Integer)
        ' ALMOST identical copy of this method exists in POFSHIP1 - probably should refactor both to POCMAIN1
        Dim rowWHTWREC8 As DataRow = dst.Tables("WHTWREC8").NewRow
        With rowWHTWREC8
            .Item("WH_REC_NO") = WH_REC_NO
            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            .Item("CARTON_NO") = row.Item("CARTON_NO")
            .Item("STYLE_CODE") = row.Item("STYLE_CODE") & ""
            .Item("COLOR_CODE") = row.Item("COLOR_CODE") & ""
            .Item("QTY") = Val(row.Item("QTY") & "")
            .Item("DOZENS") = row.Item("DOZENS") & ""
            .Item("PPK_INNER_QTY") = Val(row.Item("PPK_INNER_QTY") & "")
            .Item("QTY_SHP") = Val(row.Item("QTY") & "")
        End With
        dst.Tables("WHTWREC8").Rows.Add(rowWHTWREC8)
    End Sub


    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("PO_SHIPMENT_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTSHIP1"
            E.COLUMN_NAME = "PO_SHIPMENT_NO"
            E.CODE_VALUE = PO_SHIPMENT_NO ' HFs("CUST_CODE")
            E.DESC_VALUE = Absx1.txtFor("PO_SHIP_VESSEL").Text & ":" & Absx1.txtFor("PO_SHIP_REF_NO").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "POTSHIP1"
        E.TABLE_KEY_CAPTION = "PO Shipment"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PO_SHIPMENT_NO").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("VESSEL_NAME").Text
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
            Case "PO_SHIPMENT_NO"
                If InquiryMode Then
                Else
                    sql_where &= "PO_SHIPMENT_NO in (Select DISTINCT PO_SHIPMENT_NO from POTSHIP2 where PO_SHIP_STATUS = 'O') "
                End If
            Case "CUST_STORE_NO"
                sql_where &= "CUST_CODE = '171659' AND CUST_ADDR_TYPE = 'MK' AND CUST_ADDR_STATUS = 'A'"
        End Select
    End Sub

    Public Function EmailInvoice() As Boolean

        Dim INV_NO As String = String.Empty
        Dim attachFileName As String = String.Empty

        Try
            If Not (ASCMAIN1.DBS_SERVER = "RGI" AndAlso ASCMAIN1.DBS_COMPANY = "RGI") Then
                Return True
            End If

            If dst.Tables("SOTINVH1").Rows.Count = 0 Then
                Return False
            End If

            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows(0)
            INV_NO = rowSOTINVH1.Item("INV_NO")
            Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO")
            Dim SREP_CODE As String = rowSOTINVH1.Item("SREP_CODE") & String.Empty
            Dim salesRepEmail As String = String.Empty
            Dim CUST_XMIT_INV_VIA As String = String.Empty

            Dim CONS_INV As String = "0"
            If rowSOTINVH1.Item("INV_NO_CONS") & String.Empty <> String.Empty Then
                CONS_INV = "1"
            End If

            ' See if the customer receives an acknowledgment
            Dim rowSOTSREP1 As DataRow = LookUp("SOTSREP1", SREP_CODE)
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

            If rowSOTSREP1 Is Nothing AndAlso rowARTCUST1 Is Nothing Then
                Return False
            End If

            ' See if we have anyone to email to
            If rowSOTSREP1 IsNot Nothing AndAlso rowSOTSREP1.Item("SREP_EMAIL") & String.Empty <> String.Empty Then
                salesRepEmail = rowSOTSREP1.Item("SREP_EMAIL") & String.Empty
            End If

            If rowARTCUST1 IsNot Nothing Then
                ' Mail, Email, Both
                CUST_XMIT_INV_VIA = (rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty).ToString.Trim
                If CUST_XMIT_INV_VIA.Length > 0 AndAlso "EB".Contains(CUST_XMIT_INV_VIA) Then
                    If rowARTCUST1.Item("CUST_INV_EMAIL") & String.Empty <> String.Empty Then
                        salesRepEmail &= ";" & rowARTCUST1.Item("CUST_INV_EMAIL") & String.Empty
                    End If

                    If rowARTCUST1.Item("CUST_INV_CC") & String.Empty <> String.Empty Then
                        salesRepEmail &= ";" & rowARTCUST1.Item("CUST_INV_CC") & String.Empty
                    End If
                End If
            End If

            ' remove double semi-colons
            salesRepEmail = salesRepEmail.Replace(",", ";")
            salesRepEmail = salesRepEmail.Replace(";;", ";")
            salesRepEmail = salesRepEmail.Replace(" ", "")

            ' should be at least 5 characters
            If salesRepEmail.Replace(";", "").Trim.Length < 5 Then
                Return False
            End If

            attachFileName = rowARTCUST1.Item("CUST_NAME") & " " & INV_NO

            For Each invalidChar As String In New String() {"\", "/", ":", "*", "?", "<", ">", "|", "."}
                attachFileName = attachFileName.Replace(invalidChar, "")
            Next
            attachFileName = attachFileName.Replace(" ", "_")


            Dim invNos As String = String.Empty
            For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                invNos &= ", '" & row.Item("INV_NO") & "'"
            Next
            invNos = invNos.Substring(1).Trim

            ASCMAIN1.Progress("Sending Sales Rep / Customer a copy of the Invoice", "")

            Dim RPT As String = "SORINVP1"
            If Not REPORTS.ContainsKey(RPT) Then
                REPORTS.Add(RPT, Load_rptClass(RPT))
                REPORTS(RPT).Prepare_dst(False, "")
            End If

            REPORTS(RPT).Fill_Records_RPT(New String() {" and SOTINVH1.INV_NO IN (" & invNos & ")"})

            Dim REPORT_NO As String = String.Empty
            With REPORTS(RPT).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", CONS_INV)
                .CR_params.Add("EXPORT_INFO", "0")

                ' Set the customers Invoice
                Select Case ASCMAIN1.DBS_COMPANY
                    Case "RGI"
                        RPT = "SORINVPR"

                End Select

                REPORT_NO = .Generate_Report(RPT, "Invoice", , True, , , "PDF", attachFileName, False)
                .Print_Report_End(True, True)
            End With

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(attachFileName & ".pdf", ASCMAIN1.Folders("Temp") & attachFileName & ".pdf")

            Dim SUBJECT As String = String.Empty
            SUBJECT = "Sales Invoice (" & INV_NO & ") for customer " & rowARTCUST1.Item("CUST_NAME")

            ' Concatentate and process all email addresses
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each emailAddress As String In (salesRepEmail).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Return True
            End If

            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    SUBJECT, IIf(ASCMAIN1.DBS_COMPANY = "RGI", "AUTOINV", "INV"),
                    True, False, CUST_CODE, rowARTCUST1.Item("CUST_NAME"), "Customer")


            ' Mark email Only Invoices as Mailed
            Try
                If CUST_XMIT_INV_VIA = "E" Then
                    For Each rowSOTINVH1 In dst.Tables("SOTINVH1").Rows
                        INV_NO = rowSOTINVH1.Item("INV_NO")
                        ASCDATA1.ExecuteSQL("Update SOTINVH1 Set INV_PRINTED = SYSDATE where INV_NO = '" & INV_NO & "'")
                    Next
                End If
            Catch ex As Exception
                ' nothing 
            End Try

            EmailInvoice = True

        Catch ex As Exception
            EmailInvoice = False
        End Try

    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTSHIP3, "BBBBBB", "PO Inquiry", "Style Status Inquiry", "Style Master File", "Move Receiving Shortage to New Shipment Line",
                        "Allow Change to Qty Received", "Switch PO & Line", "Add PO Line")
        Load_Popup_Menu(grdPOTSHIP5, "B", "Voucher Inquiry")
        Load_Popup_Menu(grdPOTSHIPX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Shipment Inquiry")
        Load_Popup_Menu(grdPOTSHIPS, "S", "Show Filter")
        Load_Popup_Menu(grdPOTSHIPI, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Shipment Inquiry", "Style Status Inquiry")
        Load_Popup_Menu(grdPOTSHIPF, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Shipment Inquiry")
        Load_Popup_Menu(grdPOTSHIPC, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Shipment Inquiry")
        Load_Popup_Menu(grdPOTSHIP4, "B", "Create Containers from BOL Data")
        Load_Popup_Menu(grdEDT944T1, "B", "Delete 944 from Queue")
        Load_Popup_Menu(grdPOTSHIP7, "BBBBB",
                        "Copy Carton Dimensions to All Cartons Referencing PO",
                        "Copy Carton Dimensions to All Cartons Referencing PO (only if blank)",
                        "Calculate Cartons as Qty Shipped / Carton Pack",
                        "Calculate Qty/Carton Proportionately",
                        "Add Remaining Units")
        Load_Popup_Menu(grdPOTSHIPR, "BBBBS",
                     "Create a carton type containing all selected Style/Colors",
                     "Create an individual carton type for All Style/Colors",
                     "Create an individual carton type for each selected Style/Color",
                     "Add Style to Carton", "Show Multi-Pack")

        Load_Popup_Menu(grdPOTSHIP8, "S", "Show All Carton Details")
        Load_Popup_Menu(grdPOTSHIP2, "SSBBSBB", "Show Filter", "Show GroupBox", "Pro-Forma Invoice", "Sales Order Inquiry",
                        "FIFO Previous Period", "Import Packing List from XLS", "Consolidate Inv/BOL")
        Load_Popup_Menu(grdICTIRECX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Shipment Inquiry")

        Load_Popup_Menu(grdAPTINVH1, "B", "Voucher Inquiry")

        Load_Popup_Menu(grdAPTCHCKV, "SSB", "Show Filter", "Show GroupBox", "Vendor Inquiry")
        Load_Popup_Menu(grdAPTCHCKP, "SSBB", "Show Filter", "Show GroupBox", "Vendor Inquiry", "Voucher Inquiry")
        Load_Popup_Menu(grdAPTCHCKQ, "SSBB", "Show Filter", "Show GroupBox", "Vendor Inquiry", "Voucher Inquiry")

        If ASCMAIN1.CLIENT = "VAN" Then
            Load_Popup_Menu(grdWHT3PLR1, "B", "Mark as Deleted")
            Load_Popup_Menu(grdATSHIPS, "B", "Delete AT Transmission")
            Load_Popup_Menu(grdPOTSHIP5_ALL, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Shipment Inquiry")
        End If

        Load_Popup_Menu(grdPOTPACKG, "B", "Split Balance Open to New Line")
        Load_Popup_Menu(grdPOTPACKR, "BBB", "PO Inquiry", "Style Status Inquiry", "Style Master File")
        Load_Popup_Menu(grdPOTSHPIE, "BBBB", "PO Inquiry", "Style Status Inquiry", "Style Master File")

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        Try
            grd = GRDs(Mid(e.SourceControl.Name, 4))

        Catch ex As Exception

        End Try
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

            Case "grdATSHIPS"
                tlb_btn = DirectCast(tlb.Tools("Delete AT Transmission"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = ship_entry And grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Band.Index = 0)


            Case "grdPOTSHIP2"
                tlb_btn = DirectCast(tlb.Tools("Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = receipt_mode And grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("PO_SHIP_STATUS").Value = "X" And grd.ActiveRow.Cells("ORDR_NO").Value & "" <> "")
                tlb_btn = DirectCast(tlb.Tools("Sales Order Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("ORDR_NO").Value & "" <> "")
                tlb_sbt = DirectCast(tlb.Tools("FIFO Previous Period"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = cost_calc And EntryMode = "E" AndAlso grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("OPS_YYYYPP").Value & "" <> "")
                If tlb_sbt.SharedProps.Visible Then
                    tlb_sbt.Checked = (grd.ActiveRow.Cells("OPS_YYYYPP_FIFO").Value & "" <> "")
                End If
                tlb_btn = DirectCast(tlb.Tools("Import Packing List from XLS"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Packing_Import_Allowed()

                tlb_sbt = DirectCast(tlb.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = cost_calc
                tlb_sbt = DirectCast(tlb.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = cost_calc

                tlb_btn = DirectCast(tlb.Tools("Consolidate Inv/BOL"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ship_entry) And (EntryMode = "E") And Not (WH_REC_NOsInProcess.Count > 0)
                tlb_btn.SharedProps.Visible = False ' until we add logic to disable this for POTSHIP2 records created from bookings import

            Case "grdPOTSHIP3"
                tlb_btn = DirectCast(tlb.Tools("Move Receiving Shortage to New Shipment Line"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ship_entry) And (EntryMode = "E") And Not (WH_REC_NOsInProcess.Count > 0) _
                    And grd.ActiveRow IsNot Nothing AndAlso Val(grd.ActiveRow.Cells("PO_QTY_VAR").Value & "") <> 0 _
                    And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")

                tlb_btn = DirectCast(tlb.Tools("Allow Change to Qty Received"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (receipt_mode) And (EntryMode = "E") And Not (WH_REC_NOsInProcess.Count > 0) _
                    And grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Cells("PO_QTY_REC").Column.CellActivation = UltraWinGrid.Activation.NoEdit _
                    And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")

                tlb_btn = DirectCast(tlb.Tools("Switch PO & Line"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = ASCMAIN1.Running_in_VS And Not (WH_REC_NOsInProcess.Count > 0)

                tlb_btn = DirectCast(tlb.Tools("Add PO Line"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = ASCMAIN1.Running_in_VS And Not (WH_REC_NOsInProcess.Count > 0)

            Case "grdPOTPACKG"
                tlb_btn = DirectCast(tlb.Tools("Split Balance Open to New Line"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Index = 0 And grd.Selected.Rows.Count < 2 _
                    And grd.ActiveRow.Cells("PO_QTY_PACK").Value > 0 And grd.ActiveRow.Cells("PO_QTY_BAL").Value > 0)

            Case "grdPOTPACKR"
                tlb_btn = DirectCast(tlb.Tools("PO Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Index = 1)

            Case "grdPOTSHIPR"
                If EntryMode = "V" Or cost_calc Then e.Cancel = True

                tlb_btn = DirectCast(tlb.Tools("Add Style to Carton"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Manual_Entry_Allowed() And Not (WH_REC_NOsInProcess.Count > 0)

                If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") Then
                    tlb_btn = DirectCast(tlb.Tools("Create a carton type containing all selected Style/Colors"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Manual_Entry_Allowed() And Not (WH_REC_NOsInProcess.Count > 0)

                    tlb_btn = DirectCast(tlb.Tools("Create an individual carton type for All Style/Colors"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Manual_Entry_Allowed() And Not (WH_REC_NOsInProcess.Count > 0)

                    tlb_btn = DirectCast(tlb.Tools("Create an individual carton type for each selected Style/Color"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Manual_Entry_Allowed() ' And Not (WH_REC_NOsInProcess.Count > 0)
                End If

                tlb_sbt = DirectCast(tlb.Tools("Show Multi-Pack"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = Not cost_calc And (EntryMode = "E" Or EntryMode = "N")

                tlb_sbt.Tag = "X"
                tlb_sbt.Checked = Not splCartonQ.Panel2Collapsed
                tlb_sbt.Tag = ""

            Case "grdPOTSHIP4"
                'If EntryMode = "V" Or cost_calc Then e.Cancel = True
                If EntryMode = "V" Then e.Cancel = True
                'tlb_btn = DirectCast(tlb.Tools("Create Containers from BOL Data"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") And Not POTVBKG2_RECORDS

            Case "grdPOTSHIP7"
                If cost_calc Then
                    e.Cancel = True
                    Exit Sub
                End If
                tlb_btn = DirectCast(tlb.Tools("Copy Carton Dimensions to All Cartons Referencing PO"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" Or EntryMode = "E") And ScreenMode And ship_entry And Not (WH_REC_NOsInProcess.Count > 0)
                tlb_btn = DirectCast(tlb.Tools("Copy Carton Dimensions to All Cartons Referencing PO (only if blank)"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" Or EntryMode = "E") And ScreenMode And ship_entry And Not (WH_REC_NOsInProcess.Count > 0)

                tlb_btn = DirectCast(tlb.Tools("Calculate Qty/Carton Proportionately"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" Or EntryMode = "E") And ScreenMode And ship_entry And Not (WH_REC_NOsInProcess.Count > 0)
                tlb_btn = DirectCast(tlb.Tools("Add Remaining Units"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" Or EntryMode = "E") And ScreenMode And ship_entry And Not (WH_REC_NOsInProcess.Count > 0)

                tlb_btn = DirectCast(tlb.Tools("Calculate Cartons as Qty Shipped / Carton Pack"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not (WH_REC_NOsInProcess.Count > 0)

            Case "grdPOTSHIP8"
                If cost_calc Then
                    e.Cancel = True
                    Exit Sub
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdPOTSHIP3"
                '    ' If EntryMode = "V" Then e.Cancel = True
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.OwningMenu Is Nothing OrElse Not GRDs.ContainsKey(Mid(e.Tool.OwningMenu.Key, 4)) Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Delete AT Transmission"

                Dim SD As String = grdATSHIPS.ActiveRow.Cells("ShipDate").Value & ""
                Dim SHIPDATE As Date = Nothing

                If SD <> "" Then
                    SHIPDATE = grdATSHIPS.ActiveRow.Cells("ShipDate").Value
                    SD = Format(SHIPDATE, "MM/dd/yyyy")
                End If
                Dim CARRIER As String = grdATSHIPS.ActiveRow.Cells("Carrier").Value & ""
                If MsgBox("Are you sure you want to delete this Transission: " & vbCrLf & SD & ":" & CARRIER,
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

                BeginTrans()
                Dim R As Integer = 0

                ASCMAIN1.sql = "Update VAN.POTIHDRA Set STATUS = 'D'" & vbCrLf _
                    & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "', LAST_DATE = SYSDATE" & vbCrLf _
                    & " where VAN_REF IN (" & vbCrLf _
                    & " Select V.VAN_REF FROM VAN.POTIHDRA V,AT.`invhdr` A" & vbCrLf _
                    & " where V.VAN_REF = A.VAN_REF AND V.STATUS = 'W'" & vbCrLf _
                    & "   and NVL(A.`Carrier`,'?') = NVL(:PARM1,'?')"

                If SD = "" Then
                    ASCMAIN1.sql &= "   and A.`ShipDate` is null)"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    R = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {CARRIER})
                Else
                    ASCMAIN1.sql &= "   and A.`ShipDate` = :PARM2)"
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                    R = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VD", New Object() {CARRIER, SHIPDATE})
                End If

                CommitTrans()

                MsgBox("Transmission Deleted", MsgBoxStyle.OkOnly, "Confirmation")

                Setup_AT_Shipments()

            Case "Split Balance Open to New Line"
                If grd.Selected.Rows.Count = 0 Then grd.ActiveRow.Selected = True

                Dim PO_SHIPMENT_LNO As Integer = 0

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If Not ASCMAIN1.Logical_Lock("POTSHIP1", grow.Cells("PO_SHIPMENT_NO").Value) Then Exit Sub

                    PO_SHIPMENT_NO = grow.Cells("PO_SHIPMENT_NO").Value
                    EnforceConstraints(False)
                    Fill_Records("POTSHIP1", grow.Cells("PO_SHIPMENT_NO").Value, False)
                    Fill_Records("POTSHIP2", grow.Cells("PO_SHIPMENT_NO").Value, False)
                    Fill_Records("POTSHIP3", New Object() {grow.Cells("PO_SHIPMENT_NO").Value}, False)
                    Fill_Records("POTSHIP7", grow.Cells("PO_SHIPMENT_NO").Value, False)
                    Fill_Records("POTSHIP8", grow.Cells("PO_SHIPMENT_NO").Value, False)

                    Fill_Records("POTORDRO", PO_SHIPMENT_NO, False)

                    dst.Tables("POTSHIPR").Rows.Clear()
                    For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3"), New String() {"PO_SHIPMENT_LNO", "STYLE_CODE", "COLOR_CODE"}).Rows
                        Create_POTSHIPR(row.Item("PO_SHIPMENT_LNO"), row.Item("STYLE_CODE"), row.Item("COLOR_CODE"))
                    Next

                    Dim PO_QTY_PACK As Int64 = Val(grow.Cells("PO_QTY_PACK").Value & "")
                    Dim PO_QTY_BAL As Int64 = Val(grow.Cells("PO_QTY_BAL").Value & "")

                    If PO_QTY_BAL > 0 Then

                        PO_SHIPMENT_LNO = Val(dst.Tables("POTSHIP2").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1

                        Dim row2 As DataRow = dst.Tables("POTSHIP2").Rows.Find _
                                                (New Object() {grow.Cells("PO_SHIPMENT_NO").Value,
                                                                grow.Cells("PO_SHIPMENT_LNO").Value})

                        Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").NewRow
                        rowPOTSHIP2.ItemArray = row2.ItemArray
                        rowPOTSHIP2.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                        rowPOTSHIP2.Item("PO_SHIP_STATUS") = "O"
                        rowPOTSHIP2.Item("TRAN_NO") = DBNull.Value
                        rowPOTSHIP2.Item("OPS_YYYYPP") = DBNull.Value
                        rowPOTSHIP2.Item("OPS_YYYYPP_FIFO") = DBNull.Value
                        rowPOTSHIP2.Item("PO_SOURCE_DOC") = DBNull.Value
                        rowPOTSHIP2.Item("PO_DATE_RECEIVED") = DBNull.Value
                        rowPOTSHIP2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowPOTSHIP2.Item("LAST_DATE") = DATETIME_STAMP
                        rowPOTSHIP2.Item("PO_SHIP_CTNS") = 0
                        'PO_SHIP_CTNS
                        'TOTAL_WEIGHT()
                        'TRUCKING()
                        'rowPOTSHIP2.Item("ORDR_NO") = DBNull.Value

                        dst.Tables("POTSHIP2").Rows.Add(rowPOTSHIP2)

                        For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                            Dim STYLE_CODE As String = grow2.Cells("STYLE_CODE").Value
                            Dim COLOR_CODE As String = grow2.Cells("COLOR_CODE").Value
                            Dim PO_QTY_SHP As Integer = grow2.Cells("PO_QTY_BAL").Value
                            PO_QTY_PACK = grow2.Cells("PO_QTY_PACK").Value

                            Dim row3 As DataRow = dst.Tables("POTSHIP3").Rows.Find _
                                                 (New Object() {grow2.Cells("PO_SHIPMENT_NO").Value,
                                                                grow2.Cells("PO_SHIPMENT_LNO").Value,
                                                                grow2.Cells("PO_ORDER_NO").Value,
                                                                grow2.Cells("PO_ORDER_LNO").Value})

                            Dim rowPOTSHIP3 As DataRow = dst.Tables("POTSHIP3").NewRow
                            rowPOTSHIP3.ItemArray = row3.ItemArray
                            rowPOTSHIP3.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            rowPOTSHIP3.Item("PO_QTY_SHP") = PO_QTY_SHP
                            rowPOTSHIP3.Item("PO_QTY_REC") = 0
                            rowPOTSHIP3.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            rowPOTSHIP3.Item("LAST_DATE") = DATETIME_STAMP



                            row3("PO_QTY_SHP") = Val(row3("PO_QTY_SHP") & "") - PO_QTY_SHP
                            grow.Update()

                            Create_POTSHIPR(PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE)

                            If rowPOTSHIP3.Item("PO_QTY_SHP") > 0 Then

                                dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)

                                Dim row7 As DataRow = dst.Tables("POTSHIP7").Select($"PO_SHIPMENT_NO = '{grow2.Cells("PO_SHIPMENT_NO").Value}' " &
                                                                                $"and PO_SHIPMENT_LNO = '{grow2.Cells("PO_SHIPMENT_LNO").Value}'" &
                                                                                $" and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'").FirstOrDefault
                                row7.Item("CARTONS") = (PO_QTY_PACK / Val(row7("PO_QTY_PER_CTN")))
                                row2.Item("PO_SHIP_CTNS") = row2.Item("PO_SHIP_CTNS") - (PO_QTY_SHP / Val(row7("PO_QTY_PER_CTN")))

                                Dim rowPOTSHIP7 As DataRow = dst.Tables("POTSHIP7").NewRow
                                rowPOTSHIP7.ItemArray = row7.ItemArray
                                rowPOTSHIP7.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                                rowPOTSHIP7.Item("CARTONS") = (PO_QTY_SHP / Val(row7("PO_QTY_PER_CTN")))
                                rowPOTSHIP2.Item("PO_SHIP_CTNS") = rowPOTSHIP2.Item("PO_SHIP_CTNS") + (PO_QTY_SHP / Val(row7("PO_QTY_PER_CTN")))

                                dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)

                                Dim row8 As DataRow = dst.Tables("POTSHIP8").Select($"PO_SHIPMENT_NO = '{grow2.Cells("PO_SHIPMENT_NO").Value}' " &
                                                                                $"and PO_SHIPMENT_LNO = '{grow2.Cells("PO_SHIPMENT_LNO").Value}'" &
                                                                                $" and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'").FirstOrDefault

                                Dim rowPOTSHIP8 As DataRow = dst.Tables("POTSHIP8").NewRow
                                rowPOTSHIP8.ItemArray = row8.ItemArray
                                rowPOTSHIP8.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO

                                dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)

                            End If


                        Next
                    End If
                Next
                EnforceConstraints(True)

                If MsgBox("New PO Shipment line has been created, Continue to Save changes?",
          MsgBoxStyle.YesNo, "Please Confirm Action") = MsgBoxResult.Yes Then

                    BeginTrans()
                    Update_Record_TDA("POTSHIP2")
                    Update_Record_TDA("POTSHIP3")
                    Update_Record_TDA("POTSHIP7")
                    Update_Record_TDA("POTSHIP8")
                    CommitTrans("New Shipment Line created")

                    Setup_PackingSLips()
                End If
                grd.Selected.Rows.Clear()
                PO_SHIPMENT_NO = ""

            Case "Move Receiving Shortage to New Shipment Line"
                If grd.Selected.Rows.Count = 0 Then grd.ActiveRow.Selected = True

                Dim PO_SHIPMENT_LNO As Integer = 0

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows

                    Dim PO_QTY_SHP As Int64 = Val(grow.Cells("PO_QTY_SHP").Value & "")
                    Dim PO_QTY_REC As Int64 = Val(grow.Cells("PO_QTY_REC").Value & "")

                    PO_QTY_SHP = PO_QTY_SHP - PO_QTY_REC
                    If PO_QTY_SHP > 0 Then

                        If PO_SHIPMENT_LNO = 0 Then
                            PO_SHIPMENT_LNO = Val(dst.Tables("POTSHIP2").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1

                            Dim row2 As DataRow = dst.Tables("POTSHIP2").Rows.Find _
                                                    (New Object() {grow.Cells("PO_SHIPMENT_NO").Value,
                                                                   grow.Cells("PO_SHIPMENT_LNO").Value})

                            Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").NewRow
                            rowPOTSHIP2.ItemArray = row2.ItemArray
                            rowPOTSHIP2.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            rowPOTSHIP2.Item("PO_SHIP_STATUS") = "O"
                            rowPOTSHIP2.Item("TRAN_NO") = DBNull.Value
                            rowPOTSHIP2.Item("OPS_YYYYPP") = DBNull.Value
                            rowPOTSHIP2.Item("OPS_YYYYPP_FIFO") = DBNull.Value
                            rowPOTSHIP2.Item("PO_SOURCE_DOC") = DBNull.Value
                            rowPOTSHIP2.Item("PO_DATE_RECEIVED") = DBNull.Value
                            rowPOTSHIP2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            rowPOTSHIP2.Item("LAST_DATE") = DATETIME_STAMP
                            rowPOTSHIP2.Item("PO_SHIP_CTNS") = 0
                            'PO_SHIP_CTNS
                            'TOTAL_WEIGHT()
                            'TRUCKING()

                            rowPOTSHIP2.Item("ORDR_NO") = DBNull.Value

                            dst.Tables("POTSHIP2").Rows.Add(rowPOTSHIP2)
                        End If

                        Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                        Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value

                        Dim row3 As DataRow = dst.Tables("POTSHIP3").Rows.Find _
                                             (New Object() {grow.Cells("PO_SHIPMENT_NO").Value,
                                                            grow.Cells("PO_SHIPMENT_LNO").Value,
                                                            grow.Cells("PO_ORDER_NO").Value,
                                                            grow.Cells("PO_ORDER_LNO").Value})

                        Dim rowPOTSHIP3 As DataRow = dst.Tables("POTSHIP3").NewRow
                        rowPOTSHIP3.ItemArray = row3.ItemArray
                        rowPOTSHIP3.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                        rowPOTSHIP3.Item("PO_QTY_SHP") = PO_QTY_SHP
                        rowPOTSHIP3.Item("PO_QTY_REC") = 0
                        rowPOTSHIP3.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowPOTSHIP3.Item("LAST_DATE") = DATETIME_STAMP

                        Create_POTSHIPR(PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE)

                        grow.Cells("PO_QTY_SHP").Value = Val(grow.Cells("PO_QTY_SHP").Value & "") - PO_QTY_SHP
                        grow.Update()

                        dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)

                        If PO_QTY_REC = 0 Then
                            For Each rowPOTSHIP7 As DataRow In dst.Tables("POTSHIP7").Select("PO_SHIPMENT_LNO = " & grow.Cells("PO_SHIPMENT_LNO").Value & " and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                                rowPOTSHIP7.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            Next
                        End If
                    End If
                Next

                grd.Selected.Rows.Clear()

            Case "Create Containers from BOL Data"
                Create_Containers_from_BOL()

            Case "Add Remaining Units"

                If grdPOTSHIP7.ActiveRow IsNot Nothing AndAlso grdPOTSHIP7.ActiveRow.IsDataRow Then
                    Dim PO_SHIPMENT_NO As String = grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                    Dim PO_SHIPMENT_LNO As Int64 = Val(grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                    Dim CARTON_NO As Int64 = Val(grdPOTSHIP7.ActiveRow.Cells("CARTON_NO").Value & "")
                    Dim CARTONS As Int64 = Val(grdPOTSHIP7.ActiveRow.Cells("CARTONS").Value & "")
                    'If MsgBox("This option will set the Qty per Carton for all Styles in this Carton (Type " & CStr(CARTON_NO) & ")" _
                    '          & vbCrLf & vbCrLf & "This option is best utilized with each style set to its default qty of 1 unit per case" _
                    '          & vbCrLf & " and the Actual number of Cartons for this carton type already specified" _
                    '          & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    '    Exit Sub
                    'End If

                    Dim sql As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO)
                    For Each rowPOTSHIP8 As DataRow In dst.Tables("POTSHIP8").Select(sql)
                        Dim STYLE_CODE As String = rowPOTSHIP8.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowPOTSHIP8.Item("COLOR_CODE")
                        Dim rowPOTSHIPR As DataRow = dst.Tables("POTSHIPR").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE})
                        Dim QTY_SHP As Int64 = Val(rowPOTSHIPR.Item("QTY_SHP") & "")
                        Dim QTY_CTN As Int64 = Val(rowPOTSHIPR.Item("QTY_CTN") & "")
                        Dim f As Decimal = 0
                        If CARTONS <> 0 Then
                            f = (QTY_SHP - QTY_CTN) / CARTONS
                        End If
                        If f >= 1 Then
                            rowPOTSHIP8.Item("QTY") = Val(rowPOTSHIP8.Item("QTY") & "") + CInt(f)
                        End If
                    Next

                End If

            Case "Calculate Qty/Carton Proportionately"

                If grdPOTSHIP7.ActiveRow IsNot Nothing AndAlso grdPOTSHIP7.ActiveRow.IsDataRow Then
                    Dim PO_SHIPMENT_NO As String = grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                    Dim PO_SHIPMENT_LNO As Int64 = Val(grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                    Dim CARTON_NO As Int64 = Val(grdPOTSHIP7.ActiveRow.Cells("CARTON_NO").Value & "")
                    If MsgBox("This option will set the Qty per Carton for all Styles in this Carton (Type " & CStr(CARTON_NO) & ")" _
                              & vbCrLf & vbCrLf & "This option is best utilized with each style set to its default qty of 1 unit per case" _
                              & vbCrLf & " and the Actual number of Cartons for this carton type already specified" _
                              & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    Dim sql As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO)
                    For Each rowPOTSHIP8 As DataRow In dst.Tables("POTSHIP8").Select(sql)
                        Dim STYLE_CODE As String = rowPOTSHIP8.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowPOTSHIP8.Item("COLOR_CODE")
                        Dim rowPOTSHIPR As DataRow = dst.Tables("POTSHIPR").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE})
                        Dim QTY_SHP As Int64 = Val(rowPOTSHIPR.Item("QTY_SHP") & "")
                        Dim QTY_CTN As Int64 = Val(rowPOTSHIPR.Item("QTY_CTN") & "")
                        Dim f As Decimal = 0
                        If QTY_CTN <> 0 And QTY_SHP <> 0 Then
                            f = QTY_SHP / QTY_CTN
                        End If
                        If f >= 2 Then
                            rowPOTSHIP8.Item("QTY") = Val(rowPOTSHIP8.Item("QTY") & "") * CInt(f)
                        End If
                    Next

                End If

            Case "Calculate Cartons as Qty Shipped / Carton Pack"
                If grdPOTSHIP7.ActiveRow IsNot Nothing AndAlso grdPOTSHIP7.ActiveRow.IsDataRow Then
                    Dim PO_SHIPMENT_NO As String = grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                    Dim PO_SHIPMENT_LNO As Int64 = Val(grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                    If MsgBox("This option will set the Carton Count for All Carton Types in this Container (Line " & CStr(PO_SHIPMENT_LNO) & ")" _
                              & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    With grdPOTSHIP7.ActiveRow


                        Dim sql As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)

                        For Each rowPOTSHIP7 As DataRow In dst.Tables("POTSHIP7").Select(sql)
                            Dim rowPOTSHIP8() As DataRow = dst.Tables("POTSHIP8").Select(sql & " and CARTON_NO = " & rowPOTSHIP7.Item("CARTON_NO"))
                            If rowPOTSHIP8.Length = 1 Then
                                Dim QTY_SHP As Int64 = Val(rowPOTSHIP8(0).GetParentRow("POTSHIPR_POTSHIP8").Item("QTY_SHP") & "")
                                Dim QTY As Int64 = Val(rowPOTSHIP8(0).Item("QTY") & "")
                                If QTY <> 0 Then
                                    Dim CARTONS As Int64 = QTY_SHP / QTY
                                    rowPOTSHIP7.Item("CARTONS") = CARTONS
                                End If
                            End If
                        Next
                    End With
                End If

            Case "Copy Carton Dimensions to All Cartons Referencing PO", "Copy Carton Dimensions to All Cartons Referencing PO (only if blank)"
                If grdPOTSHIP7.ActiveRow IsNot Nothing AndAlso grdPOTSHIP7.ActiveRow.IsDataRow Then

                    Dim PO_ORDER_NO_carton = Get_PO_for_Carton(grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_NO").Value,
                                                         grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value,
                                                         grdPOTSHIP7.ActiveRow.Cells("CARTON_NO").Value)

                    If MsgBox("This option will set the Carton Dimensions" _
                              & vbCrLf & " for ALL Cartons on this Shipment" _
                              & vbCrLf & " connected to PO " & PO_ORDER_NO_carton _
                              & IIf(e.Tool.Key = "Copy Carton Dimensions to All Cartons Referencing PO (only if blank)", " (only if blank)", "") _
                              & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    With grdPOTSHIP7.ActiveRow
                        Dim CARTON_DIMS As String = .Cells("CARTON_DIMS").Value & ""

                        Dim sql As String = ""
                        If e.Tool.Key = "Copy Carton Dimensions to All Cartons Referencing PO (only if blank)" Then
                            sql = "ISNULL(CARTON_DIMS,'') = ''"
                        End If
                        For Each rowPOTSHIP7 As DataRow In dst.Tables("POTSHIP7").Select(sql)
                            If PO_ORDER_NO_carton = Get_PO_for_Carton(rowPOTSHIP7.Item("PO_SHIPMENT_NO"),
                                                rowPOTSHIP7.Item("PO_SHIPMENT_LNO"),
                                                rowPOTSHIP7.Item("CARTON_NO")) Then
                                rowPOTSHIP7.Item("CARTON_DIMS") = CARTON_DIMS
                                rowPOTSHIP7.Item("CARTON_VOLUME") = Get_Volume_from_Dims(CARTON_DIMS)
                            End If
                        Next
                    End With
                End If


            Case "Allow Change to Qty Received"
                grd.ActiveRow.Cells("PO_QTY_REC").Column.CellActivation = UltraWinGrid.Activation.AllowEdit

            Case "Show Multi-Pack"

                tlb_sbt = DirectCast(tlb.Tools("Show Multi-Pack"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then
                    Exit Sub
                End If
                splCartonQ.Panel2Collapsed = Not tlb_sbt.Checked


            Case "Import Packing List from XLS"
                Dim openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                openFileDialog1.RestoreDirectory = True
                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim FILENAME As String = openFileDialog1.FileName
                    ASCMAIN1.Progress("Now Importing Packing XLS")
                    Import_Packing_List(FILENAME)
                End If


            Case "Consolidate Inv/BOL"

                Dim tbl As New DataTable
                With tbl
                    .Columns.Add("COMM_INV_NO")
                    .Columns.Add("BOL_NO")
                    .Columns.Add("CONTAINER_NOS")
                    .Columns.Add("CONTAINER_COUNT", GetType(System.Int32))
                    .Columns.Add("PO_SHIPMENT_LNO", GetType(System.Int32))
                    .PrimaryKey = New DataColumn() { .Columns("COMM_INV_NO"), .Columns("BOL_NO")}
                End With

                For Each ROW As DataRow In dst.Tables("POTSHIP2").Select("")
                    Dim COMM_INV_NO As String = ROW.Item("COMM_INV_NO")
                    Dim BOL_NO As String = ROW.Item("BOL_NO")
                    If COMM_INV_NO <> "" And BOL_NO <> "" Then
                        Dim row2 As DataRow = tbl.Rows.Find(New String() {COMM_INV_NO, BOL_NO})
                        If row2 Is Nothing Then
                            row2 = tbl.NewRow
                            row2.Item("COMM_INV_NO") = COMM_INV_NO
                            row2.Item("BOL_NO") = BOL_NO
                            row2.Item("PO_SHIPMENT_LNO") = ROW.Item("PO_SHIPMENT_LNO")
                            tbl.Rows.Add(row2)
                        End If
                        Dim CONTAINER_NOS As String = row2.Item("CONTAINER_NOS") & ""
                        If CONTAINER_NOS <> "" Then CONTAINER_NOS &= ";"
                        CONTAINER_NOS &= ROW.Item("CONTAINER_NO")
                        row2.Item("CONTAINER_NOS") = CONTAINER_NOS
                        row2.Item("CONTAINER_COUNT") = Val(row2.Item("CONTAINER_COUNT") & "") + 1
                    End If
                Next

                ASCDATA1.DeleteRows(tbl, "CONTAINER_COUNT < 2")
                If tbl.Rows.Count = 0 Then
                    MsgBox("Nothing to Consolidate", MsgBoxStyle.OkOnly, "Cannot Consolidate")
                    Exit Sub
                End If


                Using frm As New ASFMSGBF
                    frm.Show_grd(tbl, Me, "Commercial Invoice / BOL Nos that are Eligible to be Consolidated")
                End Using

                If MsgBox("OK to Consolidate these records?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Consolidating")

                    EnforceConstraints(False)

                    dst.Tables("POTSHIPR").Rows.Clear()
                    ' dst.Tables("POTPACKR").Rows.Clear()

                    For Each row As DataRow In tbl.Select()
                        Dim COMM_INV_NO As String = row.Item("COMM_INV_NO")
                        Dim BOL_NO As String = row.Item("BOL_NO")
                        Dim sql As String = "COMM_INV_NO = '" & COMM_INV_NO & "' and BOL_NO = '" & BOL_NO & "'"
                        Dim PO_SHIPMENT_LNO_MAIN As Integer = Val(row.Item("PO_SHIPMENT_LNO") & "")
                        Dim CARTON_NO_MAX As Integer = Val(dst.Tables("POTSHIP7").Compute("MAX(CARTON_NO)", "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO_MAIN)))
                        Dim rowPOTSHIP2_MAIN As DataRow = dst.Tables("POTSHIP2").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO_MAIN})

                        For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select(sql)
                            Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
                            If PO_SHIPMENT_LNO = PO_SHIPMENT_LNO_MAIN Then
                                ' DO NOTHING - THIS IS THE PARENT LINE
                            Else
                                Dim PO_SHIP_CTNS As Integer = Val(rowPOTSHIP2.Item("PO_SHIP_CTNS") & "")
                                rowPOTSHIP2_MAIN.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP2_MAIN.Item("PO_SHIP_CTNS") & "") + PO_SHIP_CTNS

                                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO))
                                    rowPOTSHIP3.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_MAIN
                                Next

                                For Each rowPOTSHIP7 As DataRow In dst.Tables("POTSHIP7").Select("PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO))
                                    CARTON_NO_MAX += 1
                                    rowPOTSHIP7.Item("PO_SHIPMENT_LNO") = 0
                                    rowPOTSHIP7.Item("CARTON_NO") = CARTON_NO_MAX
                                    rowPOTSHIP7.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_MAIN
                                Next

                                rowPOTSHIP2.Delete()
                            End If
                        Next
                    Next

                    For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3"), New String() {"PO_SHIPMENT_LNO", "STYLE_CODE", "COLOR_CODE"}).Rows
                        Create_POTSHIPR(row.Item("PO_SHIPMENT_LNO"), row.Item("STYLE_CODE"), row.Item("COLOR_CODE"))
                    Next

                    EnforceConstraints(True)

                    dst.Tables("POTSHIP4").Rows.Clear()

                    Create_Containers_from_BOL()

                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")

                    Sort_grdColumns(grdPOTSHIP2, "PO_SHIPMENT_LNO")

                End If


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Mark as Deleted"
            '    Dim TRANS_SEQ As String = grd.ActiveRow.Cells("TRANS_SEQ").Value

            '    If MsgBox("Do you Really want to mark Trans Seq " & TRANS_SEQ & " as Deleted?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
            '        If ASCDATA1.ExecuteSQL("Update ADS.RCPTHDR@ADSIIS Set STATUS = 'D' where TRANS_SEQ = " & TRANS_SEQ & " and (STATUS = '0' or STATUS = 'V')") = 1 Then
            '            grd.ActiveRow.Delete(False)
            '        End If
            '    End If

            '    MsgBox("Trans Seq " & TRANS_SEQ & " has been Deleted", MsgBoxStyle.OkOnly, "Verification")

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Value
                Dim rowAPTINVH1 As DataRow = LookUp("APTINVH1", VOUCHER_NO)
                If rowAPTINVH1 IsNot Nothing Then
                    Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                End If

            Case "Delete 944 from Queue"

                Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value
                Dim EDI_WH_RECEIPT_NO As String = grd.ActiveRow.Cells("EDI_WH_RECEIPT_NO").Value
                Dim EDI_PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("EDI_PO_SHIPMENT_NO").Value

                If MsgBox("OK to delete:" _
                          & vbCrLf & vbCrLf & " Receipt " & EDI_WH_RECEIPT_NO _
                          & vbCrLf & " PO Shipment No " & EDI_PO_SHIPMENT_NO _
                          & vbCrLf & vbCrLf & " from EDI Queue?", MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then Exit Sub

                If Not ASCMAIN1.Logical_Lock("POTSHIP1", EDI_PO_SHIPMENT_NO) Then Exit Sub

                ASCMAIN1.sql = "Update EDT944T1 Set EDI_PROCESS_IND = 'D' where EDI_DOC_SEQ_NO = :PARM1 and NVL(EDI_PROCESS_IND,'0') = '0'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {EDI_DOC_SEQ_NO})

                ASCMAIN1.MultiTask_Release()

                MsgBox("EDI 944 Record has been marked as Deleted", MsgBoxStyle.OkOnly, "Success")

                Get_Receipt_Data_from_3PL()

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")

            Case "Shipment Inquiry"
                If grd.Name = "grdPOTSHIPC" And grd.ActiveRow.Band.Key <> "POTSHIPC_POTSHIPC2" Then
                    Exit Sub
                End If
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")

            Case "Create a carton type containing all selected Style/Colors"
                If grdPOTSHIPR.Selected.Rows.Count = 0 Then
                    If grdPOTSHIPR.ActiveRow IsNot Nothing Then
                        grdPOTSHIPR.Selected.Rows.Add(grdPOTSHIPR.ActiveRow)
                    End If
                End If
                If grdPOTSHIPR.Selected.Rows.Count = 0 Then
                    MsgBox("You must first select rows before creating a carton", MsgBoxStyle.OkOnly, "Cannot Create Carton")
                ElseIf grdPOTSHIPR.ActiveRow IsNot Nothing AndAlso Not grdPOTSHIPR.ActiveRow.Selected Then
                    MsgBox("Active Row is not Selected", MsgBoxStyle.OkOnly, "Cannot Create Carton")
                Else
                    Dim PO_SHIPMENT_LNO As Int32 = 0
                    For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIPR.Selected.Rows
                        If PO_SHIPMENT_LNO = 0 Then
                            PO_SHIPMENT_LNO = Val(grow.Cells("PO_SHIPMENT_LNO").Value & "")
                        Else
                            If PO_SHIPMENT_LNO <> Val(grow.Cells("PO_SHIPMENT_LNO").Value & "") Then
                                MsgBox("Cannot Mix Style/Colors from different Shipment Lines in the Same Carton", MsgBoxStyle.OkOnly, "Cannot Create Carton")
                                Exit Sub
                            End If
                        End If
                    Next

                    Create_Carton_for_Selected_Styles(PO_SHIPMENT_LNO)
                    grdPOTSHIPR.Selected.Rows.Clear()
                End If
                Sort_grdColumns(grdPOTSHIP7, "CARTON_NO")

            Case "Create an individual carton type for All Style/Colors"
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIPR.Rows
                    grdPOTSHIPR.Selected.Rows.Clear()
                    grdPOTSHIPR.Selected.Rows.Add(grow)
                    Dim PO_SHIPMENT_LNO As Int64 = Val(grow.Cells("PO_SHIPMENT_LNO").Value & "")
                    Create_Carton_for_Selected_Styles(PO_SHIPMENT_LNO)
                    If grow.Cells("QTY_VAR").Value > 0 Then
                        Create_Carton_for_Selected_Styles(PO_SHIPMENT_LNO, grow.Cells("QTY_VAR").Value)
                    End If
                Next
                grdPOTSHIPR.Selected.Rows.Clear()
                Sort_grdColumns(grdPOTSHIP7, "CARTON_NO")

            Case "Create an individual carton type for each selected Style/Color"
                If grdPOTSHIPR.Selected.Rows.Count = 0 Then
                    If grdPOTSHIPR.ActiveRow IsNot Nothing Then
                        grdPOTSHIPR.Selected.Rows.Add(grdPOTSHIPR.ActiveRow)
                    End If
                End If
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIPR.Selected.Rows
                    grdPOTSHIPR.Selected.Rows.Clear()
                    grdPOTSHIPR.Selected.Rows.Add(grow)
                    Dim PO_SHIPMENT_LNO As Int64 = Val(grow.Cells("PO_SHIPMENT_LNO").Value & "")
                    Create_Carton_for_Selected_Styles(PO_SHIPMENT_LNO)
                Next
                grdPOTSHIPR.Selected.Rows.Clear()
                Sort_grdColumns(grdPOTSHIP7, "CARTON_NO")

            Case "Show All Carton Details"
                Setup_grdPOTSHIP8()

            Case "Pro-Forma Invoice"
                Print_Invoice()

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Voucher Inquiry"
                Dim CTL_NO As String = grd.ActiveRow.Cells("CTL_NO").Value & ""
                Dim rowPOTLCST1 As DataRow = LookUp("POTLCST1", CTL_NO)
                If rowPOTLCST1 IsNot Nothing Then
                    Dim VOUCHER_NO As String = rowPOTLCST1.Item("VOUCHER_NO") & ""
                    If VOUCHER_NO <> "" Then
                        Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                    End If
                End If

            Case "Style Master File"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If

            Case "Add Style to Carton"
                If grdPOTSHIP7.ActiveRow IsNot Nothing AndAlso grdPOTSHIP7.ActiveRow.IsDataRow Then
                    If dst.Tables("POTSHIP8").Rows.Find(New Object() {grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_NO").Value,
                                                              grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value,
                                                              grdPOTSHIP7.ActiveRow.Cells("CARTON_NO").Value,
                                                              grdPOTSHIPR.ActiveRow.Cells("STYLE_CODE").Value,
                                                              grdPOTSHIPR.ActiveRow.Cells("COLOR_CODE").Value}) IsNot Nothing Then
                        MsgBox("Style/Color is already in this carton", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Exit Sub
                    Else
                        Dim rowPOTSHIP8 As DataRow = dst.Tables("POTSHIP8").NewRow
                        rowPOTSHIP8.Item("PO_SHIPMENT_NO") = grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                        rowPOTSHIP8.Item("PO_SHIPMENT_LNO") = grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value
                        rowPOTSHIP8.Item("CARTON_NO") = grdPOTSHIP7.ActiveRow.Cells("CARTON_NO").Value
                        rowPOTSHIP8.Item("STYLE_CODE") = grdPOTSHIPR.ActiveRow.Cells("STYLE_CODE").Value
                        rowPOTSHIP8.Item("COLOR_CODE") = grdPOTSHIPR.ActiveRow.Cells("COLOR_CODE").Value
                        dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)
                        Check_for_PPK(grdPOTSHIP7.ActiveRow)
                    End If
                End If

            Case "FIFO Previous Period"
                tlb_sbt = DirectCast(tlb.Tools("FIFO Previous Period"), UltraWinToolbars.ButtonTool)
                If tlb_sbt.Checked Then
                    Dim YP As String = grd.ActiveRow.Cells("OPS_YYYYPP").Value
                    grd.ActiveRow.Cells("OPS_YYYYPP_FIFO").Value = ASCMAIN1.Period_Calc(YP, -1)
                Else
                    grd.ActiveRow.Cells("OPS_YYYYPP_FIFO").Value = ""
                End If
                grd.ActiveRow.Update()

            Case "Switch PO & Line"
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PO_ORDER_LNO", , "PO_STATUS = 'O' AND PO_QTY_OPN <> 0")

                If ASCMAIN1.CodeSelector.SQL = "" Then
                    Exit Sub
                End If

                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""

                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()

                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    Dim PO_ORDER_NO As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("PO_ORDER_NO")
                    Dim PO_ORDER_LNO As Int32 = Val(ASCMAIN1.CodeSelector.SelectedRows(0).Item("PO_ORDER_LNO"))
                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, False, False, False, 1) Then
                        MsgBox("Cannot Lock PO " & PO_ORDER_NO, MsgBoxStyle.OkOnly, "Cannot Switch to this PO at this Time")
                        Exit Sub
                    End If

                    Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                    Dim PO_SHIPMENT_LNO As Int32 = Val(grd.ActiveRow.Cells("PO_SHIPMENT_LNO").Value)
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

                    ASCDATA1.ExecuteSP("POPSHIP3_SWAP", "VNVNV",
                                       New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, PO_ORDER_NO, PO_ORDER_LNO, WHSE_CODE},
                                       New String() {-"PO_SHIPMENT_NO_in", "PO_SHIPMENT_LNO_in", "PO_ORDER_NO_in", "PO_ORDER_lNO_in", "WHSE_CODE_in"})

                End If


            Case "Add PO Line"
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PO_ORDER_LNO", , "PO_STATUS = 'O' AND PO_QTY_OPN <> 0")

                If ASCMAIN1.CodeSelector.SQL = "" Then
                    Exit Sub
                End If

                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""

                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()

                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    Dim PO_ORDER_NO As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("PO_ORDER_NO")
                    Dim PO_ORDER_LNO As Int32 = Val(ASCMAIN1.CodeSelector.SelectedRows(0).Item("PO_ORDER_LNO"))
                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, False, False, False, 1) Then
                        MsgBox("Cannot Lock PO " & PO_ORDER_NO, MsgBoxStyle.OkOnly, "Cannot Switch to this PO at this Time")
                        Exit Sub
                    End If
                End If

        End Select
    End Sub

#End Region

    Sub Import_Packing_List(fileName As String)

        If fileName <> "" Then
            Dim eMsg As String = ""
            packingListPOs.Clear()

            ASCMAIN1.Progress("Now Importing Packing XLS")

            Dim oWB As SpreadsheetGear.IWorkbook

            If Not packingFromXLS Then
                dst.Tables("POTSHPIE").Rows.Clear()
                tabBOL.Tabs("Import Errors").Visible = True
                tabBOL.Tabs("Packing Discrepancies").Visible = True
            Else
                If dicWorkbooks.Contains(fileName) Then
                    MsgBox("Workbook " & fileName & " has already been selected for Import", MsgBoxStyle.OkOnly, "Cannot select same workbook more than once in Import session")
                    Exit Sub
                End If
            End If

            Try
                oWB = SpreadsheetGear.Factory.GetWorkbook(fileName)

                If Workbook_Is_Valid(oWB, eMsg) Then
                    ASCMAIN1.Progress("Now Importing Workbook")
                    packingFromXLS = True
                    dicWorkbooks.Add(fileName)
                    Set_Entry_Mode_Controls(packingFromXLS)

                    ASCMAIN1.Progress("Now Obtaining PO Locks")

                    If Got_PO_Locks(eMsg) Then
                        Initialize_Workbook_Memory()
                        WORKBOOK_COUNTER += 1
                        For ws As Integer = 0 To oWB.Worksheets.Count - 1
                            Initialize_Worksheet_Memory()
                            eMsg = ""

                            Import_Packing_Worksheet(oWB.Worksheets(ws), eMsg)
                            If eMsg <> "" Then
                                MsgBox(eMsg, MsgBoxStyle.Critical, "Import Error")

                            End If
                        Next
                        Create_Containers_from_BOL()
                    Else
                        MsgBox(eMsg, MsgBoxStyle.Critical, "Multitasking Error")
                        Exit Sub
                    End If


                    For Each rowPOTSHPIE As DataRow In dst.Tables("POTSHPIE").Select("WORKBOOK = '" & oWB.Name & "'")
                        ' Dim QTY_PACK As Int64 = 0

                        Dim PO_ORDER_NO As String = rowPOTSHPIE.Item("PO_ORDER_NO") & ""

                        Dim STYLE_CODE As String = rowPOTSHPIE.Item("STYLE_CODE") & ""
                        Dim COLOR_CODE As String = rowPOTSHPIE.Item("COLOR_CODE") & ""
                        ASCMAIN1.sql = "Select POTORDR2.PO_ORDER_NO, POTORDR1.PO_REFERENCE, Sum (PO_QTY_OPN) PO_QTY_OPN from POTORDR2,POTORDR1" & vbCrLf _
                            & " where POTORDR1.VEND_CODE = 'YINTAK'" & vbCrLf _
                            & "  and POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
                            & "  and POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                            & "  and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                            & "  and POTORDR2.PO_QTY_OPN <> 0" & vbCrLf _
                            & "  group by POTORDR2.PO_ORDER_NO, POTORDR1.PO_REFERENCE"

                        For Each rowPO As DataRow In ASCDATA1.GetDataTable.Select("")
                            Dim PO_ORDER_NO_OTHER As String = rowPO.Item("PO_ORDER_NO") & ""
                            Dim PO_REFERENCE_OTHER As String = rowPO.Item("PO_REFERENCE") & ""
                            Dim PO_QTY_OPN As Int64 = Val(rowPO.Item("PO_QTY_OPN") & "")
                            If PO_ORDER_NO_OTHER = PO_ORDER_NO Then
                                rowPOTSHPIE.Item("QTY_OPEN_THIS_PO") = PO_QTY_OPN
                            Else
                                rowPOTSHPIE.Item("QTY_OPEN_OTHER_POS") = Val(rowPOTSHPIE.Item("QTY_OPEN_OTHER_POS") & "") + PO_QTY_OPN
                                If rowPOTSHPIE.Item("PO_ORDER_NO1") & "" = "" Then
                                    rowPOTSHPIE.Item("PO_REFERENCE1") = PO_REFERENCE_OTHER
                                    rowPOTSHPIE.Item("PO_ORDER_NO1") = PO_ORDER_NO_OTHER
                                    rowPOTSHPIE.Item("QTY_OPEN_PO_ORDER_NO1") = PO_QTY_OPN
                                ElseIf rowPOTSHPIE.Item("PO_ORDER_NO2") & "" = "" Then
                                    rowPOTSHPIE.Item("PO_REFERENCE2") = PO_REFERENCE_OTHER
                                    rowPOTSHPIE.Item("PO_ORDER_NO2") = PO_ORDER_NO_OTHER
                                    rowPOTSHPIE.Item("QTY_OPEN_PO_ORDER_NO2") = PO_QTY_OPN
                                End If

                            End If
                        Next
                        Dim QTY_OPEN As Int64 = Val(ASCDATA1.GetDataValue)
                        rowPOTSHPIE.Item("QTY_PACK") = rowPOTSHPIE.Item("QTY")
                    Next
                Else

                    MsgBox(eMsg, MsgBoxStyle.Critical, "Import Error")
                    Exit Sub
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error has occurred")
            End Try


        End If
        ASCMAIN1.Progress("")

    End Sub

    Sub Import_Packing_Worksheet(ws As SpreadsheetGear.IWorksheet, ByRef eMsg As String)

        ASCMAIN1.Progress("Now Importing " & ws.Name)

        Dim wbName As String = ws.Workbook.Name
        Dim rCount As Int64 = ws.UsedRange.RowCount
        Dim ccMergedCode As String = ""
        Dim COMM_INV_NO As String = Trim(Replace(Replace(Replace(ws.Cells(4, 0).Text & "", " ", ""), "INVOICENO", ""), ":", ""))
        Dim BOL_NO As String = Trim(Replace(ws.Cells(10, 8).Text & "", ":", ""))
        Dim CONTAINER_NO As String = Trim(Replace(ws.Cells(6, 8).Text & "", ":", ""))

        Dim packingLinesStart As Integer = 0

        For TEST_LINE As Integer = 13 To 23
            If ws.Cells(TEST_LINE, 0).Text & "" = "NUMBER" And ws.Cells(TEST_LINE - 1, 0).Text & "" = "CARTON" Then
                packingLinesStart = TEST_LINE + 1
                Exit For
            End If
        Next

        Dim COLOR_CODEs As New List(Of String)
        Dim COLORs As String = ws.Cells(packingLinesStart - 3, 0).Text
        If COLORs <> "" Then
            For Each COLOR_CODE As String In Split(COLORs, ",")
                If COLOR_CODE.Length <> 3 Then
                    COLOR_CODEs.Clear()
                    Exit For
                Else
                    If LookUp("ICTCOLR1", COLOR_CODE) Is Nothing Then
                        COLOR_CODEs.Clear()
                        Exit For
                    End If
                End If
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If Trim(BOL_NO).StartsWith("ETD/CTG") Then
            BOL_NO = Trim(Replace(ws.Cells(11, 7).Text & "", ":", ""))
        End If
        If Trim(CONTAINER_NO).StartsWith("CONT. NO") Then
            CONTAINER_NO = Trim(Replace(ws.Cells(8, 7).Text & "", ":", ""))
        End If

        Dim eMsgsStartEnd As New List(Of String)

        ASCMAIN1.Progress("Now Importing Worksheet", ws.Name)

        Dim poadj As Integer = 0
        If ws.Cells(packingLinesStart - 2, 3).Text = "PO" Or ws.Cells(packingLinesStart - 2, 4).Text = "PO" Then
            poadj = 1
        End If

        Dim TOTAL_PCS_style As Integer = 0
        Dim TOTAL_PCS_STYLE_COLOR As String = ""

        Dim PO_REFERENCEs() As String = New String() {}

        For r As Int64 = packingLinesStart To rCount - 1
            Dim PACKING_LNO As Integer = (r - packingLinesStart) + 1

            Dim PO_ORDER_NO As String = ""
            Dim PO_REFERENCE As String = ""
            Dim PO_REFERENCE_X As String = ""
            Dim PO_SPEC_ORDR_NO As String = ""

            If poadj = 1 Then
                'PO_REFERENCE = ws.Cells(r, 4).Text & "" '  ws.Cells(r, 3).Text & ""
                PO_REFERENCE_X = ws.Cells(r, 4).Text & "" '  ws.Cells(r, 3).Text & ""
                PO_REFERENCE_X = Replace(Replace(Replace(PO_REFERENCE_X, " ", ""), "+", ","), "&", ",")

                Dim PO_ORDER_NO_LAST As String = ""
                For Each PO_REFERENCE In Split(PO_REFERENCE_X, ",")
                    Dim rowPOTORDR1 As DataRow = Get_PO_Header_By_Ref_No(PO_REFERENCE)

                    If rowPOTORDR1 IsNot Nothing Then
                        PO_ORDER_NO = rowPOTORDR1.Item("PO_ORDER_NO")
                        PO_SPEC_ORDR_NO = rowPOTORDR1.Item("PO_SPEC_ORDR_NO") & ""
                        Got_PO_Lock(PO_ORDER_NO)
                    End If

                    If Not packingListPOs.Contains(PO_ORDER_NO) Then
                        packingListPOs.Add(PO_ORDER_NO)
                    End If

                    If PO_ORDER_NO_LAST <> "" Then
                        dicPOTORDR2(PO_ORDER_NO).Merge(dicPOTORDR2(PO_ORDER_NO_LAST))
                        'dicPOTORDR2.Remove(PO_ORDER_NO_LAST)
                        dicPOTORDR2(PO_ORDER_NO_LAST) = dicPOTORDR2(PO_ORDER_NO)
                    End If
                    PO_ORDER_NO_LAST = PO_ORDER_NO
                Next
            Else
                PO_ORDER_NO = dicWorksheetPOs(ws.Name)
                PO_REFERENCE = "?"
                Dim row1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                If row1 IsNot Nothing Then
                    PO_REFERENCE = row1.Item("PO_REFERENCE")
                    PO_SPEC_ORDR_NO = row1.Item("PO_SPEC_ORDR_NO") & ""
                End If
            End If

            If PO_ORDER_NO = "" Then

                MsgBox("Cannot Find Open PO with PO Reference " & PO_REFERENCE, MsgBoxStyle.OkOnly, "Error")
                Exit Sub

            End If

            Dim STYLE_CODE As String = Trim(ws.Cells(r, 3 + poadj).Text & "")
            Dim COLOR_CODE_m As String = Extract_Color_Code(ws.Cells(r, 4 + poadj).Text & "") 'merged column cells do not return value after first row
            If COLOR_CODE_m <> "" Then
                ccMergedCode = COLOR_CODE_m
            End If
            Dim COLOR_CODE As String = Trim(ccMergedCode)
            Dim SIZE As String = Trim(ws.Cells(r, 5 + poadj).Text & "")
            Dim TOTAL_PCS As Integer = Val(ws.Cells(r, 9 + poadj).Text & "")

            If STYLE_CODE & "-" & COLOR_CODE <> TOTAL_PCS_STYLE_COLOR Then
                TOTAL_PCS_STYLE_COLOR = STYLE_CODE & "-" & COLOR_CODE
                TOTAL_PCS_style = TOTAL_PCS
                Dim rnext As Integer = 1
                Do While ws.Cells(r + rnext, 3 + poadj).Text & "" = STYLE_CODE  ' And ws.Cells(r + rnext, 5 + poadj).Text & "" = COLOR_CODE
                    TOTAL_PCS_style += Val(ws.Cells(r + rnext, 9 + poadj).Text & "")
                    rnext += 1
                Loop
            End If

            Dim tblPOTORDR2 As DataTable = dicPOTORDR2(PO_ORDER_NO)
            Dim rowPOTORDR2 As DataRow = Nothing
            Dim splitLine As Boolean = False

            If COLOR_CODE = "ASSORTED" Or SIZE = "ASSORTED" Or UCase(Trim(PO_SPEC_ORDR_NO)) = "INITIAL" Then
                eMsg = ""
                Handle_Prepacks(ws, eMsg, r, poadj, PO_ORDER_NO, PO_REFERENCE, COMM_INV_NO, BOL_NO, CONTAINER_NO, COLOR_CODEs)
                Exit For
            End If

            Dim PO_SHIPMENT_LNO As Integer = 0

            If Valid_Style(STYLE_CODE) Then
                If Valid_Color(COLOR_CODE) Then

                    Dim rowICTSTYTL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    Dim SUB_UNIT_PACK_QTY As Integer = Val(rowICTSTYTL1.Item("SUB_UNIT_PACK_QTY") & "")
                    If SUB_UNIT_PACK_QTY = 0 Then SUB_UNIT_PACK_QTY = 1
                    If SUB_UNIT_PACK_QTY <> 1 Then
                        TOTAL_PCS = TOTAL_PCS / SUB_UNIT_PACK_QTY
                        TOTAL_PCS_style = TOTAL_PCS_style / SUB_UNIT_PACK_QTY
                    End If
                    '  If ASCMAIN1.Running_in_VS And STYLE_CODE = "WM194057" Then Stop

                    If TOTAL_PCS > 0 Then
                        Dim rowPOTSHIP2 As DataRow = Get_Shipment_Line(COMM_INV_NO, BOL_NO, CONTAINER_NO)
                        Dim rowPOTSHPWB As DataRow = Get_Worksheet_Container_Line(wbName, ws.Name, rowPOTSHIP2)
                        PO_SHIPMENT_LNO = rowPOTSHIP2.Item("PO_SHIPMENT_LNO")
                        Dim packingLine As poPackingLine = Get_Packing_Line_By_Qty(PO_ORDER_NO, tblPOTORDR2, STYLE_CODE, COLOR_CODE, TOTAL_PCS_style) 'TOTAL_PCS)
                        If packingLine.eMsg = "" Then
                            Dim rowPOTSHPXL As DataRow = Get_Packing_Line_By_Lno(PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE, PACKING_LNO)
                            Dim TOTAL_CTN As Integer = Val(ws.Cells(r, 6 + poadj).Text & "")

                            Dim CTN_NO_START As Integer = Val(Trim(ws.Cells(r, 0).Text & ""))
                            Dim CTN_NO_END As Integer = Val(Trim(ws.Cells(r, 2).Text & ""))

                            If CTN_NO_END - CTN_NO_START + 1 <> TOTAL_CTN Then
                                Dim eMsgStartEnd As String = "Problem with Starting/Ending Carton Nos (" & CStr(CTN_NO_START) & "/" & CStr(CTN_NO_END) & ") vs Total Cartons (" & CStr(TOTAL_CTN) & ") on " & ws.Name
                                If Not eMsgsStartEnd.Contains(eMsgStartEnd) Then
                                    eMsgsStartEnd.Add(eMsgStartEnd)
                                    MsgBox(eMsgStartEnd, MsgBoxStyle.OkOnly, "Please do NOT Update without Fixing XLS")
                                End If
                            End If

                            rowPOTSHPXL.Item("CTN_NO_START") = Val(Trim(ws.Cells(r, 0).Text & ""))
                            rowPOTSHPXL.Item("CTN_NO_END") = Val(Trim(ws.Cells(r, 2).Text & ""))
                            rowPOTSHPXL.Item("PO_ORDER_NO") = PO_ORDER_NO
                            rowPOTSHPXL.Item("CONTAINER_NO") = CONTAINER_NO
                            rowPOTSHPXL.Item("SIZE") = SIZE
                            rowPOTSHPXL.Item("TOTAL_CTN") = TOTAL_CTN
                            ' wjz adding division by SUB_UNIT_PACK_QTY 10/23/2021 because packing lists are listing pcs per carton not units per cartpn
                            rowPOTSHPXL.Item("PER_CTN") = Val(ws.Cells(r, 7 + poadj).Text & "") / SUB_UNIT_PACK_QTY
                            rowPOTSHPXL.Item("TOTAL_PCS") = TOTAL_PCS

                            rowPOTSHPXL.Item("GW") = Val(ws.Cells(r, 12 + poadj).Text & "")
                            rowPOTSHPXL.Item("NW") = Val(ws.Cells(r, 13 + poadj).Text & "")
                            rowPOTSHPXL.Item("TTL_GW") = Val(ws.Cells(r, 10 + poadj).Text & "")
                            rowPOTSHPXL.Item("TTL_NW") = Val(ws.Cells(r, 11 + poadj).Text & "")

                            Dim measCM As String = Validate_Carton_Dimensions(ws.Cells(r, 14 + poadj).Text & "").CTN_DIMS_CM
                            rowPOTSHPXL.Item("MEAS") = measCM
                            rowPOTSHPXL.Item("IS_SPLIT") = ""
                            rowPOTSHPXL.Item("WORKBOOK") = wbName
                            rowPOTSHPXL.Item("WORKSHEET") = ws.Name
                            rowPOTSHPXL.Item("PO_ORDER_LNO") = Val(packingLine.rowPOTORDR2.Item("PO_ORDER_LNO") & "")

                            ' packingLine.rowPOTORDR2.Item("PO_QTY_OPN") = Val(packingLine.rowPOTORDR2.Item("PO_QTY_OPN") & "") - TOTAL_PCS

                            rowPOTSHPXL.Item("IS_SPLIT") = IIf(packingLine.splitLine, "1", "0")
                            If rowPOTSHPXL.RowState <> DataRowState.Added Then
                                dst.Tables("POTSHPXL").Rows.Add(rowPOTSHPXL)
                            End If

                            rowPOTSHPWB.Item("IS_PPK") = "0"
                            rowPOTSHPWB.Item("TOTAL_CTNS") = Val(rowPOTSHPWB.Item("TOTAL_CTNS") & "") + TOTAL_CTN

                            rowPOTSHPWB.Item("NW") = Val(ws.Cells(r, 14 + poadj).Text & "")
                            rowPOTSHPWB.Item("MEAS") = measCM

                        Else
                            eMsg &= packingLine.eMsg & ws.Name & ", row " & r + 1 & ": " & STYLE_CODE & "/" & COLOR_CODE & "/" & ws.Cells(r, 9 + poadj).Text & vbCrLf
                            Log_Import_Error(ws, packingLine.eMsg, STYLE_CODE, COLOR_CODE, TOTAL_PCS, PO_ORDER_NO, PO_REFERENCE, COMM_INV_NO, BOL_NO, CONTAINER_NO, PO_SHIPMENT_LNO)
                            ' Exit For
                        End If
                    Else
                        eMsg &= "Invalid QTY on sheet " & ws.Name & ", row " & r + 1 & ": " & STYLE_CODE & "/" & COLOR_CODE & "/" & ws.Cells(r, 9 + poadj).Text & vbCrLf
                        Log_Import_Error(ws, "Invalid QTY on sheet", STYLE_CODE, COLOR_CODE, TOTAL_PCS, PO_ORDER_NO, PO_REFERENCE, COMM_INV_NO, BOL_NO, CONTAINER_NO, PO_SHIPMENT_LNO)
                    End If
                Else
                    eMsg &= "Invalid Color Code on sheet " & ws.Name & ", row " & r + 1 & ": " & STYLE_CODE & "/" & COLOR_CODE & "/" & vbCrLf
                    Log_Import_Error(ws, "Invalid Color Code on sheet", STYLE_CODE, COLOR_CODE, TOTAL_PCS, PO_ORDER_NO, PO_REFERENCE, COMM_INV_NO, BOL_NO, CONTAINER_NO, PO_SHIPMENT_LNO)
                End If
            Else
                Dim checkCell4 As String = Trim(ws.Cells(r, 4 + poadj).Text & "")
                Dim checkCell5 As String = Trim(ws.Cells(r, 5 + poadj).Text & "")
                Dim checkCell6 As String = Trim(ws.Cells(r, 6 + poadj).Text & "")
                If checkCell4.StartsWith("TOTAL") Or checkCell5.StartsWith("TOTAL") Or checkCell6.StartsWith("TOTAL") Then
                    Exit For
                Else
                    If STYLE_CODE <> "" Then
                        eMsg &= "Invalid Style Code on sheet " & ws.Name & ", row " & r + 1 & ": " & STYLE_CODE & vbCrLf
                        Log_Import_Error(ws, "Invalid Style Code on sheet", STYLE_CODE, COLOR_CODE, TOTAL_PCS, PO_ORDER_NO, PO_REFERENCE, COMM_INV_NO, BOL_NO, CONTAINER_NO, PO_SHIPMENT_LNO)
                    End If
                End If
            End If
        Next

        If eMsg = "" Then

            Dim SQLB As String = "IMPORTED = '0' AND WORKBOOK = '" & wbName & "' AND WORKSHEET = '" & ws.Name & "'"

            For Each rowPOTSHPWB As DataRow In dst.Tables("POTSHPWB").Select(SQLB, "WORKBOOK, WORKSHEET, PO_SHIPMENT_LNO")
                Dim PO_SHIPMENT_LNO As Integer = rowPOTSHPWB.Item("PO_SHIPMENT_LNO")
                Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
                Dim IS_PPK As Boolean = (rowPOTSHPWB.Item("IS_PPK") & "" = "1")
                Dim lineToPack As Boolean = False
                Dim sqlPackingLines As String = "WORKBOOK = '" & wbName & "' AND WORKSHEET = '" & ws.Name _
                                               & "' and PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO
                sqlPackingLines = "WORKBOOK = '" & rowPOTSHPWB.Item("WORKBOOK") & "' AND WORKSHEET = '" & rowPOTSHPWB.Item("WORKSHEET") _
                               & "' and PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO


                For Each rowPOTSHPXL As DataRow In dst.Tables("POTSHPXL").Select(sqlPackingLines, "WORKBOOK, WORKSHEET, PO_SHIPMENT_LNO, PACKING_LNO")
                    lineToPack = True


                    Dim WJZSTYLE As String = rowPOTSHPXL.Item("STYLE_CODE")
                    '  If ASCMAIN1.Running_in_VS And (WJZSTYLE = "WM194057") Then Stop


                    Dim PO_ORDER_NO As String = rowPOTSHPXL.Item("PO_ORDER_NO")
                    Dim PO_ORDER_LNO As Integer = Val(rowPOTSHPXL.Item("PO_ORDER_LNO") & "")

                    ' here is where we need to use the "other POs" table
                    Dim tblPOTORDR2 As DataTable = dicPOTORDR2(PO_ORDER_NO)

                    Dim maxLno As Integer = Val(tblPOTORDR2.Compute("MAX(PO_ORDER_LNO)", "") & "")
                    Dim maxLnoSplit As Integer = Val(dst.Tables("POTORDR2_SPLIT").Compute("MAX(PO_ORDER_LNO)", "PO_ORDER_NO = '" & PO_ORDER_NO & "'") & "")
                    Dim isSplitLine As Boolean = (rowPOTSHPXL.Item("IS_SPLIT") & "" = "1")

                    If maxLnoSplit = 0 Or maxLnoSplit < maxLno Then
                        maxLnoSplit = maxLno
                    End If

                    If Not POTORDR1_added.Contains(PO_ORDER_NO) Then
                        POTORDR1_added.Add(PO_ORDER_NO)
                        Dim rowsDC() As DataRow = dst.Tables("POTORDR1_SPLIT").Select("PO_ORDER_NO = '" & PO_ORDER_NO & "'") 'WHY DO I NEED TO DO THIS?
                        If rowsDC.Length = 0 Then
                            Dim rowPOTORDR1 As DataRow = dicPOTORDR1(PO_ORDER_NO)
                            Dim rowPOTORDR1_SPLIT As DataRow = dst.Tables("POTORDR1_SPLIT").NewRow
                            For i As Int16 = 0 To rowPOTORDR1.ItemArray.Length - 1
                                rowPOTORDR1_SPLIT.Item(i) = rowPOTORDR1.Item(i)
                            Next
                            dst.Tables("POTORDR1_SPLIT").Rows.Add(rowPOTORDR1_SPLIT)
                        End If
                    End If

                    Dim rowPOTORDR2 As DataRow = Nothing
                    Dim rowPOTORDR2_orig As DataRow = tblPOTORDR2.Select("PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & PO_ORDER_LNO)(0)
                    Dim rowPOTORDR2_pack() As DataRow = dst.Tables("POTORDR2_SPLIT").Select("PO_ORDER_NO = '" & PO_ORDER_NO _
                                                                                            & "' and PO_ORDER_LNO" & IIf(isSplitLine, "_ORIG", "") & " = " & PO_ORDER_LNO)
                    Dim PO_QTY_SHP As Integer = Val(rowPOTSHPXL.Item("TOTAL_PCS") & "")
                    Dim PO_QTY_ORD As Integer = Val(rowPOTORDR2_orig.Item("PO_QTY_ORD") & "")
                    Dim PO_QTY_SHP_TOT As Integer = PO_QTY_SHP
                    Dim newSplitLine As Boolean = (rowPOTORDR2_pack.Length = 0)

                    If newSplitLine Then
                        Dim PO_ORDER_LNO_split As Integer = IIf(isSplitLine, maxLnoSplit + 1, PO_ORDER_LNO)
                        rowPOTORDR2 = dst.Tables("POTORDR2_SPLIT").NewRow()
                        For i As Int16 = 0 To rowPOTORDR2_orig.ItemArray.Length - 1
                            rowPOTORDR2.Item(i) = rowPOTORDR2_orig.Item(i)
                        Next

                        rowPOTORDR2.Item("PO_ORDER_LNO") = PO_ORDER_LNO_split

                        If isSplitLine Then
                            rowPOTORDR2.Item("PO_ORDER_LNO_ORIG") = PO_ORDER_LNO
                        End If
                        rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_SHP
                        rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP
                        rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_SHP
                        rowPOTORDR2.Item("STYLE_DESC") = dicStyleDesc(rowPOTORDR2.Item("STYLE_CODE"))
                        rowPOTORDR2.Item("COLOR_DESC") = dicColorDesc(rowPOTORDR2.Item("COLOR_CODE"))
                        rowPOTORDR2.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                        dst.Tables("POTORDR2_SPLIT").Rows.Add(rowPOTORDR2)

                    Else
                        rowPOTORDR2 = rowPOTORDR2_pack(0)
                        Dim PO_QTY_SHP_new As Integer = Val(rowPOTORDR2.Item("PO_QTY_SHP") & "") + PO_QTY_SHP
                        Dim PO_QTY_OPN_new As Integer = PO_QTY_ORD - PO_QTY_SHP_new
                        rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_SHP_new
                        rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP_new
                        rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_SHP_new
                    End If

                Next

                If lineToPack Then
                    For Each rowPO As DataRow In dst.Tables("POTORDR1_SPLIT").Select()
                        Dim PO_ORDER_NO As String = rowPO.Item("PO_ORDER_NO") & ""
                        If packingListPOs.Contains(PO_ORDER_NO) Then
                            Load_POs_into_POTSHIP3(PO_ORDER_NO, True, PO_SHIPMENT_LNO, COLOR_CODEs)
                        End If
                    Next
                    If ASCMAIN1.Running_in_VS And rowPOTSHPWB.Item("IMPORTED") & "" = "1" Then Stop
                    'rowPOTSHIP2.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP2.Item("PO_SHIP_CTNS") & "") + Create_Cartons_From_Import(PO_SHIPMENT_LNO, ws)
                    Create_Cartons_From_Import(PO_SHIPMENT_LNO, ws)
                    rowPOTSHIP2.Item("PO_SHIP_CTNS") = Val(dst.Tables("POTSHIP7").Compute("SUM(CARTONS)", "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)) & "")
                    rowPOTSHPWB.Item("IMPORTED") = "1"
                End If

            Next

        Else
            Exit Sub
        End If

    End Sub

    Sub Create_Cartons_From_Import(PO_SHIPMENT_LNO As Integer, ws As SpreadsheetGear.IWorksheet)

        Dim wbName As String = ws.Workbook.Name
        Dim CARTON_NO As Int32 = Val(dst.Tables("POTSHIP7").Compute("MAX (CARTON_NO)", "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)) & "")  ' + 1
        Dim rowPOTSHPWB As DataRow = dst.Tables("POTSHPWB").Rows.Find(New Object() {wbName, ws.Name, PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
        Dim isPPK As Boolean = (rowPOTSHPWB.Item("IS_PPK") & "" = "1")

        If isPPK Then
            Dim NW As Decimal = 0
            Dim MEAS As String = ""

            If rowPOTSHPWB IsNot Nothing Then
                NW = rowPOTSHPWB.Item("NW")
                MEAS = rowPOTSHPWB.Item("MEAS") & ""
            End If

            Dim CTN_NO_START_last As Int32 = 0

            For Each rowPOTSHPXL As DataRow In dst.Tables("POTSHPXL").Select("WORKSHEET = '" & ws.Name & "' and PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO & " and TOTAL_PCS <> 0", "PACKING_LNO")
                Dim CTN_NO_START As Int32 = Val(rowPOTSHPXL.Item("CTN_NO_START") & "")
                Dim new_carton_type As Boolean = False
                If CTN_NO_START <> CTN_NO_START_last Then
                    CARTON_NO += 1
                    new_carton_type = True
                    CTN_NO_START_last = CTN_NO_START
                End If
                Dim rowPOTSHIP7 As DataRow = Create_Carton_Type_From_XLS_Packing_Line(PO_SHIPMENT_LNO, CARTON_NO, rowPOTSHPXL, True, new_carton_type)
            Next
        Else
            For Each rowPOTSHPXL As DataRow In dst.Tables("POTSHPXL").Select("WORKSHEET = '" & ws.Name & "'")
                '  If ASCMAIN1.Running_in_VS And rowPOTSHPXL.Item("STYLE_CODE") = "WM194057" Then Stop

                CARTON_NO += 1
                Dim rowPOTSHIP7 As DataRow = Create_Carton_Type_From_XLS_Packing_Line(PO_SHIPMENT_LNO, CARTON_NO, rowPOTSHPXL)
            Next
        End If

    End Sub

    Function Get_Shipment_Line(COMM_INV_NO As String, BOL_NO As String, CONTAINER_NO As String) As DataRow
        Dim rowPOTSHIP2 As DataRow = Nothing
        Dim sqlShipLine As String = "COMM_INV_NO = '" & Mid(COMM_INV_NO, 1, 20) & "' AND BOL_NO ='" & Mid(BOL_NO, 1, 20) & "' AND CONTAINER_NO = '" & Mid(CONTAINER_NO, 1, 20) & "'"
        If WORKBOOK_COUNTER <> 0 Then
            sqlShipLine &= " and WORKBOOK_COUNTER = " & CStr(WORKBOOK_COUNTER)
        End If
        Dim rowsShip() As DataRow = dst.Tables("POTSHIP2").Select(sqlShipLine)
        If rowsShip.Length = 0 Then
            rowPOTSHIP2 = dst.Tables("POTSHIP2").NewRow
            rowPOTSHIP2.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP2.Item("PO_SHIPMENT_LNO") = Val(dst.Tables("POTSHIP2").Compute("Max(PO_SHIPMENT_LNO)", "") & "") + 1
            rowPOTSHIP2.Item("COMM_INV_NO") = Mid(COMM_INV_NO, 1, 20)
            rowPOTSHIP2.Item("BOL_NO") = Mid(BOL_NO, 1, 20)
            rowPOTSHIP2.Item("CONTAINER_NO") = Mid(CONTAINER_NO, 1, 20)
            rowPOTSHIP2.Item("PO_SHIP_STATUS") = "O"
            rowPOTSHIP2.Item("ACCRUAL_STATUS") = "0"
            rowPOTSHIP2.Item("WORKBOOK_COUNTER") = WORKBOOK_COUNTER
            dst.Tables("POTSHIP2").Rows.Add(rowPOTSHIP2)
        Else
            rowPOTSHIP2 = rowsShip(0)
        End If
        Return rowPOTSHIP2
    End Function
    Function Get_Packing_Line_By_Lno(PO_SHIPMENT_NO As String, PO_SHIPMENT_LNO As Integer, STYLE_CODE As String, COLOR_CODE As String, PACKING_LNO As Integer) As DataRow
        Dim rowPOTSHPXL As DataRow = Nothing
        Dim sqlPackingLine As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO & " and " _
                                       & " STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and PACKING_LNO = " & PACKING_LNO
        Dim rowPackingLine() As DataRow = dst.Tables("POTSHPXL").Select(sqlPackingLine)
        If rowPackingLine.Length = 0 Then
            rowPOTSHPXL = dst.Tables("POTSHPXL").NewRow
            rowPOTSHPXL.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHPXL.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            rowPOTSHPXL.Item("STYLE_CODE") = STYLE_CODE
            rowPOTSHPXL.Item("COLOR_CODE") = COLOR_CODE
            rowPOTSHPXL.Item("PACKING_LNO") = PACKING_LNO
        Else
            rowPOTSHPXL = rowPackingLine(0)
        End If
        Return rowPOTSHPXL
    End Function
    Function Get_Packing_Line_By_Qty(PO_ORDER_NO As String, tblPOTORDR2 As DataTable, STYLE_CODE As String, COLOR_CODE As String, PO_QTY_OPEN As Integer, Optional comparison As String = "=") As poPackingLine
        Dim packingLine As New poPackingLine
        packingLine.eMsg = ""
        Dim rows() As DataRow = Get_Matching_PO_Details(tblPOTORDR2, STYLE_CODE, COLOR_CODE, PO_QTY_OPEN)

        If rows.Length = 1 Then
            packingLine.rowPOTORDR2 = rows(0)
        Else
            rows = Get_Matching_PO_Details(tblPOTORDR2, STYLE_CODE, COLOR_CODE, PO_QTY_OPEN, ">=")
            If rows.Length = 1 Then
                packingLine.splitLine = True
                packingLine.rowPOTORDR2 = rows(0)
            Else
                If rows.Length = 0 Then
                    ' packingLine.eMsg = "No Matching PO details for Style/Color/Qty " & vbCrLf & STYLE_CODE & "/" & COLOR_CODE & "/" & CStr(PO_QTY_OPEN) & " for PO " & PO_ORDER_NO & vbCrLf & "on sheet "
                    packingLine.eMsg = "No Matching PO details for Style/Color/Qty"
                Else

                    Dim qty_already_packed As Int32 = 0

                    For Each row As DataRow In rows ' look for an already split row with the exact qty - not sure what to do if there are multiple rows with exact qty so just going with the 1st one we find
                        If Val(row.Item("PO_QTY_OPN") & "") = PO_QTY_OPEN Then
                            packingLine.rowPOTORDR2 = row
                            Exit For
                        End If
                    Next
                    If packingLine.rowPOTORDR2 Is Nothing Then
                        ' packingLine.eMsg = "Multiple Matching PO details for Style/Color/Qty " & vbCrLf & STYLE_CODE & "/" & COLOR_CODE & "/" & CStr(PO_QTY_OPEN) & " for PO " & PO_ORDER_NO & vbCrLf & " on sheet "
                        packingLine.eMsg = "Multiple Matching PO details for Style/Color/Qty"
                    End If
                End If
            End If
        End If
        Return packingLine
    End Function
    Function Get_Matching_PO_Details(tblPOTORDR2 As DataTable, STYLE_CODE As String, COLOR_CODE As String, PO_QTY_OPEN As Integer, Optional comparison As String = "=") As DataRow()
        Return tblPOTORDR2.Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and PO_QTY_OPN " & comparison & PO_QTY_OPEN)
    End Function

    Function Get_Worksheet_Container_Line(wbName As String, wsName As String, rowPOTSHIP2 As DataRow) As DataRow
        Dim PO_SHIPMENT_LNO As Integer = rowPOTSHIP2.Item("PO_SHIPMENT_LNO")
        Dim rowPOTSHPWB As DataRow = dst.Tables("POTSHPWB").Rows.Find(New Object() {wbName, wsName, PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
        If rowPOTSHPWB Is Nothing Then
            rowPOTSHPWB = dst.Tables("POTSHPWB").NewRow
            rowPOTSHPWB.Item("WORKBOOK") = wbName
            rowPOTSHPWB.Item("WORKSHEET") = wsName
            rowPOTSHPWB.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHPWB.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            rowPOTSHPWB.Item("CONTAINER_NO") = rowPOTSHIP2.Item("CONTAINER_NO") & ""
            rowPOTSHPWB.Item("IMPORTED") = "0"
            dst.Tables("POTSHPWB").Rows.Add(rowPOTSHPWB)
        End If
        Return rowPOTSHPWB
    End Function

    Sub Handle_Prepacks(ws As SpreadsheetGear.IWorksheet,
                        ByRef eMsg As String,
                        r As Integer,
                        poadj As Integer,
                        PO_ORDER_NO As String,
                        PO_REFERENCE As String,
                        COMM_INV_NO As String,
                        BOL_NO As String,
                        CONTAINER_NO As String,
                        COLOR_CODEs As List(Of String))

        Dim PO_ORDER_NO_LAST As String = PO_ORDER_NO
        Dim PO_REFERENCE_LAST As String = PO_REFERENCE

        Dim wbName As String = ws.Workbook.Name
        Dim wsName As String = ws.Name
        Dim PACKING_LNO As Integer = 0

        Dim CARTON_NO_START As Integer = Val(Trim(ws.Cells(r, 0).Text & ""))

        Dim eMsgsStartEnd As New List(Of String)

        Dim isRange As Boolean = False
        Dim STYLE_CODEs() As String = New String() {}

        Do While CARTON_NO_START > 0
            Dim STYLE_RANGE As String = ws.Cells(r, 3 + poadj).Text

            If STYLE_RANGE.Contains("-") Then
                STYLE_CODEs = Split(STYLE_RANGE, "-")
                isRange = True
            End If

            Dim TOTAL_CTNS As Integer = Val(ws.Cells(r, 6 + poadj).Text & "")
            Dim PER_CTN_PCS As Integer = Val(ws.Cells(r, 7 + poadj).Text & "")
            Dim TOTAL_PCS As Integer = Val(ws.Cells(r, 9 + poadj).Text & "")
            Dim NW As Decimal = Val(ws.Cells(r, 14 + poadj).Text & "")
            Dim MEAS As String = ws.Cells(r, 10 + poadj).Text & ""
            Dim rowPOTSHIP2 As DataRow = Get_Shipment_Line(COMM_INV_NO, BOL_NO, CONTAINER_NO)
            Dim rowPOTSHPWB As DataRow = Get_Worksheet_Container_Line(wbName, wsName, rowPOTSHIP2)

            'Dim PO_ORDER_NO As String = ""
            If poadj = 1 Then

                ' I DON'T THINK THAT THIS EVER WORKED BECAUSE WE SHOULD HAVE BEEN LOOKING at ws.Cells(r, 4).Text & ""
                'PO_REFERENCE = ws.Cells(r, 3).Text & ""

                'Dim rowPOTORDR1 As DataRow = Get_PO_Header_By_Ref_No(PO_REFERENCE)

                'If rowPOTORDR1 IsNot Nothing Then
                '    PO_ORDER_NO = rowPOTORDR1.Item("PO_ORDER_NO")
                '    Got_PO_Lock(PO_ORDER_NO)
                'End If

                'If Not packingListPOs.Contains(PO_ORDER_NO) Then
                '    packingListPOs.Add(PO_ORDER_NO)
                'End If

            Else
                PO_ORDER_NO = dicWorksheetPOs(ws.Name)
            End If



            Dim sqlw As String = ""
            If COLOR_CODEs IsNot Nothing Then
                For Each COLOR_CODE As String In COLOR_CODEs
                    sqlw &= " OR COLOR_CODE = '" & COLOR_CODE & "'"
                Next
                sqlw = Mid(sqlw, 5)
            End If

            If sqlw = "" Then
                sqlw = "PO_QTY_OPN > 0"
            Else
                sqlw = "(" & sqlw & ") AND PO_QTY_OPN > 0"
            End If

            If isRange Then
                sqlw &= $" and (STYLE_CODE >= '{STYLE_CODEs(0)}' and STYLE_CODE <= '{STYLE_CODEs(1)}')"
            End If

            Dim tblPOTORDR2 As DataTable = dicPOTORDR2(PO_ORDER_NO)
            Dim PO_QTY_OPEN_TOT As Integer = Val(tblPOTORDR2.Compute("SUM(PO_QTY_OPN)", sqlw) & "")

            rowPOTSHPWB.Item("IS_PPK") = "1"
            rowPOTSHPWB.Item("TOTAL_CTNS") = TOTAL_CTNS
            rowPOTSHPWB.Item("NW") = NW
            If MEAS <> "" Then
                Dim measCM As String = Validate_Carton_Dimensions(MEAS).CTN_DIMS_CM
                rowPOTSHPWB.Item("MEAS") = measCM
            End If

            Dim PCT_SHP As Decimal = 0
            If PO_QTY_OPEN_TOT <> 0 Then PCT_SHP = (TOTAL_CTNS * PER_CTN_PCS) / PO_QTY_OPEN_TOT
            If PO_QTY_OPEN_TOT <> 0 Then PCT_SHP = TOTAL_PCS / PO_QTY_OPEN_TOT
            If PCT_SHP > 1 Then
                MsgBox("Warning - Total Packing Units (" & CStr(TOTAL_PCS) & ") for PO Reference " & PO_REFERENCE & vbCrLf & " Is greater than Total Open Qty on PO (" & CStr(PO_QTY_OPEN_TOT) & ")")
                PCT_SHP = 1
            End If

            'Stop
            'PCT_SHP = 0.5

            ' Dim PACKING_LNO As Integer = 0
            For Each rowPOTORDR2 As DataRow In tblPOTORDR2.Select(sqlw, "PO_ORDER_NO,PO_ORDER_LNO")
                PACKING_LNO += 1
                Dim PO_SHIPMENT_NO As String = rowPOTSHIP2.Item("PO_SHIPMENT_NO")
                Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
                PO_ORDER_NO = rowPOTORDR2.Item("PO_ORDER_NO")
                Dim PO_ORDER_LNO As Integer = Val(rowPOTORDR2.Item("PO_ORDER_LNO") & "")
                Dim splitLine As Boolean = False
                Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowPOTORDR2.Item("COLOR_CODE")
                Dim validStyle As Boolean = Valid_Style(STYLE_CODE)
                Dim validColor As Boolean = Valid_Color(COLOR_CODE)
                Dim PO_QTY_OPEN As Integer = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")
                Dim sqlTAKEN As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' AND PO_ORDER_LNO  = " & PO_ORDER_LNO
                Dim PO_QTY_OPEN_Taken As Integer = Val(dst.Tables("POTSHPXL").Compute("SUM(TOTAL_PCS)", sqlTAKEN) & "")
                Dim STYLE_PCS As Integer = CInt(PO_QTY_OPEN * PCT_SHP)

                Dim TOTAL_CTN As Integer = Val(ws.Cells(r, 6 + poadj).Text & "")

                Dim CTN_NO_START As Integer = Val(Trim(ws.Cells(r, 0).Text & ""))
                Dim CTN_NO_END As Integer = Val(Trim(ws.Cells(r, 2).Text & ""))

                If CTN_NO_END - CTN_NO_START + 1 <> TOTAL_CTN Then
                    Dim eMsgStartEnd As String = "Problem with Starting/Ending Carton Nos (" & CStr(CTN_NO_START) & "/" & CStr(CTN_NO_END) & ") vs Total Cartons (" & CStr(TOTAL_CTN) & ") on " & ws.Name
                    If Not eMsgsStartEnd.Contains(eMsgStartEnd) Then
                        eMsgsStartEnd.Add(eMsgStartEnd)
                        MsgBox(eMsgStartEnd, MsgBoxStyle.OkOnly, "Please do NOT Update without Fixing XLS")
                    End If
                End If

                Dim rowPOTSHPXL As DataRow = Get_Packing_Line_By_Lno(PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE, PACKING_LNO)
                rowPOTSHPXL.Item("CTN_NO_START") = Val(Trim(ws.Cells(r, 0).Text & ""))
                rowPOTSHPXL.Item("CTN_NO_END") = Val(Trim(ws.Cells(r, 2).Text & ""))
                rowPOTSHPXL.Item("PO_ORDER_NO") = PO_ORDER_NO
                rowPOTSHPXL.Item("CONTAINER_NO") = CONTAINER_NO
                rowPOTSHPXL.Item("SIZE") = Trim(ws.Cells(r, 5 + poadj).Text & "")
                rowPOTSHPXL.Item("TOTAL_CTN") = TOTAL_CTNS
                rowPOTSHPXL.Item("PER_CTN") = PER_CTN_PCS
                rowPOTSHPXL.Item("TOTAL_PCS") = STYLE_PCS
                rowPOTSHPXL.Item("GW") = Val(ws.Cells(r, 13 + poadj).Text & "")
                rowPOTSHPXL.Item("NW") = Val(ws.Cells(r, 14 + poadj).Text & "")
                rowPOTSHPXL.Item("TTL_GW") = Val(ws.Cells(r, 15 + poadj).Text & "")
                rowPOTSHPXL.Item("TTL_NW") = Val(ws.Cells(r, 16 + poadj).Text & "")
                Dim measCM As String = Validate_Carton_Dimensions(ws.Cells(r, 10 + poadj).Text & "").CTN_DIMS_CM
                rowPOTSHPXL.Item("MEAS") = measCM
                rowPOTSHPXL.Item("WORKBOOK") = wbName
                rowPOTSHPXL.Item("WORKSHEET") = ws.Name

                rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_OPEN - PO_QTY_OPEN_Taken

                If STYLE_PCS > 0 Then
                    Dim packingLine As poPackingLine = Get_Packing_Line_By_Qty(PO_ORDER_NO, tblPOTORDR2, STYLE_CODE, COLOR_CODE, STYLE_PCS)
                    If packingLine.eMsg = "" Then
                        rowPOTSHPXL.Item("PO_ORDER_LNO") = Val(packingLine.rowPOTORDR2.Item("PO_ORDER_LNO") & "")
                        rowPOTSHPXL.Item("IS_SPLIT") = IIf(packingLine.splitLine, "1", "0")
                        If Not rowPOTSHPXL.RowState = DataRowState.Added Then
                            dst.Tables("POTSHPXL").Rows.Add(rowPOTSHPXL)
                        End If
                    Else
                        eMsg &= packingLine.eMsg & ws.Name & ", PO " & PO_ORDER_NO & ", Style " & STYLE_CODE & ", Color " & COLOR_CODE & ", Qty " & CStr(STYLE_PCS) & ", Row " & r + 1 & ": " & ws.Cells(r, 9 + poadj).Text & vbCrLf
                        Log_Import_Error(ws, packingLine.eMsg, STYLE_CODE, COLOR_CODE, STYLE_PCS, PO_ORDER_NO, PO_REFERENCE, COMM_INV_NO, BOL_NO, CONTAINER_NO, PO_SHIPMENT_LNO)
                        ' Exit For
                    End If
                Else
                    ' MsgBox("WHY 0")
                    eMsg &= "Cannot Locate Open PO to pack" & vbCrLf & ws.Name & ", PO " & PO_ORDER_NO & ", Style " & STYLE_CODE & ", Color " & COLOR_CODE & ", Qty " & CStr(STYLE_PCS) & ", Row " & r + 1 & ": " & ws.Cells(r, 9 + poadj).Text & vbCrLf
                    Log_Import_Error(ws, "Cannot Locate Open PO to pack", STYLE_CODE, COLOR_CODE, STYLE_PCS, PO_ORDER_NO, PO_REFERENCE, COMM_INV_NO, BOL_NO, CONTAINER_NO, PO_SHIPMENT_LNO)
                    ' Exit For
                End If

            Next
            r += 1
            CARTON_NO_START = Val(Trim(ws.Cells(r, 0).Text & ""))
        Loop

        PO_ORDER_NO = PO_ORDER_NO_LAST
        PO_REFERENCE = PO_REFERENCE_LAST
    End Sub
    Function Workbook_Is_Valid(wb As SpreadsheetGear.IWorkbook, ByRef eMsg As String) As Boolean
        Dim isValid As Boolean = True

        If wb.Worksheets.Count = 0 Then
            eMsg &= "No Sheets Found" & vbCrLf
            Return False
        End If

        For wsi As Integer = 0 To wb.Worksheets.Count - 1
            If Not Worksheet_Is_Valid(wb.Worksheets(wsi), eMsg, wsi) Then
                isValid = False
            End If
        Next

        Return isValid
    End Function
    Function Worksheet_Is_Valid(ws As SpreadsheetGear.IWorksheet, ByRef eMsg As String, sheetIndex As Integer) As Boolean
        Dim isValid As Boolean = True
        Dim checkCell As String = ws.Cells(9, 7).Text & ""
        Dim poRefCol As Integer = IIf(checkCell.StartsWith("P.O. NO"), 7, 8)
        Dim PO_REFERENCE As String = Trim(Replace((ws.Cells(9, poRefCol).Text & ""), ":", ""))
        ' Dim PO_REFERENCE2 As String = Trim(Replace(ws.Cells(9, 6).Text & "", ":", ""))
        Dim rowPOTORDR1 As DataRow = Get_PO_Header_By_Ref_No(PO_REFERENCE)
        Dim PO_ORDER_NO As String = ""
        If rowPOTORDR1 IsNot Nothing Then
            PO_ORDER_NO = rowPOTORDR1.Item("PO_ORDER_NO")
        Else
            PO_REFERENCE = Trim(Replace((ws.Cells(0, 13).Text & ""), "FOR PO #", ""))
            rowPOTORDR1 = Get_PO_Header_By_Ref_No(PO_REFERENCE)
            If rowPOTORDR1 IsNot Nothing Then
                PO_ORDER_NO = rowPOTORDR1.Item("PO_ORDER_NO")
            Else
                'eMsg = IIf(PO_REFERENCE <> "", "Invalid", "Missing") & " PO Reference on sheet " & sheetIndex
                'Log_Import_Error(ws, eMsg)
                'isValid = False
                PO_ORDER_NO = ""
            End If
        End If
        If isValid Then
            Record_Sheet_PO(ws.Name, PO_ORDER_NO)
        Else
            If ASCMAIN1.Running_in_VS Then Stop
        End If
        Return isValid
    End Function

    Function Validate_Carton_Dimensions(cd As String) As CARTON_DIMENSIONS
        Dim inInches As Boolean = (InStr(cd, "INCH") > 0 Or InStr(cd, Chr(34)) > 0)
        Dim clsDims As New CARTON_DIMENSIONS
        Dim cleanDims As String = Trim(Replace(Replace(Replace(cd, "INCH", ""), " ", ""), Chr(34), "")).ToUpper
        Dim arrayDims As String() = cleanDims.Split("X")
        If arrayDims.Length <> 3 Then
            clsDims.CTN_DIMS_IN = ""
            clsDims.CTN_DIMS_CM = ""
            Return clsDims
        Else
            Dim dimIn As String = ""
            Dim dimCm As String = ""
            For Each d As String In arrayDims

                If d.EndsWith("'") Then
                    d = Mid(d, 1, d.Length - 1)
                    d = CStr(Val(d) * 12)
                End If
                Dim dIN As Double = IIf(inInches, CDbl(d), CentimetersToInches(CDbl(d)))
                Dim dCM As Double = IIf(inInches, InchesToCentimeters(CDbl(d)), CDbl(d))
                dimIn &= dIN.ToString() & "X"
                dimCm &= dCM.ToString() & "X"
            Next
            clsDims.CTN_DIMS_IN = dimIn.TrimEnd("X")
            clsDims.CTN_DIMS_CM = dimCm.TrimEnd("X")
        End If
        Return clsDims
    End Function
    Sub Initialize_Workbook_Memory()
        dicStyleDesc.Clear()
        dicColorDesc.Clear()
        POTORDR1_added.Clear()
    End Sub
    Sub Initialize_Worksheet_Memory()

    End Sub
    Sub Log_Import_Error(ws As SpreadsheetGear.IWorksheet, eMsg As String,
                         STYLE_CODE As String,
                         COLOR_CODE As String,
                         QTY As Int64,
                         PO_ORDER_NO As String,
                         PO_REFERENCE As String,
                         COMM_INV_NO As String,
                         BOL_NO As String,
                         CONTAINER_NO As String,
                         PO_SHIPMENT_LNO As Int64,
                         Optional xlsRef As String = "")
        Dim wbName As String = ws.Workbook.Name
        Dim errorLNO As Int64 = Val(dst.Tables("POTSHPIE").Compute("MAX(IE_LNO)", "") & "") + 1
        Dim rowPOTSHPIE As DataRow = dst.Tables("POTSHPIE").NewRow
        With rowPOTSHPIE
            .Item("WORKBOOK") = wbName
            .Item("WORKSHEET") = ws.Name
            .Item("IE_LNO") = errorLNO
            .Item("ERROR_MSG") = eMsg
            .Item("XLS_REF") = xlsRef

            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("QTY") = QTY
            .Item("PO_ORDER_NO") = PO_ORDER_NO
            .Item("PO_REFERENCE") = PO_REFERENCE

            .Item("COMM_INV_NO") = COMM_INV_NO
            .Item("BOL_NO") = BOL_NO
            .Item("CONTAINER_NO") = CONTAINER_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
        End With


        dst.Tables("POTSHPIE").Rows.Add(rowPOTSHPIE)
    End Sub

    Function InchesToCentimeters(ByVal inches As Double) As Double
        Return inches * 2.54
    End Function

    Function CentimetersToInches(ByVal centimeters As Double) As Double
        Return centimeters / 2.54
    End Function

    Sub Record_Sheet_PO(sheetName As String, PO_ORDER_NO As String)
        If PO_ORDER_NO <> "" Then
            If Not packingListPOs.Contains(PO_ORDER_NO) Then
                packingListPOs.Add(PO_ORDER_NO)
            End If
            If Not dicWorksheetPOs.ContainsKey(sheetName) Then
                dicWorksheetPOs.Add(sheetName, PO_ORDER_NO)
            End If
        End If
    End Sub

    Function Got_PO_Locks(ByRef eMsg As String) As Boolean

        Dim gotLocks As Boolean = True
        For Each PO_ORDER_NO As String In packingListPOs

            If Not Got_PO_Lock(PO_ORDER_NO) Then
                gotLocks = False
            End If

        Next
        Return gotLocks
    End Function

    Function Got_PO_Lock(PO_ORDER_NO As String, Optional show_message As Boolean = False) As Boolean
        Dim gotLock As Boolean = True
        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, False, show_message, False) Then
            EMsg &= "Could not lock PO " & PO_ORDER_NO & " for editing."
            gotLock = False
        Else
            If Not dicPOTORDR1.ContainsKey(PO_ORDER_NO) Then
                Dim rowPOTORDR1 As DataRow = ASCDATA1.GetDataRow("Select * from POTORDR1 WHERE PO_ORDER_NO = :PARM1 AND PO_STATUS = 'O'", "V", PO_ORDER_NO)
                dicPOTORDR1.Add(PO_ORDER_NO, rowPOTORDR1)
                Dim tblPOTORDR2 As DataTable = ASCDATA1.GetDataTable("Select * from POTORDR2 WHERE PO_ORDER_NO = :PARM1", "", "V", PO_ORDER_NO)
                dicPOTORDR2.Add(PO_ORDER_NO, tblPOTORDR2)
            End If
        End If

        Return gotLock

    End Function

    Function Get_PO_Header_By_Ref_No(PO_REFERENCE As String) As DataRow
        ASCMAIN1.sql = "Select * from POTORDR1 WHERE PO_REFERENCE = :PARM1 AND PO_STATUS = 'O'"
        Return ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", PO_REFERENCE)
    End Function

    Function Valid_Style(STYLE_CODE As String) As Boolean
        ASCMAIN1.sql = "Select * from ICTSTYL1 WHERE STYLE_CODE = :PARM1 "
        Dim rowICTSTYL1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", STYLE_CODE)
        If rowICTSTYL1 IsNot Nothing AndAlso Not dicStyleDesc.ContainsKey(STYLE_CODE) Then
            Dim STYLE_DESC As String = rowICTSTYL1.Item("STYLE_DESC") & ""
            dicStyleDesc.Add(STYLE_CODE, STYLE_DESC)
        End If
        Return (rowICTSTYL1 IsNot Nothing)
    End Function
    Function Valid_Color(colorCode As String) As Boolean
        Dim COLOR_CODE As String = IIf(colorCode.IndexOf("(") > -1, Extract_Color_Code(colorCode), colorCode)
        ASCMAIN1.sql = "Select * from ICTCOLR1 WHERE COLOR_CODE = :PARM1 "
        Dim rowICTCOLR1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", COLOR_CODE)
        If rowICTCOLR1 IsNot Nothing AndAlso Not dicColorDesc.ContainsKey(COLOR_CODE) Then
            Dim COLOR_DESC As String = rowICTCOLR1.Item("COLOR_DESC") & ""
            dicColorDesc.Add(COLOR_CODE, COLOR_DESC)
        End If
        Return (rowICTCOLR1 IsNot Nothing)
    End Function
    Function Extract_Color_Code(ccDesc As String) As String
        Dim a As Integer = ccDesc.IndexOf("(")
        Dim cCode As String = ccDesc
        If a > -1 Then
            Dim ccDesc2 As String = ccDesc & ")"
            Dim b As Integer = ccDesc2.IndexOf(")", a + 1)
            cCode = ccDesc.Substring(a + 1, b - a - 1)
            'Dim b As Integer = ccDesc.IndexOf(")", a + 1)
            'If b > -1 Then
            '    cCode = ccDesc.Substring(a + 1, b - a - 1)
            'End If
        End If
        Return cCode
    End Function

    Function Packing_Import_Allowed() As Boolean
        Dim importAllowed As Boolean = False
        If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") And EntryMode = "N" Then
            Dim POTSHIP2_empty As Boolean = (dst.Tables("POTSHIP2").Rows.Count = 0)
            If packingFromXLS Or POTSHIP2_empty Then
                importAllowed = True
            End If
        End If
        Return importAllowed
    End Function
    Function Manual_Entry_Allowed() As Boolean
        Dim manualEntryAllowed As Boolean = False
        If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") And (EntryMode = "N" Or EntryMode = "E") And ScreenMode And ship_entry Then
            Dim POTSHIP2_empty As Boolean = (dst.Tables("POTSHIP2").Rows.Count = 0)
            If Not packingFromXLS Then
                manualEntryAllowed = True
            End If
        End If
        Return manualEntryAllowed
    End Function
    Sub Print_Invoice()
        ' NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM
        ' AND SHOULD BE USING THE DATALAYER OF SORSHIP1

        'Fill_Records("SOTSVIA1", SHIP_CODE)

        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Dim RPT As String = "SORSHIP1" ' unneccesary if Report Name is Like Form Name
        'Generate_Report(RPT, "Shipper Invoice Report", , , , , False)
        'Print_Report_End()

        Dim REPORTFILE As String = "SORINVP1"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        Dim RPT As String = "SORINVP1"
        Dim AR_PARM_INVOICE_RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If AR_PARM_INVOICE_RPT <> "" Then RPT = AR_PARM_INVOICE_RPT ' "SORINVP1"

        Dim ORDR_NO As String = grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and SOTORDR1.ORDR_NO = '" & ORDR_NO & "'", "1", "O"})
        'For Each row As DataRow In REPORTS(REPORTFILE).clsASCBASE1.dst.Tables("SOTINVH1").Select("")
        '    row.Item("INV_NO") = "A123"
        'Next


        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()

            .CR_params.Add("SUBT", "")
            .CR_params.Add("CONS_INV", "0")
            .Generate_Report(RPT, "Sales Order Confirmation", , True, , , , , False)
            .Print_Report_End()
        End With
    End Sub


    Function Create_Carton_for_Selected_Styles(PO_SHIPMENT_LNO As Int64, Optional PARTIAL_QTY As Int64 = 0) As DataRow

        Dim CARTON_NO As Int32 = Val(dst.Tables("POTSHIP7").Compute("MAX (CARTON_NO)", "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)) & "") 'THIS IS THE CARTON TYPE
        Dim CARTONS As Int32 = 0
        Dim rowPOTSHIP7 As DataRow = Nothing

        If Not packingFromXLS Then
            rowPOTSHIP7 = dst.Tables("POTSHIP7").NewRow
            rowPOTSHIP7.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP7.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            CARTON_NO += 1
            rowPOTSHIP7.Item("CARTON_NO") = CARTON_NO
            'rowPOTSHIP7.Item("CARTON_COMMENTS") = ""
            dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIPR.Selected.Rows
            If packingFromXLS Then
                Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value

                For Each rowPOTSHPXL As DataRow In dst.Tables("POTSHPXL").Select("PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " AND STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                    CARTON_NO += 1
                    Dim WORKBOOK As String = rowPOTSHPXL.Item("WORKBOOK")
                    Dim WORKSHEET As String = rowPOTSHPXL.Item("WORKSHEET")
                    Dim rowPOTSHPWB As DataRow = dst.Tables("POTSHPWB").Rows.Find(New Object() {WORKBOOK, WORKSHEET})
                    Dim isPPK As Boolean = rowPOTSHPWB.Item("IS_PPK") & "" = "1"
                    'rowPOTSHIP7 = Create_Carton_Type_From_XLS_Packing_Line(PO_SHIPMENT_LNO, CARTON_NO, rowPOTSHPXL, isPPK)
                Next
            Else
                Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value

                Dim rowPOTSHIP8 As DataRow = dst.Tables("POTSHIP8").NewRow
                rowPOTSHIP8.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                rowPOTSHIP8.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                rowPOTSHIP8.Item("CARTON_NO") = CARTON_NO
                rowPOTSHIP8.Item("STYLE_CODE") = STYLE_CODE
                rowPOTSHIP8.Item("COLOR_CODE") = COLOR_CODE
                dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)

                Dim rowPOTSHIPR As DataRow = dst.Tables("POTSHIPR").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE})
                Dim rowPOTSHIP3 As DataRow = rowPOTSHIPR.GetChildRows("POTSHIPR_POTSHIP3")(0)

                Dim rows() As DataRow = dst.Tables("POTSHIP3").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")

                Dim CARTON_PACK_QTY As Int32 = IIf(rows.Length = 0, 1, Val(rows(0).Item("CARTON_PACK_QTY") & ""))
                If CARTON_PACK_QTY = 0 Then CARTON_PACK_QTY = 1
                'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                'Dim CARTON_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
                If PARTIAL_QTY <> 0 Then
                    rowPOTSHIP8.Item("QTY") = PARTIAL_QTY
                    rowPOTSHIP7.Item("CARTONS") = 1
                Else
                    rowPOTSHIP8.Item("QTY") = CARTON_PACK_QTY
                    If CARTONS = 0 Then
                        Dim QTY As Int32 = Val(rowPOTSHIPR.Item("QTY_VAR") & "")
                        If CARTON_PACK_QTY <> 0 And QTY > 0 Then
                            CARTONS = QTY \ CARTON_PACK_QTY
                            rowPOTSHIP7.Item("CARTONS") = CARTONS
                        End If
                    End If
                End If
            End If

            For Each grow7 As UltraWinGrid.UltraGridRow In grdPOTSHIP7.Rows
                If Val(grow7.Cells("CARTON_NO").Value) = CARTON_NO Then
                    grdPOTSHIP7.ActiveRow = grow7
                    Check_for_PPK(grow7)
                    'If grow7.DataChanged Then
                    '    grow7.Update()
                    'End If
                End If
            Next

        Next

        Dim PO_ORDER_NO_carton As String = Get_PO_for_Carton(PO_SHIPMENT_NO, PO_SHIPMENT_LNO, CARTON_NO)
        If CARTON_DIMS_by_PO.ContainsKey(PO_ORDER_NO_carton) Then
            rowPOTSHIP7.Item("CARTON_DIMS") = CARTON_DIMS_by_PO(PO_ORDER_NO_carton)
            rowPOTSHIP7.Item("CARTON_VOLUME") = Get_Volume_from_Dims(CARTON_DIMS_by_PO(PO_ORDER_NO_carton))
        Else
            Dim CARTON_DIMS As String = rowPOTSHIP7.Item("CARTON_DIMS") & ""
            If CARTON_DIMS <> "" Then
                CARTON_DIMS_by_PO.Add(PO_ORDER_NO_carton, CARTON_DIMS)
            End If
        End If

        Return rowPOTSHIP7
    End Function
    Function Create_Carton_Type_From_XLS_Packing_Line(PO_SHIPMENT_LNO As Int64, CARTON_NO As Integer, rowPOTSHPXL As DataRow, Optional isPPK As Boolean = False, Optional new_carton_type As Boolean = True) As DataRow

        Dim STYLE_CODE As String = rowPOTSHPXL.Item("STYLE_CODE")
        Dim COLOR_CODE As String = rowPOTSHPXL.Item("COLOR_CODE")
        Dim SIZE As String = rowPOTSHPXL.Item("SIZE")
        Dim TOTAL_CTN As Int32 = Val(rowPOTSHPXL.Item("TOTAL_CTN") & "")
        Dim TOTAL_PCS As Int32 = Val(rowPOTSHPXL.Item("TOTAL_PCS") & "")
        Dim CARTON_PACK_QTY As Int32 = Val(rowPOTSHPXL.Item("PER_CTN") & "")
        Dim rowPOTSHIP7 As DataRow = Nothing

        '  If ASCMAIN1.Running_in_VS And STYLE_CODE = "WM194057" Then Stop

        If new_carton_type Then
            rowPOTSHIP7 = dst.Tables("POTSHIP7").NewRow()
            rowPOTSHIP7.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP7.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            rowPOTSHIP7.Item("CARTON_NO") = CARTON_NO
            Dim MEAS As String = rowPOTSHPXL.Item("MEAS")
            Dim NW As Decimal = Val(rowPOTSHPXL.Item("NW") & "")
            rowPOTSHIP7.Item("CARTON_WEIGHT") = NW
            rowPOTSHIP7.Item("CARTON_DIMS") = MEAS
            If MEAS <> "" Then
                rowPOTSHIP7.Item("CARTON_VOLUME") = Get_Volume_from_Dims(MEAS)
            End If
            rowPOTSHIP7.Item("CARTONS") = TOTAL_CTN
            dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
        End If

        If isPPK Then
            CARTON_PACK_QTY = TOTAL_PCS / TOTAL_CTN
        End If

        Dim rowPOTSHIPR As DataRow = dst.Tables("POTSHIPR").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE})
        If rowPOTSHIPR Is Nothing Then
            Create_POTSHIPR(PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE)
        End If

        Dim rowPOTSHIP8 As DataRow = dst.Tables("POTSHIP8").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, CARTON_NO, STYLE_CODE, COLOR_CODE})
        If rowPOTSHIP8 IsNot Nothing Then
            rowPOTSHIP8.Item("QTY") = Val(rowPOTSHIP8.Item("QTY") & "") + CARTON_PACK_QTY
        Else
            rowPOTSHIP8 = dst.Tables("POTSHIP8").NewRow
            With rowPOTSHIP8
                .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                .Item("CARTON_NO") = CARTON_NO
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("QTY") = CARTON_PACK_QTY
            End With
            dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)
        End If

        Return rowPOTSHIP7
    End Function
#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PO_SHIPMENT_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If receipt_mode Then
                        Click_Command("Select", e)
                    Else
                        Click_Command("View", e)
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                Setup_Warehouse_Attributes()
            Case "CUST_STORE_NO"
                If ctl.Text & "" = "" Then Return
                'Store Adress needs to be updated
                Dim row As DataRow = LookUp("ARTCUST2", New String() {"171659", "MK", ctl.Text & ""})
                Absx1.txtFor("ADDRESS").Text = row("CUST_NAME") & vbCrLf & row("CUST_ADDR1") & vbCrLf & row("CUST_CITY") & " " & row("CUST_STATE") & ", " & row("CUST_ZIP_CODE")

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PO_SHIPMENT_NO"
                If receipt_mode Then
                    Click_Command("Select")
                Else
                    Click_Command("View")
                End If
        End Select
    End Sub

    Public Overrides Sub CheckedChanged_Special(COLUMN_NAME As String, chk As Infragistics.Win.UltraWinEditors.UltraCheckEditor)
        MyBase.CheckedChanged_Special(COLUMN_NAME, chk)
        Select Case COLUMN_NAME
            Case "REVIEW"
                If chkFlag.Checked Then
                    If chkCostComplete.Checked Then
                        MsgBox("You May Not Mark A Shipment For Review That Is Also Costed Complete.", MsgBoxStyle.OkOnly, "Un-Check The Cost Complete Option")
                        chkFlag.Checked = False
                    End If
                End If

            Case "COST_COMPLETE"
                If chkCostComplete.Checked Then
                    If chkFlag.Checked Then
                        MsgBox("You May Not Mark A Shipment Complete That Also Requires Review.", MsgBoxStyle.OkOnly, "Un-Check The Needs Review Option")
                        chkCostComplete.Checked = False
                    End If
                End If
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "FREIGHT_ENTERED_BY"
                Toggle_Columns()
        End Select
    End Sub
#End Region

    Private Sub grdPOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTSHIPX.DoubleClickRow
        If grdPOTSHIPX.ActiveRow IsNot Nothing AndAlso grdPOTSHIPX.ActiveRow.IsDataRow Then
            Absx1.txtFor("PO_SHIPMENT_NO").Text = grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Text
            If receipt_mode Then
                Click_Command("Select")
            Else
                Click_Command("View")
            End If
        End If
    End Sub

    Sub Load_WHTWRECX()
        Fill_Records("WHTWRECX")
        Sort_grdColumns(grdWHTWRECX, "PO_SHIPMENT_NO,CONTAINER_NO")
    End Sub

    Sub Load_POTSHIPX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If receipt_mode Then Get_Receipt_Data_from_3PL()

        ASCMAIN1.sql = sqlPOTSHIPX
        If Not InquiryMode Then
            If cost_calc Then
                ASCMAIN1.sql &= " and NVL(POTSHIP1.COST_COMPLETE,'0') <> '1'"
                grdPOTSHIPX.Text = "Shipments Not Costed Completely"
            Else

                ASCMAIN1.sql &= " and POTSHIP1.PO_SHIPMENT_NO in " _
                    & "(Select DISTINCT PO_SHIPMENT_NO from POTSHIP2 where PO_SHIP_STATUS = 'O')"
                grdPOTSHIPX.Text = "Shipments In Transit Not Received"

                If receipt_mode Then
                    If optReceiptType.Value = "BTB" Then
                        ASCMAIN1.sql &= " and POTSHIP1.PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from POTSHIP2 where PO_SHIP_STATUS = 'O' and ORDR_NO is Not Null)"
                        grdPOTSHIPX.Text &= " - Back-to-Back Shipments"
                    Else
                        ASCMAIN1.sql &= " and POTSHIP1.PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from POTSHIP2 where PO_SHIP_STATUS = 'O' and ORDR_NO is Null)"
                        grdPOTSHIPX.Text &= " - Warehouse Arrivals"
                    End If
                End If

            End If
        End If

        ASCDATA1.ExecuteSQL("Delete from " & POTSHIPX)

        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
            ASCMAIN1.sql &= " AND POTSHIP1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")"
        End If

        ASCMAIN1.sql = "Insert into " & POTSHIPX & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        'Fill_Records("POTSHIPX", "", True, ASCMAIN1.sql)
        Fill_Records("POTSHIPX")
        Sort_grdColumns(grdPOTSHIPX, "PO_SHIP_ETA".ToLower)

        If cost_calc Then
            ASCMAIN1.sql = "Select PO_SHIPMENT_NO, COUNT (*) LINES, SUM (DECODE(PO_SHIP_STATUS,'C',1,0)) LINES_REC" _
                & " from POTSHIP2 where PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from (" & POTSHIPX & "))" _
                & " group by PO_SHIPMENT_NO"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim PO_SHIPMENT_NO As String = row.Item("PO_SHIPMENT_NO")
                Dim rowPOTSHIPX As DataRow = dst.Tables("POTSHIPX").Rows.Find(PO_SHIPMENT_NO)
                rowPOTSHIPX.Item("LINES") = row.Item("LINES")
                rowPOTSHIPX.Item("LINES_REC") = row.Item("LINES_REC")
            Next
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.sql = "Select PO_SHIPMENT_NO, SUM (DECODE(WH_REC_STATUS,'C',1,0)) WH_LINES_REC" _
                & " from WHTWREC1 where PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from (" & POTSHIPX & "))" _
                & " group by PO_SHIPMENT_NO"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim PO_SHIPMENT_NO As String = row.Item("PO_SHIPMENT_NO")
                Dim rowPOTSHIPX As DataRow = dst.Tables("POTSHIPX").Rows.Find(PO_SHIPMENT_NO)
                rowPOTSHIPX.Item("WH_LINES_REC") = row.Item("WH_LINES_REC")
            Next
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub tabPOTSHIP1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        SETUP_tabPOTSHIP1()
    End Sub

    Sub SETUP_tabPOTSHIP1()

    End Sub

    Sub Calculate_Landed_Cost()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Calculating Landed Costs")

        Synch_TABLE_NAME("POTSHIP1")

        Dim DZ As Boolean = (optUD.Value = "D")

        lblMissingCBMs.Visible = False
        If optCOST_FRT_METHOD.Value & "" = "V" Then
            If dst.Tables("POTSHIP3").Select("CBM = 0").Length <> 0 Then
                optCOST_FRT_METHOD.ForeColor = Drawing.Color.Red
            End If
        End If

        Dim sqlP As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
        Dim TOTAL_FREIGHT As Decimal = Val(dst.Tables("POTSHIP4").Compute("SUM(FREIGHT_AMT)", sqlP) & "")
        Dim sqlp2 As String = " and ISNULL(PO_SHIPMENT_LNO_DIST,0) = 0"
        ' Note: TOTAL_TRUCKING is spread by Weight Factor across the entire Shipment
        ' If Frt is done by BOL, then there is a TRUKING amount that is spread to each BOL independently, down below.
        ' This is in Addition to any TRUCKING which may have been entered on the Customs/Duty/Trucking grid
        Dim TOTAL_TRUCKING As Decimal = Val(dst.Tables("POTSHIP4").Compute("SUM(TRUCKING)", sqlP) & "")
        If rowPOTSHIP1.Item("FREIGHT_ENTERED_BY") = "B" Or rowPOTSHIP1.Item("FREIGHT_ENTERED_BY") = "I" Then
            'TOTAL_TRUCKING = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_AMT)", sqlP & sqlp2 & " and LANDING_COST_DIST = 'T'") & "")
            TOTAL_TRUCKING = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_T)", sqlP & sqlp2) & "")
        End If

        Dim TOTAL_WEIGHT As Decimal = Val(rowPOTSHIP1.Item("TOTAL_WEIGHT_FACTOR") & "")
        Dim TOTAL_CBM As Decimal = Val(rowPOTSHIP1.Item("TOTAL_CBM") & "")
        ' Dim TOTAL_DUTY As Decimal = Val(rowPOTSHIP1.Item("TOTAL_DUTY") & "")
        Dim TOTAL_QTY_THIS_SHIPMENT As Int64 = Val(dst.Tables("POTSHIP3").Compute("SUM(PO_QTY_SR)", "") & "")

        Dim TOTAL_MISC_INVOICED As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_M)", sqlP & sqlp2) & "")
        Dim TOTAL_DUTY_INVOICED As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_D)", sqlP & sqlp2) & "")
        Dim TOTAL_CUSTOMS_INVOICED As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_W)", sqlP & sqlp2) & "")

        For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")

            With rowPOTSHIP2

                Dim PO_SHIPMENT_LNO As Integer = Val(.Item("PO_SHIPMENT_LNO") & "")

                Dim sqlp_LNO As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO_DIST = " & CStr(PO_SHIPMENT_LNO)
                Dim TOTAL_TRUCKING_LNO As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_T)", sqlp_LNO) & "")
                Dim TOTAL_MISC_INVOICED_LNO As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_M)", sqlp_LNO) & "")
                Dim TOTAL_DUTY_INVOICED_LNO As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_D)", sqlp_LNO) & "")
                Dim TOTAL_CUSTOMS_INVOICED_LNO As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_W)", sqlp_LNO) & "")

                Dim TOTAL_WEIGHT_B As Decimal = 0
                Dim TOTAL_CBM_B As Decimal = 0
                Dim TOTAL_FREIGHT_B As Decimal = 0
                Dim TOTAL_TRUCKING_BOL As Decimal = 0

                If rowPOTSHIP1.Item("FREIGHT_ENTERED_BY") = "B" Then
                    If Val(.Item("CBM") & "") = 0 Then
                        TOTAL_FREIGHT_B = (Val(.Item("TOTAL_WEIGHT") & "") * Val(.Item("CBM_RATE") & "")) + Val(.Item("BOL_FEE") & "")
                    Else
                        TOTAL_FREIGHT_B = (Val(.Item("CBM") & "") * Val(.Item("CBM_RATE") & "")) + Val(.Item("BOL_FEE") & "")
                    End If
                    TOTAL_WEIGHT_B = Val(rowPOTSHIP2.Item("TOTAL_WEIGHT_FACTOR") & "")
                    TOTAL_CBM_B = Val(rowPOTSHIP2.Item("TOTAL_CBM") & "")
                    TOTAL_TRUCKING_BOL = Val(.Item("TRUCKING") & "")
                End If

                Dim TOTAL_QTY_THIS_BOL As Decimal = 0

                For Each rowPOTSHIP3 As DataRow In rowPOTSHIP2.GetChildRows("POTSHIP2_POTSHIP3")
                    Dim PXWF As Decimal = 0
                    Dim PXWF_B As Decimal = 0
                    With rowPOTSHIP3
                        Dim PO_QTY_SR As Int64 = Val(.Item("PO_QTY_SR") & "")
                        Dim SUB_UNIT_PACK_QTY As Int64 = Val(.Item("SUB_UNIT_PACK_QTY") & "")

                        If optCOST_FRT_METHOD.Value & "" = "V" Then
                            PXWF = IIf(TOTAL_CBM = 0, 0, Val(.Item("CBM") & "") / TOTAL_CBM)
                            PXWF_B = IIf(TOTAL_CBM_B = 0, 0, Val(.Item("CBM") & "") / TOTAL_CBM_B)
                        Else
                            PXWF = IIf(TOTAL_WEIGHT = 0, 0, Val(.Item("EXT_WEIGHT_FACTOR") & "") / TOTAL_WEIGHT)
                            PXWF_B = IIf(TOTAL_WEIGHT_B = 0, 0, Val(.Item("EXT_WEIGHT_FACTOR") & "") / TOTAL_WEIGHT_B)
                        End If

                        If dst.Tables("POTSHIP3").Select("").Length = 1 Then
                            If PXWF = 0 Then PXWF = 1
                            If PXWF_B = 0 Then PXWF_B = 1
                        End If

                        If rowPOTSHIP1.Item("FREIGHT_ENTERED_BY") = "C" Then
                            Dim CONTAINER_NO As String = rowPOTSHIP2.Item("CONTAINER_NO") & ""
                            Dim sqlpC As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and CONTAINER_NO = '" & CONTAINER_NO & "'"
                            TOTAL_FREIGHT = Val(dst.Tables("POTSHIP4").Compute("SUM(FREIGHT_AMT)", sqlpC) & "")
                            If optCOST_FRT_METHOD.Value & "" = "V" Then
                                Dim TOTAL_CBM_THIS_BOL As Decimal = Val(dst.Tables("POTSHIP3").Compute("SUM(CBM)", sqlpC) & "")
                                Dim PXWF_THIS_BOL As Decimal = IIf(TOTAL_CBM_THIS_BOL = 0, 0, Val(.Item("CBM") & "") / TOTAL_CBM_THIS_BOL)
                                .Item("PO_COST_FREIGHT_IN") = IIf(PO_QTY_SR = 0, 0, System.Math.Round((TOTAL_FREIGHT * PXWF_THIS_BOL) / PO_QTY_SR, 6))
                            Else
                                Dim TOTAL_WEIGHT_THIS_BOL As Decimal = Val(dst.Tables("POTSHIP3").Compute("SUM(EXT_WEIGHT_FACTOR)", sqlpC) & "")
                                Dim PXWF_THIS_BOL As Decimal = IIf(TOTAL_WEIGHT_THIS_BOL = 0, 0, Val(.Item("EXT_WEIGHT_FACTOR") & "") / TOTAL_WEIGHT_THIS_BOL)
                                .Item("PO_COST_FREIGHT_IN") = IIf(PO_QTY_SR = 0, 0, System.Math.Round((TOTAL_FREIGHT * PXWF_THIS_BOL) / PO_QTY_SR, 6))
                            End If
                        ElseIf rowPOTSHIP1.Item("FREIGHT_ENTERED_BY") = "I" Then
                            Dim sqlpC As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and (ISNULL(PO_SHIPMENT_LNO_DIST,0) = 0 OR ISNULL(PO_SHIPMENT_LNO_DIST,0) = " & CStr(PO_SHIPMENT_LNO) & ")"
                            TOTAL_FREIGHT = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_F)", sqlpC) & "")
                            sqlpC = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
                            If optCOST_FRT_METHOD.Value & "" = "V" Then
                                Dim TOTAL_CBM_THIS_BOL As Decimal = Val(dst.Tables("POTSHIP3").Compute("SUM(CBM)", sqlpC) & "")
                                Dim PXWF_THIS_BOL As Decimal = IIf(TOTAL_CBM_THIS_BOL = 0, 0, Val(.Item("CBM") & "") / TOTAL_CBM_THIS_BOL)
                                .Item("PO_COST_FREIGHT_IN") = IIf(PO_QTY_SR = 0, 0, System.Math.Round((TOTAL_FREIGHT * PXWF) / PO_QTY_SR, 6))
                            Else
                                Dim TOTAL_WEIGHT_THIS_BOL As Decimal = Val(dst.Tables("POTSHIP3").Compute("SUM(EXT_WEIGHT_FACTOR)", sqlpC) & "")
                                Dim PXWF_THIS_BOL As Decimal = IIf(TOTAL_WEIGHT_THIS_BOL = 0, 0, Val(.Item("EXT_WEIGHT_FACTOR") & "") / TOTAL_WEIGHT_THIS_BOL)
                                .Item("PO_COST_FREIGHT_IN") = IIf(PO_QTY_SR = 0, 0, System.Math.Round((TOTAL_FREIGHT * PXWF) / PO_QTY_SR, 6))
                            End If
                        Else
                            .Item("PO_COST_FREIGHT_IN") = IIf(PO_QTY_SR = 0, 0, System.Math.Round((TOTAL_FREIGHT_B * PXWF_B) / PO_QTY_SR, 6))
                        End If

                        .Item("PO_COST_TRUCKING") = System.Math.Round((TOTAL_TRUCKING * PXWF + TOTAL_TRUCKING_BOL) / PO_QTY_SR, 6)

                        Dim a As Decimal = Val(.Item("DUTY_RATE") & "") / 100
                        If chkNoDuty.Checked Then a = 0
                        Dim b As Decimal = Val(.Item("PO_COST_VCOST") & "") + Val(.Item("PO_COST_MATLS") & "")
                        Dim c As Decimal = Val(.Item("PO_COST_QUOTA") & "") + Val(.Item("PO_COST_OTHER") & "")
                        Dim d As Decimal = Val(.Item("SUB_UNIT_PACK_QTY") & "")
                        Dim f As Decimal = Val(.Item("PO_COST_BUFFER") & "") / 100

                        Dim PO_COST_COMM As Decimal = Val(.Item("PO_COST_COMM") & "")
                        Dim PO_COST_COMM_paid_by_supplier As Decimal = 0
                        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                            ' THIS SHOULD BE CODED AS "IF POTORDR1.PO_COMM_CHGBACK_TO_SUPP = '1' THEN set to commission
                            ' this is to REDUCE DUTY BY THE AMT OF THE BUYERS COMMISSION
                            ' PO_COST_COMM_paid_by_supplier = (Val(.Item("PO_COST_VCOST") & "") + Val(.Item("PO_COST_OTHER") & "")) * PO_COST_COMM / 100

                            'Leslie Shalom Monday, October 19, 2015 3:09 PM 
                            ' I just noticed that when an invoice is subject to the 2% buyer’s commission, 
                            ' ABS calculates it against the VENDOR cost only 
                            ' as opposed to the VENDOR + OTHER costs 
                            ' which is truer to our experience 
                            ' and is usually included by the shipper/factory 
                            ' as part of the calculation of the buyer’s commission. 
                            ' Can you change the formula in ABS to reflect that?
                            'PO_COST_COMM_paid_by_supplier = Val(.Item("PO_COST_VCOST") & "") * PO_COST_COMM / 100
                            PO_COST_COMM_paid_by_supplier = (Val(.Item("PO_COST_VCOST") & "") + Val(.Item("PO_COST_OTHER") & "")) * PO_COST_COMM / 100

                        End If

                        TOTAL_QTY_THIS_BOL = TOTAL_QTY_THIS_BOL + PO_QTY_SR

                        Dim PO_COST_DUTY As Decimal = System.Math.Round(a * (b + c - PO_COST_COMM_paid_by_supplier), 6)
                        ' for NYA, take the commission out of the PO Cost.  
                        ' The 2% is a fee we charge to the supplier to cover our overhead costs for JA. 
                        ' We will pay less duty

                        .Item("PO_COST_DUTY") = PO_COST_DUTY
                        .Item("PO_COST_CUSTOMS") = 0
                        .Item("PO_COST_CUSTOMS") = System.Math.Round((Val(.Item("PO_COST_CUSTOMS") & "") + TOTAL_CUSTOMS_INVOICED * PXWF) / PO_QTY_SR, 6)
                        .Item("PO_COST_MISC") = 0
                        .Item("PO_COST_MISC") = System.Math.Round((Val(.Item("PO_COST_MISC") & "") + TOTAL_MISC_INVOICED * PXWF + TOTAL_MISC_INVOICED_LNO) / PO_QTY_SR, 6)

                        Dim otherCost As Decimal = (Val(.Item("PO_COST_OTHER") & "") + Val(.Item("PO_COST_QUOTA") & ""))

                        Dim LANDING_COST As Decimal = 0
                        If PO_QTY_SR <> 0 Then
                            LANDING_COST = System.Math.Round((Val(.Item("PO_COST_FREIGHT_IN") & "") + Val(.Item("PO_COST_DUTY") & "") + Val(.Item("PO_COST_TRUCKING") & "") + Val(.Item("PO_COST_MISC") & "") + Val(.Item("PO_COST_CUSTOMS") & "")) + otherCost, 6)
                        End If

                        Dim g As Decimal = System.Math.Round(LANDING_COST + Val(.Item("PO_COST_MATLS") & "") + (Val(.Item("PO_COST_VCOST") & "")), 6)
                        .Item("PO_COST_LANDED") = g + Val(.Item("COMMISSION_COST") & "") - PO_COST_COMM_paid_by_supplier + Val(.Item("PO_COST_QUOTA_DF") & "")
                    End With
                Next

                If TOTAL_QTY_THIS_SHIPMENT <> 0 Then
                    If Not chkNoDuty.Checked Then
                        If ASCMAIN1.CLIENT = "RGIX" Then
                        Else
                            ' Add Duty from Grid
                            Dim e As Decimal = TOTAL_DUTY_INVOICED / TOTAL_QTY_THIS_SHIPMENT
                            For Each rowPOTSHIP3 As DataRow In rowPOTSHIP2.GetChildRows("POTSHIP2_POTSHIP3")
                                rowPOTSHIP3.Item("PO_COST_DUTY") += System.Math.Round(e, 6)
                                rowPOTSHIP3.Item("PO_COST_LANDED") += System.Math.Round(e, 6)
                            Next
                        End If
                    End If

                    Dim m As Decimal = TOTAL_MISC_INVOICED / TOTAL_QTY_THIS_SHIPMENT
                    For Each rowPOTSHIP3 As DataRow In rowPOTSHIP2.GetChildRows("POTSHIP2_POTSHIP3")
                        If ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "VAN" Then
                            ' 20190617 email Leslie - does not want MISC mixed with Duty, but wants it included in Landed
                            ' LOOK ABOVE  .Item("PO_COST_MISC") IS ADDING IT TO PO_COST_MISC - SO WHY THEN DO WE NEED IT IN PO_COST_DUTY?
                        Else
                            rowPOTSHIP3.Item("PO_COST_DUTY") += System.Math.Round(m, 6)
                            rowPOTSHIP3.Item("PO_COST_LANDED") += System.Math.Round(m, 6)
                        End If

                    Next

                End If
            End With
        Next


        If Not chkNoDuty.Checked Then
            ' Replace Duty with Grid Total
            If ASCMAIN1.CLIENT = "RGIX" Then
                Dim TOTAL_DUTY_shipment As Decimal = Val(dst.Tables("POTSHIP3").Compute("SUM(TOTAL_DUTY)", "") & "")
                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
                    Dim PO_QTY_SR As Int64 = Val(rowPOTSHIP3.Item("PO_QTY_SR") & "")
                    Dim TOTAL_DUTY As Decimal = Val(rowPOTSHIP3.Item("TOTAL_DUTY") & "")
                    Dim TOTAL_DUTY_NEW As Decimal = 0
                    If TOTAL_DUTY_shipment <> 0 Then TOTAL_DUTY_NEW = TOTAL_DUTY_INVOICED * TOTAL_DUTY / TOTAL_DUTY_shipment

                    Dim e As Decimal = (TOTAL_DUTY_NEW - TOTAL_DUTY) / PO_QTY_SR

                    rowPOTSHIP3.Item("PO_COST_DUTY") += System.Math.Round(e, 6)
                    rowPOTSHIP3.Item("PO_COST_LANDED") += System.Math.Round(e, 6)
                Next
            Else

            End If
        End If

        rowPOTSHIP1.Item("COST_IND") = "1"
        Set_Landed_Cost_Needs_to_be_Calculated_Indicator(True)

        '  Calculate_DUTY_DIST()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Receive_BOL()
        'if I'm dereceiving I need to hit po details with qty adjustments

        Absx1.dteFor("PO_DATE_RECEIVED").ReadOnly = True
        ' WHY DO I NEED TO DO THIS NOW AND NOT AT UPDATE, WHERE i AM ALREADY DOING IT? XYP = ASCMAIN1.CYP ' Set_Period()

        Dim PO_SHIPMENT_LNO As Integer = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
        Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)

        With grdPOTSHIP2.ActiveRow
            If .Cells("PO_SHIP_STATUS").Value = "O" Then

                .Cells("PO_SHIP_STATUS").Value = "X"
                .Update()

                If select_from_3PL_list Or Select_from_Whse_Receipt Then
                    grdPOTSHIP3.DisplayLayout.Bands(0).Columns("PO_QTY_REC").CellActivation = UltraWinGrid.Activation.NoEdit
                    ' USE THE NEXT LINE TO ALLOW HUE TO CHANGE THE 944 RECIPT QTY
                    grdPOTSHIP3.DisplayLayout.Bands(0).Columns("PO_QTY_REC").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    grdPOTSHIP3.DisplayLayout.Bands(0).Columns("PO_QTY_REC").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If

                'For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO)
                '    rowPOTSHIP2.Item("OPS_YYYYPP") = XYP
                'Next

                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select(sqlw)
                    rowPOTSHIP3.Item("PO_QTY_REC") = rowPOTSHIP3.Item("PO_QTY_SHP")
                Next

            ElseIf .Cells("PO_SHIP_STATUS").Value = "X" Then
                .Cells("PO_SHIP_STATUS").Value = "O"
                .Update()
                grdPOTSHIP3.DisplayLayout.Bands(0).Columns("PO_QTY_REC").CellActivation = UltraWinGrid.Activation.NoEdit

                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select(sqlw)
                    rowPOTSHIP3.Item("PO_QTY_REC") = 0
                Next

            ElseIf .Cells("PO_SHIP_STATUS").Value = "C" Then
                .Cells("PO_SHIP_STATUS").Value = "R"
                .Update()
                grdPOTSHIP3.DisplayLayout.Bands(0).Columns("PO_QTY_REC").CellActivation = UltraWinGrid.Activation.NoEdit

            ElseIf .Cells("PO_SHIP_STATUS").Value = "R" Then
                .Cells("PO_SHIP_STATUS").Value = "C"
                .Update()
                grdPOTSHIP3.DisplayLayout.Bands(0).Columns("PO_QTY_REC").CellActivation = UltraWinGrid.Activation.NoEdit
            End If

        End With

        sqlw = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select(sqlw)
            Dim PO_QTY_REC As Integer = Val(rowPOTSHIP3.Item("PO_QTY_REC") & "")
            Dim SUB_UNIT_PACK_QTY As Integer = Val(rowPOTSHIP3.Item("SUB_UNIT_PACK_QTY") & "")
            rowPOTSHIP3.Item("PO_QTY_REC_DZ") = PO_QTY_REC / (12 / SUB_UNIT_PACK_QTY)
        Next

        Calculate_Landed_Cost()
        grdPOTSHIP3.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Sub Get_1st_Cost_from_PO()

        Fill_Records("POTORDR2_COSTS", PO_SHIPMENT_NO)

        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'")
            Dim PO_ORDER_NO As String = rowPOTSHIP3.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(rowPOTSHIP3.Item("PO_ORDER_LNO") & "")
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2_COSTS").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
            With rowPOTORDR2
                rowPOTSHIP3.Item("PO_COST_VCOST") = Val(.Item("PO_COST_VCOST") & "")
                rowPOTSHIP3.Item("PO_COST_MATLS") = Val(.Item("PO_COST_MATLS") & "")
                rowPOTSHIP3.Item("PO_COST_VCOST_UM") = Val(.Item("PO_COST_VCOST") & "")
                rowPOTSHIP3.Item("PO_COST_VCOST_DZ") = Val(.Item("PO_COST_VCOST_DZ") & "")
                rowPOTSHIP3.Item("PO_COST_MATLS_UM") = Val(.Item("PO_COST_MATLS") & "")
                rowPOTSHIP3.Item("PO_COST_MATLS_DZ") = Val(.Item("PO_COST_MATLS_DZ") & "")

                If .Item("DFQUOTA") & "" = "1" Then
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF_DZ") = Val(.Item("PO_COST_QUOTA") & "") ' REALLY? - YES POTORDR2 STORES /DZ IN THIS FIELD
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF") = Val(.Item("PO_COST_QUOTA") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DZ") = 0
                    rowPOTSHIP3.Item("PO_COST_QUOTA") = 0
                Else
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DZ") = Val(.Item("PO_COST_QUOTA") & "") ' REALLY? - YES POTORDR2 STORES /DZ IN THIS FIELD
                    rowPOTSHIP3.Item("PO_COST_QUOTA") = Val(.Item("PO_COST_QUOTA") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF_DZ") = 0
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF") = 0
                End If

                rowPOTSHIP3.Item("PO_COST") = Val(.Item("PO_COST") & "")

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    rowPOTSHIP3.Item("PO_COST_MATLS_UM") = Val(.Item("PO_COST_MATLS_DZ") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))
                    rowPOTSHIP3.Item("PO_COST_MATLS") = Val(.Item("PO_COST_MATLS_DZ") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))
                End If

                rowPOTSHIP3.Item("PO_COST_OTHER_DZ") = Val(.Item("PO_COST_OTHER") & "") ' REALLY? - YES POTORDR2 STORES /DZ IN THIS FIELD
                rowPOTSHIP3.Item("PO_COST_OTHER") = Val(.Item("PO_COST_OTHER") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))

                '    rowPOTSHIP3.Item("PO_COST_COMM") = Val(.Item("PO_COST_COMM") & "") ' THIS IS A %
                '    rowPOTSHIP3.Item("PO_COST_BUFFER") = Val(.Item("PO_COST_BUFFER") & "") ' THIS IS A %

                If .Item("DFQUOTA") & "" = "1" Then
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF_DZ") = Val(.Item("PO_COST_QUOTA") & "") ' REALLY? - YES POTORDR2 STORES /DZ IN THIS FIELD
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF") = Val(.Item("PO_COST_QUOTA") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DZ") = 0
                    rowPOTSHIP3.Item("PO_COST_QUOTA") = 0
                Else
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DZ") = Val(.Item("PO_COST_QUOTA") & "") ' REALLY? - YES POTORDR2 STORES /DZ IN THIS FIELD
                    rowPOTSHIP3.Item("PO_COST_QUOTA") = Val(.Item("PO_COST_QUOTA") & "") / (12 / Val(.Item("SUB_UNIT_PACK_QTY") & ""))
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF_DZ") = 0
                    rowPOTSHIP3.Item("PO_COST_QUOTA_DF") = 0
                End If
                '  rowPOTSHIP3.Item("PO_COST_QUOTA_DF_DZ") = Val(.Item("PO_COST_QUOTA_DF") & "") ' REALLY? - YES POTORDR2 STORES /DZ IN THIS FIELD
                '.Item("SHIP_COST_CHANGE_DATE") = DATETIME_STAMP
                '.Item("SHIP_COST_CHANGE_USER") = ASCMAIN1.USER_ID
            End With
        Next

        Set_Landed_Cost_Needs_to_be_Calculated_Indicator(False)
        Calculate_Landed_Cost()
    End Sub

    Sub Select_POs()

        'ASCMAIN1.sql = "Select SUBSTR(POTORDR1.PO_REFERENCE,1,10) REF" & vbCrLf _
        '    & ", POTORDR2.STYLE_CODE STYLE, POTORDR2.COLOR_CODE COLOR " & vbCrLf _
        '    & ", TRUNC(POTORDR2.PO_QTY_OPN / (12 / nvl(POTORDR2.SUB_UNIT_PACK_QTY,1)) * 100) / 100 DZ " & vbCrLf _
        '    & ", SUBSTR(POTORDR1.PO_SPEC_ORDR_NO,1,10) SPECIAL" & vbCrLf _
        '    & ", POTORDR2.PO_ORDER_NO PO_NO, POTORDR2.PO_ORDER_LNO LNO" & vbCrLf _
        '    & ", POTORDR1.VEND_CODE VENDOR, POTORDR1.FOB_CMT TYPE " & vbCrLf _
        '    & ", POTORDR2.PO_COST_VCOST V_COST, POTORDR2.PO_COST_MATLS M_COST" & vbCrLf _
        '    & ", POTORDR2.PO_DATE_ETA ETA, POTORDR2.PO_DATE_SHIP_BY SHIP_BY, POTORDR2.PO_QTY_OPN, POTORDR2.PO_COST_QUOTA" & vbCrLf _
        '    & " from POTORDR1,POTORDR2 " & vbCrLf _
        '    & " where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
        '    & " and POTORDR2.PO_STATUS = 'O' " & vbCrLf _
        '    & " and (POTORDR2.PO_QTY_OPN > 0 or POTORDR1.PO_ORDER_NO in (Select DISTINCT PO_ORDER_NO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'))"

        Dim sqlPOTORDR1 As String = "Select PO_ORDER_NO, VEND_CODE, PO_DATE_ORDERED, PO_REFERENCE" _
            & ", PO_DATE_ETA, PO_SPEC_ORDR_NO, FACTORY_CODE" _
            & " from POTORDR1"

        ASCMAIN1.sql = "" _
            & sqlPOTORDR1 & " where PO_STATUS = 'O'" _
            & IIf(ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA", " and WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")", "") _
            & " union " _
            & sqlPOTORDR1 & " where PO_ORDER_NO in (Select Distinct PO_ORDER_NO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')"

        Dim tbl As DataTable = ASCDATA1.GetDataTable

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PO_ORDER_NO")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = (optPOLines.Value <> "S")
            ASCMAIN1.CodeSelector.UseDataFromTable = tbl
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                Dim POs_checked As New List(Of String)
                Dim POs_locked As New List(Of String)
                Dim POs_not_locked As String = ""

                ASCMAIN1.Progress("-", "1")

                For Each PO_ORDER_NO As String In ASCMAIN1.CodeSelector.SelectedCodes
                    ASCMAIN1.Progress("-", "1 - " & PO_ORDER_NO)
                    If ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, False, False, False, 1) Then
                        POs_locked.Add(PO_ORDER_NO)
                        Load_POs_into_POTSHIP3(PO_ORDER_NO)
                    Else
                        POs_not_locked &= "," & PO_ORDER_NO
                    End If
                    POs_checked.Add(PO_ORDER_NO)
                Next

                ASCMAIN1.Progress("-", "2")
                If POs_not_locked <> "" Then
                    MsgBox(Mid(POs_not_locked, 2), vbOKOnly, "The following PO(s) are currently being edited and could not be locked.")
                End If

                Sort_grdColumns(grdPOTSHIP3, "PO_ORDER_NO,PO_ORDER_LNO")
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If

    End Sub

    Sub Load_POs_into_POTSHIP3(PO_ORDER_NO As String,
                               Optional fromXLS_Import As Boolean = False,
                               Optional shipmentLno As Integer = 0,
                               Optional COLOR_CODEs As List(Of String) = Nothing)


        Try

            Dim PO_SHIPMENT_LNO As Integer = IIf(fromXLS_Import Or shipmentLno > 0, shipmentLno, 0)

            If Not fromXLS_Import AndAlso grdPOTSHIP2.ActiveRow IsNot Nothing Then
                PO_SHIPMENT_LNO = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            End If

            ASCMAIN1.Progress("-", "1A - " & PO_ORDER_NO)

            ASCMAIN1.sql = "Select POTORDR2.*,ICTSTYL1.STYLE_DESC,ICTCOLR1.COLOR_DESC" _
                & " from POTORDR2,ICTSTYL1,ICTCOLR1" _
                & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'" _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" _
                & "   and ICTCOLR1.COLOR_CODE = POTORDR2.COLOR_CODE"
            If dst.Tables("POTORDRO").Select("PO_ORDER_NO = '" & PO_ORDER_NO & "'").Length = 0 Then
                ASCMAIN1.sql &= " and NVL(POTORDR2.PO_QTY_OPN,0) > 0"
            End If

            If COLOR_CODEs IsNot Nothing Then
                ASCMAIN1.sql &= " and NVL(POTORDR2.PO_QTY_OPN,0) > 0"
                ASCMAIN1.sql &= " and POTORDR2.COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
            End If

            ASCMAIN1.Progress("-", "2A - " & PO_ORDER_NO)

            Dim TBL As DataTable = IIf(fromXLS_Import, dst.Tables("POTORDR2_SPLIT"), ASCDATA1.GetDataTable)
            Dim PO_ORDER_LNOs As New List(Of Integer)

            If Not fromXLS_Import And shipmentLno <> 0 Then
                For Each rowPO_LNO As DataRow In TBL.Select("PO_ORDER_NO = '" & PO_ORDER_NO & "'")
                    Dim PO_ORDER_LNO As String = rowPO_LNO.Item("PO_ORDER_LNO").ToString
                    PO_ORDER_LNOs.Add(PO_ORDER_LNO)
                Next
            Else

                If optPOLines.Value = "S" Then
                    If Not fromXLS_Import Then
                        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PO_ORDER_LNO")

                        ASCMAIN1.Progress("-", "3A - " & PO_ORDER_NO)

                        If ASCMAIN1.CodeSelector.SQL = "" Then
                            Exit Sub
                        Else
                            ASCMAIN1.CodeSelector.MultipleSelections = True
                            ASCMAIN1.CodeSelector.UseDataFromTable = TBL
                            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                            ASCMAIN1.Progress("-", "4A - " & PO_ORDER_NO)
                            Try
                                Dim F As New ASFCODE1
                                F.ShowDialog()
                                F.Dispose()

                                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                                    For Each PO_ORDER_LNO As String In ASCMAIN1.CodeSelector.SelectedCodes
                                        PO_ORDER_LNOs.Add(PO_ORDER_LNO)
                                    Next
                                End If
                            Catch ex As Exception
                                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Please Call Walter 201 400 2156")
                            End Try
                        End If
                    Else
                        For Each rowPO_LNO As DataRow In TBL.Select("PO_ORDER_NO = '" & PO_ORDER_NO & "'")
                            Dim PO_ORDER_LNO As String = rowPO_LNO.Item("PO_ORDER_LNO").ToString
                            PO_ORDER_LNOs.Add(PO_ORDER_LNO)
                        Next

                    End If
                End If
            End If

            If optPOLines.Value = "S" And PO_ORDER_LNOs.Count = 0 Then Exit Sub

            'Dim rowPOTORDR1 As DataRow = IIf(fromImport, dst.Tables("POTORDR1_SPLIT").Select("PO_ORDER_NO ='" & PO_ORDER_NO & "'")(0), LookUp("POTORDR1", PO_ORDER_NO))
            Dim rowPOTORDR1 As DataRow = Nothing
            Dim sqlW As String = IIf(fromXLS_Import, "PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO, "")
            If fromXLS_Import Then
                rowPOTORDR1 = dst.Tables("POTORDR1_SPLIT").Select("PO_ORDER_NO ='" & PO_ORDER_NO & "'")(0)
            Else
                rowPOTORDR1 = LookUp("POTORDR1", PO_ORDER_NO)
            End If


            For Each rowPOTORDR2 As DataRow In TBL.Select(sqlW)

                Dim PO As String = rowPOTORDR2.Item("PO_ORDER_NO")
                If PO <> PO_ORDER_NO Then
                    PO_ORDER_NO = PO
                    rowPOTORDR1 = dst.Tables("POTORDR1_SPLIT").Select("PO_ORDER_NO ='" & PO_ORDER_NO & "'")(0)
                End If

                Dim PO_ORDER_LNO As Integer = Val(rowPOTORDR2.Item("PO_ORDER_LNO") & "")
                If optPOLines.Value = "A" Or PO_ORDER_LNOs.Contains(PO_ORDER_LNO) Then

                    '  Dim rowPOTORDR2 As DataRow = LookUp("POTORDR2", New String() {PO_ORDER_NO, PO_ORDER_LNO})
                    Dim SUB_UNIT_PACK_QTY As Integer = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")

                    Dim SPQ As Integer = IIf(SUB_UNIT_PACK_QTY = 0, 12, 12 / SUB_UNIT_PACK_QTY)

                    'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    '    ' TOO SCARED TO CHANGE THIS WITHOUT CHECKING TO SEE IF IT REALLY IS A BUG
                    '    ' IF IT IS A BUG THEN WHENEVER NADINE CREATES A SHIPMENT, THE COSTS ARE AFU, AND THEN VISHA MUST ALWAYS HAVE TO REFRESH 1ST COST, WHICH APPEARS TO DO THIS RIGHT
                    'Else
                    '    SPQ = IIf(SUB_UNIT_PACK_QTY = 0, 12, 12 / SUB_UNIT_PACK_QTY)
                    'End If

                    Dim PO_QTY_OPN As Int64 = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")
                    Dim PO_QTY_SHP As Int64 = IIf(fromXLS_Import, Val(rowPOTORDR2.Item("PO_QTY_SHP") & ""), PO_QTY_OPN)

                    If PO_QTY_OPN = 0 And PO_QTY_SHP = 0 And fromXLS_Import Then
                        Continue For
                    End If

                    Dim rowPOTSHIP3 As DataRow = dst.Tables("POTSHIP3").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, PO_ORDER_NO, PO_ORDER_LNO})
                    If rowPOTSHIP3 Is Nothing Then
                        rowPOTSHIP3 = dst.Tables("POTSHIP3").NewRow
                        With rowPOTSHIP3

                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            .Item("STYLE_CODE") = rowPOTORDR2.Item("STYLE_CODE")
                            .Item("COLOR_CODE") = rowPOTORDR2.Item("COLOR_CODE")

                            .Item("PO_QTY_SHP") = PO_QTY_SHP
                            .Item("PO_QTY_OPN") = PO_QTY_OPN
                            .Item("PO_QTY_REC") = 0

                            .Item("PO_QTY_UOM") = rowPOTORDR2.Item("PO_QTY_UOM")
                            .Item("PO_COST") = Val(rowPOTORDR2.Item("PO_COST") & "")
                            .Item("PO_COST_VCOST") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "")
                            .Item("PO_COST_MATLS") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "")

                            .Item("PO_COST_VCOST_UM") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "")
                            .Item("PO_COST_MATLS_UM") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "")

                            ' IMPORTANT - note that this field is currently maintained in POTORDR2 per Dozen units, and is per unit in POTSHIP3
                            .Item("PO_COST_OTHER") = Val(rowPOTORDR2.Item("PO_COST_OTHER") & "") / SPQ

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then ' VAN PARANOIA
                                .Item("PO_COST_COMM") = Val(rowPOTORDR2.Item("PO_COST_COMM") & "")
                            Else
                                If rowPOTORDR1.Item("PO_COMM_PAYABLE_TO_BRKR") & "" = "1" Then
                                    .Item("PO_COST_COMM") = Val(rowPOTORDR1.Item("PO_COMM_PCT") & "")
                                Else
                                    .Item("PO_COST_COMM") = 0
                                End If
                            End If

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                .Item("PO_COST_BUFFER") = 5
                            End If

                            ' this is not exactly true- but we can let the calculation routines fix it later
                            .Item("PO_COST_LANDED") = Val(rowPOTORDR2.Item("PO_COST") & "")

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                If Val(rowPOTORDR2.Item("DFQUOTA") & "") = 1 Then
                                    .Item("PO_COST_QUOTA_DF") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "") / SPQ
                                    .Item("PO_COST_QUOTA_DF_DZ") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "")
                                Else
                                    .Item("PO_COST_QUOTA") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "") / SPQ
                                    .Item("PO_COST_QUOTA_DZ") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "")
                                End If
                            End If

                            If Val(rowPOTORDR2.Item("PO_COST_VCOST_DZ") & "") = 0 Then
                                .Item("PO_COST_VCOST_DZ") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "") * SPQ
                            Else
                                .Item("PO_COST_VCOST_DZ") = Val(rowPOTORDR2.Item("PO_COST_VCOST_DZ") & "")
                            End If

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                If Val(rowPOTORDR2.Item("PO_COST_MATLS_DZ") & "") = 0 Then
                                    .Item("PO_COST_MATLS_DZ") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "") * SPQ
                                Else
                                    .Item("PO_COST_MATLS_DZ") = Val(rowPOTORDR2.Item("PO_COST_MATLS_DZ") & "")
                                End If
                            End If

                            .Item("PO_COST_OTHER_DZ") = Val(rowPOTORDR2.Item("PO_COST_OTHER") & "") ' see note above regarding unit of measure for PO_COST_OTHER
                            .Item("SUB_UNIT_PACK_QTY") = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")
                            .Item("CARTON_PACK_QTY") = Val(rowPOTORDR2.Item("CARTON_PACK_QTY") & "")
                            If Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "") = 0 Then
                                .Item("PO_QTY_SHP_DZ") = 0
                                .Item("NET_OPEN_DZ") = 0
                            Else
                                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                    .Item("PO_QTY_SHP_DZ") = PO_QTY_SHP / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                                Else
                                    .Item("PO_QTY_SHP_DZ") = PO_QTY_OPN / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                                End If
                                .Item("NET_OPEN_DZ") = PO_QTY_OPN / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                            End If
                            .Item("PO_QTY_REC_DZ") = 0

                            .Item("PO_REFERENCE") = rowPOTORDR1.Item("PO_REFERENCE")
                            .Item("PO_DATE_SHIP_BY") = rowPOTORDR2.Item("PO_DATE_SHIP_BY")
                            .Item("FOB_CMT") = (rowPOTORDR1.Item("FOB_CMT") & "")
                            .Item("VEND_CODE") = rowPOTORDR1.Item("VEND_CODE")

                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowPOTORDR2.Item("STYLE_CODE"))
                            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                        End With

                        If dst.Tables("POTORDRO").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO}) Is Nothing Then
                            ASCMAIN1.sql = "Select Sum (PO_QTY_SHP) from POTSHIP3 where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
                            Dim PO_QTY_SHP_PRE As Int64 = Val(ASCDATA1.GetDataValue & "")
                            Dim rowPOTORDRO As DataRow = dst.Tables("POTORDRO").NewRow
                            rowPOTORDRO.Item("PO_ORDER_NO") = PO_ORDER_NO
                            rowPOTORDRO.Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            rowPOTORDRO.Item("PO_QTY_OPN") = PO_QTY_OPN
                            If EntryMode = "N" Or EntryMode = "E" Then
                                rowPOTORDRO.Item("PO_QTY_OPN_PRE") = PO_QTY_OPN + PO_QTY_SHP_PRE
                            Else
                                rowPOTORDRO.Item("PO_QTY_OPN_PRE") = PO_QTY_OPN
                            End If
                            dst.Tables("POTORDRO").Rows.Add(rowPOTORDRO)
                        End If
                        Create_POTSHIPR(PO_SHIPMENT_LNO, rowPOTORDR2.Item("STYLE_CODE"), rowPOTORDR2.Item("COLOR_CODE"))

                        dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)
                    End If
                End If
            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Please take a screenshot and send to ABS - do NOT Update")
        End Try

    End Sub

    Sub Calc_Cost_Variance()

        Dim varianceReportNeeded As Boolean = False

        If MsgBox("Do you really want to calculate the PO Cost Variance?",
                  MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Verification.") = MsgBoxResult.Cancel Then
            Exit Sub
        End If

        dst.Tables("POTCOSTV").Rows.Clear()

        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'")
            Dim variance As Boolean = False
            Dim PO_ORDER_NO As String = rowPOTSHIP3.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(rowPOTSHIP3.Item("PO_ORDER_LNO"))
            Dim rowPOTCOSTF As DataRow = Fill_Record("POTCOSTF", New Object() {PO_ORDER_NO, PO_ORDER_LNO})

            If rowPOTCOSTF IsNot Nothing Then
                With rowPOTCOSTF
                    If Val(.Item("PO_COST_VCOST") & "") <> Val(rowPOTSHIP3.Item("PO_COST_VCOST") & "") _
                    Or Val(.Item("PO_COST_MATLS") & "") <> Val(rowPOTSHIP3.Item("PO_COST_MATLS") & "") _
                    Or Val(.Item("PO_COST_VCOST_DZ") & "") <> Val(rowPOTSHIP3.Item("PO_COST_VCOST_DZ") & "") _
                    Or Val(.Item("PO_COST_MATLS_DZ") & "") <> Val(rowPOTSHIP3.Item("PO_COST_MATLS_DZ") & "") _
                    Or Val(.Item("PO_COST_OTHER") & "") <> Val(rowPOTSHIP3.Item("PO_COST_OTHER_DZ") & "") _
                    Or Val(.Item("PO_COST_COMM") & "") <> Val(rowPOTSHIP3.Item("PO_COST_COMM") & "") _
                    Or Val(.Item("PO_COST_QUOTA") & "") <> Val(rowPOTSHIP3.Item("PO_COST_QUOTA_DZ") & "") Then
                        variance = True
                    End If
                End With
                'ElseIf Val(.item("PO_COST_QUOTA_DF") & "") <> Val(rowPOTSHIP3.item("PO_COST_QUOTA_DF_DZ") & "") Then
                '   variance = True

                If variance = True Then
                    varianceReportNeeded = True
                    Dim rowPOTCOSTV As DataRow = dst.Tables("POTCOSTV").NewRow
                    With rowPOTCOSTV
                        For Each COLUMN_NAME As String In New String() _
                            {"PO_SHIPMENT_NO", "PO_SHIPMENT_LNO", "PO_ORDER_NO", "PO_ORDER_LNO"}
                            .Item(COLUMN_NAME) = rowPOTSHIP3.Item(COLUMN_NAME)
                        Next
                        .Item("PO_REFERENCE") = rowPOTCOSTF.Item("PO_REFERENCE")
                        For Each COLUMN_NAME As String In New String() _
                            {"PO_COST_VCOST", "PO_COST_VCOST_DZ", "PO_COST_MATLS", "PO_COST_MATLS_DZ", "PO_COST_OTHER_DZ", "PO_COST_COMM", "PO_COST_QUOTA_DZ"}
                            .Item(Replace(COLUMN_NAME, "PO_COST_", "SHIP3_")) = rowPOTSHIP3.Item(COLUMN_NAME)
                            .Item(Replace(COLUMN_NAME, "PO_COST_", "ORDR2_")) = rowPOTCOSTF.Item(COLUMN_NAME)
                        Next
                        For Each COLUMN_NAME As String In New String() _
                             {"PO_COST_VCOST", "PO_COST_VCOST_DZ", "PO_COST_MATLS", "PO_COST_MATLS_DZ", "PO_COST_OTHER_DZ", "PO_COST_COMM", "PO_COST_QUOTA_DZ"}
                            Stop ' WHAT WE SUPPOSED TO DO IN HERE?
                        Next
                    End With
                    dst.Tables("POTCOSTV").Rows.Add(rowPOTCOSTV)
                Else
                    'costs match
                End If
            End If
        Next

        If varianceReportNeeded Then
            Print_Report_Begin()
            CR_params.Add("SUBT", "")
            Generate_Report("PORCOSTV")
            Print_Report_End()

            Using F As New ASFMSGBF
                F.Show_grd(dst.Tables("POTCOSTV"), Me, "Cost Variances")
            End Using

        Else
            MsgBox("There is no variance between the PO and Shipment cost details", MsgBoxStyle.OkOnly, "Costs Match!")
        End If
    End Sub

    Function Set_Period() As String
        Dim PO_DATE_RECEIVED As Date = Absx1.dteFor("PO_DATE_RECEIVED").Value
        If Format(PO_DATE_RECEIVED, "yyyyMM") > ASCMAIN1.CYM Then
            Return ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)
        Else
            Return ASCMAIN1.CYP
        End If
    End Function

    Private Sub Get_Duty()

        STYLE_CODEs_No_Duty.Clear()

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
        Dim WHSE_COUNTRY As String = rowICTWHSE1.Item("WHSE_COUNTRY") & ""
        Dim COUNTRY_CODE As String = ""
        If WHSE_COUNTRY <> "" And WHSE_COUNTRY <> "USA" Then
            COUNTRY_CODE = WHSE_COUNTRY
        End If

        For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
            Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
            Dim OPS_YYYYPP As String = rowPOTSHIP2.Item("OPS_YYYYPP") & ""
            If OPS_YYYYPP = "" Then
                OPS_YYYYPP = ASCMAIN1.CYP ' XYP
            End If

            If ASCMAIN1.CLIENT = "NYA - NOT" Then
                Fill_Records("ICTSTYLD", New String() {COUNTRY_CODE, PO_SHIPMENT_NO, COUNTRY_CODE, Mid(OPS_YYYYPP, 1, 4)})
            Else
                Fill_Records("ICTSTYLD", New String() {PO_SHIPMENT_NO, Mid(OPS_YYYYPP, 1, 4)})
            End If

            ' ALTER TABLE ICTPORT1 ADD COUNTRY_CODE VARCHAR2(3)

            ' Apply Duty Rate Modifier is importing from a country with Duty Rate modifications

            Dim COST_NO_DUTY As String = rowPOTSHIP1.Item("COST_NO_DUTY") & ""
            If COST_NO_DUTY <> "1" Then

                Dim PORT_CODE_ORIG As String = rowPOTSHIP1.Item("PORT_CODE_ORIG") & ""
                Dim rowICTPORT1 As DataRow = LookUp("ICTPORT1", PORT_CODE_ORIG)

                Dim COUNTRY_CODE_import_from As String = ""
                If rowICTPORT1 IsNot Nothing Then
                    COUNTRY_CODE_import_from = rowICTPORT1.Item("COUNTRY_CODE") & ""
                End If

                Dim shipment_received As Boolean = False
                If rowPOTSHIP2.Item("PO_DATE_RECEIVED") & "" <> "" Then
                    shipment_received = True
                End If
                If COUNTRY_CODE_import_from <> "" And COUNTRY_CODE_import_from <> "USA" Then

                    dst.Tables("ICTDUTY4").Rows.Clear()

                    Dim PO_DATE_RECEIVED As Date
                    If shipment_received Then
                        If (ASCMAIN1.CLIENT = "NYA" AndAlso rowPOTSHIP2.Item("PO_DATE_RECEIVED_PORT") & "" <> "") Then
                            PO_DATE_RECEIVED = rowPOTSHIP2.Item("PO_DATE_RECEIVED_PORT")
                        Else
                            PO_DATE_RECEIVED = rowPOTSHIP2.Item("PO_DATE_RECEIVED")
                        End If

                    Else
                        PO_DATE_RECEIVED = rowPOTSHIP1.Item("PO_SHIP_ETA")
                    End If
                    Dim DATE_RECEIVED_IN_PORT_sql As String = Format(PO_DATE_RECEIVED, "dd-MMM-yyyy")
                    For Each rowICTSTYLD As DataRow In dst.Tables("ICTSTYLD").Select("")
                        Dim DUTY_RATE_CODE As String = rowICTSTYLD.Item("DUTY_RATE_CODE")
                        Dim DUTY_RATE_ICTSTYLD As Decimal = Val(rowICTSTYLD.Item("DUTY_RATE") & "")

                        Dim rowICTDUTY4 As DataRow = dst.Tables("ICTDUTY4").Rows.Find(DUTY_RATE_CODE)
                        If rowICTDUTY4 Is Nothing Then
                            rowICTDUTY4 = dst.Tables("ICTDUTY4").NewRow
                            rowICTDUTY4.Item("DUTY_RATE_CODE") = DUTY_RATE_CODE
                            rowICTDUTY4.Item("COUNTRY_CODE") = COUNTRY_CODE_import_from

                            ASCMAIN1.sql = "Select * from ICTDUTY4" & vbCrLf _
                                & " where DUTY_RATE_CODE = '" & DUTY_RATE_CODE & "'" & vbCrLf _
                                & "   and COUNTRY_CODE = '" & COUNTRY_CODE_import_from & "'" & vbCrLf _
                                & "   and DUTY_RATE_BEGIN <= '" & DATE_RECEIVED_IN_PORT_sql & "'" & vbCrLf _
                                & "   and (DUTY_RATE_END is Null or DUTY_RATE_END >= '" & DATE_RECEIVED_IN_PORT_sql & "')"
                            Dim row As DataRow = ASCDATA1.GetDataRow

                            If row Is Nothing Then
                                rowICTDUTY4.Item("DUTY_RATE") = DUTY_RATE_ICTSTYLD
                                rowICTDUTY4.Item("DUTY_RATE_ADD") = "0"
                            Else
                                rowICTDUTY4.Item("DUTY_RATE") = row.Item("DUTY_RATE")
                                rowICTDUTY4.Item("DUTY_RATE_ADD") = row.Item("DUTY_RATE_ADD")
                            End If

                            dst.Tables("ICTDUTY4").Rows.Add(rowICTDUTY4)
                        End If

                        Dim DUTY_RATE As Decimal = Val(rowICTDUTY4.Item("DUTY_RATE") & "")
                        Dim DUTY_RATE_ADD As String = rowICTDUTY4.Item("DUTY_RATE_ADD") & ""
                        If DUTY_RATE_ADD = "1" Then
                            DUTY_RATE += DUTY_RATE_ICTSTYLD
                        End If
                        rowICTSTYLD.Item("DUTY_RATE") = DUTY_RATE
                    Next

                End If
            End If



            For Each rowPOTSHIP3 As DataRow In rowPOTSHIP2.GetChildRows("POTSHIP2_POTSHIP3")
                Dim rowICTSTYLD As DataRow = dst.Tables("ICTSTYLD").Rows.Find(rowPOTSHIP3.Item("STYLE_CODE"))
                If rowICTSTYLD Is Nothing Then
                    STYLE_CODEs_No_Duty.Add(rowPOTSHIP3.Item("STYLE_CODE"))
                Else
                    rowPOTSHIP3.Item("DUTY_RATE") = Val(rowICTSTYLD.Item("DUTY_RATE") & "")
                    rowPOTSHIP3.Item("DUTY_RATE_CODE") = rowICTSTYLD.Item("DUTY_RATE_CODE") & ""
                End If
            Next
        Next

        If cost_calc Then
            If STYLE_CODEs_No_Duty.Count <> 0 Then
                Dim no_duty_required As Boolean = False
                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    Dim FRT_TERMS As String = ""
                    If grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & "" <> "" Then
                        Dim ORDR_NO As String = grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & ""
                        Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO)
                        FRT_TERMS = rowSOTORDR1.Item("FRT_TERMS")
                    End If

                    If rowPOTSHIP1.Item("WHSE_CODE") & "" = "FE" And FRT_TERMS = "COL" Then
                        no_duty_required = True
                    End If

                End If
                If no_duty_required Then
                Else
                    If automated_cost_complete Then
                    Else
                        MsgBox("Warning - No Duty Rate for the following Styles:" & vbCrLf & Join(STYLE_CODEs_No_Duty.ToArray, ","), MsgBoxStyle.OkOnly, "Update will not be permitted")
                    End If
                End If
            End If
        End If

        Set_Landed_Cost_Needs_to_be_Calculated_Indicator(False)
        Calculate_Landed_Cost()
    End Sub

    Private Sub Get_Weight_Factor()

        Fill_Records("ICTSTYLW", PO_SHIPMENT_NO)
        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'")
            Dim STYLE_CODE As String = rowPOTSHIP3.Item("STYLE_CODE")
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                Dim rowICTSTYLW As DataRow = dst.Tables("ICTSTYLW").Rows.Find(STYLE_CODE)
                rowPOTSHIP3.Item("WEIGHT_FACTOR") = rowICTSTYLW.Item("WEIGHT_FACTOR")
            Else
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                Dim WEIGHT_FACTOR As Decimal = Val(rowICTSTYL1.Item("STYLE_WEIGHT") & "")
                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    Dim CASE_WEIGHT_GRS As Decimal = Val(rowICTSTYL1.Item("CASE_WEIGHT_GRS") & "")
                    Dim CARTON_PACK_QTY As Decimal = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
                    If CARTON_PACK_QTY = 0 Then
                        WEIGHT_FACTOR = 0
                    Else
                        WEIGHT_FACTOR = CASE_WEIGHT_GRS / CARTON_PACK_QTY
                    End If

                End If
                rowPOTSHIP3.Item("WEIGHT_FACTOR") = WEIGHT_FACTOR
            End If
        Next

        Set_Landed_Cost_Needs_to_be_Calculated_Indicator(False)
        Calculate_Landed_Cost()
    End Sub

    Sub Set_Container()
        If grdPOTSHIP2.ActiveRow Is Nothing Then Exit Sub
        If grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "" = "" Then
            grdPOTSHIP3.Visible = False
        Else
            Sort_grdColumns(grdPOTSHIP3, "PO_REFERENCE,PO_ORDER_NO,PO_ORDER_LNO,PO_SHIPMENT_LNO")

            grdPOTSHIP3.Visible = True
        End If
    End Sub

    Sub Set_Receipt_Options()

        Dim col As UltraWinGrid.UltraGridColumn = grdPOTSHIP3.DisplayLayout.Bands(0).Columns("PO_QTY_REC")

        With grdPOTSHIP2.ActiveRow
            If .Cells("PO_SHIP_STATUS").Value = "O" Then
                col.CellActivation = UltraWinGrid.Activation.NoEdit
            ElseIf .Cells("PO_SHIP_STATUS").Value = "X" Then
                If select_from_3PL_list Or Select_from_Whse_Receipt Then
                    col.CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    col.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            ElseIf .Cells("PO_SHIP_STATUS").Value = "C" Then
                col.CellActivation = UltraWinGrid.Activation.NoEdit
            ElseIf .Cells("PO_SHIP_STATUS").Value = "R" Then
                col.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
        grdPOTSHIP3.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Public Overrides Function CustomSummary_End(
      ByVal summarySettings As UltraWinGrid.SummarySettings,
      ByVal rows As UltraWinGrid.RowsCollection,
      ByVal CustomValue As Double,
      ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        'Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name

            Case "grdPOTSHIP3"
                Dim KEY As String = summarySettings.Key

                Dim COLUMN_NAME_QTY As String = "PO_QTY_SR"
                Dim F As Decimal = 1
                If KEY.EndsWith("_DZ") Then
                    F = 12
                End If

                For Each grow As UltraWinGrid.UltraGridRow In rows
                    Dim SUB_UNIT_PACK_QTY As Integer = Val(grow.Cells("SUB_UNIT_PACK_QTY").Value & "")
                    If KEY.EndsWith("_DZ") Then
                        F = 12 / SUB_UNIT_PACK_QTY
                    End If
                    CustomValue += Val(grow.Cells(COLUMN_NAME_QTY).Value & "") * Val(grow.Cells(KEY).Value & "") / F
                    'Dim AMT As Decimal = Val(grow.Cells(COLUMN_NAME_QTY).Value & "") * Val(grow.Cells(KEY).Value & "") / F
                    'If System.Math.Abs(AMT - Val(grow.Cells("EXT_VCOST").Value & "")) > 0.01 Then
                    '    Stop
                    'End If

                Next
                '  Stop
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue

    End Function

#Region "grdPOTSHIP2"

    Private Sub grdPOTSHIP2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "NON_AR"
            '    If e.Cell.Value & "" = "1" Then
            '        grdPOTSHIP2.ActiveRow.Cells("CUST_CODE").Value = ""
            '        grdPOTSHIP2.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
            '    Else
            '        grdPOTSHIP2.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
            '    End If
        End Select
    End Sub

    Private Sub grdPOTSHIP2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTSHIP2.AfterExitEditMode
    End Sub

    Private Sub grdPOTSHIP2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTSHIP2.AfterRowActivate
        If grdPOTSHIP2.ActiveRow.IsGroupByRow Then
            tabBOL.Visible = False
        Else
            Setup_grdPOTSHIP2_ActiveRow()
            tabBOL.Visible = True
        End If

    End Sub

    Sub Setup_grdPOTSHIP2_ActiveRow()

        If grdPOTSHIP2.ActiveRow IsNot Nothing Then

            If grdPOTSHIP2.ActiveRow.IsGroupByRow Then
                Exit Sub
            End If

            With grdPOTSHIP2.DisplayLayout.Bands(0).Columns("ACTION")
                If grdPOTSHIP2.ActiveRow.IsAddRow Then
                    .Style = UltraWinGrid.ColumnStyle.Default
                    '.ButtonDisplayStyle = UltraWinGrid.ColumnStyle.Default
                    .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                Else
                    .Style = UltraWinGrid.ColumnStyle.Button
                    .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                End If
            End With
        End If

        If grdPOTSHIP2.ActiveRow Is Nothing OrElse Not grdPOTSHIP2.ActiveRow.IsDataRow OrElse grdPOTSHIP2.ActiveRow.IsAddRow Then
            grdPOTSHIP3.Visible = False
            splMain.Panel2Collapsed = True
            If grdPOTSHIP2.Rows.Count = 0 Then Exit Sub
            tabBOL.SelectedTab = tabBOL.Tabs("PO Details")
        Else
            splMain.Panel2Collapsed = False
            grdPOTSHIP3.Visible = True
            Dim COMM_INV_NO As String = grdPOTSHIP2.ActiveRow.Cells("COMM_INV_NO").Value & ""
            Dim BOL_NO As String = grdPOTSHIP2.ActiveRow.Cells("BOL_NO").Value & ""
            Dim CONTAINER_NO As String = ""
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                CONTAINER_NO = grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO").Value & ""
            End If
            grdPOTSHIP3.Text = "" _
                & IIf(COMM_INV_NO = "", "", "Commercial Invoice '" & COMM_INV_NO & "', ") _
                & grdPOTSHIP2.DisplayLayout.Bands(0).Columns("BOL_NO").Header.Caption & " '" & BOL_NO & "'" _
                & IIf(CONTAINER_NO = "", "", ", " & grdPOTSHIP2.DisplayLayout.Bands(0).Columns("CONTAINER_NO").Header.Caption & " '" & CONTAINER_NO & "'") & " Contents"

            Dim PO_SHIPMENT_LNO As Int32 = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            Dim dvw As DataView = DirectCast(grdPOTSHIP3.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
            Sort_grdColumns(grdPOTSHIP3, "PO_REFERENCE,PO_ORDER_NO,PO_ORDER_LNO")
            If ASCMAIN1.CLIENT = "RGI" Then
                Dim dvwr As DataView = DirectCast(grdWHTPREC3.DataSource, DataTable).DefaultView
                CONTAINER_NO = grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO").Value & ""
                LoadWhseReceipts()
                dvwr.RowFilter = "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
                Sort_grdColumns(grdWHTPREC3, "PO_REFERENCE,PO_ORDER_NO,PO_ORDER_LNO")
                Dim WH_REC_STATUS As String = ASCDATA1.GetDataValue("SELECT WH_REC_STATUS from WHTWREC1 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and CONTAINER_NO = '" & CONTAINER_NO & "'")
                'btnRcvClose.Enabled = (grdPOTSHIP2.ActiveRow.Cells("PO_SHIP_STATUS").Value = "O") And WH_REC_STATUS = "O"
                btnRcvClose.Text = If((grdPOTSHIP2.ActiveRow.Cells("PO_SHIP_STATUS").Value = "O") And WH_REC_STATUS = "O", "Close Container", "Reprint Report")
            End If
        End If

        UltraExplorerBar1.Groups("Back-to-Back").Visible = False
        If grdPOTSHIP2.ActiveRow IsNot Nothing Then
            If grdPOTSHIP2.ActiveRow.IsAddRow Then
            Else
                If grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & "" <> "" Then
                    Dim ORDR_NO As String = grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & ""
                    Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO)
                    UltraExplorerBar1.Groups("Back-to-Back").Visible = True And ScreenMode
                    UltraExplorerBar1.Groups("Back-to-Back").Text = "Back-to-Back" & " " & rowSOTORDR1.Item("FRT_TERMS")
                End If
            End If
        End If

        'If ship_entry Then
        '    grdPOTSHIP2.DisplayLayout.Bands(0).Columns("ACTION").Hidden = (grdPOTSHIP2.ActiveRow Is Nothing OrElse Not grdPOTSHIP2.ActiveRow.IsDataRow OrElse grdPOTSHIP2.ActiveRow.IsAddRow)
        'End If

        If tabBOL.SelectedTab.Key = "Cartons" Then
            If grdPOTSHIP2.ActiveRow Is Nothing OrElse Not grdPOTSHIP2.ActiveRow.IsDataRow Then
                'grdPOTSHIPR.Visible = False
                splCartonQ.Visible = False
            Else
                'grdPOTSHIPR.Visible = True
                splCartonQ.Visible = True
                grdPOTSHIPR.Text = "Style/Color Unit Recap for BOL" & " " & grdPOTSHIP2.ActiveRow.Cells("BOL_NO").Value & ""

                Dim PO_SHIPMENT_LNO As Int32 = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                Dim dvw As DataView = DirectCast(grdPOTSHIPR.DataSource, DataTable).DefaultView
                dvw.RowFilter = "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
            End If

            If grdPOTSHIP2.ActiveRow Is Nothing OrElse Not grdPOTSHIP2.ActiveRow.IsDataRow Then
                grdPOTSHIP7.Visible = False
                grdPOTSHIP8.Visible = False
            Else
                grdPOTSHIP7.Visible = True
                grdPOTSHIP8.Visible = True
                grdPOTSHIP7.Text = "Carton Types for BOL" & " " & grdPOTSHIP2.ActiveRow.Cells("BOL_NO").Value & ""

                Dim PO_SHIPMENT_LNO As Int32 = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                Dim dvw As DataView = DirectCast(grdPOTSHIP7.DataSource, DataTable).DefaultView
                dvw.RowFilter = "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)

                Setup_grdPOTSHIP8()
            End If
        End If

        If grdPOTSHIP2.ActiveRow IsNot Nothing AndAlso grdPOTSHIP2.ActiveRow.IsDataRow Then
            Dim COLUMN_NAME = "COMM_INV_NO"
            ' "CONTAINER_NO"
            If Trim(grdPOTSHIP2.ActiveRow.Cells(COLUMN_NAME).Value & "") = "" And
                 (grdPOTSHIP2.ActiveCell Is Nothing OrElse grdPOTSHIP2.ActiveCell.Column.Key <> COLUMN_NAME) Then
                grdPOTSHIP2.ActiveCell = grdPOTSHIP2.ActiveRow.Cells(COLUMN_NAME)
                Exit Sub
            End If
        End If

        If grdPOTSHIP2.ActiveRow IsNot Nothing Then
            If grdPOTSHIP2.ActiveRow.IsAddRow Then
                grdPOTSHIP3.Visible = False
                If grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO").Value & "" = "" Then
                    If grdPOTSHIP2.ActiveCell Is Nothing OrElse grdPOTSHIP2.ActiveCell.Column.Key <> "CONTAINER_NO" Then
                        grdPOTSHIP2.ActiveCell = grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO")
                    End If
                End If
                Exit Sub
            Else
                If receipt_mode Then
                    Set_Receipt_Options()
                End If
                Set_Container()
            End If
        End If
    End Sub

    Private Sub grdPOTSHIP2_AfterRowCancelUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTSHIP2.AfterRowCancelUpdate

    End Sub

    Private Sub grdPOTSHIP2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTSHIP2.AfterRowsDeleted
        Delete_Rows_from_Summary_Tables()
        Setup_grdPOTSHIP2_ActiveRow()
    End Sub

    Private Sub grdPOTSHIP2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTSHIP2.AfterRowUpdate

    End Sub

    Private Sub grdPOTSHIP2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdPOTSHIP2.BeforeCellUpdate

    End Sub

    Private Sub grdPOTSHIP2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTSHIP2.BeforeExitEditMode
        With grdPOTSHIP2.ActiveCell
            Select Case .Column.Key
                'Case "CUST_CODE"
                '    'If .Row.IsAddRow Then
                '    If .Text <> "" Then
                '        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                '    End If
                '    'End If
            End Select
        End With
    End Sub

    Private Sub grdPOTSHIP2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTSHIP2.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Not grow.IsAddRow Then
                Dim PO_SHIPMENT_LNO As Integer = Val(grow.Cells("PO_SHIPMENT_LNO").Value & "")
                Dim sqlw As String = "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " AND PO_QTY_REC <> 0"
                If Val(dst.Tables("POTSHIP3").Compute _
                       ("Count(PO_SHIPMENT_LNO)", sqlw) & "") <> 0 Then
                    MsgBox("Receipts have been entered", MsgBoxStyle.OkOnly, "Deletion Denied")
                    e.Cancel = True
                End If
            End If
        Next
    End Sub

    Private Sub grdPOTSHIP2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTSHIP2.BeforeRowUpdate
        If grdPOTSHIP2.Visible = False Then
            Exit Sub
        End If

        With grdPOTSHIP2
            If e.Row.Cells("CONTAINER_NO").Value & "" = "" Then
                MsgBox("Container No is Mandatory", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            ElseIf e.Row.Cells("BOL_NO").Value & "" = "" Then
                MsgBox(e.Row.Cells("BOL_NO").Column.Header.Caption & " is Mandatory", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If Not e.Cancel Then
                If Val(e.Row.Cells("PO_SHIPMENT_LNO").Value & "") = 0 Then
                    .ActiveRow.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO
                    .ActiveRow.Cells("PO_SHIPMENT_LNO").Value = Val(dst.Tables("POTSHIP2").Compute("Max(PO_SHIPMENT_LNO)", "") & "") + 1
                    .ActiveRow.Cells("PO_SHIP_STATUS").Value = "O"
                    .ActiveRow.Cells("ACCRUAL_STATUS").Value = "0"
                End If

                ' .ActiveRow.Cells("OPS_YYYYPP").Value = Set_Period()

                Set_Landed_Cost_Needs_to_be_Calculated_Indicator(False)
            End If
        End With
    End Sub

    Private Sub grdPOTSHIP2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ACTION" ' "PO_SHIP_STATUS"

                If (WH_REC_NOsInProcess.Count > 0) Then
                    If grdPOTSHIP2.ActiveRow.IsAddRow Then
                        Exit Sub
                    Else
                        Dim WH_REC_NO As String = grdPOTSHIP2.ActiveRow.Cells("WH_REC_NO").Value & ""
                        If WH_REC_NO <> "" Then
                            If Not WH_REC_NOsInProcess.Contains(WH_REC_NO) Then
                                MsgBox("WH Receipt posted against this line is Complete - no changes permitted", MsgBoxStyle.OkOnly, "Cannot Add PO Lines")
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                If ship_entry Then
                    If EntryMode = "N" Or EntryMode = "E" Then
                        If grdPOTSHIP2.ActiveRow.IsAddRow Then
                            grdPOTSHIP2.ActiveRow.Update()
                            If Not grdPOTSHIP2.ActiveRow.IsAddRow Then
                                Setup_grdPOTSHIP2_ActiveRow()
                                Select_POs()
                            End If
                        Else
                            Select_POs()
                        End If
                    End If

                ElseIf receipt_mode Then
                    If e.Cell.Row.IsAddRow Then Exit Sub
                    'O -> X Clear Receipt of this Container
                    'X -> O Receive this BOL
                    'C -> R Un-Reverse Receipt of this Container
                    'R -> C Reverse Receipt of this Container

                    With e.Cell.Row.Cells("PO_SHIP_STATUS")
                        If .Value & "" = "O" Then
                            .Value = "X"
                            If Not select_from_3PL_list And Not Select_from_Whse_Receipt Then
                                Dim sql As String = "PO_SHIPMENT_LNO = " & e.Cell.Row.Cells("PO_SHIPMENT_LNO").Value
                                For Each row As DataRow In dst.Tables("POTSHIP3").Select(sql)
                                    row.Item("PO_QTY_REC") = row.Item("PO_QTY_SHP")
                                Next
                            End If

                            ' the next block of code is repeated NEARLY identically in a section below dealing with reverse receipts
                            Dim ORDR_NO As String = e.Cell.Row.Cells("ORDR_NO").Value & ""
                            ' We bill the customer on BTB orders if receiving into a warehouse whose type is P
                            If ORDR_NO <> "" And (WHSE_TYPE = "P" Or (ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "NC")) Then

                                If dst.Tables("SOTORDP1").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                                    Dim rowSOTORDP1 As DataRow = dst.Tables("SOTORDP1").NewRow
                                    rowSOTORDP1.Item("ORDR_NO") = ORDR_NO
                                    rowSOTORDP1.Item("INV_NO") = ORDR_NO
                                    If Absx1.dteFor("PO_DATE_RECEIVED").Value & "" <> "" Then rowSOTORDP1.Item("INV_DATE") = Absx1.dteFor("PO_DATE_RECEIVED").Value
                                    If Absx1.txtFor("PO_SOURCE_DOC").Text & "" <> "" Then rowSOTORDP1.Item("INV_REF") = Absx1.txtFor("PO_SOURCE_DOC").Text
                                    dst.Tables("SOTORDP1").Rows.Add(rowSOTORDP1)

                                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                                        ASCMAIN1.sql = "Select * from SOTORDP1 where ORDR_NO = '" & ORDR_NO & "'"
                                        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                                        If rows.Count = 1 Then
                                            rowSOTORDP1.Item("INV_NO_PREV") = rows(0).Item("INV_NO")
                                            rowSOTORDP1.Item("INV_DATE") = rows(0).Item("INV_DATE")
                                            rowSOTORDP1.Item("INV_REF") = rows(0).Item("INV_REF")
                                            rowSOTORDP1.Item("INV_COMMENT") = rows(0).Item("INV_COMMENT")
                                        End If
                                    End If

                                    Fill_Records("SOTORDR2", ORDR_NO, False)

                                End If

                                Dim sql As String = "PO_SHIPMENT_LNO = " & e.Cell.Row.Cells("PO_SHIPMENT_LNO").Value
                                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select(sql)

                                    Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find _
                                                                 (New Object() {rowPOTSHIP3.Item("PO_ORDER_NO"),
                                                                                rowPOTSHIP3.Item("PO_ORDER_LNO")})
                                    Dim ORDR_LNO As Integer = Val(rowPOTORDR2.Item("ORDR_LNO") & "")

                                    Dim rowSOTORDP2 As DataRow = dst.Tables("SOTORDP2").Rows.Find(New Object() {ORDR_NO, ORDR_NO, ORDR_LNO})
                                    If rowSOTORDP2 Is Nothing Then
                                        rowSOTORDP2 = dst.Tables("SOTORDP2").NewRow
                                        rowSOTORDP2.Item("ORDR_NO") = ORDR_NO
                                        rowSOTORDP2.Item("INV_NO") = ORDR_NO
                                        rowSOTORDP2.Item("ORDR_LNO") = ORDR_LNO
                                        rowSOTORDP2.Item("ORDR_QTY_SHIP") = 0
                                        dst.Tables("SOTORDP2").Rows.Add(rowSOTORDP2)
                                    End If
                                    rowSOTORDP2.Item("ORDR_QTY_SHIP") += Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "")
                                Next
                            End If

                        ElseIf .Value & "" = "X" Then
                            .Value = "O"

                            Dim ORDR_NO As String = e.Cell.Row.Cells("ORDR_NO").Value & ""

                            Dim sql As String = "PO_SHIPMENT_LNO = " & e.Cell.Row.Cells("PO_SHIPMENT_LNO").Value
                            For Each row As DataRow In dst.Tables("POTSHIP3").Select(sql)

                                If ORDR_NO <> "" Then
                                    Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find _
                                                               (New Object() {row.Item("PO_ORDER_NO"),
                                                                              row.Item("PO_ORDER_LNO")})
                                    Dim ORDR_LNO As Integer = Val(rowPOTORDR2.Item("ORDR_LNO") & "")

                                    Dim rowSOTORDP2 As DataRow = dst.Tables("SOTORDP2").Rows.Find(New Object() {ORDR_NO, ORDR_NO, ORDR_LNO})
                                    rowSOTORDP2.Item("ORDR_QTY_SHIP") -= Val(row.Item("PO_QTY_REC") & "")
                                    If Val(rowSOTORDP2.Item("ORDR_QTY_SHIP") & "") = 0 Then
                                        rowSOTORDP2.Delete()
                                    End If
                                End If

                                row.Item("PO_QTY_REC") = 0
                            Next

                            If e.Cell.Row.Cells("ORDR_NO").Value & "" <> "" Then
                                Dim row As DataRow = dst.Tables("SOTORDP1").Select("ORDR_NO = '" & ORDR_NO & "'")(0) ' expecting 1 and only 1
                                If row.GetChildRows("SOTORDP1_SOTORDP2").Count = 0 Then
                                    row.Delete()
                                End If
                                '    ASCDATA1.DeleteRows(dst.Tables("SOTORDP1"), "ORDR_NO = '" & e.Cell.Row.Cells("ORDR_NO").Value & "'") Then
                            End If

                        ElseIf .Value & "" = "C" Then
                            ' Status was Received, clicked Reverse Receipt, Status will now show Reverse Now, button will now show Cancel Reverse
                            Dim ORDR_NO As String = e.Cell.Row.Cells("ORDR_NO").Value & ""
                            ' We bill the customer on BTB orders if receiving into a warehouse whose type is P
                            If ORDR_NO <> "" And (WHSE_TYPE = "P" Or (ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "NC")) Then

                                If dst.Tables("SOTORDP1").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                                    Dim rowSOTORDP1 As DataRow = dst.Tables("SOTORDP1").NewRow
                                    rowSOTORDP1.Item("ORDR_NO") = ORDR_NO
                                    rowSOTORDP1.Item("INV_NO") = ORDR_NO
                                    Dim rowICTIREC1 As DataRow = LookUp("ICTIREC1", e.Cell.Row.Cells("TRAN_NO").Value)
                                    rowSOTORDP1.Item("INV_DATE") = rowICTIREC1.Item("RECEIPT_DATE")
                                    rowSOTORDP1.Item("INV_REF") = rowICTIREC1.Item("SOURCE_DOC_NO")
                                    dst.Tables("SOTORDP1").Rows.Add(rowSOTORDP1)

                                    Fill_Records("SOTORDR2", ORDR_NO, False)

                                End If

                                Dim sql As String = "PO_SHIPMENT_LNO = " & e.Cell.Row.Cells("PO_SHIPMENT_LNO").Value
                                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select(sql)

                                    Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find _
                                                                 (New Object() {rowPOTSHIP3.Item("PO_ORDER_NO"),
                                                                                rowPOTSHIP3.Item("PO_ORDER_LNO")})
                                    Dim ORDR_LNO As Integer = Val(rowPOTORDR2.Item("ORDR_LNO") & "")

                                    Dim rowSOTORDP2 As DataRow = dst.Tables("SOTORDP2").Rows.Find(New Object() {ORDR_NO, ORDR_NO, ORDR_LNO})
                                    If rowSOTORDP2 Is Nothing Then
                                        rowSOTORDP2 = dst.Tables("SOTORDP2").NewRow
                                        rowSOTORDP2.Item("ORDR_NO") = ORDR_NO
                                        rowSOTORDP2.Item("INV_NO") = ORDR_NO
                                        rowSOTORDP2.Item("ORDR_LNO") = ORDR_LNO
                                        rowSOTORDP2.Item("ORDR_QTY_SHIP") = 0
                                        dst.Tables("SOTORDP2").Rows.Add(rowSOTORDP2)
                                    End If

                                    rowSOTORDP2.Item("ORDR_QTY_SHIP") -= Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "")

                                Next
                            End If

                            'If e.Cell.Row.Cells("ORDR_NO").Value & "" <> "" And WHSE_TYPE = "P" Then
                            'MsgBox("You May NOT De-Receive a BTB Shipment which has been Invoiced", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                            'Else
                            .Value = "R"
                            'End If
                        ElseIf .Value & "" = "R" Then
                            ' Status was Reverse Now, clicked Cancel Reverse, Status will now show Received, button will now show Reverse Receipt

                            Dim ORDR_NO As String = e.Cell.Row.Cells("ORDR_NO").Value & ""
                            ' remove this POTSHIP2 record's contribution to the BTB sales order

                            If ORDR_NO <> "" And (WHSE_TYPE = "P" Or (ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "NC")) Then

                                Dim sql As String = "PO_SHIPMENT_LNO = " & e.Cell.Row.Cells("PO_SHIPMENT_LNO").Value
                                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select(sql)
                                    Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find _
                                                                 (New Object() {rowPOTSHIP3.Item("PO_ORDER_NO"),
                                                                                rowPOTSHIP3.Item("PO_ORDER_LNO")})
                                    Dim ORDR_LNO As Integer = Val(rowPOTORDR2.Item("ORDR_LNO") & "")
                                    Dim rowSOTORDP2 As DataRow = dst.Tables("SOTORDP2").Rows.Find(New Object() {ORDR_NO, ORDR_NO, ORDR_LNO})
                                    rowSOTORDP2.Item("ORDR_QTY_SHIP") += Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "")
                                    If Val(rowSOTORDP2.Item("ORDR_QTY_SHIP") & "") = 0 Then
                                        rowSOTORDP2.Delete()
                                    End If
                                Next

                                If dst.Tables("SOTORDP2").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                                    Dim rowSOTORDP1 As DataRow = dst.Tables("SOTORDP1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)
                                    rowSOTORDP1.Delete()
                                End If
                            End If



                            .Value = "C"
                        Else
                            Stop
                        End If
                        grdPOTSHIP2.ActiveRow.Update()
                        '                        e.Cell.Row.Update()
                    End With
                    Set_Receipt_Options()
                End If


            Case ""
                Dim sql_where As String = ""
                grdClickCellButton(grdPOTSHIP2, sql_where, sql_where <> "")
        End Select
    End Sub

    Private Sub grdPOTSHIP2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdPOTSHIP2.Error
        grdPOTSHIP2.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdPOTSHIP2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTSHIP2.InitializeRow
        'If ship_entry Then
        '    'If Not e.Row.IsAddRow Then
        '    e.Row.Cells("ACTION").Value = "Select POs"
        '    'End If
        'End If

        If e.Row.Cells("ORDR_NO").Value & "" <> "" Then
            e.Row.Cells("PO_SHIPMENT_LNO").Appearance.BackColor = Drawing.Color.LightBlue
            e.Row.Cells("PO_SHIPMENT_LNO").ToolTipText = "This line is associated with a BTB Order"
        Else
            e.Row.Cells("PO_SHIPMENT_LNO").Appearance.BackColor = Drawing.Color.Empty
        End If

        If ship_entry Then
            If WH_REC_NOsInProcess.Count <> 0 Then
                Dim WH_REC_NO As String = e.Row.Cells("WH_REC_NO").Value & ""
                If WH_REC_NO = "" Then
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Empty
                    e.Row.Cells("PO_SHIP_STATUS").ToolTipText = "Warehouse Receipt not Started"
                ElseIf WH_REC_NOsInProcess.Contains(WH_REC_NO) Then
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.BackColor = Drawing.Color.Yellow
                    e.Row.Cells("PO_SHIP_STATUS").ToolTipText = "Partially Received by the Warehouse"
                Else
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("PO_SHIP_STATUS").ToolTipText = "Completely Received by the Warehouse"
                End If
            End If
        End If


        If receipt_mode OrElse e.Row.Cells("TRAN_NO").Value & "" <> "" Then
            If Val(e.Row.Cells("PO_QTY_SHP").Value & "") <> Val(e.Row.Cells("PO_QTY_REC").Value & "") Then
                e.Row.Cells("PO_QTY_REC").Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Cells("PO_QTY_REC").Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If

        If receipt_mode Then
            'If Not e.Row.IsAddRow Then
            'If Val(e.Row.Cells("PO_QTY_SHP").Value & "") <> Val(e.Row.Cells("PO_QTY_REC").Value & "") Then
            '    e.Row.Cells("PO_QTY_REC").Appearance.ForeColor = Drawing.Color.Red
            'Else
            '    e.Row.Cells("PO_QTY_REC").Appearance.ForeColor = Drawing.Color.Empty
            'End If
            Select Case e.Row.Cells("PO_SHIP_STATUS").Value
                Case "C"
                    e.Row.Cells("ACTION").Value = "Reverse Receipt"
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.BackColor = Drawing.Color.LightGray
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Empty
                Case "O"
                    e.Row.Cells("ACTION").Value = "Receive"
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.BackColor = Drawing.Color.LightGreen
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Empty
                Case "R"
                    e.Row.Cells("ACTION").Value = "Cancel Reverse"
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.BackColor = Drawing.Color.Empty
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Red
                Case "X"
                    e.Row.Cells("ACTION").Value = "Cancel Receipt"
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.BackColor = Drawing.Color.Empty
                    e.Row.Cells("PO_SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Green
                Case Else
                    Stop

            End Select
            'End If

            'O -> X Clear Receipt of this Container
            'X -> O Receive this BOL
            'C -> R Un-Reverse Receipt of this Container
            'R -> C Reverse Receipt of this Container

            If Val(e.Row.Cells("LINES_SHORT").Value & "") > 0 Then
                e.Row.Cells("LINES_SHORT").Appearance.ForeColor = Drawing.Color.White
                e.Row.Cells("LINES_SHORT").Appearance.BackColor = Drawing.Color.Red
            Else
                e.Row.Cells("LINES_SHORT").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("LINES_SHORT").Appearance.BackColor = Drawing.Color.Empty
            End If
            If Val(e.Row.Cells("LINES_ZERO").Value & "") > 0 Then
                e.Row.Cells("LINES_ZERO").Appearance.BackColor = Drawing.Color.Red
                e.Row.Cells("LINES_ZERO").Appearance.ForeColor = Drawing.Color.White
            Else
                e.Row.Cells("LINES_ZERO").Appearance.BackColor = Drawing.Color.Empty
                e.Row.Cells("LINES_ZERO").Appearance.ForeColor = Drawing.Color.Empty
            End If
            If Val(e.Row.Cells("LINES_OVER").Value & "") > 0 Then
                e.Row.Cells("LINES_OVER").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("LINES_OVER").Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Cells("LINES_OVER").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("LINES_OVER").Appearance.BackColor = Drawing.Color.Empty
            End If

        End If

        If cost_calc Then
            Dim TOTAL_FIRST As Decimal = Val(e.Row.Cells("TOTAL_FIRST").Value & "")
            Dim TOTAL_FIRST_CALC As Decimal = Val(e.Row.Cells("TOTAL_FIRST_CALC").Value & "")

            With e.Row.Cells("TOTAL_FIRST")
                If System.Math.Abs(TOTAL_FIRST_CALC - TOTAL_FIRST) > 1 Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText = "Calculated Value of " & Format(TOTAL_FIRST_CALC, "#,##0.00") & " varies by > $1"
                Else
                    .Appearance.ForeColor = Drawing.Color.Empty
                    .ToolTipText = ""
                End If
            End With

            Dim TOTAL_LANDED As Decimal = Val(e.Row.Cells("TOTAL_LANDED").Value & "")
            Dim TOTAL_LANDED_CALC As Decimal = Val(e.Row.Cells("TOTAL_LANDED_CALC").Value & "")

            With e.Row.Cells("TOTAL_LANDED")
                If System.Math.Abs(TOTAL_LANDED_CALC - TOTAL_LANDED) > 2 Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText = "Calculated Value of " & Format(TOTAL_LANDED_CALC, "#,##0.00") & " varies by > $2"
                Else
                    .Appearance.ForeColor = Drawing.Color.Empty
                    .ToolTipText = ""
                End If
            End With

            If e.Row.Cells("OPS_YYYYPP_FIFO").Value & "" <> "" And e.Row.Cells("OPS_YYYYPP_FIFO").Value & "" <> e.Row.Cells("OPS_YYYYPP").Value & "" Then
                e.Row.Cells("OPS_YYYYPP").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("OPS_YYYYPP").ToolTipText = "This Receipt is used in FIFO calculation for Previous Period"
            Else
                e.Row.Cells("OPS_YYYYPP").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("OPS_YYYYPP").ToolTipText = ""
            End If
        End If
    End Sub

    Private Sub grdPOTSHIP2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdPOTSHIP2.KeyPress
        'If grdPOTSHIP2.ActiveRow IsNot Nothing Then
        '    Try
        '        If grdPOTSHIP2.ActiveCell.Column.Key = "CUST_NAME" Then
        '            If grdPOTSHIP2.ActiveRow.Cells("CUST_CODE").Text <> "" Then
        '                e.KeyChar = Chr(0)
        '                e.Handled = True
        '            End If
        '        End If
        '    Catch ex As Exception

        '    End Try
        'End If
    End Sub

#End Region

#Region "grdPOTSHIP3"

    Private Sub grdPOTSHIP3_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "DUTY_RATE_CODE"
                Dim DUTY_RATE_CODE As String = e.Cell.Value & ""
                Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", DUTY_RATE_CODE)
                If rowICTDUTY1 IsNot Nothing Then
                    e.Cell.Row.Cells("DUTY_RATE").Value = rowICTDUTY1.Item("DUTY_RATE")
                End If

                Set_Landed_Cost_Needs_to_be_Calculated_Indicator(True)
        End Select
    End Sub

    Private Sub grdPOTSHIP3_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIP3.AfterRowActivate
        With grdPOTSHIP3.ActiveRow
            If Trim(.Cells("STYLE_CODE").Value & "") = "" And (grdPOTSHIP3.ActiveCell Is Nothing OrElse grdPOTSHIP3.ActiveCell.Column.Key <> "STYLE_CODE") Then
                grdPOTSHIP3.ActiveCell = .Cells("STYLE_CODE")
                Exit Sub
            End If

            If Not receipt_mode And (EntryMode = "N" Or EntryMode = "E") Then
                If Val(.Cells("PO_QTY_REC").Value & "") = 0 Then
                    .Cells("PO_QTY_SHP").Column.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Cells("PO_QTY_SHP").Column.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End If

            If .IsAddRow Then
                If .Cells("STYLE_CODE").Value & "" = "" Then
                    If grdPOTSHIP3.ActiveCell.Column.Key = "STYLE_CODE" Then
                        grdPOTSHIP3.ActiveCell = .Cells("STYLE_CODE")
                    End If
                End If
                Exit Sub
            End If
        End With
    End Sub

    Private Sub grdPOTSHIP3_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTSHIP3.AfterRowsDeleted
        Dim DT As DataTable = ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3").Select(""), New String() {"PO_SHIPMENT_LNO", "STYLE_CODE", "COLOR_CODE"})
        DT.PrimaryKey = New DataColumn() {DT.Columns("PO_SHIPMENT_LNO"), DT.Columns("STYLE_CODE"), DT.Columns("COLOR_CODE")}
        dst.Tables("POTSHIPR").AcceptChanges()
        If dst.Tables("POTSHIPR").Rows.Count > 0 Then
            For R As Integer = dst.Tables("POTSHIPR").Rows.Count - 1 To 0 Step -1
                Dim rowPOTSHIPR As DataRow = dst.Tables("POTSHIPR").Rows(R)
                If DT.Rows.Count > 0 Then
                    If DT.Rows.Find(New String() {rowPOTSHIPR.Item("PO_SHIPMENT_LNO"), rowPOTSHIPR.Item("STYLE_CODE"), rowPOTSHIPR.Item("COLOR_CODE")}) Is Nothing Then
                        rowPOTSHIPR.Delete()
                    End If
                Else
                    rowPOTSHIPR.Delete()
                End If
            Next
        End If

        DT = ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3").Select(""), New String() {"PO_ORDER_NO", "PO_ORDER_LNO"})
        DT.PrimaryKey = New DataColumn() {DT.Columns("PO_ORDER_NO"), DT.Columns("PO_ORDER_LNO")}
        dst.Tables("POTORDRO").AcceptChanges()
        If dst.Tables("POTORDRO").Rows.Count > 0 Then
            For R As Integer = dst.Tables("POTORDRO").Rows.Count - 1 To 0 Step -1
                Dim rowPOTORDRO As DataRow = dst.Tables("POTORDRO").Rows(R)
                If DT.Rows.Count > 0 Then
                    If DT.Rows.Find(New String() {rowPOTORDRO.Item("PO_ORDER_NO"), rowPOTORDRO.Item("PO_ORDER_LNO")}) Is Nothing Then
                        rowPOTORDRO.Delete()
                    End If
                Else
                    rowPOTORDRO.Delete()
                End If
            Next
        End If

        Delete_Rows_from_Summary_Tables()

    End Sub

    Private Sub grdPOTSHIP3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTSHIP3.AfterRowUpdate
        With e.Row
            Dim SUB_UNIT_PACK_QTY As Integer = Val(.Cells("SUB_UNIT_PACK_QTY").Value & "")
            If SUB_UNIT_PACK_QTY = 0 Then SUB_UNIT_PACK_QTY = 1
            Dim SPQ As Decimal = (12 / SUB_UNIT_PACK_QTY)
            If ship_entry Then
                ' .Cells("PO_QTY_SHP_DZ").Value = Val(.Cells("PO_QTY_SHP").Value / SPQ & "")
            Else
                Dim NEW_FIRST_COST As Decimal
                If optUD.Value = "U" Then
                    ' .Cells("PO_QTY_SHP_DZ").Value = Val(.Cells("PO_QTY_SHP").Value / SPQ & "")
                    For Each COLUMN_NAME As String In New String() {"PO_COST_MATLS", "PO_COST_VCOST", "PO_COST_OTHER", "PO_COST_QUOTA", "PO_COST_QUOTA_DF"}
                        .Cells(COLUMN_NAME & "_DZ").Value = Val(.Cells(COLUMN_NAME).Value & "") * 12
                    Next
                    NEW_FIRST_COST = (Val(.Cells("PO_COST_MATLS").Value & "") * 12) _
                                    + (Val(.Cells("PO_COST_VCOST").Value & "") * 12) _
                                    + (Val(.Cells("PO_COST_OTHER").Value & "") * 12)
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        ' this column is read only
                        ' .Cells("FIRST_COST_TOTAL_DZ").Value = NEW_FIRST_COST
                    End If
                    ' this column is read only
                    ' .Cells("FIRST_COST_TOTAL").Value = NEW_FIRST_COST / 12
                Else
                    ' .Cells("PO_QTY_SHP").Value = Val(.Cells("PO_QTY_SHP_DZ").Value * SPQ)
                    For Each COLUMN_NAME As String In New String() {"PO_COST_MATLS", "PO_COST_VCOST", "PO_COST_OTHER", "PO_COST_QUOTA", "PO_COST_QUOTA_DF"}
                        .Cells(COLUMN_NAME).Value = Val(.Cells(COLUMN_NAME & "_DZ").Value & "") / SPQ
                    Next
                    NEW_FIRST_COST = (Val(.Cells("PO_COST_MATLS_DZ").Value & "") / SPQ) _
                        + (Val(.Cells("PO_COST_VCOST_DZ").Value & "") / SPQ) _
                        + (Val(.Cells("PO_COST_OTHER_DZ").Value & "") / SPQ)
                    '.Cells("FIRST_COST_TOTAL").Value = NEW_FIRST_COST
                    '.Cells("FIRST_COST_TOTAL_DZ").Value = NEW_FIRST_COST * 12
                End If
            End If
        End With

        If receipt_mode Then
            If (WHSE_TYPE = "P" Or (ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "NC")) Then
                Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {e.Row.Cells("PO_ORDER_NO").Value, e.Row.Cells("PO_ORDER_LNO").Value})
                If rowPOTORDR2.Item("ORDR_NO") & "" <> "" Then
                    Dim rowSOTORDP2 As DataRow = dst.Tables("SOTORDP2").Rows.Find(New Object() {rowPOTORDR2.Item("ORDR_NO"),
                                                                                                rowPOTORDR2.Item("ORDR_NO"),
                                                                                                rowPOTORDR2.Item("ORDR_LNO")})
                    Dim sql As String = "PO_ORDER_NO = '" & e.Row.Cells("PO_ORDER_NO").Value & "' and PO_ORDER_LNO = " & e.Row.Cells("PO_ORDER_LNO").Value
                    Dim ORDR_QTY_SHIP As Int64 = Val(dst.Tables("POTSHIP3").Compute("SUM(PO_QTY_REC)", sql) & "")
                    rowSOTORDP2.Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
                End If
            End If
        End If

        Set_Landed_Cost_Needs_to_be_Calculated_Indicator(False)
        'grdPOTSHIP3.ActiveRow.Update()
    End Sub

    Private Sub grdPOTSHIP3_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdPOTSHIP3.BeforeCellUpdate

        'If e.Cell.Row.IsAddRow Then
        '    If Absx1.dteFor("PO_DATE_SHIP_BY").Value & "" = "" Or Absx1.dteFor("PO_DATE_ETA").Value & "" = "" Then
        '        MsgBox("Please provide default values for Ship-By and ETA Date above in the PO Header before entering PO Details", _
        '               MsgBoxStyle.OkOnly, "Cannot Enter New PO Details")
        '        e.Cancel = True
        '        Exit Sub
        '    End If
        'End If


        Select Case e.Cell.Column.Key

            Case "DUTY_RATE_CODE"
                If e.NewValue & "" <> "" Then
                    Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", New String() {e.NewValue})
                    If rowICTDUTY1 Is Nothing Then
                        MsgBox("Duty Rate Code " & e.NewValue & " is not valid",
                               MsgBoxStyle.OkOnly, "Invalid Value for Duty Rate Code")
                        e.Cancel = True
                    End If
                End If

        End Select
    End Sub

    Private Sub grdPOTSHIP3_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTSHIP3.BeforeExitEditMode

    End Sub

    Private Sub grdPOTSHIP3_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTSHIP3.BeforeRowActivate

    End Sub

    Private Sub grdPOTSHIP3_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTSHIP3.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Not grow.IsAddRow Then
                If receipt_mode Then
                    If grdPOTSHIP2.ActiveRow.Cells("PO_SHIP_STATUS").Value = "X" Then
                        grow.Cells("PO_QTY_REC").Value = 0
                    End If
                Else
                    If Val(grow.Cells("PO_QTY_REC").Value & "") <> 0 Then
                        MsgBox("Receipts have been entered", MsgBoxStyle.OkOnly, "Deletion Denied")
                        e.Cancel = True
                    Else
                        If grow.Cells("PO_SHIP_STATUS").Value & "" <> "O" Then
                            ASCMAIN1.sql = "Select * from ICTTRAN1 where TRAN_TYPE_ORIG = 'P'" & vbCrLf _
                            & " and TRAN_NO_ORIG = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                            & " and TRAN_STATUS_UPD = 'U'"
                            Dim row As DataRow = ASCDATA1.GetDataRow
                            If row IsNot Nothing Then
                                MsgBox("Receipts have been entered", MsgBoxStyle.OkOnly, "Deletion Denied")
                                e.Cancel = True
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If receipt_mode Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdPOTSHIP3_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTSHIP3.BeforeRowUpdate
        If Not grdPOTSHIP3.Visible Then Exit Sub
    End Sub

    Private Sub grdPOTSHIP3_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP3.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdPOTSHIP3.ActiveCell.Column.Key
            'Case "ACCT_CODE"
            '    sql_where = "NVL(ACCT_STATUS,'X') = 'A' and NVL(ACCT_SUB_CTL,'0') <> '1'"
        End Select

        grdClickCellButton(grdPOTSHIP3, sql_where, False)
    End Sub

    Private Sub grdPOTSHIP3_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTSHIP3.DoubleClickRow
        If cost_calc And Not InquiryMode And EntryMode = "E" And e.Row.IsDataRow And grdPOTSHIP3.ActiveCell IsNot Nothing Then
            If grdPOTSHIP3.ActiveCell.Column.Key = "PO_COST_COMM" Then
                Dim PO_COST_COMM As Decimal = Val(e.Row.Cells("PO_COST_COMM").Value & "")
                If MsgBox("Do You Want To Update All Commissions on this BOL to " & Format(PO_COST_COMM, "#0.00") & "%",
                          MsgBoxStyle.YesNo, "Commission Update") = MsgBoxResult.Yes Then
                    Dim PO_SHIPMENT_NO As String = e.Row.Cells("PO_SHIPMENT_NO").Value
                    Dim PO_SHIPMENT_LNO As Integer = Val(e.Row.Cells("PO_SHIPMENT_LNO").Value & "")
                    For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select _
                            ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO))
                        rowPOTSHIP3.Item("PO_COST_COMM") = PO_COST_COMM
                    Next
                End If
            End If
            If grdPOTSHIP3.ActiveCell.Column.Key = "PO_COST_BUFFER" Then
                Dim PO_COST_BUFFER As Decimal = Val(e.Row.Cells("PO_COST_BUFFER").Value & "")
                If MsgBox("Do You Want To Update All Commission Buffers on this BOL to " & Format(PO_COST_BUFFER, "#0.00") & "%",
                          MsgBoxStyle.YesNo, "Commission Update") = MsgBoxResult.Yes Then
                    Dim PO_SHIPMENT_NO As String = e.Row.Cells("PO_SHIPMENT_NO").Value
                    Dim PO_SHIPMENT_LNO As Integer = Val(e.Row.Cells("PO_SHIPMENT_LNO").Value & "")
                    For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select _
                            ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO))
                        rowPOTSHIP3.Item("PO_COST_BUFFER") = PO_COST_BUFFER
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub grdPOTSHIP3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTSHIP3.InitializeRow
        If e.Row.Cells("PO_QTY_REC").Column.CellActivation = UltraWinGrid.Activation.NoEdit Then
            e.Row.Cells("PO_QTY_REC").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("PO_QTY_REC_DZ").Appearance.BackColor = Drawing.Color.Empty
        Else
            e.Row.Cells("PO_QTY_REC").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("PO_QTY_REC_DZ").Appearance.BackColor = Drawing.Color.Yellow
        End If

        If receipt_mode OrElse (grdPOTSHIP2.ActiveRow IsNot Nothing AndAlso grdPOTSHIP2.ActiveRow.Cells("TRAN_NO").Value & "" <> "") Then
            If grdPOTSHIP2.ActiveRow.Cells("PO_SHIP_STATUS").Value & "" <> "O" And Val(e.Row.Cells("PO_QTY_SHP").Value & "") > Val(e.Row.Cells("PO_QTY_REC").Value & "") Then
                e.Row.Cells("PO_QTY_VAR").Appearance.ForeColor = Drawing.Color.White
                e.Row.Cells("PO_QTY_VAR").Appearance.BackColor = Drawing.Color.Red
            ElseIf grdPOTSHIP2.ActiveRow.Cells("PO_SHIP_STATUS").Value & "" <> "O" And Val(e.Row.Cells("PO_QTY_SHP").Value & "") < Val(e.Row.Cells("PO_QTY_REC").Value & "") Then
                e.Row.Cells("PO_QTY_VAR").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("PO_QTY_VAR").Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Cells("PO_QTY_VAR").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("PO_QTY_VAR").Appearance.BackColor = Drawing.Color.Empty
            End If
        End If

        ' Dim rowPOTORDRO As DataRow = dst.Tables("POTORDRO").Rows.Find(New Object() {e.Row.Cells("PO_ORDER_NO").Value, e.Row.Cells("PO_ORDER_LNO").Value})
        Dim PO_QTY_OPN_PRE As Int64 = Val(e.Row.Cells("PO_QTY_OPN_PRE").Value & "")
        Dim PO_QTY_SHP As Int64 = Val(e.Row.Cells("PO_QTY_SHP").Value & "")
        Dim sqlw As String = "PO_ORDER_NO = '" & e.Row.Cells("PO_ORDER_NO").Value & "' and PO_ORDER_LNO = " & e.Row.Cells("PO_ORDER_LNO").Value
        Dim PO_QTY_SHP_others As Int64 = Val(dst.Tables("POTSHIP3").Compute("SUM(PO_QTY_SHP)", sqlw) & "") - PO_QTY_SHP
        Dim NET_OPEN As Int64 = PO_QTY_OPN_PRE - PO_QTY_SHP_others - PO_QTY_SHP ' Val(rowPOTORDRO.Item("PO_QTY_OPN_PRE") & "") - Val(rowPOTORDRO.Item("PO_QTY_SHP") & "")
        If NET_OPEN < 0 Then
            e.Row.Cells("PO_QTY_SHP").Appearance.BackColor = Drawing.Color.Orange
            Dim msg As String = "Total Qty Shipped > Qty Open PO"
            If PO_QTY_SHP_others <> 0 Then
                msg &= vbCrLf & "There are multiple lines on this Shipment referring to this PO Detail"
            End If
            e.Row.Cells("PO_QTY_SHP").ToolTipText = msg
        Else
            e.Row.Cells("PO_QTY_SHP").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("PO_QTY_SHP").ToolTipText = ""
        End If


        If cost_calc Then
            Dim EXT_FIRST As Decimal = Val(e.Row.Cells("EXT_FIRST").Value & "")
            Dim EXT_FIRST_CALC As Decimal = Val(e.Row.Cells("EXT_FIRST_CALC").Value & "")

            With e.Row.Cells(IIf(optUD.Value = "U", "FIRST_COST_TOTAL", "FIRST_COST_TOTAL_DZ"))
                If System.Math.Abs(EXT_FIRST_CALC - EXT_FIRST) > 1 Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText = "Calculated Value of " & Format(EXT_FIRST_CALC, "#,##0.00") & " varies by > $1"
                Else
                    .Appearance.ForeColor = Drawing.Color.Empty
                    .ToolTipText = ""
                End If
            End With

            Dim EXT_LANDED As Decimal = Val(e.Row.Cells("EXT_LANDED").Value & "")
            Dim EXT_LANDED_CALC As Decimal = Val(e.Row.Cells("EXT_LANDED_CALC").Value & "")

            With e.Row.Cells(IIf(optUD.Value = "U", "PO_COST_LANDED", "PO_COST_LANDED"))
                If System.Math.Abs(EXT_LANDED_CALC - EXT_LANDED) > 2 Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText = "Calculated Value of " & Format(EXT_LANDED_CALC, "#,##0.00") & " varies by > $2"
                Else
                    .Appearance.ForeColor = Drawing.Color.Empty
                    .ToolTipText = ""
                End If
            End With
        End If
    End Sub
#End Region

#Region "grdPOTSHIP4"

    Private Sub grdPOTSHIP4_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP4.AfterCellUpdate

    End Sub

    Private Sub grdPOTSHIP4_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTSHIP4.AfterRowUpdate
        Set_Landed_Cost_Needs_to_be_Calculated_Indicator(False)
    End Sub

    Private Sub grdPOTSHIP4_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTSHIP4.BeforeRowUpdate
        If grdPOTSHIP4.Visible = False Then Exit Sub

        If e.Row.Cells("CONTAINER_NO").Value & "" = "" Then
            MsgBox("Container No is Mandatory", vbOKOnly, "Cannot Proceed")
            e.Cancel = True
        End If

        If Not e.Cancel Then
            If Val(e.Row.Cells("PO_SHIPMENT_LNO").Value & "") = 0 Then
                e.Row.Cells("PO_SHIPMENT_LNO").Value = Val(dst.Tables("POTSHIP4").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
                e.Row.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO
                e.Row.Cells("PO_SHIP_STATUS").Value = "O"
                e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            End If
        End If
    End Sub

    Private Sub grdPOTSHIP4_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP4.ClickCellButton

        Select Case e.Cell.Column.Key
            Case "PO_SHIP_STATUS"
                If e.Cell.Value = "O" Then
                    e.Cell.Value = "R"
                ElseIf e.Cell.Value = "R" Then
                    e.Cell.Value = "O"
                End If

            Case "CONTAINER_NO"
                Stop ' SHOW THE CONTAINERS IN THIS SHIPMENT

        End Select

    End Sub

#End Region

#Region "grdPOTSHIP5"

    Private Sub grdPOTSHIP5_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP5.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "COST_CATGY_CODE"
                Dim COST_CATGY_CODE As String = e.Cell.Value & ""
                Dim rowPOTCATG1 As DataRow = LookUp("POTCATG1", COST_CATGY_CODE)
                If rowPOTCATG1 IsNot Nothing Then
                    e.Cell.Row.Cells("COST_CATGY_DESC").Value = rowPOTCATG1.Item("COST_CATGY_DESC")
                    If Val(e.Cell.Row.Cells("LANDING_COST_AMT").Value & "") = 0 Then
                        e.Cell.Row.Cells("LANDING_COST_AMT").Value = rowPOTCATG1.Item("LANDING_COST_AMT")
                    End If
                    'If e.Cell.Row.Cells("LANDING_COST_DIST").Value & "" = "" Then
                    e.Cell.Row.Cells("LANDING_COST_DIST").Value = rowPOTCATG1.Item("LANDING_COST_DIST")
                    'End If
                End If
        End Select
    End Sub

    Private Sub grdPOTSHIP5_AfterEnterEditMode(sender As Object, e As System.EventArgs) Handles grdPOTSHIP5.AfterEnterEditMode

    End Sub

    Private Sub grdPOTSHIP5_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIP5.AfterRowActivate
        With grdPOTSHIP5
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("COST_CATGY_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                For Each COLUMN_NAME As String In New String() {"COST_CATGY_CODE", "LANDING_COST_AMT", "LANDING_COST_DIST", "LANDING_COST_COMMENT"}
                    grdPOTSHIP5.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                Next
                .ActiveCell = grdPOTSHIP5.ActiveRow.Cells("COST_CATGY_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                With grdPOTSHIP5.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() {"COST_CATGY_CODE", "LANDING_COST_AMT", "LANDING_COST_DIST", "LANDING_COST_COMMENT"}
                        If grdPOTSHIP5.ActiveRow.Cells("CTL_NO").Value & "" = "" Then
                            .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                        Else
                            .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                        End If
                    Next
                End With
            End If
        End With
    End Sub

    Private Sub grdPOTSHIP5_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTSHIP5.AfterRowsDeleted
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else
            Calculate_Landed_Cost()
        End If
    End Sub

    Private Sub grdPOTSHIP5_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTSHIP5.AfterRowUpdate
        Set_Landed_Cost_Needs_to_be_Calculated_Indicator(False)
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else
            Calculate_Landed_Cost()
        End If
    End Sub

    Private Sub grdPOTSHIP5_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTSHIP5.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("CTL_NO").Value & "" <> "" Then
                e.Cancel = True
            End If
        Next
    End Sub

    Private Sub grdPOTSHIP5_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTSHIP5.BeforeRowUpdate

        If Not grdPOTSHIP5.Visible Then Exit Sub

        If e.Row.Cells("COST_CATGY_CODE").Value & "" = "" Then
            e.Cancel = True
            Exit Sub
        Else
            Dim rowPOTCATG1 As DataRow = LookUp("POTCATG1", e.Row.Cells("COST_CATGY_CODE").Value)
            If rowPOTCATG1 Is Nothing Then
                e.Cancel = True
                Exit Sub
            End If
        End If

        '  e.Row.Cells("LANDING_COST_DIST").Value = Mid(e.Row.Cells("LANDING_COST_DIST").Value & "", 1, 1)

        If e.Row.Cells("PO_SHIPMENT_NO").Value & "" = "" Then
            e.Row.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO
            e.Row.Cells("PO_SHIPMENT_LNO").Value = Val(dst.Tables("POTSHIP5").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
        End If
    End Sub

    Private Sub grdPOTSHIP5_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP5.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdPOTSHIP5.ActiveCell.Column.Key
            'Case "ACCT_CODE"
            '    sql_where = "NVL(ACCT_STATUS,'X') = 'A' and NVL(ACCT_SUB_CTL,'0') <> '1'"
        End Select

        grdClickCellButton(grdPOTSHIP5, sql_where, False)
    End Sub

    Private Sub grdPOTSHIP5_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTSHIP5.InitializeRow
        If e.Row.Cells("CTL_NO").Value & "" <> "" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGray
        End If
    End Sub
#End Region

#Region "grdWHTPREC3"
    Private Sub grdWHTPREC3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTPREC3.AfterRowUpdate
        With e.Row
            Dim Sql = "BEGIN " & vbCrLf _
            & " INSERT INTO WHTPREC3 (PO_SHIPMENT_NO, PO_SHIPMENT_LNO, PO_ORDER_NO, PO_ORDER_LNO, PO_QTY_SHP, PO_REC_NOTE, LOCATION_CODE)" & vbCrLf _
            & " VALUES ('" & .Cells("PO_SHIPMENT_NO").Value & "', " & .Cells("PO_SHIPMENT_LNO").Value & ",'" & .Cells("PO_ORDER_NO").Value & "'," & .Cells("PO_ORDER_LNO").Value & "," & .Cells("PO_QTY_SHP").Value & ",'" & .Cells("PO_REC_NOTE").Value & "','" & .Cells("LOCATION_CODE").Value & "'); " & vbCrLf _
            & " Exception" & vbCrLf _
            & "  WHEN DUP_VAL_ON_INDEX THEN" & vbCrLf _
            & "     Update WHTPREC3" & vbCrLf _
            & "     SET    PO_REC_NOTE = '" & .Cells("PO_REC_NOTE").Value & "'" & vbCrLf _
            & "     ,LOCATION_CODE = '" & .Cells("LOCATION_CODE").Value & "'" & vbCrLf _
            & "     WHERE PO_SHIPMENT_NO = '" & .Cells("PO_SHIPMENT_NO").Value & "'" & vbCrLf _
            & "     and PO_SHIPMENT_LNO =  " & .Cells("PO_SHIPMENT_LNO").Value & "" & vbCrLf _
            & "     and PO_ORDER_NO = '" & .Cells("PO_ORDER_NO").Value & "'" & vbCrLf _
            & "     and PO_ORDER_LNO = " & .Cells("PO_ORDER_LNO").Value & ";" & vbCrLf _
            & " End;"
            ASCDATA1.ExecuteSQL(Sql)
        End With

    End Sub

    Private Sub grdWHTPREC3_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTPREC3.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

                Case "LOCATION_CODE"
                    Dim sql_where As String = "WHSE_CODE = '" & WHSE_CODE & "' and nvl(LOCATION_USE, 'A') = 'A'"
                    grdClickCellButton(grdWHTPREC3, sql_where)


            End Select
        End With

    End Sub

#End Region

    Private Sub optUD_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optUD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Toggle_UD()

    End Sub

    Sub Toggle_UD()
        Dim tf As Boolean = (optUD.Value = "D")

        Dim TF_VAN As Boolean = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")

        With grdPOTSHIP3.DisplayLayout.Bands(0)
            If Not cost_calc Then
                .Columns("PO_QTY_SHP").Hidden = tf
                .Columns("PO_QTY_SHP_DZ").Hidden = Not tf
                .Columns("NET_OPEN").Hidden = tf
                .Columns("NET_OPEN_DZ").Hidden = Not tf
                .Columns("PO_QTY_REC").Hidden = tf Or (InquiryMode And ship_entry And (EntryMode <> "E"))
                .Columns("PO_QTY_REC_DZ").Hidden = Not tf Or (ship_entry And (InquiryMode And EntryMode <> "E"))
                .Columns("PO_QTY_VAR").Hidden = tf Or (Not receipt_mode And Not InquiryMode And Not (ship_entry Or (EntryMode = "E")))
            End If

            If cost_calc Then
                .Columns("PO_QTY_REC").Hidden = tf
                .Columns("PO_QTY_SHP").Hidden = tf
                .Columns("PO_COST_VCOST_UM").Hidden = tf
                .Columns("PO_COST_MATLS_UM").Hidden = tf Or Not TF_VAN
                .Columns("PO_COST_OTHER").Hidden = tf
                .Columns("PO_COST_QUOTA").Hidden = tf Or Not TF_VAN
                .Columns("PO_COST_QUOTA_DF").Hidden = tf Or Not TF_VAN
                .Columns("FIRST_COST_TOTAL").Hidden = tf
                .Columns("COMMISSION_COST").Hidden = tf Or Not TF_VAN

                .Columns("PO_QTY_REC_DZ").Hidden = Not tf
                .Columns("PO_QTY_SHP_DZ").Hidden = Not tf
                .Columns("PO_COST_VCOST_DZ").Hidden = Not tf
                .Columns("PO_COST_MATLS_DZ").Hidden = Not tf Or Not TF_VAN
                .Columns("PO_COST_OTHER_DZ").Hidden = Not tf
                .Columns("PO_COST_QUOTA_DZ").Hidden = Not tf Or Not TF_VAN
                .Columns("PO_COST_QUOTA_DF_DZ").Hidden = Not tf Or Not TF_VAN
                .Columns("FIRST_COST_TOTAL_DZ").Hidden = Not tf
                .Columns("COMMISSION_COST_DZ").Hidden = Not tf Or Not TF_VAN
            End If


        End With
    End Sub

    Sub Update_ICTSTAT2(STYLE_CODE As String, COLOR_CODE As String, WHSE_CODE As String, COLUMN_NAME As String, QTY As Integer, Optional OPS_YYYYPP As String = "")
        ASCMAIN1.sql = "" _
                   & "Begin" _
                   & " Update ICTSTAT2 Set COLUMN_NAME = NVL(COLUMN_NAME,0) + " & CStr(QTY) _
                   & "  where STYLE_CODE = '" & STYLE_CODE & "'" _
                   & "    and COLOR_CODE = '" & COLOR_CODE & "'" _
                   & "    and WHSE_CODE = '" & WHSE_CODE & "';" _
                   & " If SQL%NOTFOUND Then " _
                   & "  Insert into ICTSTAT2 (STYLE_CODE,COLOR_CODE,WHSE_CODE,COLUMN_NAME)" _
                   & "   Values ('" & STYLE_CODE & "','" & COLOR_CODE & "','" & WHSE_CODE & "'," & CStr(QTY) & ");" _
                   & " End If;" _
                   & "End;"
        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "COLUMN_NAME", COLUMN_NAME)
        ASCDATA1.ExecuteSQL()

        If COLUMN_NAME = "WHSE_QTY_ON_HAND" Then
            ASCMAIN1.sql = "" _
               & "Begin" _
               & " Update ICTSTAT1 Set WHSE_QTY_REC = NVL(WHSE_QTY_REC,0) + " & CStr(QTY) _
               & "  where STYLE_CODE = '" & STYLE_CODE & "'" _
               & "    and COLOR_CODE = '" & COLOR_CODE & "'" _
               & "    and WHSE_CODE = '" & WHSE_CODE & "'" _
               & "    and OPS_YYYYPP = '" & OPS_YYYYPP & "';" _
               & " If SQL%NOTFOUND Then " _
               & "  Insert into ICTSTAT1 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_REC)" _
               & "   Values ('" & STYLE_CODE & "','" & COLOR_CODE & "','" & WHSE_CODE & "','" & OPS_YYYYPP & "'," & CStr(QTY) & ");" _
               & " End If;" _
               & "End;"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Sub Set_Landed_Cost_Needs_to_be_Calculated_Indicator(tf As Boolean)
        If cost_calc Then
            cost_ind = tf
            If Not cost_ind Then
                UltraExplorerBar1.Groups("Screen Control").Items("Calculate Costs").Settings.AppearancesSmall.Appearance.ForeColor = Drawing.Color.Red
            Else
                UltraExplorerBar1.Groups("Screen Control").Items("Calculate Costs").Settings.AppearancesSmall.Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If

        Calculate_DUTY_DIST()

    End Sub

    Sub Create_POTSHIPR(PO_SHIPMENT_LNO As Int64, STYLE_CODE As String, COLOR_CODE As String)
        If dst.Tables("POTSHIPR").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE}) Is Nothing Then
            Dim rowPOTSHIPR As DataRow = dst.Tables("POTSHIPR").NewRow
            rowPOTSHIPR.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIPR.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            rowPOTSHIPR.Item("STYLE_CODE") = STYLE_CODE
            rowPOTSHIPR.Item("COLOR_CODE") = COLOR_CODE
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            rowPOTSHIPR.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC") & ""
            dst.Tables("POTSHIPR").Rows.Add(rowPOTSHIPR)
        End If
    End Sub

    Private Sub grdPOTSHIP8_AfterCellActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIP8.AfterCellActivate

    End Sub

    Private Sub grdPOTSHIP8_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP8.AfterCellUpdate
        'If e.Cell.Column.Key = "DOZENS" Then
        '    e.Cell.Row.Update()
        'End If
        e.Cell.Row.Update()
    End Sub

    Private Sub grdPOTSHIP8_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTSHIP8.InitializeLayout

    End Sub

    Private Sub grdPOTSHIP7_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTSHIP7.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "PPK_CODE"
                Check_for_PPK(e.Cell.Row)

            Case "CARTON_DIMS"
                Dim CARTON_DIMS As String = e.Cell.Value & ""
                If CARTON_DIMS <> "" Then
                    e.Cell.Row.Cells("CARTON_VOLUME").Value = Get_Volume_from_Dims(CARTON_DIMS)
                End If
        End Select
    End Sub

    Private Sub grdPOTSHIP7_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIP7.AfterRowActivate
        Setup_grdPOTSHIP8()
    End Sub

    Sub Check_for_PPK(grow As UltraWinGrid.UltraGridRow)
        If Checking_for_PPK Then Exit Sub
        Checking_for_PPK = True
        If Val(grow.Cells("STYLES").Value & "") > 1 _
        Or grow.Cells("CUSTOM_PPK").Value & "" = "1" Then
            ' Or grow.Cells("CARTON_COMMENTS").Value & "" <> ""
            grow.Cells("PPK_CODE").Value = Get_Next_PPK_CODE()
        Else
            If grow.Cells("PPK_CODE").Value & "" <> "" Then
                grow.Cells("PPK_CODE").Value = ""
            End If
        End If
        If grow.DataChanged Then
            grow.Update()
        End If
        Checking_for_PPK = False
    End Sub

    Function Get_Next_PPK_CODE() As String
        PPK_CODE_ctr += 1
        Return "TMP" & Format(PPK_CODE_ctr, "0000000")
    End Function

    Sub Setup_grdPOTSHIP8()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show All Carton Details"), UltraWinToolbars.StateButtonTool)
        grdPOTSHIP8.DisplayLayout.Bands(0).Columns("CARTON_NO").Hidden = Not tlb_sbt.Checked

        If grdPOTSHIP7.ActiveRow Is Nothing OrElse Not grdPOTSHIP7.ActiveRow.IsDataRow Then
            grdPOTSHIP8.Visible = False
        Else

            Dim dvw As DataView = DirectCast(grdPOTSHIP8.DataSource, DataTable).DefaultView
            Dim CARTON_NO As Integer = Val(grdPOTSHIP7.ActiveRow.Cells("CARTON_NO").Value & "")
            Dim PO_SHIPMENT_LNO As Integer = Val(grdPOTSHIP7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            If tlb_sbt.Checked Then
                dvw.RowFilter = ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO))
                grdPOTSHIP8.Text = "Carton Type " & CStr(CARTON_NO) & " by Style/Color for All Carton Types"
            Else
                dvw.RowFilter = ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO))
                grdPOTSHIP8.Text = "Carton Type " & CStr(CARTON_NO) & " by Style/Color"
            End If
            grdPOTSHIP8.Visible = True
        End If
    End Sub

    Private Sub tabBOL_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabBOL.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabBOL()
    End Sub

    Sub Setup_tabBOL()
        If ship_entry Then
            grdPOTSHIP4.Visible = (tabBOL.SelectedTab.Key = "PO Details")
            'grdPOTSHIPR.Visible = (tabBOL.SelectedTab.Key = "Cartons")
            'grdPOTSHIPQ.Visible = (tabBOL.SelectedTab.Key = "Cartons")
            splCartonQ.Visible = (tabBOL.SelectedTab.Key = "Cartons")
            cmdCreate.Visible = (tabBOL.SelectedTab.Key = "Cartons")

            grdPOTSHIP2.DisplayLayout.Bands(0).Columns("ACTION").Hidden = Not (tabBOL.SelectedTab.Key = "PO Details") Or packingFromXLS
            UltraExplorerBar1.Groups("Options").Visible = (tabBOL.SelectedTab.Key = "PO Details")

            If tabBOL.SelectedTab.Key = "Cartons" Then
                splBOL.SplitterDistance = splPOTSHIP7.SplitterDistance
            Else
                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    splBOL.SplitterDistance = splBOL.Width * 0.7
                Else
                    splBOL.SplitterDistance = splBOL.Width * 0.6
                End If
            End If

            If tabBOL.SelectedTab.Key = "Packing Discrepancies" And ScreenMode Then
                Packing_Discrepancies()
            End If
        Else
            splBOL.Panel2Collapsed = (tabBOL.SelectedTab.Key <> "Cartons")
            UltraExplorerBar1.Groups("Options").Visible = False

        End If

        With grdPOTSHIP2.DisplayLayout.Bands(0)
            .Columns("PO_QTY_SHP").Hidden = Not receipt_mode And Not InquiryMode And Not cost_calc
            .Columns("PO_QTY_REC").Hidden = Not receipt_mode And Not InquiryMode And Not cost_calc
            .Columns("PO_QTY_VAR").Hidden = Not receipt_mode And Not InquiryMode And Not cost_calc
            'For Each COLUMN_NAME As String In New String() _
            '    {"TOTAL_WEIGHT", "CBM_RATE", "CBM", "BOL_FEE", "TRUCKING"}
            '    .Columns(COLUMN_NAME).Hidden = (tabBOL.SelectedTab.Key = "Cartons") Or ship_entry Or receipt_mode
            'Next
            .Columns("CLOSE").Hidden = True

            For Each COLUMN_NAME As String In New String() _
                {"INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE"}
                .Columns(COLUMN_NAME).Hidden = (tabBOL.SelectedTab.Key = "Cartons") Or Not (EntryMode = "V") Or receipt_mode
            Next
            For Each COLUMN_NAME As String In New String() _
                {"TRAN_NO", "OPS_YYYYPP", "PO_SOURCE_DOC", "PO_DATE_RECEIVED"}
                .Columns(COLUMN_NAME).Hidden = (tabBOL.SelectedTab.Key = "Cartons") Or Not (cost_calc Or EntryMode = "V") Or receipt_mode
            Next
        End With

        With grdPOTSHIP4.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"TOTAL_WEIGHT"}
                .Columns(COLUMN_NAME).Hidden = ship_entry
            Next
            For Each COLUMN_NAME As String In New String() _
                {"INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE"}
                .Columns(COLUMN_NAME).Hidden = Not (EntryMode = "V")
            Next
            'For Each COLUMN_NAME As String In New String() _
            '    {"TRAN_NO", "OPS_YYYYPP", "PO_SOURCE_DOC", "PO_DATE_RECEIVED"}
            '    .Columns(COLUMN_NAME).Hidden = Not (tabBOL.SelectedTab.Key = "PO Details") Or Not (EntryMode = "V" Or (EntryMode = "E" And receipt_mode))
            'Next
            .Columns("PO_SHIP_STATUS").Hidden = True
        End With

        SETUP_tabPOTSHIP1()
        Setup_grdPOTSHIP2_ActiveRow()

        Dim bln As Boolean = ScreenMode And Not ship_entry And tabBOL.SelectedTab IsNot Nothing
        With UltraExplorerBar1
            .Groups("BOL Data").Visible = bln And Not receipt_mode AndAlso (tabBOL.SelectedTab.Key = "PO Details" Or tabBOL.SelectedTab.Key = "Container Summary")
            .Groups("Customs/Duty").Visible = bln And Not receipt_mode AndAlso (tabBOL.SelectedTab.Key = "Customs / Other")

            .Groups("Cost Options").Visible = bln And Not receipt_mode AndAlso (tabBOL.SelectedTab.Key = "PO Details") And Not InquiryMode And cost_calc And EntryMode = "E"
        End With

        UltraExplorerBar1.Groups("Print Labels").Visible = (tabBOL.SelectedTab.Key = "Cartons")
    End Sub

    Sub Packing_Discrepancies()
        EnforceConstraints(False)
        dst.Tables("POTPACKR").Rows.Clear()
        dst.Tables("POTPACKD").Rows.Clear()
        EnforceConstraints(True)
        For Each rowPOTSHIPR As DataRow In dst.Tables("POTSHIPR").Select("QTY_VAR <> 0")
            Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").Rows.Find(New Object() {rowPOTSHIPR.Item("PO_SHIPMENT_NO"), rowPOTSHIPR.Item("PO_SHIPMENT_LNO")})
            Dim rowPOTPACKR As DataRow = dst.Tables("POTPACKR").NewRow
            With rowPOTPACKR
                For Each C As String In New String() _
                    {"PO_SHIPMENT_NO", "PO_SHIPMENT_LNO", "STYLE_CODE", "COLOR_CODE",
                     "QTY_SHP", "QTY_CTN", "COLOR_DESC"}
                    .Item(C) = rowPOTSHIPR.Item(C)
                Next
                .Item("CONTAINER_NO") = rowPOTSHIP2.Item("CONTAINER_NO")
                .Item("COMM_INV_NO") = rowPOTSHIP2.Item("COMM_INV_NO")
                Dim QTY_PACKED_KEY As String = rowPOTSHIPR.Item("PO_SHIPMENT_LNO") & ":" & rowPOTSHIPR.Item("STYLE_CODE") & ":" & rowPOTSHIPR.Item("COLOR_CODE")
                If QTY_PACKED.ContainsKey(QTY_PACKED_KEY) Then
                    .Item("QTY_PCK") = QTY_PACKED(QTY_PACKED_KEY)
                End If

            End With
            dst.Tables("POTPACKR").Rows.Add(rowPOTPACKR)
            '            If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE.StartsWith("ME20294") Or PO_REFERENCE = "ME20295") Then Stop
            Dim sqlw As String = "PO_SHIPMENT_NO = '" & rowPOTSHIPR.Item("PO_SHIPMENT_NO") & "'" _
                                 & " AND PO_SHIPMENT_LNO = " & rowPOTSHIPR.Item("PO_SHIPMENT_LNO") _
                                 & " AND STYLE_CODE = '" & rowPOTSHIPR.Item("STYLE_CODE") & "'" _
                                 & " AND COLOR_CODE = '" & rowPOTSHIPR.Item("COLOR_CODE") & "'"

            Dim QTY_PCK As Integer = Val(rowPOTPACKR.Item("QTY_PCK") & "")
            Dim QTY_CTN As Integer = Val(rowPOTPACKR.Item("QTY_CTN") & "")
            Dim QTY_SHP As Integer = Val(rowPOTPACKR.Item("QTY_SHP") & "")
            Dim QTY_SHP_TOTAL As Integer = 0
            Dim blnDone As Boolean = False

            For Each row As DataRow In dst.Tables("POTSHIP3").Select(sqlw)

                If blnDone Then
                    'rowPOTSHIPR.Item("QTY_SHP") = Val(rowPOTSHIPR.Item("QTY_SHP") & "") - Val(row.Item("PO_QTY_SHP") & "")
                    row.Item("PO_QTY_SHP") = 0

                Else
                    Dim PO_QTY_SHP As Integer = Val(row.Item("PO_QTY_SHP") & "")
                    QTY_SHP_TOTAL += PO_QTY_SHP
                    'If QTY_SHP_TOTAL = QTY_PCK Then
                    '    blnDone = True
                    'End If
                    'If QTY_SHP_TOTAL = QTY_CTN Then
                    '    blnDone = True
                    'End If

                    Dim rowPOTPACKD As DataRow = dst.Tables("POTPACKD").NewRow
                    With rowPOTPACKD
                        For Each C As String In New String() _
                            {"PO_SHIPMENT_NO", "PO_SHIPMENT_LNO", "PO_ORDER_NO", "PO_ORDER_LNO",
                             "STYLE_CODE", "COLOR_CODE", "PO_QTY_OPN", "PO_QTY_SHP",
                             "PO_REFERENCE", "PO_DATE_SHIP_BY"}
                            .Item(C) = row.Item(C)
                        Next
                    End With

                    dst.Tables("POTPACKD").Rows.Add(rowPOTPACKD)
                End If
            Next

            'If QTY_SHP_TOTAL = QTY_PCK Then
            '    rowPOTPACKR.Delete()
            'End If
            'If QTY_SHP_TOTAL = QTY_CTN Then
            '    rowPOTPACKR.Delete()
            'End If
        Next

        grdPOTPACKR.Rows.ExpandAll(True)

    End Sub

    Private Sub grdPOTSHIP7_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTSHIP7.InitializeRow
        If e.Row.Cells("PPK_CODE").Value & "" <> "" Then
            e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.LightGreen
        Else
            e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdPOTSHIPR_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTSHIPR.InitializeRow
        If e.Row.Band.Key = "POTSHIPR" Then
            If Val(e.Row.Cells("QTY_VAR").Value & "") <> 0 Then
                e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Sub Setup_Warehouse_Attributes()
        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            LP_CODE = ""
            WHSE_LOCATOR = ""
            WHSE_CTN_CTL = ""
        Else
            LP_CODE = rowICTWHSE1.Item("LP_CODE") & ""
            WHSE_LOCATOR = rowICTWHSE1.Item("WHSE_LOCATOR") & ""
            WHSE_CTN_CTL = rowICTWHSE1.Item("WHSE_CTN_CTL") & ""
        End If

        Toggle_chkFinalize()
    End Sub

    Sub Toggle_chkFinalize()
        If LP_CODE <> "" Or (WHSE_LOCATOR = "1" And WHSE_CTN_CTL = "C") Then
            'chkFinalize.Enabled = True
            chkFinalize.Visible = True
            If LP_CODE <> "" Then
                chkFinalize.Text = "Send to 3PL"
            Else
                chkFinalize.Text = "Rel to Whse"
            End If
        Else
            chkFinalize.Checked = False
            'chkFinalize.Enabled = False
            chkFinalize.Visible = False
        End If
    End Sub

    Sub Release_Shipment_Send_3PL()

        Dim LP_XNO As String = TAC.WHCMAIN1.Get_LP_XNO(MENU_ITEM_OBJECT, 1)

        If LP_CODE <> "" Then ' IE, IF THIS WAREHOUSE IS A 3PL

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then

            ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then

                If Not dst.Tables.Contains("EDT943O1") Then
                    Create_TDA(dst.Tables.Add, "EDT943O1", "*")
                    Create_TDA(dst.Tables.Add, "EDT943O2", "*")
                End If

                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowPOTSHIP1.Item("WHSE_CODE"))
                Dim rowEDTTRPM1 As DataRow = LookUp("EDTTRPM1",
                                                    New String() {rowICTWHSE1.Item("WHSE_EDI_QUAL"), rowICTWHSE1.Item("WHSE_EDI_ID"), "943"})

                ' BeginTrans()
                ASCMAIN1.sql = "Select Distinct PO_SHIPMENT_NO, CONTAINER_NO from POTSHIP2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIP_STATUS = 'O'"
                For Each rowPOTSHIPX As DataRow In ASCDATA1.GetDataTable.Select("", "CONTAINER_NO")
                    Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                    Dim CONTAINER_NO As String = rowPOTSHIPX.Item("CONTAINER_NO")

                    ASCMAIN1.sql = "Select * from POTSHIP4 where PO_SHIPMENT_NO = :PARM1 and CONTAINER_NO = :PARM2"
                    Dim rowPOTSHIP4 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {PO_SHIPMENT_NO, CONTAINER_NO})

                    Dim rowEDT943O1 As DataRow = dst.Tables("EDT943O1").NewRow
                    With rowEDT943O1
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        .Item("EDI_REPORTING_CODE") = "J"
                        .Item("EDI_DEPOSITOR_ORDER_NO") = ""
                        .Item("EDI_SHIPMENT_DATE") = rowPOTSHIP1.Item("PO_DATE_SHIPPED")
                        ' PO_SHIPMENT_NO & "-" & FORMAT(PO_SHIPMENT_LNO,"000") & "-" & CONTAINER_NO
                        ' 123456-001-CNFNDNDWWDD22
                        ' 123456-002-CNFNDNDWWDD22
                        ' 123477-CNFNDNDWWDD22
                        .Item("EDI_PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        .Item("EDI_NAME") = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME")
                        .Item("EDI_WH_ID_CODE") = ""
                        .Item("EDI_DIVISION_CODE") = rowICTWHSE1.Item("LP_WHSE_ID")
                        .Item("EDI_ARRIVAL_DATE") = rowPOTSHIP1.Item("PO_SHIP_ETA")
                        .Item("EDI_CARRIER_SCAC") = ""
                        .Item("EDI_PALLET_QTY") = 0
                        .Item("EDI_SEAL_NUMBER_CONTAINER") = CONTAINER_NO
                        If rowPOTSHIP4 Is Nothing Then
                            .Item("EDI_SEAL_NUMBER") = ""
                        Else
                            .Item("EDI_SEAL_NUMBER") = rowPOTSHIP4.Item("CONTAINER_SEAL_NO")
                        End If
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    End With
                    dst.Tables("EDT943O1").Rows.Add(rowEDT943O1)

                    ASCMAIN1.sql = "Select POTSHIP3.*,POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE" & vbCrLf _
                        & ",ICTSTYL1.STYLE_DESC,ICTSTYL1.CUST_CODE,ICTSTYL1.STYLE_GROUP_CODE" & vbCrLf _
                        & ",ICTSTYL1.CARTON_PACK_QTY,ICTSTYL1.INNER_PACK_QTY,ICTSTYL1.CASE_WEIGHT_GRS" & vbCrLf _
                        & ",ICTSTYC1.HIDE_COLOR_3PL,ICTSTYC1.UPC_CODE" & vbCrLf _
                        & " from POTSHIP3,POTSHIP2,POTORDR2,ICTSTYL1,ICTSTYC1" & vbCrLf _
                        & " where POTSHIP2.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                        & "   and POTSHIP2.CONTAINER_NO = '" & CONTAINER_NO & "'" & vbCrLf _
                        & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.COLOR_CODE = POTORDR2.COLOR_CODE" & vbCrLf _
                        & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                        & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                        & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                        & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO"

                    Dim lead_item_processed As Boolean = False

                    Dim EDI_DOC_LNO As Int64 = 0
                    For Each rowPOTSHIPD As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_CODE,COLOR_CODE")

                        Dim STYLE_CODE As String = rowPOTSHIPD.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowPOTSHIPD.Item("COLOR_CODE")
                        Dim HIDE_COLOR_3PL As String = rowPOTSHIPD.Item("HIDE_COLOR_3PL") & ""
                        Dim PO_SHIPMENT_LNO As Int32 = Val(rowPOTSHIPD.Item("PO_SHIPMENT_LNO") & "")
                        Dim CUST_CODE As String = rowPOTSHIPD.Item("CUST_CODE") & ""
                        Dim STYLE_GROUP_CODE As String = rowPOTSHIPD.Item("STYLE_GROUP_CODE") & ""

                        If Not lead_item_processed Then
                            Dim EDI_DIVISION_CODE As String = rowEDT943O1.Item("EDI_DIVISION_CODE")
                            If CUST_CODE = "DOLGEN" Then
                                EDI_DIVISION_CODE = "NYDG"
                            ElseIf CUST_CODE = "WALMART" Then
                                If STYLE_GROUP_CODE = "07" Then
                                    EDI_DIVISION_CODE = "NYWB"
                                Else
                                    EDI_DIVISION_CODE = "NYWM"
                                End If
                            End If
                            rowEDT943O1.Item("EDI_DIVISION_CODE") = EDI_DIVISION_CODE
                            lead_item_processed = True
                        End If


                        ASCMAIN1.sql = "Select POTSHIP7.* from POTSHIP7,POTSHIP8" & vbCrLf _
                            & " where POTSHIP8.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                            & "   and POTSHIP8.PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & vbCrLf _
                            & "   and POTSHIP8.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                            & "   and POTSHIP8.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                            & "   and POTSHIP7.PO_SHIPMENT_NO = POTSHIP8.PO_SHIPMENT_NO" & vbCrLf _
                            & "   and POTSHIP7.PO_SHIPMENT_LNO = POTSHIP8.PO_SHIPMENT_LNO" & vbCrLf _
                            & "   and POTSHIP7.CARTON_NO = POTSHIP8.CARTON_NO"
                        Dim rowPOTSHIP7() As DataRow = ASCDATA1.GetDataTable.Select("")
                        Dim CARTON_DIMS As String = ""
                        Dim CARTON_VOLUME As Decimal = 0
                        If rowPOTSHIP7 IsNot Nothing AndAlso rowPOTSHIP7.Length > 0 Then
                            CARTON_DIMS = rowPOTSHIP7(0).Item("CARTON_DIMS") & ""
                            CARTON_VOLUME = Val(rowPOTSHIP7(0).Item("CARTON_VOLUME") & "")
                        End If

                        Dim CUST_STYLE_CODE As String = ""
                        If CUST_CODE <> "" Then
                            ' we may add CUST_STYLE_CODE in the PO in the future
                            ASCMAIN1.sql = "Select Min (CUST_STYLE_CODE) from SOTCSTY1" _
                                & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                                & "   and SOTCSTY1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                                & "   and SOTCSTY1.COLOR_CODE = '" & COLOR_CODE & "'"
                            CUST_STYLE_CODE = ASCDATA1.GetDataValue
                        End If

                        Dim rowEDT943O2 As DataRow = dst.Tables("EDT943O2").NewRow
                        With rowEDT943O2
                            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            EDI_DOC_LNO += 1
                            .Item("EDI_DOC_LNO") = EDI_DOC_LNO
                            .Item("EDI_UNITS_SHIPPED") = rowPOTSHIPD.Item("PO_QTY_SHP")
                            .Item("EDI_UPC_CASE_CODE") = rowPOTSHIPD.Item("UPC_CODE")
                            .Item("EDI_STYLE_NO") = IIf(HIDE_COLOR_3PL = "1", STYLE_CODE, STYLE_CODE & "" & COLOR_CODE)
                            .Item("EDI_CUST_STYLE_NO") = CUST_STYLE_CODE
                            .Item("EDI_ITEM_DESC") = rowPOTSHIPD.Item("STYLE_DESC")

                            Dim EDI_SUB_INNER_QTY As Int64 = 0
                            Dim CARTON_PACK_QTY As Int64 = Val(rowPOTSHIPD.Item("CARTON_PACK_QTY") & "")
                            Dim INNER_PACK_QTY As Int64 = Val(rowPOTSHIPD.Item("INNER_PACK_QTY") & "")
                            If INNER_PACK_QTY <> 0 And CARTON_PACK_QTY <> 0 AndAlso CARTON_PACK_QTY Mod INNER_PACK_QTY = 0 Then
                                EDI_SUB_INNER_QTY = CARTON_PACK_QTY / INNER_PACK_QTY
                            End If
                            .Item("EDI_SUB_INNER_QTY") = EDI_SUB_INNER_QTY

                            .Item("EDI_PACK_QTY") = CARTON_PACK_QTY
                            '.Item("EDI_SIZE") = CARTON_VOLUME
                            '.Item("EDI_WEIGHT") = rowPOTSHIPD.Item("CASE_WEIGHT_GRS")
                            .Item("EDI_PO_ORDER_NO") = rowPOTSHIPD.Item("PO_ORDER_NO")
                        End With
                        dst.Tables("EDT943O2").Rows.Add(rowEDT943O2)
                    Next

                    '4012453780      TAYLORED

                    ASCMAIN1.sql = "Insert into EDTSYSIH (COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_APPLICATION_ID,EDI_PROCESS_IND," _
                        & "EDI_OUR_ID,EDI_TP_ID,INIT_DATE,INIT_OPER)" _
                        & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,'" & ASCMAIN1.USER_ID & "')"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVVV",
                            New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, "AR", "1",
                                          rowEDTTRPM1.Item("EDI_OUR_ID"), rowICTWHSE1.Item("WHSE_EDI_ID")})
                Next
                Update_Record_TDA("EDT943O1")
                Update_Record_TDA("EDT943O2")

                ' CommitTrans()
            End If
        End If

        ASCMAIN1.sql = "Update POTSHIP1 Set LP_STATUS = '1',LP_XNO = '" & LP_XNO & "'" _
                & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
        ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub grdPOTSHIPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTSHIPX.InitializeRow

        If cost_calc Then Exit Sub

        If e.Row.IsDataRow Then
            If e.Row.Cells("LP_CODE").Value & "" <> "" Then
                Dim PO_SHIP_ETA As Date = e.Row.Cells("PO_SHIP_ETA").Value
                If e.Row.Cells("LP_CODE").Value & "" <> "" And e.Row.Cells("LP_STATUS").Value & "" = "1" Then
                    e.Row.Cells("WHSE_CODE").Appearance.ForeColor = Drawing.Color.Green
                    e.Row.Cells("WHSE_CODE").ToolTipText = "LP Provider has been sent this PO Shipment"
                Else
                    If Format(PO_SHIP_ETA, "yyyyMMdd") < Format(Now.Date.AddDays(10), "yyyyMMdd") Then
                        e.Row.Cells("WHSE_CODE").Appearance.BackColor = Drawing.Color.Red
                        e.Row.Cells("WHSE_CODE").ToolTipText = "LP Provider has not been sent this PO Shipment"
                    End If
                End If

                If Format(PO_SHIP_ETA, "yyyyMMdd") < Format(Now.Date.AddDays(0), "yyyyMMdd") Then
                    e.Row.Cells("PO_SHIP_ETA").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("PO_SHIP_ETA").ToolTipText = "ETA Date has arrived - Receipt past due"
                ElseIf Format(PO_SHIP_ETA, "yyyyMMdd") < Format(Now.Date.AddDays(10), "yyyyMMdd") Then
                    e.Row.Cells("PO_SHIP_ETA").Appearance.BackColor = Drawing.Color.Yellow
                    e.Row.Cells("PO_SHIP_ETA").ToolTipText = "ETA Date is within 10 days"
                End If
            End If
        End If
    End Sub

    Sub Toggle_Columns()
        Dim blnFREIGHT_ENTERED_BY_Container As Boolean = (Absx1.optFor("FREIGHT_ENTERED_BY").Value = "C")
        Dim blnFREIGHT_ENTERED_BY_Invoice As Boolean = (Absx1.optFor("FREIGHT_ENTERED_BY").Value = "I")
        With grdPOTSHIP2.DisplayLayout.Bands(0)
            .Columns("CBM_RATE").Hidden = blnFREIGHT_ENTERED_BY_Container Or blnFREIGHT_ENTERED_BY_Invoice Or ship_entry Or receipt_mode
            .Columns("CBM").Hidden = blnFREIGHT_ENTERED_BY_Container Or blnFREIGHT_ENTERED_BY_Invoice Or ship_entry Or receipt_mode
            .Columns("BOL_FEE").Hidden = blnFREIGHT_ENTERED_BY_Container Or blnFREIGHT_ENTERED_BY_Invoice Or ship_entry Or receipt_mode
            .Columns("TRUCKING").Hidden = blnFREIGHT_ENTERED_BY_Container Or blnFREIGHT_ENTERED_BY_Invoice Or ship_entry Or receipt_mode
            .Columns("ACTION").Hidden = Not receipt_mode And Not ((EntryMode = "N" Or EntryMode = "E") And (ship_entry Or receipt_mode))

            If receipt_mode Then
                If select_from_3PL_list Or Select_from_Whse_Receipt Then
                    .Columns("ACTION").Hidden = True
                End If
                '.Columns("ACTION").Hidden = select_from_3PL_list
            End If

            For Each COLUMN_NAME As String In New String() {"LINES", "LINES_EXACT", "LINES_OVER", "LINES_SHORT", "LINES_ZERO"}
                .Columns(COLUMN_NAME).Hidden = ship_entry ' Not receipt_mode And Not InquiryMode And (Not ship_entry And (EntryMode = "N" Or EntryMode = "E"))
            Next

        End With
        With grdPOTSHIP4.DisplayLayout.Bands(0)
            .Columns("CBM").Hidden = Not blnFREIGHT_ENTERED_BY_Container Or blnFREIGHT_ENTERED_BY_Invoice Or ship_entry
            .Columns("FREIGHT_AMT").Hidden = Not blnFREIGHT_ENTERED_BY_Container Or blnFREIGHT_ENTERED_BY_Invoice Or ship_entry
            .Columns("TRUCKING").Hidden = Not blnFREIGHT_ENTERED_BY_Container Or blnFREIGHT_ENTERED_BY_Invoice Or ship_entry
        End With
    End Sub

    Sub Get_Receipt_Data_from_3PL()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Checking 3PL for Receipts Data")

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

        ElseIf ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Initialize_EDI()
            Fill_Records("EDT944T1")
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdWHT3PLR1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHT3PLR1.AfterRowActivate
        If ScreenMode Then
            Setup_WHT3PLR1()
        End If
    End Sub

    Private Sub grdWHT3PLR1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHT3PLR1.DoubleClickRow

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

        ElseIf ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then

        End If

    End Sub

    Function Load_3PL_Receipt_Details(TRANS_SEQ As String, LP_CODE As String, WHSE_CODE As String, ARRDTE As Date, REF1 As String) As Boolean

        ASCMAIN1.sql = "Select POTSHIP3.*,POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE, '0' RECEIVED" & vbCrLf _
            & " from POTSHIP3,POTSHIP2,POTORDR2" & vbCrLf _
            & " where POTSHIP3.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO"

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Dim rowEDT944T1 As DataRow = LookUp("EDT944T1", EDI_DOC_SEQ_NO) '  dst.Tables("EDT944T1").Rows.Find(EDI_DOC_SEQ_NO)

            If rowEDT944T1.Item("EDI_PROCESS_IND") & "" <> "0" Then
                MsgBox("3PL Receipt No Longer Available for Receipt", MsgBoxStyle.OkOnly, "Need to Refresh Grid")
                Return False
            End If

            Dim EDI_SEAL_NUMBER_CONTAINER As String = rowEDT944T1.Item("EDI_SEAL_NUMBER_CONTAINER") & ""
            ASCMAIN1.sql &= vbCrLf & " and REPLACE(REPLACE(REPLACE(POTSHIP2.CONTAINER_NO,'-',''),'/',''),':','') = '" & EDI_SEAL_NUMBER_CONTAINER & "'"
        End If

        Dim POTSHIP3 As String = ASCMAIN1.Temp_Table()

        ASCMAIN1.sql = "Select Sum (PO_QTY_REC) from " & POTSHIP3
        Dim PO_QTY_REC As Int64 = Val(ASCDATA1.GetDataValue)
        If PO_QTY_REC <> 0 Then
            MsgBox("Receipt Qtys were found in Open Shipment Records",
                   vbOKOnly, "Cannot proceed with Loading 3PL Receipt")
            Return False
        End If

        Dim sql3PLReceipts As String = ""

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "Select WHT3PLR2.*" & vbCrLf _
                & ", WHTSTYLX.ITEM_TYPE, WHTSTYLX.PPK_CODE" & vbCrLf _
                & ", DECODE(WHTSTYLX.ITEM_TYPE,'P',WHTPPKM2.STYLE_CODE,WHTSTYLX.STYLE_CODE) STYLE_CODE" & vbCrLf _
                & ", DECODE(WHTSTYLX.ITEM_TYPE,'P',WHTPPKM2.COLOR_CODE,WHTSTYLX.COLOR_CODE) COLOR_CODE" & vbCrLf _
                & ", DECODE(WHTSTYLX.ITEM_TYPE,'P',WHTPPKM2.PPK_QTY,1) PPK_QTY" & vbCrLf _
                & " from " & ASW("WHT3PLR2") & " WHT3PLR2,WHTSTYLX,WHTPPKM2" & vbCrLf _
                & " where WHTSTYLX.LP_CODE (+) = WHT3PLR2.LP_CODE" & vbCrLf _
                & "   and WHTSTYLX.ITEM_CODE (+) = WHT3PLR2.ITEM_CODE" & vbCrLf _
                & "   and WHTPPKM2.PPK_CODE (+) = WHTSTYLX.PPK_CODE"
            Dim WHT3PLRX As String = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & WHT3PLRX & " Add PPK_QTY_TOTAL NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Update " & WHT3PLRX & " WHT3PLRX Set PPK_QTY_TOTAL = (Select Sum (PPK_QTY) from WHTPPKM2 where PPK_CODE = WHT3PLRX.PPK_CODE) where ITEM_TYPE = 'P'")
            ASCDATA1.ExecuteSQL("Update " & WHT3PLRX & " WHT3PLRX Set RCVQTY = RCVQTY * PPK_QTY / PPK_QTY_TOTAL where ITEM_TYPE = 'P'")


            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is Select * from " & WHT3PLRX & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Begin" & vbCrLf _
                & "    Declare Cursor C2 is" & vbCrLf _
                & "     Select * from " & POTSHIP3 & vbCrLf _
                & "      where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
                & "      for Update;" & vbCrLf _
                & "     REC NUMBER(8,0);" & vbCrLf _
                & "     BAL NUMBER(8,0);" & vbCrLf _
                & "    Begin" & vbCrLf _
                & "     BAL := R1.RCVQTY;" & vbCrLf _
                & "     For R2 in C2 Loop" & vbCrLf _
                & "      If NVL(R2.PO_QTY_SHP,0) - NVL(R2.PO_QTY_REC,0) >= BAL Then" & vbCrLf _
                & "       REC := BAL;" & vbCrLf _
                & "      Else" & vbCrLf _
                & "       REC := NVL(R2.PO_QTY_SHP,0) - NVL(R2.PO_QTY_REC,0);" & vbCrLf _
                & "      End If;" & vbCrLf _
                & "      BAL := BAL - REC;" & vbCrLf _
                & "      Update " & POTSHIP3 & " Set PO_QTY_REC = NVL(PO_QTY_REC,0) + REC, RECEIVED = '1'" & vbCrLf _
                & "       where current of C2;" & vbCrLf _
                & "     End Loop;" & vbCrLf _
                & "     If BAL<>0 Then" & vbCrLf _
                & "      Update " & POTSHIP3 & " Set PO_QTY_REC = NVL(PO_QTY_REC,0) + BAL, RECEIVED = '1'" & vbCrLf _
                & "       where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
                & "         and ROWNUM <=1;" & vbCrLf _
                & "     End If;" & vbCrLf _
                & "    End;" & vbCrLf _
                & "   End;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            sql3PLReceipts = "Select STYLE_CODE, COLOR_CODE, 0 PO_QTY_SHP, 0 PO_QTY_REC, SUM (RCVQTY) RCVQTY" & vbCrLf _
                & ", MIN(ITEM_CODE) ITEM_MIN, MAX(ITEM_CODE) ITEM_MAX from " & WHT3PLRX & vbCrLf _
                & " group by STYLE_CODE, COLOR_CODE"

        Else

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " _
                & "  Select EDI_STYLE_NO STYLE_CODE, 'AST' COLOR_CODE, EDI_PO_ORDER_NO PO_ORDER_NO, EDI_UNITS_RECIEVED RCVQTY" & vbCrLf _
                & "   from EDT944T2 WHERE EDI_DOC_SEQ_NO = '" & TRANS_SEQ & "'" & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Begin" & vbCrLf _
                & "    Declare Cursor C2 is" & vbCrLf _
                & "     Select * from " & POTSHIP3 & vbCrLf _
                & "      where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and PO_ORDER_NO = R1.PO_ORDER_NO" & vbCrLf _
                & "      for Update;" & vbCrLf _
                & "     REC NUMBER(8,0);" & vbCrLf _
                & "     BAL NUMBER(8,0);" & vbCrLf _
                & "    Begin" & vbCrLf _
                & "     BAL := R1.RCVQTY;" & vbCrLf _
                & "     For R2 in C2 Loop" & vbCrLf _
                & "      If NVL(R2.PO_QTY_SHP,0) - NVL(R2.PO_QTY_REC,0) >= BAL Then" & vbCrLf _
                & "       REC := BAL;" & vbCrLf _
                & "      Else" & vbCrLf _
                & "       REC := NVL(R2.PO_QTY_SHP,0) - NVL(R2.PO_QTY_REC,0);" & vbCrLf _
                & "      End If;" & vbCrLf _
                & "      BAL := BAL - REC;" & vbCrLf _
                & "      Update " & POTSHIP3 & " Set PO_QTY_REC = NVL(PO_QTY_REC,0) + REC, RECEIVED = '1'" & vbCrLf _
                & "       where current of C2;" & vbCrLf _
                & "     End Loop;" & vbCrLf _
                & "     If BAL<>0 Then" & vbCrLf _
                & "      Update " & POTSHIP3 & " Set PO_QTY_REC = NVL(PO_QTY_REC,0) + BAL, RECEIVED = '1'" & vbCrLf _
                & "       where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
                & "         and ROWNUM <=1;" & vbCrLf _
                & "     End If;" & vbCrLf _
                & "    End;" & vbCrLf _
                & "   End;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            sql3PLReceipts = "Select EDI_STYLE_NO STYLE_CODE, 'AST' COLOR_CODE, 0 PO_QTY_SHP, 0 PO_QTY_REC, SUM (EDI_UNITS_RECIEVED) RCVQTY" & vbCrLf _
                & ", MIN(EDI_WH_PRODUCT_CODE) ITEM_MIN, MAX(EDI_WH_PRODUCT_CODE) ITEM_MAX from EDT944T2 WHERE EDI_DOC_SEQ_NO = '" & TRANS_SEQ & "'" & vbCrLf _
                & " group by EDI_STYLE_NO"
        End If

        ASCMAIN1.sql = "Select Sum (PO_QTY_REC) from " & POTSHIP3
        PO_QTY_REC = Val(ASCDATA1.GetDataValue)

        Dim RCVQTY As Int64 = 0
        ASCMAIN1.sql = "Select Sum (RCVQTY) from (" & sql3PLReceipts & ")"
        RCVQTY = Val(ASCDATA1.GetDataValue)

        If PO_QTY_REC <> RCVQTY Then

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, SUM (PO_QTY_SHP) PO_QTY_SHP, SUM (PO_QTY_REC) PO_QTY_REC, SUM (RCVQTY) RCVQTY" & vbCrLf _
                & ", CASE WHEN SUM(NVL(RCVQTY,0)) <> 0 THEN SUM(NVL(RCVQTY,0)) - SUM(NVL(PO_QTY_SHP,0)) ELSE 0 END PO_QTY_VAR" & vbCrLf _
                & ", MAX(ITEM_MIN) ITEM_MIN, MAX(ITEM_MAX) ITEM_MAX from (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, SUM (PO_QTY_SHP) PO_QTY_SHP, SUM (PO_QTY_REC) PO_QTY_REC, 0 RCVQTY" & vbCrLf _
                & ", NULL ITEM_MIN, NULL ITEM_MAX from " & POTSHIP3 & vbCrLf _
                & " group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & "union" & vbCrLf _
                & sql3PLReceipts & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE"
            Dim DT As DataTable = ASCDATA1.GetDataTable
            Using F As New ASFMSGBF
                F.Show_grd(DT, Me, "Receipt Qtys mismatch trying to load Receipt records for PO Shipment " & PO_SHIPMENT_NO)
            End Using
            MsgBox("Receipt Qtys mismatch trying to load PO Shipment records for Receipt",
                vbOKOnly, "Cannot proceed with Loading 3PL Receipt")
            Return False
        End If

        ASCMAIN1.sql = "Select PO_SHIPMENT_NO, PO_SHIPMENT_LNO, PO_ORDER_NO, PO_ORDER_LNO, PO_QTY_REC from " & POTSHIP3
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowPOTSHIP3 As DataRow = dst.Tables("POTSHIP3").Rows.Find _
                                         (New Object() {row.Item("PO_SHIPMENT_NO"),
                                                        row.Item("PO_SHIPMENT_LNO"),
                                                        row.Item("PO_ORDER_NO"),
                                                        row.Item("PO_ORDER_LNO")})
            rowPOTSHIP3.Item("PO_QTY_REC") = row.Item("PO_QTY_REC")
        Next


        If dst.Tables("POTSHIP3").Select("ISNULL(PO_QTY_REC,0) > ISNULL(PO_QTY_SHP,0)").Length <> 0 Then
            '    ' NEED TO SPREAD AN OVERAGE QTY TO OTHER LINES ON THE SHIPMENT FOR THE SAME STYLE/COLOR BEFORE JUMPING UGLY HERE
            '    MsgBox("Receipt Qtys mismatch trying to load PO Shipment records for Receipt", _
            '        vbOKOnly, "Cannot proceed with Loading 3PL Receipt")
            '    Return False
            MsgBox("For some styles (in the receiving detail)" & vbCrLf & " the total number of units reported as received by 3PL" & vbCrLf & " is more than Qty Shipped", MsgBoxStyle.OkOnly, "Warning")
        End If

        Dim PO_SHIPMENT_LNOs As New List(Of Integer)
        ASCMAIN1.sql = "Select Distinct PO_SHIPMENT_LNO from " & POTSHIP3 & " where RECEIVED = '1'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            PO_SHIPMENT_LNOs.Add(Val(row.Item("PO_SHIPMENT_LNO") & ""))
        Next

        For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIP2.Rows
            If PO_SHIPMENT_LNOs.Contains(Val(grow.Cells("PO_SHIPMENT_LNO").Value & "")) Then
                If grow.Cells("PO_SHIP_STATUS").Value <> "O" Then
                    MsgBox("Receipt Qtys were found in Open Shipment Records",
                         vbOKOnly, "Cannot proceed with Loading 3PL Receipt")
                    Return False
                Else
                    grow.Cells("PO_SHIP_STATUS").Value = "X"
                    grow.Update()
                End If
            End If
        Next

        Absx1.txtFor("PO_SOURCE_DOC").Text = REF1
        Absx1.dteFor("PO_DATE_RECEIVED").Value = ARRDTE.Date

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Fill_Records("WHT3PLR1")
            Fill_Records("WHT3PLR2")
            Fill_Records("WHT3PLR3")
        End If

        Return True
    End Function

    Sub Delete_Rows_from_Summary_Tables()
        For Each TABLE_NAME As String In New String() {"POTSHIPR", "POTORDRO"}
            Dim rows As New List(Of DataRow)
            Dim sqlw As String = "ISNULL(PO_QTY_SHP,0) = 0"
            If TABLE_NAME = "POTSHIPR" Then sqlw = "ISNULL(QTY_SHP,0) = 0"
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlw)
                If row.GetChildRows(TABLE_NAME & "_POTSHIP3").Length = 0 Then
                    rows.Add(row)
                End If
            Next
            If rows.Count > 0 Then
                For i As Integer = rows.Count - 1 To 0 Step 1
                    Dim row As DataRow = rows(i)
                    row.Delete()
                Next
            End If
        Next
    End Sub

    Sub Print_Record()
        Synch_TABLE_NAME("POTSHIP1")
        Print_Report_Begin()
        Dim MODE As String = "S"
        If receipt_mode Then MODE = "R"
        If cost_calc Then MODE = "C"
        CR_params.Add("MODE", MODE)

        Dim RPT As String = "PORSHIP1"
        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If MENU_ITEM_OBJECT = "POFSHIP1" Or MENU_ITEM_OBJECT = "POFSHIPI" Then
                RPT = "PORSHIPR"
                'check for BTB QVC for GlenRaven NC
                If WHSE_CODE = "NC" Then
                    For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
                        If rowPOTSHIP2("ORDR_NO") <> "" Then
                            Dim ORDR_NO As String = rowPOTSHIP2("ORDR_NO")
                            Dim PO_SHIPMENT_LNO As String = rowPOTSHIP2("PO_SHIPMENT_LNO")
                            RPT = "PORSHPGR"
                            Fill_Record("SOTORDR1", ORDR_NO, False, False)
                            Dim row As DataRow = TBLs("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}'").FirstOrDefault
                            If row Is Nothing Then
                                Fill_Records("SOTORDR2", ORDR_NO, False)
                            End If

                        End If
                    Next
                End If
            End If
        End If
        If MENU_ITEM_OBJECT = "POFSHIPC" Then
            RPT = "PORSHIPC"

            If ASCMAIN1.CLIENT = "VAN" Then
                RPT = "PORSHIPV"
            End If
        End If

        Dim FILTER As String = ""

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            If receipt_mode Then
                For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'X'")
                    Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO"))
                    FILTER = " or {POTSHIP2.PO_SHIPMENT_LNO} = " & CStr(PO_SHIPMENT_LNO)
                Next
                FILTER = Mid(FILTER, 5)
            End If
        End If

        Dim RPT_TITLE As String = "PO Shipment Status Report"
        If MENU_ITEM_OBJECT = "POFSHIPC" Then
            RPT_TITLE = "Landed Costs Report"
            ' PROBABLY SHOULD BE PARAMETERIZED, AND ALSO USED IN THE SCREEN
            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                CR_params.Add("COST_DTL_DEC", "0")
            Else
                CR_params.Add("COST_DTL_DEC", "1")
            End If
        End If
        Generate_Report(RPT, RPT_TITLE,
                        "Shipment " & PO_SHIPMENT_NO & IIf((EntryMode = "E"), " - Edit in Process (" & Me.Text & ")", IIf(Absx1.chkFor("COST_COMPLETE").Checked, " - Costs Complete", " - Incomplete")),
                        FILTER)

        Print_Report_End()
    End Sub

    Function Print_Receiving() As String
        Synch_TABLE_NAME("POTSHIP1")
        Print_Report_Begin()
        Dim MODE As String = "D"
        Dim REPORT_NO As String = ""
        Dim AllClosed As Boolean = True
        Dim FILE_NAME As String = Absx1.txtFor("PO_SHIPMENT_NO").Text

        Fill_Records("WHTWRECX", , , "Select WHTWREC1.*, POTSHIP1.PO_SHIP_VESSEL from WHTWREC1, POTSHIP1 where POTSHIP1.PO_SHIPMENT_NO = WHTWREC1.PO_SHIPMENT_NO and POTSHIP1.PO_SHIPMENT_NO = '" & Absx1.txtFor("PO_SHIPMENT_NO").Text & "'")

        CR_params.Add("MODE", MODE)
        Dim PO_SHIPMENT_LNO As String = grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & ""
        Dim CONTAINER_NO As String = grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO").Value & ""
        Dim RPT As String = "PORSHIPW"
        Dim FILTER As String = "{POTSHIP2.PO_SHIPMENT_LNO} = " & CStr(PO_SHIPMENT_LNO)
        Dim RPT_TITLE As String = "PO Shipment Received Report"

        Generate_Report(RPT, RPT_TITLE, "Shipment " & PO_SHIPMENT_NO & " Lno " & CStr(PO_SHIPMENT_LNO), FILTER, "PDF", FILE_NAME)

        CR_params.Add("MODE", MODE)
        Generate_Report(RPT, RPT_TITLE, "Shipment " & PO_SHIPMENT_NO & " Lno " & CStr(PO_SHIPMENT_LNO), FILTER)

        For Each row As DataRow In ASCDATA1.SelectDistinct("POTSHIP2", "CONTAINER_NO").Select("")
            If dst.Tables("WHTWRECX").Compute("Count(CONTAINER_NO)", "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and CONTAINER_NO = '" & row.Item("CONTAINER_NO") & "'") = 0 Then
                AllClosed = False
            End If
        Next

        If AllClosed Then
            MODE = "S"
            CR_params.Add("MODE", MODE)
            REPORT_NO = Generate_Report(RPT, RPT_TITLE, "Shipment Summary")
        End If

        Print_Report_End()

        Dim frmASFMSGBF As New ASFMSGBF
        Dim Label As New System.Text.StringBuilder With {.Length = 0}
        Label.AppendLine("Enter Email Message for shipment:" & PO_SHIPMENT_NO)
        Dim Caption As String = "Warehouse Receipt"
        Dim emailNote As String = frmASFMSGBF.Get_txtblock_from_User(Label.ToString, Caption, "", False, 0)

        Try
            Dim clsASCNOTE1 As New TAC.ASCNOTE1("PORSHIPW", dst)
            clsASCNOTE1.Note = String.Format("PO Shipment:{0} Received by Whse", PO_SHIPMENT_NO) & vbCrLf & emailNote
            clsASCNOTE1.ReplaceEmailSubject = "Container " & CONTAINER_NO & " for Shipment " & PO_SHIPMENT_NO & " Received by Whse"
            clsASCNOTE1.Attachments.Add(ASCMAIN1.Folders("Temp") & FILE_NAME & ".pdf")
            clsASCNOTE1.CreateComponents()
            clsASCNOTE1.EmailDocument()


            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                & " Select 'POTSHIP1', PO_SHIPMENT_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'CLS_RCV','PO Shipment Received', ''" _
                & " from POTSHIP1 " & vbCrLf _
                & " where (PO_SHIPMENT_NO) in ('" & PO_SHIPMENT_NO & "')"
            ASCDATA1.ExecuteSQL()
        Catch ex As Exception
            MessageBox.Show("Error emailing Warehouse receipts." & ex.Message, "Email Error", MessageBoxButtons.OK)
        End Try

        Return REPORT_NO
    End Function

    Sub Show_Invoice()
        Synch_TABLE_NAME("POTSHIP1")

        Dim REPORTFILE As String = "SORINVP1"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        Dim RPT As String = "SORINVP1"
        Dim AR_PARM_INVOICE_RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If AR_PARM_INVOICE_RPT <> "" Then RPT = AR_PARM_INVOICE_RPT ' "SORINVP1"

        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows(0)
        Dim ORDR_NO As String = Absx1.txtFor("ORDR_NO").Text
        Dim INV_NO As String = "" ' Absx1.txtFor("INV_NO").Text

        If INV_NO <> "" Then ' RECEIPT IS UPDATED
            REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and SOTINVH1.INV_NO = '" & INV_NO & "'"})
            With REPORTS(REPORTFILE).clsASCBASE1
                .Print_Report_Begin()

                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "0")
                .Generate_Report(RPT, "Sales Invoice", , True, , , , , False)
                .Print_Report_End()
            End With
        Else
            REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and SOTORDR1.ORDR_NO = '" & ORDR_NO & "'", "1", "O"})

            For Each rowSOTINVH2 As DataRow In REPORTS(REPORTFILE).dst.Tables("SOTINVH2").Select("")
                rowSOTINVH2.Item("ORDR_QTY_SHIP") = 0
            Next
            Dim INV_SALES As Decimal = 0
            For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
                Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {rowPOTSHIP3.Item("PO_ORDER_NO"), rowPOTSHIP3.Item("PO_ORDER_LNO")})
                Dim rowSOTINVH2() As DataRow = REPORTS(REPORTFILE).dst.Tables("SOTINVH2").Select("INV_LNO = " & CStr(Val(rowPOTORDR2.Item("ORDR_LNO") & "")))
                rowSOTINVH2(0).Item("ORDR_QTY_SHIP") += Val(rowPOTSHIP3.Item("PO_QTY_REC") & "")
                INV_SALES += Val(rowSOTINVH2(0).Item("ORDR_UNIT_PRICE")) * Val(rowPOTSHIP3.Item("PO_QTY_REC"))
            Next

            Dim rowSOTINVH1 As DataRow = REPORTS(REPORTFILE).dst.Tables("SOTINVH1").Rows(0)
            rowSOTINVH1.Item("INV_SALES") = INV_SALES
            rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = INV_SALES + Val(rowSOTINVH1.Item("INV_FREIGHT") & "") + Val(rowSOTINVH1.Item("INV_MISC_CHG") & "")

            With REPORTS(REPORTFILE).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "0")
                .Generate_Report(RPT, "Pro-Forma Sales Invoice", , True, , , , , False)
                .Print_Report_End()
            End With
        End If
    End Sub

    Sub Setup_WHT3PLR1()

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            If grdWHT3PLR1.ActiveRow Is Nothing Then
                spl3PL.Panel2Collapsed = True
            Else
                spl3PL.Panel2Collapsed = False
                Dim dvw As DataView
                Dim TRANS_SEQ As String = grdWHT3PLR1.ActiveRow.Cells("TRANS_SEQ").Value

                dvw = DirectCast(grdWHT3PLR2.DataSource, DataTable).DefaultView
                dvw.RowFilter = "TRANS_SEQ = '" & TRANS_SEQ & "'"
                grdWHT3PLR2.Text = "Receiving Details for Transaction " & TRANS_SEQ

                dvw = DirectCast(grdWHT3PLR3.DataSource, DataTable).DefaultView
                dvw.RowFilter = "TRANS_SEQ = '" & TRANS_SEQ & "'"
                grdWHT3PLR3.Text = "Individual Carton Scans for Transaction " & TRANS_SEQ
            End If
        Else
            If grdEDT944T1.ActiveRow Is Nothing Then
                spl3PL.Panel2Collapsed = True
            Else
                spl3PL.Panel2Collapsed = False
                'Dim dvw As DataView
                Dim EDI_DOC_SEQ_NO As String = grdEDT944T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value

                ASCMAIN1.sql = "Select * from EDT944T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                Fill_Records("WHT3PLR2", "", , ASCMAIN1.sql)

                ASCMAIN1.sql = "Select * from EDT944T3 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                Fill_Records("WHT3PLR3", "", , ASCMAIN1.sql)

                grdWHT3PLR2.Text = "Receiving Details for Transaction " & EDI_DOC_SEQ_NO
                grdWHT3PLR3.Text = "Receiving Exceptions for Transaction " & EDI_DOC_SEQ_NO
            End If
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        UltraExplorerBar1.Groups("Receipts History").Visible = (tab0.SelectedTab.Key = "Receipts History")
        UltraExplorerBar1.Groups("Receipt Type").Visible = (tab0.SelectedTab.Key = "Shipments") And receipt_mode
        If tab0.SelectedTab.Key = "Costing Summaries" Then
            Load_POTSHIPI()
        End If
        If (tab0.SelectedTab.Key = "Payments") Then
            splPaymentOptions.Parent = splPayments.Panel1

            With grdAPTCHCKV.DisplayLayout.Bands(0)
                For Each C As String In New String() {"AMT_SHP", "AMT_REC", "AMT_INV", "AMT_ADV", "AMT_PMT"}
                    .Columns(C).Hidden = False
                Next
            End With

            chkShowAllColumns.Visible = False
            chkIncludeLCs.Visible = True
            Fetch_Payments()

        ElseIf (tab0.SelectedTab.Key = "Open && Paid") Then
            ShowAllColumns()
            splPaymentOptions.Parent = splOpenAndPaid.Panel1

            With grdAPTCHCKV.DisplayLayout.Bands(0)
                For Each C As String In New String() {"AMT_SHP", "AMT_REC", "AMT_INV", "AMT_ADV", "AMT_PMT"}
                    .Columns(C).Hidden = True
                Next
            End With

            chkShowAllColumns.Visible = True
            chkIncludeLCs.Visible = False
            Fetch_Payments()

        ElseIf (tab0.SelectedTab.Key = "AT Shipments") Then
            Setup_AT_Shipments()
        ElseIf (tab0.SelectedTab.Key = "Tariffs") Then
            Fill_Records("POTSHIP5_ALL")
        ElseIf (tab0.SelectedTab.Key = "Bookings") Then
            UltraExplorerBar1.Groups("Screen Control").Items("Import Bookings").Visible = True

            ASCMAIN1.sql = sqlPOTVBKGX & " and VBKG_STATUS = 'F' and PO_SHIPMENT_NO is Null"
            Fill_Records("POTVBKGX", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdPOTVBKGX, "VBKG_NO".ToLower)
        ElseIf (tab0.SelectedTab.Key = "Glen Raven") Then
            Setup_PackingSLips()
        Else
            UltraExplorerBar1.Groups("Screen Control").Visible = True
            UltraExplorerBar1.Groups("Options").Visible = True
            UltraExplorerBar1.Groups("Screen Control").Items("Import Bookings").Visible = False
            UltraExplorerBar1.Groups("Packing Slips").Visible = False
            spl.Panel1Collapsed = False
        End If
    End Sub


    Sub Setup_AT_Shipments()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 Is" & vbCrLf _
            & " Select INVOICE_HDR_KEY, MIN(VAN_REF) MIN_VAN_REF, MAX (VAN_REF) MAX_VAN_REF, COUNT (*) INVS" & vbCrLf _
            & "  from POTIHDRA where STATUS = 'W' group BY INVOICE_HDR_KEY HAVING COUNT (*) > 1;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update POTIHDRA Set STATUS = 'X' where INVOICE_HDR_KEY = R1.INVOICE_HDR_KEY and VAN_REF <> R1.MAX_VAN_REF;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 Is" & vbCrLf _
            & " Select INVHDRKEY, MIN(VAN_REF) MIN_VAN_REF, MAX (VAN_REF) MAX_VAN_REF, COUNT (*) INVS" & vbCrLf _
            & "  from POTPACKA where STATUS = 'W' group BY INVHDRKEY HAVING COUNT (*) > 1;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update POTPACKA Set STATUS = 'X' where INVHDRKEY = R1.INVHDRKEY and VAN_REF <> R1.MAX_VAN_REF;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        EnforceConstraints(False)
        Fill_Records("ATSHIPS")
        Fill_Records("ATINVHDR")
        Fill_Records("ATPACKHDR")

        Fill_Records("ATPACKBAG")
        Fill_Records("ATPACKPO")
        Fill_Records("ATPACKCARTON")

        'Sort_grdColumns(grdATSHIPS, "VAN_REF,packhdrkey", False, 2)
        Sort_grdColumns(grdATSHIPS, "SHIPDATE,CARRIER", False, 0)
        Sort_grdColumns(grdATSHIPS, "VAN_REF,INVNO", False, 1)
        Sort_grdColumns(grdATSHIPS, "VAN_REF,PACKREFNO", False, 2)
        '  grdATSHIPS.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        grdATSHIPS.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
        EnforceConstraints(True)

    End Sub
    Sub Load_POTSHIPI()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Style Shipment Information")

        ASCDATA1.ExecuteSQL("Delete from " & POTSHIPI)
        ASCDATA1.ExecuteSQL("Insert into " & POTSHIPI & " " & sql_POTSHIPI)

        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP from SOTINVH2" & vbCrLf _
            & " where ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " and (STYLE_CODE, COLOR_CODE) in (Select Distinct STYLE_CODE, COLOR_CODE from " & POTSHIPI & ")" & vbCrLf _
            & " group by STYLE_CODE, COLOR_CODE"

        ASCMAIN1.sql = "Select POTSHIPI.*, X.ORDR_QTY_SHIP" _
            & " from " & POTSHIPI & " POTSHIPI,(" & ASCMAIN1.sql & ") X" & vbCrLf _
            & " where X.STYLE_CODE (+) = POTSHIPI.STYLE_CODE" & vbCrLf _
            & "   and X.COLOR_CODE = POTSHIPI.COLOR_CODE" & vbCrLf

        Fill_Records("POTSHIPI", "", True, ASCMAIN1.sql)

        EnforceConstraints(False)
        Fill_Records("POTSHIPF", ASCMAIN1.CYP)
        Fill_Records("POTSHIPC", ASCMAIN1.CYP)
        Fill_Records("POTSHIPC2", ASCMAIN1.CYP)
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdFetchReceipts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetchReceipts.Click

        Load_ICTIRECX()
    End Sub

    Sub Load_ICTIRECX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Receipts History")
        Dim OPS_YYYYPP As String = cbeReceipts.Value
        Dim OPS_YYYYPP2 As String = cbeReceipts2.Value
        Fill_Records("ICTIRECX", New String() {OPS_YYYYPP, OPS_YYYYPP2})
        Sort_grdColumns(grdICTIRECX, "RECEIPT_NO")
        grdICTIRECX.Text = "Receipts History for " & cbeReceipts.Text & IIf(cbeReceipts.Value = cbeReceipts2.Value, "", " thru " & cbeReceipts2.Text)
        Setup_ICTIRECX()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Setup_ICTIRECX()
        If grdICTIRECX.ActiveRow Is Nothing OrElse Not grdICTIRECX.ActiveRow.IsDataRow Then
            splICTIRECX.Panel2Collapsed = True
        Else
            splICTIRECX.Panel2Collapsed = False
            Dim RECEIPT_NO As String = grdICTIRECX.ActiveRow.Cells("RECEIPT_NO").Value
            ASCMAIN1.sql = "Select * from ICTIREC2 where RECEIPT_NO = '" & RECEIPT_NO & "'"
            Fill_Records("ICTIREC2", "", True, ASCMAIN1.sql)
            grdICTIREC2.Text = "Receipts Details for " & RECEIPT_NO
        End If
    End Sub

    Private Sub grdPOTSHIPI_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTSHIPI.DoubleClickRow
        Absx1.txtFor("PO_SHIPMENT_NO").Text = e.Row.Cells("PO_SHIPMENT_NO").Value
        Click_Command("View")
    End Sub

    Private Sub grdPOTSHIPI_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTSHIPI.InitializeLayout

    End Sub

    Private Sub grdPOTSHIPC_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTSHIPC.DoubleClickRow
        If e.Row.Band.Key = "POTSHIPC_POTSHIPC2" Then
            Absx1.txtFor("PO_SHIPMENT_NO").Text = e.Row.Cells("PO_SHIPMENT_NO").Value
            Click_Command("View")
        End If
    End Sub


    Private Sub grdPOTSHIPF_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTSHIPF.DoubleClickRow
        Absx1.txtFor("PO_SHIPMENT_NO").Text = e.Row.Cells("PO_SHIPMENT_NO").Value
        Click_Command("View")
    End Sub

    Private Sub grdPOTSHIP7_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTSHIP7.InitializeLayout

    End Sub

    Private Sub grdPOTSHIP7_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTSHIP7.AfterRowUpdate

        Dim CARTON_DIMS As String = e.Row.Cells("CARTON_DIMS").Value & ""
        If CARTON_DIMS <> "" Then
            Dim PO_ORDER_NO_carton As String = Get_PO_for_Carton(
                                                e.Row.Cells("PO_SHIPMENT_NO").Value,
                                                e.Row.Cells("PO_SHIPMENT_LNO").Value,
                                                e.Row.Cells("CARTON_NO").Value)
            If PO_ORDER_NO_carton <> "" Then
                If CARTON_DIMS_by_PO.ContainsKey(PO_ORDER_NO_carton) Then
                    CARTON_DIMS_by_PO(PO_ORDER_NO_carton) = CARTON_DIMS
                Else
                    CARTON_DIMS_by_PO.Add(PO_ORDER_NO_carton, CARTON_DIMS)
                End If
            End If
        End If
    End Sub

    Function Get_PO_for_Carton(PO_SHIPMENT_NO As String, PO_SHIPMENT_LNO As Int32, CARTON_NO As Int32) As String
        Dim PO_ORDER_NO As String = ""
        Dim rowPOTSHIP7 As DataRow = dst.Tables("POTSHIP7").Rows.Find _
                               (New Object() {PO_SHIPMENT_NO,
                                              PO_SHIPMENT_LNO,
                                              CARTON_NO})

        Dim rowPOTSHIP8() As DataRow = rowPOTSHIP7.GetChildRows("POTSHIP7_POTSHIP8")
        If rowPOTSHIP8.Length > 0 Then
            Dim rowPOTSHIPR As DataRow = rowPOTSHIP8(0).GetParentRow("POTSHIPR_POTSHIP8")
            Dim rowPOTSHIP3() As DataRow = rowPOTSHIPR.GetChildRows("POTSHIPR_POTSHIP3")
            If rowPOTSHIP3.Length > 0 Then
                PO_ORDER_NO = rowPOTSHIP3(0).Item("PO_ORDER_NO")
            End If
        End If

        Return PO_ORDER_NO
    End Function

    Private Sub grdPOTSHIP7_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTSHIP7.BeforeRowUpdate
        Dim CARTON_DIMS As String = e.Row.Cells("CARTON_DIMS").Value & ""
        If CARTON_DIMS = "" Then
            Dim PO_ORDER_NO_carton As String = Get_PO_for_Carton(
                                                e.Row.Cells("PO_SHIPMENT_NO").Value,
                                                e.Row.Cells("PO_SHIPMENT_LNO").Value,
                                                e.Row.Cells("CARTON_NO").Value)
            If PO_ORDER_NO_carton <> "" Then
                If CARTON_DIMS_by_PO.ContainsKey(PO_ORDER_NO_carton) Then
                    e.Row.Cells("CARTON_DIMS").Value = CARTON_DIMS_by_PO(PO_ORDER_NO_carton)
                End If
            End If
        End If
    End Sub

    Function Get_Volume_from_Dims(CARTON_DIMS As String) As Decimal
        Dim CARTON_VOLUME As Decimal = 0
        Dim D() As String = Split(CARTON_DIMS.ToUpper, "X")
        For I As Integer = 1 To D.Length
            If Val(D(I - 1)) <> 0 Then
                If CARTON_VOLUME = 0 Then CARTON_VOLUME = 1
                CARTON_VOLUME *= Val(D(I - 1))
            End If
        Next

        Return CARTON_VOLUME
    End Function

    Private Sub optReceiptType_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optReceiptType.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_POTSHIPX()
    End Sub

    Private Sub grdWHT3PLR1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHT3PLR1.InitializeLayout

    End Sub

    Sub Initialize_EDI()
        ASCDATA1.ExecuteSQL("Update EDT944T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_PROCESS_IND is Null")

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            ' Make Transfer Receipts go away
            '  ASCDATA1.ExecuteSQL("UPDATE EDT944T1 SET EDI_PROCESS_IND = 'T' WHERE EDI_PROCESS_IND = '0' AND LENGTH(EDI_PO_SHIPMENT_NO) <> 6 AND SHIPPER_NAME = 'XXXX'")
            ASCDATA1.ExecuteSQL("UPDATE EDT944T1 SET EDI_PROCESS_IND = 'T' WHERE EDI_PROCESS_IND = '0' AND LENGTH(EDI_PO_SHIPMENT_NO) <> 6")
        End If

        ASCMAIN1.sql = "Update EDT944T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
               & " where TRIM(EDI_TP_ID) = TRIM(EDT944T1.EDI_TP_ID) and TRIM(EDI_OUR_ID) = TRIM(EDT944T1.EDI_OUR_ID))" _
               & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT944T1 Set WHSE_CODE = (Select WHSE_CODE from ICTWHSE1" _
               & " where WHSE_EDI_QUAL = EDT944T1.EDI_TP_QUAL and WHSE_EDI_ID = EDT944T1.EDI_TP_ID and LP_WHSE_ID = EDT944T1.EDI_WH_ID_CODE)" _
               & " where EDI_PROCESS_IND = '0' and WHSE_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT944T1 Set WHSE_CODE = '95'" _
               & " where EDI_PROCESS_IND = '0' and WHSE_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" _
               & "   and EDI_WH_ID_CODE in ('NYDG','NYWM','NYWB')"
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Select * from EDT944T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE is Null"
        'If EDI_DOC_SEQ_NOs_no_company.Count <> 0 Then
        '    ASCMAIN1.sql &= " and EDI_DOC_SEQ_NO Not in ('" & Join(EDI_DOC_SEQ_NOs_no_company.ToArray, "','") & "')"
        'End If
        'Dim dt As DataTable = ASCDATA1.GetDataTable
        'If dt.Rows.Count <> 0 Then
        '    For Each row As DataRow In dt.Rows
        '        EDI_DOC_SEQ_NOs_no_company.Add(row.Item("EDI_DOC_SEQ_NO"))
        '    Next
        '    Using frm As New ASFMSGBF
        '        frm.Show_grd(dt, Me, "EDI Transactions which could not be mapped to an ABSolution Company")
        '    End Using
        'End If
    End Sub

    Private Sub grdEDT944T1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT944T1.DoubleClickRow
        EDI_DOC_SEQ_NO = e.Row.Cells("EDI_DOC_SEQ_NO").Value
        Dim WHSE_CODE As String = e.Row.Cells("WHSE_CODE").Value
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim LP_CODE As String = rowICTWHSE1.Item("LP_CODE")

        Dim CLOSE_DTE As Date = e.Row.Cells("EDI_RECEIPT_DATE").Value
        Dim ARRDTE As Date = e.Row.Cells("EDI_ARRIVAL_DATE").Value
        Dim REF1 As String = e.Row.Cells("EDI_WH_RECEIPT_NO").Value & ""
        If REF1 = "" Then REF1 = "3PL " & EDI_DOC_SEQ_NO

        PO_SHIPMENT_NO = e.Row.Cells("EDI_PO_SHIPMENT_NO").Value
        Absx1.txtFor("PO_SHIPMENT_NO").Text = PO_SHIPMENT_NO

        ',
        'EDI_SEAL_NUMBER_CONTAINER,
        'EDI_SEAL_NUMBER,
        'EDI_TOTAL_QTY_RECEIVED,

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        select_from_3PL_list = True
        Click_Command("Select")

        If ScreenMode Then
            If Not Load_3PL_Receipt_Details(EDI_DOC_SEQ_NO, LP_CODE, WHSE_CODE, ARRDTE, REF1) Then
                EntryMode = "X"
                Click_Command("Cancel")
            Else
                Setup_WHT3PLR1()
            End If
        Else
            select_from_3PL_list = False
        End If
    End Sub

    Function Check_Cartons_in_Balance(skip_message_if_no_cartons As Boolean) As String

        Dim EMsg As String = ""

        For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("")
            Dim PO_SHIPMENT_LNO As Int64 = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
            Dim CARTONS As Int64 = Val(dst.Tables("POTSHIP7").Compute("SUM(CARTONS)", "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)) & "")
            Dim PO_SHIP_CTNS As Int64 = Val(rowPOTSHIP2.Item("PO_SHIP_CTNS") & "")
            If PO_SHIP_CTNS <> CARTONS Then
                EMsg &= vbCr & "Carton Count for Some BOLs do not match Packing Details" _
                & vbCr & " (See Line " & rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & ", " & CStr(PO_SHIP_CTNS) & " vs " & CStr(CARTONS) & ")"
                Exit For
            End If
        Next

        If dst.Tables("POTSHIPR").Select("ISNULL(QTY_VAR,0) <> 0").Length <> 0 Then
            EMsg &= vbCr & "Some Lines are NOT Cartonized in Balance with Shipment Qtys"
        End If

        If skip_message_if_no_cartons Then
            If dst.Tables("POTSHIPR").Select("ISNULL(QTY_CTN,0) <> 0").Length = 0 Then
                EMsg = ""
            End If
        End If

        Return EMsg
    End Function

    Sub Create_Containers_from_BOL()

        tabBOL.SelectedTab = tabBOL.Tabs("PO Details")

        If grdPOTSHIP4.ActiveRow IsNot Nothing Then
            grdPOTSHIP4.ActiveRow.CancelUpdate()
        End If

        Dim allowAddNew As Boolean = True
        If grdPOTSHIP4.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No Then
            allowAddNew = False
            grdPOTSHIP4.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        End If

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTSHIP2"), New String() {"CONTAINER_NO"}).Rows
            Dim CONTAINER_NO As String = row.Item("CONTAINER_NO") & ""
            CONTAINER_NO = ASCMAIN1.Format_Field(CONTAINER_NO, "CONTAINER_NO")
            If CONTAINER_NO <> "" Then
                Dim sqlw As String = "CONTAINER_NO = '" & CONTAINER_NO & "'"
                Dim rowContainer() As DataRow = dst.Tables("POTSHIP4").Select(sqlw)
                If rowContainer.Length = 0 Then
                    Dim PO_SHIP_CTNS As Integer = Val(dst.Tables("POTSHIP2").Compute("SUM(PO_SHIP_CTNS)", sqlw) & "")
                    Dim PO_SHIPMENT_LNO As Integer = Val(dst.Tables("POTSHIP4").Compute("MAX(PO_SHIPMENT_LNO)", "") & "")
                    grdPOTSHIP4.DisplayLayout.Bands(0).AddNew()
                    With grdPOTSHIP4.ActiveRow
                        .Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO
                        .Cells("PO_SHIPMENT_LNO").Value = PO_SHIPMENT_LNO + 1
                        .Cells("CONTAINER_NO").Value = CONTAINER_NO
                        .Cells("PO_SHIP_CTNS").Value = PO_SHIP_CTNS
                        .Update()
                    End With
                Else
                    If rowContainer.Length = 1 AndAlso packingFromXLS Then
                        Dim PO_SHIP_CTNS As Integer = Val(dst.Tables("POTSHIP2").Compute("SUM(PO_SHIP_CTNS)", sqlw) & "")
                        Dim rowPOTSHIP4 As DataRow = rowContainer(0)
                        rowPOTSHIP4.Item("PO_SHIP_CTNS") = PO_SHIP_CTNS
                    End If
                End If
            End If
        Next
        Sort_grdColumns(grdPOTSHIP4, "PO_SHIPMENT_LNO")

        If Not allowAddNew Then
            grdPOTSHIP4.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        End If
    End Sub

    Private Sub grdSOTORDP1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDP1.ClickCellButton
        Dim ORDR_NO As String = e.Cell.Row.Cells("ORDR_NO").Value

        ' Dim sql_where As String = "ORDR_NO = '" & grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & "' and NVL(INV_STATUS,'0') = '0'"
        Dim sql_where As String = "ORDR_NO = '" & ORDR_NO & "' and NVL(INV_STATUS,'0') = '0'"
        grdClickCellButton(grdSOTORDP1, sql_where, sql_where <> "")
        If e.Cell.Row.Cells("INV_NO_PREV").Value & "" <> "" Then
            'If e.Cell.Row.Cells("INV_DATE").Value & "" = "" Then
            Dim rowSOTORDP1 As DataRow = LookUp("SOTORDP1", New String() {grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value, e.Cell.Row.Cells("INV_NO_PREV").Value & ""})
            If rowSOTORDP1 IsNot Nothing Then
                e.Cell.Row.Cells("INV_DATE").Value = rowSOTORDP1.Item("INV_DATE")
                e.Cell.Row.Cells("INV_REF").Value = rowSOTORDP1.Item("INV_REF")
                e.Cell.Row.Cells("INV_COMMENT").Value = rowSOTORDP1.Item("INV_COMMENT")
            End If
            'End If
        End If

    End Sub

    Function Check_INV_NO_PREV(ORDR_NO As String, INV_NO_PREV As String) As Boolean

        For Each row As DataRow In dst.Tables("SOTORDP1").Select("ISNULL(INV_NO_PREV,'') <> ''")
            ASCMAIN1.sql = "Select * from SOTORDP2 where ORDR_NO = '" & ORDR_NO & "' and INV_NO = '" & INV_NO_PREV & "'"
            Dim SCQ As String = ""
            For Each row2 As DataRow In ASCDATA1.GetDataTable.Select("ORDR_QTY_SHIP<>0", "ORDR_LNO")
                SCQ &= row2.Item("ORDR_LNO") & ":" & row2.Item("ORDR_QTY_SHIP")
            Next
            Dim SCQ2 As String = ""
            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and INV_NO = '" & ORDR_NO & "'"
            For Each row2 As DataRow In dst.Tables("SOTORDP2").Select(sqlw, "ORDR_LNO")
                SCQ2 &= row2.Item("ORDR_LNO") & ":" & row2.Item("ORDR_QTY_SHIP")
            Next

            Return SCQ = SCQ2
        Next

    End Function

    Private Sub numDuty_ValueChanged(sender As System.Object, e As System.EventArgs) Handles numDuty.ValueChanged
        Calculate_DUTY_DIST()
    End Sub

    Sub Calculate_DUTY_DIST()
        If SELECTION_NO = 0 Then Exit Sub

        Dim TOTAL_DUTY As Decimal = System.Math.Round(Val(dst.Tables("POTSHIP2").Compute("SUM(TOTAL_DUTY)", "") & ""), 2)
        Dim DUTY As Decimal = Val(dst.Tables("POTSHIP5").Compute("SUM(LANDING_COST_D)", "") & "")
        Dim CUSTOMS_DUTY_AMT_DIST As Decimal = TOTAL_DUTY '  + DUTY - ALREADY ADDED INTO TOTAL_DUTY SINCE WE ARE CALLING CALCULATE COSTS SO MUCH
        '  Absx1.numFor("CUSTOMS_DUTY_AMT_DIST").Value = CUSTOMS_DUTY_AMT_DIST
        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            CUSTOMS_DUTY_AMT_DIST = DUTY
            Absx1.numFor("CUSTOMS_DUTY_AMT").Value = DUTY
            Absx1.txtFor("CUSTOMS_ENTRY_NO").Text = "REGENCY"
        End If
        numDutyDist.Value = CUSTOMS_DUTY_AMT_DIST
        numDutyNotDist.Value = Val(Absx1.numFor("CUSTOMS_DUTY_AMT").Value & "") - CUSTOMS_DUTY_AMT_DIST

    End Sub

    Private Sub grdICTIRECX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTIRECX.AfterRowActivate
        Setup_ICTIRECX()
    End Sub

    Sub DutyBalance()
        Calculate_Landed_Cost()
        Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and LANDING_COST_DIST = 'D' and CTL_NO is Null"
        Dim rowPOTSHIP5s() As DataRow = dst.Tables("POTSHIP5").Select(sqlw)
        Dim rowPOTSHIP5 As DataRow = Nothing
        If rowPOTSHIP5s.Length = 0 Then
            rowPOTSHIP5 = dst.Tables("POTSHIP5").NewRow
            rowPOTSHIP5.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP5.Item("PO_SHIPMENT_LNO") = Val(dst.Tables("POTSHIP5").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
            rowPOTSHIP5.Item("COST_CATGY_CODE") = "DUTY"
            rowPOTSHIP5.Item("COST_CATGY_DESC") = "Duty Adjustment"
            rowPOTSHIP5.Item("LANDING_COST_DIST") = "D"
            dst.Tables("POTSHIP5").Rows.Add(rowPOTSHIP5)
        Else
            rowPOTSHIP5 = rowPOTSHIP5s(0)
        End If
        rowPOTSHIP5.Item("LANDING_COST_AMT") = Val(rowPOTSHIP5.Item("LANDING_COST_AMT") & "") + Val(numDuty.Value & "") - Val(numDutyDist.Value & "")
        Calculate_Landed_Cost()
    End Sub

    Private Sub cmdDutyBalance_Click(sender As System.Object, e As System.EventArgs) Handles cmdDutyBalance.Click
        DutyBalance()
    End Sub

    Private Sub chkNoDuty_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkNoDuty.CheckedChanged
        If Not ScreenMode Or IsLoading Then Exit Sub
        Calculate_Landed_Cost()
    End Sub

    Private Sub grdPOTSHIP3_LostFocus(sender As Object, e As System.EventArgs) Handles grdPOTSHIP3.LostFocus
        'For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIP3.Rows
        '    If grow.DataChanged Then
        '        grow.Update()
        '    End If
        'Next
    End Sub

    Private Sub grdICTIRECX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIRECX.DoubleClickRow
        Absx1.txtFor("PO_SHIPMENT_NO").Text = e.Row.Cells("PO_SHIPMENT_NO").Text
        Click_Command("View")
    End Sub

    Private Sub grdWHTWRECX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTWRECX.DoubleClickRow
        Dim Not_Completed_Shipments As String = ""
        ASCMAIN1.sql = "Select * from WHTWREC1" _
        & " Where PO_SHIPMENT_NO = '" & e.Row.Cells("PO_SHIPMENT_NO").Value & "' and WH_REC_STATUS <> 'C'"
        If ASCMAIN1.CLIENT = "VAN" Then ASCMAIN1.sql &= "  and WH_REC_STATUS <> 'R'"
        For Each rowWHTWREC1 As DataRow In ASCDATA1.GetDataTable.Rows
            Select Case rowWHTWREC1.Item("WH_REC_STATUS")
                Case "R"
                    Not_Completed_Shipments &= "  Container No: " & rowWHTWREC1.Item("CONTAINER_NO") & " - Received" & vbCrLf
                Case Else
                    Not_Completed_Shipments &= "  Container No: " & rowWHTWREC1.Item("CONTAINER_NO") & " - Not Complete" & vbCrLf
            End Select
        Next

        ASCMAIN1.sql = "Select Distinct CONTAINER_NO from POTSHIP2" _
        & " Where PO_SHIPMENT_NO = '" & e.Row.Cells("PO_SHIPMENT_NO").Value & "' and WH_REC_NO is Null"
        For Each rowPOTSHIP2 As DataRow In ASCDATA1.GetDataTable.Rows
            Not_Completed_Shipments &= "  Container No: " & rowPOTSHIP2.Item("CONTAINER_NO") & " - Not Entered" & vbCrLf
        Next

        If Not_Completed_Shipments <> "" Then

            MsgBox("There are Whse Receipts for this Shipment that are NOT Complete" & vbCrLf _
                                                   & Not_Completed_Shipments & vbCrLf & vbCrLf _
                                                   , MsgBoxStyle.OkOnly, "Pay Attention!")
            If ASCMAIN1.CLIENT = "RGI" Then
            Else
                Exit Sub
            End If
        End If
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Loading Receipt")
        PO_SHIPMENT_NO = e.Row.Cells("PO_SHIPMENT_NO").Value
        Absx1.txtFor("PO_SHIPMENT_NO").Text = PO_SHIPMENT_NO


        Select_from_Whse_Receipt = True
        Click_Command("Select")
        If ScreenMode Then
            Setup_Whse_Receipt_Details()
            Absx1.txtFor("PO_SOURCE_DOC").Text = "Whse Receipt"
            Absx1.dteFor("PO_DATE_RECEIVED").Value = Now
        Else
            Select_from_Whse_Receipt = False
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Whse_Receipt_Details()
        Dim PO_Shipment_No_Fill As String = ""
        Dim PO_Shipment_Lno_Fill As Integer = 0
        Dim Style_Code_Fill As String = ""
        Dim Color_Code_Fill As String = ""
        Dim Units_Received As Integer = 0
        Dim Total_Units_Received_Balance As Integer = 0
        Dim POTSHIP3_Row_Count As Integer = 0
        Dim POTSHIP3_Current_Row As Integer = 1
        Dim rowWHTWRECD As DataRow = Nothing
        Dim RCV_OPEN As Integer = 0

        EMsg = ""

        If ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.sql = " Select '0' WH_REC_NO, C2.PO_SHIPMENT_NO, C2.PO_SHIPMENT_LNO," & vbCrLf _
            & " C2.STYLE_CODE, C2.COLOR_CODE, " & vbCrLf _
            & " Sum(Nvl(C2.PO_QTY_REC, 0)) As UNITS_REC" & vbCrLf _
            & " from WHTWREC1 C1,WHTPREC2 C2" & vbCrLf _
            & " Where C1.WH_REC_NO = C2.WH_REC_NO" & vbCrLf _
            & " And C2.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
            & " And C1.WH_REC_STATUS = 'C'" & vbCrLf _
            & " Group By '0', C2.PO_SHIPMENT_NO, C2.PO_SHIPMENT_LNO, " & vbCrLf _
            & " C2.STYLE_CODE, C2.COLOR_CODE, " & vbCrLf _
            & " '0'" & vbCrLf
            Fill_Records("WHTWRECD", , , ASCMAIN1.sql)
        Else
            ASCMAIN1.sql = " Select C7.WH_REC_NO, C7.PO_SHIPMENT_NO, C7.PO_SHIPMENT_LNO," & vbCrLf _
            & " C8.STYLE_CODE, C8.COLOR_CODE, " & vbCrLf _
            & " Sum(Nvl(C7.CARTONS_RECEIVED,0) *  QTY) As UNITS_REC" & vbCrLf _
            & " from WHTWREC1 C1,WHTWREC7 C7, WHTWREC8 C8" & vbCrLf _
            & " Where C1.WH_REC_NO = C7.WH_REC_NO" & vbCrLf _
            & " And C7.WH_REC_NO = C8.WH_REC_NO" & vbCrLf _
            & " And C7.PO_SHIPMENT_NO = C8.PO_SHIPMENT_NO" & vbCrLf _
            & " And C7.PO_SHIPMENT_LNO = C8.PO_SHIPMENT_LNO" & vbCrLf _
            & " And C7.CARTON_NO = C8.CARTON_NO" & vbCrLf _
            & " And C7.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
            & " And C1.WH_REC_STATUS = 'C'" & vbCrLf _
            & " Group By C7.WH_REC_NO, C7.PO_SHIPMENT_NO, C7.PO_SHIPMENT_LNO," & vbCrLf _
            & " C8.STYLE_CODE, C8.COLOR_CODE, " & vbCrLf _
            & " C1.CONTAINER_NO" & vbCrLf
            Fill_Records("WHTWRECD", , , ASCMAIN1.sql)
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            For Each rowWHTWRECD In dst.Tables("WHTWRECD").Select("", "PO_SHIPMENT_NO, STYLE_CODE, COLOR_CODE")
                POTSHIP3_Row_Count = dst.Tables("POTSHIP3").Select("PO_SHIP_STATUS = 'O' And PO_SHIPMENT_NO = '" & rowWHTWRECD.Item("PO_SHIPMENT_NO") & "'" _
                                 & " And STYLE_CODE = '" & rowWHTWRECD.Item("STYLE_CODE") & "'" _
                                 & " And COLOR_CODE = '" & rowWHTWRECD.Item("COLOR_CODE") & "'").Count
                Total_Units_Received_Balance = rowWHTWRECD.Item("UNITS_REC")
                Dim upc_cnt As Integer = 0
                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIP_STATUS = 'O' And PO_SHIPMENT_LNO = '" & rowWHTWRECD("PO_SHIPMENT_LNO") & "' and STYLE_CODE = '" & rowWHTWRECD("STYLE_CODE") & "' and COLOR_CODE = '" & rowWHTWRECD("COLOR_CODE") & "'", "PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE")
                    upc_cnt += 1
                    RCV_OPEN = rowPOTSHIP3.Item("PO_QTY_SHP") - Val(rowPOTSHIP3.Item("PO_QTY_REC") & "")
                    If upc_cnt = POTSHIP3_Row_Count Or RCV_OPEN >= Total_Units_Received_Balance Then
                        rowPOTSHIP3.Item("PO_QTY_REC") = Val(rowPOTSHIP3.Item("PO_QTY_REC") & "") + Total_Units_Received_Balance
                        Total_Units_Received_Balance = 0
                    ElseIf RCV_OPEN < Total_Units_Received_Balance Then
                        rowPOTSHIP3.Item("PO_QTY_REC") = Val(rowPOTSHIP3.Item("PO_QTY_REC") & "") + RCV_OPEN
                        Total_Units_Received_Balance = Total_Units_Received_Balance - RCV_OPEN
                    End If
                Next
                If Total_Units_Received_Balance > 0 Then
                    For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIP_STATUS = 'O' And PO_SHIPMENT_LNO <> '" & rowWHTWRECD("PO_SHIPMENT_LNO") & "' and STYLE_CODE = '" & rowWHTWRECD("STYLE_CODE") & "' and COLOR_CODE = '" & rowWHTWRECD("COLOR_CODE") & "'", "PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE")
                        upc_cnt += 1
                        RCV_OPEN = rowPOTSHIP3.Item("PO_QTY_SHP") - Val(rowPOTSHIP3.Item("PO_QTY_REC") & "")
                        If upc_cnt = POTSHIP3_Row_Count Or RCV_OPEN >= Total_Units_Received_Balance Then
                            rowPOTSHIP3.Item("PO_QTY_REC") = Val(rowPOTSHIP3.Item("PO_QTY_REC") & "") + Total_Units_Received_Balance
                            Total_Units_Received_Balance = 0
                        ElseIf RCV_OPEN < Total_Units_Received_Balance Then
                            rowPOTSHIP3.Item("PO_QTY_REC") = Val(rowPOTSHIP3.Item("PO_QTY_REC") & "") + RCV_OPEN
                            Total_Units_Received_Balance = Total_Units_Received_Balance - RCV_OPEN
                        End If
                    Next
                End If
                If Total_Units_Received_Balance > 0 Then
                    EMsg = EMsg & vbCrLf & rowWHTWRECD.Item("STYLE_CODE") & " - " & rowWHTWRECD.Item("COLOR_CODE") & " Qty: " & Total_Units_Received_Balance

                End If
            Next
            If EMsg <> "" Then
                MsgBox("Failed to allocate Received Qty, Contact ABS" & EMsg, MsgBoxStyle.Critical, "receiving error")
            End If
            ASCMAIN1.sql = " Select C2.WH_REC_NO, C2.PO_SHIPMENT_NO, C2.PO_SHIPMENT_LNO," & vbCrLf _
            & " C2.STYLE_CODE, C2.COLOR_CODE, " & vbCrLf _
            & " Sum(Nvl(C2.PO_QTY_REC, 0)) As UNITS_REC" & vbCrLf _
            & " from WHTWREC1 C1,WHTPREC2 C2" & vbCrLf _
            & " Where C1.WH_REC_NO = C2.WH_REC_NO" & vbCrLf _
            & " And C2.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
            & " And C1.WH_REC_STATUS = 'C'" & vbCrLf _
            & " Group By C2.WH_REC_NO, C2.PO_SHIPMENT_NO, C2.PO_SHIPMENT_LNO, " & vbCrLf _
            & " C2.STYLE_CODE, C2.COLOR_CODE, " & vbCrLf _
            & " '0'" & vbCrLf
            Fill_Records("WHTWRECD", , , ASCMAIN1.sql)
        Else

            For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIP_STATUS = 'O'", "PO_SHIPMENT_NO, PO_SHIPMENT_LNO")
                Dim PO_SHIPMENT_LNO As Int32 = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")

                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO), "PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE")

                    POTSHIP3_Row_Count = dst.Tables("POTSHIP3").Select("PO_SHIPMENT_NO = '" & rowPOTSHIP3.Item("PO_SHIPMENT_NO") & "'" _
                                         & " And PO_SHIPMENT_LNO = " & rowPOTSHIP3.Item("PO_SHIPMENT_LNO") _
                                         & " And STYLE_CODE = '" & rowPOTSHIP3.Item("STYLE_CODE") & "'" _
                                         & " And COLOR_CODE = '" & rowPOTSHIP3.Item("COLOR_CODE") & "'").Count
                    ' If rowPOTSHIP3.Item("STYLE_CODE") = "801995NMX" And rowPOTSHIP3.Item("COLOR_CODE") = "672" Then Stop

                    If PO_Shipment_No_Fill = rowPOTSHIP3.Item("PO_SHIPMENT_NO") And PO_Shipment_Lno_Fill = rowPOTSHIP3.Item("PO_SHIPMENT_LNO") _
                        And Style_Code_Fill = rowPOTSHIP3.Item("STYLE_CODE") And Color_Code_Fill = rowPOTSHIP3.Item("COLOR_CODE") Then
                        If POTSHIP3_Row_Count = POTSHIP3_Current_Row Then
                            Units_Received = Total_Units_Received_Balance
                            POTSHIP3_Current_Row = 1
                        Else
                            If Total_Units_Received_Balance <= rowPOTSHIP3.Item("PO_QTY_SHP") Then
                                Units_Received = Total_Units_Received_Balance
                                Total_Units_Received_Balance = 0
                            Else
                                Units_Received = rowPOTSHIP3.Item("PO_QTY_SHP")
                                Total_Units_Received_Balance -= Units_Received
                            End If
                            POTSHIP3_Current_Row += 1
                        End If
                    Else
                        PO_Shipment_No_Fill = rowPOTSHIP3.Item("PO_SHIPMENT_NO")
                        PO_Shipment_Lno_Fill = rowPOTSHIP3.Item("PO_SHIPMENT_LNO")
                        Style_Code_Fill = rowPOTSHIP3.Item("STYLE_CODE")
                        Color_Code_Fill = rowPOTSHIP3.Item("COLOR_CODE")

                        rowWHTWRECD = dst.Tables("WHTWRECD").Rows.Find(New Object() {PO_Shipment_No_Fill,
                                                                            PO_Shipment_Lno_Fill,
                                                                            Style_Code_Fill,
                                                                            Color_Code_Fill})
                        If rowWHTWRECD IsNot Nothing Then
                            Total_Units_Received_Balance = rowWHTWRECD.Item("UNITS_REC")
                            If POTSHIP3_Row_Count > 1 Then
                                Units_Received = rowPOTSHIP3.Item("PO_QTY_SHP")
                                Total_Units_Received_Balance -= Units_Received
                                POTSHIP3_Current_Row += 1
                            Else
                                Units_Received = rowWHTWRECD.Item("UNITS_REC")
                                POTSHIP3_Current_Row = 1
                            End If
                        Else
                            Units_Received = 0
                            POTSHIP3_Current_Row = 1
                        End If
                    End If
                    rowPOTSHIP3.Item("PO_QTY_REC") = Units_Received
                Next
            Next
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIP2.Rows
            If ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "VAN" Then
                If grow.Cells("PO_SHIP_STATUS").Value = "O" Then
                    If dst.Tables("WHTWRECD").Compute("COUNT(PO_SHIPMENT_LNO)", "WH_REC_NO = '" & grow.Cells("WH_REC_NO").Value & "'") > 0 Then
                        grow.Cells("PO_SHIP_STATUS").Value = "X"
                    End If
                End If
            Else
                grow.Cells("PO_SHIP_STATUS").Value = "X"
            End If
            grow.Update()
        Next
    End Sub

    Private Sub btnABSonly_Click(sender As System.Object, e As System.EventArgs) Handles btnABSonly.Click
        automated_cost_complete = True

        ASCMAIN1.sql = "SELECT PO_SHIPMENT_NO FROM POTSHIP1 WHERE PO_SHIPMENT_NO IN (" & vbCrLf _
            & "SELECT DISTINCT PO_SHIPMENT_NO FROM POTSHIP2 WHERE OPS_YYYYPP <= '201512')" & vbCrLf _
            & "AND NVL(COST_COMPLETE,'0') = '0'" '  and PO_SHIPMENT_NO IN ('000422','000668','000372','000150','000177','000178','000421')"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        For Each row As DataRow In tbl.Select("", "PO_SHIPMENT_NO")
            Dim PO_SHIPMENT_NO As String = row.Item("PO_SHIPMENT_NO")

            Absx1.txtFor("PO_SHIPMENT_NO").Text = PO_SHIPMENT_NO
            rowPOTSHIP1 = Fill_Record("POTSHIP1", PO_SHIPMENT_NO)

            Click_Command("Edit")

            Dim FRT_TERMS As String = ""
            If grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & "" <> "" Then
                Dim ORDR_NO As String = grdPOTSHIP2.ActiveRow.Cells("ORDR_NO").Value & ""
                Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO)
                FRT_TERMS = rowSOTORDR1.Item("FRT_TERMS")
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                If rowPOTSHIP1.Item("WHSE_CODE") & "" = "FE" And FRT_TERMS = "COL" Then
                    chkNoDuty.Checked = True
                End If
            End If


            chkCostComplete.Checked = True

            Calculate_Landed_Cost()

            'If Val(numDuty.Value & "") = 0 And Val(numDutyNotDist.Value & "") <> 0 Then
            '    numDuty.Value = -1 * (Val(numDutyNotDist.Value & ""))
            '    Absx1.txtFor("CUSTOMS_ENTRY_NO").Text = "X"
            '    If Val(numDutyNotDist.Value & "") <> 0 Then
            '        DutyBalance()
            '    End If
            'End If

            Click_Command("Update")
            If ScreenMode Then
                If EMsg.Contains("Some Styles have no Duty") Then
                    ASCDATA1.ExecuteSQL("Insert into POTSHIP1_EMSG values ('" & PO_SHIPMENT_NO & "','No Containers Entered')")
                    Click_Command("Cancel")
                Else
                    Exit For
                End If
            End If
            'If ScreenMode Then Exit For
        Next

        automated_cost_complete = False
    End Sub

    Private Sub cmdCreate_Click(sender As Object, e As EventArgs) Handles cmdCreate.Click
        If grdPOTSHIPR.ActiveRow Is Nothing Then
            Exit Sub
        End If

        grdPOTSHIPR.Selected.Rows.Clear()
        grdPOTSHIPR.ActiveRow.Selected = True
        Dim PO_SHIPMENT_LNO As Integer = Val(grdPOTSHIPR.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
        For Each row As DataRow In dst.Tables("POTSHIPQ").Select("", "CTN_NO")
            Dim CTNS As Integer = Val(row.Item("CTNS") & "")
            Dim PACK As Integer = Val(row.Item("PACK") & "")
            Dim NOTE As String = row.Item("NOTE") & ""

            Dim rowPOTSHIP7 As DataRow = Create_Carton_for_Selected_Styles(PO_SHIPMENT_LNO)
            rowPOTSHIP7.Item("CARTONS") = CTNS
            rowPOTSHIP7.Item("CARTON_COMMENTS") = NOTE
            rowPOTSHIP7.GetChildRows(("POTSHIP7_POTSHIP8"))(0).Item("QTY") = PACK
        Next
    End Sub

    Private Sub grdPOTSHIPQ_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdPOTSHIPQ.BeforeRowUpdate
        Dim CTNS As Integer = Val(e.Row.Cells("CTNS").Value & "")
        Dim PACK As Integer = Val(e.Row.Cells("PACK").Value & "")
        If CTNS <= 0 Or PACK <= 0 Then
            e.Cancel = True
        Else
            Dim CTN_NO As Integer = Val(dst.Tables("POTSHIPQ").Compute("MAX(CTN_NO)", "") & "") + 1
            e.Row.Cells("CTN_NO").Value = CTN_NO
        End If
    End Sub

    Private Sub btnPrintLabels_Click(sender As Object, e As EventArgs) Handles btnPrintLabels.Click
        PrintLabels()
    End Sub

    Sub PrintLabels()

        Dim TACMAIN1 As New TAC.TACMAIN1

        If tabBOL.SelectedTab.Key = "Cartons" Then
            If grdPOTSHIP7.Selected.Rows.Count = 0 Then
                MsgBox("Please Highlight Lines to print on left bottom Grid", vbOKOnly, "No Cartons Selected")
            End If

            Using ipp As New nsoftware.IPWorks.Ipport
                ipp.RuntimeLicense = TACMAIN1.nSoftwareIPWorksV9Key
                ipp.Connect("192.168.110.223", "4444")
                Dim data As String '= "upc123" ' & vbCrLf a new line is needed to send the data across
                Try
                    For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIP7.Selected.Rows
                        ASCMAIN1.sql = "select c1.STYLE_CODE, c1.COLOR_CODE, UPC_CODE, STYLE_DESC" & vbCrLf _
                                & " from ICTSTYC1 c1, ICTSTYL1 s1" & vbCrLf _
                                & " where s1.STYLE_CODE = c1.STYLE_CODE" & vbCrLf _
                                & " and c1.STYLE_CODE = '" & grow.Cells("STYLE_CODE").Value & "'" & vbCrLf _
                                & " and c1.COLOR_CODE = '" & grow.Cells("COLOR_CODE").Value & "'"
                        Dim row As DataRow = ASCDATA1.GetDataTable.Select("").First

                        data = cbxLabelPrinter.SelectedItem
                        data &= "|" & row.Item("STYLE_CODE") &
                        "|" & row.Item("COLOR_CODE") &
                        "|" & row.Item("STYLE_DESC") &
                        "|" & row.Item("UPC_CODE") &
                        "|" & If(ASCMAIN1.Running_in_VS, 1, grow.Cells("CARTONS").Value)
                        ipp.SendLine(data)
                    Next
                    grdPOTSHIP7.Selected.Rows.Clear()

                Catch ex As Exception

                End Try

                ipp.Disconnect()
            End Using
        End If
    End Sub

    Private Sub btnRcvRefresh_Click(sender As Object, e As EventArgs) Handles btnRvcRefresh.Click
        LoadWhseReceipts()

    End Sub

    Sub LoadWhseReceipts()
        Dim PO_SHIPMENT_NO As String = Absx1.txtFor("PO_SHIPMENT_NO").Text
        Dim CONTAINER_NO As String = String.Empty
        Dim sqlWhere As String
        Dim WH_REC_NO As String = ""
        Dim PO_SHIPMENT_LNOs As String = ""
        Dim REC_LOC_QTY As Int32 = 0
        Dim RCV_OPEN As Integer = 0

        grdWHTPREC3.Text = "Receiving Statistics"
        If grdPOTSHIP2.ActiveRow IsNot Nothing Then
            CONTAINER_NO = grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO").Value & ""
            For Each row As DataRow In ASCDATA1.GetDataTable("Select * from WHTWREC1 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and CONTAINER_NO = '" & CONTAINER_NO & "'").Rows
                grdWHTPREC3.Text = "Unloaded by " & row.Item("UNLOADED_BY_OPER") & " on " & row.Item("WH_DATE_UNLOADED")
                WH_REC_NO = row.Item("WH_REC_NO")
            Next
        End If

        For Each row As DataRow In dst.Tables("WHTPREC3").Select("")
            row.Item("PO_QTY_REC") = 0
        Next

        For Each row As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and CONTAINER_NO = '" & CONTAINER_NO & "'", "PO_SHIPMENT_LNO")
            PO_SHIPMENT_LNOs += "," & row("PO_SHIPMENT_LNO")
        Next


        ASCMAIN1.sql = "" _
            & "select PO_SHIPMENT_NO, STYLE_CODE, COLOR_CODE, sum(PO_QTY_REC) PO_QTY_REC from WHTPREC2" & vbCrLf _
            & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and WH_REC_NO = '" & WH_REC_NO & "'" & vbCrLf _
            & " group by PO_SHIPMENT_NO, STYLE_CODE, COLOR_CODE"
        For Each rowWHTPREC2 As DataRow In ASCDATA1.GetDataTable.Select("")
            'sqlWhere = "PO_SHIPMENT_LNO = '" & rowWHTPREC2("PO_SHIPMENT_LNO") & "' " _
            sqlWhere = " STYLE_CODE = '" & rowWHTPREC2("STYLE_CODE") & "'" _
                & " and COLOR_CODE = '" & rowWHTPREC2("COLOR_CODE") & "'" _
                & " and PO_SHIPMENT_LNO in (" & PO_SHIPMENT_LNOs.Substring(1) & ")"
            Dim upc_lines As Integer = dst.Tables("WHTPREC3").Compute("Count(STYLE_CODE)", sqlWhere)
            Dim upc_cnt As Integer = 0
            Dim po_qty_rec As Integer = rowWHTPREC2("PO_QTY_REC")
            If upc_lines > 1 Then
                Null = Null
            End If
            For Each row As DataRow In dst.Tables("WHTPREC3").Select(sqlWhere)
                upc_cnt += 1
                RCV_OPEN = row.Item("PO_QTY_SHP") - Val(row.Item("PO_QTY_REC") & "")
                If upc_cnt = upc_lines Or RCV_OPEN >= po_qty_rec Then
                    row.Item("PO_QTY_REC") = Val(row.Item("PO_QTY_REC") & "") + po_qty_rec
                    po_qty_rec = 0
                Else
                    row.Item("PO_QTY_REC") = Val(row.Item("PO_QTY_REC") & "") + RCV_OPEN
                    po_qty_rec = po_qty_rec - RCV_OPEN
                End If
                'these next two lines are for the warehouse to see what's left to putaway from the receiving loation
                REC_LOC_QTY = Val(ASCDATA1.GetDataValue("SELECT nvl(LOCATION_QTY,0) FROM WHTLOCB1 WHERE WHSE_CODE = '" & WHSE_CODE & "' AND LOCATION_CODE = '00-RCV' and STYLE_CODE = '" & rowWHTPREC2("STYLE_CODE") & "' and COLOR_CODE = '" & rowWHTPREC2("COLOR_CODE") & "'") & "")
                row.Item("REC_LOC_QTY") = REC_LOC_QTY
                If row.Item("VARIANCE") <> 0 Then
                    Null = Null
                End If
            Next
        Next
    End Sub

    Private Sub btnRcvClose_Click(sender As Object, e As EventArgs) Handles btnRcvClose.Click
        If btnRcvClose.Text = "Close Container" Then
            CloseForReceiving()
        Else
            Print_Receiving()
        End If

    End Sub
    Sub CloseForReceiving()
        Dim PO_SHIPMENT_NO As String = Absx1.txtFor("PO_SHIPMENT_NO").Text
        Dim PO_SHIPMENT_LNO As String = grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & ""
        Dim CONTAINER_NO As String = grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO").Value & ""
        Dim PO_QTY_REC_S As Integer

        For Each row As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and CONTAINER_NO = '" & CONTAINER_NO & "'")
            PO_QTY_REC_S = dst.Tables("WHTPREC3").Compute("Sum(PO_QTY_REC)", "PO_SHIPMENT_LNO = " & row.Item("PO_SHIPMENT_LNO"))
            If PO_QTY_REC_S = 0 Then
                MsgBox("Some lines in container have no receipts, Please Verify ", MsgBoxStyle.Critical, "Verify Close Receipt")
                'Exit Sub
            End If
        Next
        'Dim PO_QTY_SHP_S As Integer = dst.Tables("WHTPREC3").Compute("Sum(PO_QTY_SHP)", "PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO)
        'PO_QTY_REC_S = dst.Tables("WHTPREC3").Compute("Sum(PO_QTY_REC)", "PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO)

        If MsgBox("Would you like to Close this Shipment for Receiving?", MsgBoxStyle.YesNo, "Flag Shipment Received") = MsgBoxResult.Yes Then

            ASCMAIN1.sql = "Update WHTWREC1 Set WH_REC_STATUS = :PARM1 where PO_SHIPMENT_NO = :PARM2 and CONTAINER_NO = :PARM3"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {"C", PO_SHIPMENT_NO, CONTAINER_NO})
            btnRcvClose.Enabled = False
            Print_Receiving()
            btnRcvClose.Text = "Reprint Report"
            btnRcvClose.Enabled = True
        End If
    End Sub
    Private Sub btnRefreshLocations_Click(sender As Object, e As EventArgs) Handles btnRefreshLocations.Click
        RefreshLocations()
        LoadWhseReceipts()
    End Sub
    Sub RefreshLocations()
        Dim PO_SHIPMENT_LNO As Integer = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
        Dim locRow As DataRow

        For Each row As DataRow In dst.Tables("WHTPREC3").Select("PO_SHIPMENT_LNO = '" & PO_SHIPMENT_LNO & "'")
            locRow = GetLocation(row.Item("STYLE_CODE"), row.Item("COLOR_CODE"))
            If Not IsNothing(locRow) AndAlso Not locRow.Item("LOCATION_CODE") Is Null Then
                row.Item("LOCATION_CODE") = locRow.Item("LOCATION_CODE")
                Dim Sql = "BEGIN " & vbCrLf _
                & " INSERT INTO WHTPREC3 (PO_SHIPMENT_NO, PO_SHIPMENT_LNO, PO_ORDER_NO, PO_ORDER_LNO, PO_QTY_SHP, PO_REC_NOTE, LOCATION_CODE)" & vbCrLf _
                & " VALUES ('" & PO_SHIPMENT_NO & "', " & PO_SHIPMENT_LNO & ",'" & row.Item("PO_ORDER_NO") & "'," & row.Item("PO_ORDER_LNO") & "," & row.Item("PO_QTY_SHP") & ",'" & row.Item("PO_REC_NOTE") & "','" & row.Item("LOCATION_CODE") & "'); " & vbCrLf _
                & " Exception" & vbCrLf _
                & "  WHEN DUP_VAL_ON_INDEX THEN" & vbCrLf _
                & "     Update WHTPREC3" & vbCrLf _
                & "     SET    PO_REC_NOTE = '" & row.Item("PO_REC_NOTE") & "'" & vbCrLf _
                & "     ,LOCATION_CODE = '" & row.Item("LOCATION_CODE") & "'" & vbCrLf _
                & "     WHERE PO_SHIPMENT_NO = '" & row.Item("PO_SHIPMENT_NO") & "'" & vbCrLf _
                & "     and PO_SHIPMENT_LNO =  " & row.Item("PO_SHIPMENT_LNO") & "" & vbCrLf _
                & "     and PO_ORDER_NO = '" & row.Item("PO_ORDER_NO") & "'" & vbCrLf _
                & "     and PO_ORDER_LNO = " & row.Item("PO_ORDER_LNO") & ";" & vbCrLf _
                & " End;"
                ASCDATA1.ExecuteSQL(Sql)
            End If
        Next
    End Sub
    Function GetLocation(ByVal Style As String, ByVal Color As String) As DataRow
        Dim rtn_row As DataRow = Nothing

        ASCMAIN1.sql = " select b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE " & vbCrLf _
            & " from whtlocb1 b1 " & vbCrLf _
            & "  join whtlocm1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE " & vbCrLf _
            & "  where b1.STYLE_CODE = '" & Style & "' and b1.COLOR_CODE = '" & Color & "' and m1.WHSE_CODE = '" & WHSE_CODE & "' " & vbCrLf _
            & "  and  nvl(m1.LOCATION_USE,'A') = 'A' and m1.LOCATION_ROUTE_SEQ is not null" & vbCrLf _
            & "  order by b1.LOCATION_QTY DESC, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            rtn_row = row
            If row("LOCATION_QTY") >= 0 Then
                Exit For
            End If
        Next

        Return rtn_row
    End Function

    Private Sub btnShowScans_Click(sender As Object, e As EventArgs) Handles btnShowScans.Click
        Dim STYLE_CODE = grdWHTPREC3.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE = grdWHTPREC3.ActiveRow.Cells("COLOR_CODE").Value & ""
        Dim PO_SHIPMENT_NO As String = Absx1.txtFor("PO_SHIPMENT_NO").Text

        If STYLE_CODE = "" Then
            MsgBox("Select a Style * Color to display scans for", vbOKOnly, "No Line Selected")
            Exit Sub
        End If

        ASCMAIN1.sql = "select GUN_ID Gun,INIT_OPER Oper,to_char(INIT_DATE,'yyyy-mm-dd hh24:mi:ss') Scan_Date_and_Time, " _
                        & " REC_KEYIN Entered,PO_QTY_REC Qty, STYLE_CODE Style,COLOR_CODE Color,UPC_CODE UPC, " _
                        & " PO_SHIPMENT_NO Shipment,PO_SHIPMENT_LNO Lno,WHSE_TRAN_NO TransNo,WHSE_TRAN_LNO TrnasLno, WH_REC_NO RecNo" _
                        & " from whtprec2 " _
                        & " where po_shipment_no = :PARM1 and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3"
        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVV", New String() {PO_SHIPMENT_NO, STYLE_CODE, COLOR_CODE})
        Using F As New ASFMSGBF
            F.Show_grd(dt, Me, "The following Scans were made for this style/Color")

        End Using
    End Sub

    Private Sub cmdFetchSummary_Click(sender As Object, e As EventArgs) Handles cmdFetchSummary.Click
        Fetch_Payments()
    End Sub

    Private Sub chkIncludeLCs_CheckedChanged(sender As Object, e As EventArgs) Handles chkIncludeLCs.CheckedChanged
        Fetch_Payments()
    End Sub

    Sub Fetch_Payments()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now creating work tables")

        Dim YP1 As String = cbeYPSFrom.Value
        Dim YP2 As String = cbeYPSTo.Value

        dst.Tables("POTSHIPP").Rows.Clear()
        dst.Tables("APTCHCKQ").Rows.Clear()
        dst.Tables("APTCHCKP").Rows.Clear()
        dst.Tables("APTCHCKv").Rows.Clear()

        If (tab0.SelectedTab.Key = "Payments") Then
            ' QUITE A BIT MORE THOUGHT NEEDS TO GO INTO SUPPORTING AMT_SHP (? MAYBE A SNAPSHOT TABLE) AND AMT_REC (ICTIREC1)

            grdAPTCHCKV.Text = "Vendors with Payments from " & cbeYPSFrom.Text & " to " & cbeYPSTo.Text
            grdAPTCHCKP.Text = "Payments from " & cbeYPSFrom.Text & " to " & cbeYPSTo.Text
            grdAPTCHCKQ.Text = "Payment Details from " & cbeYPSFrom.Text & " to " & cbeYPSTo.Text

            ASCMAIN1.sql = "" _
                & "Select 'I' TYPE, 'V' STATUS, APTINVH1.INV_AMT AMT, APTINVH1.* from APTINVH1" & vbCrLf _
                & " where APTINVH1.OPS_YYYYPP between '" & YP1 & "' and '" & YP2 & "'" & vbCrLf _
                & IIf(chkIncludeLCs.Checked, "", "   and NVL(APTINVH1.INV_PYMT_METHOD,'?') <> 'LC'" & vbCrLf) _
                & "   and APTINVH1.VEND_CODE in (Select Distinct VEND_CODE from POTORDR1)" & vbCrLf _
                & " UNION " & vbCrLf _
                & "Select 'P' TYPE, CHECK_STATUS STATUS, APTCHCK2.INV_AMT_APPLIED AMT, APTINVH1.* from APTINVH1,APTCHCK2,APTCHCK1" & vbCrLf _
                & " where APTCHCK1.OPS_YYYYPP between '" & YP1 & "' and '" & YP2 & "'" & vbCrLf _
                & "   and APTCHCK1.VEND_CODE in (Select Distinct VEND_CODE from POTORDR1)" & vbCrLf _
                & "   and APTINVH1.INV_STATUS <> 'D'" & vbCrLf _
                & "   and APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
                & "   and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM" & vbCrLf _
                & IIf(chkIncludeLCs.Checked, "", "   and NVL(APTINVH1.INV_PYMT_METHOD,'?') <> 'LC'" & vbCrLf) _
                & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO"
            Dim APTCHCKI As String = ASCMAIN1.Temp_Table

            EnforceConstraints(False)

            ASCMAIN1.sql = "Select X.*, APTVEND1.VEND_NAME from APTVEND1, (" & vbCrLf _
                & "Select APTCHCKI.VEND_CODE" & vbCrLf _
                & ", 0 AMT_SHP" & vbCrLf _
                & ", 0 AMT_REC" & vbCrLf _
                & ", Sum (Decode(TYPE,'I',AMT,0)) AMT_INV" & vbCrLf _
                & ", Sum (CASE WHEN TYPE = 'P' AND INV_TYPE = 'A' THEN AMT ELSE 0 END) AMT_ADV" & vbCrLf _
                & ", Sum (Decode(TYPE,'P',AMT,0)) AMT_PMT" & vbCrLf _
                & " from " & APTCHCKI & " APTCHCKI group by APTCHCKI.VEND_CODE" & vbCrLf _
                & ") X where APTVEND1.VEND_CODE = X.VEND_CODE "
            Fill_Records("APTCHCKV", , True, ASCMAIN1.sql)

            Sort_grdColumns(grdAPTCHCKV, "VEND_CODE")

            ASCMAIN1.sql = "Select X.*, APTVEND1.VEND_NAME from APTVEND1, (" & vbCrLf _
                & "Select APTCHCKI.CHECK_DATE, APTCHCKI.CHECK_NUM, APTCHCKI.VEND_CODE, APTCHCKI.VOUCHER_NO, APTCHCKI.INV_TYPE, APTCHCKI.INV_NUM, APTCHCKI.AMT" & vbCrLf _
                & " from " & APTCHCKI & " APTCHCKI where TYPE = 'P'" & vbCrLf _
                & ") X where APTVEND1.VEND_CODE = X.VEND_CODE "
            Fill_Records("APTCHCKP", , True, ASCMAIN1.sql)

            Sort_grdColumns(grdAPTCHCKP, "CHECK_DATE, VEND_CODE")

            ASCMAIN1.sql = "Select X.*, APTVEND1.VEND_NAME FROM APTVEND1, (" & vbCrLf _
                & "Select APTCHCKI.CHECK_DATE, APTCHCKI.CHECK_NUM, APTCHCKI.VEND_CODE, APTCHCKI.VOUCHER_NO" & vbCrLf _
                & ", APTCHCKI.INV_NUM, APTCHCKI.INV_TYPE, APTINVH5.PO_SHIPMENT_NO, POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & ", SUM (APTINVH5.INV_QTY * APTINVH5.INV_COST) AMT" & vbCrLf _
                & " from " & APTCHCKI & " APTCHCKI, APTINVH5, POTSHIP1" & vbCrLf _
                & "where APTCHCKI.TYPE = 'P' AND APTINVH5.VOUCHER_NO = APTCHCKI.VOUCHER_NO" & vbCrLf _
                & "and POTSHIP1.PO_SHIPMENT_NO = APTINVH5.PO_SHIPMENT_NO" & vbCrLf _
                & "GROUP BY APTCHCKI.CHECK_DATE, APTCHCKI.CHECK_NUM, APTCHCKI.VOUCHER_NO, APTCHCKI.VEND_CODE" & vbCrLf _
                & ", APTCHCKI.INV_NUM, APTCHCKI.INV_TYPE, APTINVH5.PO_SHIPMENT_NO, POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & "UNION" & vbCrLf _
                & "Select APTCHCKI.CHECK_DATE, APTCHCKI.CHECK_NUM, APTCHCKI.VEND_CODE, APTCHCKI.VOUCHER_NO" & vbCrLf _
                & ", APTCHCKI.INV_NUM, APTCHCKI.INV_TYPE, 'ADV' PO_SHIPMENT_NO, NULL PO_DATE_SHIPPED, NULL PO_SHIP_ETA" & vbCrLf _
                & ", APTCHCKI.AMT " & vbCrLf _
                & " from " & APTCHCKI & " APTCHCKI " & vbCrLf _
                & "where APTCHCKI.TYPE = 'P' AND APTCHCKI.INV_TYPE = 'A'" & vbCrLf _
                & "UNION " & vbCrLf _
                & "Select APTCHCKI.CHECK_DATE, APTCHCKI.CHECK_NUM, APTCHCKI.VEND_CODE, APTCHCKI.VOUCHER_NO" & vbCrLf _
                & ", APTCHCKI.INV_NUM, APTCHCKI.INV_TYPE, 'GL' PO_SHIPMENT_NO, NULL PO_DATE_SHIPPED, NULL PO_SHIP_ETA" & vbCrLf _
                & ", SUM (APTINVH2.INV_LINE_AMT) AMT" & vbCrLf _
                & " from " & APTCHCKI & " APTCHCKI,APTINVH2" & vbCrLf _
                & "where APTCHCKI.TYPE = 'P' AND APTINVH2.VOUCHER_NO = APTCHCKI.VOUCHER_NO AND APTINVH2.INV_LTYP IS NULL" & vbCrLf _
                & "GROUP BY APTCHCKI.CHECK_DATE, APTCHCKI.CHECK_NUM, APTCHCKI.VOUCHER_NO, APTCHCKI.VEND_CODE" & vbCrLf _
                & ", APTCHCKI.INV_NUM, APTCHCKI.INV_TYPE" & vbCrLf _
                & ") X WHERE APTVEND1.VEND_CODE = X.VEND_CODE"
            Fill_Records("APTCHCKQ", , True, ASCMAIN1.sql)

            Sort_grdColumns(grdAPTCHCKQ, "CHECK_DATE, VEND_CODE, VOUCHER_NO")

            EnforceConstraints(True)

            Filter_Details()

        ElseIf (tab0.SelectedTab.Key = "Open && Paid") Then

            grdAPTCHCKV.Text = "Vendors with Payments from " & cbeYPSFrom.Text & " to " & cbeYPSTo.Text
            grdPOTSHIPP.Text = "Open and Paid from " & cbeYPSFrom.Text & " to " & cbeYPSTo.Text

            ASCDATA1.ExecuteSQL("Delete from " & POTSHIPP)
            ASCMAIN1.sql = "Insert into " & POTSHIPP & " Select x.*, 0 SHP, 0 REC, 0 ACC, 0 OPN, 0 INV from (" _
                & Replace(sqlPOTSHIPP, "APTINVH1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'", "APTINVH1.OPS_YYYYPP between '" & cbeYPSFrom.Value & "' and '" & cbeYPSTo.Value & "'") & ") x"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is " & vbCrLf _
                & " Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & ", SUM (POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST) SHP" & vbCrLf _
                & ", SUM (POTSHIP3.PO_QTY_REC * POTSHIP3.PO_COST) REC" & vbCrLf _
                & ", SUM (DECODE(POTSHIPP.PO_SHIP_STATUS,'O',POTSHIP3.PO_QTY_SHP,POTSHIP3.PO_QTY_REC) * POTSHIP3.PO_COST) ACC" & vbCrLf _
                & " from POTSHIP3," & POTSHIPP & " POTSHIPP" & vbCrLf _
                & " where POTSHIP3.PO_SHIPMENT_NO = POTSHIPP.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIPP.PO_SHIPMENT_LNO" & vbCrLf _
                & " group by POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO;" & vbCrLf _
                & " Begin For R1 in C1 loop" & vbCrLf _
                & "  Update " & POTSHIPP & " POTSHIPP Set SHP = R1.SHP, REC = R1.REC, ACC = R1.ACC" & vbCrLf _
                & " , OPN = Decode(ACCRUAL_STATUS,'1',0,R1.SHP)" & vbCrLf _
                & "   where POTSHIPP.PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO" & vbCrLf _
                & "     and POTSHIPP.PO_SHIPMENT_LNO = R1.PO_SHIPMENT_LNO;" & vbCrLf _
                & " End Loop; End; " & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            '& "   and APTINVH5.VOUCHER_NO = POTSHIPP.VOUCHER_NO" & vbCrLf _
            '& "   and POTSHIPP.VOUCHER_NO is Not Null" & vbCrLf _

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is " & vbCrLf _
                & " Select POTSHIPP.PO_SHIPMENT_NO, POTSHIPP.PO_SHIPMENT_LNO" & vbCrLf _
                & ", SUM (APTINVH5.INV_QTY * APTINVH5.INV_COST) INV" & vbCrLf _
                & " from APTINVH5," & POTSHIPP & " POTSHIPP" & vbCrLf _
                & " where APTINVH5.PO_SHIPMENT_NO = POTSHIPP.PO_SHIPMENT_NO" & vbCrLf _
                & "   and APTINVH5.PO_SHIPMENT_LNO = POTSHIPP.PO_SHIPMENT_LNO" & vbCrLf _
                & " group by POTSHIPP.PO_SHIPMENT_NO, POTSHIPP.PO_SHIPMENT_LNO;" & vbCrLf _
                & " Begin For R1 in C1 loop" & vbCrLf _
                & "  Update " & POTSHIPP & " POTSHIPP Set INV = R1.INV" & vbCrLf _
                & "   where POTSHIPP.PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO" & vbCrLf _
                & "     and POTSHIPP.PO_SHIPMENT_LNO = R1.PO_SHIPMENT_LNO;" & vbCrLf _
                & " End Loop; End; " & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update " & POTSHIPP & " POTSHIPP Set VEND_CODE = (Select MIN (VEND_CODE)" & vbCrLf _
                & " from POTORDR1,POTSHIP3" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" _
                & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIPP.PO_SHIPMENT_NO" _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIPP.PO_SHIPMENT_LNO)" _
                & " where VEND_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            EnforceConstraints(False)

            Fill_Records("POTSHIPP")

            ASCMAIN1.sql = "Select X.*, APTVEND1.VEND_NAME from APTVEND1, (" & vbCrLf _
                & "Select APTCHCKI.VEND_CODE" & vbCrLf _
                & ", 0 AMT_SHP" & vbCrLf _
                & ", 0 AMT_REC" & vbCrLf _
                & ", 0 AMT_INV" & vbCrLf _
                & ", 0 AMT_ADV" & vbCrLf _
                & ", 0 AMT_PMT" & vbCrLf _
                & " from " & POTSHIPP & " APTCHCKI group by APTCHCKI.VEND_CODE" & vbCrLf _
                & ") X where APTVEND1.VEND_CODE = X.VEND_CODE "
            Fill_Records("APTCHCKV", , True, ASCMAIN1.sql)

            EnforceConstraints(True)
        End If


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Filter_Details()
        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim dvwAPTCHCKP As DataView = DirectCast(grdAPTCHCKP.DataSource, DataTable).DefaultView
        If optShowDetailsFor.Value = "A" Then
            dvwAPTCHCKP.RowFilter = ""
        Else
            dvwAPTCHCKP.RowFilter = "SEL = '1'"
        End If

        Dim dvwAPTCHCKQ As DataView = DirectCast(grdAPTCHCKQ.DataSource, DataTable).DefaultView
        If optShowDetailsFor.Value = "A" Then
            dvwAPTCHCKQ.RowFilter = ""
        Else
            dvwAPTCHCKQ.RowFilter = "SEL = '1'"
        End If

        Dim dvwPOTSHIPP As DataView = DirectCast(grdPOTSHIPP.DataSource, DataTable).DefaultView
        If optShowDetailsFor.Value = "A" Then
            dvwPOTSHIPP.RowFilter = ""
        Else
            dvwPOTSHIPP.RowFilter = "SEL = '1'"
        End If
    End Sub

    Private Sub optShowDetailsFor_ValueChanged(sender As Object, e As EventArgs) Handles optShowDetailsFor.ValueChanged
        Filter_Details()
    End Sub

    Private Sub chkShowAllColumns_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAllColumns.CheckedChanged
        ShowAllColumns()
    End Sub
    Sub ShowAllColumns()
        With grdPOTSHIPP.DisplayLayout.Bands(0)
            For Each c As String In New String() {
                "INV_DATE", "INV_STATUS", "CHECK_NUM", "PO_SHIP_VESSEL", "WHSE_CODE", "CONTAINER_NO", "BOL_NO",
                "PO_SHIP_CTNS", "OPS_YYYYPP", "PO_DATE_RECEIVED", "CONTAINER_SIZE", "VOUCHER_NO",
                "SHP", "REC", "ACC", "OPN"}
                .Columns(c).Hidden = Not chkShowAllColumns.Checked
            Next
        End With
    End Sub

    Private Sub grdATSHIPS_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdATSHIPS.DoubleClickRow
        If e.Row.IsDataRow Then
            If e.Row.Band.Index = 0 Then



                Dim SHIPDATE As Date = e.Row.Cells("ShipDate").Value
                Dim CARRIER As String = e.Row.Cells("Carrier").Value
                Dim IDX As Integer = e.Row.ListIndex

                If Not ASCMAIN1.Logical_Lock("ATSHIPS", CARRIER & Format(SHIPDATE, "yyyyMMdd")) Then
                    Exit Sub
                End If

                loading_AT = True
                rowATSHIPS = dst.Tables("ATSHIPS").Rows(IDX)
                Load_AT_Shipment(SHIPDATE, CARRIER)

                'If Not ASCMAIN1.Logical_Lock("ATSHIPS", CARRIER & Format(SHIPDATE, "yyyyMMdd")) Then
                '    Click_Command("Cancel")
                '    Exit Sub
                'End If
                loading_AT = False

            End If
        End If
    End Sub

    Sub Load_AT_Shipment(SHIPDATE As Date, CARRIER As String)

        Click_Command("New")
        If Not ScreenMode Then
            Exit Sub
        End If

        'Absx1.txtFor("PO_SHIP_VESSEL").Text = CARRIER
        'Absx1.dteFor("PO_DATE_SHIPPED").Value = SHIPDATE.Date
        rowPOTSHIP1.Item("PO_SHIP_VESSEL") = Mid(CARRIER, 1, 20)
        rowPOTSHIP1.Item("PO_DATE_SHIPPED") = SHIPDATE.Date

        ' eta?

        Dim sql As String = "ShipDate = '" & Format(SHIPDATE, "MM/dd/yyyy") & " and CARRIER = '" & CARRIER & "'"

        Dim SHIP2s As New Dictionary(Of String, Integer)
        Dim PO_SHIPMENT_LNO_ctr As Integer = 0
        Dim rowPOTSHIP2 As DataRow = Nothing

        Dim eMsgs As String = ""
        Dim msg As String = ""

        ASCMAIN1.Progress("Now Loading AT Packing", "")

        For Each rowINVHDR As DataRow In rowATSHIPS.GetChildRows("ATSHIPS_ATINVHDR")
            Dim INVNO As String = rowINVHDR.Item("INVNO")
            For Each rowPACKHDR As DataRow In rowINVHDR.GetChildRows("ATINVHDR_ATPACKHDR")
                Dim VAN_REF As String = rowPACKHDR.Item("VAN_REF")
                Dim CONTRNO As String = rowPACKHDR.Item("CONTRNO") & ""
                If CONTRNO = "" Then CONTRNO = "MISSING"
                Dim PACKREFNO As String = rowPACKHDR.Item("PACKREFNO")
                Dim BILLNO As String = rowPACKHDR.Item("BILLNO") & ""
                If BILLNO = "" Then BILLNO = "MISSING"
                Dim PO_REFERENCE As String = ""
                Dim PO_REFERENCE_modified As String = PACKREFNO

                If rowPACKHDR.GetChildRows("ATPACKHDR_ATPACKPO").Count = 0 Then
                    'PO_REFERENCE = PACKREFNO
                    msg = "unexpected - No POs related to a single packhdr - VAN_REF = " & VAN_REF
                    eMsgs &= vbCrLf & msg
                    ' Throw New Exception(msg)
                    Continue For
                Else

                    ' NADINE SAYS THIS HAPPENS ALL OF THE TIME

                    If rowPACKHDR.GetChildRows("ATPACKHDR_ATPACKPO").Count <> 1 Then

                        Dim ponos As String = ""
                        '    For Each rowPACKPO As DataRow In rowPACKHDR.GetChildRows("ATPACKHDR_ATPACKPO")
                        '        ponos &= "," & rowPACKPO.Item("pono")
                        '    Next
                        '    msg = "unexpected - IMPORTANT CALL ABS - multiple POs related to a single packhdr - VAN_REF = " & VAN_REF _
                        '        & vbCrLf & " packhdr = " & rowPACKHDR.Item("packhdrkey") & ", pono's = " & Mid(ponos, 2)
                        '    eMsgs &= vbCrLf & msg
                        '    ' Throw New Exception(msg)
                        '    Continue For
                        'Else
                        '    PO_REFERENCE = rowPACKHDR.GetChildRows("ATPACKHDR_ATPACKPO")(0).Item("PONO")
                    End If
                End If

                Dim ctnpack As New Dictionary(Of String, Int32)

                For Each rowPACKPO As DataRow In rowPACKHDR.GetChildRows("ATPACKHDR_ATPACKPO")
                    PO_REFERENCE = rowPACKPO.Item("PONO")
                    PO_REFERENCE_modified = PO_REFERENCE

                    If ASCMAIN1.Running_in_VS And PO_REFERENCE = "GB19008" Then Stop

                    If PO_REFERENCE <> PACKREFNO Then

                        ' email from lawrence 08/27/2019
                        ' Dear Walter
                        ' Please do not use  packhdr.packrefno to check against po no  in packpo.  PackRefNo is just a reference number. Usually they put pono, but not a must, since in some cases, one packing list will got multi-PO.
                        ' PackPo is the PO No 

                        ' SETTING PACKREFNO TO PO_REFERENCE FOR NOW BECAUSE AT IS SENDING BAD VALUES AT TIMES IN PACKHDR
                        ' NOTE THE IMPORTANT NOTE ABOUT MULTIPLES ABOVE - DON'T KNOW WHAT EXACTLY WE NEED TO DO IF WE ENCOUNTER MULTIPLE POS
                        PACKREFNO = PO_REFERENCE


                        'msg = "unexpected - PO Reference  " & PACKREFNO & " in packhdr not the same as PO Reference  " & PO_REFERENCE & " in packpo - VAN_REF = " & VAN_REF
                        'eMsgs &= vbCrLf & msg
                        '' Throw New Exception(msg)
                        'Continue For
                    End If

                    Dim PO_ORDER_NO As String = ""

                    Dim sqlPO As String = "Select PO_ORDER_NO from POTORDR1 where VEND_CODE = 'AT' and PO_STATUS = 'O' and PO_REFERENCE = :PARM1"

                    Dim rowPO() As DataRow = ASCDATA1.GetDataTable(sqlPO, "", "V", New Object() {PACKREFNO}).Select("")

                    If rowPO.Length = 0 Then

                        If PO_REFERENCE_modified.Length = 8 And (PO_REFERENCE_modified.EndsWith("A") Or PO_REFERENCE_modified.EndsWith("B")) Then
                            PO_REFERENCE_modified = PO_REFERENCE_modified.Substring(0, PO_REFERENCE_modified.Length - 1)
                            rowPO = ASCDATA1.GetDataTable(sqlPO, "", "V", New Object() {PO_REFERENCE_modified}).Select("")
                        End If
                    End If
                    If rowPO Is Nothing OrElse rowPO.Length < 1 Then
                        msg = "unexpected - PO Reference " & PACKREFNO & " in packhdr could not be found in Open PO File - VAN_REF = " & VAN_REF
                        eMsgs &= vbCrLf & msg
                        'Throw New Exception(msg)
                        Continue For
                    Else
                        If rowPO.Length <> 1 Then
                            'MsgBox("unexpected - PO Reference  " & PACKREFNO & " in packhdr points to multiple records in Open PO File - VAN_REF = " & VAN_REF, _
                            '       MsgBoxStyle.OkOnly, "Warning - You will have to manually include the additional POs")

                            'MsgBox("yo mama", MsgBoxStyle.OkOnly, "test")
                            'PO_ORDER_NO = rowPO(0).Item("PO_ORDER_NO")


                            ' CHANGING THIS CONDITION TO JUST RUN WITH THE 1ST PO FOUND (AS CODED ABOVE) AND ISSUE A WARNING TO NADINE THAT SHE WILL HAVE TO MANUALLY ADD THE 2ND PO

                            msg = "unexpected - PO Reference  " & PACKREFNO & " in packhdr points to multiple records in Open PO File - VAN_REF = " & VAN_REF
                            eMsgs &= vbCrLf & msg
                            ' Throw New Exception(msg)
                            Continue For
                        Else
                            PO_ORDER_NO = rowPO(0).Item("PO_ORDER_NO")
                        End If
                    End If

                    rowPOTSHIP2 = Get_Shipment_Line(INVNO, BILLNO, CONTRNO)

                    Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO"))

                    Dim COLOR_CODEs As New List(Of String)
                    For Each rowpackcarton As DataRow In rowPACKHDR.GetChildRows("ATPACKHDR_ATPACKCARTON")
                        Dim colorcode As String = rowpackcarton.Item("colorcode") & ""
                        If colorcode.StartsWith("#") Then
                            colorcode = colorcode.Substring(1)
                        End If
                        colorcode = Trim(colorcode)
                        If colorcode <> "" Then
                            If Not COLOR_CODEs.Contains(colorcode) Then
                                COLOR_CODEs.Add(colorcode)
                            End If
                        End If
                    Next
                    If COLOR_CODEs.Count = 0 Then
                        For Each rowpackbag As DataRow In rowPACKHDR.GetChildRows("ATPACKHDR_ATPACKBAG")

                            Dim pono As String = rowpackbag.Item("pono") & ""
                            If pono = "" Or pono = PO_REFERENCE_modified Then    ' If pono = "" Or pono = PO_REFERENCE Then
                                Dim colorcode As String = rowpackbag.Item("colorcode") & ""
                                If colorcode.StartsWith("#") Then
                                    colorcode = colorcode.Substring(1)
                                End If
                                colorcode = Trim(colorcode)
                                If colorcode <> "" Then
                                    If Not COLOR_CODEs.Contains(colorcode) Then
                                        COLOR_CODEs.Add(colorcode)
                                    End If
                                End If
                            End If
                        Next
                    End If

                    If Not Got_PO_Lock(PO_ORDER_NO, True) Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Clearing Shipment Data")

                        ASCMAIN1.MultiTask_Release()
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Clearing Shipment Data")
                        Click_Command("Cancel")

                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")

                        Exit Sub
                    End If

                    'If PO_ORDER_NO = "140617" Then Stop

                    Load_POs_into_POTSHIP3(PO_ORDER_NO, False, PO_SHIPMENT_LNO, COLOR_CODEs)

                    'If ASCMAIN1.Running_in_VS AndAlso (PO_ORDER_NO = "135247" Or PO_ORDER_NO = "135246") Then Stop

                    Dim CARTON_NO As Integer = Val(dst.Tables("POTSHIP7").Compute("MAX(CARTON_NO)", "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)) & "")
                    'Dim msg2 As String = Load_Invoice_Packing(rowPACKHDR, PO_REFERENCE, PO_SHIPMENT_LNO, False, CARTON_NO, COLOR_CODEs, ctnpack)
                    Dim msg2 As String = Load_Invoice_Packing(rowPACKHDR, PO_REFERENCE_modified, PO_SHIPMENT_LNO, False, CARTON_NO, COLOR_CODEs, ctnpack, INVNO)
                    If msg2 <> "" Then
                        eMsgs &= vbCrLf & msg2
                    End If
                Next
            Next

        Next

        ASCMAIN1.Progress("", "")

        For Each rowPOTSHIP2 In dst.Tables("POTSHIP2").Select("")
            Dim PO_SHIPMENT_LNO As Int32 = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
            rowPOTSHIP2.Item("PO_SHIP_CTNS") = Val(dst.Tables("POTSHIP7").Compute("SUM(CARTONS)", "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)) & "")
        Next


        Create_Containers_from_BOL()

        AT_Packing_Errors = eMsgs
        AT_Packing = True

        tabBOL.Tabs("Packing Discrepancies").Visible = True

        If eMsgs <> "" Then
            MsgBox(eMsgs, MsgBoxStyle.OkOnly, "Errors Encountered during Processing - Update Disabled")
        End If


        If eMsgs = "" Then
            'For Each rowPOTSHIPR As DataRow In dst.Tables("POTSHIPR").Select("QTY_VAR <> 0")
            '    Dim QTY_CTN As Integer = Val(rowPOTSHIPR.Item("QTY_CTN") & "")
            '    Dim foundit As Boolean = False
            '    For Each rowPOTSHIP3 As DataRow In rowPOTSHIPR.GetChildRows("POTSHIPR_POTSHIP3")

            '        Dim PO_ORDER_NO As String = rowPOTSHIP3.Item("PO_ORDER_NO")
            '        'If PO_ORDER_NO = "140617" Then Stop

            '        Dim PO_QTY_SHP As Integer = Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "")
            '        If Not foundit AndAlso PO_QTY_SHP = QTY_CTN Then
            '            foundit = True
            '        Else
            '            ' NEED TO DO THIS BY PO
            '            If foundit Then
            '                rowPOTSHIP3.Item("PO_QTY_SHP") = 0
            '            End If
            '            ' rowPOTSHIP3.Item("PO_QTY_SHP") = 0
            '        End If
            '    Next
            'Next



            Dim PACKS As New Dictionary(Of String, Int64)
            For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("PO_QTY_SHP <> PO_QTY_PCK")
                Dim POSC As String = rowPOTSHIP3.Item("PO_ORDER_NO") & ":" & rowPOTSHIP3.Item("STYLE_CODE") & ":" & rowPOTSHIP3.Item("COLOR_CODE")
                Dim PO_QTY_PCK As Int64 = Val(rowPOTSHIP3.Item("PO_QTY_PCK") & "")
                If PACKS.ContainsKey(POSC) Then
                    PACKS(POSC) += PO_QTY_PCK
                Else
                    PACKS.Add(POSC, PO_QTY_PCK)
                End If
                rowPOTSHIP3.Item("PO_QTY_SHP") = 0
            Next

            For Each POSC As String In PACKS.Keys
                Dim PO_ORDER_NO As String = Split(POSC, ":")(0)
                Dim STYLE_CODE As String = Split(POSC, ":")(1)
                Dim COLOR_CODE As String = Split(POSC, ":")(2)
                Dim PO_QTY_PCK As Int64 = PACKS(POSC)
                Dim sqlw As String = $"PO_ORDER_NO = '{PO_ORDER_NO}' and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' and PO_QTY_SHP = 0 and PO_QTY_OPN = {CStr(PO_QTY_PCK)}"
                Dim rows() As DataRow = dst.Tables("POTSHIP3").Select(sqlw)
                If rows.Length = 1 Then
                    rows(0).Item("PO_QTY_SHP") = PO_QTY_PCK
                End If

            Next



        End If
    End Sub

    Function Load_Invoice_Packing(rowpackhdr As DataRow,
                                  PO_REFERENCE As String,
                                  PO_SHIPMENT_LNO As Integer,
                                  isPPK As Boolean,
                                  ByRef CARTON_NO As String,
                                  COLOR_CODEs As List(Of String),
                                  ctnpack As Dictionary(Of String, Int32), INVNO As String) As String

        Dim msg As String = ""

        Dim packhdrkey As Integer = Val(rowpackhdr.Item("packhdrkey"))

        Dim grosswgt As Decimal = Val(rowpackhdr.Item("grosswgt"))
        Dim packcartondimensions As String = rowpackhdr.Item("dim") & ""
        packcartondimensions = Replace(packcartondimensions, "*", "x")

        If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE = "ME22243" Or PO_REFERENCE = "MX220419") Then Stop

        Dim sqlw As String = "packhdrkey = " & CStr(packhdrkey) & " and ISNULL(pono,'" & PO_REFERENCE & "') = '" & PO_REFERENCE & "'"
        Dim packbag_qty As Integer = Val(dst.Tables("ATPACKBAG").Compute("Sum(qty)", sqlw) & "")
        If packbag_qty = 0 Then packbag_qty = 1 ' see ME19052, VAN_REF 0000001474 - no packbag rows for invoice I-7470-19

        'If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE = "ME20236" Or PO_REFERENCE = "ME20238") Then Stop
        '        If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE = "MX20833" Or PO_REFERENCE = "MX20835") Then Stop

        Dim SIZEs As String = ""
        Dim QTYs As String = ""
        Dim STYLE_CODEs As String = ""
        Dim STYLE_CODEs_in_packbag As New List(Of String)

        For Each rowpackbag As DataRow In rowpackhdr.GetChildRows("ATPACKHDR_ATPACKBAG") ' For Each rowpackbag As DataRow In packbag.Select(sqlw, "packbagkey")
            Dim pono As String = rowpackbag.Item("pono") & ""
            If pono = "" Or pono = PO_REFERENCE Then
                Dim packsize As String = rowpackbag.Item("size") & ""
                Dim packqty As Integer = Val(rowpackbag.Item("qty"))

                Dim styleno As String = rowpackbag.Item("styleno") & ""
                ' need to know if packbag has 1 styleno or several - see ME22243 vs MX220419
                If Not STYLE_CODEs_in_packbag.Contains(styleno) Then
                    STYLE_CODEs_in_packbag.Add(styleno)
                End If

                If packsize <> "" Or styleno <> "" Then
                    SIZEs &= "/" & Trim(packsize)
                    QTYs &= "/" & CStr(packqty)
                    ' If styleno <> "" Then STYLE_CODEs &= "/" & Trim(styleno)
                    ' MAKING THE FOLLOWING CHANGE TO IMPORT MX20444/MX20445
                    ' BECAUSE THE SAME STYLE WAS ADDED 4X BECAUSE THERE WERE 4 DIFFERENT SIZES, 
                    ' And THIS CAUSED THE QTY PACKED TO * 4
                    If styleno <> "" Then
                        If Not (STYLE_CODEs & "/").Contains("/" & Trim(styleno) & "/") Then
                            STYLE_CODEs &= "/" & Trim(styleno)
                        End If
                    End If

                End If
            End If
        Next

        If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE = "MX20832" Or PO_REFERENCE = "MX20833" Or PO_REFERENCE = "MX20834" Or PO_REFERENCE = "MX20835") Then Stop

        Dim total_cartons As Integer = 0
        Dim row7s As New List(Of DataRow)

        '  If packhdrkey = 10136 Then Stop

        Dim TOTAL_PCS_PACKED_by_color As New Dictionary(Of String, Int32)
        For Each rowpackcarton As DataRow In rowpackhdr.GetChildRows("ATPACKHDR_ATPACKCARTON")
            Dim ctnfrom As Integer = Val(rowpackcarton.Item("ctnfrom"))
            Dim ctnto As Integer = Val(rowpackcarton.Item("ctnto"))
            Dim bagperctn As Integer = Val(rowpackcarton.Item("bagperctn"))
            Dim bagqty As Integer = Val(rowpackcarton.Item("bagqty") & "")
            If bagqty = 0 Then bagqty = 1 ' new since ME19052
            'If bagqty = 0 Then
            '    If packbag_qty > 0 Then
            '        bagqty = packbag_qty ' new since ME22243
            '    Else
            '        bagqty = 1 ' new since ME19052
            '    End If
            'End If

            Dim CARTONS As Int32 = 0
            If ctnto = 0 Then
                CARTONS = 1
            Else
                CARTONS = ctnto - ctnfrom + 1
            End If
            Dim TOTAL_PCS As Int32 = packbag_qty * bagperctn * bagqty * CARTONS

            Dim colorcode As String = rowpackcarton.Item("colorcode") & ""
            If Not TOTAL_PCS_PACKED_by_color.ContainsKey(colorcode) Then
                TOTAL_PCS_PACKED_by_color.Add(colorcode, 0)
            End If

            TOTAL_PCS_PACKED_by_color(colorcode) += TOTAL_PCS
        Next

        Dim PO_LINE_for_color As New Dictionary(Of String, Int32)

        Dim sizes_used_as_styles As Boolean = False

        For Each rowpackcarton As DataRow In rowpackhdr.GetChildRows("ATPACKHDR_ATPACKCARTON") 'For Each rowpackcarton As DataRow In packcarton.Select(sqlw, "packcartonkey")
            Dim ctnfrom As Integer = Val(rowpackcarton.Item("ctnfrom"))
            Dim ctnto As Integer = Val(rowpackcarton.Item("ctnto"))
            Dim bagperctn As Integer = Val(rowpackcarton.Item("bagperctn"))
            Dim bagqty As Integer = Val(rowpackcarton.Item("bagqty"))
            If bagqty = 0 Then bagqty = 1 ' new since ME19052

            'If bagqty = 0 Then
            '    If packbag_qty > 0 Then
            '        bagqty = packbag_qty ' new since ME22243
            '    Else
            '        bagqty = 1 ' new since ME19052
            '    End If
            'End If

            Dim cartonwgt As Decimal = Val(rowpackcarton.Item("cartonwgt"))

            Dim packcartonkey As Integer = Val(rowpackcarton.Item("packcartonkey"))
            'If ASCMAIN1.Running_in_VS And packcartonkey = 38374 Then Stop
            isPPK = False

            Dim STYLE_CODE_by_size As String = ""
            Dim size As String = rowpackcarton.Item("size") & ""
            Dim styleno As String = rowpackcarton.Item("styleno") & ""
            If styleno = "" And size <> "" Then
                ' they place the VAN style code in the size field
                STYLE_CODE_by_size = size
            ElseIf styleno <> "" Then
                STYLE_CODE_by_size = styleno
                If size <> "" And size.StartsWith(styleno) Then
                    ' use the size field if both size nd styleno are not null and size starts with styleno
                    STYLE_CODE_by_size = size
                End If
            End If


            Dim CARTONS As Int32 = 0
            If ctnto = 0 Then
                CARTONS = 1
            Else
                CARTONS = ctnto - ctnfrom + 1
            End If

            Dim TOTAL_PCS As Int32 = packbag_qty * bagperctn * bagqty * CARTONS

            Dim CARTON_PACK_QTY As Int32 = packbag_qty * bagperctn * bagqty

            Dim colorcode As String = rowpackcarton.Item("colorcode") & ""
            Dim COLOR_CODE As String = colorcode
            Dim first_time_for_colorcode As Boolean = Not PO_LINE_for_color.ContainsKey(colorcode)

            If COLOR_CODE.StartsWith("#") Then
                COLOR_CODE = COLOR_CODE.Substring(1)
            End If
            COLOR_CODE = Trim(COLOR_CODE)

            Dim rowPOTSHIP3 As DataRow = Nothing
            Dim sqlpo As String = "PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and PO_REFERENCE = '" & PO_REFERENCE & "'"
            If COLOR_CODE <> "" Then
                sqlpo &= " and COLOR_CODE = '" & COLOR_CODE & "'"
            Else
                'Dim sqlpo_cc As String = ""
                'For Each cc As String In COLOR_CODEs
                '    sqlpo &= " or PO_ORDER_LNO = " & CStr(PO_LINE_for_color(cc))
                'Next
                'sqlpo &= " and (" & Mid(sqlpo_cc, 5) & ")"
            End If


            If Not first_time_for_colorcode AndAlso PO_LINE_for_color(colorcode) <> 0 Then ' this line is probably not good for color prepacks
                ' need to check this out for splits, where a PO is shipped on 2 different dates, probably for both color prepacks as well as regular
                sqlpo &= " and PO_ORDER_LNO = " & CStr(PO_LINE_for_color(colorcode))
            End If

            'If colorcode = "" Then ' MULTIPLE COLORS IN A SINGLE CARTON
            '    Dim sqlpo_cc As String = ""
            '    For Each cc As String In COLOR_CODEs
            '        sqlpo &= " or PO_ORDER_LNO = " & CStr(PO_LINE_for_color(cc))
            '    Next
            '    sqlpo &= " and (" & Mid(sqlpo_cc, 5) & ")"
            'Else
            '    If Not first_time_for_colorcode AndAlso PO_LINE_for_color(colorcode) <> 0 Then

            '        sqlpo &= " and PO_ORDER_LNO = " & CStr(PO_LINE_for_color(colorcode))
            '    End If
            'End If


            Dim STYLEs() As String



            If STYLE_CODEs = "" And SIZEs <> "" Then
                ' ME20236 - SIZE PREPACK WHERE 4 SIZES WERE IN A SINGLE BAG AND EACH SIZE HAD ITS OWN STYLE CODE
                ' BUT DO THIS ONLY IF ALL OF THE "SIZES" ARE > 6 CHARACTERS - BECAUSE LB20228 HAD REAL SIZES AND NOT STYLES WITH SIZES IN PACKBAG
                Dim STYLEs_MAYBE() As String = Split(Mid(SIZEs, 2), "/")
                Dim do_all_sizes_look_like_styles As Boolean = (STYLEs_MAYBE.Length) > 0
                For Each style_with_size As String In STYLEs_MAYBE
                    If Len(style_with_size) <= 6 Then
                        do_all_sizes_look_like_styles = False
                    End If
                Next
                If do_all_sizes_look_like_styles Then
                    STYLE_CODEs = SIZEs
                End If
                'If ctnfrom = 1 And do_all_sizes_look_like_styles Then sizes_used_as_styles = True
                ' Length = 1 clause added to handle ME20265 on 01/30, which had 1 packcarton record, from 18 to 18
                ' i am still not sure whether this logic is sound
                ' If (ctnfrom = 1 Or rowpackhdr.GetChildRows("ATPACKHDR_ATPACKCARTON").Length = 1) And do_all_sizes_look_like_styles Then sizes_used_as_styles = True
                If do_all_sizes_look_like_styles Then sizes_used_as_styles = True
            End If

            'If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE = "ME20236" Or PO_REFERENCE = "ME20238" Or PO_REFERENCE = "LB20194") Then Stop
            'If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE.StartsWith("ME20264") Or PO_REFERENCE = "ME20266") Then Stop
            'If ASCMAIN1.Running_in_VS AndAlso (PO_REFERENCE.StartsWith("ME20280") Or PO_REFERENCE = "ME20282") Then Stop

            'Dim sqlpo2 As String = ""
            If STYLE_CODE_by_size <> "" And Not sizes_used_as_styles Then
                sqlpo &= " and (STYLE_CODE = '" & STYLE_CODE_by_size & "')"
                'sqlpo2 = " and (STYLE_CODE = '" & STYLE_CODE_by_size & "')"


                'sqlpo &= " and (STYLE_CODE = '" & STYLE_CODE_by_size & "'"
                'sqlpo &= "  or  STYLE_CODE = '" & Replace(STYLE_CODE_by_size, "-", "") & "')"
            End If


            If STYLE_CODEs = "" Then
                ReDim STYLEs(0)
            Else
                STYLEs = Split(Mid(STYLE_CODEs, 2), "/")
            End If

            Dim TOTAL_PCS_bagStyle As Integer = 0
            Dim bagqtys() As String = Split(Mid(QTYs, 2), "/")

            Dim bagStyleNo As Integer = 0
            For Each STYLE_CODE_in_bag As String In STYLEs
                '     If STYLE_CODE_in_bag = "JS51256BMX" Then
                '    Stop
                '   End If
                bagStyleNo += 1
                TOTAL_PCS_bagStyle = TOTAL_PCS
                If packbag_qty <> 0 And SIZEs <> "" And QTYs <> "" And (sizes_used_as_styles Or (packbag_qty > 1 And STYLE_CODEs_in_packbag.Count > 1)) Then
                    'If packbag_qty <> 0 And SIZEs <> "" And QTYs <> "" And sizes_used_as_styles Then
                    TOTAL_PCS_bagStyle = TOTAL_PCS * bagqtys(bagStyleNo - 1) / packbag_qty
                End If

                Dim sqlPO_STYLE_CODE As String = ""
                If STYLE_CODE_in_bag <> "" Then
                    sqlPO_STYLE_CODE &= " and STYLE_CODE = '" & STYLE_CODE_in_bag & "'"
                Else
                    'sqlPO_STYLE_CODE &= sqlpo2
                End If


                Dim rowPOTSHIP3s() As DataRow = dst.Tables("POTSHIP3").Select(sqlpo & sqlPO_STYLE_CODE)
                If rowPOTSHIP3s.Length = 0 Then
                    msg = $"Could not Find Color {colorcode} in Inv {INVNO} PO {PO_REFERENCE} for Style {IIf(STYLE_CODE_in_bag = "", STYLE_CODE_by_size, STYLE_CODE_in_bag)}"

                    Return msg
                Else

                    If colorcode = "" Then ' deal with multiple colors in a single carton

                    Else

                    End If

                    Dim QTY_PACKED_KEY As String = CStr(PO_SHIPMENT_LNO) & ":" & rowPOTSHIP3s(0).Item("STYLE_CODE") & ":" & COLOR_CODE
                    If Not QTY_PACKED.ContainsKey(QTY_PACKED_KEY) Then
                        QTY_PACKED.Add(QTY_PACKED_KEY, 0)
                    End If
                    QTY_PACKED(QTY_PACKED_KEY) += TOTAL_PCS_bagStyle

                    If rowPOTSHIP3s.Length = 1 Then
                        rowPOTSHIP3 = rowPOTSHIP3s(0)
                    ElseIf rowPOTSHIP3s.Length = 2 Or rowPOTSHIP3s.Length = 3 Then
                        rowPOTSHIP3 = rowPOTSHIP3s(0)
                        For Each rowx As DataRow In rowPOTSHIP3s
                            If rowx.Item("PO_QTY_SHP") = TOTAL_PCS_bagStyle Then
                                rowPOTSHIP3 = rowx
                                Exit For
                            End If
                        Next



                    Else

                        If colorcode = "" Then ' this really should be looking at a boolean like isPPK, where isPPK means color prepack
                            Dim qtyMatched As Boolean = False
                            For Each rowMatchQty As DataRow In rowPOTSHIP3s
                                If rowMatchQty.Item("PO_QTY_SHP") = TOTAL_PCS_bagStyle Then
                                    rowPOTSHIP3 = rowMatchQty
                                    qtyMatched = True
                                    Exit For
                                End If
                            Next
                            If Not qtyMatched Then
                                rowPOTSHIP3 = rowPOTSHIP3s(0) ' just spitballing here - I am sure we will revisit this when we get split POs for color prepacks
                            End If
                        Else

                            ' NADINE MAY COME BACK TO SHOW THAT WE PICKED THE WRONG ONE
                            For Each rowPOTSHIP3matchqty As DataRow In rowPOTSHIP3s
                                Dim PO_QTY_SHP As Int32 = Val(rowPOTSHIP3matchqty.Item("PO_QTY_SHP") & "")
                                If PO_QTY_SHP = TOTAL_PCS_bagStyle Or (first_time_for_colorcode And PO_QTY_SHP = TOTAL_PCS_PACKED_by_color(colorcode)) Then
                                    rowPOTSHIP3 = rowPOTSHIP3matchqty

                                    If first_time_for_colorcode And PO_QTY_SHP = TOTAL_PCS_PACKED_by_color(colorcode) Then
                                        PO_LINE_for_color.Add(colorcode, rowPOTSHIP3.Item("PO_ORDER_LNO"))
                                    End If

                                    If PO_REFERENCE = "ME19158" Then
                                        ' SKIP THIS
                                    Else
                                        For Each rowPOTSHIP3_zero As DataRow In rowPOTSHIP3s
                                            If rowPOTSHIP3_zero.Item("PO_ORDER_LNO") <> rowPOTSHIP3matchqty.Item("PO_ORDER_LNO") Then
                                                rowPOTSHIP3_zero.Item("PO_QTY_SHP") = 0
                                            End If
                                        Next
                                    End If

                                    Exit For
                                End If
                            Next
                        End If

                        If rowPOTSHIP3 IsNot Nothing Then
                            ' found one
                        Else
                            msg &= vbCrLf & "Multiple Instances of Color Code " & COLOR_CODE & " in packhdrkey " & CStr(packhdrkey) & ", PO Reference " & PO_REFERENCE & vbCrLf & "Style " & rowPOTSHIP3s(0).Item("STYLE_CODE") & ", PO " & rowPOTSHIP3s(0).Item("PO_ORDER_NO") & ", Packed " & CStr(TOTAL_PCS)
                            '  Return msg
                        End If
                    End If
                End If


                If Not PO_LINE_for_color.ContainsKey(colorcode) Then
                    PO_LINE_for_color.Add(colorcode, 0)
                End If

                If rowPOTSHIP3 IsNot Nothing Then
                    Dim STYLE_CODE As String = rowPOTSHIP3.Item("STYLE_CODE")
                    total_cartons += CARTONS
                    Dim rowPOTSHIP7 As DataRow = Nothing

                    Dim ctnpackrange As String = Format(ctnfrom, "000000") & Format(ctnto, "000000")

                    If ctnpack.ContainsKey(ctnpackrange) Then
                        isPPK = True
                    End If

                    rowPOTSHIP3.Item("PO_QTY_PCK") = Val(rowPOTSHIP3.Item("PO_QTY_PCK")) + TOTAL_PCS_bagStyle

                    If Not isPPK Then
                        rowPOTSHIP7 = dst.Tables("POTSHIP7").NewRow()
                        rowPOTSHIP7.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        rowPOTSHIP7.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                        CARTON_NO += 1


                        ctnpack.Add(ctnpackrange, CARTON_NO)

                        rowPOTSHIP7.Item("CARTON_NO") = CARTON_NO

                        rowPOTSHIP7.Item("CARTON_WEIGHT") = grosswgt / CARTONS ' s/b cartonwgt
                        rowPOTSHIP7.Item("CARTON_DIMS") = packcartondimensions ' s.b cartondimensions
                        If packcartondimensions <> "" Then
                            rowPOTSHIP7.Item("CARTON_VOLUME") = Get_Volume_from_Dims(packcartondimensions)
                        End If
                        rowPOTSHIP7.Item("CARTONS") = CARTONS
                        dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
                        row7s.Add(rowPOTSHIP7)
                    Else
                        CARTON_PACK_QTY = TOTAL_PCS_bagStyle / CARTONS
                    End If

                    ' maybe we don't need the else above?
                    CARTON_PACK_QTY = TOTAL_PCS_bagStyle / CARTONS

                    Dim COLOR_CODEs_Packed As New List(Of String)
                    If colorcode <> "" Then
                        COLOR_CODEs_Packed.Add(COLOR_CODE)
                    ElseIf COLOR_CODEs.Count = 1 Then ' 08/31 SHIPMENT INVOICE I-7407-19
                        COLOR_CODEs_Packed.Add(COLOR_CODEs(0))
                    Else

                        For Each CC As String In COLOR_CODEs
                            COLOR_CODEs_Packed.Add(CC)
                        Next
                        ' needed to divide by color_codes.count twice to get the 7/27 shipment to load
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        Dim SUB_UNIT_PACK_QTY As Integer = Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "")
                        If SUB_UNIT_PACK_QTY = 0 Then SUB_UNIT_PACK_QTY = 1

                        CARTON_PACK_QTY = TOTAL_PCS_bagStyle / CARTONS / COLOR_CODEs.Count / SUB_UNIT_PACK_QTY
                        ' the carton_pack_qty maybe needs to be figured out by looking at packbag
                    End If

                    For Each COLOR_CODE_PACKED As String In COLOR_CODEs_Packed

                        Dim rowPOTSHIP8 As DataRow = Nothing
                        If isPPK Then
                            CARTON_NO = ctnpack(ctnpackrange)
                            rowPOTSHIP8 = dst.Tables("POTSHIP8").Rows.Find _
                                (New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, CARTON_NO, STYLE_CODE, COLOR_CODE_PACKED})
                        End If

                        If rowPOTSHIP8 IsNot Nothing Then
                            rowPOTSHIP8.Item("QTY") = Val(rowPOTSHIP8.Item("QTY") & "") + CARTON_PACK_QTY
                        Else
                            '  If STYLE_CODE = "JS51256BMX" Then
                            ' Stop
                            'End If

                            rowPOTSHIP8 = dst.Tables("POTSHIP8").NewRow
                            rowPOTSHIP8.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            rowPOTSHIP8.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO

                            rowPOTSHIP8.Item("CARTON_NO") = CARTON_NO
                            rowPOTSHIP8.Item("STYLE_CODE") = STYLE_CODE
                            rowPOTSHIP8.Item("COLOR_CODE") = COLOR_CODE_PACKED
                            rowPOTSHIP8.Item("QTY") = CARTON_PACK_QTY
                            dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)
                        End If
                    Next
                End If
            Next
        Next

        For Each rowPOTSHIP7 As DataRow In row7s
            rowPOTSHIP7.Item("CARTON_WEIGHT") = grosswgt / total_cartons
        Next
        'Next

        Return msg
    End Function

    Private Sub grdPOTSHPIE_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdPOTSHPIE.ClickCellButton
        Dim COL As String = e.Cell.Column.Key
        Dim POX As String = COL.Substring(COL.Length - 1, 1)

        Dim QTY_OPEN As Int64 = Val(e.Cell.Row.Cells("QTY_OPEN_PO_ORDER_NO" & POX).Value & "")
        Dim QTY_OPEN_THIS_PO As Int64 = Val(e.Cell.Row.Cells("QTY_OPEN_THIS_PO").Value & "")
        Dim QTY_NEEDED As Int64 = Val(e.Cell.Row.Cells("QTY_NEEDED").Value & "")

        Dim PO_ORDER_NO As String = e.Cell.Row.Cells("PO_ORDER_NO").Value & ""
        Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = e.Cell.Row.Cells("COLOR_CODE").Value & ""

        If QTY_OPEN <= 0 And QTY_NEEDED > 0 Then
            MsgBox("Nothing to Steal", MsgBoxStyle.OkOnly, "Cannot Figure Out what to Steal")
        Else
            Dim PO_ORDER_NOX As String = e.Cell.Value & ""
            Dim QTY_OPEN_PO_ORDER_NOX As Int64 = Val(e.Cell.Row.Cells("QTY_OPEN_PO_ORDER_NO" & POX).Value & "")

            ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = :PARM1 and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3 and PO_QTY_OPN = :PARM4"
            Dim rowPO As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVVN", New Object() {PO_ORDER_NO, STYLE_CODE, COLOR_CODE, QTY_OPEN_THIS_PO})
            Dim rowPOX As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVVN", New Object() {PO_ORDER_NOX, STYLE_CODE, COLOR_CODE, QTY_OPEN})

            If rowPO Is Nothing Or rowPOX Is Nothing Then
                MsgBox("Cannot do a Simple Open Qty Transfer between these 2 POs", MsgBoxStyle.OkOnly, "Please move the Open PO Qty Manually")
            Else
                If MsgBox("OK to Steal " & CStr(QTY_NEEDED) & " from PO " & PO_ORDER_NOX & vbCrLf & " and add to PO " & PO_ORDER_NO & "?",
          MsgBoxStyle.YesNo, "Please Confirm Action") = MsgBoxResult.Yes Then

                    BeginTrans()
                    ASCMAIN1.sql = "Update POTORDR2 Set PO_QTY_ORD = PO_QTY_ORD + :PARM1 where PO_ORDER_NO = :PARM1 and PO_ORDER_LNO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NVN", New Object() {QTY_NEEDED, PO_ORDER_NO, Val(rowPO.Item("PO_ORDER_LNO") & "")})
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NVN", New Object() {-1 * QTY_NEEDED, PO_ORDER_NOX, Val(rowPO.Item("PO_ORDER_LNO") & "")})
                    CommitTrans()
                End If
            End If

        End If

    End Sub

    Private Sub grdPOTSHPIE_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTSHPIE.InitializeRow
        Dim PO_ORDER_NO1 As String = e.Row.Cells("PO_ORDER_NO1").Value & ""
        If PO_ORDER_NO1 <> "" Then
            e.Row.Cells("PO_ORDER_NO1").ToolTipText = "Click to Steal from " & PO_ORDER_NO1
        Else
            e.Row.Cells("PO_ORDER_NO1").ToolTipText = ""
        End If
        Dim PO_ORDER_NO2 As String = e.Row.Cells("PO_ORDER_NO2").Value & ""
        If PO_ORDER_NO1 <> "" Then
            e.Row.Cells("PO_ORDER_NO2").ToolTipText = "Click to Steal from " & PO_ORDER_NO2
        Else
            e.Row.Cells("PO_ORDER_NO2").ToolTipText = ""
        End If
    End Sub

    Private Sub btnPlusDuty_Click(sender As Object, e As EventArgs) Handles btnPlusDuty.Click
        If Not ScreenMode Or EntryMode <> "E" Then
            Exit Sub
        End If

        Dim DP As Decimal = Val(numPlusDuty.Value & "")
        If DP = 0 Then
            MsgBox("Cannot add $0 Tariff", MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        'For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
        '    Dim PO_COST As Decimal = Val(rowPOTSHIP3.Item("PO_COST") & "")
        '    Dim PO_COST_QUOTA As Decimal = PO_COST * DP / 100
        '    rowPOTSHIP3.Item("PO_COST_QUOTA") = PO_COST_QUOTA
        'Next

        Calculate_Landed_Cost()
        Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and LANDING_COST_DIST = 'D' and COST_CATGY_CODE = 'TARIFF' and CTL_NO is Null"
        Dim rowPOTSHIP5s() As DataRow = dst.Tables("POTSHIP5").Select(sqlw)
        Dim rowPOTSHIP5 As DataRow = Nothing
        If rowPOTSHIP5s.Length = 0 Then
            rowPOTSHIP5 = dst.Tables("POTSHIP5").NewRow
            rowPOTSHIP5.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP5.Item("PO_SHIPMENT_LNO") = Val(dst.Tables("POTSHIP5").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
            rowPOTSHIP5.Item("COST_CATGY_CODE") = "TARIFF"
            rowPOTSHIP5.Item("COST_CATGY_DESC") = "Tariff (Additional Duty)"
            rowPOTSHIP5.Item("LANDING_COST_DIST") = "D"
            dst.Tables("POTSHIP5").Rows.Add(rowPOTSHIP5)
        Else
            rowPOTSHIP5 = rowPOTSHIP5s(0)
        End If

        Dim FIRST_COST As Decimal = 0
        For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
            Dim PO_COST As Decimal = Val(rowPOTSHIP3.Item("PO_COST") & "")
            Dim PO_QTY_SR As Int64 = Val(rowPOTSHIP3.Item("PO_QTY_SR") & "")
            FIRST_COST += PO_COST * PO_QTY_SR
        Next

        Dim TARIFF As Decimal = FIRST_COST * DP / 100
        'rowPOTSHIP5.Item("LANDING_COST_AMT") = Val(rowPOTSHIP5.Item("LANDING_COST_AMT") & "") + TARIFF
        rowPOTSHIP5.Item("LANDING_COST_AMT") = TARIFF
        Calculate_Landed_Cost()

        MsgBox("Tariff added to Duty", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Sub Transmit_Discrepancies(PO_SHIPMENT_NO As String, Optional email_to_myself As Boolean = False)


        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim REPORT_NAME As String = "PORDISC1"
        Print_Report_Begin()
        Dim SUBT As String = "Packing Discrepancies for Vessel: " & Absx1.txtFor("PO_SHIP_VESSEL").Text & "  Ship Date: " & Absx1.dteFor("PO_DATE_SHIPPED").Value
        CR_params.Add("SUBT", SUBT)
        Dim REPORT_NO As String = Generate_Report(REPORT_NAME, "Discrepancy Print Out", "",, "PDF",, False)
        Print_Report_End(True)
        ATTACHMENTs.Add(REPORT_NO & ".pdf", ASCMAIN1.Folders("Temp") & ASCMAIN1.CLIENT & "_" & REPORT_NO & ".pdf")

        Dim SUBJECT As String = "Packing Discrepancies for Vessel: " & Absx1.txtFor("PO_SHIP_VESSEL").Text & "  Ship Date: " & Absx1.dteFor("PO_DATE_SHIPPED").Value
        Dim PFX As String = ""

        Dim SEND_CC_to_USER_ID As Boolean = True

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        If ASCMAIN1.Running_in_VS Then
            EMAIL_ADDRESSs.Add("dgj@absolution.com", "Darrin Joscelyn")
        Else
            EMAIL_ADDRESSs.Add("jtrinh@vandale.com", "Joann Trinh")
            EMAIL_ADDRESSs.Add("asiegel@vandale.com", "Annamaria Siegel")
            EMAIL_ADDRESSs.Add("humbach@vandale.com", "Hui Umbach")
        End If

        If email_to_myself Then
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
            SEND_CC_to_USER_ID = False
            'Else
            '    EMAIL_ADDRESSs.Add("dgj@absolution.com", "Darrin Joscelyn")
        End If

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, "PORDISC1", True, SEND_CC_to_USER_ID, "", "", "Supplier")

        If SEND_NO <> "" Then
            MsgBox("email has been sent", MsgBoxStyle.OkOnly, "Verification")
        End If

    End Sub

    Private Sub cmdDiscrepancy_Click(sender As Object, e As EventArgs) Handles cmdDiscrepancy.Click
        Transmit_Discrepancies(PO_SHIPMENT_NO, True) ' eItemKey = "email to myself")
    End Sub

    Sub Import_Bookings()

        packingFromBooking = False

        For Each TABLE_NAME As String In New String() {"POTVBKG1", "POTVBKG2", "POTVBKG3", "POTPACK2", "POTPACK3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EntryMode = "N"
        Load_Record()

        Dim PO_SPLITS As New Dictionary(Of String, Dictionary(Of Integer, Integer))
        eMsg_Booking = ""
        For Each grow As UltraWinGrid.UltraGridRow In grdPOTVBKGX.Selected.Rows
            Dim VBKG_NO As String = grow.Cells("VBKG_NO").Value & ""
            Book2ShiP(VBKG_NO, PO_SHIPMENT_NO, PO_SPLITS)
        Next
        If dst.Tables("POTSHPIE").Rows.Count <> 0 Then
            '         dst.Tables("POTSHPIE").Rows.Clear()
            tabBOL.Tabs("Import Errors").Visible = True
            MsgBox("There are errors that need be addressed", MsgBoxStyle.OkOnly, "Please refer to Import Error Tab below")

        Else
            If eMsg_Booking = "" Then
                MsgBox($"Shipment {PO_SHIPMENT_NO} has been Imported", MsgBoxStyle.OkOnly, "Verification")
            Else
                MsgBox(eMsg_Booking, MsgBoxStyle.OkOnly, "Update Disabled")
            End If
        End If
    End Sub

    Function Get_Volume_from_Dims2(CARTON_DIMENSIONS As String) As Decimal ' BELONGS IN TAC - SEE POFPACK1
        Dim CARTON_VOLUME As Decimal = 0
        If CARTON_DIMENSIONS <> "" Then
            Dim D() As String = Split(Replace(CARTON_DIMENSIONS, Chr(34), "").ToUpper, "X")
            If D.Length > 0 Then
                For I As Integer = 1 To D.Length
                    If Val(D(I - 1)) <> 0 Then
                        If CARTON_VOLUME = 0 Then CARTON_VOLUME = 1
                        CARTON_VOLUME *= Val(D(I - 1))
                    End If
                Next
                If D.Length <> 3 Then CARTON_VOLUME = 0
            End If
        End If
        Return CARTON_VOLUME

    End Function

    Function Book2ShiP(VBKG_NO As String, PO_SHIPMENT_NO As String, PO_SPLITS As Dictionary(Of String, Dictionary(Of Integer, Integer))) As String

        Dim rowPOTVBKG1 As DataRow = Fill_Record("POTVBKG1", VBKG_NO, , False)
        ' Dim rowPOTVBKG3 As DataRow = Fill_Record("POTVBKG3", VBKG_NO, , False)
        Dim LINE_NO As Integer = 1


        Dim TBLPOTORDR2 As DataTable
        Dim rowPOTORDR1 As DataRow
        Dim CONTAINER_NOs As List(Of String) = New List(Of String)

        Dim DUPPONOs As List(Of String) = New List(Of String)


        'If Not packingFromBooking Then
        '    EnforceConstraints(False)
        '    For Each TABLE_NAME As String In New String() {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP7", "POTSHIP8", "POTSHIPR", "POTPACK2", "POTPACK3", "WHTPPKM1", "WHTPPKM2"}
        '        dst.Tables(TABLE_NAME).Rows.Clear()
        '    Next
        '    EnforceConstraints(True)
        'End If


        Fill_Records("POTVBKG3", VBKG_NO, True)
        Dim CONTAINER_NO As String = ""
        Dim CONTAINER_SEAL_NO As String = ""
        Dim CONTAINER_SIZE As String = ""
        Dim CONTAINER_ADDITIONAL As String = ""
        Dim CONTAINER_CTR As Integer = 1

        For Each rowPOTVBKG3 As DataRow In dst.Tables("POTVBKG3").Select("")
            If CONTAINER_NO = "" Then
                CONTAINER_NO = rowPOTVBKG3.Item("CONTAINER_NO")
                CONTAINER_SEAL_NO = rowPOTVBKG3.Item("CONTAINER_SEAL_NO")
                CONTAINER_SIZE = rowPOTVBKG3.Item("CONTAINER_SIZE")
            End If
            CONTAINER_NOs.Add(rowPOTVBKG3.Item("CONTAINER_NO"))

            CONTAINER_ADDITIONAL = CONTAINER_ADDITIONAL & CONTAINER_CTR & ") Container# " & rowPOTVBKG3.Item("CONTAINER_NO") & vbCrLf & "Seal# " & rowPOTVBKG3.Item("CONTAINER_SEAL_NO") & " Size " & rowPOTVBKG3.Item("CONTAINER_SIZE") & vbCrLf & vbCrLf
            CONTAINER_CTR = CONTAINER_CTR + 1
        Next
        CONTAINER_ADDITIONAL = Mid(CONTAINER_ADDITIONAL, 1, 255)


        Dim rowPOTSHIP1 As DataRow = Nothing

        If Not packingFromBooking Then
            packingFromBooking = True
            ' PO_SHIPMENT_NO = ASCMAIN1.Next_Control_No("PO_SHIPMENT_NO")
            ' rowPOTSHIP1 = dst.Tables("POTSHIP1").NewRow
            rowPOTSHIP1 = dst.Tables("POTSHIP1").Rows(0)
            With rowPOTSHIP1
                ' .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIP_VESSEL") = rowPOTVBKG1.Item("VESSEL_NAME")
                .Item("PO_SHIP_ETA") = rowPOTVBKG1.Item("VBKG_ETA")
                ' .Item("PO_SHIP_LANDING_LEAD_DAYS") = ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR")
                ' .Item("PO_SHIP_REF_NO") = Val(PO_SHIPMENT_NO)
                '.Item("PO_SHIP_ADV_DATE") = DATETIME_STAMP.Date
                .Item("PO_DATE_SHIPPED") = rowPOTVBKG1.Item("VBKG_ETD")
                ' .Item("PORT_CODE") = ""
                ' .Item("WHSE_CODE") = ROWs("POTPARM1").Item("PO_PARM_DEF_WHSE_CODE")
                '.Item("INIT_OPER") = ASCMAIN1.USER_ID
                '.Item("LAST_OPER") = ASCMAIN1.USER_ID
                '.Item("INIT_DATE") = DATETIME_STAMP
                '.Item("LAST_DATE") = DATETIME_STAMP
                .Item("COST_IND") = "1"
                '.Item("FREIGHT_ENTERED_BY") = "C"
                '  .Item("PO_NOTES") = "YINTAK BOOKING"
                ' DGJ
                .Item("PO_NOTES") = CONTAINER_ADDITIONAL
                '.Item("REVIEW") = "0"
                .Item("AIR_SHIP") = IIf(rowPOTVBKG1.Item("VBKG_SHIP_BY") & "" = "AIR", "1", "0")
                .Item("COST_COMPLETE") = "0"
                .Item("LP_STATUS") = "0"
                .Item("PORT_CODE_ORIG") = rowPOTVBKG1.Item("PORT_CODE_ORIG")
                .Item("PORT_CODE_DEST") = rowPOTVBKG1.Item("PORT_CODE_DEST")
                '.Item("COST_FRT_METHOD") = "W"
                '.Item("COST_NO_DUTY") = "0"
            End With
            ' dst.Tables("POTSHIP1").Rows.Add(rowPOTSHIP1)
        Else


            'For Each TABLE_NAME As String In New String() {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP7", "POTSHIP8"}
            '    Fill_Record(TABLE_NAME, PO_SHIPMENT_NO)
            'Next

            rowPOTSHIP1 = dst.Tables("POTSHIP1").Rows.Find(PO_SHIPMENT_NO)

        End If

        '  Dim CONTAINER_NO As String = rowPOTVBKG1.Item("CONTAINER_NO") & "" ' EVENTUALLY, THIS COMES FROM POTVBKG3
        'DGJ


        Dim VEND_INV_NO As String = rowPOTVBKG1.Item("VEND_INV_NO") & ""

        Dim PO_SHIPMENT_LNO_ctr As Integer = 0
        'PO_SHIPMENT_LNO_ctr = Val(dst.Tables("POTSHIP2").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
        'Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").NewRow
        'With rowPOTSHIP2
        '    .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
        '    .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
        '    .Item("CONTAINER_NO") = CONTAINER_NO
        '    .Item("BOL_NO") = rowPOTVBKG1.Item("VBKG_BOL_NO")
        '    .Item("PO_SHIP_CTNS") = 0
        '    .Item("PO_SHIP_STATUS") = "O"
        '    '.Item("PO_SOURCE_DOC") = ""
        '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
        '    .Item("INIT_DATE") = DATETIME_STAMP
        '    .Item("LAST_OPER") = ASCMAIN1.USER_ID
        '    .Item("LAST_DATE") = DATETIME_STAMP
        '    .Item("CONTAINER_SIZE") = rowPOTVBKG1.Item("CONTAINER_SIZE")
        '    .Item("COMM_INV_NO") = rowPOTVBKG1.Item("VEND_INV_NO")
        '    .Item("ACCRUAL_STATUS") = "0"
        'End With
        'dst.Tables("POTSHIP2").Rows.Add(rowPOTSHIP2)
        rowPOTVBKG1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
        'rowPOTVBKG1.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr



        Fill_Records("POTVBKG2", VBKG_NO, False)
        Dim TOTAL_CARTONS As Integer = 0
        Dim CARTON_NO_ctr As Integer = 0

        For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select($"VBKG_NO = '{VBKG_NO}'", "PACK_LIST_NO")

            Dim rowPOTSHIP2 As DataRow = Nothing
            Dim PO_SHIPMENT_LNO As Integer = 0

            Dim sql2 As String = $"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}' and CONTAINER_NO = '{CONTAINER_NO}' and COMM_INV_NO = '{VEND_INV_NO}'"
            Dim rowPOTSHIP2s() As DataRow = dst.Tables("POTSHIP2").Select(sql2)
            If rowPOTSHIP2s.Length = 1 Then
                PO_SHIPMENT_LNO_ctr = Val(rowPOTSHIP2s(0).Item("PO_SHIPMENT_LNO") & "")
                PO_SHIPMENT_LNO = PO_SHIPMENT_LNO_ctr
                rowPOTSHIP2 = rowPOTSHIP2s(0)
            Else
                PO_SHIPMENT_LNO_ctr = Val(dst.Tables("POTSHIP2").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
                PO_SHIPMENT_LNO = PO_SHIPMENT_LNO_ctr
                rowPOTSHIP2 = dst.Tables("POTSHIP2").NewRow
                With rowPOTSHIP2
                    .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                    .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                    .Item("CONTAINER_NO") = CONTAINER_NO
                    .Item("BOL_NO") = rowPOTVBKG1.Item("VBKG_BOL_NO")
                    .Item("PO_SHIP_CTNS") = 0
                    .Item("PO_SHIP_STATUS") = "O"
                    '.Item("PO_SOURCE_DOC") = ""
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    ' .Item("CONTAINER_SIZE") = rowPOTVBKG1.Item("CONTAINER_SIZE")
                    ' DGJ
                    .Item("CONTAINER_SIZE") = CONTAINER_SIZE
                    .Item("COMM_INV_NO") = VEND_INV_NO
                    .Item("ACCRUAL_STATUS") = "0"
                End With
                dst.Tables("POTSHIP2").Rows.Add(rowPOTSHIP2)
            End If

            rowPOTVBKG2.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTVBKG2.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr



            Dim PACK_LIST_NO As String = rowPOTVBKG2.Item("PACK_LIST_NO") & ""
            Dim rowPOTPACK1 As DataRow = LookUp("POTPACK1", PACK_LIST_NO)

            Dim INITIAL_ORDER As String = rowPOTPACK1.Item("INITIAL_ORDER") & ""
            Dim CUST_CODE As String = rowPOTPACK1.Item("CUST_CODE") & ""
            Dim rowPOTPACKC As DataRow = LookUp("POTPACKC", CUST_CODE) ' dst.Tables("POTPACKC").Rows.Find(CUST_CODE)

            Fill_Records("POTPACK2", PACK_LIST_NO, False)
            Fill_Records("POTPACK3", PACK_LIST_NO, False)
            'Dim CARTON_NO_ctr As Integer = 0

            'Dim PO_SPLIT_LINES As New Dictionary(Of Integer, Integer)
            'If PO_SPLITS.ContainsKey(PO_ORDER_NO) Then
            '    PO_SPLIT_LINES = PO_SPLITS(PO_ORDER_NO)
            'End If

            Dim PPK_CODE As String = ""
            For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select($"PACK_LIST_NO = '{PACK_LIST_NO}'", "PACK_LIST_SHEET_NO")
                Dim PACK_LIST_SHEET_NO As Integer = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

                Dim CARTON_COUNT2 As Integer = 0

                If INITIAL_ORDER = "1" Then ' If CUST_CODE = "WALMART" And INITIAL_ORDER = "1" Then
                    Dim CARTON_COUNT As Integer = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                    TOTAL_CARTONS += CARTON_COUNT
                    CARTON_COUNT2 = CARTON_COUNT
                    Dim CARTON_PACK As Integer = Val(rowPOTPACK2.Item("CARTON_PACK") & "")
                    ' CARTON PACK FOR PREPACKS MAY NOT BE CORRECT - BECAUSE OF SUB UNIT PACK QTY

                    Dim CARTON_DIMENSIONS As String = rowPOTPACK2.Item("CARTON_DIMENSIONS") & ""
                    CARTON_DIMENSIONS = Validate_Carton_Dimensions(CARTON_DIMENSIONS).CTN_DIMS_CM



                    rowPOTSHIP2.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP2.Item("PO_SHIP_CTNS") & "") + CARTON_COUNT


                    If rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" <> "1" Then ' WALMART & COSTCO
                        PPK_CODE = ASCMAIN1.Next_Control_No("PPK_CODE") & "PPK"
                        PPK_CODE = Mid(PPK_CODE, 2)

                        Dim rowWHTPPKM1 As DataRow = dst.Tables("WHTPPKM1").NewRow
                        With rowWHTPPKM1
                            .Item("PPK_CODE") = PPK_CODE
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("PPK_DESC") = "" ' SHOULD BE SAME AS WHAT WAS LOADED ITO rowPOTSHIP7.Item("CARTON_COMMENTS")
                            .Item("LAST_DATE") = DATETIME_STAMP
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("CUSTOM_PPK") = "1"
                            .Item("PPK_QTY_TOTAL") = CARTON_PACK
                        End With
                        dst.Tables("WHTPPKM1").Rows.Add(rowWHTPPKM1)
                    End If

                    Dim rowPOTSHIP7 As DataRow = dst.Tables("POTSHIP7").NewRow()
                    With rowPOTSHIP7
                        .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                        CARTON_NO_ctr += 1
                        .Item("CARTON_NO") = CARTON_NO_ctr
                        .Item("CARTONS") = CARTON_COUNT
                        .Item("CARTON_COMMENTS") = ""
                        .Item("CUSTOM_PPK") = "1"
                        .Item("PPK_CODE") = PPK_CODE
                        .Item("PO_QTY_PER_CTN") = CARTON_PACK
                        .Item("STYLE_CODE") = ""
                        .Item("COLOR_CODE") = ""
                        .Item("PPK_INNER_QTY") = 0 ' ? CARTON_PACK
                        .Item("CARTON_DIMS") = CARTON_DIMENSIONS
                        Dim CARTON_VOLUME As Decimal = Get_Volume_from_Dims2(CARTON_DIMENSIONS)
                        .Item("CARTON_VOLUME") = CARTON_VOLUME
                        .Item("CARTON_WEIGHT") = rowPOTPACK2.Item("CARTON_GRS_WGT")
                    End With
                    dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
                    rowPOTPACK2.Item("CARTON_NO") = CARTON_NO_ctr

                End If

                For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_NO = '{PACK_LIST_NO}' AND PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, COLOR_CODE, PACK_LIST_SHEET_LNO")

                    Dim CARTON_COUNT As Integer = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                    Dim CARTON_PACK As Integer = Val(rowPOTPACK3.Item("CARTON_PACK") & "")
                    Dim CARTON_DIMENSIONS As String = rowPOTPACK3.Item("CARTON_DIMENSIONS") & ""
                    CARTON_DIMENSIONS = Validate_Carton_Dimensions(CARTON_DIMENSIONS).CTN_DIMS_CM

                    If INITIAL_ORDER = "1" Then
                        CARTON_COUNT = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                    End If

                    Dim STYLE_CODE As String = rowPOTPACK3.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowPOTPACK3.Item("COLOR_CODE")

                    '   If STYLE_CODE = "WM223251" Or STYLE_CODE = "WM223252" Or STYLE_CODE = "WM223647" Or STYLE_CODE = "WM223649" Then
                    '  Stop
                    ' End If

                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)

                    Dim PO_ORDER_NO As String = rowPOTPACK3.Item("PO_ORDER_NO") & ""
                    Dim PO_ORDER_LNO As Integer = Val(rowPOTPACK3.Item("PO_ORDER_LNO") & "")
                    Dim PO_ORDER_LNO_ORIG As Integer = Val(rowPOTPACK3.Item("PO_ORDER_LNO") & "")
                    Dim PO_ORDER_LNO_ORIG_PO As Integer = Val(rowPOTPACK3.Item("PO_ORDER_LNO") & "")

                    rowPOTORDR1 = dicPOTORDR1(PO_ORDER_NO)
                    TBLPOTORDR2 = dicPOTORDR2(PO_ORDER_NO)

                    Dim PO_SPLIT_LINES As New Dictionary(Of Integer, Integer)
                    If PO_SPLITS.ContainsKey(PO_ORDER_NO) Then
                        PO_SPLIT_LINES = PO_SPLITS(PO_ORDER_NO)
                    Else
                        PO_SPLITS.Add(PO_ORDER_NO, PO_SPLIT_LINES)
                    End If

                    If PO_SPLIT_LINES.ContainsKey(PO_ORDER_LNO_ORIG) Then
                        PO_ORDER_LNO = PO_SPLIT_LINES(PO_ORDER_LNO_ORIG)
                    End If

                    Dim rowPOTORDR2 As DataRow = TBLPOTORDR2.Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})


                    Dim SUB_UNIT_PACK_QTY As Integer = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")
                    If SUB_UNIT_PACK_QTY = 0 Then SUB_UNIT_PACK_QTY = 1

                    ' NEXT 3 LINES OF CODE ARE UNNEC IF IT TURNS OUT THAT POTPACK3 (POTPACK2) CARTON_PACK IS WRONG
                    'If SUB_UNIT_PACK_QTY <> 1 And INITIAL_ORDER = "1" Then
                    '    CARTON_PACK = CARTON_PACK / SUB_UNIT_PACK_QTY
                    'End If

                    Dim PO_QTY_SHP As Int32 = CARTON_COUNT * CARTON_PACK
                    If CUST_CODE = "WALMART" And INITIAL_ORDER = "1" Then
                        PO_QTY_SHP = CARTON_COUNT2 * CARTON_PACK
                    End If

                    Dim PO_QTY_OPN As Int32 = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")
                    Dim PO_QTY_ORD As Int32 = Val(rowPOTORDR2.Item("PO_QTY_ORD") & "")

                    Dim SATISFY_PO_QTY_SHP As Boolean = False
                    If PO_QTY_OPN = 0 Then
                        '  Dim rowPOTORDR2x As DataRow = TBLPOTORDR2.Rows.Find(New Object() {PO_ORDER_NO, STYLE_CODE, COLOR_CODE})
                        For Each rowPOTORDR2x As DataRow In TBLPOTORDR2.Select($"PO_ORDER_NO = '{PO_ORDER_NO}' and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' AND PO_QTY_OPN >= {CStr(PO_QTY_SHP)}") ' GREATER THAN OR = SHIP
                            '           If Val(rowPOTORDR2x.Item("PO_QTY_OPN") & "") >= PO_QTY_SHP Then
                            PO_QTY_OPN = Val(rowPOTORDR2x.Item("PO_QTY_OPN") & "")
                            PO_QTY_ORD = Val(rowPOTORDR2x.Item("PO_QTY_ORD") & "")
                            PO_ORDER_LNO = Val(rowPOTORDR2x.Item("PO_ORDER_LNO") & "")
                            PO_ORDER_LNO_ORIG = Val(rowPOTORDR2x.Item("PO_ORDER_LNO") & "")
                            SATISFY_PO_QTY_SHP = True
                            Exit For
                            '     Else
                            ' ADD QTY TO MESSAGE 
                            'End If
                        Next
                        If Not SATISFY_PO_QTY_SHP Then
                            eMsg_Booking &= vbCr & $"Insufficient Open PO Qty for Style {STYLE_CODE} and COLOR_CODE = {COLOR_CODE} in Po {PO_ORDER_NO}, Qty needed to ship {PO_QTY_SHP}"
                        End If
                    End If

                    If PO_QTY_SHP > PO_QTY_OPN Then
                        Dim rowPOTSHPIE As DataRow = dst.Tables("POTSHPIE").NewRow
                        With rowPOTSHPIE
                            .Item("WORKBOOK") = "Bk# " & VBKG_NO
                            .Item("WORKSHEET") = "PL# " & PACK_LIST_NO
                            .Item("IE_LNO") = PACK_LIST_SHEET_NO
                            .Item("ERROR_MSG") = $"Pack Qty {PO_QTY_SHP} is greater than Qty Open {PO_QTY_OPN}"
                            '  .Item("ERROR_MSG") = $"Pack Qty {PO_QTY_SHP} is greater than Qty Open {PO_QTY_OPN}"
                            '.Item("XLS_REF") = xlsRef

                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            .Item("QTY") = PO_QTY_SHP
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_REFERENCE") = rowPOTORDR1.Item("PO_REFERENCE")

                            '.Item("COMM_INV_NO") = COMM_INV_NO
                            '.Item("BOL_NO") = BOL_NO
                            .Item("CONTAINER_NO") = CONTAINER_NO
                            '.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            '  dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)
                        End With
                        dst.Tables("POTSHPIE").Rows.Add(rowPOTSHPIE)
                        eMsg_Booking &= vbCr & $"Insufficient Open PO Qty for Style {STYLE_CODE} and COLOR_CODE = {COLOR_CODE} in Po {PO_ORDER_NO}, Qty needed to ship {PO_QTY_SHP}"
                        Dim POLNO As String = PO_ORDER_NO & Format(PO_ORDER_LNO, "0000")
                        DUPPONOs.Add(POLNO)

                    ElseIf PO_QTY_SHP = 0 Then

                        Dim rowPOTSHPIE As DataRow = dst.Tables("POTSHPIE").NewRow
                        With rowPOTSHPIE
                            .Item("WORKBOOK") = VBKG_NO
                            .Item("WORKSHEET") = PACK_LIST_NO
                            .Item("IE_LNO") = PACK_LIST_SHEET_NO
                            .Item("ERROR_MSG") = $"Pack Qty {PO_QTY_SHP} is Zero"
                            '.Item("XLS_REF") = xlsRef

                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            .Item("QTY") = PO_QTY_SHP
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_REFERENCE") = rowPOTORDR1.Item("PO_REFERENCE")

                            '.Item("COMM_INV_NO") = COMM_INV_NO
                            '.Item("BOL_NO") = BOL_NO
                            .Item("CONTAINER_NO") = CONTAINER_NO
                            '.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                        End With
                        dst.Tables("POTSHPIE").Rows.Add(rowPOTSHPIE)
                        eMsg_Booking &= vbCr & $"Pack Qty {PO_QTY_SHP} is Zero For Style {STYLE_CODE} And COLOR_CODE = {COLOR_CODE} In Po {PO_ORDER_NO}, Qty needed To ship {PO_QTY_SHP}"
                        Dim POLNO As String = PO_ORDER_NO & Format(PO_ORDER_LNO, "0000")
                        DUPPONOs.Add(POLNO)



                    ElseIf PO_QTY_SHP < PO_QTY_OPN Then

                        Dim PO_QTY_BAL As Integer = PO_QTY_OPN - PO_QTY_SHP
                        ' split the po line

                        'If maxLnoSplit = 0 Or maxLnoSplit < maxLno Then
                        '    maxLnoSplit = maxLno
                        'End If

                        If Not POTORDR1_added.Contains(PO_ORDER_NO) Then
                            POTORDR1_added.Add(PO_ORDER_NO)
                            'Dim rowsDC() As DataRow = dst.Tables("POTORDR1_SPLIT").Select("PO_ORDER_NO = '" & PO_ORDER_NO & "'") 'WHY DO I NEED TO DO THIS?
                            'If rowsDC.Length = 0 Then
                            '    Dim rowPOTORDR1 As DataRow = dicPOTORDR1(PO_ORDER_NO)

                            Dim rowPOTORDR1_SPLIT As DataRow = dst.Tables("POTORDR1_SPLIT").NewRow
                            For i As Int16 = 0 To rowPOTORDR1.ItemArray.Length - 1
                                rowPOTORDR1_SPLIT.Item(i) = rowPOTORDR1.Item(i)
                            Next
                            dst.Tables("POTORDR1_SPLIT").Rows.Add(rowPOTORDR1_SPLIT)
                            'End If
                        End If

                        'Dim rowPOTORDR2 As DataRow = Nothing
                        'Dim rowPOTORDR2_orig As DataRow = tblPOTORDR2.Select("PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & PO_ORDER_LNO)(0)
                        'Dim rowPOTORDR2_pack() As DataRow = dst.Tables("POTORDR2_SPLIT").Select("PO_ORDER_NO = '" & PO_ORDER_NO _
                        '                                                                        & "' and PO_ORDER_LNO" & IIf(isSplitLine, "_ORIG", "") & " = " & PO_ORDER_LNO)
                        'Dim PO_QTY_SHP As Integer = Val(rowPOTSHPXL.Item("TOTAL_PCS") & "")
                        'Dim PO_QTY_ORD As Integer = Val(rowPOTORDR2_orig.Item("PO_QTY_ORD") & "")
                        'Dim PO_QTY_SHP_TOT As Integer = PO_QTY_SHP

                        Dim rowPOTORDR2_orig As DataRow = TBLPOTORDR2.Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})

                        Dim PO_ORDER_LNO_split As Integer = Val(TBLPOTORDR2.Compute("MAX (PO_ORDER_LNO)", $"PO_ORDER_NO = '{PO_ORDER_NO}'")) + 1

                        If PO_SPLIT_LINES.ContainsKey(PO_ORDER_LNO_ORIG) Then
                            PO_SPLIT_LINES(PO_ORDER_LNO_ORIG) = PO_ORDER_LNO_split
                        Else
                            ' If ASCMAIN1.Running_in_VS AndAlso (PO_ORDER_NO = "146260" Or rowPOTORDR2_orig.Item("PO_ORDER_NO") = "146260") And PO_ORDER_LNO_split = 91 Then Stop
                            PO_SPLIT_LINES.Add(PO_ORDER_LNO_ORIG, PO_ORDER_LNO_split)
                        End If


                        Dim TBLX As DataTable = Nothing
                        For Each TABLE_NAME As String In New String() {"POTORDR2", "POTORDR2_SPLIT"}
                            If TABLE_NAME = "POTORDR2" Then TBLX = TBLPOTORDR2
                            If TABLE_NAME = "POTORDR2_SPLIT" Then TBLX = dst.Tables("POTORDR2_SPLIT")
                            ' If ASCMAIN1.Running_in_VS AndAlso PO_ORDER_NO = "146260" And PO_ORDER_LNO_split = 91 Then Stop

                            rowPOTORDR2 = TBLX.NewRow()
                            For i As Int16 = 0 To rowPOTORDR2_orig.ItemArray.Length - 1
                                rowPOTORDR2.Item(i) = rowPOTORDR2_orig.Item(i)
                            Next
                            rowPOTORDR2.Item("PO_ORDER_LNO") = PO_ORDER_LNO_split
                            If TABLE_NAME = "POTORDR2_SPLIT" Then rowPOTORDR2.Item("PO_ORDER_LNO_ORIG") = PO_ORDER_LNO

                            rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_BAL
                            rowPOTORDR2.Item("PO_QTY_SHP") = 0
                            rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_BAL
                            ' rowPOTORDR2.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                            ' rowPOTORDR2.Item("COLOR_DESC") = dicColorDesc(rowPOTORDR2.Item("COLOR_CODE"))
                            If TABLE_NAME = "POTORDR2_SPLIT" Then rowPOTORDR2.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            TBLX.Rows.Add(rowPOTORDR2)
                        Next

                        rowPOTORDR2 = dst.Tables("POTORDR2_SPLIT").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                        If rowPOTORDR2 Is Nothing Then ' FIRST SPLIT
                            rowPOTORDR2 = dst.Tables("POTORDR2_SPLIT").NewRow()
                            For i As Int16 = 0 To rowPOTORDR2_orig.ItemArray.Length - 1
                                rowPOTORDR2.Item(i) = rowPOTORDR2_orig.Item(i)
                            Next
                            rowPOTORDR2.Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            'rowPOTORDR2.Item("PO_ORDER_LNO_ORIG") = PO_ORDER_LNO_ORIG
                            rowPOTORDR2.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            dst.Tables("POTORDR2_SPLIT").Rows.Add(rowPOTORDR2)
                        End If
                        rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_SHP
                        rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP
                        rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_SHP

                        rowPOTORDR2 = TBLPOTORDR2.Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                        rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_SHP
                        rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP
                        rowPOTORDR2.Item("PO_QTY_OPN") = 0

                    Else

                        ' Dim PO_QTY_SHP_new As Integer = Val(rowPOTORDR2.Item("PO_QTY_SHP") & "") + PO_QTY_SHP
                        ' Dim PO_QTY_OPN_new As Integer = PO_QTY_ORD - PO_QTY_SHP_new
                        'dgj 12/26 rem out next 3 lines 
                        '    rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_SHP ' PO_QTY_SHP_new
                        '    rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP ' PO_QTY_SHP_new
                        '    rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_SHP ' PO_QTY_SHP_new

                        'rowPOTORDR2 = TBLPOTORDR2.Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                        'rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_SHP
                        'rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP
                        'rowPOTORDR2.Item("PO_QTY_OPN") = 0

                        ' ?? CAUSED DOUBLING IN PO_QTY_SHIP FIELD 5/10/22 for lines that are fully satisfied no split
                        If PO_QTY_SHP = PO_QTY_OPN Then
                            rowPOTORDR2 = TBLPOTORDR2.Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                            rowPOTORDR2.Item("PO_QTY_ORD") = PO_QTY_SHP
                            rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP
                            rowPOTORDR2.Item("PO_QTY_OPN") = 0

                            ' NEED TO FIX 11/11/2022
                            '''If PO_ORDER_LNO = PO_ORDER_LNO_ORIG_PO Then
                            '''    rowPOTORDR2.Item("PO_QTY_SHP") = 0
                            '''    rowPOTORDR2.Item("PO_QTY_OPN") = PO_QTY_SHP
                            '''Else
                            '''    rowPOTORDR2.Item("PO_QTY_SHP") = PO_QTY_SHP
                            '''    rowPOTORDR2.Item("PO_QTY_OPN") = 0
                            '''End If
                            ''''DGJ
                        End If


                        Dim rowPOTORDR2_SPLIT As DataRow = dst.Tables("POTORDR2_SPLIT").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                        If rowPOTORDR2_SPLIT IsNot Nothing Then
                            rowPOTORDR2_SPLIT.Item("PO_QTY_ORD") = PO_QTY_SHP ' PO_QTY_SHP_new
                            rowPOTORDR2_SPLIT.Item("PO_QTY_SHP") = 0 ' PO_QTY_SHP ' PO_QTY_SHP_new - WJZ
                            rowPOTORDR2_SPLIT.Item("PO_QTY_OPN") = PO_QTY_SHP ' 0 ' PO_QTY_SHP_new - WJZ
                        Else
                            '  Stop
                            'Stop DGJ
                        End If

                    End If

                    rowPOTPACK3.Item("PO_ORDER_LNO") = PO_ORDER_LNO
                    rowPOTPACK3.Item("PO_QTY_OPN") = PO_QTY_SHP

                    Dim SPQ As Integer = IIf(SUB_UNIT_PACK_QTY = 0, 12, 12 / SUB_UNIT_PACK_QTY)

                    If INITIAL_ORDER = "1" Then
                    Else
                        rowPOTSHIP2.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP2.Item("PO_SHIP_CTNS") & "") + CARTON_COUNT
                        TOTAL_CARTONS += CARTON_COUNT
                        Dim rowPOTSHIP7 As DataRow = dst.Tables("POTSHIP7").NewRow()
                        With rowPOTSHIP7
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                            CARTON_NO_ctr += 1
                            .Item("CARTON_NO") = CARTON_NO_ctr

                            rowPOTPACK3.Item("CARTON_NO") = CARTON_NO_ctr

                            .Item("CARTONS") = CARTON_COUNT
                            .Item("CARTON_COMMENTS") = ""
                            .Item("CUSTOM_PPK") = ""
                            .Item("PPK_CODE") = ""
                            .Item("PO_QTY_PER_CTN") = CARTON_PACK
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            .Item("PPK_INNER_QTY") = 0
                            .Item("CARTON_DIMS") = CARTON_DIMENSIONS
                            Dim CARTON_VOLUME As Decimal = Get_Volume_from_Dims2(CARTON_DIMENSIONS)
                            .Item("CARTON_VOLUME") = CARTON_VOLUME
                            .Item("CARTON_WEIGHT") = rowPOTPACK3.Item("CARTON_GRS_WGT")
                        End With
                        dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
                    End If

                    Dim rowPOTSHIPR As DataRow
                    If dst.Tables("POTSHIPR").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO_ctr, rowPOTPACK3.Item("STYLE_CODE"), rowPOTPACK3.Item("COLOR_CODE")}) Is Nothing Then
                        rowPOTSHIPR = dst.Tables("POTSHIPR").NewRow()
                        With rowPOTSHIPR
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                        End With
                        dst.Tables("POTSHIPR").Rows.Add(rowPOTSHIPR)
                    End If

                    ' need to repeat for 1 to carton_count, and record the lpn
                    Dim rowPOTSHIP8 As DataRow = dst.Tables("POTSHIP8").NewRow()
                    With rowPOTSHIP8
                        .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                        .Item("CARTON_NO") = CARTON_NO_ctr
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE
                        .Item("QTY") = CARTON_PACK
                        .Item("DOZENS") = ""
                        '.Item("PPK_INNER_QTY") = 
                    End With
                    dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)

                    If PPK_CODE <> "" Then
                        Dim rowWHTPPKM2 As DataRow = dst.Tables("WHTPPKM2").NewRow
                        rowWHTPPKM2.Item("PPK_CODE") = PPK_CODE
                        rowWHTPPKM2.Item("STYLE_CODE") = STYLE_CODE
                        rowWHTPPKM2.Item("COLOR_CODE") = COLOR_CODE
                        rowWHTPPKM2.Item("PPK_QTY") = Val(rowPOTSHIP8.Item("QTY") & "") * IIf(rowPOTSHIP8.Item("DOZENS") & "" = "1", 12, 1)
                        dst.Tables("WHTPPKM2").Rows.Add(rowWHTPPKM2)
                    End If


                    Dim rowPOTORDRO As DataRow = dst.Tables("POTORDRO").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                    If rowPOTORDRO Is Nothing Then
                        rowPOTORDRO = dst.Tables("POTORDRO").NewRow()
                        With rowPOTORDRO
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        End With
                        dst.Tables("POTORDRO").Rows.Add(rowPOTORDRO)
                    End If
                    rowPOTORDRO.Item("PO_QTY_OPN_PRE") = PO_QTY_SHP ' PO_QTY_OPN




                    Dim rowPOTSHIP3 As DataRow = Nothing
                    ' Dim rowPOTSHIP3s() As DataRow = dst.Tables("POTSHIP3").Select($"PO_ORDER_NO ='{PO_ORDER_NO}' and PO_ORDER_LNO = {CStr(PO_ORDER_LNO)}", "")

                    ' If rowPOTSHIP3s.Length = 0 Then


                    ' CODE BLOCK BELOW AS RECLOATED UPWARDS

                    'Dim rowPOTORDR2 As DataRow = LookUp("POTORDR2", New String() {PO_ORDER_NO, PO_ORDER_LNO})
                    'Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", New String() {PO_ORDER_NO})

                    'Dim PO_QTY_SHP As Int32 = CARTON_COUNT * CARTON_PACK
                    'Dim PO_QTY_OPN As Int32 = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")

                    'If CUST_CODE = "WALMART" And INITIAL_ORDER = "1" Then
                    '    PO_QTY_SHP = CARTON_COUNT2 * CARTON_PACK
                    'End If

                    '' Dim SUB_UNIT_PACK_QTY As Integer = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")

                    'Dim SPQ As Integer = IIf(SUB_UNIT_PACK_QTY = 0, 12, 12 / SUB_UNIT_PACK_QTY)



                    rowPOTSHIP3 = dst.Tables("POTSHIP3").NewRow()
                    With rowPOTSHIP3
                        .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                        .Item("PO_ORDER_NO") = PO_ORDER_NO
                        .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        .Item("PO_QTY_SHP") = PO_QTY_SHP
                        ' .Item("PO_QTY_OPN_PRE") = PO_QTY_SHP ' PO_QTY_OPN
                        .Item("PO_QTY_REC") = 0
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = DATETIME_STAMP

                        'DUTY_RATE_CODE
                        'DUTY_RATE
                        'WEIGHT_FACTOR
                        ' PO_COST_BUFFER = 1

                        '.Item("PO_QTY_UOM") = rowPOTORDR2.Item("PO_QTY_UOM")
                        .Item("PO_COST") = Val(rowPOTORDR2.Item("PO_COST") & "")
                        .Item("PO_COST_VCOST") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "")
                        .Item("PO_COST_MATLS") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "")
                        .Item("PO_COST_VCOST_UM") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "")
                        .Item("PO_COST_MATLS_UM") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "")

                        ' IMPORTANT - note that this field is currently maintained in POTORDR2 per Dozen units, and is per unit in POTSHIP3
                        .Item("PO_COST_OTHER") = Val(rowPOTORDR2.Item("PO_COST_OTHER") & "") / SPQ

                        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then ' VAN PARANOIA
                            .Item("PO_COST_COMM") = Val(rowPOTORDR2.Item("PO_COST_COMM") & "")
                        Else
                            If rowPOTORDR1.Item("PO_COMM_PAYABLE_TO_BRKR") & "" = "1" Then
                                .Item("PO_COST_COMM") = Val(rowPOTORDR1.Item("PO_COMM_PCT") & "")
                            Else
                                .Item("PO_COST_COMM") = 0
                            End If
                        End If

                        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                            .Item("PO_COST_BUFFER") = 5
                        End If

                        ' this is not exactly true- but we can let the calculation routines fix it later
                        .Item("PO_COST_LANDED") = Val(rowPOTORDR2.Item("PO_COST") & "")

                        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                            If Val(rowPOTORDR2.Item("DFQUOTA") & "") = 1 Then
                                .Item("PO_COST_QUOTA_DF") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "") / SPQ
                                .Item("PO_COST_QUOTA_DF_DZ") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "")
                            Else
                                .Item("PO_COST_QUOTA") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "") / SPQ
                                .Item("PO_COST_QUOTA_DZ") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "")
                            End If
                        End If

                        If Val(rowPOTORDR2.Item("PO_COST_VCOST_DZ") & "") = 0 Then
                            .Item("PO_COST_VCOST_DZ") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "") * SPQ
                        Else
                            .Item("PO_COST_VCOST_DZ") = Val(rowPOTORDR2.Item("PO_COST_VCOST_DZ") & "")
                        End If

                        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                            If Val(rowPOTORDR2.Item("PO_COST_MATLS_DZ") & "") = 0 Then
                                .Item("PO_COST_MATLS_DZ") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "") * SPQ
                            Else
                                .Item("PO_COST_MATLS_DZ") = Val(rowPOTORDR2.Item("PO_COST_MATLS_DZ") & "")
                            End If
                        End If

                        .Item("PO_COST_OTHER_DZ") = Val(rowPOTORDR2.Item("PO_COST_OTHER") & "") ' see note above regarding unit of measure for PO_COST_OTHER
                        .Item("SUB_UNIT_PACK_QTY") = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")
                        .Item("CARTON_PACK_QTY") = Val(rowPOTORDR2.Item("CARTON_PACK_QTY") & "")
                        'If Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "") = 0 Then
                        '    .Item("PO_QTY_SHP_DZ") = 0
                        '    .Item("NET_OPEN_DZ") = 0
                        'Else
                        '    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        '        .Item("PO_QTY_SHP_DZ") = PO_QTY_SHP / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                        '    Else
                        '        .Item("PO_QTY_SHP_DZ") = PO_QTY_OPN / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                        '    End If
                        '    .Item("NET_OPEN_DZ") = PO_QTY_OPN / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                        'End If
                        '.Item("PO_QTY_REC_DZ") = 0

                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE

                        .Item("PO_REFERENCE") = rowPOTORDR1.Item("PO_REFERENCE")
                        .Item("PO_DATE_SHIP_BY") = rowPOTORDR2.Item("PO_DATE_SHIP_BY")
                        .Item("FOB_CMT") = (rowPOTORDR1.Item("FOB_CMT") & "")
                        .Item("VEND_CODE") = rowPOTORDR1.Item("VEND_CODE")


                        .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")

                    End With

                    '  Stop
                    ' fix below crfeate a list of errors po_order_no _ Format po_line_no "0000" i
                    ' if po,polno in Error list avoide new row
                    'End If

                    Dim BADPO As Integer = 0
                    Dim DUPPONO As String = PO_ORDER_NO & Format(PO_ORDER_LNO, "0000")
                    If DUPPONOs IsNot Nothing Then
                        For Each DUPPONO In DUPPONOs
                            If PO_ORDER_NO & Format(PO_ORDER_LNO, "0000") = DUPPONO Then
                                BADPO += 1
                                Exit For
                            End If
                        Next
                    End If
                    If BADPO = 0 Then
                        dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)
                    End If
                    'Else
                    '    rowPOTSHIP3 = rowPOTSHIP3s(0)
                    '    rowPOTSHIP3.Item("PO_QTY_SHP") = Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "") + CARTON_COUNT * CARTON_PACK
                    'End If
                Next
            Next
        Next


        Dim rowPOTSHIP4 As DataRow = Nothing
        Dim rowPOTSHIP4s() As DataRow = dst.Tables("POTSHIP4").Select($"CONTAINER_NO = '{CONTAINER_NO}'")
        If rowPOTSHIP4s.Length = 0 Then
            PO_SHIPMENT_LNO_ctr = Val(dst.Tables("POTSHIP4").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
            rowPOTSHIP4 = dst.Tables("POTSHIP4").NewRow
            With rowPOTSHIP4
                .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                .Item("CONTAINER_NO") = CONTAINER_NO
                .Item("CONTAINER_TYPE_CODE") = ""
                .Item("PO_SHIP_CTNS") = TOTAL_CARTONS
                '.Item("PO_SHIP_STATUS") = "?"
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                '.Item("TOTAL_WEIGHT") = -1
                '.Item("CBM") = -1
                '.Item("TRUCKING") = -1
                '.Item("FREIGHT_AMT") = -1
                ' .Item("CONTAINER_SEAL_NO") = rowPOTVBKG1.Item("CONTAINER_SEAL_NO")
                'DGJ
                .Item("CONTAINER_SEAL_NO") = CONTAINER_SEAL_NO
                .Item("TRAILER_NO") = "?"
                .Item("CONTAINER_SEAL_INTACT") = "?"
            End With
            dst.Tables("POTSHIP4").Rows.Add(rowPOTSHIP4)
        Else
            rowPOTSHIP4 = rowPOTSHIP4s(0)
            rowPOTSHIP4.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP4.Item("PO_SHIP_CTNS") & "") + TOTAL_CARTONS
        End If

        Return PO_SHIPMENT_NO

    End Function

    Private Sub grdPOTSHIP2_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTSHIP2.InitializeLayout

    End Sub

#Region "Packing Slips"
    Sub Setup_PackingSLips()
        EnforceConstraints(False)
        Fill_Records("POTPACKG")
        Fill_Records("POTPACKH")
        Fill_Records("POTPCKS1")
        Fill_Records("POTPCKS2")
        EnforceConstraints(True)
        Sort_grdColumns(grdPOTPACKG, "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")
        UltraExplorerBar1.Groups("Screen Control").Visible = False
        UltraExplorerBar1.Groups("Packing Slips").Visible = True
        UltraExplorerBar1.Groups("Options").Visible = False
        spl.Panel1Collapsed = True
        For Each rowPOTPACKG As DataRow In dst.Tables("POTPACKG").Select("")
            Dim PO_SHIPMENT_NO As String = rowPOTPACKG("PO_SHIPMENT_NO")
            Dim PO_SHIPMENT_LNO As String = rowPOTPACKG("PO_SHIPMENT_LNO")
            'ASCMAIN1.sql = "select sum(po_qty_shp) from potship3 where po_shipment_no = :parm1 and po_shipment_lno - :parm2"
            'Dim shp As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "vv", New Object() {po_shipment_no, po_shipment_lno})
            Dim pack As String = dst.Tables("POTPCKS2").Compute("sum(PO_QTY_PACK)", $"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}' and PO_SHIPMENT_LNO = '{PO_SHIPMENT_LNO}'").ToString
            rowPOTPACKG("PO_QTY_PACK") = Val(pack + "")
        Next
        If EntryMode <> "L" Then PackingSlipModes(False)

    End Sub

    Private Sub grdPOTPACKG_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTPACKG.DoubleClickRow

        If grdPOTPACKG.ActiveRow IsNot Nothing AndAlso grdPOTPACKG.ActiveRow.IsDataRow Then
            If UltraExplorerBar1.Groups("Packing Slips").Visible = False Then
                Exit Sub
            Else
                If tabPackSlips.Tabs("Packing Slip Details").Selected = True Then
                    Dim PACK_SLIP_NO As String = Absx1.txtFor("PACK_SLIP_NO").Text
                    If grdPOTPACKG.ActiveRow.Band.Key = "POTPACKG" Then
                        For Each gr2 As UltraWinGrid.UltraGridRow In grdPOTPACKG.ActiveRow.ChildBands(0).Rows
                            CreatePackSlipDtl(PACK_SLIP_NO, gr2)
                        Next
                    Else
                        CreatePackSlipDtl(PACK_SLIP_NO, grdPOTPACKG.ActiveRow)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub PackingSlipModes(active As Boolean)
        If EntryMode = "L" Then
            UltraExplorerBar1.Groups("Packing Slips").Items("Update New Lines").Settings.Enabled = (EntryMode = "L")
        Else
            If active Then
                tabPackSlips.Tabs("Packing Slip Details").Selected = active
            Else
                tabPackSlips.Tabs("Packing Slips").Selected = True
                ASCMAIN1.MultiTask_Release()
                EntryMode = "V"
            End If
            SplitContainer3.SplitterDistance = 120
            tab0.Tabs("Shipments").Enabled = Not active
            UltraExplorerBar1.Groups("Packing Slips").Items("New Packing Slip").Settings.Enabled = active
            UltraExplorerBar1.Groups("Packing Slips").Items("Edit Packing Slip").Settings.Enabled = (EntryMode = "N" Or EntryMode = "E")
            UltraExplorerBar1.Groups("Packing Slips").Items("Print Packing Slip").Settings.Enabled = Not (active And EntryMode = "V")
            UltraExplorerBar1.Groups("Packing Slips").Items("Email Packing Slip").Settings.Enabled = Not (active And EntryMode = "V")
            UltraExplorerBar1.Groups("Packing Slips").Items("Update Packing Slip").Settings.Enabled = (EntryMode = "V")
            UltraExplorerBar1.Groups("Packing Slips").Items("Update New Lines").Settings.Enabled = (EntryMode = "V")
            UltraExplorerBar1.Groups("Packing Slips").Items("Cancel").Settings.Enabled = Not active
        End If
    End Sub

    Private Sub Update_New_Balance_Bookings()

        BeginTrans()
        Update_Record_TDA("POTSHIP2")
        Update_Record_TDA("POTSHIP3")
        Update_Record_TDA("POTSHIP7")
        CommitTrans("New Shipment Line created")

        EntryMode = "V"
        PackingSlipModes(False)
    End Sub


    Private Sub CreatePackingSlip()
        'Review FillPackingList if changing this sub
        Dim PACK_SLIP_NO As String = ASCMAIN1.Next_Control_No("POTPCKS1.PACK_SLIP_NO")
        Dim CUST_STORE_NO As String = ""
        Dim WHSE_CODE As String = ""
        Dim rowPOTPCKS1 As DataRow = Nothing

        Dim dvw As DataView = DirectCast(grdPOTPCKS2.DataSource, DataTable).DefaultView
        dvw.RowFilter = $"PACK_SLIP_NO = '{PACK_SLIP_NO}'"


        For Each grow As UltraWinGrid.UltraGridRow In grdPOTPACKG.Selected.Rows
            If CUST_STORE_NO = "" Then
                CUST_STORE_NO = grow.Cells("CUST_STORE_NO").Value & ""
                WHSE_CODE = grow.Cells("WHSE_CODE").Value & ""

                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {"171659", "MK", CUST_STORE_NO})

                rowPOTPCKS1 = dst.Tables("POTPCKS1").NewRow
                With rowPOTPCKS1
                    .Item("PACK_SLIP_NO") = PACK_SLIP_NO
                    .Item("PACK_SLIP_DATE") = DATETIME_STAMP.Date
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("TRAILER_NO") = ""
                    .Item("CUST_NAME") = rowARTCUST2("CUST_NAME") & ""
                    .Item("CUST_ADDR1") = rowARTCUST2("CUST_ADDR1") & ""
                    .Item("CUST_ADDR2") = rowARTCUST2("CUST_ADDR2") & ""
                    .Item("CUST_CITY") = rowARTCUST2("CUST_CITY") & ", " & rowARTCUST2("CUST_STATE") & " " & rowARTCUST2("CUST_ZIP_CODE")
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                End With
                dst.Tables("POTPCKS1").Rows.Add(rowPOTPCKS1)
                FillPackSlipHeader(rowPOTPCKS1)
            End If
            For Each gr2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                CreatePackSlipDtl(PACK_SLIP_NO, gr2)
            Next
        Next

    End Sub

    Private Sub CreatePackSlipDtl(PACK_SLIP_NO As String, gr2 As UltraWinGrid.UltraGridRow)
        'PACK_SLIP_NO, PO_SHIPMENT_NO, PO_SHIPMENT_LNO , PO_ORDER_NO, PO_ORDER_LNO
        Dim row As DataRow = dst.Tables("POTPCKS2").Rows.Find(New Object() {PACK_SLIP_NO, gr2.Cells("PO_SHIPMENT_NO").Value, gr2.Cells("PO_SHIPMENT_LNO").Value, gr2.Cells("PO_ORDER_NO").Value, gr2.Cells("PO_ORDER_LNO").Value})
        If Not row Is Nothing Then
            MsgBox("Unable to add row, already part of packing slip", vbOKOnly, "Cannot Add")
            Exit Sub
        End If

        Dim rowPOTPCKS2 As DataRow = dst.Tables("POTPCKS2").NewRow
        Dim rowPOTSHIP2 As DataRow = dst.Tables("POTPACKG").Rows.Find(New Object() {gr2.Cells("PO_SHIPMENT_NO").Value & "", gr2.Cells("PO_SHIPMENT_LNO").Value & ""})
        Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", gr2.Cells("ORDR_NO").Value & "")
        Dim rowSOTORDR2 As DataRow = LookUp("SOTORDR2", New String() {gr2.Cells("ORDR_NO").Value & "", gr2.Cells("ORDR_LNO").Value & ""})
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", gr2.Cells("STYLE_CODE").Value & "")

        With rowPOTPCKS2
            .Item("PACK_SLIP_NO") = PACK_SLIP_NO
            .Item("PO_SHIPMENT_NO") = gr2.Cells("PO_SHIPMENT_NO").Value & ""
            .Item("PO_SHIPMENT_LNO") = gr2.Cells("PO_SHIPMENT_LNO").Value & ""
            .Item("PO_ORDER_NO") = gr2.Cells("PO_ORDER_NO").Value & ""
            .Item("PO_ORDER_LNO") = gr2.Cells("PO_ORDER_LNO").Value & ""
            .Item("STYLE_CODE") = gr2.Cells("STYLE_CODE").Value & ""
            .Item("COLOR_CODE") = gr2.Cells("COLOR_CODE").Value & ""
            .Item("CARTON_PACK_QTY") = rowICTSTYL1("CARTON_PACK_QTY") + 0
            .Item("CUST_SKU") = rowSOTORDR2("CUST_SKU") & ""
            .Item("ORDR_CUST_PO") = rowSOTORDR1("ORDR_CUST_PO") & ""
            .Item("CONTAINER_NO") = rowPOTSHIP2("CONTAINER_NO") & ""
            .Item("CUST_STORE_NO") = rowSOTORDR1("CUST_STORE_NO") & ""
            .Item("ORDR_NO") = gr2.Cells("ORDR_NO").Value & ""
            .Item("LOAD_NO") = ""
            .Item("PO_QTY_BAL") = gr2.Cells("PO_QTY_BAL").Value + 0
            .Item("PO_QTY_PACK") = gr2.Cells("PO_QTY_BAL").Value + 0
            .Item("IN_ERR") = "0"
        End With
        dst.Tables("POTPCKS2").Rows.Add(rowPOTPCKS2)

    End Sub
    Private Sub EditPackingList()
        'Review CreatePackingList if changing this sub
        Dim PACK_SLIP_NO As String = grdPOTPCKS1.ActiveRow.Cells("PACK_SLIP_NO").Value

        Dim rowPOTPCKS1 As DataRow = Nothing
        Dim rowPOTPCKS2 As DataRow = Nothing

        Dim dvw As DataView = DirectCast(grdPOTPCKS2.DataSource, DataTable).DefaultView
        dvw.RowFilter = $"PACK_SLIP_NO = '{PACK_SLIP_NO}'"

        For Each row As DataRow In dst.Tables("POTPCKS2").Select($"PACK_SLIP_NO = '{PACK_SLIP_NO}'")
            Dim row2 As DataRow = dst.Tables("POTPACKH").Rows.Find(New Object() {row("PO_SHIPMENT_NO"), row("PO_SHIPMENT_LNO"), row("STYLE_CODE"), row("COLOR_CODE")})
            row("PO_QTY_BAL") = row2("PO_QTY_BAL") + row("PO_QTY_PACK")
            row("IN_ERR") = "0"
        Next

        rowPOTPCKS1 = dst.Tables("POTPCKS1").Rows.Find(PACK_SLIP_NO)

        FillPackSlipHeader(rowPOTPCKS1)
        'grdPOTPACKG.Enabled = False

    End Sub
    Private Sub FillPackSlipHeader(row As DataRow)
        Absx1.txtFor("PACK_SLIP_NO").Text = row("PACK_SLIP_NO")
        Absx1.dteFor("PACK_SLIP_DATE").Value = row("PACK_SLIP_DATE")
        Absx1.txtFor("PACK_WHSE_CODE").Text = row("WHSE_CODE")
        Absx1.txtFor("CUST_STORE_NO").Text = row("CUST_STORE_NO")
        Absx1.txtFor("ADDRESS").Text = row("CUST_NAME") & vbCrLf & row("CUST_ADDR1") & vbCrLf & row("CUST_CITY") & " " & row("CUST_STATE") & ", " & row("CUST_ZIP_CODE")
        Absx1.txtFor("TRAILER_NO").Text = row("TRAILER_NO") & ""
        PackingSlipModes(True)

    End Sub

    Private Sub PackSlipUpdt()
        Dim PACK_SLIP_NO As String = Absx1.txtFor("PACK_SLIP_NO").Text
        Dim rowPOTPCKS1 As DataRow = dst.Tables("POTPCKS1").Rows.Find(PACK_SLIP_NO)

        rowPOTPCKS1("PACK_SLIP_DATE") = Absx1.dteFor("PACK_SLIP_DATE").Value
        rowPOTPCKS1("CUST_STORE_NO") = Absx1.txtFor("CUST_STORE_NO").Text
        rowPOTPCKS1("TRAILER_NO") = Absx1.txtFor("TRAILER_NO").Text

        rowPOTPCKS1("LAST_OPER") = ASCMAIN1.USER_ID
        rowPOTPCKS1("LAST_DATE") = Now + ASCMAIN1.NowTSD

        BeginTrans()
        Update_Record_TDA("POTPCKS1")
        Update_Record_TDA("POTPCKS2")
        CommitTrans("Packing Slip Updated")

        Setup_PackingSLips()

        PackingSlipModes(False)

    End Sub

    Private Sub PackSlipCancel()
        tabPackSlips.Tabs("Packing Slips").Selected = True
        EnforceConstraints(False)
        Fill_Records("POTPCKS1")
        Fill_Records("POTPCKS2")
        EnforceConstraints(True)

        PackingSlipModes(False)

    End Sub

    Private Sub grdPOTPCKS1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTPCKS1.DoubleClickRow
        If grdPOTPCKS1.Selected.Rows.Count = 0 Then
            Dim grow As UltraWinGrid.UltraGridRow = e.Row
            grow.Selected = True
        End If
        Click_Command("View Packing Slip")

    End Sub

    Private Sub grdPOTPCKS2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTPCKS2.InitializeRow
        If (e.Row.Cells("PO_QTY_PACK").Value & "" <> "" And Val(e.Row.Cells("CARTON_PACK_QTY").Value & "") > 0) AndAlso (e.Row.Cells("PO_QTY_PACK").Value Mod (e.Row.Cells("CARTON_PACK_QTY").Value + 0)) <> 0 Then
            e.Row.Cells("PO_QTY_PACK").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PO_QTY_PACK").ToolTipText = "Qty selected does not match Carton Pack"
        End If
    End Sub
    Private Sub grdPOTPCKS2_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdPOTPCKS2.BeforeCellUpdate
        Select Case e.Cell.Column.Key

            Case "PO_QTY_PACK"
                If e.NewValue & "" <> "" Then
                    If (e.NewValue Mod grdPOTPCKS2.ActiveRow.Cells("CARTON_PACK_QTY").Value + 0) <> 0 Then
                        grdPOTPCKS2.ActiveRow.Cells("PO_QTY_PACK").Appearance.ForeColor = Drawing.Color.Red
                        grdPOTPCKS2.ActiveRow.Cells("PO_QTY_PACK").ToolTipText = "Qty selected does not match Carton Pack"
                    End If
                    If e.NewValue > grdPOTPCKS2.ActiveRow.Cells("PO_QTY_BAL").Value + 0 Then
                        MsgBox($"{e.NewValue} is more than Qty Available {grdPOTPCKS2.ActiveRow.Cells("PO_QTY_BAL").Value + 0}", vbOKOnly + vbCritical, "Qty Error")
                        e.Cancel = True
                    End If
                End If
        End Select

    End Sub
    Private Sub Print_PackingSlip(Email As Boolean)
        Print_Report_Begin()
        Dim REPORT_NO As String = ""
        Dim AllClosed As Boolean = True
        Dim FILE_NAME As String = Absx1.txtFor("PACK_SLIP_NO").Text
        Dim PACK_SLIP_NO As String = Absx1.txtFor("PACK_SLIP_NO").Text

        Dim RPT As String = "PORPCKS1"
        Dim FILTER As String = "{POTPCKS1.PACK_SLIP_NO} = """ & CStr(PACK_SLIP_NO) & """"
        Dim RPT_TITLE As String = "PO Packing Slip Report"

        Dim PO_SHIPMENT_NOs As String = ""
        Dim PO_ORDER_NOs As String = ""
        Dim ORDR_NOs As String = ""
        For Each row As DataRow In dst.Tables("POTPCKS2").Select($"PACK_SLIP_NO = '{PACK_SLIP_NO}'")
            If Not PO_SHIPMENT_NOs.Contains(row("PO_SHIPMENT_NO")) Then
                PO_SHIPMENT_NOs &= ",'" & row("PO_SHIPMENT_NO") & "'"
            End If
            If Not PO_ORDER_NOs.Contains(row("PO_ORDER_NO")) Then
                PO_ORDER_NOs &= ",'" & row("PO_ORDER_NO") & "'"
            End If
            If Not ORDR_NOs.Contains(row("ORDR_NO")) Then
                ORDR_NOs &= ",'" & row("ORDR_NO") & "'"
            End If
        Next

        If (Email) Then Generate_Report(RPT, RPT_TITLE, "", FILTER, "PDF", FILE_NAME)
        Generate_Report(RPT, RPT_TITLE, "", FILTER)

        Print_Report_End()

        If (Email) Then
            Dim frmASFMSGBF As New ASFMSGBF
            Dim Label As New System.Text.StringBuilder With {.Length = 0}
            Label.AppendLine("Enter Email Message for Packing Slip:" & PACK_SLIP_NO)
            Dim Caption As String = "Packing Slip"
            Dim emailNote As String = frmASFMSGBF.Get_txtblock_from_User(Label.ToString, Caption, "", False, 0)

            Try
                Dim clsASCNOTE1 As New TAC.ASCNOTE1("PORPCKS1", dst)
                clsASCNOTE1.Note = String.Format("Packing Slip:{0} Created", PACK_SLIP_NO) & vbCrLf & emailNote
                clsASCNOTE1.ReplaceEmailSubject = $"New QVC Packing Slip:{PACK_SLIP_NO}"
                clsASCNOTE1.Attachments.Add(ASCMAIN1.Folders("Temp") & FILE_NAME & ".pdf")
                clsASCNOTE1.CreateComponents()
                clsASCNOTE1.EmailDocument()


                ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                    & " Select 'POTPCKS1', PACK_SLIP_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'QVC_PS','QVC Packing Slip emailed', ''" _
                    & " from POTPCKS1 " & vbCrLf _
                    & " where (PO_SHIPMENT_NO) in ('" & PACK_SLIP_NO & "')"
                ASCDATA1.ExecuteSQL()
            Catch ex As Exception
                MessageBox.Show("Error emailing Warehouse receipts." & ex.Message, "Email Error", MessageBoxButtons.OK)
            End Try

        End If
    End Sub

#End Region

    Sub BOOKING_INTEGRITY(PO_SHIPMENT_NO As String)

        ASCMAIN1.sql = "UPDATE POTPACK2 SET CARTON_NO = NULL WHERE PACK_LIST_NO IN (" _
         & " Select PACK_LIST_NO from POTVBKG2 where (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) In (" _
         & " Select DISTINCT PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTVBKG2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
         & " MINUS" _
         & " Select  PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTSHIP2 where PO_SHIPMENT_NO  = '" & PO_SHIPMENT_NO & "'" _
         & "))"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "UPDATE POTPACK3 SET CARTON_NO = NULL WHERE PACK_LIST_NO IN (" _
         & " Select PACK_LIST_NO from POTVBKG2 where (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) In (" _
         & " Select DISTINCT PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTVBKG2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
         & " MINUS" _
         & " Select  PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTSHIP2 where PO_SHIPMENT_NO  = '" & PO_SHIPMENT_NO & "'" _
         & "))"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update POTLPNL1 Set PO_SHIPMENT_NO = NULL, PO_SHIPMENT_LNO = Null, CARTON_NO = Null WHERE (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) IN (" _
         & " Select DISTINCT PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTVBKG2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
         & " MINUS" _
         & " Select  PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTSHIP2 where PO_SHIPMENT_NO  = '" & PO_SHIPMENT_NO & "'" _
         & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update POTVBKG1 SET PO_SHIPMENT_NO = NULL WHERE VBKG_NO IN (" _
         & " Select VBKG_NO from POTVBKG2 where (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) In (" _
         & " Select DISTINCT PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTVBKG2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
         & " MINUS" _
         & " Select  PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTSHIP2 where PO_SHIPMENT_NO  = '" & PO_SHIPMENT_NO & "'" _
         & "))"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update POTVBKG2 Set PO_SHIPMENT_NO = NULL, PO_SHIPMENT_LNO = Null WHERE VBKG_NO IN (" _
         & " Select VBKG_NO from POTVBKG2 where (PO_SHIPMENT_NO, PO_SHIPMENT_LNO) In (" _
         & " Select DISTINCT PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTVBKG2 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
         & " MINUS" _
         & " Select  PO_SHIPMENT_NO, PO_SHIPMENT_LNO from POTSHIP2 where PO_SHIPMENT_NO  = '" & PO_SHIPMENT_NO & "'" _
         & "))"
        ASCDATA1.ExecuteSQL()


    End Sub

End Class




Public Class poPackingLine
    Public rowPOTORDR2 As DataRow
    Public splitLine As Boolean
    Public eMsg As String
End Class

Public Class CARTON_DIMENSIONS
    Public CTN_DIMS_IN As String
    Public CTN_DIMS_CM As String
End Class
