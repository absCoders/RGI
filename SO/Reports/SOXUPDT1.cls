VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SOXUPDT1"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim RYP As String
Dim NYP As String   'next year period
Dim NYM As String
Dim RYPLegend As String

Dim SOTINVH1_temp As String
Dim SUBTITLE As String
Dim WHSECODES() As String
Dim chkSEARSCAN As String

Sub Build_Workfile()

    Call Build_WorkFile_DB_Init
    
    Call Track("Loading Run-Time Options", "")
    
    Dim dyn As OraDynaset
    Dim dynWK As Recordset
    
    NYM = Period_Calc(CYM, 1)
    NYP = Period_Calc(CYP, 1)
    
    RYP = CYP
    RYPLegend = Get_Legend(RYP)
    xYPD = RYPLegend
    SUBTITLE = "New Invoices"
    
    chkSEARSCAN = SRead(opts, "CHKSEARSCAN", 2)
    
' Load SOTINVH1 - Based on Run-Time Options

    Call Track("Now Building Work File", "")
    
    Call Track("-", "Invoices")
    sql = "Select SOTINVH1.*"
    sql = sql & ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.CUST_DC_NO, SOTORDR1.EDI_APPOINTMENT"
    sql = sql & " , 0 TOTAL_UNITS "
    sql = sql & " , 0 TOTAL_UNITS_CANC "
    sql = sql & " , 0 TOTAL_UNITS_BACK "
    sql = sql & " from SOTINVH1,SOTORDR1 "
    sql = sql & " where INV_TYPE = 'I'"
    sql = sql & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO"
    sql = sql & "   and SOTINVH1.ORDR_YYYYPP_UPDATED is Null"
    SOTINVH1_temp = Temp_Table(sql)
    sql = "Create Index I_" & SOTINVH1_temp & "_1 on " & SOTINVH1_temp & " (INV_NO)"
    OraD.ExecuteSQL sql
    sql = "Create Index I_" & SOTINVH1_temp & "_2 on " & SOTINVH1_temp & " (PICK_NO)"
    OraD.ExecuteSQL sql
    sql = "Create Index I_" & SOTINVH1_temp & "_3 on " & SOTINVH1_temp & " (ORDR_NO)"
    OraD.ExecuteSQL sql
    sql = "Create Index I_" & SOTINVH1_temp & "_4 on " & SOTINVH1_temp & " (SHIP_BOL_NO)"
    OraD.ExecuteSQL sql
    
    OraD.ExecuteSQL "ALTER SESSION SET OPTIMIZER_MODE = " & "choose"
    OraD.ExecuteSQL "Analyze Table " & SOTINVH1_temp & " Compute Statistics"
    
    sql = "Select SOTINVH1.* from " & SOTINVH1_temp & " SOTINVH1"
    Call Ora_to_Acc(Nothing, "SOWINVH1", 2, "", sql)
    
    sql = "Update SOWINVH1 set ORDR_YYYYPP_UPDATED = '" & CYP & "'"
    AccD.Execute sql
    
    sql = "Update SOWINVH1 set ORDR_YYYYPP_UPDATED = '" & NYP & "'"
    sql = sql & " where INV_DATE >= #" & Format(Right(NYM, 2) & "/01/" & Left$(NYM, 4), "MM/DD/YY") & "# "
    AccD.Execute sql
    
    Call Track("-", "Credits")
    sql = " SELECT  Distinct " & vbCr
    sql = sql & " H1.INV_TYPE, H1.INV_NO, H1.CUST_CODE, H1.CUST_STORE_NO, H1.ORDR_CUST_PO, H1.ORDR_NO," & vbCr
    sql = sql & " H1.WHSE_CODE, H1.INV_SALES, H1.INV_COGS, H1.INV_FREIGHT, H1.INV_MISC_CHG, H1.INV_TOTAL_AMOUNT," & vbCr
    sql = sql & " H1.REASON_CODE, H1.INV_DATE, H1.ORDR_DATE_UPDATED, H1.ORDR_YYYYPP_UPDATED, H1.ORDR_BILL_TO_CUST," & vbCr
    sql = sql & " H1.POST_CODE, H1.SHIP_BOL_NO, ST.SALES_DIVISION_CODE, H1.INV_PRINTED, H1.INV_810_BATCH_NO," & vbCr
    sql = sql & " H1.INV_NO_CONS, H1.TERM_CODE, H1.INIT_DATE, H1.INIT_OPER, H1.PICK_NO, H1.INV_NO_REV, H1.CUST_FACTOR_IND, " & vbCr
    sql = sql & " H1.CUST_SURCHARGE_IND, H1.SREP_CODE, H1.INV_COMMENT, H1.REGISTER_XNO, H1.INV_NO_REV_BY, " & vbCr
    sql = sql & " H1.SREP2_CODE, H1.REVISED_CUST_STORE_NO, H1.LAST_REVISED_DATE, H1.LAST_REVISED_OPER, " & vbCr
    sql = sql & " H1.EDI_RETRANSMIT_IND, H1.ORIG_CUST_STORE_NO, H1.INV_TOTAL_AMOUNT_CURR, H1.CURR_CODE, " & vbCr
    sql = sql & " H1.CURR_EXCH_RATE, H1.INV_SALES_CURR, H1.INV_FREIGHT_CURR, H1.INV_MISC_CHG_CURR, " & vbCr
    sql = sql & " H1.INV_TOTAL_AMT_CURR, H1.GST_TAX, H1.GST_TAX_CURR" & vbCr
    sql = sql & " from SOTINVH1 H1, SOTINVH2 H2, ICTSTYL1 ST" & vbCr
    sql = sql & " where H1.INV_TYPE = 'C'" & vbCr
    sql = sql & " AND H1.INV_TYPE = H2.INV_TYPE" & vbCr
    sql = sql & " AND H1.INV_NO = H2.INV_NO" & vbCr
    sql = sql & " AND H2.STYLE_CODE = ST.STYLE_CODE" & vbCr
    sql = sql & "   and H1.ORDR_YYYYPP_UPDATED is Null"
    Call Ora_to_Acc(Nothing, "SOWINVHR", 2, "", sql)
    
    sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, 'N' EDI, 'N' MULTI_STORE from ARTCUST1"
    Call Ora_to_Acc(Nothing, "ARWCUST1", 1, "", sql)
    
    Call Track("-", "EDI Parameters")
    Dim dynEDTTRPM1 As OraDynaset
    sql = "Select distinct CUST_CODE from EDTTRPM1"
    sql = sql & " WHERE EDI_STATUS = 'P'"
    sql = sql & " AND EDI_DOC_NO = '810'"
    Set dynEDTTRPM1 = OraD.CreateDynaset(sql, 8&)
    Do While Not dynEDTTRPM1.EOF
        sql = "Update ARWCUST1 set EDI = 'Y' where CUST_CODE = '" & dynEDTTRPM1.Fields("CUST_CODE").Value & "'"
        AccD.Execute sql
        dynEDTTRPM1.MoveNext
    Loop
    dynEDTTRPM1.Close
     
    Call Track("-", "Multiple Store")
    sql = "Select CUST_CODE from ARTCUST2 "
    sql = sql & " WHERE CUST_ADDR_TYPE = 'MK'"
    sql = sql & " GROUP BY CUST_CODE HAVING COUNT (*) > 1"
    Set dyn = OraD.CreateDynaset(sql, 8&)
    Do While Not dyn.EOF
        sql = "Update ARWCUST1 set MULTI_STORE = 'Y' where CUST_CODE = '" & dyn.Fields("CUST_CODE").Value & "'"
        AccD.Execute sql
        dyn.MoveNext
    Loop
    dyn.Close
    
    

' Load SOTINVH2, SOTPICK1, & SOTPICK2 records corresponding to SOTINVH1
    
    sql = "Select * from SOTFPCT1"
    Call Ora_to_Acc(Nothing, "SOWFPCT1", 1, "", sql)
    sql = "Select '" & CYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT"
    sql = sql & " from SOTPARM1 where SO_PARM_KEY = 'Z'"
    Call Ora_to_Acc(Nothing, "SOWFPCT1", 1, "N", sql)
    sql = "Select '" & NYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT"
    sql = sql & " from SOTPARM1 where SO_PARM_KEY = 'Z'"
    Call Ora_to_Acc(Nothing, "SOWFPCT1", 1, "N", sql)
    
    Call Track("-", "Invoice Details")
    sql = "Select SOTINVH2.*, ICTSTYL1.SALES_DIVISION_CODE from SOTINVH2, ICTSTYL1, " & SOTINVH1_temp & " T "
    sql = sql & " where SOTINVH2.INV_TYPE = 'I'"
    sql = sql & " AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE"
    sql = sql & "   and SOTINVH2.INV_NO = T.INV_NO "
    Call Ora_to_Acc(Nothing, "SOWINVH2", 3, "", sql)
    
    Call Track("-", "Pick Tickets")
    sql = "Select SOTPICK1.* from SOTPICK1"
    sql = sql & " where SOTPICK1.PICK_NO in (Select PICK_NO from " & SOTINVH1_temp & ")"
    Call Ora_to_Acc(Nothing, "SOWPICK1", 1, "", sql)
    
    Call Track("-", "Shipments")
    sql = "Select SOTSHIP1.*, DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'MK','G:' || SOTSHIP1.ORDR_GROUP_NO, 'S:' || SOTSHIP1.SHIP_BOL_NO) SHIP_BOL_NO_X, 'N' MULTI_STORE from SOTSHIP1"
    sql = sql & " where SOTSHIP1.SHIP_BOL_NO in (Select DISTINCT SHIP_BOL_NO from " & SOTINVH1_temp & ")"
    Call Ora_to_Acc(Nothing, "SOWSHIP1", 1, "", sql)
    
    Call Track("-", "Verify Integrity")
    sql = "(Select SHIP_BOL_NO, Count (*), Min (INV_NO), Max (INV_NO), Sum (INV_SALES) from " & SOTINVH1_temp & " group by SHIP_BOL_NO)"
    sql = sql & " minus "
    sql = sql & " (Select SOTSHIP1.SHIP_BOL_NO, Count (*), Min (SOTINVH1.INV_NO), Max (SOTINVH1.INV_NO), Sum (SOTINVH1.INV_SALES) from SOTINVH1,SOTSHIP1 where SOTSHIP1.SHIP_BOL_NO = SOTINVH1.SHIP_BOL_NO (+) AND SOTINVH1.ORDR_YYYYPP_UPDATED IS NULL and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from SOTSHIP1 where REGISTER_XNO is Null) group by SOTSHIP1.SHIP_BOL_NO)"
    Set dyn = OraD.CreateDynaset(sql, 8&)
    If Not dyn.EOF Then
        objASCCALLB.xerrmsg = "Shipments Header Record not in synch w/Invoices; Call ABS"
        dyn.Close
        Exit Sub
    Else
        dyn.Close
    End If
    
    sql = "Select * from SOWPICK1 where PICK_STATUS <> 'F'"
    Set dynWK = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    If Not dynWK.EOF Then
        objASCCALLB.xerrmsg = "Unconfirmed Pick Tickets linked to Invoices; Call ABS"
        dynWK.Close
        Exit Sub
    Else
        dynWK.Close
    End If
    
    Call Track("-", "Pick Ticket Details")
    sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE from SOTPICK2,SOTORDR2"
    sql = sql & " where SOTPICK2.PICK_NO in (Select PICK_NO from " & SOTINVH1_temp & ")"
    sql = sql & " and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO"
    sql = sql & " and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO"
    Call Ora_to_Acc(Nothing, "SOWPICK2", 2, "", sql)
    
    sql = "Update SOWINVH1,SOWPICK2"
    sql = sql & " Set SOWINVH1.TOTAL_UNITS = SOWINVH1.TOTAL_UNITS + SOWPICK2.PICK_QTY_CONF"
    sql = sql & " ,   SOWINVH1.TOTAL_UNITS_CANC = SOWINVH1.TOTAL_UNITS_CANC + IIF (ISNULL(SOWPICK2.PICK_QTY_CANC),0,SOWPICK2.PICK_QTY_CANC)"
    sql = sql & " ,   SOWINVH1.TOTAL_UNITS_BACK = SOWINVH1.TOTAL_UNITS_BACK + IIF (ISNULL(SOWPICK2.PICK_QTY_BACK),0,SOWPICK2.PICK_QTY_BACK)"
    sql = sql & " where SOWPICK2.PICK_NO = SOWINVH1.PICK_NO"
    AccD.Execute sql
    
    Call Track("-", "Report Summaries")
    sql = "Select SOWINVH2.SALES_DIVISION_CODE,  SOWINVH1.CUST_CODE,  SOWINVH1.INV_DATE" & vbCrLf
    sql = sql & " , SOWINVH1.SREP_CODE, SOWINVH1.ORDR_GROUP_NO" & vbCrLf
    sql = sql & " , IIF (SOWINVH1.ORDR_ADDR_TYPE_ST = 'MK','G:' & SOWINVH1.ORDR_GROUP_NO, 'S:' & SOWINVH1.SHIP_BOL_NO) AS SHIP_BOL_NO_X" & vbCrLf
    sql = sql & " , IIF (SOWINVH1.ORDR_ADDR_TYPE_ST = 'MK','MK','DC') AS SHIP_ADDR_TYPE" & vbCrLf
    sql = sql & " , IIF (SOWINVH1.ORDR_ADDR_TYPE_ST = 'MK','000000',SOWINVH1.CUST_DC_NO) AS SHIP_ADDR_CODE" & vbCrLf
    sql = sql & " , SOWINVH2.STYLE_CODE, SOWINVH2.COLOR_CODE" & vbCrLf
    sql = sql & " , Sum (SOWINVH2.ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf
    sql = sql & " , Sum (SOWINVH2.ORDR_QTY_SHIP * SOWINVH2.ORDR_UNIT_PRICE) as TOTAL_SALES" & vbCrLf
    sql = sql & " , Sum (SOWINVH2.ORDR_QTY_SHIP * SOWINVH2.ORDR_UNIT_COST) as TOTAL_COSTS" & vbCrLf
    sql = sql & " into SOWINVHD from SOWINVH1,SOWINVH2 " & vbCrLf
    sql = sql & "  where SOWINVH1.INV_TYPE = SOWINVH2.INV_TYPE" & vbCrLf
    sql = sql & "    and SOWINVH1.INV_NO = SOWINVH2.INV_NO" & vbCrLf
    'If El speaks unto you and says something is missing ... remove the line below.
    'sql = sql & "    and INV_NO_REV_BY IS NULL" & vbCrLf
    sql = sql & " group by SOWINVH2.SALES_DIVISION_CODE, SOWINVH1.CUST_CODE,  SOWINVH1.INV_DATE" & vbCrLf
    sql = sql & " , SOWINVH1.SREP_CODE, SOWINVH1.ORDR_GROUP_NO" & vbCrLf
    sql = sql & " , IIF (SOWINVH1.ORDR_ADDR_TYPE_ST = 'MK','G:' & SOWINVH1.ORDR_GROUP_NO, 'S:' & SOWINVH1.SHIP_BOL_NO)" & vbCrLf
    sql = sql & " , IIF (SOWINVH1.ORDR_ADDR_TYPE_ST = 'MK','MK','DC')" & vbCrLf
    sql = sql & " , IIF (SOWINVH1.ORDR_ADDR_TYPE_ST = 'MK','000000',SOWINVH1.CUST_DC_NO)" & vbCrLf
    sql = sql & " , SOWINVH2.STYLE_CODE, SOWINVH2.COLOR_CODE" & vbCrLf
    AccD.Execute sql
    Call Create_Index("SOWINVHD", "I_SOWINVHD_1", "SHIP_BOL_NO_X")
    '--- BEGIN ADDITION OF CANCELLED QTY ADDITION FOR MAURICE - WR - 6/30/05
    Call Track("-", "Cancellations")
    Dim dynSOTCANCG As OraDynaset
    Dim dynSOTCANCS As OraDynaset
    Dim QTY_CANC As Long
    Dim AMT_CANC As Double
    sql = " SELECT SUM(NVL(ORDR_QTY_CANC,0)) QTY_CANC," & vbCrLf
    sql = sql & " SUM(NVL(ORDR_UNIT_PRICE,0) * NVL(ORDR_QTY_CANC,0)) AMT_CANC" & vbCrLf
    sql = sql & " FROM SOTORDR1, SOTORDR2" & vbCrLf
    sql = sql & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf
    sql = sql & " AND ORDR_GROUP_NO = :ORDR_GROUP_NO" & vbCrLf
    Set dynSOTCANCG = OraD.CreateDynaset(sql, 8&)
    
    sql = " SELECT SUM(NVL(ORDR_QTY_CANC,0)) QTY_CANC," & vbCrLf
    sql = sql & " SUM(NVL(ORDR_UNIT_PRICE,0) * NVL(ORDR_QTY_CANC,0)) AMT_CANC" & vbCrLf
    sql = sql & " FROM SOTORDR1, SOTORDR2" & vbCrLf
    sql = sql & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf
    sql = sql & " AND SOTORDR1.ORDR_GROUP_NO IN (" & vbCrLf
    sql = sql & " SELECT DISTINCT ORDR_GROUP_NO" & vbCrLf
    sql = sql & " FROM SOTORDR1 WHERE ORDR_NO IN (" & vbCrLf
    sql = sql & " SELECT ORDR_NO " & vbCrLf
    sql = sql & " FROM SOTPICK1" & vbCrLf
    sql = sql & " WHERE SHIP_BOL_NO = :CODE))" & vbCrLf
    Set dynSOTCANCS = OraD.CreateDynaset(sql, 8&)
    
    sql = "SELECT SHIP_BOL_NO_X,"
    sql = sql & " 0 AS QTY_CANC,"
    sql = sql & " 0.01 AS AMT_CANC"
    sql = sql & " INTO SOWINVHN"
    sql = sql & " FROM SOWINVHD"
    sql = sql & " GROUP BY SHIP_BOL_NO_X"
    AccD.Execute sql
    Call Create_Index("SOWINVHN", "I_SOWINVHN_1", "SHIP_BOL_NO_X")
    Dim dynSOWINVHN As Recordset
    sql = "SELECT * FROM SOWINVHN"
    Set dynSOWINVHN = AccD.OpenRecordset(sql, dbOpenDynaset)
    Do While Not dynSOWINVHN.EOF
        Select Case Left(dynSOWINVHN.Fields("SHIP_BOL_NO_X").Value, 1)
            Case Is = "G"
                OraD.Parameters("ORDR_GROUP_NO").Value = Mid(dynSOWINVHN.Fields("SHIP_BOL_NO_X").Value, 3, 10)
                dynSOTCANCG.Refresh
                AMT_CANC = Val(dynSOTCANCG.Fields("AMT_CANC").Value & "")
                QTY_CANC = Val(dynSOTCANCG.Fields("QTY_CANC").Value & "")
            Case Is = "S"
                OraD.Parameters("CODE").Value = Mid(dynSOWINVHN.Fields("SHIP_BOL_NO_X").Value, 3, 10)
                dynSOTCANCS.Refresh
                AMT_CANC = Val(dynSOTCANCS.Fields("AMT_CANC").Value & "")
                QTY_CANC = Val(dynSOTCANCS.Fields("QTY_CANC").Value & "")
            Case Else
                Stop
                'THIS SHOULD NEVER HAPPEN
        End Select
        dynSOWINVHN.Edit
        dynSOWINVHN.Fields("AMT_CANC").Value = AMT_CANC
        dynSOWINVHN.Fields("QTY_CANC").Value = QTY_CANC
        dynSOWINVHN.Update
        dynSOWINVHN.MoveNext
    Loop
    dynSOWINVHN.Close
    '--- END ADDITION OF CANCELLED QTY ADDITION FOR MAURICE - WR - 6/30/05
    
    sql = "Select SOWSHIP1.SHIP_BOL_NO_X"
    sql = sql & " , Min (SOWSHIP1.SHIP_DATE_SHIPPED) as SHIP_DATE_SHIPPED"
    sql = sql & " , Min (SOWSHIP1.SHIP_ADDR_TYPE) as SHIP_ADDR_TYPE"
    sql = sql & " , Min (SOWSHIP1.SHIP_ADDR_CODE) as SHIP_ADDR_CODE "
    sql = sql & " , Min (SOWSHIP1.SHIP_VIA_CODE) as SHIP_VIA_CODE "
    sql = sql & " into SOWSHIPX from SOWSHIP1 group by SOWSHIP1.SHIP_BOL_NO_X"
    AccD.Execute sql
    Call Create_Index("SOWSHIPX", "I_SOWSHIPX_1", "SHIP_BOL_NO_X")
    
   
    sql = "Select * from SOTORDR0 where ORDR_GROUP_NO in (Select DISTINCT ORDR_GROUP_NO from " & SOTINVH1_temp & ")"
    Call Ora_to_Acc(Nothing, "SOWORDR0", 1, "", sql)
    

    Call Track("-", "Consolidated Invoices")
    
    sql = "SELECT DISTINCT  H1.INV_TYPE, H1.INV_NO, H1.CUST_CODE, H1.CUST_STORE_NO, H1.ORDR_CUST_PO, H1.ORDR_NO,"
    sql = sql & " Sum (H2.ORDR_QTY_SHIP * H2.ORDR_UNIT_PRICE) as TOTAL_SALES,"
    sql = sql & " H1.INV_SALES, H1.INV_FREIGHT,  H1.INV_MISC_CHG,  H1.INV_TOTAL_AMOUNT,"
    sql = sql & " H1.INV_DATE, H1.POST_CODE, H1.SHIP_BOL_NO, H2.SALES_DIVISION_CODE,"
    sql = sql & " H1.INV_NO_CONS , H1.INIT_DATE, H1.PICK_NO , H1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE, H1.GST_TAX"
    sql = sql & " into SOWINVHZ"
    sql = sql & " FROM SOWINVH1 AS H1, SOWINVH2 AS H2"
    sql = sql & " WHERE H1.INV_TYPE = H2.INV_TYPE"
    sql = sql & " AND H1.INV_NO = H2.INV_NO"
    sql = sql & " AND INV_NO_CONS is Null"
    sql = sql & " GROUP BY "
    sql = sql & " H1.INV_TYPE, H1.INV_NO, H1.CUST_CODE, H1.CUST_STORE_NO, H1.ORDR_CUST_PO, H1.ORDR_NO,"
    sql = sql & " H1.INV_SALES, H1.INV_FREIGHT,  H1.INV_MISC_CHG,  H1.INV_TOTAL_AMOUNT,"
    sql = sql & "  H1.INV_DATE, H1.POST_CODE, H1.SHIP_BOL_NO, H2.SALES_DIVISION_CODE,"
    sql = sql & " H1.INV_NO_CONS,  H1.INIT_DATE,H1.PICK_NO, H1.SALES_DIVISION_CODE, H1.GST_TAX"
    AccD.Execute sql
    
    
    sql = " (SELECT DISTINCT  H1.INV_TYPE, H1.INV_NO, H1.CUST_CODE, H1.CUST_STORE_NO, H1.ORDR_CUST_PO, H1.ORDR_NO," & vbCr
    sql = sql & "  Sum (H2.ORDR_QTY_SHIP * H2.ORDR_UNIT_PRICE) as TOTAL_SALES," & vbCr
    sql = sql & "  H1.INV_SALES, H1.INV_FREIGHT,  H1.INV_MISC_CHG,  H1.INV_TOTAL_AMOUNT," & vbCr
    sql = sql & "  H1.INV_DATE, H1.POST_CODE, H1.SHIP_BOL_NO, H2.SALES_DIVISION_CODE," & vbCr
    sql = sql & "  H1.INV_NO_CONS , H1.INIT_DATE, H1.PICK_NO, H1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE, H1.GST_TAX" & vbCr
    sql = sql & " into SOWINVHT"
    sql = sql & "  FROM SOWINVH1 AS H1, SOWINVH2 AS H2" & vbCr
    sql = sql & "  WHERE H1.INV_TYPE = H2.INV_TYPE" & vbCr
    sql = sql & "  AND H1.INV_NO = H2.INV_NO" & vbCr
    sql = sql & "  AND INV_NO_CONS is Not Null" & vbCr
    sql = sql & "  GROUP BY " & vbCr
    sql = sql & "  H1.INV_TYPE, H1.INV_NO, H1.CUST_CODE, H1.CUST_STORE_NO, H1.ORDR_CUST_PO, H1.ORDR_NO," & vbCr
    sql = sql & "  H1.INV_SALES, H1.INV_FREIGHT,  H1.INV_MISC_CHG,  H1.INV_TOTAL_AMOUNT," & vbCr
    sql = sql & "   H1.INV_DATE, H1.POST_CODE, H1.SHIP_BOL_NO, H2.SALES_DIVISION_CODE," & vbCr
    sql = sql & "  H1.INV_NO_CONS,  H1.INIT_DATE,H1.PICK_NO, H1.SALES_DIVISION_CODE, H1.GST_TAX)" & vbCr
    AccD.Execute sql
    
    
    Dim dynSOWINVHZ As Recordset
    Set dynSOWINVHZ = AccD.OpenRecordset("SOWINVHZ", dbOpenDynaset)
    Dim dynARM As Recordset
    sql = " Select INV_NO_CONS,  CUST_CODE, SALES_DIVISION_CODE, H_SALES_DIVISION_CODE" & vbCr
    sql = sql & " , MAX (INV_DATE) AS INV_DT,  " & vbCr
    sql = sql & "   MAX (ORDR_CUST_PO) AS ORDR_PO, " & vbCr
    sql = sql & "   Sum (TOTAL_SALES) as TOT_SALES," & vbCr
    sql = sql & "   SUM (INV_SALES) AS INV_SALE, " & vbCr
    sql = sql & "   SUM (INV_FREIGHT) AS INV_FR, " & vbCr
    sql = sql & "   SUM (INV_MISC_CHG) AS INV_MISC, " & vbCr
    sql = sql & "   SUM (INV_TOTAL_AMOUNT) AS INV_TOT_AMOUNT," & vbCr
    sql = sql & "   SUM (SOWINVHT.GST_TAX) AS GST_TAX" & vbCr
    sql = sql & "  from SOWINVHT" & vbCr
    sql = sql & "  group by INV_NO_CONS, CUST_CODE, SALES_DIVISION_CODE , H_SALES_DIVISION_CODE" & vbCr
    Set dynARM = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Do While Not dynARM.EOF
        dynSOWINVHZ.AddNew
        dynSOWINVHZ.Fields("INV_TYPE").Value = "I"
        dynSOWINVHZ.Fields("INV_NO").Value = dynARM.Fields("INV_NO_CONS").Value
        dynSOWINVHZ.Fields("INV_DATE").Value = dynARM.Fields("INV_DT").Value
        dynSOWINVHZ.Fields("CUST_CODE").Value = dynARM.Fields("CUST_CODE").Value
        dynSOWINVHZ.Fields("SALES_DIVISION_CODE").Value = dynARM.Fields("SALES_DIVISION_CODE").Value
        dynSOWINVHZ.Fields("TOTAL_SALES").Value = dynARM.Fields("TOT_SALES").Value
        dynSOWINVHZ.Fields("H_SALES_DIVISION_CODE").Value = dynARM.Fields("H_SALES_DIVISION_CODE").Value
        dynSOWINVHZ.Fields("ORDR_CUST_PO").Value = dynARM.Fields("ORDR_PO").Value
        dynSOWINVHZ.Fields("INV_SALES").Value = dynARM.Fields("INV_SALE").Value
        dynSOWINVHZ.Fields("INV_FREIGHT").Value = dynARM.Fields("INV_FR").Value
        dynSOWINVHZ.Fields("INV_MISC_CHG").Value = dynARM.Fields("INV_MISC").Value
        dynSOWINVHZ.Fields("INV_TOTAL_AMOUNT").Value = dynARM.Fields("INV_TOT_AMOUNT").Value
        dynSOWINVHZ.Fields("GST_TAX").Value = dynARM.Fields("GST_TAX").Value
        dynSOWINVHZ.Update
        dynARM.MoveNext
    Loop
    dynARM.Close
    
    
    sql = "SELECT DISTINCT INV_NO, 'X' as UPD"
    sql = sql & " Into SOWTEMP"
    sql = sql & " From SOWINVHZ"
    AccD.Execute sql
    
    sql = "Update SOWTEMP, SOWINVHZ set UPD = '1' Where SOWINVHZ.INV_NO = SOWTEMP.INV_NO"
    sql = sql & " and SOWINVHZ.SALES_DIVISION_CODE = SOWINVHZ.H_SALES_DIVISION_CODE"
    AccD.Execute sql
    
    sql = "UPDATE SOWINVHZ, SOWTEMP SET H_SALES_DIVISION_CODE = SALES_DIVISION_CODE "
    sql = sql & " WHERE SOWTEMP.UPD = 'X' AND SOWTEMP.INV_NO = SOWINVHZ.INV_NO"
    AccD.Execute sql
    
    
    sql = "Select SOWINVH1.* into SOWINVHX from SOWINVH1"
    sql = sql & " where INV_NO_CONS is Null"
    AccD.Execute sql
    
    Dim dynSOWINVHX As Recordset
    Set dynSOWINVHX = AccD.OpenRecordset("SOWINVHX", dbOpenDynaset)
    Dim h
    sql = "Select INV_NO_CONS, ORDR_BILL_TO_CUST as CUST_CODE"
    sql = sql & ", MAX (SOWINVH1.INV_DATE) AS INV_DATE, MAX (SOWINVH1.REASON_CODE) as REASON_CODE, MAX (SOWINVH1.SALES_DIVISION_CODE) as SALES_DIVISION_CODE, MAX (SOWINVH1.ORDR_CUST_PO) AS ORDR_CUST_PO"
    sql = sql & ", SUM (SOWINVH1.INV_SALES) AS INV_SALES, SUM (SOWINVH1.INV_FREIGHT) AS INV_FREIGHT, SUM (SOWINVH1.INV_MISC_CHG) AS INV_MISC_CHG, SUM (SOWINVH1.INV_TOTAL_AMOUNT) AS INV_TOTAL_AMOUNT"
    sql = sql & ", SUM (SOWINVH1.TOTAL_UNITS) AS TOTAL_UNITS, SUM (SOWINVH1.TOTAL_UNITS_CANC) AS TOTAL_UNITS_CANC, SUM (SOWINVH1.TOTAL_UNITS_BACK) AS TOTAL_UNITS_BACK"
    sql = sql & ", SUM(SOWINVH1.GST_TAX) AS GST_TAX, MIN(SOWINVH1.CUST_FACTOR_IND) AS CUST_FACTOR_IND, MIN(SOWINVH1.CUST_SURCHARGE_IND) AS CUST_SURCHARGE_IND"
    sql = sql & " from SOWINVH1"
    sql = sql & " where INV_NO_CONS is Not Null"
    sql = sql & " group by INV_NO_CONS, ORDR_BILL_TO_CUST"
    Set dynARM = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Do While Not dynARM.EOF
        dynSOWINVHX.AddNew
        dynSOWINVHX.Fields("INV_TYPE").Value = "I"
        dynSOWINVHX.Fields("INV_NO").Value = dynARM.Fields("INV_NO_CONS").Value
        dynSOWINVHX.Fields("INV_DATE").Value = dynARM.Fields("INV_DATE").Value
        dynSOWINVHX.Fields("CUST_CODE").Value = dynARM.Fields("CUST_CODE").Value
        dynSOWINVHX.Fields("REASON_CODE").Value = dynARM.Fields("REASON_CODE").Value
        dynSOWINVHX.Fields("SALES_DIVISION_CODE").Value = dynARM.Fields("SALES_DIVISION_CODE").Value
        dynSOWINVHX.Fields("ORDR_CUST_PO").Value = dynARM.Fields("ORDR_CUST_PO").Value
        dynSOWINVHX.Fields("INV_SALES").Value = dynARM.Fields("INV_SALES").Value
        dynSOWINVHX.Fields("INV_FREIGHT").Value = dynARM.Fields("INV_FREIGHT").Value
        dynSOWINVHX.Fields("INV_MISC_CHG").Value = dynARM.Fields("INV_MISC_CHG").Value
        dynSOWINVHX.Fields("INV_TOTAL_AMOUNT").Value = dynARM.Fields("INV_TOTAL_AMOUNT").Value
        dynSOWINVHX.Fields("TOTAL_UNITS").Value = dynARM.Fields("TOTAL_UNITS").Value
        dynSOWINVHX.Fields("TOTAL_UNITS_CANC").Value = dynARM.Fields("TOTAL_UNITS_CANC").Value
        dynSOWINVHX.Fields("TOTAL_UNITS_BACK").Value = dynARM.Fields("TOTAL_UNITS_BACK").Value
        dynSOWINVHX.Fields("GST_TAX").Value = dynARM.Fields("GST_TAX").Value
        dynSOWINVHX.Fields("CUST_FACTOR_IND").Value = dynARM.Fields("CUST_FACTOR_IND").Value
        dynSOWINVHX.Fields("CUST_SURCHARGE_IND").Value = dynARM.Fields("CUST_SURCHARGE_IND").Value
'        dynSOWINVHX.Fields("ORDR_SOURCE").Value = "E"
        dynSOWINVHX.Update
        dynARM.MoveNext
    Loop
    dynARM.Close
    
    Call Track("-", "Factored Invoices")
    
    sql = "Select * INTO SOWINVHF FROM SOWINVHX WHERE SOWINVHX.CUST_FACTOR_IND = '1'"
    AccD.Execute sql
    Call Create_Index("SOWINVHF", "PrimaryKey", "INV_TYPE,INV_NO,CUST_CODE")
    sql = "UPDATE SOWINVHF SET SOWINVHF.ORDR_YYYYPP_UPDATED = '" & CYP & "'"
    sql = sql & " where SOWINVHF.ORDR_YYYYPP_UPDATED is Null"
    AccD.Execute sql
    
    Call Track("-", "Master Files")
    
    sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_COST, ICTSTYL1.SALES_DIVISION_CODE from ICTSTYL1"
    Call Ora_to_Acc(Nothing, "ICWSTYL1", 1, "", sql)
    
    sql = "Select * from SOTSDIV1"
    Call Ora_to_Acc(Nothing, "SOWSDIV1", 1, "", sql)
    
End Sub

Sub Update_Invoices()

    Dim LAST_DATE As Date
    LAST_DATE = Now + NowTSD
    
    Dim i As Integer        ' general purpose
    Dim j As Integer        ' general purpose
    Dim z As String         ' working variable
    Dim r As Long           ' Record Counter
    
    Dim INV_TOTAL_AMOUNT As Double  ' Total Invoice Amount
    Dim INV_SALES As Double         ' Total Sales Amt on Invoice
    Dim INV_BALANCE As Double       ' Balance Open on Invoice
    Dim INV_TYPE As String
    Dim INV_NO As String
    
    Dim INV_DATE As String
    Dim INV_DUE_DATE As String
    
    Dim WHSE_CODE As String         ' Ship From Warehouse Code
    Dim CUST_CODE As String         ' Sold-To Customer on Current Order
    Dim ORDR_BILL_TO_CUST As String ' Bill-To Customer for Sold-To Customer on Current Order
    Dim CUST_NAME As String
    
    Dim STYLE_CODE As String        ' Current Style Code
    Dim COLOR_CODE As String        ' Current Color Code
    
    Dim PICK_QTY As Long
    Dim PICK_QTY_CONF As Long
    Dim PICK_QTY_CANC As Long
    Dim PICK_QTY_BACK As Long
    Dim WH_PARM_LOC_SHP As String
    Dim WH_PARM_LOC_SPH As String
    Dim WHSE_TRAN_LNO As Long
    
    Dim ORDR_QTY_OPEN As Long
    
    Dim ORDR_STATUS As String
    Dim PICK_NO As String
    
    Dim ORDR_UNIT_PRICE As Double   ' Net Selling Price of Current Order Detail Item
    Dim STYLE_COST As Double        ' Current Standard Cost of Current Order Detail Item
    
    Dim POST_CODE As String
    Dim TATTERM1 As String          ' Terms Table (built using Make_Table)
    Dim TERM_CODE As String
    Dim TERM_DAYS_DUE As Integer
    Dim TERM_DUE_TYPE As String
    Dim TERM_ADDL_MOS As Integer
    
    Dim dynWK As Recordset
    Dim dyn As OraDynaset
    
    Dim CURR_CODE As String
    Dim CURR_EXCH_RATE As Double
    Dim GST_TAX As Double
    Dim dynTATCURR1 As OraDynaset
    Dim dynARTPARM1 As OraDynaset
    
    sql = "Select * from TATCURR1 where CURR_CODE = :CODE"
    Set dynTATCURR1 = OraD.CreateDynaset(sql, 8&)
    sql = "Select * from ARTPARM1 where AR_PARM_KEY = 'Z'"
    Set dynARTPARM1 = OraD.CreateDynaset(sql, 8&)
        
    Dim dynARTCASH1 As OraDynaset
        sql = "Select * from ARTPYMT1 where ROWNUM < 1"
    Set dynARTCASH1 = OraD.CreateDynaset(sql, 0&)
    
    Dim dynARTCASH2 As OraDynaset
        sql = "Select * from ARTPYMT2 where ROWNUM < 1"
    Set dynARTCASH2 = OraD.CreateDynaset(sql, 0&)
    
    Dim dynARTCASH3 As OraDynaset
        sql = "Select * from ARTPYMT3 where ROWNUM < 1"
    Set dynARTCASH3 = OraD.CreateDynaset(sql, 0&)
            
    Dim PYMT_BATCH_NO As String
    
    
    sql = "Select ORDR_NO, SUM (ORDR_QTY_OPEN) OPEN, SUM (ORDR_QTY_PICK) PICK"
    sql = sql & " from SOTORDR2 where ROWNUM < 1"
    sql = sql & " group by ORDR_NO"
    Dim TT As String
    TT = Temp_Table(sql)
    sql = "Alter Table " & TT & " Add Primary Key (ORDR_NO)"
    OraD.ExecuteSQL sql
    
    OraS.BeginTrans
    sql = "Set transaction use rollback segment RBS_ABS"
    OraD.ExecuteSQL sql
    
    Dim dynSOTORDR1 As OraDynaset
    sql = "Select * from SOTORDR1 where ORDR_NO = :ORDR_NO"
    Set dynSOTORDR1 = OraD.CreateDynaset(sql, 0&)
    
    Dim dynSOTORDR2 As OraDynaset
    sql = "Select * from SOTORDR2 where ORDR_NO = :ORDR_NO"
    sql = sql & " and ORDR_LNO = :ORDR_LNO"
    Set dynSOTORDR2 = OraD.CreateDynaset(sql, 0&)
    
    Dim dynSOTPICK2 As OraDynaset
    sql = "Select * from SOTPICK2 where PICK_NO = :CODE1"
    sql = sql & " and PICK_LNO = :NUM1"
    Set dynSOTPICK2 = OraD.CreateDynaset(sql, 0&)
    
    Dim dynARTCUST1 As OraDynaset
    sql = "Select * from ARTCUST1 where CUST_CODE = :CUST_CODE"
    Set dynARTCUST1 = OraD.CreateDynaset(sql, 0&)
    
    Dim dynARTCUST6 As OraDynaset
    sql = "Select * from ARTCUST6 where CUST_CODE = :CUST_CODE"
    Set dynARTCUST6 = OraD.CreateDynaset(sql, 8&)
    
    Dim dynARTOPEN1 As OraDynaset
    sql = "Select * from ARTOPEN1 where ROWNUM < 1"
    Set dynARTOPEN1 = OraD.CreateDynaset(sql, 8&)
    
    Dim dynARTOPENX As OraDynaset
    sql = "Select Sum (INV_BALANCE) from ARTOPEN1 where CUST_CODE = :CUST_CODE"
    Set dynARTOPENX = OraD.CreateDynaset(sql, 8&)
    
    Dim dynICTSTYL1 As OraDynaset
    sql = "Select * from ICTSTYL1 where STYLE_CODE = :STYLE_CODE"
    Set dynICTSTYL1 = OraD.CreateDynaset(sql, 8&)
    Dim dynICTSTAT1 As OraDynaset
    sql = "Select * from ICTSTAT1 where STYLE_CODE = :STYLE_CODE"
    sql = sql & " and COLOR_CODE = :COLOR_CODE"
    sql = sql & " and WHSE_CODE = :WHSE_CODE"
    sql = sql & " and OPS_YYYYPP = :OPS_YYYYPP"
    Set dynICTSTAT1 = OraD.CreateDynaset(sql, 8&)
    Dim dynICTSTAT2 As OraDynaset
    sql = "Select * from ICTSTAT2 where STYLE_CODE = :STYLE_CODE"
    sql = sql & " and COLOR_CODE = :COLOR_CODE"
    sql = sql & " and WHSE_CODE = :WHSE_CODE"
    Set dynICTSTAT2 = OraD.CreateDynaset(sql, 8&)
    
    Dim dynWHTPARM1 As OraDynaset
    sql = "SELECT WH_PARM_LOC_SHP, WH_PARM_LOC_SPH FROM WHTPARM1 WHERE WH_PARM_KEY = 'Z'"
    Set dynWHTPARM1 = OraD.CreateDynaset(sql, 8&)
    WH_PARM_LOC_SHP = dynWHTPARM1.Fields("WH_PARM_LOC_SHP").Value
    WH_PARM_LOC_SPH = dynWHTPARM1.Fields("WH_PARM_LOC_SPH").Value
    dynWHTPARM1.Close
    Set dynWHTPARM1 = Nothing

    Dim dynWHTLOCB1 As OraDynaset
    sql = "Select * from WHTLOCB1 where WHSE_CODE = :WHSE_CODE" & vbCrLf
    sql = sql & " and LOCATION_CODE = :LOCATION_CODE" & vbCrLf
    sql = sql & " and BAR_CODE = :BAR_CODE" & vbCrLf
    sql = sql & " and STYLE_CODE = :STYLE_CODE" & vbCrLf
    sql = sql & " and COLOR_CODE = :COLOR_CODE" & vbCrLf
    Set dynWHTLOCB1 = OraD.CreateDynaset(sql, 0&)
    
    Dim dynWHTLOCB2 As OraDynaset
    sql = "Select * from WHTLOCB2"
    Set dynWHTLOCB2 = OraD.CreateDynaset(sql, 0&)
    
    Dim WHSE_ACTIVE As Boolean
    Dim WHSE_COUNTING As Boolean
    Dim dynICTWHSE1 As OraDynaset
    sql = "SELECT WHSE_LOCATOR, WHSE_PHYS_STATUS"
    sql = sql & " From ICTWHSE1"
    sql = sql & " WHERE WHSE_CODE = :WHSE_CODE"
    Set dynICTWHSE1 = OraD.CreateDynaset(sql, 8&)
        
    sql = "Select TERM_CODE, TERM_DAYS_DUE, TERM_DUE_TYPE, TERM_ADDL_MOS from TATTERM1"
    TATTERM1 = Make_Table()
        
    Dim dynSOWPICK2 As Recordset
    
    Dim dynSOWINVH1 As Recordset
    sql = "Select * from SOWINVH1 order by INV_NO"
    Set dynSOWINVH1 = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    
    Dim tblICWSTYL1 As Recordset
    Set tblICWSTYL1 = AccD.OpenRecordset("ICWSTYL1", dbOpenTable)
    tblICWSTYL1.Index = "PrimaryKey"
    
    
    Call Track("Adjusting Overshipped Pick Tickets", "")
    
    Dim SHIP_OVER As Long
    Dim SHIP_OVER_ADJ As Long
    sql = "Select * from SOWPICK2 where PICK_QTY_CONF > PICK_QTY"
    Set dynSOWPICK2 = AccD.OpenRecordset(sql, dbOpenDynaset)
    Do While Not dynSOWPICK2.EOF
        SHIP_OVER = Val(dynSOWPICK2.Fields("PICK_QTY_CONF").Value & "") - Val(dynSOWPICK2.Fields("PICK_QTY").Value & "")
        OraD.Parameters("ORDR_NO").Value = dynSOWPICK2.Fields("ORDR_NO").Value
        OraD.Parameters("ORDR_LNO").Value = dynSOWPICK2.Fields("ORDR_LNO").Value
        dynSOTORDR2.Refresh
        ORDR_QTY_OPEN = Val(dynSOTORDR2.Fields("ORDR_QTY_OPEN").Value & "")
        If SHIP_OVER <= ORDR_QTY_OPEN Then
            SHIP_OVER_ADJ = SHIP_OVER
        Else
            SHIP_OVER_ADJ = ORDR_QTY_OPEN
        End If
        If SHIP_OVER_ADJ <> 0 Then
            dynSOWPICK2.Edit
            dynSOWPICK2.Fields("PICK_QTY_BACK").Value = -1 * SHIP_OVER_ADJ
            dynSOWPICK2.Update
            OraD.Parameters("CODE1").Value = dynSOWPICK2.Fields("PICK_NO").Value
            OraD.Parameters("NUM1").Value = dynSOWPICK2.Fields("PICK_LNO").Value
            dynSOTPICK2.Refresh
            dynSOTPICK2.Edit
            dynSOTPICK2.Fields("PICK_QTY_BACK").Value = -1 * SHIP_OVER_ADJ
            dynSOTPICK2.Update
        End If
        dynSOWPICK2.MoveNext
    Loop
    dynSOWPICK2.Close
    
    Call Track("Updating Invoice", "")
    r = 0
    Do While Not dynSOWINVH1.EOF
        r = r + 1
        
        INV_NO = dynSOWINVH1.Fields("INV_NO").Value & ""
        PICK_NO = dynSOWINVH1.Fields("PICK_NO").Value & ""
        
        WHSE_CODE = dynSOWINVH1.Fields("WHSE_CODE").Value & ""
        
        'The Day Vandale Sets up a second Warehouse with another location system
        ' this will have to be expanded.  Hopfully by then they will not be using V1. - W.R. 11/3/14.
        ' OK, Hope is useless.  One more nail in the coffin and Im still alive and here making changes - W.R. 9/8/16
        If WHSE_CODE = "NJE" Or WHSE_CODE = "NJT" Then
            sql = "Begin WHPLOCB2('S','" & INV_NO & "', '" & SessionID & "'); END;"
            OraD.ExecuteSQL sql
        End If
        
        INV_TOTAL_AMOUNT = Val(dynSOWINVH1.Fields("INV_TOTAL_AMOUNT").Value & "")
        INV_SALES = Val(dynSOWINVH1.Fields("INV_SALES").Value & "")
        INV_BALANCE = INV_TOTAL_AMOUNT
        INV_TYPE = dynSOWINVH1.Fields("INV_TYPE").Value & ""
        ORDR_BILL_TO_CUST = dynSOWINVH1.Fields("ORDR_BILL_TO_CUST").Value & ""
        CUST_CODE = dynSOWINVH1.Fields("CUST_CODE").Value & ""
        If ORDR_BILL_TO_CUST = "" Then
            ORDR_BILL_TO_CUST = CUST_CODE
        End If
        INV_DATE = Format$(dynSOWINVH1.Fields("INV_DATE").Value & "", "mm/dd/yyyy")
        WHSE_CODE = dynSOWINVH1.Fields("WHSE_CODE").Value & ""
        
        OraD.Parameters("CUST_CODE").Value = CUST_CODE
        dynARTCUST1.Refresh
        
        'Special Code to Handle Sears Canada.  They get billed in USD but get GST of 5% Added.
        If (CUST_CODE = "SEARSCAN" Or CUST_CODE = "SEARSCANOP") And chkSEARSCAN = "1" Then
            CURR_CODE = "USD"
            CURR_EXCH_RATE = 1
            GST_TAX = 0.05
        Else
            CURR_CODE = dynARTCUST1.Fields("CURR_CODE").Value & ""
            If CURR_CODE = "" Then
                CURR_CODE = "USD"
                CURR_EXCH_RATE = 1
                GST_TAX = 0
            Else
                If dynARTCUST1.Fields("CURR_CODE").Value = dynARTPARM1.Fields("AR_PARM_CURR_CODE").Value Then
                    CURR_CODE = "USD"
                    CURR_EXCH_RATE = 1
                    GST_TAX = 0
                Else
                    OraD.Parameters("CODE").Value = dynARTCUST1.Fields("CURR_CODE").Value
                    dynTATCURR1.Refresh
                    CURR_CODE = dynARTCUST1.Fields("CURR_CODE").Value
                    CURR_EXCH_RATE = dynTATCURR1.Fields("CURR_EXCH_CUR").Value
                    GST_TAX = 0.07
                End If
            End If
        End If
        
        
        If r Mod 100 = 0 Then
            Call Track("-", INV_NO & " - " & CStr(r))
        End If
        
        If TERM_CODE <> dynSOWINVH1.Fields("TERM_CODE").Value & "" Then
            TERM_CODE = dynSOWINVH1.Fields("TERM_CODE").Value & ""
            TERM_DAYS_DUE = Val(SRead(TATTERM1, TERM_CODE, 2))
            TERM_DUE_TYPE = SRead(TATTERM1, TERM_CODE, 3)
            TERM_ADDL_MOS = Val(SRead(TATTERM1, TERM_CODE, 4))
        End If
        'INV_DUE_DATE = Format$(DateValue(INV_DATE) + TERM_DAYS_DUE, "MM/DD/YYYY")
        INV_DUE_DATE = Calc_Due_Date(INV_DATE, TERM_DAYS_DUE, TERM_DUE_TYPE, TERM_ADDL_MOS)
        
        If INV_TYPE = "I" Then
            sql = "Begin SOPORDR2_P('" & PICK_NO & "', '" & INV_NO & "'); END;"
            OraD.ExecuteSQL sql
        End If
        
                
' Update Open A/R Items File
        
        If INV_BALANCE <> 0 And dynSOWINVH1.Fields("INV_NO_CONS").Value & "" = "" Then
        'If INV_BALANCE <> 0 Then
            dynARTOPEN1.AddNew
            dynARTOPEN1.Fields("CUST_CODE").Value = ORDR_BILL_TO_CUST
            dynARTOPEN1.Fields("INV_TYPE").Value = INV_TYPE
            'INV_NO -> INV_NUM change made to accomodate conversion to AR V2. - 8/4/16 W.R.
            'dynARTOPEN1.Fields("INV_NO").Value = INV_NO
            dynARTOPEN1.Fields("INV_NUM").Value = INV_NO
            dynARTOPEN1.Fields("INV_DATE").Value = INV_DATE
            dynARTOPEN1.Fields("CUST_STORE_NO").Value = dynSOWINVH1.Fields("CUST_STORE_NO").Value & ""
            dynARTOPEN1.Fields("POST_CODE").Value = dynSOWINVH1.Fields("POST_CODE").Value & ""
            dynARTOPEN1.Fields("TERM_CODE").Value = dynSOWINVH1.Fields("TERM_CODE").Value & ""
            dynARTOPEN1.Fields("INV_DUE_DATE").Value = INV_DUE_DATE
            dynARTOPEN1.Fields("SREP_CODE").Value = dynSOWINVH1.Fields("SREP_CODE").Value & ""
            dynARTOPEN1.Fields("INV_CUST_PO").Value = dynSOWINVH1.Fields("ORDR_CUST_PO").Value & ""
            'INV_ORDR_NO -> ORDR_NO change made to accomodate conversion to AR V2. - 8/4/16 W.R.
            'dynARTOPEN1.Fields("INV_ORDR_NO").Value = dynSOWINVH1.Fields("ORDR_NO").Value & ""
            dynARTOPEN1.Fields("ORDR_NO").Value = dynSOWINVH1.Fields("ORDR_NO").Value & ""
            dynARTOPEN1.Fields("INV_SALES").Value = dynSOWINVH1.Fields("INV_SALES").Value & ""
            dynARTOPEN1.Fields("INV_DISC").Value = 0
            dynARTOPEN1.Fields("INV_FREIGHT").Value = dynSOWINVH1.Fields("INV_FREIGHT").Value & ""
            dynARTOPEN1.Fields("INV_STAX").Value = 0
            dynARTOPEN1.Fields("INV_MISC_CHG").Value = dynSOWINVH1.Fields("INV_MISC_CHG").Value & ""
            dynARTOPEN1.Fields("INV_TOTAL_AMOUNT").Value = dynSOWINVH1.Fields("INV_TOTAL_AMOUNT").Value & ""
            dynARTOPEN1.Fields("INV_BALANCE").Value = INV_BALANCE
            dynARTOPEN1.Fields("CUST_CODE_SO").Value = CUST_CODE
            dynARTOPEN1.Fields("REASON_CODE").Value = dynSOWINVH1.Fields("REASON_CODE").Value & ""
            dynARTOPEN1.Fields("SALES_DIVISION_CODE").Value = dynSOWINVH1.Fields("SALES_DIVISION_CODE").Value & ""
            dynARTOPEN1.Fields("SEG2_CODE").Value = "000"
            dynARTOPEN1.Fields("SEG3_CODE").Value = "000"
            dynARTOPEN1.Fields("SEG4_CODE").Value = "000"
            dynARTOPEN1.Fields("INIT_OPER").Value = xUserID
            dynARTOPEN1.Fields("INIT_DATE").Value = LAST_DATE
            
            'Added on 2/6/17 per WZ - WR.
            If Format(INV_DATE, "YYYYMM") >= Format(Right(NYM, 2) & "/01/" & Left$(NYM, 4), "YYYYMM") Then
                dynARTOPEN1.Fields("OPS_YYYYPP").Value = NYP
            Else
                dynARTOPEN1.Fields("OPS_YYYYPP").Value = CYP
            End If
            
            'CURRENCY
            If Len(CURR_CODE & "") = 0 Or CURR_EXCH_RATE = 0 Then
                Stop
            End If
            
            CURR_CODE = "USD"
            CURR_EXCH_RATE = 1
            dynARTOPEN1.Fields("GST_TAX").Value = dynSOWINVH1.Fields("GST_TAX").Value
            dynARTOPEN1.Fields("CURR_CODE").Value = CURR_CODE
            dynARTOPEN1.Fields("CURR_EXCH_RATE").Value = CURR_EXCH_RATE
           
           If CURR_CODE = "USD" Then
               dynARTOPEN1.Fields("INV_SALES_CURR").Value = dynSOWINVH1.Fields("INV_SALES").Value
               dynARTOPEN1.Fields("INV_DISC_CURR").Value = 0
               dynARTOPEN1.Fields("INV_FREIGHT_CURR").Value = dynSOWINVH1.Fields("INV_FREIGHT").Value
               dynARTOPEN1.Fields("INV_STAX_CURR").Value = 0
               dynARTOPEN1.Fields("INV_MISC_CHG_CURR").Value = dynSOWINVH1.Fields("INV_MISC_CHG").Value
               dynARTOPEN1.Fields("GST_TAX_CURR").Value = dynSOWINVH1.Fields("GST_TAX").Value
               dynARTOPEN1.Fields("INV_TOTAL_AMOUNT_CURR").Value = dynSOWINVH1.Fields("INV_TOTAL_AMOUNT").Value
               dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = dynSOWINVH1.Fields("INV_TOTAL_AMOUNT").Value
           Else
               dynARTOPEN1.Fields("INV_SALES_CURR").Value = dynSOWINVH1.Fields("INV_SALES_CURR").Value
               dynARTOPEN1.Fields("INV_DISC_CURR").Value = 0
               dynARTOPEN1.Fields("INV_FREIGHT_CURR").Value = dynSOWINVH1.Fields("INV_FREIGHT_CURR").Value
               dynARTOPEN1.Fields("INV_STAX_CURR").Value = 0
               dynARTOPEN1.Fields("INV_MISC_CHG_CURR").Value = dynSOWINVH1.Fields("INV_MISC_CHG_CURR").Value
               'dynARTOPEN1.Fields("GST_TAX").Value = (dynSOWINVH1.Fields("INV_SALES").Value * GST_TAX)
               ''dynARTOPEN1.Fields("GST_TAX").Value = dynSOWINVH1.Fields("GST_TAX").Value
               'dynARTOPEN1.Fields("GST_TAX_CURR").Value = (dynSOWINVH1.Fields("INV_SALES").Value * GST_TAX) / CURR_EXCH_RATE
               dynARTOPEN1.Fields("GST_TAX_CURR").Value = dynSOWINVH1.Fields("GST_TAX_CURR").Value
               dynARTOPEN1.Fields("INV_TOTAL_AMOUNT_CURR").Value = dynSOWINVH1.Fields("INV_TOTAL_AMT_CURR").Value
               'dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = INV_BALANCE / CURR_EXCH_RATE
               dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = dynSOWINVH1.Fields("INV_TOTAL_AMT_CURR").Value
           End If
'            dynARTOPEN1.Fields("CURR_CODE").Value = CURR_CODE
'            dynARTOPEN1.Fields("CURR_EXCH_RATE").Value = CURR_EXCH_RATE
'            dynARTOPEN1.Fields("INV_SALES_CURR").Value = dynSOWINVH1.Fields("INV_SALES_CURR").Value
'            dynARTOPEN1.Fields("INV_DISC_CURR").Value = 0
'            dynARTOPEN1.Fields("INV_FREIGHT_CURR").Value = dynSOWINVH1.Fields("INV_FREIGHT_CURR").Value
'            dynARTOPEN1.Fields("INV_STAX_CURR").Value = 0
'            dynARTOPEN1.Fields("INV_MISC_CHG_CURR").Value = dynSOWINVH1.Fields("INV_MISC_CHG_CURR").Value
'            'dynARTOPEN1.Fields("GST_TAX").Value = (dynSOWINVH1.Fields("INV_SALES").Value * GST_TAX)
'            dynARTOPEN1.Fields("GST_TAX").Value = dynSOWINVH1.Fields("GST_TAX").Value
'            'dynARTOPEN1.Fields("GST_TAX_CURR").Value = (dynSOWINVH1.Fields("INV_SALES").Value * GST_TAX) / CURR_EXCH_RATE
'            dynARTOPEN1.Fields("GST_TAX_CURR").Value = dynSOWINVH1.Fields("GST_TAX_CURR").Value
'            dynARTOPEN1.Fields("INV_TOTAL_AMOUNT_CURR").Value = dynSOWINVH1.Fields("INV_TOTAL_AMT_CURR").Value
'            'dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = INV_BALANCE / CURR_EXCH_RATE
'            dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = dynSOWINVH1.Fields("INV_TOTAL_AMT_CURR").Value
            dynARTOPEN1.Update
        End If
        
        dynSOWINVH1.MoveNext
    Loop
    dynSOWINVH1.Close
    
    
' Update SOTSHIP1

    Call Track("Shipments", "")
    sql = "Update SOTSHIP1 set REGISTER_XNO = '" & xXNO & "'"
    sql = sql & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTINVH1_temp & ")"
    OraD.ExecuteSQL sql
    
    sql = "UPDATE SOTSHIP1 SET CUST_FACTOR_TRANS_IND = '0' WHERE SHIP_BOL_NO IN ("
    sql = sql & " SELECT DISTINCT(SHIP_BOL_NO) FROM " & SOTINVH1_temp & ", ARTCUST1"
    sql = sql & " Where " & SOTINVH1_temp & ".CUST_CODE = ARTCUST1.CUST_CODE"
    sql = sql & " AND ARTCUST1.CUST_FACTOR_TRANS_IND = '1')"
    OraD.ExecuteSQL sql
    
    sql = "Update SOTSHIP1 set SHIP_810_BATCH_NO = 'N'"
    sql = sql & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTINVH1_temp & ")"
    OraD.ExecuteSQL sql
    sql = "Update SOTSHIP1 set SHIP_810_BATCH_NO = NULL"
    sql = sql & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTINVH1_temp & " where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '810' AND EDI_STATUS = 'P') or CUST_CODE = 'WALMARTCOM')"
    'Replaced with the line above for Rick 08/15/07 - WR
    'sql = sql & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTINVH1_temp & " where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '810') or CUST_CODE = 'NORDS1')"
    OraD.ExecuteSQL sql
    sql = "Update SOTSHIP1 set SHIP_856_BATCH_NO = 'N'"
    sql = sql & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTINVH1_temp & ")"
    OraD.ExecuteSQL sql
    sql = "Update SOTSHIP1 set SHIP_856_BATCH_NO = NULL"
    sql = sql & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTINVH1_temp & " where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '856' AND EDI_STATUS = 'P')  or CUST_CODE = 'NORDS1')"
    'Replaced with the line above for Rick 08/15/07 - WR
    'sql = sql & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTINVH1_temp & " where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '856')  or CUST_CODE = 'NORDS1')"
    OraD.ExecuteSQL sql


' Update Order Status

    Call Track("Order Status", "")
    sql = "Insert into " & TT & " Select ORDR_NO, SUM (ORDR_QTY_OPEN) OPEN, SUM (ORDR_QTY_PICK) PICK"
    sql = sql & " from SOTORDR2 where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVH1_temp & ")"
    sql = sql & " group by ORDR_NO"
    OraD.ExecuteSQL sql
    
'    sql = "UPDATE SOTORDR1 SET ORDR_STATUS = 'O' WHERE ORDR_NO IN"
'    sql = sql & " (SELECT ORDR_NO FROM " & TT & " WHERE OPEN <> 0)"
'    OraD.ExecuteSQL sql
'    sql = "UPDATE SOTORDR2 SET ORDR_STATUS = 'O' WHERE ORDR_NO IN"
'    sql = sql & " (SELECT ORDR_NO FROM " & TT & " WHERE OPEN <> 0)"
'    OraD.ExecuteSQL sql
'    sql = "UPDATE SOTORDR1 SET ORDR_STATUS = 'P' WHERE ORDR_NO IN"
'    sql = sql & " (SELECT ORDR_NO FROM " & TT & " WHERE OPEN = 0 AND PICK <> 0)"
'    OraD.ExecuteSQL sql
'    sql = "UPDATE SOTORDR2 SET ORDR_STATUS = 'P' WHERE ORDR_NO IN"
'    sql = sql & " (SELECT ORDR_NO FROM " & TT & " WHERE OPEN = 0 AND PICK <> 0)"
'    OraD.ExecuteSQL sql
'    sql = "UPDATE SOTORDR1 SET ORDR_STATUS = 'F', ORDR_YYYYPP_CLOSED = '" & CYP & "', ORDR_DATE_CLOSED = SYSDATE WHERE ORDR_NO IN"
'    sql = sql & " (SELECT ORDR_NO FROM " & TT & " WHERE OPEN = 0 AND PICK = 0)"
'    OraD.ExecuteSQL sql
'    sql = "UPDATE SOTORDR2 SET ORDR_STATUS = 'F' WHERE ORDR_NO IN"
'    sql = sql & " (SELECT ORDR_NO FROM " & TT & " WHERE OPEN = 0 AND PICK = 0)"
'    OraD.ExecuteSQL sql
    
    
    sql = "BEGIN DECLARE CURSOR C1 IS SELECT ORDR_NO FROM " & TT & " WHERE OPEN <> 0;"
    sql = sql & " BEGIN FOR R1 IN C1 LOOP"
    sql = sql & " UPDATE SOTORDR1 SET ORDR_STATUS = 'O' WHERE ORDR_NO = R1.ORDR_NO;"
    sql = sql & " UPDATE SOTORDR2 SET ORDR_STATUS = 'O' WHERE ORDR_NO = R1.ORDR_NO;"
    sql = sql & " END LOOP; END; END;"
    OraD.ExecuteSQL sql
    
    sql = "BEGIN DECLARE CURSOR C1 IS SELECT ORDR_NO FROM " & TT & " WHERE OPEN = 0 AND PICK <> 0;"
    sql = sql & " BEGIN FOR R1 IN C1 LOOP"
    sql = sql & " UPDATE SOTORDR1 SET ORDR_STATUS = 'P' WHERE ORDR_NO = R1.ORDR_NO;"
    sql = sql & " UPDATE SOTORDR2 SET ORDR_STATUS = 'P' WHERE ORDR_NO = R1.ORDR_NO;"
    sql = sql & " END LOOP; END; END;"
    OraD.ExecuteSQL sql
    
    sql = "BEGIN DECLARE CURSOR C1 IS SELECT ORDR_NO FROM " & TT & " WHERE OPEN = 0 AND PICK = 0;"
    sql = sql & " BEGIN FOR R1 IN C1 LOOP"
    sql = sql & " UPDATE SOTORDR1 SET ORDR_STATUS = 'F', ORDR_YYYYPP_CLOSED = '" & CYP & "', ORDR_DATE_CLOSED = SYSDATE WHERE ORDR_NO = R1.ORDR_NO;"
    sql = sql & " UPDATE SOTORDR2 SET ORDR_STATUS = 'F' WHERE ORDR_NO = R1.ORDR_NO;"
    sql = sql & " END LOOP; END; END;"
    OraD.ExecuteSQL sql
    
    
' Update Style Status

    Call Track("Style Status Update", "")
    
    sql = "Select SOWINVH1.ORDR_YYYYPP_UPDATED, SOWINVH1.WHSE_CODE, SOWPICK2.STYLE_CODE, "
    sql = sql & " SOWPICK2.COLOR_CODE, SOWPICK1.SHIP_BOL_NO"
    sql = sql & ", SUM (SOWPICK2.PICK_QTY) as PICK_QTY"
    sql = sql & ", SUM (SOWPICK2.PICK_QTY_CONF) as PICK_QTY_CONF "
    sql = sql & ", SUM (SOWPICK2.PICK_QTY_BACK) as PICK_QTY_BACK"
    sql = sql & " from SOWPICK1,SOWPICK2,SOWINVH1"
    sql = sql & " where SOWPICK1.PICK_NO = SOWPICK2.PICK_NO"
    sql = sql & "   and SOWINVH1.INV_TYPE = 'I'"
    sql = sql & "   and SOWINVH1.INV_NO = SOWPICK1.INV_NO"
    sql = sql & " group by SOWINVH1.ORDR_YYYYPP_UPDATED, SOWINVH1.WHSE_CODE, SOWPICK2.STYLE_CODE,"
    sql = sql & " SOWPICK2.COLOR_CODE, SOWPICK1.SHIP_BOL_NO"
    Dim dynSOWPICKX As Recordset
    Set dynSOWPICKX = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Do While Not dynSOWPICKX.EOF
        WHSE_TRAN_LNO = WHSE_TRAN_LNO + 1
        RYP = dynSOWPICKX.Fields("ORDR_YYYYPP_UPDATED").Value & ""
        WHSE_CODE = dynSOWPICKX.Fields("WHSE_CODE").Value & ""
        STYLE_CODE = dynSOWPICKX.Fields("STYLE_CODE").Value & ""
        COLOR_CODE = dynSOWPICKX.Fields("COLOR_CODE").Value & ""
        PICK_QTY = dynSOWPICKX.Fields("PICK_QTY").Value & ""
        PICK_QTY_CONF = dynSOWPICKX.Fields("PICK_QTY_CONF").Value & ""
        PICK_QTY_BACK = dynSOWPICKX.Fields("PICK_QTY_BACK").Value & ""
        
        Call Track("-", STYLE_CODE)
        
        OraD.Parameters("STYLE_CODE").Value = STYLE_CODE
        OraD.Parameters("COLOR_CODE").Value = COLOR_CODE
        OraD.Parameters("WHSE_CODE").Value = WHSE_CODE
        OraD.Parameters("OPS_YYYYPP").Value = RYP
        OraD.Parameters("BAR_CODE").Value = "0000000000"

        
        dynICTSTAT2.Refresh
        If dynICTSTAT2.EOF Then
            dynICTSTAT2.AddNew
            dynICTSTAT2.Fields("STYLE_CODE").Value = STYLE_CODE
            dynICTSTAT2.Fields("COLOR_CODE").Value = COLOR_CODE
            dynICTSTAT2.Fields("WHSE_CODE").Value = WHSE_CODE
        Else
            dynICTSTAT2.Edit
        End If
        dynICTSTAT2.Fields("WHSE_QTY_OPEN").Value = Val(dynICTSTAT2.Fields("WHSE_QTY_OPEN").Value & "") + PICK_QTY_BACK
        dynICTSTAT2.Fields("WHSE_QTY_PICK").Value = Val(dynICTSTAT2.Fields("WHSE_QTY_PICK").Value & "") - PICK_QTY
        dynICTSTAT2.Fields("WHSE_QTY_ON_HAND").Value = Val(dynICTSTAT2.Fields("WHSE_QTY_ON_HAND").Value & "") - PICK_QTY_CONF
        dynICTSTAT2.Update
        
        dynICTSTAT1.Refresh
        If dynICTSTAT1.EOF Then
            dynICTSTAT1.AddNew
            dynICTSTAT1.Fields("OPS_YYYYPP").Value = RYP
            dynICTSTAT1.Fields("STYLE_CODE").Value = STYLE_CODE
            dynICTSTAT1.Fields("COLOR_CODE").Value = COLOR_CODE
            dynICTSTAT1.Fields("WHSE_CODE").Value = WHSE_CODE
        Else
            dynICTSTAT1.Edit
        End If
        dynICTSTAT1.Fields("WHSE_QTY_SHP").Value = Val(dynICTSTAT1.Fields("WHSE_QTY_SHP").Value & "") + PICK_QTY_CONF
        dynICTSTAT1.Update
        
        'Update Shipping Location for Location System.
        dynICTWHSE1.Refresh
'        If dynICTWHSE1.Fields("WHSE_LOCATOR").Value = "1" Then
'            WHSE_COUNTING = True
'            OraD.Parameters("LOCATION_CODE").Value = WH_PARM_LOC_SPH
'        Else
'            WHSE_COUNTING = False
'            OraD.Parameters("LOCATION_CODE").Value = WH_PARM_LOC_SHP
'        End If
        
        If dynICTWHSE1.Fields("WHSE_PHYS_STATUS").Value = "1" Then
            WHSE_COUNTING = True
            OraD.Parameters("LOCATION_CODE").Value = WH_PARM_LOC_SPH
        Else
            WHSE_COUNTING = False
            OraD.Parameters("LOCATION_CODE").Value = WH_PARM_LOC_SHP
        End If
                
'        If dynICTWHSE1.Fields("WHSE_LOCATOR").Value = "1" Then
'            dynWHTLOCB1.Refresh
'            If dynWHTLOCB1.EOF Then
'                dynWHTLOCB1.AddNew
'                dynWHTLOCB1.Fields("WHSE_CODE").Value = WHSE_CODE
'                If WHSE_COUNTING Then
'                    dynWHTLOCB1.Fields("LOCATION_CODE").Value = WH_PARM_LOC_SPH
'                Else
'                    dynWHTLOCB1.Fields("LOCATION_CODE").Value = WH_PARM_LOC_SHP
'                End If
'                dynWHTLOCB1.Fields("BAR_CODE").Value = "0000000000"
'                dynWHTLOCB1.Fields("STYLE_CODE").Value = STYLE_CODE
'                dynWHTLOCB1.Fields("COLOR_CODE").Value = COLOR_CODE
'                dynWHTLOCB1.Fields("LOCATION_QTY").Value = PICK_QTY_CONF * -1
'            Else
'                dynWHTLOCB1.Edit
'                dynWHTLOCB1.Fields("LOCATION_QTY").Value = dynWHTLOCB1.Fields("LOCATION_QTY").Value - PICK_QTY_CONF
'            End If
'            dynWHTLOCB1.Update
'
'            'dynWHTLOCB2.Refresh
'            dynWHTLOCB2.AddNew
'            dynWHTLOCB2.Fields("WHSE_CODE").Value = WHSE_CODE
'            If WHSE_COUNTING Then
'                dynWHTLOCB2.Fields("LOCATION_CODE").Value = WH_PARM_LOC_SPH
'            Else
'                dynWHTLOCB2.Fields("LOCATION_CODE").Value = WH_PARM_LOC_SHP
'            End If
'            dynWHTLOCB2.Fields("BAR_CODE").Value = "0000000000"
'            dynWHTLOCB2.Fields("STYLE_CODE").Value = STYLE_CODE
'            dynWHTLOCB2.Fields("COLOR_CODE").Value = COLOR_CODE
'            dynWHTLOCB2.Fields("WHSE_TRAN_QTY").Value = PICK_QTY_CONF * -1
'            dynWHTLOCB2.Fields("WHSE_TRAN_TYPE").Value = "S"
'            dynWHTLOCB2.Fields("WHSE_TRAN_NO").Value = dynSOWPICKX.Fields("SHIP_BOL_NO").Value
'            dynWHTLOCB2.Fields("WHSE_TRAN_LNO").Value = WHSE_TRAN_LNO
'            dynWHTLOCB2.Fields("INIT_DATE").Value = Now + NowTSD
'            dynWHTLOCB2.Fields("INIT_OPER").Value = xUserID
'            dynWHTLOCB2.Fields("LOCATION_CODE_OTHER").Value = ""
'            dynWHTLOCB2.Fields("SESSION_ID").Value = ""
'            dynWHTLOCB2.Update
'        End If
        dynSOWPICKX.MoveNext
        
    Loop
    dynSOWPICKX.Close
    

' Record Consolidated Invoice

    Call Track("Consolidated Invoice", "")
    Dim dynARM As Recordset
    sql = "Select SOWINVH1.INV_NO_CONS, SOWINVH1.ORDR_BILL_TO_CUST as CUST_CODE"
    sql = sql & ", MAX (SOWINVH1.INV_DATE) AS INV_DATE, MAX (SOWINVH1.REASON_CODE) as REASON_CODE, MAX (SOWINVH1.SALES_DIVISION_CODE) as SALES_DIVISION_CODE, MAX (SOWINVH1.ORDR_CUST_PO) AS ORDR_CUST_PO"
    sql = sql & ", SUM (SOWINVH1.INV_SALES) AS INV_SALES, SUM (SOWINVH1.INV_FREIGHT) AS INV_FREIGHT, SUM (SOWINVH1.INV_MISC_CHG) AS INV_MISC_CHG, SUM (SOWINVH1.INV_TOTAL_AMOUNT) AS INV_TOTAL_AMOUNT, MAX(SOWINVH1.EDI_APPOINTMENT) AS EDI_APPOINTMENT"
    sql = sql & ", SUM (SOWINVH1.INV_SALES_CURR) AS INV_SALES_CURR, SUM (SOWINVH1.INV_FREIGHT_CURR) AS INV_FREIGHT_CURR, SUM (SOWINVH1.INV_MISC_CHG_CURR) AS INV_MISC_CHG_CURR, SUM (SOWINVH1.INV_TOTAL_AMOUNT_CURR) AS INV_TOTAL_AMOUNT_CURR"
    sql = sql & " from SOWINVH1"
    sql = sql & " where INV_NO_CONS is Not Null"
    'sql = sql & "   and SOWORDR1.ORDR_NO = SOWINVH1.ORDR_NO"
    sql = sql & " group by SOWINVH1.INV_NO_CONS, SOWINVH1.ORDR_BILL_TO_CUST"
    Set dynARM = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Do While Not dynARM.EOF
        GoSub Consolidated_Invoice
        dynARM.MoveNext
    Loop
    dynARM.Close
    
    
' Record Sales Summary

    Call Track("Customer Sales Summary", "")
    sql = "Select SOWINVH1.CUST_CODE, Count (*) as INVOICES, "
    sql = sql & " Sum (SOWINVH1.INV_SALES) as INV_SALES, "
    sql = sql & " Max (SOWINVH1.INV_DATE) as LAST_DATE, "
    sql = sql & " Min (SOWINVH1.INV_DATE) as FIRST_DATE "
    sql = sql & " from SOWINVH1 "
    sql = sql & " where INV_DATE < #" & Right(NYM, 2) & "/01/" & Left$(NYM, 4) & "# group by CUST_CODE "
    Set dynWK = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Dim dynWK2 As Recordset
    Do While Not dynWK.EOF
        OraD.Parameters("CUST_CODE").Value = dynWK.Fields("CUST_CODE").Value & ""
        dynARTCUST6.Refresh
        If dynARTCUST6.EOF Then
            dynARTCUST6.AddNew
            dynARTCUST6.Fields("CUST_CODE").Value = OraD.Parameters("CUST_CODE").Value
        Else
            dynARTCUST6.Edit
        End If
        sql = "Select * from SOWINVH1 where CUST_CODE = '" & dynWK.Fields("CUST_CODE").Value & "'"
        sql = sql & " and INV_DATE = #" & Format$(dynWK.Fields("LAST_DATE").Value & "", "MM/DD/YYYY") & "#"
        Set dynWK2 = AccD.OpenRecordset(sql, dbOpenForwardOnly)
        dynARTCUST6.Fields("CUST_LAST_INV_NUM").Value = dynWK2.Fields("INV_NO").Value
        dynARTCUST6.Fields("CUST_LAST_INV_DATE").Value = dynWK2.Fields("INV_DATE").Value
        dynARTCUST6.Fields("CUST_LAST_INV_AMT").Value = Val(dynWK2.Fields("INV_TOTAL_AMOUNT").Value & "")
        dynWK2.Close
        If IsNull(dynARTCUST6.Fields("CUST_FIRST_PURCH").Value) Then
            dynARTCUST6.Fields("CUST_FIRST_PURCH").Value = dynWK.Fields("FIRST_DATE").Value
        End If
'        If INV_TYPE = "I" Then
            dynARTCUST6.Fields("CUST_SALES_MTD").Value = Val(dynARTCUST6.Fields("CUST_SALES_MTD").Value & "") + Val(dynWK.Fields("INV_SALES").Value & "")
            dynARTCUST6.Fields("CUST_SALES_YTD").Value = Val(dynARTCUST6.Fields("CUST_SALES_YTD").Value & "") + Val(dynWK.Fields("INV_SALES").Value & "")
            dynARTCUST6.Fields("CUST_NUM_INV_MTD").Value = Val(dynARTCUST6.Fields("CUST_NUM_INV_MTD").Value & "") + Val(dynWK.Fields("INVOICES").Value & "")
            dynARTCUST6.Fields("CUST_NUM_INV_YTD").Value = Val(dynARTCUST6.Fields("CUST_NUM_INV_YTD").Value & "") + Val(dynWK.Fields("INVOICES").Value & "")
'        Else
'            dynARTCUST6.Fields("CUST_CRED_MTD").Value = Val(dynARTCUST6.Fields("CUST_CRED_MTD").Value & "") + Val(dynwk.Fields("INV_SALES").Value & "")
'            dynARTCUST6.Fields("CUST_CRED_YTD").Value = Val(dynARTCUST6.Fields("CUST_CRED_YTD").Value & "") + Val(dynwk.Fields("INV_SALES").Value & "")
'        End If
        dynARTCUST6.Update
        dynWK.MoveNext
    Loop
    dynWK.Close
    
    
' A/R Summary
    
    Call Track("Customer A/R Summary", "")
    sql = "Select ORDR_BILL_TO_CUST, MAX(INV_DATE) from SOWINVH1 group by ORDR_BILL_TO_CUST"
    Dim dynAR As Recordset
    Set dynAR = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Do While Not dynAR.EOF
        OraD.Parameters("CUST_CODE").Value = dynAR.Fields(0).Value
        dynARTOPENX.Refresh
        dynARTCUST6.Refresh
        If dynARTCUST6.EOF Then
            dynARTCUST6.AddNew
            dynARTCUST6.Fields("CUST_CODE").Value = dynAR.Fields(0).Value
            dynARTCUST6.Update
            dynARTCUST6.Refresh
        End If
        
        If Val(dynARTOPENX.Fields(0).Value & "") > Val(dynARTCUST6.Fields("CUST_HIGH_BAL_AMT").Value & "") Then
            dynARTCUST6.Edit
            dynARTCUST6.Fields("CUST_HIGH_BAL_AMT").Value = Val(dynARTOPENX.Fields(0).Value & "")
            dynARTCUST6.Fields("CUST_HIGH_BAL_DATE").Value = dynAR.Fields(1).Value
            dynARTCUST6.Update
        End If
        dynAR.MoveNext
    Loop
    dynAR.Close
    

' Stamp Ops Period

    Call Track("Mark Records with Period", "")
    
    sql = "Update SOTINVH1 set ORDR_YYYYPP_UPDATED = '" & CYP & "'"
    sql = sql & ", ORDR_DATE_UPDATED = SYSDATE, REGISTER_XNO = '" & xXNO & "'"
    sql = sql & " where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1_temp & ") AND "
    sql = sql & "INV_DATE < '" & Format(Right(NYM, 2) & "/01/" & Left$(NYM, 4), "DD-MMM-YY") & "'"
    OraD.ExecuteSQL (sql)
    'sql = "Update SOTINVH2 set ORDR_YYYYPP_UPDATED = '" & CYP & "'"
    'sql = sql & " where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1_temp & ") AND "
    'sql = sql & "INV_DATE >= '" & Format(Right(CYP, 2) & "/01/" & Left$(CYP, 4), "DD-MMM-YY") & "' AND INV_DATE <= '" & Format(dynGLTPARM2.Fields("PRD_END_DATE").Value, "DD-MMM-YY") & "' "
    sql = "Update SOTINVH2 set ORDR_YYYYPP_UPDATED = '" & CYP & "' where INV_TYPE = 'I'"
    sql = sql & " and INV_NO in (SELECT INV_NO FROM SOTINVH1 where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1_temp & ") AND INV_DATE >= '" & Format(Right(CYP, 2) & "/01/" & Left$(CYP, 4), "DD-MMM-YY") & "'"
    sql = sql & " AND INV_DATE < '" & Format(Right(NYM, 2) & "/01/" & Left$(NYM, 4), "DD-MMM-YY") & "')"
    OraD.ExecuteSQL (sql)

    'THESE SQL STATEMENT WILL UPDATE INVOICES FOR PERIODS NOT IN THE CURRENT CYP
    
    sql = "Update SOTINVH1 set ORDR_YYYYPP_UPDATED = '" & NYP & "'"
    sql = sql & ", ORDR_DATE_UPDATED = SYSDATE, REGISTER_XNO = '" & xXNO & "'"
    sql = sql & " where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1_temp & ") AND "
    sql = sql & "INV_DATE >= '" & Format(Right(NYM, 2) & "/01/" & Left$(NYM, 4), "DD-MMM-YY") & "'"
    OraD.ExecuteSQL (sql)
    
    sql = "Update SOTINVH2 set ORDR_YYYYPP_UPDATED = '" & NYP & "' where INV_TYPE = 'I'"
    sql = sql & " and INV_NO in (SELECT INV_NO FROM SOTINVH1 where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1_temp & ") AND INV_DATE >= '" & Format(Right(NYP, 2) & "/01/" & Left$(NYP, 4), "DD-MMM-YY") & "'"
    sql = sql & " AND INV_DATE >= '" & Format(Right(NYM, 2) & "/01/" & Left$(NYM, 4), "DD-MMM-YY") & "')"
    OraD.ExecuteSQL (sql)
' Order Group Summary
    
    
    
    Call Track("Updating Order Group Summary", "")
    sql = "Select Distinct ORDR_GROUP_NO from SOWSHIP1 "
    Set dynWK = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Do While Not dynWK.EOF
        sql = "BEGIN SOPORDR0_G('" & dynWK.Fields(0).Value & "'); END;"
        OraD.ExecuteSQL sql
        dynWK.MoveNext
    Loop
    dynWK.Close
    
    
' Create a Cash Receipts Batch Matching off All ReverSals with Originals

    dynARTOPEN1.Close
    sql = "Select * from ARTOPEN1 where CUST_CODE = :CUST_CODE "
    'INV_NO -> INV_NUM change made to accomodate conversion to AR V2. - 8/4/16 W.R.
    'sql = sql & " and INV_TYPE = 'I' and INV_NO = :CODE"
    sql = sql & " and INV_TYPE = 'I' and INV_NUM = :CODE"
    Set dynARTOPEN1 = OraD.CreateDynaset(sql, 8&)
    
    Dim dynSOTINVH1_CONS As OraDynaset
    sql = "Select * from SOTINVH1 where INV_TYPE = 'I' and INV_NO = :CODE"
    Set dynSOTINVH1_CONS = OraD.CreateDynaset(sql, 8&)
    
    Dim SHIP_BOL_NO As String
    SHIP_BOL_NO = ""
    
    Dim PYMT_BATCH_ILNO_ctr As Integer
    Dim INV_BALANCE_expected As Double
    Dim INV_BALANCE_expected_CURR As Double
   
    sql = "Select * from SOWINVH1 where INV_NO_REV is Not Null order by SHIP_BOL_NO"
    Set dynSOWINVH1 = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    Do While Not dynSOWINVH1.EOF
        CUST_CODE = dynSOWINVH1.Fields("CUST_CODE").Value
        INV_DATE = dynSOWINVH1.Fields("INV_DATE").Value
        If dynSOWINVH1.Fields("SHIP_BOL_NO").Value <> SHIP_BOL_NO Then
            SHIP_BOL_NO = dynSOWINVH1.Fields("SHIP_BOL_NO").Value
            RYP = dynSOWINVH1.Fields("ORDR_YYYYPP_UPDATED").Value
            GoSub Update_ARTCASH1
        End If
        INV_NO = dynSOWINVH1.Fields("INV_NO").Value
        INV_BALANCE_expected = Val(dynSOWINVH1.Fields("INV_TOTAL_AMOUNT").Value & "")
        INV_BALANCE_expected_CURR = Val(dynSOWINVH1.Fields("INV_TOTAL_AMOUNT_CURR").Value & "")
        GoSub Update_ARTCASH3
        INV_BALANCE_expected = -1 * Val(dynSOWINVH1.Fields("INV_TOTAL_AMOUNT").Value & "")
        INV_BALANCE_expected_CURR = -1 * Val(dynSOWINVH1.Fields("INV_TOTAL_AMT_CURR").Value & "")
        INV_NO = dynSOWINVH1.Fields("INV_NO_REV").Value
        GoSub Update_ARTCASH3
        dynSOWINVH1.MoveNext
    Loop
    dynSOWINVH1.Close
    
    'Update New Control File To Force SJ&U between Confirm and Deconfirm - WR - 20051024
    sql = "UPDATE SOTCTLU1"
    sql = sql & " SET CTL_UPDATE_REQ = 'U'"
    sql = sql & " WHERE UPPER(CTL_KEY) = 'Z'"
    OraD.ExecuteSQL (sql)
    
    Dim dynSOWINVHT As Recordset
    sql = "Select * from SOWINVH1 WHERE ORDR_TYPE_CODE = 'XFR'"
    Set dynSOWINVHT = AccD.OpenRecordset(sql, dbOpenDynaset)
    Do While Not dynSOWINVHT.EOF
        Call Update_Transfer(dynSOWINVHT.Fields("INV_NO").Value)
        dynSOWINVHT.MoveNext
    Loop
    dynSOWINVHT.Close
    
    OraS.CommitTrans
Exit Sub

Update_ARTCASH1:
        PYMT_BATCH_NO = CTLNO("ARTPYMT1.PYMT_BATCH_NO", 10)
    
    OraD.Parameters("CUST_CODE").Value = CUST_CODE
    dynARTCUST1.Refresh
    CUST_NAME = dynARTCUST1.Fields("CUST_NAME").Value
    
    dynARTCASH1.AddNew
    dynARTCASH1.Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    dynARTCASH1.Fields("PYMT_BATCH_DATE").Value = INV_DATE
    dynARTCASH1.Fields("BANK_CODE").Value = Null
    dynARTCASH1.Fields("STATUS").Value = "1"
    dynARTCASH1.Fields("INIT_OPER").Value = xUserID
    dynARTCASH1.Fields("LAST_OPER").Value = xUserID
    dynARTCASH1.Fields("INIT_DATE").Value = LAST_DATE
    dynARTCASH1.Fields("LAST_DATE").Value = LAST_DATE
        dynARTCASH1.Fields("PYMT_APPL_ONLY").Value = "1"
    dynARTCASH1.Fields("OPS_YYYYPP").Value = RYP
        dynARTCASH1.Fields("CURR_CODE").Value = "USD"
        dynARTCASH1.Fields("CURR_EXCH_RATE").Value = 1
        dynARTCASH1.Fields("PYMT_SOURCE").Value = "K"
    dynARTCASH1.Update
    
    dynARTCASH2.AddNew
    dynARTCASH2.Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    dynARTCASH2.Fields("PYMT_BATCH_LNO").Value = 1
    dynARTCASH2.Fields("CUST_CODE").Value = CUST_CODE
    dynARTCASH2.Fields("CUST_NAME").Value = CUST_NAME
        dynARTCASH2.Fields("CUST_PYMT_REF_NO").Value = "AUTO"
        dynARTCASH2.Fields("CUST_PYMT_REF_DATE").Value = INV_DATE
        dynARTCASH2.Fields("CUST_PYMT_AMT").Value = 0
        dynARTCASH2.Fields("PYMT_STATUS").Value = "2"
        dynARTCASH2.Fields("CUST_PYMT_AMT_CURR").Value = 0
        dynARTCASH2.Fields("CURR_CODE").Value = "USD"
        dynARTCASH2.Fields("CURR_EXCH_RATE").Value = 1
    dynARTCASH2.Update
    
    PYMT_BATCH_ILNO_ctr = 0
Return
            
Update_ARTCASH3:
    OraD.Parameters("CODE").Value = INV_NO
    dynSOTINVH1_CONS.Refresh
    If dynSOTINVH1_CONS.Fields("INV_NO_CONS").Value & "" <> "" Then
        INV_NO = dynSOTINVH1_CONS.Fields("INV_NO_CONS").Value & ""
    End If
    
    OraD.Parameters("CODE").Value = INV_NO
    dynARTOPEN1.Refresh
    If dynARTOPEN1.EOF Then
            MsgBox "Can Not Find Invoice " & INV_NO & " Required for Update!!", vbCritical, "Update Error"
            Stop ' PROBLEM WITH FINDING INVOICE NUMBER TO NET TO ZERO
    ElseIf dynSOTINVH1_CONS.Fields("INV_NO_CONS").Value & "" = "" And INV_BALANCE_expected <> Val(dynARTOPEN1.Fields("INV_BALANCE").Value & "") Then
        Stop    ' PROBLEM WITH FINDING INVOICE BALANCE EXPECTED
                ' WR - 2/16/04 - You may be finiding this problem when you they are de-confirming an order that
                ' has had cash applied against it already.  If this is true AND they are deconfirming in order to
                ' change something other than anounts and re-confirm then it is OK to step through here.
    End If
    
    If Val(dynARTOPEN1.Fields("INV_BALANCE").Value & "") <> 0 Then
        dynARTCASH3.AddNew
        dynARTCASH3.Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
        dynARTCASH3.Fields("PYMT_BATCH_LNO").Value = 1
        PYMT_BATCH_ILNO_ctr = PYMT_BATCH_ILNO_ctr + 1
        dynARTCASH3.Fields("PYMT_BATCH_ILNO").Value = PYMT_BATCH_ILNO_ctr
        dynARTCASH3.Fields("INV_TYPE").Value = dynARTOPEN1.Fields("INV_TYPE").Value
        'INV_NO -> INV_NUM change made to accomodate conversion to AR V2. - 8/4/16 W.R.
        'dynARTCASH3.Fields("INV_NO").Value = dynARTOPEN1.Fields("INV_NO").Value
            dynARTCASH3.Fields("INV_NUM").Value = dynARTOPEN1.Fields("INV_NUM").Value
        dynARTCASH3.Fields("REASON_CODE").Value = dynARTOPEN1.Fields("REASON_CODE").Value
        dynARTCASH3.Fields("INV_DATE").Value = dynARTOPEN1.Fields("INV_DATE").Value
        dynARTCASH3.Fields("INV_DUE_DATE").Value = dynARTOPEN1.Fields("INV_DUE_DATE").Value
        dynARTCASH3.Fields("CUST_CODE_SO").Value = dynARTOPEN1.Fields("CUST_CODE_SO").Value
        dynARTCASH3.Fields("CUST_STORE_NO").Value = dynARTOPEN1.Fields("CUST_STORE_NO").Value
        dynARTCASH3.Fields("INV_CUST_PO").Value = dynARTOPEN1.Fields("INV_CUST_PO").Value
        dynARTCASH3.Fields("INV_PMT").Value = INV_BALANCE_expected 'dynARTOPEN1.Fields("INV_BALANCE").Value
        dynARTCASH3.Fields("INV_DISC_TAKEN").Value = 0
        dynARTCASH3.Fields("INV_WRITE_OFF").Value = 0
        dynARTCASH3.Fields("INV_BALANCE_NEW").Value = Val(dynARTOPEN1.Fields("INV_BALANCE").Value & "") - INV_BALANCE_expected
        dynARTCASH3.Fields("POST_CODE").Value = dynARTOPEN1.Fields("POST_CODE").Value
        dynARTCASH3.Fields("SEG2_CODE").Value = dynARTOPEN1.Fields("SEG2_CODE").Value
        dynARTCASH3.Fields("SEG3_CODE").Value = dynARTOPEN1.Fields("SEG3_CODE").Value
        dynARTCASH3.Fields("SEG4_CODE").Value = dynARTOPEN1.Fields("SEG4_CODE").Value
            dynARTCASH3.Fields("INV_BALANCE").Value = dynARTOPEN1.Fields("INV_BALANCE").Value
            dynARTCASH3.Fields("INV_BALANCE_CURR").Value = dynARTOPEN1.Fields("INV_BALANCE_CURR").Value
            dynARTCASH3.Fields("INV_PMT_CURR").Value = INV_BALANCE_expected 'dynARTOPEN1.Fields("INV_BALANCE").Value
            dynARTCASH3.Fields("INV_DISC_TAKEN_CURR").Value = 0
            dynARTCASH3.Fields("INV_WRITE_OFF_CURR").Value = 0
            dynARTCASH3.Fields("INV_BALANCE_NEW_CURR").Value = Val(dynARTOPEN1.Fields("INV_BALANCE").Value & "") - INV_BALANCE_expected
            dynARTCASH3.Fields("CURR_CODE").Value = "USD"
            dynARTCASH3.Fields("CURR_EXCH_RATE").Value = 1

        dynARTCASH3.Update
    
        dynARTOPEN1.Edit
        dynARTOPEN1.Fields("INV_LAST_PMT").Value = INV_DATE
        dynARTOPEN1.Fields("INV_PMT").Value = Val(dynARTOPEN1.Fields("INV_PMT").Value & "") + INV_BALANCE_expected
        dynARTOPEN1.Fields("INV_PMT_CURR").Value = (Val(dynARTOPEN1.Fields("INV_PMT_CURR").Value & "") + INV_BALANCE_expected_CURR)
        dynARTOPEN1.Fields("INV_BALANCE").Value = Val(dynARTOPEN1.Fields("INV_BALANCE").Value & "") - INV_BALANCE_expected
        dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = Val(dynARTOPEN1.Fields("INV_BALANCE").Value & "") - INV_BALANCE_expected
        'dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = (Val(dynARTOPEN1.Fields("INV_BALANCE_CURR").Value & "") - INV_BALANCE_expected_CURR)
        dynARTOPEN1.Fields("INV_LAST_PMT_REF").Value = "AUTO"
        dynARTOPEN1.Fields("INV_LAST_PMT_REF_DT").Value = INV_DATE
        dynARTOPEN1.Fields("LAST_OPER").Value = xUserID
        dynARTOPEN1.Fields("LAST_DATE").Value = LAST_DATE
        dynARTOPEN1.Update
    End If
Return

Consolidated_Invoice:
    OraD.Parameters("CUST_CODE").Value = dynARM.Fields("CUST_CODE").Value
    dynARTCUST1.Refresh
    TERM_CODE = dynARTCUST1.Fields("TERM_CODE").Value & ""
    TERM_DAYS_DUE = Val(SRead(TATTERM1, TERM_CODE, 2))
    TERM_DUE_TYPE = SRead(TATTERM1, TERM_CODE, 3)
    TERM_ADDL_MOS = Val(SRead(TATTERM1, TERM_CODE, 4))
    'INV_DUE_DATE = Format$(DateValue(INV_DATE) + TERM_DAYS_DUE, "MM/DD/YYYY")
    INV_DATE = Format$(dynARM.Fields("INV_DATE").Value & "", "mm/dd/yyyy")
    INV_DUE_DATE = Calc_Due_Date(INV_DATE, TERM_DAYS_DUE, TERM_DUE_TYPE, TERM_ADDL_MOS)
    
    dynARTOPEN1.AddNew
    dynARTOPEN1.Fields("CUST_CODE").Value = dynARM.Fields("CUST_CODE").Value
    dynARTOPEN1.Fields("INV_TYPE").Value = "I"
    'INV_NO -> INV_NUM change made to accomodate conversion to AR V2. - 8/4/16 W.R.
    'dynARTOPEN1.Fields("INV_NO").Value = dynARM.Fields("INV_NO_CONS").Value
    dynARTOPEN1.Fields("INV_NUM").Value = dynARM.Fields("INV_NO_CONS").Value
    dynARTOPEN1.Fields("INV_DATE").Value = dynARM.Fields("INV_DATE").Value
    dynARTOPEN1.Fields("CUST_STORE_NO").Value = ""
    dynARTOPEN1.Fields("POST_CODE").Value = dynARTCUST1.Fields("POST_CODE").Value & ""
    dynARTOPEN1.Fields("TERM_CODE").Value = TERM_CODE
    dynARTOPEN1.Fields("INV_DUE_DATE").Value = INV_DUE_DATE
    dynARTOPEN1.Fields("SREP_CODE").Value = ""
    If dynARM.Fields("EDI_APPOINTMENT").Value & "" <> "" Then
        dynARTOPEN1.Fields("INV_CUST_PO").Value = dynARM.Fields("EDI_APPOINTMENT").Value
    Else
        dynARTOPEN1.Fields("INV_CUST_PO").Value = dynARM.Fields("ORDR_CUST_PO").Value
    End If
    'INV_ORDR_NO -> ORDR_NO change made to accomodate conversion to AR V2. - 8/4/16 W.R.
    'dynARTOPEN1.Fields("INV_ORDR_NO").Value = ""
    dynARTOPEN1.Fields("ORDR_NO").Value = ""
    dynARTOPEN1.Fields("INV_SALES").Value = dynARM.Fields("INV_SALES").Value
    dynARTOPEN1.Fields("INV_DISC").Value = 0
    dynARTOPEN1.Fields("INV_FREIGHT").Value = dynARM.Fields("INV_FREIGHT").Value
    dynARTOPEN1.Fields("INV_STAX").Value = 0
    dynARTOPEN1.Fields("INV_MISC_CHG").Value = dynARM.Fields("INV_MISC_CHG").Value
    dynARTOPEN1.Fields("INV_TOTAL_AMOUNT").Value = dynARM.Fields("INV_TOTAL_AMOUNT").Value
    dynARTOPEN1.Fields("INV_BALANCE").Value = dynARM.Fields("INV_TOTAL_AMOUNT").Value
    dynARTOPEN1.Fields("CUST_CODE_SO").Value = dynARM.Fields("CUST_CODE").Value
    dynARTOPEN1.Fields("REASON_CODE").Value = dynARM.Fields("REASON_CODE").Value
    dynARTOPEN1.Fields("SALES_DIVISION_CODE").Value = dynARM.Fields("SALES_DIVISION_CODE").Value
    dynARTOPEN1.Fields("SEG2_CODE").Value = "000"
    dynARTOPEN1.Fields("SEG3_CODE").Value = "000"
    dynARTOPEN1.Fields("SEG4_CODE").Value = "000"
    dynARTOPEN1.Fields("INIT_OPER").Value = xUserID
    dynARTOPEN1.Fields("INIT_DATE").Value = LAST_DATE
    'CURRENCY
    If Len(CURR_CODE & "") = 0 Or CURR_EXCH_RATE = 0 Then
        Stop
    End If
    dynARTOPEN1.Fields("CURR_CODE").Value = CURR_CODE
    dynARTOPEN1.Fields("CURR_EXCH_RATE").Value = CURR_EXCH_RATE
    dynARTOPEN1.Fields("INV_SALES_CURR").Value = dynARM.Fields("INV_SALES").Value / CURR_EXCH_RATE
    dynARTOPEN1.Fields("INV_DISC_CURR").Value = 0
    dynARTOPEN1.Fields("INV_FREIGHT_CURR").Value = dynARM.Fields("INV_FREIGHT").Value / CURR_EXCH_RATE
    dynARTOPEN1.Fields("INV_STAX_CURR").Value = 0
    dynARTOPEN1.Fields("INV_MISC_CHG_CURR").Value = dynARM.Fields("INV_MISC_CHG").Value / CURR_EXCH_RATE
    dynARTOPEN1.Fields("GST_TAX").Value = (dynARM.Fields("INV_SALES").Value * GST_TAX)
    dynARTOPEN1.Fields("GST_TAX_CURR").Value = (dynARM.Fields("INV_SALES").Value * GST_TAX) / CURR_EXCH_RATE
    dynARTOPEN1.Fields("INV_TOTAL_AMOUNT_CURR").Value = dynARM.Fields("INV_TOTAL_AMOUNT").Value / CURR_EXCH_RATE
    'dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = INV_BALANCE / CURR_EXCH_RATE 'Changed with WZ on 11/1/16.  W.R.
    dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = dynARM.Fields("INV_TOTAL_AMOUNT").Value / CURR_EXCH_RATE
    
    'Added on 2/6/17 per WZ - WR.
    If Format(INV_DATE, "YYYYMM") >= Format(Right(NYM, 2) & "/01/" & Left$(NYM, 4), "YYYYMM") Then
        dynARTOPEN1.Fields("OPS_YYYYPP").Value = NYP
    Else
        dynARTOPEN1.Fields("OPS_YYYYPP").Value = CYP
    End If
    
    dynARTOPEN1.Update

Return

End Sub


Sub Update()
    Call Update_Invoices
    Call Done
End Sub

Sub Print_Report()
    Dim z As String
    Dim i As Integer
    Dim j As Integer
    
    'Reports were removed from here to avoid the never ending confusion
    'Over which reports get used.  All reports related to shipping are
    'to be run from the Monthly Sales Reports Menu.
    
'    xRPT = "SORUPDT1"
'    xRptTitle = ""
'    Call Std_Report_Parameters
'    z = ""
'    CR_Rpt.ParameterFields("SUBT").SetCurrentValue SUBTITLE
'    Call Prepare_SPRF
'
'    xRPT = "SORMTDS1"
'    xRptTitle = ""
'    Call Std_Report_Parameters
'    z = ""
'    CR_Rpt.ParameterFields("SUBT").SetCurrentValue SUBTITLE
'    CR_Rpt.ParameterFields("CHKCOSTS").SetCurrentValue "0"
'    CR_Rpt.ParameterFields("ONLY_THOSE_CUSTS").SetCurrentValue "1"
'    Call Prepare_SPRF
'
'    xRPT = "SORINVHR"
'    xRptTitle = ""
'    Call Std_Report_Parameters
'    z = ""
'    CR_Rpt.ParameterFields("SUBT").SetCurrentValue SUBTITLE
'    Call Prepare_SPRF
'
'    xRPT = "SORFINR1"
'    xRptTitle = ""
'    Call Std_Report_Parameters
'    z = ""
'    CR_Rpt.ParameterFields("SUBT").SetCurrentValue SUBTITLE
'    Call Prepare_SPRF
    
End Sub

Public Sub Go(cliASCCALLB As Object, ProjectID As String, FormName As String, XNO As String, UID As String, PWD As String, dbName As String)
    xProjectID = ProjectID
    xFormName = FormName
    xXNO = XNO
    xUID = UID
    xPWD = PWD
    xdbname = dbName
    
    Set objASCCALLB = cliASCCALLB
    Set f = Me
    If ABSWEB Then Call Main_Process Else ASFMAINX.Timer1.Enabled = True
End Sub

Public Sub Terminate()
    Call Terminate_Std
    End
End Sub

Private Sub Update_Transfer(INV_NO As String)

    sql = "BEGIN SOPSHIP1_XFRVAN('" & INV_NO & "'); END;"
    OraD.ExecuteSQL sql
    
    Dim dynICTIXFR1 As OraDynaset
    Dim XFR_NO As String
    sql = "Select XFR_NO from ICTIXFR1 where XFR_SOURCE = 'S' and CTL_NO = '" & INV_NO & "'"
    Set dynICTIXFR1 = OraD.CreateDynaset(sql, 8&)
    Do While Not dynICTIXFR1.EOF
        XFR_NO = dynICTIXFR1.Fields("XFR_NO").Value
        dynICTIXFR1.MoveNext
    Loop
    dynICTIXFR1.Close
    Set dynICTIXFR1 = Nothing
    
    'Dim sqlw As String
    'sqlw = " where PO_SHIPMENT_NO = '" & XFR_NO & "' AND SOURCE_TYPE = 'T'"
    'OraD.ExecuteSQL ("INSERT INTO ADS.WHTPORD1@ADSIIS SELECT * FROM WHTPORD1 " & sqlw)
    'OraD.ExecuteSQL ("INSERT INTO ADS.WHTPORD2@ADSIIS SELECT * FROM WHTPORD2 " & sqlw)
End Sub

