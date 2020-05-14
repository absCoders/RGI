Public Class SPCMAIN1
  
    Public Shared Function Get_APPLIES_TO( _
    ByVal rowSPTCONT1 As DataRow, _
    ByVal tblSPTCONTA As DataTable, _
    ByVal tblSPTCONTB As DataTable, _
    ByVal CONT_NO2 As String, _
    ByVal CONT_APPL_TYPE As String, _
    ByVal CONT_APPL_LNO As Integer, _
    Optional ByVal CODE_TYPEx As String = "", _
    Optional ByVal CODE_INCL_EXCLx As String = "", _
    Optional ByVal CODE_VALUEx As String = "") As String

        Dim CODE_TYPE As String
        Dim CODE_INCL_EXCL As String
        Dim CODE_VALUEs As String
        Dim APPLIES_TO As String = ""

        Dim PROD_DESCx As String
        Dim COLUMN_SEQ As Integer

        Dim SQLW As String = "" _
        & "CONT_NO = '" & CONT_NO2 & "'" & vbCrLf _
        & " and CONT_APPL_TYPE = '" & CONT_APPL_TYPE & "'" & vbCrLf _
        & " and CONT_APPL_LNO = " & CStr(CONT_APPL_LNO)
        If CODE_TYPEx <> "" Then
            SQLW &= " and (CODE_TYPE = '" & CODE_TYPEx & "' or CODE_VALUES is not null)"
        Else
            SQLW &= " and CODE_VALUES is not null"
        End If
        'If rowSPTCONT1.Item("MARKET_TYPE") = "C" Then ' THIS SHOULD PROBABLY BE HANDLED DIFFERENTLY, LIKE SETTING UP SPWCONTB FOR THE MARKET_TYPE BEFORE STARTING WITH THE CONTRACT
        '    SQLW &= " and CODE_TYPE = 'GTIN'"
        'Else
        '    SQLW &= " and CODE_TYPE <> 'GTIN'"
        'End If

        For Each rowSPTCONTB As DataRow In tblSPTCONTB.Select(SQLW, "COLUMN_SEQ")
            PROD_DESCx = ""
            CODE_TYPE = rowSPTCONTB.Item("CODE_TYPE") & ""
            COLUMN_SEQ = Val(rowSPTCONTB.Item("COLUMN_SEQ") & "")
            If CODE_TYPE = CODE_TYPEx Then
                CODE_VALUEs = CODE_VALUEx
                CODE_INCL_EXCL = CODE_INCL_EXCLx
            Else
                CODE_VALUEs = rowSPTCONTB.Item("CODE_VALUES") & ""
                CODE_INCL_EXCL = rowSPTCONTB.Item("CODE_INCL_EXCL") & ""
            End If
            If CODE_TYPE = "BRAND_CODE" Or CODE_TYPE = "ORIG_CODE" Or CODE_TYPE = "SP_GROUP" Or CODE_TYPE = "MOP" Then
                PROD_DESCx = PROD_DESCx & "," & CODE_VALUEs
            Else
                Dim rowSPTCONTA As DataRow = tblSPTCONTA.Rows.Find(CODE_TYPE)
                ASCMAIN1.sql = "Select " & rowSPTCONTA.Item("COLUMN_NAME_CODE") & "," _
                & rowSPTCONTA.Item("COLUMN_NAME_DESC") _
                & " from " & rowSPTCONTA.Item("TABLE_NAME") _
                & " where " & rowSPTCONTA.Item("COLUMN_NAME_CODE") _
                & " IN ('" & Replace$(CODE_VALUEs, ",", "','") & "')"
                For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                    PROD_DESCx = PROD_DESCx & "," & row.Item(1)
                Next
            End If

            If PROD_DESCx <> "" Then
                PROD_DESCx = Mid$(PROD_DESCx, 2)
                If CODE_INCL_EXCL = "1" Then
                    PROD_DESCx = "Except: " & PROD_DESCx
                End If
                APPLIES_TO = APPLIES_TO & "; " & PROD_DESCx
            End If
        Next

        If APPLIES_TO = "" Then
            APPLIES_TO = "All"
        Else
            APPLIES_TO = Mid$(APPLIES_TO, 3)
            If Len(APPLIES_TO) > 255 Then
                APPLIES_TO = Mid$(APPLIES_TO, 1, 254) & "+"
            End If
        End If

        Return APPLIES_TO
    End Function

    Public Shared Function Setup_SPTCONTA() As DataTable
        Dim tblSPTCONTA As New DataTable("SPTCONTA")
        With tblSPTCONTA
            .Columns.Add("COLUMN_NAME_CODE")
            .Columns.Add("COLUMN_CAPTION")
            .Columns.Add("COLUMN_SEQ", GetType(System.Int16))
            .Columns.Add("COLUMN_NAME_DESC")
            .Columns.Add("TABLE_NAME")
            .PrimaryKey = New DataColumn() {.Columns("COLUMN_NAME_CODE")}
        End With
        Setup_SPTCONTA(tblSPTCONTA)
        Return tblSPTCONTA
    End Function

    Public Shared Sub Setup_SPTCONTA(ByVal tblSPTCONTA As DataTable) ', ByVal MARKET_TYPE As String)
        With tblSPTCONTA
            .Rows.Add(New Object() {"PROD_CODE", "Product", 1, "PROD_DESC", "ICTPROD1"})
            .Rows.Add(New Object() {"CLASS_CODE", "Class", 2, "CLASS_DESC", "ICTCLAS1"})
            .Rows.Add(New Object() {"CATEGORY_CODE", "Category", 3, "CATEGORY_DESC", "ICTCATG1"})
            .Rows.Add(New Object() {"PROD_GROUP_CODE", "Group", 4, "PROD_GROUP_DESC", "ICTPROD2"})
            .Rows.Add(New Object() {"BRAND_CODE", "Brand", 5, "BRAND_DESC", "ICTBRAN1"})
            .Rows.Add(New Object() {"SIZE_CODE", "Size", 6, "SIZE_DESC", "ICTSIZE1"})
            .Rows.Add(New Object() {"ORIG_CODE", "Origin", 7, "ORIG_DESC", "ICTORIG1"})
            .Rows.Add(New Object() {"SP_GROUP", "Sp/Group", 8, "SP_GROUP_DESC", "ICTSPGR1"})
            .Rows.Add(New Object() {"PACK_CODE", "Pack", 9, "PACK_DESC", "ICTPACK1"})
            .Rows.Add(New Object() {"GTIN", "GTIN", 10, "GTIN_DESC", "ICTUPCDX"})
            .Rows.Add(New Object() {"ITEM_CODE", "Item", 11, "ITEM_DESC", "ICTITEM0"})
            .Rows.Add(New Object() {"CUST_SKU", "Cust SKU", 12, "ITEM_DESC", "SOTCITM1"})
        End With

        '    If MARKET_TYPE = "C" Then
        '        ReDim BC(1, 4)
        '        BC(1, 1) = "GTIN" : BC(1, 0) = "GTIN" : BC(1, 2) = "GTIN_DESC" : BC(1, 3) = "ICTUPCDX"
        '    Else
        '        ReDim BC(10, 4)
        '        BC(1, 1) = "Product" : BC(1, 0) = "PROD_CODE" : BC(1, 2) = "PROD_DESC" : BC(1, 3) = "ICTPROD1"
        '        BC(2, 1) = "Class" : BC(2, 0) = "CLASS_CODE" : BC(2, 2) = "CLASS_DESC" : BC(2, 3) = "ICTCLAS1"
        '        BC(3, 1) = "Category" : BC(3, 0) = "CATEGORY_CODE" : BC(3, 2) = "CATEGORY_DESC" : BC(3, 3) = "ICTCATG1"
        '        BC(4, 1) = "Group" : BC(4, 0) = "PROD_GROUP_CODE" : BC(4, 2) = "PROD_GROUP_DESC" : BC(4, 3) = "ICTPROD2"
        '        BC(5, 1) = "Brand" : BC(5, 0) = "BRAND_CODE" : BC(5, 2) = "BRAND_DESC" : BC(5, 3) = "ICTBRAN1"
        '        BC(6, 1) = "Size" : BC(6, 0) = "SIZE_CODE" : BC(6, 2) = "SIZE_DESC" : BC(6, 3) = "ICTSIZE1"
        '        BC(7, 1) = "Origin" : BC(7, 0) = "ORIG_CODE" : BC(7, 2) = "ORIG_DESC" : BC(7, 3) = "ICTORIG1"
        '        BC(8, 1) = "SP/Group" : BC(8, 0) = "SP_GROUP" : BC(8, 2) = "SP_GROUP_DESC" : BC(8, 3) = "ICTSPGRP"
        '        BC(9, 1) = "MOP" : BC(9, 0) = "MOP" : BC(9, 2) = "MOP_DESC" : BC(9, 3) = "ASTCODE1"
        '        BC(10, 1) = "Pack" : BC(10, 0) = "PACK_CODE" : BC(10, 2) = "PACK_DESC" : BC(10, 3) = "ICTPACK1"
        '    End If

        '    colBC = New Collection
        '    For i As Integer = 1 To UBound(BC, 1)
        '        colBC.Add(CStr(i), BC(i, 0))
        '    Next i
    End Sub

    Public Shared Sub Get_BC(ByVal CONT_NO As String, ByVal CONT_NO2 As String, ByVal clsASCBASE1 As ASCBASE1)

        Dim sqlx As String = ""
        If CONT_NO2 <> CONT_NO Then
            sqlx = "'" & CONT_NO & "' CONT_NO_BASE, 1 CONT_NO_SEQ, "
        End If

        ASCMAIN1.sql = "Select " & sqlx & " SPTCONTB.*, NULL AS CODE_VALUES, 0 CODE_TYPE_SEQ from SPTCONTB " _
        & " where CONT_NO = '" & CONT_NO2 & "'"
        clsASCBASE1.Fill_Records("SPTCONTB", "", False, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select " & sqlx & " SPTCONTC.* from SPTCONTC where CONT_NO = '" & CONT_NO2 & "'"
        clsASCBASE1.Fill_Records("SPTCONTC", "", False, ASCMAIN1.sql)

        For Each rowSPTCONTC As DataRow In clsASCBASE1.dst.Tables("SPTCONTC").Select _
            ("CONT_NO = '" & CONT_NO2 & "'", "CONT_NO, CONT_APPL_TYPE, CONT_APPL_LNO, CODE_TYPE, CODE_VALUE")
            Dim rowSPTCONTB As DataRow
            If CONT_NO = CONT_NO2 Or 1 = 1 Then ' HAD TO PUT 1=1 BECAUSE i COULD NOT SEE WHERE 6 PART KEY DATA WAS SET UP
                rowSPTCONTB = clsASCBASE1.dst.Tables("SPTCONTB").Rows.Find _
                    (New Object() {rowSPTCONTC.Item("CONT_NO"), _
                                   rowSPTCONTC.Item("CONT_APPL_TYPE"), _
                                   rowSPTCONTC.Item("CONT_APPL_LNO"), _
                                   rowSPTCONTC.Item("CODE_TYPE")})
            Else
                rowSPTCONTB = clsASCBASE1.dst.Tables("SPTCONTB").Rows.Find _
                    (New Object() {CONT_NO, 1, _
                                   rowSPTCONTC.Item("CONT_NO"), _
                                   rowSPTCONTC.Item("CONT_APPL_TYPE"), _
                                   rowSPTCONTC.Item("CONT_APPL_LNO"), _
                                   rowSPTCONTC.Item("CODE_TYPE")})
            End If
            Dim CODE_VALUE As String = rowSPTCONTC.Item("CODE_VALUE") & ""
            If rowSPTCONTB.Item("CODE_VALUES") & "" = "" Then
                rowSPTCONTB.Item("CODE_VALUES") = CODE_VALUE
            Else
                rowSPTCONTB.Item("CODE_VALUES") = rowSPTCONTB.Item("CODE_VALUES") & "," & CODE_VALUE
            End If
        Next
    End Sub

    Public Shared Function Get_Fund_SQL() As String

        Return "Select SPTFUND2.*" & vbCrLf _
        & ", 0.01 FUND_AMT_EXP_NOW" & vbCrLf _
        & ", TO_CHAR(SPTCONT1.CONT_DATE_START,'YYYYMM') YP_FUND_START" & vbCrLf _
        & ", TO_CHAR(SPTCONT1.CONT_DATE_END,'YYYYMM') YP_FUND_END" & vbCrLf _
        & ", 'YYYYPP' YP_FUND_SHIP" & vbCrLf _
        & ", 'YYYYPP' YP_FUND_PYMT" & vbCrLf _
        & ", 0 FUND_MOS_SPAN" & vbCrLf _
        & ", 0 FUND_MOS_DONE" & vbCrLf _
        & ", 0 FUND_EXP_INCR_DONE" & vbCrLf _
        & ", 0.01 FUND_AMT_EXP_NOW_PCT" & vbCrLf _
        & ", 0.01 FUND_TOTAL_EXP_BASIS" & vbCrLf _
        & ", 0.01 FUND_AMT_OPEN" & vbCrLf _
        & ", SPTFUND1.FUND_ACCR_ON_SHIPMENT" & vbCrLf _
        & " from SPTFUND2,SPTCONT1,SPTFUND1 " & vbCrLf _
        & " where SPTCONT1.CONT_NO = SPTFUND2.CONT_NO" & vbCrLf _
        & "   and SPTFUND1.FUND_CODE = SPTFUND2.FUND_CODE" & vbCrLf

    End Function

    Public Shared Function Get_Fund_Accruals( _
    Optional ByVal RYP As String = "", _
    Optional ByVal sqlw As String = "") As String

        Dim TT As String = ""

        ASCMAIN1.sql = Get_Fund_SQL() & sqlw
        TT = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & TT & " Add Primary Key (CONT_NO, CONT_FNO)")
        ASCMAIN1.AnalyzeTable(TT)

        If RYP <> "" Then
            Set_Fund_Data(TT, RYP)
        End If

        Get_Fund_Accruals = TT
    End Function

    Public Shared Sub Set_Fund_Data(ByVal TT As String, ByVal RYP As String)

        ASCMAIN1.sql = "Update " & TT & " set YP_FUND_START = Null" & vbCr _
        & " where FUND_ACCR_ON_SHIPMENT = '1'" & vbCr _
        & " and NVL(FUND_ACCRUE_WOUT_SHIP,'0') <> '1'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & TT & " TT set YP_FUND_SHIP = " & vbCr _
        & " (SELECT OPS_YYYYPP from SOTINVH1 " & vbCr _
        & "  where ORDR_INV_NO = TT.ORDR_INV_NO) " & vbCr _
        & " where FUND_ACCR_ON_SHIPMENT = '1'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & TT & " TT set YP_FUND_PYMT = " & vbCr _
        & " (SELECT MIN(OPS_YYYYPP) from SPTFUND3 " & vbCr _
        & "  where CONT_NO = TT.CONT_NO and CONT_FNO = TT.CONT_FNO) " & vbCr _
        & " where FUND_ACCR_ON_SHIPMENT = '1'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & TT & " set YP_FUND_START = LEAST (NVL(YP_FUND_SHIP,YP_FUND_PYMT),NVL(YP_FUND_PYMT,YP_FUND_SHIP))" & vbCr _
        & " where FUND_ACCR_ON_SHIPMENT = '1' " & vbCr _
        & " and NVL(FUND_ACCRUE_WOUT_SHIP,'0') <> '1'" & vbCr _
        & " and (YP_FUND_SHIP is Not Null or YP_FUND_PYMT is Not Null)"
        ASCDATA1.ExecuteSQL()

        Dim delsqlw As String
        Dim delsqla As String
        If ASCMAIN1.FORM_NAME = "SPFFNDA1" Then
            ASCMAIN1.sql = "Delete from " & TT _
            & " where YP_FUND_START is Null or YP_FUND_START > '" & RYP & "'"
            ASCDATA1.ExecuteSQL()
            delsqlw = ""
            delsqla = ""
        Else
            delsqlw = " where NOT (YP_FUND_START is Null or YP_FUND_START > '" & RYP & "')"
            delsqla = "   and NOT (YP_FUND_START is Null or YP_FUND_START > '" & RYP & "')"
        End If
        ASCMAIN1.sql = "Update " & TT & " set FUND_AMT_EXP_NOW = 0, FUND_AMT_EXP_NOW_PCT = 0, FUND_TOTAL_EXP_BASIS = 0, FUND_AMT_OPEN = 0"
        ASCDATA1.ExecuteSQL()

        Dim sqlX As String

        sqlX = " set YP_FUND_END = TO_CHAR(ADD_MONTHS(TO_DATE(YP_FUND_START || '01','YYYYMMDD'),11),'YYYYMM')" _
        & " where YP_FUND_END is Null "
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)
        sqlX = " set YP_FUND_END = TO_CHAR(ADD_MONTHS(TO_DATE(YP_FUND_START || '01','YYYYMMDD'),11),'YYYYMM')" _
        & " where (YP_FUND_END > TO_CHAR(ADD_MONTHS(TO_DATE(YP_FUND_START || '01','YYYYMMDD'),11),'YYYYMM') or  YP_FUND_END is not null)"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)

        sqlX = " set YP_FUND_END = YP_FUND_START where YP_FUND_END < YP_FUND_START"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)

        sqlX = " set FUND_MOS_SPAN = 1 + (TO_NUMBER(SUBSTR(YP_FUND_END,1,4))*12 + TO_NUMBER(SUBSTR(YP_FUND_END,5,2))) - (TO_NUMBER(SUBSTR(YP_FUND_START,1,4))*12 + TO_NUMBER(SUBSTR(YP_FUND_START,5,2)))"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqlw)
        sqlX = " set FUND_MOS_DONE = 1 + " & CStr(Val(Mid$(RYP, 1, 4)) * 12 + Val(Mid$(RYP, 5, 2))) & " - (TO_NUMBER(SUBSTR(YP_FUND_START,1,4))*12 + TO_NUMBER(SUBSTR(YP_FUND_START,5,2)))"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqlw)

        sqlX = " set FUND_MOS_DONE = FUND_MOS_SPAN WHERE FUND_MOS_DONE > FUND_MOS_SPAN"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)

        sqlX = " set FUND_EXP_INCR = 1 WHERE NVL(FUND_EXP_INCR,0) < 1"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)

        sqlX = " set FUND_TOTAL_EXP_BASIS = NVL(FUND_AMT,0)"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqlw)
        sqlX = " set FUND_TOTAL_EXP_BASIS = NVL(FUND_AMT_RCOV,0) where NVL(FUND_AMT,0) = 0 and FUND_EXP_METHOD = 'R'"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)
        sqlX = " set FUND_TOTAL_EXP_BASIS = NVL(FUND_AMT_PAID,0) where NVL(FUND_AMT_PAID,0) > NVL(FUND_TOTAL_EXP_BASIS,0)"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)

        sqlX = " set FUND_EXP_INCR_DONE = CEIL(NVL(FUND_EXP_INCR,0) * (NVL(FUND_MOS_DONE,0) / NVL(FUND_MOS_SPAN,0)))" _
        & " Where NVL(FUND_MOS_SPAN,0) <> 0 "
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqla)

        sqlX = " set FUND_AMT_EXP_NOW_PCT = NVL(FUND_EXP_INCR_DONE,0) / NVL(FUND_EXP_INCR,0)"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqlw)
        sqlX = " set FUND_AMT_EXP_NOW = NVL(FUND_TOTAL_EXP_BASIS,0) * NVL(FUND_AMT_EXP_NOW_PCT,0) - NVL(FUND_AMT_EXP,0)"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX & delsqlw)

        sqlX = " set FUND_AMT_OPEN = DECODE(NVL(FUND_AMT,0),0,NVL(FUND_AMT_RCOV,0) - NVL(FUND_AMT_PAID,0),NVL(FUND_AMT,0) - NVL(FUND_AMT_PAID,0))"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX)
        sqlX = " set FUND_AMT_OPEN = 0 WHERE NVL(FUND_AMT_OPEN,0) < 0 OR FUND_STATUS = 'C'"
        ASCDATA1.ExecuteSQL("Update " & TT & sqlX)

        If ASCMAIN1.FORM_NAME = "SPFFNDA1" Or ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT = "TARPEND0" Then
            ASCMAIN1.sql = "Delete from " & TT & " where NVL(FUND_AMT_EXP_NOW,0) = 0"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    'Sub BRR_Prepare(ByVal RYP0 As String, ByVal RYP1 As String, ByVal chkHISTORY As String, ByVal chkMYITEMSONLY As String)

    '    ' Build Work Tables

    '    Call Track("Initialize Work Tables", "")

    '    Call BR_sql_SOTINVHC()

    '    If chkHISTORY = "1" Then
    '        sql = sql & " and SOTINVHC.CONT_PYMT_STATUS = 'P'" & vbCr
    '        sql = sql & " and SOTINVHC.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCr
    '    Else
    '        sql = sql & " and SOTINVHC.CONT_PYMT_STATUS = 'R'" & vbCr
    '        sql = sql & " and SOTINVHC.REGISTER_IND is Null" & vbCr
    '    End If

    '    If chkMYITEMSONLY = "1" Then
    '        sql = sql & " and SOTINVHC.INIT_OPER = '" & xUserID & "'" & vbCr
    '    End If

    '    If Sel(1, 1) <> "" Then
    '        sql = sql & " and SOTINVH1.CUST_BILL_TO_CUST " & Sel(2, 1) & " in (" & Sel(1, 1) & ")" & vbCr
    '    End If
    '    If Sel(1, 2) <> "" Then
    '        sql = sql & " and SOTINVH1.CUST_CODE " & Sel(2, 2) & " in (" & Sel(1, 2) & ")" & vbCr
    '    End If
    '    If Sel(1, 3) <> "" Then
    '        sql = sql & " and SOTINVHC.CONT_PYMT_VIA = 'P'" & vbCr
    '        sql = sql & " and SOTINVHC.CONT_PYMT_PAYEE_CODE " & Sel(2, 3) & " in (" & Sel(1, 3) & ")" & vbCr
    '    End If
    '    If Sel(1, 4) <> "" Then
    '        sql = sql & " and SOTINVH1.SREP_CODE " & Sel(2, 4) & " in (" & Sel(1, 4) & ")" & vbCr
    '    End If
    '    If Sel(1, 5) <> "" Then
    '        sql = sql & " and SOTINVH1.BRKR_CODE " & Sel(2, 5) & " in (" & Sel(1, 5) & ")" & vbCr
    '    End If

    '    Dim TTC As String
    '    Call BR_Make_TTC(TTC)

    '    Call BR_Make_TTB(TTC)

    '    ' Set OK to Pay

    '    If chkHISTORY = "1" Then
    '        sql = "Update SOWINVHC SET CONT_PYMT_OK_TO_PAY = '1' "
    '        AccD.Execute(sql)
    '    Else
    '        Call SPBMAINX.Set_OK_TO_PAY(TTC)
    '    End If

    '    ' Create Payments

    '    If chkHISTORY <> "1" Then
    '        sql = "Select Count (*) from SOWINVHC where CONT_PYMT_OK_TO_PAY = '1'"
    '        Dim dynwk As Recordset
    '        dynwk = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    '        If Val(dynwk.Fields(0).Value & "") = 0 Then
    '            xRWU = "N"
    '            'objASCCALLB.xErrMsg = "Nothing to Update"
    '        End If
    '    End If

    'End Sub

    'Sub BRR_Update()

    '    Dim dynSOWINVHC As Recordset
    '    sql = "Select * from SOWINVHC where CONT_PYMT_OK_TO_PAY = '1'"
    '    dynSOWINVHC = AccD.OpenRecordset(sql, dbOpenForwardOnly)

    '    Dim dynSOTINVHC As OraDynaset
    '    sql = "Select * from SOTINVHC where SO_ORDER_NO = :CODE"
    '    sql = sql & " and CONT_TP_TYPE = :CODE1"
    '    sql = sql & " and CONT_PYMT_VIA = :CODE2"
    '    sql = sql & " and CONT_PYMT_PAYEE_CODE = :CODE3"
    '    sql = sql & " and CONT_PYMT_SEQ = :NUM1"
    '    dynSOTINVHC = OraD.CreateDynaset(sql, 0&)

    '    Do While Not dynSOWINVHC.EOF
    '        With dynSOTINVHC
    '            OraD.Parameters("CODE").Value = dynSOWINVHC.Fields("SO_ORDER_NO").Value
    '            OraD.Parameters("CODE1").Value = dynSOWINVHC.Fields("CONT_TP_TYPE").Value
    '            OraD.Parameters("CODE2").Value = dynSOWINVHC.Fields("CONT_PYMT_VIA").Value
    '            OraD.Parameters("CODE3").Value = dynSOWINVHC.Fields("CONT_PYMT_PAYEE_CODE").Value
    '            OraD.Parameters("NUM1").Value = dynSOWINVHC.Fields("CONT_PYMT_SEQ").Value
    '            .Refresh()
    '            .Edit()
    '            .Fields("REGISTER_IND").Value = "1"
    '            .Fields("REGISTER_XNO").Value = xXNO
    '            .Update()
    '        End With
    '        dynSOWINVHC.MoveNext()
    '    Loop
    '    dynSOWINVHC.Close()

    'End Sub

    'Public Shared Function Get_Fund_sql() As String

    '    Dim sql As String = "Select SPTFUND2.*" & vbCrLf _
    '    & ", 0.01 FUND_AMT_EXP_NOW" & vbCrLf _
    '    & ", TO_CHAR(SPTCONT1.CONT_DATE_START,'YYYYMM') YP_FUND_START" & vbCrLf _
    '    & ", TO_CHAR(SPTCONT1.CONT_DATE_END,'YYYYMM') YP_FUND_END" & vbCrLf _
    '    & ", 'YYYYPP' YP_FUND_SHIP" & vbCrLf _
    '    & ", 'YYYYPP' YP_FUND_PYMT" & vbCrLf _
    '    & ", 0 FUND_MOS_SPAN" & vbCrLf _
    '    & ", 0 FUND_MOS_DONE" & vbCrLf _
    '    & ", 0 FUND_EXP_INCR_DONE" & vbCrLf _
    '    & ", 0.01 FUND_AMT_EXP_NOW_PCT" & vbCrLf _
    '    & ", 0.01 FUND_TOTAL_EXP_BASIS" & vbCrLf _
    '    & ", 0.01 FUND_AMT_OPEN" & vbCrLf _
    '    & ", SPTFUND1.FUND_ACCR_ON_SHIPMENT" & vbCrLf _
    '    & " from SPTFUND2,SPTCONT1,SPTFUND1 " & vbCrLf _
    '    & " where SPTCONT1.CONT_NO = SPTFUND2.CONT_NO" & vbCrLf _
    '    & "   and SPTFUND1.FUND_CODE = SPTFUND2.FUND_CODE"

    '    Return sql
    'End Function


    'Sub Set_OK_TO_PAY(ByVal TTC As String)

    '    ' Calculate Paid in Cash

    '    sql = "SELECT INV_TYPE, INV_NUM, INV_PMT CUST_CHECK_AMT,"
    '    sql = sql & " INV_PMT, INV_DISC_TAKEN, INV_WRITE_OFF,"
    '    sql = sql & " INV_PMT INV_PMT_I, INV_PMT INV_DISC_TAKEN_I, INV_PMT INV_WRITE_OFF_I,"
    '    sql = sql & " INV_PMT INV_PMT_NI, INV_PMT INV_DISC_TAKEN_NI, INV_PMT INV_WRITE_OFF_NI,"
    '    sql = sql & " INV_PMT GL_DIST_AMT, INV_PMT AMT_CHB, INV_PMT AMT_DED FROM ARTCASH3 WHERE ROWNUM < 1"
    '    Dim TT_PIC As String
    '    TT_PIC = Temp_Table(sql)
    '    OraD.ExecuteSQL("ALTER TABLE " & TT_PIC & " ADD PRIMARY KEY (INV_TYPE, INV_NUM)")

    '    sql = ""
    '    sql = sql & " BEGIN DECLARE CURSOR C1 IS SELECT DISTINCT ORDR_INV_TYPE INV_TYPE, ORDR_INV_NO INV_NUM FROM " & TTC & ";"
    '    sql = sql & " BEGIN FOR R1 IN C1 LOOP"
    '    sql = sql & " INSERT INTO " & TT_PIC
    '    sql = sql & " SELECT R1.INV_TYPE, R1.INV_NUM, NVL(CUST_CHECK_AMT,0) CUST_CHECK_AMT,"
    '    sql = sql & " NVL(INV_PMT,0) INV_PMT, NVL(INV_DISC_TAKEN,0) INV_DISC_TAKEN, NVL(INV_WRITE_OFF,0) INV_WRITE_OFF,"
    '    sql = sql & " NVL(INV_PMT_I,0) INV_PMT_I, NVL(INV_DISC_TAKEN_I,0) INV_DISC_TAKEN_I, NVL(INV_WRITE_OFF_I,0) INV_WRITE_OFF_I,"
    '    sql = sql & " NVL(INV_PMT_NI,0) INV_PMT_NI, NVL(INV_DISC_TAKEN_NI,0) INV_DISC_TAKEN_NI, NVL(INV_WRITE_OFF_NI,0) INV_WRITE_OFF_NI,"
    '    sql = sql & " NVL(GL_DIST_AMT,0) GL_DIST_AMT , NVL(AMT_CHB,0) AMT_CHB, NVL(AMT_DED,0) AMT_DED"
    '    sql = sql & " FROM ("
    '    sql = sql & " SELECT 'Z' C2_KEY, SUM (CUST_CHECK_AMT) CUST_CHECK_AMT FROM ARTCASH2 WHERE (PYMT_BATCH_NO, PYMT_BATCH_LNO) IN ("
    '    sql = sql & " SELECT PYMT_BATCH_NO, PYMT_BATCH_LNO FROM ARTCASH3 WHERE INV_TYPE = R1.INV_TYPE AND INV_NUM = R1.INV_NUM)"
    '    sql = sql & " ) C2, ("
    '    sql = sql & " SELECT 'Z' C3_KEY, SUM (INV_PMT) INV_PMT, SUM (INV_DISC_TAKEN) INV_DISC_TAKEN, SUM (INV_WRITE_OFF) INV_WRITE_OFF"
    '    sql = sql & " FROM ARTCASH3 WHERE INV_TYPE = R1.INV_TYPE AND INV_NUM = R1.INV_NUM"
    '    sql = sql & " ) C3, ("
    '    sql = sql & " SELECT 'Z' C3_I_KEY, SUM (INV_PMT) INV_PMT_I, SUM (INV_DISC_TAKEN) INV_DISC_TAKEN_I, SUM (INV_WRITE_OFF) INV_WRITE_OFF_I"
    '    sql = sql & " FROM ARTCASH3 WHERE INV_TYPE = 'I' AND (PYMT_BATCH_NO, PYMT_BATCH_LNO) IN ("
    '    sql = sql & " SELECT PYMT_BATCH_NO, PYMT_BATCH_LNO FROM ARTCASH3 WHERE INV_TYPE = R1.INV_TYPE AND INV_NUM = R1.INV_NUM)"
    '    sql = sql & " ) C3_I, ("
    '    sql = sql & " SELECT 'Z' C3_NI_KEY, SUM (INV_PMT) INV_PMT_NI, SUM (INV_DISC_TAKEN) INV_DISC_TAKEN_NI, SUM (INV_WRITE_OFF) INV_WRITE_OFF_NI"
    '    sql = sql & " FROM ARTCASH3 WHERE INV_TYPE IN ('R','C') AND (PYMT_BATCH_NO, PYMT_BATCH_LNO) IN ("
    '    sql = sql & " SELECT PYMT_BATCH_NO, PYMT_BATCH_LNO FROM ARTCASH3 WHERE INV_TYPE = R1.INV_TYPE AND INV_NUM = R1.INV_NUM)"
    '    sql = sql & " ) C3_NI, ("
    '    sql = sql & " SELECT 'Z' C4_KEY, SUM (GL_DIST_AMT) GL_DIST_AMT FROM ARTCASH4 WHERE (PYMT_BATCH_NO, PYMT_BATCH_LNO) IN ("
    '    sql = sql & " SELECT PYMT_BATCH_NO, PYMT_BATCH_LNO FROM ARTCASH3 WHERE INV_TYPE = R1.INV_TYPE AND INV_NUM = R1.INV_NUM)"
    '    sql = sql & " ) C4, ("
    '    sql = sql & " SELECT 'Z' C5_KEY, SUM (DECODE(NVL(CHARGEBACK_IND,0),-1,GL_DIST_AMT)) AMT_CHB"
    '    sql = sql & " , SUM (DECODE(NVL(CHARGEBACK_IND,0),0,GL_DIST_AMT)) AMT_DED"
    '    sql = sql & " FROM ARTCASH5 WHERE (PYMT_BATCH_NO, PYMT_BATCH_LNO) IN ("
    '    sql = sql & " SELECT PYMT_BATCH_NO, PYMT_BATCH_LNO FROM ARTCASH3 WHERE INV_TYPE = R1.INV_TYPE AND INV_NUM = 'W039102')"
    '    sql = sql & " ) C5"
    '    sql = sql & " WHERE C3_KEY = C2_KEY AND C3_I_KEY = C2_KEY AND C3_NI_KEY = C2_KEY AND C4_KEY (+) = C2_KEY AND C5_KEY (+) = C2_KEY;"
    '    sql = sql & " END LOOP; END; END;"
    '    OraD.ExecuteSQL(sql)

    '    sql = "Select * from " & TT_PIC
    '    Call Ora_to_Acc(Nothing, "SOWINVHP", 2, "", sql)

    '    sql = "Update SOWINVHC,SOWINVHP set SOWINVHC.CONT_PYMT_PAID_IN_CASH =  '1'"
    '    sql = sql & " where SOWINVHP.INV_TYPE = SOWINVHC.ORDR_INV_TYPE"
    '    sql = sql & "   and SOWINVHP.INV_NUM = SOWINVHC.ORDR_INV_NO"
    '    sql = sql & "   and SOWINVHP.INV_PMT_NI = 0"
    '    sql = sql & "   and SOWINVHP.AMT_CHB = 0"
    '    sql = sql & "   and SOWINVHP.AMT_DED = 0"
    '    sql = sql & "   and SOWINVHP.INV_PMT >= SOWINVHC.ORDR_AMT"
    '    AccD.Execute(sql)

    '    ' Flag all Payable Records which have been satisfactorily paid by Customers


    '    sql = "Update SOWINVHC SET CONT_PYMT_LATE = '1' "
    '    sql = sql & " where ORDR_INV_DATE < (INV_LAST_PMT - 60)"
    '    AccD.Execute(sql)

    '    If xFormName = "SPFREBJ1" Then

    '        sql = "Update SOWINVHC SET CONT_PYMT_OK_TO_PAY = '1' "
    '        '  sql = sql & " where ORDR_TOTAL_AMT <= 0 "
    '        '  sql = sql & " or ((INV_BALANCE = 0 or INV_BALANCE <= (ORDR_WD_CHG + ORDR_MISC_CHG))"
    '        '  sql = sql & " and ((CONT_PYMT_PAID_IN_CASH = '1' and CONT_PYMT_LATE <> '1') "
    '        '  sql = sql & "   or IIf(IsNull(CONT_PYMT_CRED_REL),'0',CONT_PYMT_CRED_REL) = '1'"
    '        '  sql = sql & "     )"
    '        '  sql = sql & "    )"
    '        AccD.Execute(sql)

    '    Else
    '        sql = "Update SOWINVHC SET CONT_PYMT_OK_TO_PAY = '1' "
    '        '  sql = sql & " where ORDR_TOTAL_AMT <= 0 "
    '        '  sql = sql & " or ((INV_BALANCE = 0 or INV_BALANCE <= (ORDR_WD_CHG + ORDR_MISC_CHG))"
    '        '  sql = sql & " and ((CONT_PYMT_PAID_IN_CASH = '1' and CONT_PYMT_LATE <> '1') "
    '        '  sql = sql & "   or IIf(IsNull(CONT_PYMT_CRED_REL),'0',CONT_PYMT_CRED_REL) = '1'"
    '        '   sql = sql & "     )"
    '        '   sql = sql & "    )"
    '        AccD.Execute(sql)
    '    End If


    'End Sub

    'Function ZERO_CHECK(ByVal PYMT_DATE As Date, ByVal PYMT_REF As String, ByVal LAST_DATE As Object, ByVal ar() As String, ByVal CUST_CODE As String, ByVal REASON_CODE As String, ByVal ACCT_CODE As String, ByVal SEG2_CODE As String, ByVal SEG3_CODE As String, ByVal SEG4_CODE As String, ByVal pay_in_full As Boolean) As String

    '    ' AR(0,I) = INV_TYPE
    '    ' AR(1,I) = INV_NUM
    '    ' AR(2,I) = INV_PMT

    '    Dim dynARTPARM1 As OraDynaset
    '    sql = "Select * from ARTPARM1 where AR_PARM_KEY = 'Z'"
    '    dynARTPARM1 = OraD.CreateDynaset(sql, 8&)

    '    Dim dynARTOPEN1 As OraDynaset
    '    sql = "Select * from ARTOPEN1 where CUST_CODE = :CUST_CODE "
    '    sql = sql & " and INV_TYPE = :INV_TYPE and INV_NUM = :INV_NUM"
    '    dynARTOPEN1 = OraD.CreateDynaset(sql, 0&)

    '    Dim PYMT_BATCH_NO As String
    '    PYMT_BATCH_NO = CTLNO("PYMT_BATCH_NO", 10)
    '    ZERO_CHECK = PYMT_BATCH_NO

    '    Dim dynARTCUST1 As OraDynaset
    '    sql = "Select * from ARTCUST1 where CUST_CODE = :CUST_CODE"
    '    dynARTCUST1 = OraD.CreateDynaset(sql, 8&)
    '    OraD.Parameters("CUST_CODE").Value = CUST_CODE
    '    dynARTCUST1.Refresh()

    '    Dim dynARTCASH1 As OraDynaset
    '    sql = "Select * from ARTCASH1 where ROWNUM < 1"
    '    dynARTCASH1 = OraD.CreateDynaset(sql, 0&)
    '    With dynARTCASH1
    '        .AddNew()
    '        .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '        .Fields("PYMT_BATCH_DATE").Value = PYMT_DATE
    '        .Fields("STATUS").Value = "1"
    '        .Fields("INIT_OPER").Value = xUserID
    '        .Fields("LAST_OPER").Value = xUserID
    '        .Fields("INIT_DATE").Value = LAST_DATE
    '        .Fields("LAST_DATE").Value = LAST_DATE
    '        .Fields("ZERO_CHECK").Value = "1"
    '        .Fields("OPS_YYYYPP").Value = CYP
    '        .Fields("CURR_CODE").Value = dynARTPARM1.Fields("AR_PARM_CURR_CODE").Value
    '        .Fields("CURR_EXCH_RATE").Value = 1
    '        .Update()
    '    End With

    '    Dim dynARTCASH2 As OraDynaset
    '    sql = "Select * from ARTCASH2 where ROWNUM < 1"
    '    dynARTCASH2 = OraD.CreateDynaset(sql, 0&)
    '    With dynARTCASH2
    '        .AddNew()
    '        .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '        .Fields("PYMT_BATCH_LNO").Value = 1
    '        .Fields("CUST_CODE").Value = CUST_CODE
    '        .Fields("CUST_NAME").Value = dynARTCUST1.Fields("CUST_NAME").Value & ""
    '        .Fields("CUST_CHECK_NO").Value = PYMT_REF
    '        .Fields("CUST_CHECK_DATE").Value = PYMT_DATE
    '        .Fields("CUST_CHECK_AMT").Value = 0
    '        .Fields("STATUS").Value = "2"
    '        .Fields("CUST_CHECK_AMT_CURR").Value = 0
    '        .Update()
    '    End With

    '    Dim dynARTCASH3 As OraDynaset
    '    sql = "Select * from ARTCASH3 where ROWNUM < 1"
    '    dynARTCASH3 = OraD.CreateDynaset(sql, 0&)

    '    Dim INV_PMT As Double
    '    Dim INV_BALANCE As Double
    '    Dim PYMT_APPL_TOTAL As Double

    '    PYMT_APPL_TOTAL = 0
    '    Dim i As Integer
    '    For i = 1 To UBound(ar, 2)
    '        OraD.Parameters("CUST_CODE").Value = CUST_CODE
    '        OraD.Parameters("INV_TYPE").Value = ar(0, i)
    '        OraD.Parameters("INV_NUM").Value = ar(1, i)
    '        dynARTOPEN1.Refresh()
    '        If dynARTOPEN1.EOF Then
    '            sql = "INSERT INTO ARTOPEN1 SELECT * FROM ARTOPENX WHERE CUST_CODE = '" & CUST_CODE & "'"
    '            sql = sql & " AND INV_TYPE = '" & ar(0, i) & "' AND INV_NUM = '" & ar(1, i) & "'"
    '            OraD.ExecuteSQL(sql)
    '            sql = "DELETE FROM ARTOPENX WHERE CUST_CODE = '" & CUST_CODE & "'"
    '            sql = sql & " AND INV_TYPE = '" & ar(0, i) & "' AND INV_NUM = '" & ar(1, i) & "'"
    '            OraD.ExecuteSQL(sql)
    '            dynARTOPEN1.Refresh()
    '        End If
    '        INV_BALANCE = Val(dynARTOPEN1.Fields("INV_BALANCE").Value & "")
    '        If pay_in_full Then
    '            INV_PMT = INV_BALANCE
    '        Else
    '            INV_PMT = Val(ar(2, i))
    '        End If
    '        PYMT_APPL_TOTAL = PYMT_APPL_TOTAL + INV_PMT

    '        With dynARTCASH3
    '            .AddNew()
    '            .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '            .Fields("PYMT_BATCH_LNO").Value = 1
    '            .Fields("PYMT_BATCH_ILNO").Value = i
    '            .Fields("INV_TYPE").Value = ar(0, i)
    '            .Fields("INV_NUM").Value = ar(1, i)
    '            .Fields("REASON_CODE").Value = dynARTOPEN1.Fields("REASON_CODE").Value
    '            .Fields("INV_DATE").Value = dynARTOPEN1.Fields("INV_DATE").Value
    '            .Fields("INV_DUE_DATE").Value = dynARTOPEN1.Fields("INV_DUE_DATE").Value
    '            .Fields("CUST_CODE_SO").Value = dynARTOPEN1.Fields("CUST_CODE_SO").Value
    '            .Fields("CUST_SHIP_TO_CODE").Value = dynARTOPEN1.Fields("CUST_SHIP_TO_CODE").Value
    '            .Fields("INV_CUST_PO").Value = dynARTOPEN1.Fields("INV_CUST_PO").Value
    '            .Fields("INV_BALANCE").Value = INV_BALANCE
    '            .Fields("INV_PMT").Value = INV_PMT
    '            .Fields("INV_DISC_TAKEN").Value = 0
    '            .Fields("INV_WRITE_OFF").Value = 0
    '            .Fields("INV_BALANCE_NEW").Value = INV_BALANCE - INV_PMT
    '            .Fields("POST_CODE").Value = dynARTOPEN1.Fields("POST_CODE").Value
    '            .Fields("SEG2_CODE").Value = dynARTOPEN1.Fields("SEG2_CODE").Value
    '            .Fields("SEG3_CODE").Value = dynARTOPEN1.Fields("SEG3_CODE").Value
    '            .Fields("SEG4_CODE").Value = dynARTOPEN1.Fields("SEG4_CODE").Value
    '            .Fields("INV_BALANCE_CURR").Value = INV_BALANCE - INV_PMT
    '            .Fields("INV_PMT_CURR").Value = INV_PMT
    '            .Fields("INV_DISC_TAKEN_CURR").Value = 0
    '            .Fields("INV_WRITE_OFF_CURR").Value = 0
    '            .Fields("INV_BALANCE_NEW_CURR").Value = 0
    '            .Update()
    '        End With

    '        With dynARTOPEN1
    '            .Edit()
    '            .Fields("INV_LAST_PMT").Value = Format$(LAST_DATE, "MM/DD/YYYY")
    '            .Fields("INV_PMT").Value = Val(.Fields("INV_PMT").Value & "") + INV_PMT
    '            .Fields("INV_BALANCE").Value = INV_BALANCE - INV_PMT
    '            .Fields("INV_LAST_PMT_REF").Value = PYMT_REF
    '            .Fields("INV_LAST_PMT_REF_DT").Value = PYMT_DATE
    '            .Fields("LAST_OPER").Value = xUserID
    '            .Fields("LAST_DATE").Value = LAST_DATE
    '            .Fields("INV_PMT_CURR").Value = Val(.Fields("INV_PMT").Value & "")
    '            .Fields("INV_BALANCE_CURR").Value = INV_BALANCE - INV_PMT
    '            .Update()
    '        End With
    '    Next i

    '    If PYMT_APPL_TOTAL <> 0 Then
    '        If REASON_CODE <> "" Then
    '            Dim dynARTCASH5 As OraDynaset
    '            sql = "Select * from ARTCASH5 where ROWNUM < 1"
    '            dynARTCASH5 = OraD.CreateDynaset(sql, 0&)
    '            With dynARTCASH5
    '                .AddNew()
    '                .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '                .Fields("PYMT_BATCH_LNO").Value = 1
    '                .Fields("PYMT_BATCH_DLNO").Value = 1
    '                .Fields("REASON_CODE").Value = REASON_CODE
    '                .Fields("GL_DIST_AMT").Value = -1 * PYMT_APPL_TOTAL
    '                .Fields("CHARGEBACK_IND").Value = 0
    '                .Fields("CHARGEBACK_NO").Value = Null
    '                .Fields("CUST_REFERENCE").Value = PYMT_REF

    '                .Fields("CUST_CODE_SO").Value = CUST_CODE
    '                If SEG2_CODE <> "" Then
    '                    .Fields("SEG2_CODE").Value = SEG2_CODE
    '                Else
    '                    .Fields("SEG2_CODE").Value = dynARTPARM1.Fields("AR_PARM_DEF_SEG2").Value
    '                End If
    '                If SEG3_CODE <> "" Then
    '                    .Fields("SEG3_CODE").Value = SEG3_CODE
    '                Else
    '                    .Fields("SEG3_CODE").Value = dynARTPARM1.Fields("AR_PARM_DEF_SEG3").Value
    '                End If
    '                If SEG4_CODE <> "" Then
    '                    .Fields("SEG4_CODE").Value = SEG4_CODE
    '                Else
    '                    .Fields("SEG4_CODE").Value = dynARTPARM1.Fields("AR_PARM_DEF_SEG4").Value
    '                End If
    '                .Fields("INV_TYPE_CB").Value = Null
    '                .Fields("OUR_REFERENCE").Value = ""
    '                .Fields("GL_DIST_AMT_CURR").Value = PYMT_APPL_TOTAL
    '                .Update()
    '            End With
    '        ElseIf ACCT_CODE <> "" Then
    '            Dim dynARTCASH4 As OraDynaset
    '            sql = "Select * from ARTCASH4 where ROWNUM < 1"
    '            dynARTCASH4 = OraD.CreateDynaset(sql, 0&)
    '            With dynARTCASH4
    '                .AddNew()
    '                .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '                .Fields("PYMT_BATCH_LNO").Value = 1
    '                .Fields("PYMT_BATCH_GLNO").Value = 1
    '                .Fields("ACCT_CODE").Value = ACCT_CODE
    '                .Fields("GL_DIST_AMT").Value = PYMT_APPL_TOTAL
    '                .Fields("GL_DIST_REF").Value = PYMT_REF
    '                If SEG2_CODE <> "" Then
    '                    .Fields("SEG2_CODE").Value = SEG2_CODE
    '                Else
    '                    .Fields("SEG2_CODE").Value = dynARTPARM1.Fields("AR_PARM_DEF_SEG2").Value
    '                End If
    '                If SEG3_CODE <> "" Then
    '                    .Fields("SEG3_CODE").Value = SEG3_CODE
    '                Else
    '                    .Fields("SEG3_CODE").Value = dynARTPARM1.Fields("AR_PARM_DEF_SEG3").Value
    '                End If
    '                If SEG4_CODE <> "" Then
    '                    .Fields("SEG4_CODE").Value = SEG4_CODE
    '                Else
    '                    .Fields("SEG4_CODE").Value = dynARTPARM1.Fields("AR_PARM_DEF_SEG4").Value
    '                End If
    '                .Fields("GL_DIST_AMT_CURR").Value = -1 * PYMT_APPL_TOTAL
    '                .Update()
    '            End With
    '        End If
    '    End If
    'End Function

    'Function BR_Make_TTC(ByVal TTC As String)

    '    TTC = Temp_Table(sql)
    '    OraD.ExecuteSQL("Alter Table " & TTC & " Add Primary Key (SO_ORDER_NO, CONT_TP_TYPE, CONT_PYMT_VIA, CONT_PYMT_PAYEE_CODE, CONT_PYMT_SEQ)")
    '    OraD.ExecuteSQL("Analyze Table " & TTC & " Compute Statistics")

    '    sql = "Update " & TTC & " TTC set QTY_CASES = "
    '    sql = sql & " (Select Sum (QTY_CASES) from SOTINVH2 "
    '    sql = sql & " where SO_ORDER_NO = TTC.SO_ORDER_NO)"
    '    sql = sql & " where ORDR_INV_TYPE = 'I'"
    '    OraD.ExecuteSQL(sql)
    '    sql = "Update " & TTC & " TTC set QTY_UNITS = "
    '    sql = sql & " (Select Sum (QTY_UNITS) from SOTINVH2 "
    '    sql = sql & " where SO_ORDER_NO = TTC.SO_ORDER_NO)"
    '    sql = sql & " where ORDR_INV_TYPE = 'I'"
    '    OraD.ExecuteSQL(sql)
    '    sql = "Update " & TTC & " TTC set QTY_CASES = "
    '    sql = sql & " (Select Sum (CHG_CASES) from SOTINVH7 "
    '    sql = sql & " where SO_ORDER_NO = TTC.SO_ORDER_NO)"
    '    sql = sql & " where ORDR_INV_TYPE = 'C'"
    '    OraD.ExecuteSQL(sql)
    '    sql = "Update " & TTC & " TTC set QTY_UNITS = "
    '    sql = sql & " (Select Sum (CHG_UNITS) from SOTINVH7 "
    '    sql = sql & " where SO_ORDER_NO = TTC.SO_ORDER_NO)"
    '    sql = sql & " where ORDR_INV_TYPE = 'C'"
    '    OraD.ExecuteSQL(sql)

    '    sql = "Select * from " & TTC
    '    Call Ora_to_Acc(Nothing, "SOWINVHC", 5, "", sql)

    '    sql = "Select SOTINVHH.* from SOTINVHH," & TTC & " TTC" & vbCr
    '    sql = sql & " where TTC.SO_ORDER_NO = SOTINVHH.SO_ORDER_NO"
    '    sql = sql & "   and TTC.CONT_TP_TYPE = SOTINVHH.CONT_TP_TYPE"
    '    sql = sql & "   and TTC.CONT_PYMT_VIA = SOTINVHH.CONT_PYMT_VIA"
    '    sql = sql & "   and TTC.CONT_PYMT_PAYEE_CODE = SOTINVHH.CONT_PYMT_PAYEE_CODE"
    '    sql = sql & "   and TTC.CONT_PYMT_SEQ = SOTINVHH.CONT_PYMT_SEQ"
    '    Call Ora_to_Acc(Nothing, "SOWINVHH", 8, "", sql)

    '    sql = "Update SOWINVHC set CONT_PYMT_AMT_APPL = 0, CONT_PYMT_AMT_WOFF = 0, CONT_TP_NOTES = NULL"
    '    AccD.Execute(sql)
    '    sql = "Update SOWINVHC,SOWINVHH set SOWINVHC.APPLIED = '1'"
    '    sql = sql & ", SOWINVHC.CONT_PYMT_AMT_APPL = SOWINVHC.CONT_PYMT_AMT_APPL + SOWINVHH.CONT_PYMT_AMT_APPL"
    '    sql = sql & " where SOWINVHH.SO_ORDER_NO = SOWINVHC.SO_ORDER_NO"
    '    sql = sql & "   and SOWINVHH.CONT_TP_TYPE = SOWINVHC.CONT_TP_TYPE"
    '    sql = sql & "   and SOWINVHH.CONT_PYMT_VIA = SOWINVHC.CONT_PYMT_VIA"
    '    sql = sql & "   and SOWINVHH.CONT_PYMT_PAYEE_CODE = SOWINVHC.CONT_PYMT_PAYEE_CODE"
    '    sql = sql & "   and SOWINVHH.CONT_PYMT_SEQ = SOWINVHC.CONT_PYMT_SEQ"
    '    AccD.Execute(sql)
    '    sql = "Update SOWINVHC set CONT_PYMT_AMT_BALANCE = CONT_PYMT_AMT_APPR - CONT_PYMT_AMT_APPL"
    '    AccD.Execute(sql)
    '    sql = "Update SOWINVHC set CONT_PYMT_AMT_WOFF = CONT_PYMT_AMT_APPR - CONT_PYMT_AMT_APPL, CONT_PYMT_AMT_BALANCE = 0 where CONT_PYMT_WOFF = '1'"
    '    AccD.Execute(sql)

    '    sql = "Select * from SOTSREP1"
    '    Call Ora_to_Acc(Nothing, "SOWSREP1", 1, "", sql)

    '    sql = "Select Distinct SOTINVHC.    CONT_PYMT_VIA, SOTINVHC.CONT_PYMT_PAYEE_CODE, APTVEND1.VEND_NAME CONT_PYMT_PAYEE_NAME"
    '    sql = sql & " from " & TTC & " SOTINVHC,APTVEND1"
    '    sql = sql & " where APTVEND1.VEND_CODE = SOTINVHC.CONT_PYMT_PAYEE_CODE"
    '    sql = sql & "   and SOTINVHC.CONT_PYMT_VIA = 'P'"
    '    Call Ora_to_Acc(Nothing, "SOWINVHN", 2, "", sql)

    '    sql = "Select Distinct SOTINVHC.CONT_PYMT_VIA, SOTINVHC.CONT_PYMT_PAYEE_CODE, ARTCUST1.CUST_NAME CONT_PYMT_PAYEE_NAME"
    '    sql = sql & " from " & TTC & " SOTINVHC,ARTCUST1"
    '    sql = sql & " where ARTCUST1.CUST_CODE = SOTINVHC.CONT_PYMT_PAYEE_CODE"
    '    sql = sql & "   and SOTINVHC.CONT_PYMT_VIA = 'R'"
    '    Call Ora_to_Acc(Nothing, "SOWINVHN", 2, "N", sql)

    '    sql = "Select USER_ID, USER_NAME from ASTUSER1 where USER_ID in (Select Distinct INIT_OPER from " & TTC & " union Select Distinct LAST_OPER from " & TTC & ")"
    '    Call Ora_to_Acc(Nothing, "ASWUSERX", 1, "", sql)

    'End Function

    'Sub BR_Make_TTB(ByVal TTC As String)
    '    sql = "Select SOTINVHB.*"
    '    sql = sql & ", SOTINVH2.QTY_CASES, SOTINVH2.QTY_UNITS" & vbCr
    '    sql = sql & ", SOTINVH2.ORDR_PRICE_GRS, SOTINVH2.ORDR_PRICE_NET, SOTINVH2.ALLOW_RATE" & vbCr
    '    sql = sql & ", SOTINVH2.PROD_CODE, SOTINVH2.SIZE_CODE, SOTINVH2.LINE_ITEM_DESCR" & vbCr
    '    sql = sql & ", SOTINVH2.BRAND_CODE, SOTINVH2.ORIG_CODE, SOTINVH2.PACK_CODE" & vbCr
    '    sql = sql & ", SOTINVH2.SP_GROUP, SOTINVH2.GRADE_CODE" & vbCr
    '    sql = sql & " from SOTINVHB,SOTINVH2," & TTC & " TTC" & vbCr
    '    sql = sql & " where SOTINVHB.SO_ORDER_NO = SOTINVH2.SO_ORDER_NO (+)" & vbCr
    '    sql = sql & "   and SOTINVHB.SO_ORDER_LNO = SOTINVH2.SO_ORDER_LNO (+)" & vbCr
    '    sql = sql & "   and SOTINVHB.SO_ORDER_NO = TTC.SO_ORDER_NO" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_TP_TYPE = TTC.CONT_TP_TYPE" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_PYMT_VIA = TTC.CONT_PYMT_VIA" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_PYMT_PAYEE_CODE = TTC.CONT_PYMT_PAYEE_CODE" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_PYMT_SEQ = TTC.CONT_PYMT_SEQ" & vbCr
    '    Call Ora_to_Acc(Nothing, "SOWINVHB", 4, "", sql)
    '    Call Create_Index("SOWINVHB", "I_SOWINVHB_1", "SO_ORDER_NO,CONT_TP_TYPE,CONT_PYMT_VIA,CONT_PYMT_PAYEE_CODE,CONT_PYMT_SEQ")

    '    sql = "Select SO_ORDER_NO, SO_ORDER_LNO, SO_SUB_LNO, LINE_ITEM_DESCR" & vbCr
    '    sql = sql & ", ORDR_PRICE_GRS, ORDR_PRICE_NET, ALLOW_RATE" & vbCr
    '    sql = sql & ", CHG_CASES, CHG_UNITS" & vbCr
    '    sql = sql & " from SOTINVH7 " & vbCr
    '    sql = sql & " where SO_ORDER_NO in " & vbCr
    '    sql = sql & " (Select Distinct SO_ORDER_NO from " & TTC & ")" & vbCr
    '    Dim dynSOTINVH7x As OraDynaset
    '    dynSOTINVH7x = OraD.CreateDynaset(sql, 8&)
    '    Do While Not dynSOTINVH7x.EOF
    '        sql = "Update SOWINVHB set "
    '        sql = sql & "  LINE_ITEM_DESCR = '" & dynSOTINVH7x.Fields("LINE_ITEM_DESCR").Value & "'"
    '        sql = sql & ", QTY_CASES = " & dynSOTINVH7x.Fields("CHG_CASES").Value
    '        sql = sql & ", QTY_UNITS = " & dynSOTINVH7x.Fields("CHG_UNITS").Value
    '        sql = sql & ", ORDR_PRICE_GRS = " & dynSOTINVH7x.Fields("ORDR_PRICE_GRS").Value
    '        sql = sql & ", ORDR_PRICE_NET = " & dynSOTINVH7x.Fields("ORDR_PRICE_NET").Value
    '        sql = sql & ", ALLOW_RATE = " & dynSOTINVH7x.Fields("ALLOW_RATE").Value
    '        sql = sql & " where SO_ORDER_NO = '" & dynSOTINVH7x.Fields("SO_ORDER_NO").Value & "'"
    '        sql = sql & "   and SO_ORDER_LNO = " & dynSOTINVH7x.Fields("SO_ORDER_LNO").Value
    '        sql = sql & "   and SO_SUB_LNO = " & dynSOTINVH7x.Fields("SO_SUB_LNO").Value
    '        AccD.Execute(sql)
    '        dynSOTINVH7x.MoveNext()
    '    Loop
    '    dynSOTINVH7x.Close()

    '    sql = "Select SOTINVH0.SO_ORDER_NO, SOTINVH0.SO_ORDER_LNO" & vbCr
    '    sql = sql & ", MIN (SOTINVH0.BRAND_CODE) BRAND_CODE" & vbCr
    '    sql = sql & ", COUNT (DISTINCT SOTINVH0.BRAND_CODE) BRAND_CODE_CNT" & vbCr
    '    sql = sql & ", MIN (SOTINVH0.ORIG_CODE) ORIG_CODE" & vbCr
    '    sql = sql & ", COUNT (DISTINCT SOTINVH0.ORIG_CODE) ORIG_CODE_CNT" & vbCr
    '    sql = sql & ", MIN (SOTINVH0.PACK_CODE) PACK_CODE" & vbCr
    '    sql = sql & ", COUNT (DISTINCT SOTINVH0.PACK_CODE) PACK_CODE_CNT" & vbCr
    '    sql = sql & ", MIN (SOTINVH0.SP_GROUP) SP_GROUP" & vbCr
    '    sql = sql & ", COUNT (DISTINCT SOTINVH0.SP_GROUP) SP_GROUP_CNT" & vbCr
    '    sql = sql & ", MIN (SOTINVH0.GRADE_CODE) GRADE_CODE" & vbCr
    '    sql = sql & ", COUNT (DISTINCT SOTINVH0.GRADE_CODE) GRADE_CODE_CNT" & vbCr
    '    sql = sql & " from SOTINVHB,SOTINVH0," & TTC & " TTC" & sqltables & vbCr
    '    sql = sql & " where SOTINVHB.SO_ORDER_NO = SOTINVH0.SO_ORDER_NO" & vbCr
    '    sql = sql & "   and SOTINVHB.SO_ORDER_LNO = SOTINVH0.SO_ORDER_LNO" & vbCr
    '    sql = sql & "   and SOTINVHB.SO_ORDER_NO = TTC.SO_ORDER_NO" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_TP_TYPE = TTC.CONT_TP_TYPE" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_PYMT_VIA = TTC.CONT_PYMT_VIA" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_PYMT_PAYEE_CODE = TTC.CONT_PYMT_PAYEE_CODE" & vbCr
    '    sql = sql & "   and SOTINVHB.CONT_PYMT_SEQ = TTC.CONT_PYMT_SEQ" & vbCr
    '    sql = sql & " group by SOTINVH0.SO_ORDER_NO, SOTINVH0.SO_ORDER_LNO"
    '    Call Ora_to_Acc(Nothing, "SOWINVHD", 2, "", sql)

    '    sql = "Update SOWINVHB,SOWINVHD Set "
    '    sql = sql & "  SOWINVHB.BRAND_CODE = IIF(BRAND_CODE_CNT = 1,SOWINVHD.BRAND_CODE,'*')"
    '    sql = sql & ", SOWINVHB.ORIG_CODE = IIF(ORIG_CODE_CNT = 1,SOWINVHD.ORIG_CODE,'*')"
    '    sql = sql & ", SOWINVHB.PACK_CODE = IIF(PACK_CODE_CNT = 1,SOWINVHD.PACK_CODE,'*')"
    '    sql = sql & ", SOWINVHB.SP_GROUP = IIF(SP_GROUP_CNT = 1,SOWINVHD.SP_GROUP,'*')"
    '    sql = sql & ", SOWINVHB.GRADE_CODE = IIF(GRADE_CODE_CNT = 1,SOWINVHD.GRADE_CODE,'*')"
    '    sql = sql & " where SOWINVHB.SO_ORDER_NO = SOWINVHD.SO_ORDER_NO"
    '    sql = sql & "   and SOWINVHB.SO_ORDER_LNO = SOWINVHD.SO_ORDER_LNO"
    '    AccD.Execute(sql)

    '    sql = "Update SOWINVHB,SOWINVHC Set "
    '    sql = sql & "  SOWINVHC.CONT_TP_NOTES = SOWINVHB.CONT_TP_NOTES"
    '    sql = sql & " where SOWINVHB.SO_ORDER_NO = SOWINVHC.SO_ORDER_NO"
    '    sql = sql & "   and SOWINVHB.CONT_TP_TYPE = SOWINVHC.CONT_TP_TYPE"
    '    sql = sql & "   and SOWINVHB.CONT_PYMT_VIA = SOWINVHC.CONT_PYMT_VIA"
    '    sql = sql & "   and SOWINVHB.CONT_PYMT_PAYEE_CODE = SOWINVHC.CONT_PYMT_PAYEE_CODE"
    '    sql = sql & "   and SOWINVHB.CONT_PYMT_SEQ = SOWINVHC.CONT_PYMT_SEQ"
    '    AccD.Execute(sql)

    '    sql = "Select SOWINVHC.CONT_TP_TYPE, SOWINVHC.CONT_PYMT_VIA, SOWINVHC.CONT_PYMT_PAYEE_CODE"
    '    sql = sql & ", SOWINVHC.CONT_TP_NOTES"
    '    sql = sql & ", SUM(SOWINVHC.CONT_PYMT_AMT_APPR) AS CONT_PYMT_AMT_APPR"
    '    sql = sql & ", SUM(SOWINVHC.CONT_PYMT_AMT_APPL) AS CONT_PYMT_AMT_APPL"
    '    sql = sql & ", SUM(SOWINVHC.CONT_PYMT_AMT_WOFF) AS CONT_PYMT_AMT_WOFF"
    '    sql = sql & ", SUM(SOWINVHC.CONT_PYMT_AMT_BALANCE) AS CONT_PYMT_AMT_BALANCE"
    '    sql = sql & " into SOWINVHC_RECAP"
    '    sql = sql & " from SOWINVHC"
    '    sql = sql & "  WHERE CONT_PYMT_STATUS = 'A'"
    '    sql = sql & " group by SOWINVHC.CONT_TP_TYPE, SOWINVHC.CONT_PYMT_VIA, SOWINVHC.CONT_PYMT_PAYEE_CODE"
    '    sql = sql & ", SOWINVHC.CONT_TP_NOTES"
    '    AccD.Execute(sql)
    '    Call Create_Index("SOWINVHC_RECAP", "I_SOWINVHC_RECAP_1", "CONT_TP_TYPE, CONT_PYMT_VIA, CONT_PYMT_PAYEE_CODE, CONT_TP_NOTES")

    'End Sub
    'Sub BR_sql_SOTINVHC()

    '    sql = "Select SOTINVHC.*"
    '    sql = sql & ", '0' CONT_PYMT_OK_TO_PAY, '0' CONT_PYMT_PAID_IN_CASH, '0' CONT_PYMT_LATE" & vbCr
    '    sql = sql & ", SOTINVH1.CUST_CODE, SOTINVH1.CUST_NAME, SOTINVH1.CUST_BILL_TO_CUST" & vbCr
    '    sql = sql & ", SOTINVH1.ORDR_INV_TYPE, SOTINVH1.ORDR_INV_NO, SOTINVH1.ORDR_INV_DATE" & vbCr
    '    sql = sql & ", SOTINVH1.CUST_ORDER_NO, SOTINVH1.ORDR_DATE, SOTINVH1.BRKR_CODE" & vbCr
    '    sql = sql & ", SOTINVH1.ORDR_DIV_CODE, SOTINVH1.MARKET_TYPE" & vbCr
    '    sql = sql & ", NVL(SOTINVH1.ORDR_AMT,0) ORDR_AMT" & vbCr
    '    sql = sql & ", NVL(SOTINVH1.ORDR_WD_CHG,0) ORDR_WD_CHG" & vbCr
    '    sql = sql & ", NVL(SOTINVH1.ORDR_MISC_CHG,0) ORDR_MISC_CHG" & vbCr
    '    sql = sql & ", NVL(SOTINVH1.ORDR_TOTAL_AMT,0) ORDR_TOTAL_AMT" & vbCr
    '    sql = sql & ", 0 QTY_CASES, 0.01 QTY_UNITS" & vbCr
    '    sql = sql & ", NVL(ARTOPEN1.INV_BALANCE,0) INV_BALANCE" & vbCr
    '    sql = sql & ", NVL(ARTOPEN1.INV_LAST_PMT,NVL(ARTOPENX.INV_LAST_PMT,TRUNC(SYSDATE))) INV_LAST_PMT" & vbCr
    '    sql = sql & ", '0' RELEASED, '0' CHANGE_PAYEE_SPLIT, '0' CHANGE_PAYEE_JOIN"
    '    sql = sql & ", 0.01 CONT_PYMT_AMT_APPL, 0.01 CONT_PYMT_AMT_BALANCE, 0.01 CONT_PYMT_AMT_WOFF"
    '    sql = sql & ", '0' APPROVED, '0' REJECTED, '0' APPLIED"
    '    sql = sql & ", CONT_PYMT_NOTE_APPR CONT_TP_NOTES"
    '    sql = sql & " from SOTINVHC,SOTINVH1,ARTOPEN1,ARTOPENX " & vbCr
    '    sql = sql & " where SOTINVH1.SO_ORDER_NO = SOTINVHC.SO_ORDER_NO" & vbCr
    '    sql = sql & "   and ARTOPEN1.INV_TYPE (+) = SOTINVH1.ORDR_INV_TYPE" & vbCr
    '    sql = sql & "   and ARTOPEN1.INV_NUM (+) = SOTINVH1.ORDR_INV_NO" & vbCr
    '    sql = sql & "   and ARTOPEN1.CUST_CODE (+) = SOTINVH1.CUST_BILL_TO_CUST" & vbCr
    '    sql = sql & "   and ARTOPENX.INV_TYPE (+) = SOTINVH1.ORDR_INV_TYPE" & vbCr
    '    sql = sql & "   and ARTOPENX.INV_NUM (+) = SOTINVH1.ORDR_INV_NO" & vbCr
    '    sql = sql & "   and ARTOPENX.CUST_CODE (+) = SOTINVH1.CUST_BILL_TO_CUST" & vbCr
    '    If Mid$(xFormName, 4, 3) = "BRK" Then
    '        sql = sql & "   and SOTINVHC.CONT_TP_TYPE = '" & "B" & "'" & vbCr
    '    Else
    '        sql = sql & "   and SOTINVHC.CONT_TP_TYPE = '" & "R" & "'" & vbCr
    '    End If

    'End Sub

    'Sub BRJ_Prepare(LAST_DATE As Date, INV_DATE As Date, got_GL As Boolean, chkMYITEMSONLY As String, Optional chkHISTORY As String = "", Optional RYP0 As String, Optional RYP1 As String)

    '        Dim dynAPTPARM1 As OraDynaset
    '        sql = "Select * from APTPARM1 where AP_PARM_KEY = 'Z'"
    '        dynAPTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynSPTPARM1 As OraDynaset
    '        sql = "Select * from SPTPARM1 where SP_PARM_KEY = 'Z'"
    '        dynSPTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynSOTPARM1 As OraDynaset
    '        sql = "Select * from SOTPARM1 where SO_PARM_KEY = 'Z'"
    '        dynSOTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynGLTPARM1 As OraDynaset
    '        sql = "Select * from GLTPARM1 where GL_PARM_KEY = 'Z'"
    '        dynGLTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynwk As Recordset
    '        Dim dynWK2 As Recordset

    '        ' Build Work Tables

    '        Call Track("Initialize Work Tables", "")

    '        Call BR_sql_SOTINVHC()

    '        If chkHISTORY = "1" Then
    '            sql = sql & " and SOTINVHC.CONT_PYMT_STATUS = 'P'" & vbCr
    '            sql = sql & " and SOTINVHC.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCr
    '        Else
    '            sql = sql & " and SOTINVHC.CONT_PYMT_STATUS IN ('A','X')" & vbCr
    '            sql = sql & " and SOTINVHC.REGISTER_IND = '1'" & vbCr
    '        End If

    '        If chkMYITEMSONLY = "1" Then
    '            sql = sql & " and SOTINVHC.LAST_OPER = '" & xUserID & "'" & vbCr
    '        End If

    '        Dim TTC As String
    '        Call BR_Make_TTC(TTC)

    '        Call BR_Make_TTB(TTC)

    '        ' Create Payments

    '        Dim CONT_PYMT_AMT As Double
    '        Dim CONT_PYMT_AMT_APPR As Double
    '        Dim CONT_PYMT_AMT_WOFF As Double
    '        Dim ORDR_INV_NO As String
    '        Dim CONT_TP_TYPE As String

    '        Dim VOUCHER_NO As String
    '        Dim INV_NUM As String

    '        Dim dynTATTERM1 As OraDynaset
    '        sql = "Select * from TATTERM1 where TERM_CODE = :CODE"
    '        dynTATTERM1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynSOTINVH1 As OraDynaset
    '        sql = "Select * from SOTINVH1 where SO_ORDER_NO = :CODE"
    '        dynSOTINVH1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynAPTVEND1 As OraDynaset
    '        sql = "Select * from APTVEND1 where VEND_CODE = :CODE"
    '        dynAPTVEND1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynARTCUST1 As OraDynaset
    '        sql = "Select * from ARTCUST1 where CUST_CODE = :CODE"
    '        dynARTCUST1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dynARTPARM1 As OraDynaset
    '        sql = "Select * from ARTPARM1 where AR_PARM_KEY = 'Z'"
    '        dynARTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        sql = "Select * from APTINVH1 where ROWNUM <1"
    '        Call Ora_to_Acc(Nothing, "APWINVH1", 1, "", sql)
    '        Dim dynAPWINVH1 As Recordset
    '        dynAPWINVH1 = AccD.OpenRecordset("APWINVH1", dbOpenDynaset)

    '        sql = "Select * from APTINVH2 where ROWNUM <1"
    '        Call Ora_to_Acc(Nothing, "APWINVH2", 2, "", sql)
    '        Dim dynAPWINVH2 As Recordset
    '        dynAPWINVH2 = AccD.OpenRecordset("APWINVH2", dbOpenDynaset)

    '        sql = "Select * from ARTOPEN1 where ROWNUM <1"
    '        Call Ora_to_Acc(Nothing, "ARWOPEN1", 3, "", sql)
    '        Dim dynARWOPEN1 As Recordset
    '        dynARWOPEN1 = AccD.OpenRecordset("ARWOPEN1", dbOpenDynaset)

    '        Dim ACCT_CODE_ACC As String
    '        Dim ACCT_CODE_VAR As String
    '        If Mid$(xFormName, 4, 3) = "BRK" Then
    '            ACCT_CODE_ACC = dynSOTPARM1.Fields("SO_PARM_ACCT_BRKR_ACCR").Value
    '            ACCT_CODE_VAR = dynSOTPARM1.Fields("SO_PARM_ACCT_BRKR_EXP").Value
    '        Else
    '            ACCT_CODE_ACC = dynSOTPARM1.Fields("SO_PARM_ACCT_REBATE_ACCR").Value
    '            ACCT_CODE_VAR = dynSOTPARM1.Fields("SO_PARM_ACCT_REBATE_EXP").Value
    '        End If

    '        Dim XVNO As Integer
    '        XVNO = 0

    '        sql = "Select * from SOWINVHC where CONT_PYMT_STATUS = 'A' "
    '        sql = sql & " order by CONT_PYMT_VIA, CONT_PYMT_PAYEE_CODE, CONT_PYMT_SEQ"
    '        Dim dynSOWINVHC As Recordset
    '        dynSOWINVHC = AccD.OpenRecordset(sql, dbOpenDynaset)
    '        Do While Not dynSOWINVHC.EOF
    '            CONT_TP_TYPE = dynSOWINVHC.Fields("CONT_TP_TYPE").Value
    '            CONT_PYMT_AMT = Round(Val(dynSOWINVHC.Fields("CONT_PYMT_AMT").Value & ""), 2)
    '            CONT_PYMT_AMT_APPR = Round(Val(dynSOWINVHC.Fields("CONT_PYMT_AMT_APPR").Value & ""), 2)
    '            CONT_PYMT_AMT_WOFF = Round(Val(dynSOWINVHC.Fields("CONT_PYMT_AMT_WOFF").Value & ""), 2)
    '            ORDR_INV_NO = dynSOWINVHC.Fields("ORDR_INV_NO").Value
    '            dynSOWINVHC.Edit()

    '            If dynSOWINVHC.Fields("CONT_PYMT_VIA").Value = "P" Then
    '                OraD.Parameters("CODE").Value = dynSOWINVHC.Fields("CONT_PYMT_PAYEE_CODE").Value
    '                dynAPTVEND1.Refresh()

    '                XVNO = XVNO + 1
    '                VOUCHER_NO = "X" & Format$(XVNO, "000000000")
    '                'VOUCHER_NO = CTLNO("APTINVH1", 10)
    '                dynSOWINVHC.Fields("CONT_PYMT_NO").Value = VOUCHER_NO
    '            GoSub Update_AP
    '            Else
    '                OraD.Parameters("CODE").Value = dynSOWINVHC.Fields("CONT_PYMT_PAYEE_CODE").Value
    '                dynARTCUST1.Refresh()

    '                INV_NUM = CTLNO("AR_ON_ACCT_NO", 6)
    '                dynSOWINVHC.Fields("CONT_PYMT_NO").Value = INV_NUM
    '            GoSub Update_AR
    '            End If

    '            dynSOWINVHC.Fields("CONT_PYMT_STMT_XNO").Value = xXNO
    '            dynSOWINVHC.Fields("OPS_YYYYPP").Value = CYP
    '            dynSOWINVHC.Update()
    '            dynSOWINVHC.MoveNext()
    '        Loop

    '        ' Consolidate Vouchers into 1 per Payee

    '        If XVNO <> 0 Then
    '            Dim tblAPWINVH1 As Recordset
    '            tblAPWINVH1 = AccD.OpenRecordset("APWINVH1", dbOpenTable)
    '            tblAPWINVH1.Index = "PrimaryKey"

    '            sql = "Select APWINVH1.VEND_CODE, Min (APWINVH1.VOUCHER_NO) as VOUCHER_NO"
    '            sql = sql & ", Sum (APWINVH1.INV_AMT) as INV_AMT"
    '            sql = sql & ", Sum (APWINVH1.INV_DISC_BASED_ON) as INV_DISC_BASED_ON"
    '            sql = sql & ", Sum (APWINVH1.INV_PURCHASES) as INV_PURCHASES"
    '            sql = sql & ", Sum (APWINVH1.INV_BALANCE) as INV_BALANCE"
    '            sql = sql & " from APWINVH1 group by APWINVH1.VEND_CODE"
    '            dynwk = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    '            Do While Not dynwk.EOF
    '                VOUCHER_NO = CTLNO("APTINVH1", 10)
    '                sql = "Select VOUCHER_NO from APWINVH1 "
    '                sql = sql & " where VEND_CODE = '" & dynwk.Fields("VEND_CODE").Value & "'"
    '                Dim SQLX As String
    '                SQLX = sql

    '                sql = "Update SOWINVHC set CONT_PYMT_NO = '" & VOUCHER_NO & "'"
    '                sql = sql & " where CONT_PYMT_NO in (" & SQLX & ")"
    '                AccD.Execute(sql)

    '                Dim VOUCHER_LNO_ctr As Integer
    '                VOUCHER_LNO_ctr = 0
    '                sql = "Select APWINVH2.ACCT_CODE"
    '                sql = sql & ", APWINVH2.SEG2_CODE, APWINVH2.SEG3_CODE, APWINVH2.SEG4_CODE"
    '                sql = sql & ", SUM (APWINVH2.INV_LINE_AMT) as INV_LINE_AMT"
    '                sql = sql & " from APWINVH2 where APWINVH2.VOUCHER_NO in (" & SQLX & ")"
    '                sql = sql & " group by APWINVH2.ACCT_CODE"
    '                sql = sql & ", APWINVH2.SEG2_CODE, APWINVH2.SEG3_CODE, APWINVH2.SEG4_CODE"
    '                dynWK2 = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    '                Do While Not dynWK2.EOF

    '                    With dynAPWINVH2
    '                        .AddNew()
    '                        .Fields("VOUCHER_NO").Value = VOUCHER_NO
    '                        VOUCHER_LNO_ctr = VOUCHER_LNO_ctr + 1
    '                        .Fields("VOUCHER_LNO").Value = VOUCHER_LNO_ctr
    '                        .Fields("ACCT_CODE").Value = dynWK2.Fields("ACCT_CODE").Value
    '                        .Fields("SEG2_CODE").Value = dynWK2.Fields("SEG2_CODE").Value
    '                        .Fields("SEG3_CODE").Value = dynWK2.Fields("SEG3_CODE").Value
    '                        .Fields("SEG4_CODE").Value = dynWK2.Fields("SEG4_CODE").Value
    '                        .Fields("INV_LINE_AMT").Value = Val(dynWK2.Fields("INV_LINE_AMT").Value & "")
    '                        .Update()
    '                    End With

    '                    dynWK2.MoveNext()
    '                Loop
    '                dynWK2.Close()

    '                sql = "Delete from APWINVH2 where VOUCHER_NO in (" & SQLX & ")"
    '                AccD.Execute(sql)

    '                sql = "Update APWINVH1 Set VOUCHER_NO = '" & VOUCHER_NO & "'"
    '                sql = sql & " where VOUCHER_NO = '" & dynwk.Fields("VOUCHER_NO").Value & "'"
    '                AccD.Execute(sql)

    '                sql = "Delete from APWINVH1 "
    '                sql = sql & " where VOUCHER_NO in (" & SQLX & ")"
    '                sql = sql & "   and VOUCHER_NO <> '" & VOUCHER_NO & "'"
    '                AccD.Execute(sql)

    '                tblAPWINVH1.Seek("=", VOUCHER_NO)
    '                tblAPWINVH1.Edit()
    '                tblAPWINVH1.Fields("INV_AMT").Value = dynwk.Fields("INV_AMT").Value
    '                tblAPWINVH1.Fields("INV_DISC_BASED_ON").Value = dynwk.Fields("INV_DISC_BASED_ON").Value
    '                tblAPWINVH1.Fields("INV_PURCHASES").Value = dynwk.Fields("INV_PURCHASES").Value
    '                tblAPWINVH1.Fields("INV_BALANCE").Value = dynwk.Fields("INV_BALANCE").Value
    '                tblAPWINVH1.Update()

    '                dynwk.MoveNext()
    '            Loop
    '            dynwk.Close()
    '            tblAPWINVH1.Close()
    '        End If

    '        Call SPBMAINX.BRJ_Build_GL_Update(got_GL)

    '        sql = "Select Count (*) from SOWINVHC where CONT_PYMT_STATUS IN ('A','X')"
    '        dynwk = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    '        If Val(dynwk.Fields(0).Value & "") = 0 Then
    '            xRWU = "N"
    '            'objASCCALLB.xErrMsg = "Nothing to Update"
    '        End If

    '        Exit Sub

    'Update_AP:

    '        Dim inv_amt As Double
    '        inv_amt = CONT_PYMT_AMT_APPR - CONT_PYMT_AMT_WOFF

    '        With dynAPWINVH1
    '            .AddNew()
    '            .Fields("VOUCHER_NO").Value = VOUCHER_NO
    '            .Fields("VEND_CODE").Value = dynAPTVEND1.Fields("VEND_CODE").Value
    '            .Fields("INV_TYPE").Value = "I"
    '            If CONT_TP_TYPE = "B" Then
    '                .Fields("INV_NUM").Value = "BRKR_" & xXNO
    '            Else
    '                .Fields("INV_NUM").Value = "REBATE_" & xXNO
    '            End If
    '            .Fields("INV_DATE").Value = Format$(LAST_DATE, "MM/DD/YYYY")
    '            .Fields("INV_AMT").Value = inv_amt
    '            .Fields("INV_DISC_BASED_ON").Value = inv_amt
    '            .Fields("INV_PURCHASES").Value = inv_amt
    '            .Fields("INV_FREIGHT_AMT").Value = 0
    '            .Fields("INV_SALES_TAX_AMT").Value = 0
    '            .Fields("INV_PAYMENT_CYCLE").Value = CONT_TP_TYPE
    '            .Fields("INV_STATUS").Value = "O"
    '            If dynAPTVEND1.Fields("TERM_CODE").Value & "" = "" Then
    '                .Fields("TERM_CODE").Value = dynAPTPARM1.Fields("AP_PARM_TERM_CODE").Value & ""
    '            Else
    '                .Fields("TERM_CODE").Value = dynAPTVEND1.Fields("TERM_CODE").Value & ""
    '            End If
    '            OraD.Parameters("CODE").Value = .Fields("TERM_CODE").Value & ""
    '            dynTATTERM1.Refresh()
    '            .Fields("INV_DUE_DATE").Value = DateAdd("d", Val(dynTATTERM1.Fields("TERM_DAYS_DUE").Value & ""), Format$(LAST_DATE, "MM/DD/YYYY"))
    '            .Fields("INV_DISC_DUE").Value = Null
    '            .Fields("INV_DISC_AMT").Value = 0
    '            .Fields("INV_COMMENT").Value = "See Statement " & xXNO & " for Details"
    '            .Fields("INV_PAYMENTS").Value = 0
    '            .Fields("INV_DISC_TAKEN").Value = 0
    '            .Fields("INV_BALANCE").Value = inv_amt
    '            .Fields("OPS_YYYYPP").Value = CYP
    '            .Fields("POST_CODE").Value = dynAPTVEND1.Fields("POST_CODE").Value
    '            .Fields("REASON_CODE").Value = "I"
    '            .Fields("INIT_OPER").Value = xUserID
    '            .Fields("LAST_OPER").Value = xUserID
    '            .Fields("INIT_DATE").Value = LAST_DATE
    '            .Fields("LAST_DATE").Value = LAST_DATE
    '            .Fields("INV_1099_AMT").Value = 0
    '            .Fields("INV_CURR_CODE").Value = "USD"
    '            .Fields("INV_CURR_EXCH_RATE").Value = 1
    '            .Fields("INV_WHSE_ALLOW_AMT").Value = 0
    '            .Fields("INV_OTHER_AMT").Value = 0
    '            '.Fields("VEND_SEP_CHECKS").Value = "1"
    '            If dynAPTVEND1.Fields("VEND_CODE_AP").Value & "" <> "" Then
    '                .Fields("VEND_CODE_AP").Value = dynAPTVEND1.Fields("VEND_CODE_AP").Value
    '            Else
    '                .Fields("VEND_CODE_AP").Value = dynAPTVEND1.Fields("VEND_CODE").Value
    '            End If
    '            .Fields("VEND_WHSE_ALLOW").Value = 0
    '            .Fields("INV_PURCH_DISC").Value = 0
    '            If dynAPTVEND1.Fields("BANK_CODE").Value & "" = "" Then
    '                .Fields("BANK_CODE").Value = dynAPTPARM1.Fields("AP_PARM_BANK_CODE").Value & ""
    '            Else
    '                .Fields("BANK_CODE").Value = dynAPTVEND1.Fields("BANK_CODE").Value & ""
    '            End If
    '            .Fields("INV_QTY").Value = 0
    '            .Fields("APPLY_PYMT").Value = 0
    '            .Fields("APPLY_DISC").Value = 0
    '            .Fields("INV_RETAIL").Value = 0
    '            .Fields("SEG2_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG2").Value
    '            .Fields("SEG3_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG3").Value
    '            .Fields("SEG4_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG4").Value
    '            .Fields("INV_RECUR_AMT_MAX").Value = 0
    '            .Fields("INV_RECUR_AMT_GEN").Value = 0
    '            .Fields("VEND_ADV_INT_PCT").Value = 0
    '            .Fields("REGISTER_IND").Value = Null
    '            .Fields("DATE_RECEIVED").Value = Null
    '            .Update()
    '        End With

    '        With dynAPWINVH2
    '            .AddNew()
    '            .Fields("VOUCHER_NO").Value = VOUCHER_NO
    '            .Fields("VOUCHER_LNO").Value = 1
    '            .Fields("ACCT_CODE").Value = ACCT_CODE_ACC
    '            .Fields("SEG2_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG2").Value
    '            .Fields("SEG3_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG3").Value
    '            .Fields("SEG4_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG4").Value
    '            .Fields("INV_LINE_AMT").Value = CONT_PYMT_AMT
    '            .Update()
    '            If CONT_PYMT_AMT_APPR <> CONT_PYMT_AMT Then
    '                .AddNew()
    '                .Fields("VOUCHER_NO").Value = VOUCHER_NO
    '                .Fields("VOUCHER_LNO").Value = 2
    '                .Fields("ACCT_CODE").Value = ACCT_CODE_VAR
    '                .Fields("SEG2_CODE").Value = dynSOWINVHC.Fields("ORDR_DIV_CODE").Value
    '                .Fields("SEG3_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG3").Value
    '                .Fields("SEG4_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG4").Value
    '                .Fields("INV_LINE_AMT").Value = CONT_PYMT_AMT_APPR - CONT_PYMT_AMT
    '                .Update()
    '            End If
    '            If CONT_PYMT_AMT_WOFF <> 0 Then
    '                .AddNew()
    '                .Fields("VOUCHER_NO").Value = VOUCHER_NO
    '                .Fields("VOUCHER_LNO").Value = 3
    '                .Fields("ACCT_CODE").Value = ACCT_CODE_VAR
    '                .Fields("SEG2_CODE").Value = dynSOWINVHC.Fields("ORDR_DIV_CODE").Value
    '                .Fields("SEG3_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG3").Value
    '                .Fields("SEG4_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG4").Value
    '                .Fields("INV_LINE_AMT").Value = -1 * CONT_PYMT_AMT_WOFF
    '                .Update()
    '            End If
    '        End With

    '        Return

    'Update_AR:

    '        OraD.Parameters("CODE").Value = dynSOWINVHC.Fields("SO_ORDER_NO").Value
    '        dynSOTINVH1.Refresh()

    '        With dynARWOPEN1
    '            .AddNew()

    '            .Fields("CUST_CODE").Value = dynARTCUST1.Fields("CUST_CODE").Value
    '            .Fields("INV_TYPE").Value = "O"
    '            .Fields("INV_NUM").Value = INV_NUM
    '            .Fields("INV_DATE").Value = INV_DATE
    '            .Fields("CUST_SHIP_TO_CODE").Value = ""
    '            .Fields("POST_CODE").Value = dynSPTPARM1.Fields("SP_PARM_POST_CODE_AR").Value
    '            .Fields("TERM_CODE").Value = "00"
    '            .Fields("INV_DUE_DATE").Value = INV_DATE
    '            .Fields("INV_DISC_DATE").Value = Null
    '            .Fields("SREP_CODE").Value = dynSOWINVHC.Fields("SREP_CODE").Value
    '            .Fields("STAX_CODE").Value = ""
    '            .Fields("APPLY_TO_INV_NUM").Value = ""
    '            .Fields("APPLY_TO_INV_TYPE").Value = ""
    '            .Fields("INV_CUST_PO").Value = dynSOTINVH1.Fields("CUST_ORDER_NO").Value
    '            .Fields("INV_SALES_ORDER_NUM").Value = dynSOWINVHC.Fields("SO_ORDER_NO").Value
    '            .Fields("INV_SALES").Value = -1 * CONT_PYMT_AMT_APPR
    '            .Fields("INV_DISC").Value = 0
    '            .Fields("INV_FREIGHT").Value = 0
    '            .Fields("INV_STAX").Value = 0
    '            .Fields("INV_TOTAL_AMOUNT").Value = -1 * CONT_PYMT_AMT_APPR
    '            .Fields("INV_BALANCE").Value = -1 * CONT_PYMT_AMT_APPR
    '            .Fields("CUST_CODE_SO").Value = dynSOWINVHC.Fields("CUST_CODE").Value
    '            .Fields("REASON_CODE").Value = dynARTPARM1.Fields("AR_PARM_REASON_CODE_OA").Value
    '            .Fields("INIT_OPER").Value = xUserID
    '            .Fields("INIT_DATE").Value = LAST_DATE
    '            '.Fields("SEG2_CODE").Value = dynSOTINVH1.Fields("ORDR_DIV_CODE").Value
    '            .Fields("SEG2_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG2").Value
    '            .Fields("SEG3_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG3").Value
    '            .Fields("SEG4_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG4").Value
    '            .Fields("CURR_CODE").Value = "USD"
    '            .Fields("CURR_EXCH_RATE").Value = "1"
    '            .Fields("INV_SALES_CURR").Value = -1 * CONT_PYMT_AMT_APPR
    '            .Fields("INV_DISC_CURR").Value = 0
    '            .Fields("INV_FREIGHT_CURR").Value = 0
    '            .Fields("INV_STAX_CURR").Value = 0
    '            .Fields("INV_TOTAL_AMOUNT_CURR").Value = -1 * CONT_PYMT_AMT_APPR
    '            .Fields("INV_BALANCE_CURR").Value = -1 * CONT_PYMT_AMT_APPR
    '            .Fields("SALES_DIVISION_CODE").Value = dynSOTINVH1.Fields("ORDR_DIV_CODE").Value
    '            .Update()
    '        End With

    '        Return

    '    End Sub

    '    Sub BRJ_Build_GL_Update(ByVal got_GL As Boolean)

    '        Dim RYP As String

    '        Call Track("Preparing G/L Interface", "")

    '        sql = "Select * from SPTPARM1 where SP_PARM_KEY = 'Z'"
    '        Dim dynSPTPARM1 As OraDynaset
    '        dynSPTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        Dim dyn As OraDynaset
    '        sql = "Select * from ARTPOST1 where POST_CODE = '" & dynSPTPARM1.Fields("SP_PARM_POST_CODE_AR").Value & "'"
    '        dyn = OraD.CreateDynaset(sql, 8&)
    '        Dim ACCT_CODE_AR As String
    '        ACCT_CODE_AR = dyn.Fields("POST_ACCT_RECV_ACCT").Value

    '        sql = "Select * from GLTPARM1 where GL_PARM_KEY = 'Z'"
    '        Dim dynGLTPARM1 As OraDynaset
    '        dynGLTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        sql = "Select * from SOTPARM1 where SO_PARM_KEY = 'Z'"
    '        Dim dynSOTPARM1 As OraDynaset
    '        dynSOTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        sql = "Select * from GLTINTF1 where ROWNUM < 1"
    '        Call Ora_to_Acc(Nothing, "GLWINTF1", 0, "", sql)
    '        Dim dynGLWINTF1 As Recordset
    '        dynGLWINTF1 = AccD.OpenRecordset("GLWINTF1", dbOpenDynaset)

    '        Dim JOURNAL_TYPE As String
    '        If Mid$(xFormName, 4, 3) = "BRK" Then
    '            JOURNAL_TYPE = "SPBP"
    '        Else
    '            JOURNAL_TYPE = "SPRP"
    '        End If

    '        Dim JOURNAL_NO As String
    '        Dim lno As Integer
    '        JOURNAL_NO = CTLNO("GLTJRNL1", 6)
    '        lno = 0

    '        Dim ACCT_CODE As String
    '        Dim SEG2_CODE As String
    '        Dim DETL_POSTING_AMT As Double

    '        Dim dynwk As Recordset

    '        ' For Remittance via AR, J/E to relieve accrual, CR AR, and expense Broker Expense variance
    '        ' For Remittance via AP, NO J/E here - entire entry to flow through Purchase Journal

    '        sql = "SELECT SOWINVHC.OPS_YYYYPP, SOWINVHC.ORDR_DIV_CODE"
    '        sql = sql & ", SUM (SOWINVHC.CONT_PYMT_AMT) AS CONT_PYMT_ACCR"
    '        sql = sql & ", SUM (SOWINVHC.CONT_PYMT_AMT_APPR) AS CONT_PYMT_AMT_APPR"
    '        sql = sql & ", SUM (SOWINVHC.CONT_PYMT_AMT - SOWINVHC.CONT_PYMT_AMT_APPR) AS CONT_PYMT_VAR"
    '        sql = sql & " FROM SOWINVHC"
    '        sql = sql & " where SOWINVHC.CONT_PYMT_STATUS = 'A'"
    '        sql = sql & "   and SOWINVHC.CONT_PYMT_VIA = 'R'"
    '        sql = sql & " group by SOWINVHC.OPS_YYYYPP, SOWINVHC.ORDR_DIV_CODE"
    '        dynwk = AccD.OpenRecordset(sql, dbOpenForwardOnly)

    '        Do While Not dynwk.EOF
    '            RYP = dynwk.Fields("OPS_YYYYPP").Value
    '            'RYP = CYP

    '            ' Reverse Accrual (original amount accrued)
    '            DETL_POSTING_AMT = Round(Val(dynwk.Fields("CONT_PYMT_ACCR").Value & ""), 2)
    '            If DETL_POSTING_AMT <> 0 Then
    '                If Mid$(xFormName, 4, 3) = "BRK" Then
    '                    ACCT_CODE = dynSOTPARM1.Fields("SO_PARM_ACCT_BRKR_ACCR").Value & ""
    '                Else
    '                    ACCT_CODE = dynSOTPARM1.Fields("SO_PARM_ACCT_REBATE_ACCR").Value & ""
    '                End If
    '                SEG2_CODE = dynGLTPARM1.Fields("GL_PARM_DEF_SEG2").Value & ""
    '            GoSub Write_INTF1
    '            End If

    '            ' Payment (amount approved for payment)
    '            DETL_POSTING_AMT = -1 * Round(Val(dynwk.Fields("CONT_PYMT_AMT_APPR").Value & ""), 2)
    '            If DETL_POSTING_AMT <> 0 Then
    '                ACCT_CODE = ACCT_CODE_AR
    '                'SEG2_CODE = dynwk.Fields("ORDR_DIV_CODE").Value
    '                SEG2_CODE = dynGLTPARM1.Fields("GL_PARM_DEF_SEG2").Value & ""
    '            GoSub Write_INTF1
    '            End If

    '            ' Expense (variance amount)
    '            DETL_POSTING_AMT = -1 * Round(Val(dynwk.Fields("CONT_PYMT_VAR").Value & ""), 2)
    '            If DETL_POSTING_AMT <> 0 Then
    '                ACCT_CODE = dynSOTPARM1.Fields("SO_PARM_ACCT_BRKR_EXP").Value & ""
    '                'SEG2_CODE = dynwk.Fields("ORDR_DIV_CODE").Value
    '                SEG2_CODE = dynGLTPARM1.Fields("GL_PARM_DEF_SEG2").Value & ""
    '            GoSub Write_INTF1
    '            End If

    '            dynwk.MoveNext()
    '        Loop

    '        ' Clean up
    '        dynwk.Close()
    '        dynGLWINTF1.Close()
    '        dynSPTPARM1.Close()
    '        dynGLTPARM1.Close()

    '        Exit Sub

    'Write_INTF1:
    '        got_GL = True

    '        dynGLWINTF1.AddNew()
    '        dynGLWINTF1.Fields("OPS_YYYYPP").Value = RYP
    '        dynGLWINTF1.Fields("JOURNAL_NO").Value = JOURNAL_NO
    '        lno = lno + 1
    '        dynGLWINTF1.Fields("JOURNAL_LNO").Value = lno

    '        dynGLWINTF1.Fields("ACCT_CODE").Value = ACCT_CODE

    '        dynGLWINTF1.Fields("SEG2_CODE").Value = SEG2_CODE
    '        dynGLWINTF1.Fields("SEG3_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG3").Value & ""
    '        dynGLWINTF1.Fields("SEG4_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG4").Value & ""
    '        dynGLWINTF1.Fields("DETL_CTL_DATE").Value = DateValue(Format$(Now + NowTSD, "mm/dd/yyyy"))
    '        dynGLWINTF1.Fields("DETL_EXE_NO").Value = xXNO
    '        dynGLWINTF1.Fields("DETL_POSTING_AMT").Value = Round(DETL_POSTING_AMT, 2)
    '        dynGLWINTF1.Fields("JOURNAL_TYPE").Value = JOURNAL_TYPE

    '        dynGLWINTF1.Fields("DETL_CTL_NO").Value = Null
    '        dynGLWINTF1.Fields("DETL_CTL_LNO").Value = 0
    '        dynGLWINTF1.Fields("DETL_DESC").Value = ""
    '        dynGLWINTF1.Fields("DETL_CVX_NO").Value = ""
    '        dynGLWINTF1.Fields("DETL_CVX_REF_DATE").Value = Null
    '        dynGLWINTF1.Fields("DETL_CVX_REF_NO").Value = Null
    '        dynGLWINTF1.Fields("DETL_CVX_REF_LNO").Value = 0
    '        dynGLWINTF1.Fields("DETL_CTL_TYPE").Value = ""
    '        dynGLWINTF1.Fields("DIST_CODE").Value = ""

    '        dynGLWINTF1.Update()

    '        Return

    '    End Sub

    '    Sub BRJ_Update(ByVal LAST_DATE As Date, ByVal INV_DATE As Date, ByVal got_GL As Boolean)
    '        Call Acc_to_Ora("APWINVH1", "", "")
    '        Call Acc_to_Ora("APWINVH2", "", "")
    '        Call Acc_to_Ora("ARWOPEN1", "", "")

    '        Dim dynSOTPARM1 As OraDynaset
    '        sql = "Select * from SOTPARM1 where SO_PARM_KEY = 'Z'"
    '        dynSOTPARM1 = OraD.CreateDynaset(sql, 8&)

    '        Dim ACCT_CODE_VAR As String
    '        If Mid$(xFormName, 4, 3) = "BRK" Then
    '            ACCT_CODE_VAR = dynSOTPARM1.Fields("SO_PARM_ACCT_BRKR_EXP").Value
    '        Else
    '            ACCT_CODE_VAR = dynSOTPARM1.Fields("SO_PARM_ACCT_REBATE_EXP").Value
    '        End If

    '        Dim dynSOWINVHC As Recordset
    '        sql = "Select * from SOWINVHC"
    '        sql = sql & " order by CONT_PYMT_NO"
    '        dynSOWINVHC = AccD.OpenRecordset(sql, dbOpenForwardOnly)

    '        Dim dynSOTINVHC As OraDynaset
    '        sql = "Select * from SOTINVHC where SO_ORDER_NO = :CODE"
    '        sql = sql & " and CONT_TP_TYPE = :CODE1"
    '        sql = sql & " and CONT_PYMT_VIA = :CODE2"
    '        sql = sql & " and CONT_PYMT_PAYEE_CODE = :CODE3"
    '        sql = sql & " and CONT_PYMT_SEQ = :NUM1"
    '        dynSOTINVHC = OraD.CreateDynaset(sql, 0&)

    '        Do While Not dynSOWINVHC.EOF
    '            With dynSOTINVHC
    '                OraD.Parameters("CODE").Value = dynSOWINVHC.Fields("SO_ORDER_NO").Value
    '                OraD.Parameters("CODE1").Value = dynSOWINVHC.Fields("CONT_TP_TYPE").Value
    '                OraD.Parameters("CODE2").Value = dynSOWINVHC.Fields("CONT_PYMT_VIA").Value
    '                OraD.Parameters("CODE3").Value = dynSOWINVHC.Fields("CONT_PYMT_PAYEE_CODE").Value
    '                OraD.Parameters("NUM1").Value = dynSOWINVHC.Fields("CONT_PYMT_SEQ").Value
    '                .Refresh()
    '                .Edit()
    '                If dynSOWINVHC.Fields("CONT_PYMT_STATUS").Value = "X" Then
    '                    .Fields("CONT_PYMT_STATUS").Value = "O"
    '                    .Fields("REGISTER_IND").Value = Null
    '                    .Fields("REGISTER_XNO").Value = Null
    '                    .Update()
    '                Else
    '                    ' .Fields("CONT_PYMT_STMT_XNO").ValueA
    '                    .Fields("CONT_PYMT_STATUS").Value = "P"
    '                    .Fields("OPS_YYYYPP").Value = CYP
    '                    .Fields("CONT_PYMT_STMT_XNO").Value = dynSOWINVHC.Fields("CONT_PYMT_STMT_XNO").Value
    '                    .Fields("CONT_PYMT_NO").Value = dynSOWINVHC.Fields("CONT_PYMT_NO").Value
    '                    .Update()

    '                    If dynSOWINVHC.Fields("CONT_PYMT_VIA").Value = "R" Then
    '                        Dim i As Integer
    '                        Dim z As String
    '                        Dim CONT_PYMT_AMT_APPR As Double
    '                        Dim CONT_PYMT_AMT_WOFF As Double
    '                        CONT_PYMT_AMT_APPR = Val(dynSOWINVHC.Fields("CONT_PYMT_AMT_APPR").Value & "")
    '                        CONT_PYMT_AMT_WOFF = Val(dynSOWINVHC.Fields("CONT_PYMT_AMT_WOFF").Value & "")
    '                        Dim ar() As String
    '                        ReDim ar(2, 1)
    '                        ar(0, 1) = "O"
    '                        ar(1, 1) = dynSOWINVHC.Fields("CONT_PYMT_NO").Value
    '                        ar(2, 1) = -1 * CONT_PYMT_AMT_WOFF
    '                        sql = "Select INV_TYPE, INV_NUM, CONT_PYMT_AMT_APPL from SOWINVHH "
    '                        sql = sql & " where SO_ORDER_NO = '" & dynSOWINVHC.Fields("SO_ORDER_NO").Value & "'"
    '                        sql = sql & "   and CONT_TP_TYPE = '" & dynSOWINVHC.Fields("CONT_TP_TYPE").Value & "'"
    '                        sql = sql & "   and CONT_PYMT_VIA = '" & dynSOWINVHC.Fields("CONT_PYMT_VIA").Value & "'"
    '                        sql = sql & "   and CONT_PYMT_PAYEE_CODE = '" & dynSOWINVHC.Fields("CONT_PYMT_PAYEE_CODE").Value & "'"
    '                        sql = sql & "   and CONT_PYMT_SEQ = " & dynSOWINVHC.Fields("CONT_PYMT_SEQ").Value
    '                        Dim dynwk As Recordset
    '                        dynwk = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    '                        i = 1
    '                        Do While Not dynwk.EOF
    '                            i = i + 1
    '                            ReDim Preserve ar(2, i)
    '                            ar(0, i) = dynwk.Fields("INV_TYPE").Value
    '                            ar(1, i) = dynwk.Fields("INV_NUM").Value
    '                            ar(2, i) = Val(dynwk.Fields("CONT_PYMT_AMT_APPL").Value & "")
    '                            ar(2, 1) = Val(ar(2, 1)) - Val(ar(2, i))
    '                            dynwk.MoveNext()
    '                        Loop
    '                        dynwk.Close()
    '                        z = ZERO_CHECK(INV_DATE, "", LAST_DATE, ar(), dynSOWINVHC.Fields("CONT_PYMT_PAYEE_CODE").Value, "", ACCT_CODE_VAR, dynSOWINVHC.Fields("ORDR_DIV_CODE").Value, "", "", False)
    '                    End If
    '                End If
    '            End With
    '            dynSOWINVHC.MoveNext()
    '        Loop
    '        dynSOWINVHC.Close()

    '        If got_GL Then
    '            Call GL_Update()
    '        End If
    '    End Sub
End Class
