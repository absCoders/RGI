Imports System.Math
Public Class ICCMAIN1

    Public Shared Sub EMP_UPC_GTIN( _
    ByRef row As DataRow, _
    ByVal fix_row As Boolean, _
    ByRef UPC_CODE As String, _
    ByRef GTIN As String)

        Dim SW(8) As String
        Dim C(8, 1) As String

        Dim PROD_CODE As String = row.Item("PROD_CODE")
        Dim BRAND_CODE As String = row.Item("BRAND_CODE")
        Dim SIZE_CODE As String = row.Item("SIZE_CODE")
        Dim ORIG_CODE As String = row.Item("ORIG_CODE")
        Dim PACK_CODE As String = row.Item("PACK_CODE")
        Dim SP_GROUP As String = row.Item("SP_GROUP")
        Dim GRADE_CODE As String = row.Item("GRADE_CODE")

        C(2, 0) = "SIZE_CODE"
        C(2, 1) = SIZE_CODE
        C(3, 0) = "ORIG_CODE"
        C(3, 1) = ORIG_CODE
        C(5, 0) = "PACK_CODE"
        C(5, 1) = PACK_CODE
        C(6, 0) = "SP_GROUP"
        C(6, 1) = SP_GROUP
        C(7, 0) = "GRADE_CODE"
        C(7, 1) = GRADE_CODE

        For i As Integer = 2 To 7

            If (PROD_CODE = "" Or BRAND_CODE = "") _
            Or (C(i, 1) = "" And i <> 4) Then
                UPC_CODE = ""
                GTIN = ""
                If fix_row Then
                    row.Item("UPC_CODE") = UPC_CODE
                    row.Item("GTIN") = GTIN
                End If
                Exit Sub
            End If

        Next i

        ASCMAIN1.sql = "SELECT * from ICTUPCD4" _
        & " where PROD_CODE = :PARM1" _
        & "   and BRAND_CODE = :PARM2"
        Dim rowICTUPCDX As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", _
                                    New String() {PROD_CODE, BRAND_CODE})

        If rowICTUPCDX Is Nothing Then
            BRAND_CODE = "NOBRAND"
        End If

        ASCMAIN1.sql = "SELECT * from ICTUPCD1" _
        & " where PROD_CODE = :PARM1" _
        & " and BRAND_CODE = :PARM2"
        Dim rowICTUPCD1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", _
                                    New String() {PROD_CODE, BRAND_CODE})

        If rowICTUPCD1 Is Nothing Then
            UPC_CODE = ""
            GTIN = ""
            If fix_row Then
                row.Item("UPC_CODE") = UPC_CODE
                row.Item("GTIN") = GTIN
            End If
            Exit Sub
        End If

        If rowICTUPCD1.Item("UPC_BY_SIZE") & "" = "" Then
            ASCMAIN1.sql = "SELECT * from ICTUPCD1" _
            & " where UPC_BY_SIZE IS NOT NULL" _
            & "   and PROD_CTL_NO = '" & rowICTUPCD1.Item("PROD_CTL_NO") & "'"
            rowICTUPCD1 = ASCDATA1.GetDataRow

            PROD_CODE = rowICTUPCD1.Item("PROD_CODE")
            BRAND_CODE = rowICTUPCD1.Item("BRAND_CODE")

        End If

        For i = 2 To 7
            If i <> 4 Then
                Select Case rowICTUPCD1.Item(i)
                    Case Is = "S"
                        SW(i) = " AND " & C(i, 0) & " = '" & C(i, 1) & "'"
                    Case Is = "G"
                        ASCMAIN1.sql = "SELECT UPC_GROUP" _
                        & " FROM ICTUPCD" & Trim(Str(i)) _
                        & " WHERE PROD_CODE = '" & PROD_CODE & "'" _
                        & " AND BRAND_CODE = '" & BRAND_CODE & "'" _
                        & " AND " & C(i, 0) & " = '" & C(i, 1) & "'"
                        rowICTUPCDX = ASCDATA1.GetDataRow
                        SW(i) = " and UPC_GROUP_" & C(i, 0) & " = '" & rowICTUPCDX.Item("UPC_GROUP") & "'"
                    Case Is = "N"
                        SW(i) = ""
                End Select
            End If
        Next i

        ASCMAIN1.sql = "SELECT UPC_CODE, GTIN from ICTUPCD0" _
        & " where PROD_CODE = '" & PROD_CODE & "'" _
        & "   and BRAND_CODE = '" & BRAND_CODE & "'"
        For i = 0 To 8
            If SW(i) <> "" Then
                ASCMAIN1.sql &= SW(i)
            End If
        Next i
        Dim rowICTUPCD0 As DataRow = ASCDATA1.GetDataRow

        If rowICTUPCD0 Is Nothing Then
            UPC_CODE = ""
            GTIN = ""
        Else
            UPC_CODE = rowICTUPCD0.Item("UPC_CODE")
            GTIN = rowICTUPCD0.Item("GTIN")
        End If

        If fix_row Then
            row.Item("UPC_CODE") = UPC_CODE
            row.Item("GTIN") = GTIN
        End If


    End Sub

#Region "VB6"

    '    Sub Generate_C_Xfr_Adj(TRAN_NO As String)
    '        ' warning: this routine exists also in ICBMAINX
    '        Dim dynICTPARM1 As Object
    '        SQL = "Select * from ICTPARM1 where IC_PARM_KEY = 'Z'"
    '        dynICTPARM1 = OraD.CreateDynaset(SQL, 8&)

    '        Dim dyn1x As Recordset
    '        Dim dyn2x As Recordset
    '        Dim dynICWTRAN1 As Recordset
    '        Dim dynICWTRAN2 As Recordset

    '        Dim rxno As String
    '        Dim i As Integer
    '        Dim j As Integer

    '        SQL = "Select * from ICWTRAN1 where OPS_YYYYPP = '" & CYP & "' and TRAN_TYPE = 'C' and TRAN_NO = '" & TRAN_NO & "'"
    '        dyn1x = AccD.OpenRecordset(SQL, dbOpenForwardOnly)
    '        For i = 1 To 2
    '            dynICWTRAN2 = AccD.OpenRecordset("ICWTRAN2", dbOpenDynaset)
    '            dynICWTRAN1 = AccD.OpenRecordset("ICWTRAN1", dbOpenDynaset)
    '            rxno = ""
    '            SQL = "Select * from ICWTRAN2 where OPS_YYYYPP = '" & CYP & "' and TRAN_TYPE = 'C' and TRAN_NO = '" & TRAN_NO & "'"
    '            If i = 1 Then
    '                SQL = SQL & " and TRAN_QTY1 <> 0"
    '            Else
    '                SQL = SQL & " and TRAN_QTY3 <> 0"
    '            End If
    '            dyn2x = AccD.OpenRecordset(SQL, dbOpenForwardOnly)
    '            Do While Not dyn2x.EOF
    '                If rxno = "" Then
    '                    rxno = CTLNO("REC_NO_" & Mid$("TA", i, 1), 6)
    '                End If
    '                dynICWTRAN2.AddNew()
    '                For j = 0 To dyn2x.Fields.Count - 1
    '                    dynICWTRAN2.Fields(j).Value = dyn2x.Fields(j).Value
    '                Next j
    '                dynICWTRAN2.Fields("TRAN_TYPE").Value = Mid$("TA", i, 1)
    '                dynICWTRAN2.Fields("TRAN_NO").Value = rxno
    '                If i = 1 Then
    '                    dynICWTRAN2.Fields("TRAN_QTY").Value = dyn2x.Fields("TRAN_QTY1").Value
    '                Else
    '                    dynICWTRAN2.Fields("TRAN_QTY").Value = -1 * dyn2x.Fields("TRAN_QTY3").Value
    '                End If
    '                dynICWTRAN2.Fields("TRAN_PRICE").Value = 0
    '                '            dynICWTRAN2.Fields("TRAN_REF").Value = Null
    '                '            dynICWTRAN2.Fields("TRAN_HOLD").Value = Null
    '                dynICWTRAN2.Fields("TRAN_QTY1").Value = 0
    '                dynICWTRAN2.Fields("TRAN_QTY2").Value = 0
    '                dynICWTRAN2.Fields("TRAN_QTY3").Value = 0
    '                dynICWTRAN2.Update()
    '                dyn2x.MoveNext()
    '            Loop
    '            dyn2x.Close()
    '            If rxno <> "" Then
    '                dynICWTRAN1.AddNew()
    '                For j = 0 To dyn1x.Fields.Count - 1
    '                    dynICWTRAN1.Fields(j).Value = dyn1x.Fields(j).Value
    '                Next j
    '                dynICWTRAN1.Fields("TRAN_TYPE").Value = Mid$("TA", i, 1)
    '                dynICWTRAN1.Fields("TRAN_NO").Value = rxno
    '                dynICWTRAN1.Fields("TRAN_CUST_CODE").Value = Null
    '                dynICWTRAN1.Fields("CUST_STORE_NO").Value = Null
    '                dynICWTRAN1.Fields("TRAN_CCVRW_REF").Value = Null
    '                If i = 1 Then
    '                    dynICWTRAN1.Fields("TRAN_WHSE_CODE_TO").Value = dynICTPARM1.Fields("IC_PARM_RTN_XFR_WHSE").Value
    '                    dynICWTRAN1.Fields("TRAN_CCVAT_DESC").Value = "Returns Transfer to Stock"
    '                Else
    '                    dynICWTRAN1.Fields("TRAN_ADJ_REASON_CODE").Value = dynICTPARM1.Fields("IC_PARM_RTN_ADJ_REASON").Value
    '                    dynICWTRAN1.Fields("TRAN_CCVRW_DESC").Value = "Returns Adj for Destroyed Mdse"
    '                End If
    '                dynICWTRAN1.Fields("TRAN_ORIGINATE").Value = "C"
    '                dynICWTRAN1.Fields("TRAN_FREIGHT").Value = 0
    '                dynICWTRAN1.Fields("TRAN_MISC_CHG").Value = 0
    '                '            dynICWTRAN1.Fields("TRAN_STAX").Value = 0
    '                dynICWTRAN1.Fields("TRAN_TYPE_ORIG").Value = "C"
    '                dynICWTRAN1.Fields("TRAN_NO_ORIG").Value = TRAN_NO
    '                dynICWTRAN1.Update()
    '            End If
    '            dynICWTRAN1.Close()
    '            dynICWTRAN2.Close()
    '            If rxno <> "" Then
    '                Call ICBMAIN1.TRAN1_Update(Mid$("TA", i, 1), rxno, "1", "1")
    '            End If
    '        Next i

    '    End Sub

    'Sub TRAN1_Update(TRAN_TYPE As String, TRAN_NO As String, Optional quiet As Variant, Optional no_trans As Variant, Optional FP As Integer)

    '        Dim dynICTSTAT1 As OraDynaset
    '        Dim dynICTSTAT2 As OraDynaset
    '        Dim dynICTWHSE1 As OraDynaset
    '        Dim dynWHTLOCB1 As OraDynaset
    '        Dim dynWHTLOCB2 As OraDynaset
    '        Dim dynWHTPARM1 As OraDynaset
    '        Dim dynICTSTYL1 As OraDynaset
    '        Dim dynICTPARM1 As OraDynaset
    '        Dim dynARTCUST1 As OraDynaset
    '        Dim dynARTCUST6 As OraDynaset
    '        Dim dynARTOPEN1 As OraDynaset
    '        Dim dynSOTINVH1 As OraDynaset
    '        Dim dynSOTINVH2 As OraDynaset

    '        Dim STYLE_CODE As String
    '        Dim COLOR_CODE As String
    '        Dim SALES_DIVISION_CODE As String
    '        Dim CUST_CODE As String
    '        Dim CUST_STORE_NO As String
    '        Dim TRAN_PRICE As Double
    '        Dim TRAN_COST As Double
    '        Dim TRAN_QTY As Long
    '        Dim TRAN_LNO As Long

    '        Dim CUST_BILL_TO_CUST As String
    '        Dim TRAN_CCVRW_REF As String

    '        Dim INV_NO As String
    '        Dim INV_TYPE As String

    '        Dim INV_TOTAL_AMOUNT As Double
    '        Dim INV_SALES As Double
    '        Dim INV_COGS As Double
    '        Dim CUST_BALANCE As Double
    '        Dim INV_DATE As String
    '        Dim dysdue As Integer
    '        Dim dysdsc As Integer
    '        Dim INV_DUE_DATE As String
    '        Dim INV_DISC_DATE As String
    '        Dim POST_CODE As String
    '        Dim REASON_CODE As String
    '        Dim termt As String

    '        Dim srev As Integer
    '        Dim S As Integer
    '        Dim ss As Integer
    '        Dim statsign As Integer
    '        Dim sstatsign As Integer

    '        Dim WHSE_CODE As String
    '        Dim ICTSTAT1field As String

    '        Dim i As Integer
    '        Dim j As Integer
    '        Dim Period_Add As Integer
    '        Period_Add = Val(FP)
    '        Dim single_record           ' True if TRAN_NO is passed in

    '        Dim PrepTranType As String  ' List of All TRAN_TYPE's processed, so that Initialization Routines for each type of Transaction may be processed only once
    '        PrepTranType = ""

    '        Dim w As String
    '        Dim a() As String

    '        Dim dynICTSTYC1_FIFO As OraDynaset
    '        SQL = "Select * from ICTSTYC1 where STYLE_CODE = :STYLE_CODE and COLOR_CODE = :COLOR_CODE"
    '        dynICTSTYC1_FIFO = OraD.CreateDynaset(SQL, 8&)

    '        Dim IC_PARM_REASON_SHP As String
    '        Dim IC_PARM_REASON_RTN As String
    '        SQL = "Select * from ICTPARM1 where IC_PARM_KEY = 'Z'"
    '        dynICTPARM1 = OraD.CreateDynaset(SQL, 8&)
    '        IC_PARM_REASON_SHP = dynICTPARM1.Fields("IC_PARM_REASON_SHP").Value & ""
    '        IC_PARM_REASON_RTN = dynICTPARM1.Fields("IC_PARM_REASON_RTN").Value & ""

    '        Dim LOCATION_CODE As String
    '        SQL = "SELECT *"
    '        SQL = SQL & " From WHTPARM1"
    '        SQL = SQL & " WHERE WH_PARM_KEY = 'Z'"
    '        dynWHTPARM1 = OraD.CreateDynaset(SQL, 8&)
    '        Select Case TRAN_TYPE
    '            Case Is = "C"
    '                LOCATION_CODE = dynWHTPARM1.Fields("WH_PARM_LOC_RET").Value
    '            Case Is = "M"
    '                LOCATION_CODE = dynWHTPARM1.Fields("WH_PARM_LOC_SAM").Value
    '            Case Else
    '                LOCATION_CODE = dynWHTPARM1.Fields("WH_PARM_LOC_LNF").Value
    '        End Select
    '        dynWHTPARM1.Close()

    '        Dim WHSE_ACTIVE As Boolean
    '        SQL = "SELECT WHSE_LOCATOR"
    '        SQL = SQL & " From ICTWHSE1"
    '        SQL = SQL & " WHERE WHSE_CODE = :WHSE_CODE"
    '        dynICTWHSE1 = OraD.CreateDynaset(SQL, 8&)

    '        SQL = "Select * from ICTSTAT1 "
    '        SQL = SQL & " where OPS_YYYYPP = :OPS_YYYYPP "
    '        SQL = SQL & "   and STYLE_CODE = :STYLE_CODE "
    '        SQL = SQL & "   and COLOR_CODE = :COLOR_CODE "
    '        SQL = SQL & "   and WHSE_CODE = :WHSE_CODE"
    '        dynICTSTAT1 = OraD.CreateDynaset(SQL, 8&)
    '        SQL = "Select * from ICTSTAT2 "
    '        SQL = SQL & " where STYLE_CODE = :STYLE_CODE "
    '        SQL = SQL & "   and COLOR_CODE = :COLOR_CODE "
    '        SQL = SQL & "   and WHSE_CODE = :WHSE_CODE"
    '        dynICTSTAT2 = OraD.CreateDynaset(SQL, 8&)
    '        SQL = "SELECT * FROM WHTLOCB1"
    '        SQL = SQL & " WHERE WHSE_CODE = :WHSE_CODE"
    '        SQL = SQL & " AND LOCATION_CODE = :LOCATION_CODE"
    '        SQL = SQL & " AND BAR_CODE = '0000000000'"
    '        SQL = SQL & " AND STYLE_CODE = :STYLE_CODE"
    '        SQL = SQL & " AND COLOR_CODE = :COLOR_CODE"
    '        dynWHTLOCB1 = OraD.CreateDynaset(SQL, 0&)
    '        SQL = "SELECT * FROM WHTLOCB2"
    '        SQL = SQL & " WHERE WHSE_CODE = :WHSE_CODE"
    '        SQL = SQL & " AND LOCATION_CODE = :LOCATION_CODE"
    '        SQL = SQL & " AND BAR_CODE = '0000000000'"
    '        SQL = SQL & " AND STYLE_CODE = :STYLE_CODE"
    '        SQL = SQL & " AND COLOR_CODE = :COLOR_CODE"
    '        dynWHTLOCB2 = OraD.CreateDynaset(SQL, 0&)

    '        SQL = "Select * from ICTSTYL1 where STYLE_CODE = :STYLE_CODE"
    '        dynICTSTYL1 = OraD.CreateDynaset(SQL, 8&)

    '        '    SQL = " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf
    '        '    SQL = SQL & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST," & vbCrLf
    '        '    SQL = SQL & " 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf
    '        '    SQL = SQL & " FROM ICTCOST1" & vbCrLf
    '        '    SQL = SQL & " WHERE ROWNUM < 0" & vbCrLf
    '        '    Call Ora_to_Acc(Nothing, "ICWCOST1", 0, "", SQL)
    '        SQL = "SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf
    '    SQL = SQL & " ICTCOST1.TRAN_TYPE, '" & String(50, " ") & "' TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST, 0.00 QTY_USED" & vbCrLf
    '        SQL = SQL & " FROM ICTCOST1" & vbCrLf
    '        SQL = SQL & " WHERE ROWNUM < 0"
    '        Call Ora_to_Acc(Nothing, "ICWCOST1", 0, "", SQL)

    '        If no_trans = "" Then
    '            OraS.BeginTrans()
    '        End If

    '        If TRAN_NO = "" Then
    '            single_record = False
    '        Else
    '            single_record = True
    '        End If

    '        If Not single_record Then
    '            Call Acc_to_Ora("ICWTRAN1", "")
    '            Call Acc_to_Ora("ICWTRAN2", "")
    '        End If

    '        'SQL = "Select * from ICWTRAN1 where OPS_YYYYPP = '" & CYP & "'"
    '        SQL = "Select * from ICWTRAN1 where OPS_YYYYPP = '" & Period_Calc(CYP, Period_Add) & "'"
    '        If single_record Then
    '            SQL = SQL & " and TRAN_NO = '" & TRAN_NO & "'"
    '            SQL = SQL & " and TRAN_TYPE = '" & TRAN_TYPE & "'"
    '        End If
    '        Dim dynICWTRAN1 As Recordset
    '        dynICWTRAN1 = AccD.OpenRecordset(SQL, dbOpenForwardOnly)
    '        Do While Not dynICWTRAN1.EOF
    '            TRAN_TYPE = dynICWTRAN1.Fields("TRAN_TYPE").Value
    '            If InStr(PrepTranType, TRAN_TYPE) = 0 Then
    '            GoSub Prep_TranType
    '            End If
    '            TRAN_NO = dynICWTRAN1.Fields("TRAN_NO").Value
    '            REASON_CODE = dynICWTRAN1.Fields("REASON_CODE").Value & ""
    '            TRAN_CCVRW_REF = dynICWTRAN1.Fields("TRAN_CCVRW_REF").Value & ""
    '            SALES_DIVISION_CODE = ""

    '            Call Prompt("-", TRAN_NO)
    '            Select Case TRAN_TYPE
    '                Case "A", "M"
    '                    ICTSTAT1field = "WHSE_QTY_ADJ"
    '                    statsign = 1
    '                    S = 1
    '                Case "C"
    '                    ICTSTAT1field = "WHSE_QTY_RTN"
    '                    statsign = 1
    '                    S = 1
    '                Case "R"
    '                    ICTSTAT1field = "WHSE_QTY_REC"
    '                    statsign = 1
    '                    S = 1
    '                Case "S"
    '                    ICTSTAT1field = "WHSE_QTY_SHP"
    '                    statsign = 1
    '                    S = -1
    '                Case "T"
    '                    ICTSTAT1field = "WHSE_QTY_XFR"
    '                    statsign = -1
    '                    S = -1
    '            End Select

    '            If TRAN_TYPE = "S" Or TRAN_TYPE = "C" Then
    '                INV_NO = "0000" & TRAN_NO
    '                CUST_CODE = dynICWTRAN1.Fields("TRAN_CUST_CODE").Value
    '                OraD.Parameters("CUST_CODE").Value = CUST_CODE
    '                CUST_STORE_NO = dynICWTRAN1.Fields("CUST_STORE_NO").Value & ""
    '                'OraD.Parameters("CUST_STORE_NO").Value = CUST_STORE_NO
    '                INV_SALES = 0
    '                INV_COGS = 0
    '            End If

    '            Dim sqlx As String
    '            'sqlx = " OPS_YYYYPP = '" & CYP & "'"
    '            sqlx = " OPS_YYYYPP = '" & Period_Calc(CYP, Period_Add) & "'"
    '            sqlx = sqlx & " and TRAN_TYPE = '" & TRAN_TYPE & "'"
    '            sqlx = sqlx & " and TRAN_NO = '" & TRAN_NO & "'"
    '            If dynICWTRAN1.Fields("TRAN_STATUS_UPD").Value = "R" Then
    '                srev = -1
    '                statsign = -statsign
    '                S = -S
    '                SQL = "Update ICTTRAN1 set TRAN_STATUS_UPD = 'R' where " & sqlx
    '                OraD.ExecuteSQL(SQL)
    '            Else
    '                srev = 1
    '                If single_record Then
    '                    Call Acc_to_Ora("ICWTRAN1", " where " & sqlx)
    '                    Call Acc_to_Ora("ICWTRAN2", " where " & sqlx)
    '                End If
    '            End If

    '            OraD.Parameters("OPS_YYYYPP").Value = Period_Calc(CYP, Period_Add)

    '            Dim dynICWTRAN2 As Recordset
    '            SQL = "Select * from ICWTRAN2 where " & sqlx
    '            dynICWTRAN2 = AccD.OpenRecordset(SQL, dbOpenForwardOnly)
    '            Do While Not dynICWTRAN2.EOF
    '                WHSE_CODE = dynICWTRAN1.Fields("TRAN_WHSE_CODE").Value
    '                TRAN_QTY = Val(dynICWTRAN2.Fields("TRAN_QTY").Value & "")

    '                STYLE_CODE = dynICWTRAN2.Fields("STYLE_CODE").Value
    '                COLOR_CODE = dynICWTRAN2.Fields("COLOR_CODE").Value
    '                OraD.Parameters("STYLE_CODE").Value = STYLE_CODE
    '                OraD.Parameters("COLOR_CODE").Value = COLOR_CODE

    '                dynICTSTYL1.Refresh()
    '                dynICTSTYC1_FIFO.Refresh()

    '                If SALES_DIVISION_CODE = "" Then
    '                    SALES_DIVISION_CODE = dynICTSTYL1.Fields("SALES_DIVISION_CODE").Value & ""
    '                End If

    '                '            If Val(dynICTSTYC1_FIFO.Fields("STYLE_COST_FIFO").Value & "") <> 0 Then
    '                '                TRAN_COST = Val(dynICTSTYC1_FIFO.Fields("STYLE_COST_FIFO").Value & "")
    '                '            Else
    '                '                TRAN_COST = Val(dynICTSTYL1.Fields("STYLE_COST").Value & "")
    '                '            End If

    '                '            If UserID = "wayne" Then
    '                '                Stop
    '                '                'Wayne needs to test this following costing.
    '                '            End If

    '                ReDim a(1)
    '                w = ""
    '                'w = Calc_Cost(STYLE_CODE, COLOR_CODE, False, 0, False, "SOTINVH1", "SOTINVH2", False)
    '                w = Calc_Cost_New(Period_Calc(CYP, Period_Add), STYLE_CODE, COLOR_CODE, False)
    '                a() = Split(w, "|")

    '                TRAN_COST = Val(a(0))

    '                If TRAN_TYPE = "S" Or TRAN_TYPE = "C" Then
    '                    TRAN_PRICE = dynICWTRAN2.Fields("TRAN_PRICE").Value
    '                    INV_SALES = INV_SALES - S * TRAN_QTY * TRAN_PRICE
    '                    INV_COGS = INV_COGS - S * TRAN_QTY * TRAN_COST
    '                GoSub Update_INVH2
    '                End If

    '                ss = S
    '                sstatsign = statsign
    '            GoSub Update_Stat

    '                If TRAN_TYPE = "T" Then
    '                    WHSE_CODE = dynICWTRAN1.Fields("TRAN_WHSE_CODE_TO").Value
    '                    ss = -S
    '                    sstatsign = -statsign
    '                GoSub Update_Stat
    '                End If

    '                dynICWTRAN2.MoveNext()
    '            Loop
    '            dynICWTRAN2.Close()

    '            If TRAN_TYPE = "S" Or TRAN_TYPE = "C" Then
    '            GoSub Update_CUST
    '            GoSub Update_INVH1
    '            End If

    '            dynICWTRAN1.MoveNext()
    '        Loop
    '        dynICWTRAN1.Close()

    '        If no_trans = "" Then
    '            OraS.CommitTrans()
    '        End If

    '        If quiet = "" Then
    '            MsgBox("Record Successfully Updated", 48 + 0, "Success")
    '        End If
    '        Exit Sub

    'Update_INVH1:
    '        If srev = -1 Then
    '            If TRAN_TYPE = "S" Then
    '                OraD.Parameters("INV_TYPE").Value = "I"
    '            Else
    '                OraD.Parameters("INV_TYPE").Value = "C"
    '            End If
    '            OraD.Parameters("INV_NO").Value = INV_NO
    '            dynSOTINVH1.Refresh()
    '            dynSOTINVH1.Delete()
    '        Else
    '            dynSOTINVH1.AddNew()
    '            If TRAN_TYPE = "S" Then
    '                dynSOTINVH1.Fields("INV_TYPE").Value = "I"
    '                dynSOTINVH1.Fields("REASON_CODE").Value = IC_PARM_REASON_SHP
    '            Else
    '                dynSOTINVH1.Fields("INV_TYPE").Value = "C"
    '                dynSOTINVH1.Fields("REASON_CODE").Value = IC_PARM_REASON_RTN
    '            End If
    '            dynSOTINVH1.Fields("INV_NO").Value = INV_NO
    '            dynSOTINVH1.Fields("CUST_CODE").Value = CUST_CODE
    '            dynSOTINVH1.Fields("CUST_STORE_NO").Value = CUST_STORE_NO
    '            dynSOTINVH1.Fields("ORDR_CUST_PO").Value = TRAN_CCVRW_REF
    '            dynSOTINVH1.Fields("WHSE_CODE").Value = WHSE_CODE
    '            dynSOTINVH1.Fields("REASON_CODE").Value = REASON_CODE
    '            dynSOTINVH1.Fields("INV_SALES").Value = INV_SALES
    '            dynSOTINVH1.Fields("INV_COGS").Value = INV_COGS
    '            dynSOTINVH1.Fields("INV_FREIGHT").Value = -1 * S * Val(dynICWTRAN1.Fields("TRAN_FREIGHT").Value & "")
    '            dynSOTINVH1.Fields("INV_MISC_CHG").Value = -1 * S * Val(dynICWTRAN1.Fields("TRAN_MISC_CHG").Value & "")
    '            dynSOTINVH1.Fields("INV_TOTAL_AMOUNT").Value = INV_TOTAL_AMOUNT
    '            dynSOTINVH1.Fields("INV_DATE").Value = INV_DATE
    '            dynSOTINVH1.Fields("ORDR_DATE_UPDATED").Value = Now + NowTSD
    '            dynSOTINVH1.Fields("ORDR_YYYYPP_UPDATED").Value = CYP
    '            dynSOTINVH1.Fields("ORDR_BILL_TO_CUST").Value = CUST_BILL_TO_CUST
    '            dynSOTINVH1.Fields("POST_CODE").Value = POST_CODE
    '            dynSOTINVH1.Fields("TERM_CODE").Value = "00"
    '            dynSOTINVH1.Fields("SREP_CODE").Value = dynICWTRAN1.Fields("SREP_CODE").Value & ""
    '            dynSOTINVH1.Fields("SHIP_BOL_NO").Value = Null
    '            'dynSOTINVH1.Fields("EDI").Value = Null
    '            dynSOTINVH1.Fields("SALES_DIVISION_CODE").Value = SALES_DIVISION_CODE
    '            dynSOTINVH1.Fields("INV_COMMENT").Value = dynICWTRAN1.Fields("TRAN_COMMENT").Value
    '            'CURRENCY
    '            dynSOTINVH1.Fields("CURR_CODE").Value = "USD"
    '            dynSOTINVH1.Fields("CURR_EXCH_RATE").Value = 1
    '            dynSOTINVH1.Fields("INV_SALES_CURR").Value = INV_SALES
    '            dynSOTINVH1.Fields("INV_FREIGHT_CURR").Value = -1 * S * Val(dynICWTRAN1.Fields("TRAN_FREIGHT").Value & "")
    '            dynSOTINVH1.Fields("INV_MISC_CHG_CURR").Value = -1 * S * Val(dynICWTRAN1.Fields("TRAN_MISC_CHG").Value & "")
    '            dynSOTINVH1.Fields("INV_TOTAL_AMT_CURR").Value = INV_TOTAL_AMOUNT
    '            dynSOTINVH1.Fields("GST_TAX").Value = 0
    '            dynSOTINVH1.Fields("GST_TAX_CURR").Value = 0
    '            dynSOTINVH1.Update()
    '        End If
    '        Return

    'Update_INVH2:
    '        If srev = -1 Then
    '            If TRAN_TYPE = "S" Then
    '                OraD.Parameters("INV_TYPE").Value = "I"
    '            Else
    '                OraD.Parameters("INV_TYPE").Value = "C"
    '            End If
    '            OraD.Parameters("INV_NO").Value = INV_NO
    '            dynSOTINVH2.Refresh()
    '            Do While Not dynSOTINVH2.EOF
    '                dynSOTINVH2.Delete()
    '                dynSOTINVH2.MoveNext()
    '            Loop
    '        Else
    '            dynSOTINVH2.AddNew()
    '            If TRAN_TYPE = "S" Then
    '                dynSOTINVH2.Fields("INV_TYPE").Value = "I"
    '            Else
    '                dynSOTINVH2.Fields("INV_TYPE").Value = "C"
    '            End If
    '            dynSOTINVH2.Fields("INV_NO").Value = INV_NO
    '            dynSOTINVH2.Fields("INV_LNO").Value = dynICWTRAN2.Fields("TRAN_LNO").Value
    '            dynSOTINVH2.Fields("STYLE_CODE").Value = STYLE_CODE
    '            dynSOTINVH2.Fields("COLOR_CODE").Value = COLOR_CODE
    '            dynSOTINVH2.Fields("ORDR_UNIT_COST").Value = TRAN_COST
    '            dynSOTINVH2.Fields("ORDR_UNIT_PRICE").Value = TRAN_PRICE
    '            dynSOTINVH2.Fields("ORDR_QTY_SHIP").Value = -1 * S * TRAN_QTY
    '            dynSOTINVH2.Fields("ORDR_YYYYPP_UPDATED").Value = CYP
    '            dynSOTINVH2.Fields("CUST_CODE").Value = CUST_CODE
    '            dynSOTINVH2.Fields("ORDR_UNIT_PRICE_CURR").Value = TRAN_PRICE
    '            '        dynSOTINVH2.Fields("ORDR_CUST_PO").Value = dynICWTRAN1.Fields("TRAN_CCVAT_REF").Value
    '            '        dynSOTINVH2.Fields("ORDR_SHIP_DATE").Value = invdate
    '            '        dynSOTINVH2.Fields("WHSE_CODE").Value = WHSE_CODE
    '            dynSOTINVH2.Update()
    '        End If
    '        Return

    'Update_CUST:
    '        If TRAN_TYPE = "S" Then
    '            INV_TYPE = "I"
    '        Else
    '            INV_TYPE = "R"
    '        End If

    '        INV_TOTAL_AMOUNT = INV_SALES _
    '                         - S * Val(dynICWTRAN1.Fields("TRAN_FREIGHT").Value & "") _
    '                         - S * Val(dynICWTRAN1.Fields("TRAN_MISC_CHG").Value & "")
    '        '                     + Val(dynICWTRAN1.Fields("TRAN_STAX").Value & "")
    '        INV_DATE = dynICWTRAN1.Fields("TRAN_DATE").Value

    '        dynARTCUST1.Refresh()
    '        If Not IsNull(dynARTCUST1.Fields("CUST_BILL_TO_CUST").Value) Then
    '            CUST_BILL_TO_CUST = dynARTCUST1.Fields("CUST_BILL_TO_CUST").Value
    '        Else
    '            CUST_BILL_TO_CUST = CUST_CODE
    '        End If
    '        If CUST_BILL_TO_CUST <> CUST_CODE Then
    '        GoSub Update_CUST6
    '            OraD.Parameters("CUST_CODE").Value = CUST_BILL_TO_CUST
    '            dynARTCUST1.Refresh()
    '        End If
    '        'dynARTCUST1.Edit

    '        POST_CODE = dynARTCUST1.Fields("POST_CODE").Value & ""
    '        dysdue = Val(SRead(termt, dynARTCUST1.Fields("TERM_CODE").Value & "", 2))
    '        dysdsc = Val(SRead(termt, dynARTCUST1.Fields("TERM_CODE").Value & "", 3))

    '        If TRAN_TYPE = "C" Then
    '            INV_DUE_DATE = INV_DATE
    '            INV_DISC_DATE = INV_DATE
    '        Else
    '            INV_DUE_DATE = DateAdd("d", dysdue, INV_DATE)
    '            If dysdsc = 0 Then
    '                INV_DISC_DATE = ""
    '            Else
    '                INV_DISC_DATE = DateAdd("d", dysdsc, INV_DATE)
    '            End If
    '        End If

    '        'CUST_BALANCE = Val(dynARTCUST1.Fields("CUST_BALANCE").Value & "") + INV_TOTAL_AMOUNT * srev
    '        'dynARTCUST1.Fields("CUST_BALANCE").Value = CUST_BALANCE
    '    GoSub Update_CUST6
    '        'dynARTCUST1.Update

    '        If srev = -1 Then
    '            SQL = "Delete from ARTOPEN1 where CUST_CODE = '" & CUST_BILL_TO_CUST & "'"
    '            SQL = SQL & " and INV_TYPE = '" & INV_TYPE & "'"
    '            SQL = SQL & " and INV_NO = '" & INV_NO & "'"
    '            OraD.ExecuteSQL(SQL)
    '        Else
    '            dynARTOPEN1.AddNew()
    '            dynARTOPEN1.Fields("CUST_CODE").Value = CUST_BILL_TO_CUST
    '            dynARTOPEN1.Fields("INV_TYPE").Value = INV_TYPE
    '            dynARTOPEN1.Fields("INV_NO").Value = INV_NO
    '            dynARTOPEN1.Fields("INV_DATE").Value = INV_DATE
    '            dynARTOPEN1.Fields("CUST_STORE_NO").Value = CUST_STORE_NO
    '            dynARTOPEN1.Fields("POST_CODE").Value = POST_CODE
    '            dynARTOPEN1.Fields("TERM_CODE").Value = dynARTCUST1.Fields("TERM_CODE").Value
    '            dynARTOPEN1.Fields("INV_DUE_DATE").Value = INV_DUE_DATE
    '            dynARTOPEN1.Fields("INV_DISC_DATE").Value = INV_DISC_DATE
    '            dynARTOPEN1.Fields("SREP_CODE").Value = dynARTCUST1.Fields("SREP_CODE").Value
    '            dynARTOPEN1.Fields("STAX_CODE").Value = Null ' dynARTCUST1.Fields("STAX_CODE").Value
    '            dynARTOPEN1.Fields("APPLY_TO_INV_NO").Value = ""
    '            dynARTOPEN1.Fields("APPLY_TO_INV_TYPE").Value = ""
    '            dynARTOPEN1.Fields("INV_CUST_PO").Value = TRAN_CCVRW_REF
    '            dynARTOPEN1.Fields("INV_ORDR_NO").Value = ""
    '            dynARTOPEN1.Fields("INV_SALES").Value = INV_SALES
    '            dynARTOPEN1.Fields("INV_DISC").Value = 0
    '            dynARTOPEN1.Fields("INV_FREIGHT").Value = -S * Val(dynICWTRAN1.Fields("TRAN_FREIGHT").Value & "")
    '            dynARTOPEN1.Fields("INV_MISC_CHG").Value = -S * Val(dynICWTRAN1.Fields("TRAN_MISC_CHG").Value & "")
    '            dynARTOPEN1.Fields("INV_STAX").Value = 0 ' Val(dynICWTRAN1.Fields("TRAN_STAX").Value & "")
    '            dynARTOPEN1.Fields("INV_TOTAL_AMOUNT").Value = INV_TOTAL_AMOUNT
    '            dynARTOPEN1.Fields("INV_BALANCE").Value = INV_TOTAL_AMOUNT
    '            dynARTOPEN1.Fields("CUST_CODE_SO").Value = CUST_CODE
    '            If TRAN_TYPE = "S" Then
    '                dynARTOPEN1.Fields("REASON_CODE").Value = IC_PARM_REASON_SHP
    '            Else
    '                dynARTOPEN1.Fields("REASON_CODE").Value = REASON_CODE
    '                '            dynARTOPEN1.Fields("REASON_CODE").Value = IC_PARM_REASON_RTN
    '            End If
    '            dynARTOPEN1.Fields("INIT_OPER").Value = UserID
    '            dynARTOPEN1.Fields("INIT_DATE").Value = Now + NowTSD
    '            dynARTOPEN1.Fields("SALES_DIVISION_CODE").Value = SALES_DIVISION_CODE
    '            'CURRENCY HARD CODED AS USD VALUES FOR NOW.
    '            dynARTOPEN1.Fields("CURR_CODE").Value = "USD"
    '            dynARTOPEN1.Fields("CURR_EXCH_RATE").Value = 1
    '            dynARTOPEN1.Fields("INV_SALES_CURR").Value = INV_SALES
    '            dynARTOPEN1.Fields("INV_DISC_CURR").Value = 0
    '            dynARTOPEN1.Fields("INV_FREIGHT_CURR").Value = -S * Val(dynICWTRAN1.Fields("TRAN_FREIGHT").Value & "")
    '            dynARTOPEN1.Fields("INV_STAX_CURR").Value = 0
    '            dynARTOPEN1.Fields("INV_MISC_CHG_CURR").Value = -S * Val(dynICWTRAN1.Fields("TRAN_MISC_CHG").Value & "")
    '            dynARTOPEN1.Fields("INV_TOTAL_AMOUNT_CURR").Value = INV_TOTAL_AMOUNT
    '            dynARTOPEN1.Fields("INV_BALANCE_CURR").Value = INV_TOTAL_AMOUNT
    '            dynARTOPEN1.Fields("GST_TAX").Value = 0
    '            dynARTOPEN1.Fields("GST_TAX_CURR").Value = 0
    '            dynARTOPEN1.Update()
    '        End If
    '        OraD.Parameters("CUST_CODE").Value = CUST_CODE ' in case it was changed to bt
    '        Return

    'Update_CUST6:
    '        dynARTCUST6.Refresh()
    '        If dynARTCUST6.EOF Then
    '            dynARTCUST6.AddNew()
    '            dynARTCUST6.Fields("CUST_CODE").Value = OraD.Parameters("CUST_CODE").Value
    '        Else
    '            dynARTCUST6.Edit()
    '        End If
    '        If INV_TYPE = "I" And srev <> -1 Then
    '            dynARTCUST6.Fields("CUST_LAST_INV_NUM").Value = INV_NO
    '            dynARTCUST6.Fields("CUST_LAST_INV_DATE").Value = INV_DATE
    '            dynARTCUST6.Fields("CUST_LAST_INV_AMT").Value = INV_TOTAL_AMOUNT
    '            If IsNull(dynARTCUST6.Fields("CUST_FIRST_PURCH").Value) Then
    '                dynARTCUST6.Fields("CUST_FIRST_PURCH").Value = INV_DATE
    '            End If
    '            If dynARTCUST6.Fields("CUST_CODE").Value = CUST_BILL_TO_CUST Then
    '                If CUST_BALANCE > Val(dynARTCUST6.Fields("CUST_HIGH_BAL_AMT").Value & "") Then
    '                    dynARTCUST6.Fields("CUST_HIGH_BAL_DATE").Value = INV_DATE
    '                    dynARTCUST6.Fields("CUST_HIGH_BAL_AMT").Value = CUST_BALANCE
    '                End If
    '            End If
    '        End If
    '        If INV_TYPE = "I" Then
    '            dynARTCUST6.Fields("CUST_SALES_MTD").Value = Val(dynARTCUST6.Fields("CUST_SALES_MTD").Value & "") + INV_SALES
    '            dynARTCUST6.Fields("CUST_SALES_YTD").Value = Val(dynARTCUST6.Fields("CUST_SALES_YTD").Value & "") + INV_SALES
    '            dynARTCUST6.Fields("CUST_NUM_INV_MTD").Value = Val(dynARTCUST6.Fields("CUST_NUM_INV_MTD").Value & "") + srev
    '            dynARTCUST6.Fields("CUST_NUM_INV_YTD").Value = Val(dynARTCUST6.Fields("CUST_NUM_INV_YTD").Value & "") + srev
    '        Else
    '            dynARTCUST6.Fields("CUST_CRED_MTD").Value = Val(dynARTCUST6.Fields("CUST_CRED_MTD").Value & "") - INV_SALES
    '            dynARTCUST6.Fields("CUST_CRED_YTD").Value = Val(dynARTCUST6.Fields("CUST_CRED_YTD").Value & "") - INV_SALES
    '        End If
    '        dynARTCUST6.Update()
    '        Return

    'Update_Stat:
    '        OraD.Parameters("WHSE_CODE").Value = WHSE_CODE
    '        dynICTSTAT1.Refresh()
    '        If dynICTSTAT1.EOF Then
    '            dynICTSTAT1.AddNew()
    '            dynICTSTAT1.Fields("OPS_YYYYPP").Value = Period_Calc(CYP, Period_Add)
    '            dynICTSTAT1.Fields("STYLE_CODE").Value = STYLE_CODE
    '            dynICTSTAT1.Fields("COLOR_CODE").Value = COLOR_CODE
    '            dynICTSTAT1.Fields("WHSE_CODE").Value = WHSE_CODE
    '        Else
    '            dynICTSTAT1.Edit()
    '        End If
    '        dynICTSTAT1.Fields(ICTSTAT1field).Value = Val(dynICTSTAT1.Fields(ICTSTAT1field).Value & "") + TRAN_QTY * sstatsign
    '        dynICTSTAT1.Update()

    '        dynICTSTAT2.Refresh()
    '        If dynICTSTAT2.EOF Then
    '            dynICTSTAT2.AddNew()
    '            dynICTSTAT2.Fields("STYLE_CODE").Value = STYLE_CODE
    '            dynICTSTAT2.Fields("COLOR_CODE").Value = COLOR_CODE
    '            dynICTSTAT2.Fields("WHSE_CODE").Value = WHSE_CODE
    '        Else
    '            dynICTSTAT2.Edit()
    '        End If
    '        dynICTSTAT2.Fields("WHSE_QTY_ON_HAND").Value = Val(dynICTSTAT2.Fields("WHSE_QTY_ON_HAND").Value & "") + TRAN_QTY * ss
    '        dynICTSTAT2.Update()
    '    GoSub Update_Locations
    '        Return

    'Update_Locations:

    '        OraD.Parameters("WHSE_CODE").Value = WHSE_CODE
    '        OraD.Parameters("LOCATION_CODE").Value = LOCATION_CODE
    '        OraD.Parameters("BAR_CODE").Value = "0000000000"
    '        OraD.Parameters("STYLE_CODE").Value = STYLE_CODE
    '        OraD.Parameters("COLOR_CODE").Value = COLOR_CODE
    '        dynICTWHSE1.Refresh()
    '        If dynICTWHSE1.Fields("WHSE_LOCATOR").Value = "1" Then
    '            dynWHTLOCB1.Refresh()
    '            If dynWHTLOCB1.EOF Then
    '                dynWHTLOCB1.AddNew()
    '                dynWHTLOCB1.Fields("WHSE_CODE").Value = WHSE_CODE
    '                dynWHTLOCB1.Fields("LOCATION_CODE").Value = LOCATION_CODE
    '                dynWHTLOCB1.Fields("BAR_CODE").Value = "0000000000"
    '                dynWHTLOCB1.Fields("STYLE_CODE").Value = STYLE_CODE
    '                dynWHTLOCB1.Fields("COLOR_CODE").Value = COLOR_CODE
    '                dynWHTLOCB1.Fields("LOCATION_QTY").Value = TRAN_QTY * ss
    '            Else
    '                dynWHTLOCB1.Edit()
    '                dynWHTLOCB1.Fields("LOCATION_QTY").Value = dynWHTLOCB1.Fields("LOCATION_QTY").Value + (TRAN_QTY * ss)
    '            End If
    '            dynWHTLOCB1.Update()

    '            dynWHTLOCB2.Refresh()
    '            dynWHTLOCB2.AddNew()
    '            dynWHTLOCB2.Fields("WHSE_CODE").Value = WHSE_CODE
    '            dynWHTLOCB2.Fields("LOCATION_CODE").Value = LOCATION_CODE
    '            dynWHTLOCB2.Fields("BAR_CODE").Value = "0000000000"
    '            dynWHTLOCB2.Fields("STYLE_CODE").Value = STYLE_CODE
    '            dynWHTLOCB2.Fields("COLOR_CODE").Value = COLOR_CODE
    '            dynWHTLOCB2.Fields("WHSE_TRAN_QTY").Value = TRAN_QTY * ss
    '            dynWHTLOCB2.Fields("WHSE_TRAN_TYPE").Value = TRAN_TYPE
    '            dynWHTLOCB2.Fields("WHSE_TRAN_NO").Value = TRAN_NO
    '            TRAN_LNO = TRAN_LNO + 1  'TEST THIS WELL.
    '            dynWHTLOCB2.Fields("WHSE_TRAN_LNO").Value = TRAN_LNO
    '            dynWHTLOCB2.Fields("INIT_DATE").Value = Now + NowTSD
    '            dynWHTLOCB2.Fields("INIT_OPER").Value = UserID
    '            dynWHTLOCB2.Fields("LOCATION_CODE_OTHER").Value = ""
    '            dynWHTLOCB2.Fields("SESSION_ID").Value = ""
    '            dynWHTLOCB2.Update()
    '        End If
    '        dynICTWHSE1.Close()
    '        Return

    'Prep_TranType:
    '        If TRAN_TYPE = "S" Or TRAN_TYPE = "C" Then
    '            SQL = "Select * from ARTCUST1 where CUST_CODE = :CUST_CODE"
    '            dynARTCUST1 = OraD.CreateDynaset(SQL, 0&)
    '            SQL = "Select * from ARTCUST6 where CUST_CODE = :CUST_CODE"
    '            dynARTCUST6 = OraD.CreateDynaset(SQL, 8&)
    '            SQL = "Select * from ARTOPEN1 where ROWNUM < 1"
    '            dynARTOPEN1 = OraD.CreateDynaset(SQL, 8&)
    '            SQL = "Select * from SOTINVH1 where INV_TYPE = :INV_TYPE and INV_NO = :INV_NO"
    '            dynSOTINVH1 = OraD.CreateDynaset(SQL, 8&)
    '            SQL = "Select * from SOTINVH2 where INV_TYPE = :INV_TYPE and INV_NO = :INV_NO"
    '            dynSOTINVH2 = OraD.CreateDynaset(SQL, 8&)

    '            SQL = "Select TERM_CODE, TERM_DAYS_DUE, TERM_DAYS_DISC from TATTERM1"
    '            termt = Make_Table()
    '            PrepTranType = PrepTranType & "SC"
    '        End If

    '        If TRAN_TYPE = "A" Then
    '            PrepTranType = PrepTranType & "A"
    '        End If
    '        If TRAN_TYPE = "T" Then
    '            PrepTranType = PrepTranType & "T"
    '        End If
    '        Return

    '    End Sub

#End Region

    Public Shared Sub Calculate_FIFO( _
        frmASFBASE0 As ASFBASE0, _
        YP As String, _
        Optional ByRef calculate_usage As Boolean = False,
        Optional ByRef ICTCOSTA As String = "", _
        Optional ByRef ICTCOSTL As String = "", _
        Optional ByRef ICTCOSTU As String = "", _
        Optional ByRef ICTCOSTG As String = "", _
        Optional STYLE_CODE_single As String = "")

        Dim sqlSC As String = ""
        If STYLE_CODE_single <> "" Then sqlSC = " and STYLE_CODE = '" & STYLE_CODE_single & "'" & vbCrLf

        Dim rowGLTPARM2 As DataRow = frmASFBASE0.LookUp("GLTPARM2", YP)
        Dim DT As Date = rowGLTPARM2.Item("PRD_END_DATE")
        If YP = ASCMAIN1.CYP And Format(Now, "yyyyMMdd") < Format(DT, "yyyyMMdd") Then DT = Now.Date

        Dim POTSHIP5_SUM As String = ""
        Dim POTSHIP5_DTL As String = ""

        If calculate_usage Then
            ' Prepare Lots Usage (Work) Tables in Oracle and ADO.Net


            'If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.Progress("Create Tariff Work Tables")
            ' ALL OF THIS TARIFF STUFF SHOULD PROBABLY BE DONE IN A PERMANENT TABLE OR FIELD WITHIN SHIPMENT COST ENTRY
            ASCMAIN1.sql = "Select PO_SHIPMENT_NO, SUM (LANDING_COST_AMT) TARIFF" & vbCrLf _
                & " from POTSHIP5 WHERE COST_CATGY_CODE = 'TARIFF'" & vbCrLf _
                & " group by PO_SHIPMENT_NO"
            POTSHIP5_SUM = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & POTSHIP5_SUM & " Add DUTY_TOTAL NUMBER (13,2)")
            ASCDATA1.ExecuteSQL("Update " & POTSHIP5_SUM & " POTSHIP5_SUM Set DUTY_TOTAL = (Select SUM (PO_QTY_SHP * PO_COST_DUTY) from POTSHIP3 where PO_SHIPMENT_NO = POTSHIP5_SUM.PO_SHIPMENT_NO)")
            ASCDATA1.ExecuteSQL("Delete from " & POTSHIP5_SUM & " where NVL(TARIFF,0) = 0 OR NVL(DUTY_TOTAL,0) = 0")
            ASCDATA1.ExecuteSQL("Alter Table " & POTSHIP5_SUM & " Add FIRST_COST_TOTAL NUMBER (13,2)")
            If ASCMAIN1.CLIENT = "VAN" Then
                ASCDATA1.ExecuteSQL("Update " & POTSHIP5_SUM & " POTSHIP5_SUM Set FIRST_COST_TOTAL = (Select SUM (PO_QTY_SHP *(PO_COST_VCOST + PO_COST_MATLS + PO_COST_OTHER)) from POTSHIP3 where PO_SHIPMENT_NO = POTSHIP5_SUM.PO_SHIPMENT_NO)")
            Else
                ASCDATA1.ExecuteSQL("Update " & POTSHIP5_SUM & " POTSHIP5_SUM Set FIRST_COST_TOTAL = (Select SUM (PO_QTY_SHP * PO_COST) from POTSHIP3 where PO_SHIPMENT_NO = POTSHIP5_SUM.PO_SHIPMENT_NO)")
            End If

            ASCMAIN1.sql = "Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                & ", Sum (POTSHIP3.PO_QTY_SHP) QTY, Sum (POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST_DUTY) DUTY" & vbCrLf _
                & " from POTSHIP3,POTORDR2, " & POTSHIP5_SUM & " POTSHIP5_SUM" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTSHIP5_SUM.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & " group by POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                & " having Sum (POTSHIP3.PO_QTY_SHP) <> 0"

            ASCMAIN1.sql = "Select X.*, ROUND (X.DUTY * POTSHIP5_SUM.TARIFF / POTSHIP5_SUM.DUTY_TOTAL,2) TARIFF" & vbCrLf _
                & ", ROUND ((X.DUTY * POTSHIP5_SUM.TARIFF / POTSHIP5_SUM.DUTY_TOTAL)/QTY,6) TARIFF_UNIT_COST" & vbCrLf _
                & ", 100 * ROUND(POTSHIP5_SUM.TARIFF/POTSHIP5_SUM.FIRST_COST_TOTAL,6) TPCT" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") X, " & POTSHIP5_SUM & " POTSHIP5_SUM" & vbCrLf _
                & " where POTSHIP5_SUM.PO_SHIPMENT_NO = X.PO_SHIPMENT_NO"
            POTSHIP5_DTL = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & POTSHIP5_DTL & " Add Primary Key (PO_SHIPMENT_NO, PO_SHIPMENT_LNO, STYLE_CODE, COLOR_CODE)")
            'End If


            ASCMAIN1.Progress("Create Lots Usage Work Table")

            If ICTCOSTU = "" Then
                ASCMAIN1.sql = "Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                  & ", SOTINVH2.INV_TYPE TRAN_TYPE, SOTINVH2.INV_NO TRAN_NO, SOTINVH2.INV_LNO TRAN_LNO" & vbCrLf _
                  & ", SOTINVH1.INV_DATE TRAN_DATE" & vbCrLf _
                  & ", SOTINVH2.ORDR_QTY_SHIP TRAN_QTY, SOTINVH2.ORDR_UNIT_COST TRAN_COST" & vbCrLf _
                  & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST TRAN_COST_EXT" & vbCrLf _
                  & ", SOTINVH2.TARIFF_UNIT_COST" & vbCrLf _
                  & ", SOTINVH2.TARIFF_FLAG" & vbCrLf _
                  & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.TARIFF_UNIT_COST TARIFF_COST_EXT" & vbCrLf _
                  & ", SOTINVH1.ORDR_NO, SOTINVH1.ORDR_TYPE_CODE, SOTORDR1.ORDR_GROUP_NO from SOTINVH2,SOTINVH1,SOTORDR1 where ROWNUM < 1"
                ICTCOSTU = ASCMAIN1.Temp_Table
                ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTU & " Modify TRAN_LNO NUMBER (4,0)")
                ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTU & " Modify TRAN_COST_EXT NUMBER (18,6)")
                ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTU & " Modify TARIFF_COST_EXT NUMBER (18,6)")
                ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTU & " Modify TRAN_LNO NUMBER (6,0)")

                ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTU & " Add Primary Key (STYLE_CODE, COLOR_CODE, TRAN_TYPE, TRAN_NO, TRAN_LNO)")
                ASCDATA1.ExecuteSQL("Create Unique Index I_" & ICTCOSTU & "_1 on " & ICTCOSTU & " (TRAN_TYPE, TRAN_NO, TRAN_LNO)")
                ASCDATA1.ExecuteSQL("Create Index I_" & ICTCOSTU & "_2 on " & ICTCOSTU & " (STYLE_CODE, COLOR_CODE, TRAN_TYPE, TRAN_DATE, ORDR_GROUP_NO)")

                frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ICTCOSTU"), ICTCOSTU, "*")



                ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, ORDR_GROUP_NO, TRAN_TYPE  " & vbCrLf _
                      & ", TRAN_DATE, TRAN_QTY, TRAN_COST, TRAN_COST_EXT, ORDR_TYPE_CODE, ORDR_NO, TRAN_NO" & vbCrLf _
                      & ", TARIFF_UNIT_COST, TARIFF_COST_EXT, TARIFF_FLAG" & vbCrLf _
                      & " from " & ICTCOSTU & " where ROWNUM < 1"
                ICTCOSTG = ASCMAIN1.Temp_Table

                ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTG & " Add Primary Key (STYLE_CODE, COLOR_CODE, ORDR_GROUP_NO, TRAN_TYPE, TRAN_DATE)")

                frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ICTCOSTG"), ICTCOSTG, "*")

            Else
                ASCDATA1.ExecuteSQL("Delete from " & ICTCOSTU)
                ASCDATA1.ExecuteSQL("Delete from " & ICTCOSTG)
            End If

            ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTU & vbCrLf _
                                & " Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH2.INV_TYPE TRAN_TYPE, SOTINVH2.INV_NO TRAN_NO, SOTINVH2.INV_LNO TRAN_LNO, SOTINVH1.INV_DATE TRAN_DATE, -1 * SOTINVH2.ORDR_QTY_SHIP TRAN_QTY, 0 TRAN_COST, 0 TRAN_COST_EXT, 0 TARIFF_COST, 0 TARIFF_COST_EXT, NULL TARIFF_FLAG, SOTINVH1.ORDR_NO, SOTINVH1.ORDR_TYPE_CODE, DECODE(SOTINVH1.INV_TYPE,'C',SOTINVH1.INV_NO,SOTORDR1.ORDR_GROUP_NO) ORDR_GROUP_NO " & vbCrLf _
                                & " from SOTINVH2,SOTINVH1,SOTORDR1 where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO AND SOTORDR1.ORDR_NO(+) = SOTINVH1.ORDR_NO AND NVL(SOTINVH1.ORDR_TYPE_CODE,'REG') <> 'XFR' and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & YP & "'" _
                                & Replace(sqlSC, "and STYLE_CODE =", "and SOTINVH2.STYLE_CODE ="))

            ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTU & vbCrLf _
                                & " Select ICTIADJ2.STYLE_CODE, ICTIADJ2.COLOR_CODE, 'A' TRAN_TYPE, ICTIADJ2.ADJ_NO, ICTIADJ2.ADJ_LNO TRAN_LNO, ICTIADJ1.ADJ_DATE TRAN_DATE, ICTIADJ2.ADJ_QTY TRAN_QTY, 0 TRAN_COST, 0 TRAN_COST_EXT, 0 TARIFF_COST, 0 TARIFF_COST_EXT, NULL TARIFF_FLAG, NULL ORDR_NO,  NULL ORDR_TYPE_CODE, ICTIADJ2.ADJ_NO ORDR_GROUP_NO" & vbCrLf _
                                & " from ICTIADJ2,ICTIADJ1 where ICTIADJ1.ADJ_NO = ICTIADJ2.ADJ_NO and ICTIADJ1.OPS_YYYYPP = '" & YP & "'" _
                                & Replace(sqlSC, "and STYLE_CODE =", "and ICTIADJ2.STYLE_CODE ="))
            ASCMAIN1.AnalyzeTable(ICTCOSTU)

            '      ASCMAIN1.sql = "Select * from " & ICTCOSTU
            '      frmASFBASE0.Fill_Records("ICTCOSTU", "", True, ASCMAIN1.sql)


            ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTG & vbCrLf _
              & " Select STYLE_CODE, COLOR_CODE, ORDR_GROUP_NO, TRAN_TYPE, TRAN_DATE" & vbCrLf _
              & ", SUM(TRAN_QTY) TRAN_QTY, 0 TRAN_COST, 0 TRAN_COST_EXT, MIN(ORDR_TYPE_CODE) ORDR_TYPE_CODE" & vbCrLf _
              & ", MIN(ORDR_NO) ORDR_NO, DECODE(TRAN_TYPE ,'A',TRAN_NO,NULL) TRAN_NO, 0 TARIFF_UNIT_COST, 0 TARIFF_COST_EXT, NULL TARIFF_FLAG" & vbCrLf _
              & " from " & ICTCOSTU & " GROUP BY STYLE_CODE, COLOR_CODE, ORDR_GROUP_NO, TRAN_TYPE, TRAN_DATE, DECODE(TRAN_TYPE ,'A',TRAN_NO,NULL)")
            ASCMAIN1.AnalyzeTable(ICTCOSTG)

            ASCMAIN1.sql = "Select * from " & ICTCOSTG
            frmASFBASE0.Fill_Records("ICTCOSTG", "", True, ASCMAIN1.sql)

        End If


        ' Prepare Oracle Work Table for ICTCOSTA

        ASCMAIN1.Progress("Create Oracle Staging Table for FIFO History")

        If ICTCOSTA = "" Then
            ASCMAIN1.sql = "Select * from ICTCOSTA where ROWNUM < 1"
            ICTCOSTA = ASCMAIN1.Temp_Table

            ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTA & " Add Primary Key (OPS_YYYYPP, STYLE_CODE, COLOR_CODE)")
            ASCDATA1.ExecuteSQL("Create Unique Index I_" & ICTCOSTA & "_1 on " & ICTCOSTA & " (STYLE_CODE, COLOR_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Delete from " & ICTCOSTA)
        End If

        ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTA & vbCrLf _
                            & " (OPS_YYYYPP, STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND) " & vbCrLf _
                            & " Select '" & YP & "' OPS_YYYYPP, STYLE_CODE, COLOR_CODE, SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND " & vbCrLf _
                            & IIf(YP <> ASCMAIN1.CYP, _
                                  " from ICTSTAT5 where WHSE_QTY_ON_HAND <> 0 and OPS_YYYYPP = '" & YP & "'", _
                                  " from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0") & vbCrLf _
                              & sqlSC _
                              & " group by STYLE_CODE, COLOR_CODE")

        If calculate_usage Then
            ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTA & vbCrLf _
                    & " (OPS_YYYYPP, STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND) " & vbCrLf _
                    & " Select '" & YP & "' OPS_YYYYPP, STYLE_CODE, COLOR_CODE, 0 WHSE_QTY_ON_HAND " & vbCrLf _
                    & " from ICTSTAT1 where OPS_YYYYPP = '" & YP & "'" _
                    & sqlSC _
                    & " and (STYLE_CODE, COLOR_CODE) in (Select Distinct STYLE_CODE, COLOR_CODE from ICTSTAT1 where OPS_YYYYPP = '" & YP & "' minus Select Distinct STYLE_CODE, COLOR_CODE from " & ICTCOSTA & ")" _
                    & " group by STYLE_CODE, COLOR_CODE")
        End If

        ASCMAIN1.Progress("Establish Baseline Period using Markdowns")
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ASCDATA1.ExecuteSQL("Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = '200607'")
        Else
            ASCDATA1.ExecuteSQL("Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = '000000'")
        End If

        ' & "   where (TRAN_TYPE = 'B' OR TRAN_TYPE = 'M')" & vbCrLf _
        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is" _
            & "  Select STYLE_CODE, COLOR_CODE, MAX(OPS_YYYYPP) OPS_YYYYPP from ICTCOST1" & vbCrLf _
            & "   where (TRAN_TYPE = 'B')" & vbCrLf _
            & "     and OPS_YYYYPP <= '" & YP & "' group by STYLE_CODE, COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = R1.OPS_YYYYPP" & vbCrLf _
            & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Last Date Shipped/Received")

        ASCMAIN1.Progress("Establish Baseline Period using Markdowns")
        Dim OPS_YYYYPP_BASE As String = "000000"
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            'ASCDATA1.ExecuteSQL("Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = '200607'")
            OPS_YYYYPP_BASE = "200607"
        Else
            'ASCDATA1.ExecuteSQL("Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = '000000'")
        End If

        ' & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= R1.OPS_YYYYPP_BASE" & vbCrLf _
        Dim sqlLAST_SHP As String = "" _
            & "Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
            & ", MAX(SOTINVH1.INV_DATE) AS DATE_LAST_SHP" & vbCrLf _
            & " from SOTINVH1, SOTINVH2" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & OPS_YYYYPP_BASE & "'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & YP & "'" & vbCrLf _
            & " group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE"

        ' & "   and POTSHIP2.OPS_YYYYPP >= R1.OPS_YYYYPP_BASE" & vbCrLf _
        Dim sqlLAST_REC As String = "" _
            & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ", MAX(POTSHIP2.PO_DATE_RECEIVED) AS DATE_LAST_REC" & vbCrLf _
            & " from POTSHIP2, POTSHIP3, POTORDR2" & vbCrLf _
            & " where POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP >= '" & OPS_YYYYPP_BASE & "'" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP <= '" & YP & "'" & vbCrLf _
            & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE"

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & " Select ICTCOSTA.STYLE_CODE, ICTCOSTA.COLOR_CODE" & vbCrLf _
            & ", S.DATE_LAST_SHP, R.DATE_LAST_REC" & vbCrLf _
            & " from " & ICTCOSTA & " ICTCOSTA" & vbCrLf _
            & ", (" & sqlLAST_SHP & ") S" & vbCrLf _
            & ", (" & sqlLAST_REC & ") R" & vbCrLf _
            & " where S.STYLE_CODE (+) = ICTCOSTA.STYLE_CODE" & vbCrLf _
            & "   and S.COLOR_CODE (+) = ICTCOSTA.COLOR_CODE" & vbCrLf _
            & "   and R.STYLE_CODE (+) = ICTCOSTA.STYLE_CODE" & vbCrLf _
            & "   and R.COLOR_CODE (+) = ICTCOSTA.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTCOSTA & vbCrLf _
            & "    Set DATE_LAST_SHP = R1.DATE_LAST_SHP, DATE_LAST_REC = R1.DATE_LAST_REC" & vbCrLf _
            & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Loading FIFO Cost Work Table")

        ASCMAIN1.sql = "Select * from " & ICTCOSTA & " ICTCOSTA"
        If Not frmASFBASE0.dst.Tables.Contains("ICTCOSTA") Then
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ICTCOSTA"), ICTCOSTA, "**", 0, True)
            'With frmASFBASE0.dst.Tables("ICTCOSTA").Columns
            '    .Add("DAYS_X_VALUE", GetType(System.Decimal))
            '    .Add("DAYS", GetType(System.Int32))
            'End With
        End If
        frmASFBASE0.Fill_Records("ICTCOSTA", "", True, ASCMAIN1.sql)


        ' Cost Lost History plus Markdowns and Cost Adjustments
        ' A = Cost Adjustment
        ' B = Baseline
        ' M = Markdown
        ' R = Receipt

        ASCMAIN1.Progress("Get Cost Lots from Receipts History, plus Markdowns")

        Dim adding_cost = False
        If ASCMAIN1.DBS_SERVER = "VAN" And ASCMAIN1.DBS_COMPANY = "TST" Then adding_cost = True
        If ASCMAIN1.DBS_SERVER = "" And ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_COMPANY = "VAN" Then
            adding_cost = True
            'MsgBox("Adding Cost")
        End If
        adding_cost = False

        Dim RGITEMP As String = ", NVL(POTSHIP3.PO_COST_LANDED,0)"

        If ASCMAIN1.CLIENT = "RGI" Then
            RGITEMP = ", NVL(POTSHIP3.PO_COST,0)"
        End If

        Dim sqlICTCOSTL As String = "Select" & vbCrLf _
            & " STYLE_CODE, COLOR_CODE, OPS_YYYYPP, TRAN_DATE, TRAN_NO, TRAN_TYPE, TRAN_REF, TRAN_QTY, TRAN_COST, TRAN_LNO" & vbCrLf _
            & ", INIT_DATE, INIT_OPER, LAST_DATE, LAST_OPER" & vbCrLf _
            & ", WHSE_CODE, ORDR_NO, TRAN_STATUS, TARIFF_UNIT_COST, TARIFF_FLAG" & vbCrLf _
            & ", TRIM(TO_CHAR(ROWNUM,'0000000000')) RECORD_NO" & vbCrLf _
            & " from (" & vbCrLf _
            & " Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP" & vbCrLf _
            & ", ICTCOST1.TRAN_NO, ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST, 0 TRAN_LNO" & vbCrLf _
            & ", ICTCOST1.INIT_DATE, ICTCOST1.INIT_OPER, ICTCOST1.LAST_DATE, ICTCOST1.LAST_OPER" & vbCrLf _
            & ", NULL WHSE_CODE, NULL ORDR_NO, NULL TRAN_STATUS, 0 TARIFF_UNIT_COST, NULL TARIFF_FLAG" & vbCrLf _
            & " from ICTCOST1, " & ICTCOSTA & " ICTCOSTA" & vbCrLf _
            & " where ICTCOST1.OPS_YYYYPP >= ICTCOSTA.OPS_YYYYPP_BASE" & vbCrLf _
            & "   and ICTCOST1.OPS_YYYYPP <= '" & YP & "'" & vbCrLf _
            & "   and ICTCOST1.STYLE_CODE = ICTCOSTA.STYLE_CODE" & vbCrLf _
            & "   and ICTCOST1.COLOR_CODE = ICTCOSTA.COLOR_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & " Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE," & vbCrLf _
            & " ICTIREC2.OPS_YYYYPP, ICTIREC2.RECEIPT_NO TRAN_NO, 'R' TRAN_TYPE, POTSHIP1.PO_SHIPMENT_NO ||' - ' || POTSHIP1.PO_SHIP_VESSEL TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY" & vbCrLf _
            & RGITEMP & vbCrLf _
            & IIf(adding_cost, " + (NVL(POTSHIP3.PO_COST_VCOST,0) + NVL(POTSHIP3.PO_COST_MATLS,0) + NVL(POTSHIP3.PO_COST_OTHER,0)) * .02", "") _
            & " TRAN_COST, ICTIREC2.RECEIPT_LNO TRAN_LNO" & vbCrLf _
            & ", ICTIREC1.INIT_DATE, ICTIREC1.INIT_OPER, ICTIREC1.LAST_DATE, ICTIREC1.LAST_OPER" & vbCrLf _
            & ", ICTIREC1.WHSE_CODE, POTSHIP2.ORDR_NO, CASE WHEN ICTIREC1.REVERSED_BY_RECEIPT_NO IS NOT NULL THEN 'D' ELSE CASE WHEN REVERSES_RECEIPT_NO IS NOT NULL THEN 'S' ELSE NULL END END TRAN_STATUS" & vbCrLf _
            & ", 0 TARIFF_UNIT_COST, NULL TARIFF_FLAG" & vbCrLf _
            & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, ICTIREC1, ICTIREC2, " & ICTCOSTA & " ICTCOSTA" & vbCrLf _
            & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.STYLE_CODE = ICTCOSTA.STYLE_CODE" & vbCrLf _
            & "   and ICTIREC2.COLOR_CODE = ICTCOSTA.COLOR_CODE" & vbCrLf _
            & "   and NVL(POTSHIP2.OPS_YYYYPP_FIFO,ICTIREC2.OPS_YYYYPP) >= ICTCOSTA.OPS_YYYYPP_BASE" & vbCrLf _
            & "   and NVL(POTSHIP2.OPS_YYYYPP_FIFO,ICTIREC2.OPS_YYYYPP) <= '" & YP & "'" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP2.TRAN_NO = ICTIREC2.RECEIPT_NO " & vbCrLf _
            & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & ") order by STYLE_CODE, COLOR_CODE, TRAN_DATE, TRAN_TYPE, TRAN_NO" & vbCrLf

        sqlICTCOSTL = "Select '" & YP & "' OPS_YYYYPP_FIFO, X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
            & ", X.RECORD_NO, X.TRAN_NO, X.TRAN_TYPE, X.TRAN_REF, X.TRAN_DATE, X.OPS_YYYYPP" & vbCrLf _
            & ", X.TRAN_QTY, X.TRAN_COST, 0 LOT_QTY_ONHD, 0 LOT_AMT_ONHD, 0 LOT_QTY_USED, 0 LOT_AMT_USED" & vbCrLf _
            & ", 0 LOT_QTY_SHP, 0 LOT_AMT_SHP, 0 LOT_QTY_RTN, 0 LOT_AMT_RTN, 0 LOT_QTY_ADJ, 0 LOT_AMT_ADJ, '0' LOT_FIFO, '0' LOT_USED" & vbCrLf _
            & ", NULL BOUNDARY_CONDITION, 0 LOT_DAYS, X.TRAN_LNO, X.INIT_DATE, X.INIT_OPER, X.LAST_DATE, X.LAST_OPER" & vbCrLf _
            & ", X.WHSE_CODE, X.ORDR_NO, X.TRAN_STATUS" & vbCrLf _
            & IIf(calculate_usage, ", POTSHIP5_DTL.TARIFF_UNIT_COST", ", X.TARIFF_UNIT_COST") & vbCrLf _
            & IIf(calculate_usage, ", TO_CHAR(X.TRAN_DATE,'YYYYMMDD') || SUBSTR(TRIM(TO_CHAR(NVL(POTSHIP5_DTL.TPCT,0)-.01,'00')),-2,2) TARIFF_FLAG", ", X.TARIFF_FLAG") & vbCrLf _
            & " from (" & sqlICTCOSTL & ") X" & vbCrLf _
            & IIf(calculate_usage, ", " & POTSHIP5_DTL & " POTSHIP5_DTL, ICTIREC2" & vbCrLf, "") _
            & IIf(calculate_usage, " where ICTIREC2.RECEIPT_NO (+) = DECODE(X.TRAN_TYPE,'R',X.TRAN_NO,'~')" & vbCrLf _
                                 & "   and ICTIREC2.RECEIPT_LNO (+) = X.TRAN_LNO" & vbCrLf _
                                 & "   and POTSHIP5_DTL.PO_SHIPMENT_NO (+) = ICTIREC2.PO_SHIPMENT_NO" & vbCrLf _
                                 & "   and POTSHIP5_DTL.PO_SHIPMENT_LNO (+) = ICTIREC2.PO_SHIPMENT_LNO" & vbCrLf _
                                 & "   and POTSHIP5_DTL.STYLE_CODE (+) = ICTIREC2.STYLE_CODE" & vbCrLf _
                                 & "   and POTSHIP5_DTL.COLOR_CODE (+) = ICTIREC2.COLOR_CODE" & vbCrLf, "")


        Dim sqlICTCOSTL_BTB_PREV As String = "Select" & vbCrLf _
            & " STYLE_CODE, COLOR_CODE, OPS_YYYYPP, TRAN_DATE, TRAN_NO, TRAN_TYPE, TRAN_REF, TRAN_QTY, TRAN_COST, TRAN_LNO" & vbCrLf _
            & ", INIT_DATE, INIT_OPER, LAST_DATE, LAST_OPER" & vbCrLf _
            & ", WHSE_CODE, ORDR_NO, TRAN_STATUS" & vbCrLf _
            & ", TRIM(TO_CHAR(ROWNUM,'0000000000')) RECORD_NO" & vbCrLf _
            & " from (" & vbCrLf _
            & " Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE," & vbCrLf _
            & " ICTIREC2.OPS_YYYYPP, ICTIREC2.RECEIPT_NO TRAN_NO, 'R' TRAN_TYPE, POTSHIP1.PO_SHIPMENT_NO ||' - ' || POTSHIP1.PO_SHIP_VESSEL TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY" & vbCrLf _
            & RGITEMP & vbCrLf _
            & IIf(adding_cost, " + (NVL(POTSHIP3.PO_COST_VCOST,0) + NVL(POTSHIP3.PO_COST_MATLS,0) + NVL(POTSHIP3.PO_COST_OTHER,0)) * .02", "") _
            & " TRAN_COST, ICTIREC2.RECEIPT_LNO TRAN_LNO" & vbCrLf _
            & ", ICTIREC1.INIT_DATE, ICTIREC1.INIT_OPER, ICTIREC1.LAST_DATE, ICTIREC1.LAST_OPER" & vbCrLf _
            & ", ICTIREC1.WHSE_CODE, POTSHIP2.ORDR_NO, CASE WHEN ICTIREC1.REVERSED_BY_RECEIPT_NO IS NOT NULL THEN 'D' ELSE CASE WHEN REVERSES_RECEIPT_NO IS NOT NULL THEN 'S' ELSE NULL END END TRAN_STATUS" & vbCrLf _
            & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, ICTIREC1, ICTIREC2" & vbCrLf _
            & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP2.TRAN_NO = ICTIREC2.RECEIPT_NO " & vbCrLf _
            & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & ") order by STYLE_CODE, COLOR_CODE, TRAN_DATE, TRAN_TYPE, TRAN_NO" & vbCrLf

        sqlICTCOSTL_BTB_PREV = "Select '" & YP & "' OPS_YYYYPP_FIFO, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", RECORD_NO, TRAN_NO, TRAN_TYPE, TRAN_REF, TRAN_DATE, OPS_YYYYPP" & vbCrLf _
            & ", TRAN_QTY, TRAN_COST, 0 LOT_QTY_ONHD, 0 LOT_AMT_ONHD, 0 LOT_QTY_USED, 0 LOT_AMT_USED" & vbCrLf _
            & ", 0 LOT_QTY_SHP, 0 LOT_AMT_SHP, 0 LOT_QTY_RTN, 0 LOT_AMT_RTN, 0 LOT_QTY_ADJ, 0 LOT_AMT_ADJ, '0' LOT_FIFO, '0' LOT_USED" & vbCrLf _
            & ", NULL BOUNDARY_CONDITION, 0 LOT_DAYS, TRAN_LNO, INIT_DATE, INIT_OPER, LAST_DATE, LAST_OPER" & vbCrLf _
            & ", WHSE_CODE, ORDR_NO, TRAN_STATUS, 0 TARIFF_UNIT_COST, NULL TARIFF_FLAG" & vbCrLf _
            & " from (" & sqlICTCOSTL_BTB_PREV & ")"

        'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        '    ASCMAIN1.sql = Replace(sqlICTCOSTL, "RECEIPT_NO", "TRAN_NO")
        '    ASCMAIN1.sql = Replace(sqlICTCOSTL, "ICTIREC2", "ICTTRAN2")
        'End If

        If ICTCOSTL = "" Then
            ASCMAIN1.sql = "Select * from ICTCOSTL where ROWNUM < 1"
            ICTCOSTL = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTL & " Add Primary Key (OPS_YYYYPP_FIFO, STYLE_CODE, COLOR_CODE, RECORD_NO)")
        Else
            ASCDATA1.ExecuteSQL("Delete from " & ICTCOSTL)
        End If

        ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTL & " " & sqlICTCOSTL)


        ASCMAIN1.sql = "Select * from " & ICTCOSTL & " ICTCOSTL"
        ASCMAIN1.sql &= " where TRAN_STATUS is Null" ' to avoid using Reversed and Reversing Transactions
        ' note that the OH match to lots does NOT occur by WHSE_CODE
        ' - matching by WHSE_CODE would make no sense to VAN, while making perfect sense to NYA and to RGI
        ' - but matching by WHSE_CODE is a bit tricky when considering Cost Adjustments and Markdowns

        If Not frmASFBASE0.dst.Tables.Contains("ICTCOSTL") Then
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ICTCOSTL"), ICTCOSTL, "**", 0, True)
        End If
        frmASFBASE0.Fill_Records("ICTCOSTL", "", True, ASCMAIN1.sql)

        ' FIFO Calculation

        ASCMAIN1.Progress("Now Calculating FIFO Costs for Period " & YP)

        For Each rowICTCOSTA As DataRow In frmASFBASE0.dst.Tables("ICTCOSTA").Select("", "STYLE_CODE,COLOR_CODE")
            Dim STYLE_CODE As String = rowICTCOSTA.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTCOSTA.Item("COLOR_CODE")

            If ASCMAIN1.Running_in_VS AndAlso STYLE_CODE = "HL657-05" Then Stop

            ASCMAIN1.Progress("-", STYLE_CODE & "-" & COLOR_CODE)
            If ASCMAIN1.Running_in_VS And STYLE_CODE = "D3586SDM" Then Stop
            Dim OH_REMAINING As Int64 = Val(rowICTCOSTA.Item("WHSE_QTY_ON_HAND") & "")

            Dim LOT_QTY_ONHD As Int64 = 0
            Dim LOT_AMT_ONHD As Decimal = 0
            Dim LOT_DAYS_X_VALUE As Decimal = 0
            Dim LOT_DAYS_TOTAL As Int64 = 0
            Dim LOT_DAYS As Int64 = 0
            Dim TRAN_COST_M As Decimal = -1
            Dim TRAN_COST As Decimal = 0

            Dim sqlw As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            Dim sqlw_no_BTB = " and ORDR_NO is Null" ' do not want to FIFO OH anything from a BTB shipment receipt

            Dim sql_filter As String = "OPS_YYYYPP_FIFO = '" & YP & "' and " & sqlw & sqlw_no_BTB
            Dim sql_sortby As String = "RECORD_NO DESC"
            'If ASCMAIN1.CLIENT = "NYA" Then
            '    sql_sortby = "TRAN_DATE, RECORD_NO DESC"
            'End If

            'Dim rows() As DataRow = frmASFBASE0.dst.Tables("ICTCOSTL").Select _
            '                        ("OPS_YYYYPP_FIFO = '" & YP & "' and " & sqlw & sqlw_no_BTB, "RECORD_NO DESC")

            Dim rows() As DataRow = frmASFBASE0.dst.Tables("ICTCOSTL").Select _
                                   (sql_filter, sql_sortby)

            Dim irow As Integer = -1
            Dim irow_markdown As Integer = -1
            For Each rowICTCOSTL As DataRow In rows
                'If OH_REMAINING = 0 Then Exit For
                Dim TRAN_TYPE As String = rowICTCOSTL.Item("TRAN_TYPE") & ""
                Dim TRAN_QTY As Int64 = Val(rowICTCOSTL.Item("TRAN_QTY") & "")

                Dim ORDR_NO As String = rowICTCOSTL.Item("ORDR_NO") & ""
                Dim WHSE_CODE As String = rowICTCOSTL.Item("WHSE_CODE") & ""

                If TRAN_TYPE = "M" Then
                    If TRAN_COST_M < 0 Then
                        TRAN_COST = Val(rowICTCOSTL.Item("TRAN_COST") & "")
                        TRAN_COST_M = TRAN_COST
                        irow_markdown = irow + 1
                    End If
                Else
                    TRAN_COST = Val(rowICTCOSTL.Item("TRAN_COST") & "")
                    If TRAN_COST_M >= 0 Then
                        TRAN_COST = TRAN_COST_M
                    End If
                    Dim QTY As Int64 = TRAN_QTY
                    If OH_REMAINING <= TRAN_QTY Or TRAN_TYPE = "B" Then QTY = OH_REMAINING

                    LOT_QTY_ONHD += QTY
                    LOT_AMT_ONHD += QTY * TRAN_COST

                    If QTY < 0 Then rowICTCOSTL.Item("BOUNDARY_CONDITION") = "1"

                    OH_REMAINING -= QTY
                    rowICTCOSTL.Item("LOT_QTY_ONHD") = QTY
                    rowICTCOSTL.Item("LOT_AMT_ONHD") = QTY * TRAN_COST
                    rowICTCOSTL.Item("LOT_FIFO") = "1"

                    LOT_DAYS = DT.Subtract(CDate(rowICTCOSTL.Item("TRAN_DATE"))).Days
                    rowICTCOSTL.Item("LOT_DAYS") = LOT_DAYS
                    LOT_DAYS_X_VALUE += LOT_DAYS * QTY * TRAN_COST
                    LOT_DAYS_TOTAL += LOT_DAYS
                End If


                irow += 1
                If OH_REMAINING = 0 Then Exit For
            Next

            If irow = -1 And rows.Length > 0 Then
                irow = 0
            End If

            If OH_REMAINING <> 0 Or irow = -1 Then
                ' you have more on hand than you have lots
                ' if row_lot is nothing, then you had no lots
                ' if row_lot is not nothing, then you have the last lot in row_lot, so extend it
                If rows.Length <> 0 Then
                    Dim QTY As Int64 = OH_REMAINING
                    rows(irow).Item("LOT_QTY_ONHD") += QTY
                    rows(irow).Item("LOT_AMT_ONHD") += QTY * TRAN_COST ' Val(rows(irow).Item("TRAN_COST") & "")
                    LOT_QTY_ONHD += QTY
                    LOT_AMT_ONHD += QTY * TRAN_COST ' Val(rows(irow).Item("TRAN_COST") & "")
                    rows(irow).Item("BOUNDARY_CONDITION") = "1"

                    LOT_DAYS = DT.Subtract(CDate(rows(irow).Item("TRAN_DATE"))).Days
                    rows(irow).Item("LOT_DAYS") = LOT_DAYS
                    LOT_DAYS_X_VALUE += LOT_DAYS * QTY * TRAN_COST
                    LOT_DAYS_TOTAL += LOT_DAYS
                Else
                    Dim rowICTCOSTL As DataRow = frmASFBASE0.dst.Tables("ICTCOSTL").NewRow
                    With rowICTCOSTL
                        .Item("OPS_YYYYPP_FIFO") = YP
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE
                        .Item("RECORD_NO") = "0000000000"
                        .Item("TRAN_NO") = DBNull.Value
                        .Item("TRAN_TYPE") = "Z"
                        .Item("TRAN_REF") = DBNull.Value
                        .Item("TRAN_DATE") = DT
                        .Item("OPS_YYYYPP") = YP
                        .Item("TRAN_QTY") = 0
                        .Item("TRAN_COST") = 0
                        .Item("LOT_QTY_ONHD") = 0
                        .Item("LOT_AMT_ONHD") = 0
                        .Item("LOT_QTY_USED") = 0
                        .Item("LOT_AMT_USED") = 0
                        .Item("LOT_QTY_SHP") = 0
                        .Item("LOT_AMT_SHP") = 0
                        .Item("LOT_QTY_RTN") = 0
                        .Item("LOT_AMT_RTN") = 0
                        .Item("LOT_QTY_ADJ") = 0
                        .Item("LOT_AMT_ADJ") = 0
                        .Item("LOT_FIFO") = "1"
                        .Item("BOUNDARY_CONDITION") = "1"
                        .Item("INIT_DATE") = frmASFBASE0.DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    End With
                    frmASFBASE0.dst.Tables("ICTCOSTL").Rows.Add(rowICTCOSTL)
                    irow = 0
                    rows = New DataRow() {rowICTCOSTL}
                End If
            End If

            Dim STYLE_COST As Decimal = 0
            If LOT_QTY_ONHD > 0 Then STYLE_COST = LOT_AMT_ONHD / LOT_QTY_ONHD
            'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            '    ' KINDA THINK THAT THIS SHOULD BE FOR ALL COMPANIES
            '    If LOT_QTY_ONHD < 0 Then STYLE_COST = LOT_AMT_ONHD / LOT_QTY_ONHD
            'End If
            ' ADDING THE NEXT LINE 02/15/2023 BECAUSE NEGATIVE INVENTORIES ARE NOT GETTING COSTED
            If LOT_QTY_ONHD <> 0 Then STYLE_COST = LOT_AMT_ONHD / LOT_QTY_ONHD

            rowICTCOSTA.Item("LOT_QTY_ONHD") = LOT_QTY_ONHD
            rowICTCOSTA.Item("LOT_AMT_ONHD") = LOT_AMT_ONHD
            rowICTCOSTA.Item("STYLE_COST") = STYLE_COST

            If LOT_AMT_ONHD = 0 Then
                LOT_DAYS = 0
            Else
                LOT_DAYS = LOT_DAYS_X_VALUE / LOT_AMT_ONHD
            End If
            rowICTCOSTA.Item("LOT_DAYS") = LOT_DAYS

            If calculate_usage Then ' And irow <> -1 Then

                Dim LOT_QTY_USED As Int64 = 0
                Dim LOT_AMT_USED As Decimal = 0

                ' TRAN_NO TO GROUP NO U TO G, RIP THRU G RECORDS
                For Each rowICTCOSTG As DataRow In frmASFBASE0.dst.Tables("ICTCOSTG").Select _
                    (sqlw, "TRAN_DATE DESC, TRAN_TYPE DESC, ORDR_GROUP_NO DESC")
                    Dim TRAN_NO As String = rowICTCOSTG.Item("TRAN_NO") & ""
                    Dim TRAN_TYPE As String = rowICTCOSTG.Item("TRAN_TYPE") & ""
                    Dim TRAN_QTY As Int32 = Val(rowICTCOSTG.Item("TRAN_QTY") & "")
                    Dim ORDR_NO As String = rowICTCOSTG.Item("ORDR_NO") & ""
                    Dim ORDR_TYPE_CODE As String = rowICTCOSTG.Item("ORDR_TYPE_CODE") & ""

                    ' positive TRAN_QTY is a Return or positive Adjustment
                    ' negative TRAN_QTY is a Shipment or nexgative Adjustment
                    Dim TRAN_QTY_REMAINING As Int32 = TRAN_QTY
                    Dim BOUNDARY_CONDITION As String = ""

                    Dim TRAN_COST_EXT As Decimal = 0
                    Dim TARIFF_COST_EXT As Decimal = 0
                    Dim TARIFF_FLAG As String = ""

                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        If TRAN_TYPE = "A" And (TRAN_NO = "000673" Or TRAN_NO = "000687" Or TRAN_NO = "000688" Or TRAN_NO = "000689" Or TRAN_NO = "000696" Or TRAN_NO = "000697") Then
                            TRAN_QTY_REMAINING = 0
                        End If
                    End If

                    Do While TRAN_QTY_REMAINING <> 0
                        Dim lot_changed As Boolean = False
                        Dim LOT_QTY_TO_USE As Int32 = 0

                        If ORDR_TYPE_CODE = "BTB" Then

                            LOT_QTY_TO_USE = -1 * TRAN_QTY_REMAINING

                            sqlw = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
                            sqlw &= " and ORDR_NO = '" & ORDR_NO & "'" ' we want the specific BTB order so that we can match the supply to the demand
                            Dim rowICTCOSTL_BTB() As DataRow = frmASFBASE0.dst.Tables("ICTCOSTL").Select _
                                                    ("OPS_YYYYPP_FIFO = '" & YP & "' and " & sqlw, "RECORD_NO DESC")
                            If rowICTCOSTL_BTB.Length = 0 Then
                                Dim RECORD_NO_max As Int64 = Val(ASCDATA1.GetDataValue("SELECT MAX(RECORD_NO) FROM " & ICTCOSTL))

                                ASCMAIN1.sql = "Insert into " & ICTCOSTL & " " & Replace(sqlICTCOSTL_BTB_PREV, "TRIM(TO_CHAR(ROWNUM", "TRIM(TO_CHAR(" & CStr(RECORD_NO_max) & " + ROWNUM") & vbCrLf _
                                    & " where ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                                    & "   and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
                                ASCDATA1.ExecuteSQL()
                                ASCMAIN1.sql = "Select * from " & ICTCOSTL & " ICTCOSTL where RECORD_NO > '" & Format(RECORD_NO_max, "0000000000") & "'"
                                frmASFBASE0.Fill_Records("ICTCOSTL", "", False, ASCMAIN1.sql)
                                rowICTCOSTL_BTB = frmASFBASE0.dst.Tables("ICTCOSTL").Select _
                                                    ("OPS_YYYYPP_FIFO = '" & YP & "' and " & sqlw, "RECORD_NO DESC")
                            End If
                            If rowICTCOSTL_BTB.Length = 0 Then
                                'STYLE_CODE = 'MTX62163' and COLOR_CODE = 'MULT' and ORDR_NO = '0000452717'
                                ' why is this happening at RGI?



                                Dim LOT_COST As Decimal = 0
                                Dim LOT_QTY_REMAINING As Int32 =0
                                ' LOT_QTY_USED is negative for returns and positive adjustments
                                ' LOT_QTY_USED is positive for shipments and negative adjustments
                                '   Dim LOT_QTY_REMAINING_NEW As Int32 = LOT_QTY_REMAINING + TRAN_QTY_REMAININ

                                LOT_QTY_USED += LOT_QTY_TO_USE
                                LOT_AMT_USED += LOT_QTY_TO_USE * LOT_COST
                                TRAN_QTY_REMAINING += LOT_QTY_TO_USE

                            Else


                                With rowICTCOSTL_BTB(0)

                                    Dim LOT_COST As Decimal = Val(.Item("TRAN_COST") & "")
                                    Dim LOT_QTY_REMAINING As Int32 = Val(.Item("TRAN_QTY") & "") _
                                                                   - Val(.Item("LOT_QTY_ONHD") & "") _
                                                                   - Val(.Item("LOT_QTY_USED") & "")
                                    ' LOT_QTY_USED is negative for returns and positive adjustments
                                    ' LOT_QTY_USED is positive for shipments and negative adjustments
                                    '   Dim LOT_QTY_REMAINING_NEW As Int32 = LOT_QTY_REMAINING + TRAN_QTY_REMAINING

                                    .Item("LOT_USED") = "1"
                                    .Item("LOT_QTY_USED") += LOT_QTY_TO_USE
                                    .Item("LOT_AMT_USED") += LOT_QTY_TO_USE * LOT_COST

                                    Dim T As String = "SHP"
                                    If TRAN_TYPE = "C" Then T = "RTN"
                                    If TRAN_TYPE = "A" Then T = "ADJ"

                                    .Item("LOT_QTY_" & T) += LOT_QTY_TO_USE
                                    .Item("LOT_AMT_" & T) += LOT_QTY_TO_USE * LOT_COST

                                    TRAN_COST_EXT -= LOT_QTY_TO_USE * LOT_COST
                                    LOT_QTY_USED += LOT_QTY_TO_USE
                                    LOT_AMT_USED += LOT_QTY_TO_USE * LOT_COST
                                    TRAN_QTY_REMAINING += LOT_QTY_TO_USE
                                End With
                            End If
                        Else

                            If irow_markdown >= 0 Then
                                irow = irow_markdown
                            End If

                            Dim LOT_QTY As Int32 = Val(rows(irow).Item("TRAN_QTY") & "")
                            Dim LOT_COST As Decimal = Val(rows(irow).Item("TRAN_COST") & "")
                            Dim LOT_QTY_REMAINING As Int32 = Val(rows(irow).Item("TRAN_QTY") & "") _
                                                           - Val(rows(irow).Item("LOT_QTY_ONHD") & "") _
                                                           - Val(rows(irow).Item("LOT_QTY_USED") & "")
                            ' LOT_QTY_USED is negative for returns and positive adjustments
                            ' LOT_QTY_USED is positive for shipments and negative adjustments
                            Dim LOT_QTY_REMAINING_NEW As Int32 = LOT_QTY_REMAINING + TRAN_QTY_REMAINING

                            Dim irow_lot As Integer = irow
                            If irow_markdown >= 0 Then
                                LOT_QTY_TO_USE = -1 * TRAN_QTY_REMAINING
                            ElseIf LOT_QTY_REMAINING_NEW >= 0 And LOT_QTY_REMAINING_NEW <= LOT_QTY Then
                                LOT_QTY_TO_USE = -1 * TRAN_QTY_REMAINING
                            ElseIf LOT_QTY_REMAINING_NEW < 0 Then
                                If LOT_QTY_REMAINING > 0 Or irow >= rows.Length - 1 Then
                                    If LOT_QTY_REMAINING <= 0 Then
                                        BOUNDARY_CONDITION = "1"
                                        LOT_QTY_TO_USE = -1 * TRAN_QTY_REMAINING
                                    Else
                                        LOT_QTY_TO_USE = LOT_QTY_REMAINING
                                    End If
                                End If
                                If irow < rows.Length - 1 Then
                                    irow += 1
                                    lot_changed = True
                                End If
                            Else
                                If LOT_QTY_REMAINING < LOT_QTY Or irow <= 0 Then
                                    If LOT_QTY_REMAINING >= LOT_QTY Then
                                        BOUNDARY_CONDITION = "1"
                                        LOT_QTY_TO_USE = -1 * TRAN_QTY_REMAINING
                                    Else
                                        LOT_QTY_TO_USE = LOT_QTY_REMAINING - LOT_QTY
                                    End If
                                End If
                                If irow > 0 Then
                                    irow -= 1
                                    lot_changed = True
                                End If
                            End If

                            If LOT_QTY_TO_USE <> 0 Then
                                rows(irow_lot).Item("LOT_USED") = "1"
                                rows(irow_lot).Item("LOT_QTY_USED") += LOT_QTY_TO_USE
                                rows(irow_lot).Item("LOT_AMT_USED") += LOT_QTY_TO_USE * LOT_COST
                                Dim T As String = "SHP"
                                If TRAN_TYPE = "C" Then T = "RTN"
                                If TRAN_TYPE = "A" Then T = "ADJ"
                                rows(irow_lot).Item("LOT_QTY_" & T) += LOT_QTY_TO_USE
                                rows(irow_lot).Item("LOT_AMT_" & T) += LOT_QTY_TO_USE * LOT_COST
                                TRAN_COST_EXT -= LOT_QTY_TO_USE * LOT_COST
                                LOT_QTY_USED += LOT_QTY_TO_USE
                                LOT_AMT_USED += LOT_QTY_TO_USE * LOT_COST
                                TRAN_QTY_REMAINING += LOT_QTY_TO_USE

                                'If TARIFF_FLAG = "" Then
                                TARIFF_FLAG = rows(irow).Item("TARIFF_FLAG") & ""
                                'End If

                                Dim TARIFF_UNIT_COST As Decimal = Val(rows(irow).Item("TARIFF_UNIT_COST") & "")
                                If TARIFF_UNIT_COST <> 0 Then
                                    rows(irow_lot).Item("TARIFF_UNIT_COST") = TARIFF_UNIT_COST
                                    ' rows(irow_lot).Item("TARIFF_UNIT_COST") += LOT_QTY_TO_USE * TARIFF_UNIT_COST ' THIS FIELD SHOULD BE RENAMED TO TARIFF_COST - IT IS NOT A UNIT COST
                                    TARIFF_COST_EXT -= LOT_QTY_TO_USE * TARIFF_UNIT_COST
                                End If
                            End If
                        End If

                        If LOT_QTY_TO_USE = 0 And Not lot_changed Then Exit Do
                    Loop

                    rowICTCOSTG.Item("TRAN_COST_EXT") = TRAN_COST_EXT
                    If TRAN_QTY <> 0 Then rowICTCOSTG.Item("TRAN_COST") = TRAN_COST_EXT / TRAN_QTY
                    rowICTCOSTG.Item("TARIFF_COST_EXT") = TARIFF_COST_EXT
                    If TRAN_QTY <> 0 Then rowICTCOSTG.Item("TARIFF_UNIT_COST") = TARIFF_COST_EXT / TRAN_QTY
                    If TRAN_COST_EXT <> 0 Then rowICTCOSTG.Item("TARIFF_FLAG") = TARIFF_FLAG

                Next

                rowICTCOSTA.Item("LOT_QTY_USED") = LOT_QTY_USED
                rowICTCOSTA.Item("LOT_AMT_USED") = LOT_AMT_USED
            End If

        Next

        ASCMAIN1.Progress("Now Eliminating Lots which were not used")

        If STYLE_CODE_single = "" Then
            ASCDATA1.DeleteRows(frmASFBASE0.dst.Tables("ICTCOSTL"), "LOT_FIFO = '0' AND LOT_USED = '0' AND TRAN_TYPE <> 'M'")
        End If

        If STYLE_CODE_single = "" Then
            ASCMAIN1.Progress("Now Uploading Work Tables to Server")
            ASCMAIN1.Progress("-", "ICTCOSTL")
            '  frmASFBASE0.Update_Record_TDA("ICTCOSTL")

            If ICTCOSTL.StartsWith("ASW") Then
                ASCDATA1.ExecuteSQL("Truncate Table " & ICTCOSTL)
                frmASFBASE0.dst.Tables("ICTCOSTL").AcceptChanges()
                For Each row As DataRow In frmASFBASE0.dst.Tables("ICTCOSTL").Rows
                    row.SetAdded()
                Next
                frmASFBASE0.Create_BAs("ICTCOSTL")
                frmASFBASE0.Update_BAs("ICTCOSTL")
            Else
                frmASFBASE0.Update_Record_TDA("ICTCOSTL")
            End If



            ASCMAIN1.Progress("-", "ICTCOSTA")

            '  frmASFBASE0.Update_Record_TDA("ICTCOSTA")

            If ICTCOSTA.StartsWith("ASW") Then
                ASCDATA1.ExecuteSQL("Truncate Table " & ICTCOSTA)
                frmASFBASE0.dst.Tables("ICTCOSTA").AcceptChanges()
                For Each row As DataRow In frmASFBASE0.dst.Tables("ICTCOSTA").Rows
                    row.SetAdded()
                Next
                frmASFBASE0.Create_BAs("ICTCOSTA")
                frmASFBASE0.Update_BAs("ICTCOSTA")
            Else
                frmASFBASE0.Update_Record_TDA("ICTCOSTA")
            End If


            ASCMAIN1.Progress("-", "ICTCOSTG")
            '  frmASFBASE0.Update_Record_TDA("ICTCOSTU")

            If ICTCOSTG.StartsWith("ASW") Then 'If ICTCOSTG.StartsWith("ASW") And 1 <> 1 Then - i am not sure why this 1<>1 was installed
                ASCDATA1.ExecuteSQL("Truncate Table " & ICTCOSTG)
                frmASFBASE0.dst.Tables("ICTCOSTG").AcceptChanges()
                For Each row As DataRow In frmASFBASE0.dst.Tables("ICTCOSTG").Rows
                    row.SetAdded()
                Next
                frmASFBASE0.Create_BAs("ICTCOSTG")
                frmASFBASE0.Update_BAs("ICTCOSTG")
            Else
                frmASFBASE0.Update_Record_TDA("ICTCOSTG")
            End If

            If calculate_usage And ICTCOSTU <> "" Then
                ASCMAIN1.sql = "" _
                   & "Begin" & vbCrLf _
                   & " Declare Cursor C1 is " & vbCrLf _
                   & "  Select * " & vbCrLf _
                   & " from " & ICTCOSTG & ";" & vbCrLf _
                   & " Begin" & vbCrLf _
                   & "  For R1 in C1 Loop" & vbCrLf _
                   & "   Update " & ICTCOSTU & vbCrLf _
                   & "    Set TRAN_COST = R1.TRAN_COST,TRAN_COST_EXT = TRAN_QTY * R1.TRAN_COST" & vbCrLf _
                   & "       ,TARIFF_UNIT_COST = R1.TARIFF_UNIT_COST,TARIFF_COST_EXT = TRAN_QTY * R1.TARIFF_UNIT_COST" & vbCrLf _
                   & "       ,TARIFF_FLAG = R1.TARIFF_FLAG" & vbCrLf _
                   & "    where STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
                   & "      and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
                   & "      and TRAN_TYPE = R1.TRAN_TYPE" & vbCrLf _
                   & "      and TRAN_DATE = R1.TRAN_DATE" & vbCrLf _
                   & "      and ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                   & "  End Loop;" & vbCrLf _
                   & " End;" & vbCrLf _
                   & "End;"
                ASCDATA1.ExecuteSQL()
            End If

        End If

        ASCMAIN1.Progress("")
    End Sub


    Public Shared Sub Calculate_FIFO_Cost_OH(frmASFBASE0 As ASFBASE0, YP As String)
        '  MsgBox("Calculate_FIFO_Cost_OH", vbOKOnly, "Please Contact ABS")
        ASCMAIN1.Progress("Create Cost Lots Work Table")
        ASCMAIN1.sql = "Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE" & vbCrLf _
            & ", ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP, ICTCOST1.TRAN_TYPE" & vbCrLf _
            & ", ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
            & " from ICTCOST1"
        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "ICTCOSTL", "**", 0, False)
        With frmASFBASE0.dst.Tables("ICTCOSTL").Columns
            .Add("TRAN_REF")
            .Add("LOT_QTY_USED", GetType(System.Int64))
            .Add("LOT_AMT_TOTAL", GetType(System.Decimal), "ISNULL(TRAN_COST,0) * ISNULL(LOT_QTY_USED,0)")
        End With

        Dim VYP As String = YP
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            If YP = ASCMAIN1.CYP Then
                VYP = ASCMAIN1.Period_Calc(VYP, 1) ' GO OUT 1 MORE PERIOD B/C VAN IS ALWAYS LATE TO CLOSE AND A RECEIPT IN MAR WON'T GET PICKED UP PROPERLY FOR COSTING IF WE WE REALLY IN MAR YET VAN HASN'T CLOSED FEB
            End If
        End If

        ASCMAIN1.Progress("Create Oracle Staging Table for FIFO History")
        ASCMAIN1.sql = "Select * from ICTCOSTA where ROWNUM < 1"
        Dim ICTCOSTA As String = ASCMAIN1.Temp_Table

        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTA & " Add Primary Key (OPS_YYYYPP, STYLE_CODE, COLOR_CODE)")
        ASCDATA1.ExecuteSQL("Create Unique Index I_" & ICTCOSTA & "_1 on " & ICTCOSTA & " (STYLE_CODE, COLOR_CODE)")
        ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTA _
                            & " (OPS_YYYYPP, STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND) " _
                            & " Select '" & YP & "' OPS_YYYYPP, STYLE_CODE, COLOR_CODE, SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND " _
                            & IIf(YP <> ASCMAIN1.CYP, _
                                  " from ICTSTAT5 where WHSE_QTY_ON_HAND <> 0 and OPS_YYYYPP = '" & YP & "' group by OPS_YYYYPP, STYLE_CODE, COLOR_CODE", _
                                  " from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0 group by STYLE_CODE, COLOR_CODE"))

        ASCMAIN1.Progress("Establish Baseline Period using Markdowns")
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ASCDATA1.ExecuteSQL("Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = '200607'")
        Else
            ASCDATA1.ExecuteSQL("Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = '000000'")
        End If

        '& "   where (TRAN_TYPE = 'B' OR TRAN_TYPE = 'M')" & vbCrLf _
        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is" _
            & "  Select STYLE_CODE, COLOR_CODE, MAX(OPS_YYYYPP) OPS_YYYYPP from ICTCOST1" & vbCrLf _
            & "   where (TRAN_TYPE = 'B')" & vbCrLf _
            & "     and OPS_YYYYPP <= '" & YP & "' group by STYLE_CODE, COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTCOSTA & " Set OPS_YYYYPP_BASE = R1.OPS_YYYYPP" & vbCrLf _
            & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Last Date Shipped/Received")

        Dim sqlLAST_SHP As String = "" _
            & "Select MAX(SOTINVH1.INV_DATE) AS DATE_LAST_SHP" & vbCrLf _
            & " from SOTINVH1, SOTINVH2" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH2.COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
            & "   and SOTINVH2.ORDR_QTY_SHIP <> 0" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= R1.OPS_YYYYPP_BASE" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & VYP & "'"

        Dim sqlLAST_REC As String = "" _
            & "Select MAX(POTSHIP2.PO_DATE_RECEIVED) AS DATE_LAST_REC" & vbCrLf _
            & " from POTSHIP2, POTSHIP3, POTORDR2" & vbCrLf _
            & " where POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP3.PO_QTY_REC <> 0" & vbCrLf _
            & "   and POTORDR2.STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
            & "   and POTORDR2.COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP >= R1.OPS_YYYYPP_BASE" & vbCrLf _
            & "   and POTSHIP2.OPS_YYYYPP <= '" & VYP & "'"

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select * from " & ICTCOSTA & " for Update; " & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTCOSTA & " Set DATE_LAST_SHP = (" & sqlLAST_SHP & ")" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "   Update " & ICTCOSTA & " Set DATE_LAST_REC = (" & sqlLAST_REC & ")" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Loading FIFO Cost Work Table")

        ASCMAIN1.sql = "Select * from " & ICTCOSTA & " ICTCOSTA"
        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ICTCOSTA"), ICTCOSTA, "**", 0, True)
        frmASFBASE0.Fill_Records("ICTCOSTA")


        ASCMAIN1.Progress("Get Cost Lots from Receipts History, plus Markdowns")

        Dim adding_cost = False
        If ASCMAIN1.DBS_SERVER = "VAN" And ASCMAIN1.DBS_COMPANY = "TST" Then adding_cost = True
        If ASCMAIN1.DBS_SERVER = "" And ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_COMPANY = "VAN" Then
            adding_cost = True
            'MsgBox("Adding Cost")
        End If
        adding_cost = False

        ASCMAIN1.sql = "Select" & vbCrLf _
            & " STYLE_CODE, COLOR_CODE, TRAN_DATE, OPS_YYYYPP, TRAN_TYPE, TRAN_REF, TRAN_QTY," & vbCrLf _
            & " TRAN_COST, 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
            & " from (" & vbCrLf _
            & " Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
            & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
            & " from ICTCOST1, " & ICTCOSTA & " ICTCOSTA" & vbCrLf _
            & " where ICTCOST1.OPS_YYYYPP >= ICTCOSTA.OPS_YYYYPP_BASE" & vbCrLf _
            & "   and ICTCOST1.OPS_YYYYPP <= '" & YP & "'" & vbCrLf _
            & "   and ICTCOST1.STYLE_CODE = ICTCOSTA.STYLE_CODE" & vbCrLf _
            & "   and ICTCOST1.COLOR_CODE = ICTCOSTA.COLOR_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & " Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE," & vbCrLf _
            & " ICTTRAN2.OPS_YYYYPP, 'R' TRAN_TYPE, POTSHIP1.PO_SHIPMENT_NO ||' - ' || POTSHIP1.PO_SHIP_VESSEL TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY," & vbCrLf _
            & " NVL(POTSHIP3.PO_COST_LANDED,0) " & vbCrLf _
            & IIf(adding_cost, " + (NVL(POTSHIP3.PO_COST_VCOST,0) + NVL(POTSHIP3.PO_COST_MATLS,0) + NVL(POTSHIP3.PO_COST_OTHER,0)) * .02", "") _
            & " TRAN_COST" & vbCrLf _
            & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, ICTTRAN2, " & ICTCOSTA & " ICTCOSTA" & vbCrLf _
            & " where POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = ICTTRAN2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = ICTTRAN2.PO_ORDER_LNO" & vbCrLf _
            & "   and POTORDR2.STYLE_CODE = ICTCOSTA.STYLE_CODE" & vbCrLf _
            & "   and POTORDR2.COLOR_CODE = ICTCOSTA.COLOR_CODE" & vbCrLf _
            & "   and ICTTRAN2.OPS_YYYYPP >= ICTCOSTA.OPS_YYYYPP_BASE" & vbCrLf _
            & "   and ICTTRAN2.OPS_YYYYPP <= '" & VYP & "')" & vbCrLf

        Dim ICTCOSTR As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTCOSTR & "_1 on " & ICTCOSTR & " (STYLE_CODE, COLOR_CODE)")


        ASCMAIN1.Progress("Now saving FIFO Costs for Period " & YP)

        For Each rowICTCOSTA As DataRow In frmASFBASE0.dst.Tables("ICTCOSTA").Select("", "STYLE_CODE,COLOR_CODE")
            Dim STYLE_CODE As String = rowICTCOSTA.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTCOSTA.Item("COLOR_CODE")
            ASCMAIN1.Progress("-", STYLE_CODE & "-" & COLOR_CODE)
            ' If ASCMAIN1.Running_in_VS And STYLE_CODE = "5200IZR" Then Stop
            Dim OH_REMAINING As Int64 = Val(rowICTCOSTA.Item("WHSE_QTY_ON_HAND") & "")

            Dim sqlw As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            ASCMAIN1.sql = "Select * from " & ICTCOSTR & " where " & sqlw
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "TRAN_DATE DESC")
                If OH_REMAINING <= 0 Then Exit For
                Dim TRAN_QTY As Int64 = Val(row.Item("TRAN_QTY") & "")
                If OH_REMAINING <= TRAN_QTY Then
                    Calculate_FIFO_Cost_OH_Add_ICTCOSTL(frmASFBASE0, row, OH_REMAINING)
                Else
                    Calculate_FIFO_Cost_OH_Add_ICTCOSTL(frmASFBASE0, row, TRAN_QTY)
                End If
                OH_REMAINING = OH_REMAINING - TRAN_QTY
            Next

            Dim STYLE_COST As Decimal = 0
            Dim LOT_AMT_USED As Decimal = Val(frmASFBASE0.dst.Tables("ICTCOSTL").Compute("SUM(LOT_AMT_TOTAL)", sqlw) & "")
            Dim LOT_QTY_USED As Decimal = Val(frmASFBASE0.dst.Tables("ICTCOSTL").Compute("SUM(LOT_QTY_USED)", sqlw) & "")

            If LOT_QTY_USED <= 0 Then
                STYLE_COST = 0
            Else
                STYLE_COST = LOT_AMT_USED / LOT_QTY_USED
            End If

            rowICTCOSTA.Item("LOT_QTY_USED") = LOT_QTY_USED
            rowICTCOSTA.Item("STYLE_COST") = STYLE_COST
        Next
        frmASFBASE0.Update_Record_TDA("ICTCOSTA")
        ASCDATA1.ExecuteSQL("Delete from ICTCOSTA where OPS_YYYYPP = '" & YP & "'")
        ASCDATA1.ExecuteSQL("Insert into ICTCOSTA Select * from " & ICTCOSTA)

    End Sub

    Public Shared Sub Calculate_FIFO_Cost_OH_Add_ICTCOSTL(frmASFBASE0 As ASFBASE0, row As DataRow, LOT_QTY_USED As Int64)
        '    MsgBox("Calculate_FIFO_Cost_OH_Add_ICTCOSTL", vbOKOnly, "Please Contact ABS")
        Dim rowICTCOSTL As DataRow = frmASFBASE0.dst.Tables("ICTCOSTL").NewRow
        With rowICTCOSTL
            .Item("STYLE_CODE") = row.Item("STYLE_CODE")
            .Item("COLOR_CODE") = row.Item("COLOR_CODE")
            .Item("TRAN_DATE") = row.Item("TRAN_DATE")
            .Item("OPS_YYYYPP") = row.Item("OPS_YYYYPP")
            .Item("TRAN_TYPE") = row.Item("TRAN_TYPE")
            .Item("TRAN_REF") = row.Item("TRAN_REF")
            .Item("TRAN_QTY") = row.Item("TRAN_QTY")
            .Item("TRAN_COST") = row.Item("TRAN_COST")
            .Item("LOT_QTY_USED") = LOT_QTY_USED
        End With
        frmASFBASE0.dst.Tables("ICTCOSTL").Rows.Add(rowICTCOSTL)
    End Sub


    Public Shared Function Calc_Cost_OH( _
                        frmASFBASE0 As ASFBASE0, _
                        PERIOD As String, _
                        STYLE_CODE As String, _
                        COLOR_CODE As String, _
                        Calc_function_needs_to_create_ICTCOST1 As Boolean)
        ' I WOULD LIKE TO DELETE THIS ROUTINE
        ' 03/12/13 WJZ
        ' IT IS USED BY ICRISTA1 AND IS CALLED FOR EACH STYLE
        ' ITS SISTER ROUTINE, CALC_FIFO_COST_OH, DOES NOT PERMIT CALLING FOR AN INDIVIDUAL STYLE/COLOR
        ' SO UNTIL THEN, WE HAVE REDUNDANT ROUTINES
        '  MsgBox("Calc_Cost_OH", vbOKOnly, "Please Contact ABS")

        If ASCMAIN1.Running_in_VS Then
            MsgBox("This routine was removed from being called by the SSA - why are you here", _
                   MsgBoxStyle.OkOnly, "Please report to Walter")
        End If

        Dim A() As String
        ReDim A(1)

        ASCMAIN1.sql = "Select MAX(OPS_YYYYPP) FROM ICTCOST1" & vbCrLf _
         & " where (TRAN_TYPE = 'B' OR TRAN_TYPE = 'M')" & vbCrLf _
         & "   and (STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "')" & vbCrLf _
         & "   and OPS_YYYYPP <= '" & PERIOD & "'"
        Dim YP_Baseline_Cost As String = ASCDATA1.GetDataValue
        If YP_Baseline_Cost = "" Then
            YP_Baseline_Cost = "200607"
        End If

        If Calc_function_needs_to_create_ICTCOST1 Then
            ASCMAIN1.sql = "Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE" & vbCrLf _
                & ", ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP, ICTCOST1.TRAN_TYPE" & vbCrLf _
                & ", ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST, ICTCOST1.INIT_OPER, ICTCOST1.INIT_DATE" & vbCrLf _
                & " from ICTCOST1"
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "ICTCOST1", "**", 0, False)
            With frmASFBASE0.dst.Tables("ICTCOST1").Columns
                .Add("TRAN_REF")
                .Add("QTY_USED", GetType(System.Int64))
                .Add("COST_TOTAL", GetType(System.Decimal), "ISNULL(TRAN_COST,0) * ISNULL(QTY_USED,0)")
            End With
        Else
            frmASFBASE0.dst.Tables("ICTCOST1").Rows.Clear()
        End If

        If PERIOD = "" Then PERIOD = ASCMAIN1.CYP

        ASCMAIN1.sql = "Select SUM(NVL(WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND" & vbCrLf _
          & " from ICTSTAT2" & vbCrLf _
          & " where STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
          & "   and COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf

        If PERIOD <> ASCMAIN1.CYP Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "ICTSTAT2", "ICTSTAT5") & vbCrLf _
                & " and OPS_YYYYPP = '" & PERIOD & "'"
        End If
        Dim TOT_OH As Int64 = Val(ASCDATA1.GetDataValue)
        Dim OH_REMAINS As Int64 = TOT_OH


        Dim adding_cost = False
        If ASCMAIN1.DBS_SERVER = "VAN" And ASCMAIN1.DBS_COMPANY = "TST" Then adding_cost = True
        If ASCMAIN1.DBS_SERVER = "" And ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_COMPANY = "VAN" Then
            adding_cost = True
            ' MsgBox("Adding Cost")
        End If
        adding_cost = False

        'Find The Lot In Memory That Represents The Highest Sales.
        ASCMAIN1.sql = "Select" & vbCrLf _
        & " STYLE_CODE, COLOR_CODE, TRAN_DATE, OPS_YYYYPP, TRAN_TYPE, TRAN_REF, TRAN_QTY," & vbCrLf _
        & " TRAN_COST, 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
        & " FROM (" & vbCrLf _
        & " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
        & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
        & " FROM ICTCOST1" & vbCrLf _
        & " WHERE ICTCOST1.STYLE_CODE =  '" & STYLE_CODE & "'" & vbCrLf _
        & " AND ICTCOST1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
        & " AND ICTCOST1.OPS_YYYYPP >= '" & YP_Baseline_Cost & "'" & vbCrLf _
        & " AND ICTCOST1.OPS_YYYYPP <= '" & PERIOD & "'" & vbCrLf _
        & " UNION" & vbCrLf _
        & " SELECT POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE," & vbCrLf _
        & " ICTTRAN2.OPS_YYYYPP, 'R' TRAN_TYPE, POTSHIP1.PO_SHIPMENT_NO ||' - ' || POTSHIP1.PO_SHIP_VESSEL TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY," & vbCrLf _
        & " NVL(POTSHIP3.PO_COST_LANDED,0) " & vbCrLf _
        & IIf(adding_cost, " + (NVL(POTSHIP3.PO_COST_VCOST,0) + NVL(POTSHIP3.PO_COST_MATLS,0) + NVL(POTSHIP3.PO_COST_OTHER,0)) * .02", "") _
        & " TRAN_COST" & vbCrLf _
        & " FROM POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, ICTTRAN2" & vbCrLf _
        & " WHERE POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
        & " AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
        & " AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
        & " AND POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
        & " AND POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
        & " AND POTORDR2.PO_ORDER_NO = ICTTRAN2.PO_ORDER_NO" & vbCrLf _
        & " AND POTORDR2.PO_ORDER_LNO = ICTTRAN2.PO_ORDER_LNO" & vbCrLf _
        & " AND POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
        & " AND POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
        & " AND ICTTRAN2.OPS_YYYYPP >= '" & YP_Baseline_Cost & "'" & vbCrLf _
        & " AND ICTTRAN2.OPS_YYYYPP <= '" & PERIOD & "')" & vbCrLf

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "TRAN_DATE DESC")
            If OH_REMAINS <= 0 Then Exit For
            Dim TRAN_QTY As Int64 = Val(row.Item("TRAN_QTY") & "")
            Dim TRAN_TYPE As String = row.Item("TRAN_TYPE") & ""
            If TRAN_TYPE = "M" And TRAN_QTY = 0 Then
                Dim OPS_YYYYPP As String = row.Item("OPS_YYYYPP")
                ASCMAIN1.sql = "Select Sum (WHSE_QTY_ON_HAND) OH" _
                    & " from ICTSTAT5 where OPS_YYYYPP = '" & OPS_YYYYPP & "'" _
                    & "  and STYLE_CODE = '" & STYLE_CODE & "'" _
                    & "  and COLOR_CODE = '" & COLOR_CODE & "'"
                TRAN_QTY = Val(ASCDATA1.GetDataValue)
            End If
            If OH_REMAINS <= TRAN_QTY Then
                Calc_Cost_OH_Add_ICTCOST1(frmASFBASE0, row, OH_REMAINS)
            Else
                Calc_Cost_OH_Add_ICTCOST1(frmASFBASE0, row, TRAN_QTY)
            End If
            OH_REMAINS = OH_REMAINS - TRAN_QTY
        Next

        Dim CostPerUnit As Decimal
        Dim CostTotal As Decimal = Val(frmASFBASE0.dst.Tables("ICTCOST1").Compute("SUM(COST_TOTAL)", "") & "")
        Dim CostQty As Decimal = Val(frmASFBASE0.dst.Tables("ICTCOST1").Compute("SUM(QTY_USED)", "") & "")

        If CostQty <= 0 Then
            CostPerUnit = 0
        Else
            CostPerUnit = CostTotal / CostQty
        End If
        A(0) = Val(CostPerUnit)
        A(1) = Val(TOT_OH)
        Return Join(A, "|")
    End Function

    Public Shared Sub Calc_Cost_OH_Add_ICTCOST1(frmASFBASE0 As ASFBASE0, row As DataRow, QTY_USED As Int64)
        ' I WOULD LIKE TO DELETE THIS ROUTINE
        Dim rowICTCOST1 As DataRow = frmASFBASE0.dst.Tables("ICTCOST1").NewRow
        With rowICTCOST1
            .Item("STYLE_CODE") = row.Item("STYLE_CODE")
            .Item("COLOR_CODE") = row.Item("COLOR_CODE")
            .Item("TRAN_DATE") = row.Item("TRAN_DATE")
            .Item("OPS_YYYYPP") = row.Item("OPS_YYYYPP")
            .Item("TRAN_TYPE") = row.Item("TRAN_TYPE")
            .Item("TRAN_REF") = row.Item("TRAN_REF")
            .Item("TRAN_QTY") = row.Item("TRAN_QTY")
            .Item("TRAN_COST") = row.Item("TRAN_COST")
            .Item("QTY_USED") = QTY_USED
        End With
        frmASFBASE0.dst.Tables("ICTCOST1").Rows.Add(rowICTCOST1)
    End Sub

    'Function MAKE_COSTMF(PERIOD As String, SQLGROUPS As String, ZERO_COST As Boolean, USING_TEMP_INV As Boolean, TINVH1 As String, TINVH2 As String, Optional ORDR_TYPE As String)
    '        'IMPORTANT!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '        'An Exact copy of this function exists in VAN and VANX.
    '        'Any Changes you make to this function must be replicated in that function as well.
    '        'Failure to do so will result in your impending doom
    '        '
    '        ' This function will create a temp table of all styles in system with their respective currently calculated costs
    '        ' for a list of groups handed to it by the SQL SQLGROUPS
    '        '--------------------------------------------------------------------------------------------------------------

    '        Dim TT As String
    '        Dim dynORA As OraDynaset
    '        Dim TABLE As String

    '        Dim W As String
    '        Dim A() As String
    '        ReDim A(1)

    '        If ORDR_TYPE = "R" Then
    '            TBL1 = "SOTRSRV1"
    '            TBL2 = "SOTRSRV2"
    '        Else
    '            TBL1 = "SOTORDR1"
    '            TBL2 = "SOTORDR2"
    '        End If

    '        '    sql = " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf
    '        '    sql = sql & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST," & vbCrLf
    '        '    sql = sql & " 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf
    '        '    sql = sql & " FROM ICTCOST1" & vbCrLf
    '        '    sql = sql & " WHERE ROWNUM < 0" & vbCrLf
    '        Sql = "SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf
    '    sql = sql & " ICTCOST1.TRAN_TYPE, '" & String(50, " ") & "' TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST, 0.00 QTY_USED" & vbCrLf
    '        Sql = Sql & " FROM ICTCOST1" & vbCrLf
    '        Sql = Sql & " WHERE ROWNUM < 0"
    '        Call Ora_to_Acc(Nothing, "ICWCOST1", 0, "", Sql)

    '        Sql = "  SELECT " & TBL2 & ".STYLE_CODE, " & TBL2 & ".COLOR_CODE, ICTSTYL1.STYLE_DESC, SUM(1.11) STYLE_COST" & vbCrLf
    '        Sql = Sql & "  FROM " & TBL1 & ", " & TBL2 & ", ICTSTYL1" & vbCrLf
    '        If ORDR_TYPE = "R" Then
    '            Sql = Sql & "  WHERE " & TBL1 & ".RSRV_NO = " & TBL2 & ".RSRV_NO" & vbCrLf
    '            Sql = Sql & "  AND " & TBL1 & ".RSRV_NO IN (" & SQLGROUPS & ")" & vbCrLf
    '        Else
    '            Sql = Sql & "  WHERE " & TBL1 & ".ORDR_NO = " & TBL2 & ".ORDR_NO" & vbCrLf
    '            Sql = Sql & "  AND ORDR_GROUP_NO IN (" & SQLGROUPS & ")" & vbCrLf
    '        End If
    '        Sql = Sql & "  AND " & TBL2 & ".STYLE_CODE = ICTSTYL1.STYLE_CODE"
    '        Sql = Sql & "  GROUP BY " & TBL2 & ".STYLE_CODE, " & TBL2 & ".COLOR_CODE, ICTSTYL1.STYLE_DESC"
    '        TT = Temp_Table(Sql)
    '        Sql = "UPDATE " & TT & " SET STYLE_COST = NULL"
    '        OraD.ExecuteSQL(Sql)

    '        Sql = "ALTER Table " & TT & " MODIFY STYLE_COST NUMBER(12,6)"
    '        OraD.ExecuteSQL(Sql)

    '        Sql = "UPDATE " & TT & " SET STYLE_COST = 0"
    '        OraD.ExecuteSQL(Sql)

    '        Sql = "Create Index I_" & TT & "_1 ON " & TT & " (STYLE_CODE, COLOR_CODE)"
    '        OraD.ExecuteSQL(Sql)

    '        If Not ZERO_COST Then
    '            Sql = "SELECT * FROM " & TT & " ORDER BY STYLE_CODE, COLOR_CODE"

    '            dynORA = OraD.CreateDynaset(Sql, 0&)
    '            Do While Not dynORA.EOF
    '                'w = Calc_Cost(dynORA.Fields("STYLE_CODE").Value, dynORA.Fields("COLOR_CODE").Value, False, 0, USING_TEMP_INV, TINVH1, TINVH2)
    '                W = Calc_Cost_New(PERIOD, dynORA.Fields("STYLE_CODE").Value, dynORA.Fields("COLOR_CODE").Value, False)
    '                A() = Split(W, "|")
    '                dynORA.Edit()
    '                dynORA.Fields("STYLE_COST").Value = Val(A(0))
    '                dynORA.Update()
    '                dynORA.MoveNext()
    '            Loop
    '        End If
    '        MAKE_COSTMF = TT
    '    End Function

    '    Public Shared Function Calc_Cost_New( _
    '                                        frmASFBASE0 As ASFBASE0, _
    '                                        PERIOD As String, _
    '                                        STYLE_CODE As String, _
    '                                        COLOR_CODE As String, _
    '                                        Calc_function_needs_to_create_ICTCOST1 As Boolean)

    '        Dim A() As String
    '        ReDim A(1)

    '        ASCMAIN1.sql = "Select MAX(OPS_YYYYPP) FROM ICTCOST1" & vbCrLf _
    '            & " where (TRAN_TYPE = 'B' OR TRAN_TYPE = 'M')" & vbCrLf _
    '            & "   and (STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "')"
    '        Dim YP_Baseline_Cost As String = ASCDATA1.GetDataValue
    '        If YP_Baseline_Cost = "" Then
    '            YP_Baseline_Cost = "200607"
    '        End If


    '        If Calc_function_needs_to_create_ICTCOST1 Then
    '            ASCMAIN1.sql = "Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE" & vbCrLf _
    '                & ", ICTCOST1.OPS_YYYYPP" & vbCrLf _
    '                & ", ICTCOST1.TRAN_TYPE, '" & "".PadLeft(50, " ") & "' TRAN_REF, ICTCOST1.TRAN_QTY" & vbCrLf _
    '                & ", ICTCOST1.TRAN_COST, 0.00 QTY_USED" & vbCrLf _
    '                & " from ICTCOST1"
    '            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "ICTCOST1", "**", 0, False)
    '            frmASFBASE0.dst.Tables("ICTCOST1").Columns.Add("COST_TOTAL", GetType(System.Decimal), "ISNULL(TRAN_COST,0) * ISNULL(QTY_USED,0)")
    '        Else
    '            frmASFBASE0.dst.Tables("ICTCOST1").Rows.Clear()
    '        End If


    '        Dim W As Long

    '        Dim SHIP_REMAINS As Long
    '        Dim TOT_CONSUMED As Long
    '        Dim TOT_R_COST As Double
    '        Dim LOT_REMAIN As Long
    '        Dim LOT_USED As Long

    '        'Calculate Sales For This Month.
    '        ASCMAIN1.sql = "Select SUM(D.ORDR_QTY_SHIP) TRAN_QTY" & vbCrLf _
    '            & " FROM SOTINVH1 H, SOTINVH2 D" & vbCrLf _
    '            & " WHERE H.INV_TYPE = D.INV_TYPE" & vbCrLf _
    '            & " AND H.INV_NO = D.INV_NO" & vbCrLf _
    '            & " AND D.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
    '            & " AND D.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
    '            & " AND H.INV_TYPE = 'I'" & vbCrLf _
    '            & " AND D.ORDR_YYYYPP_UPDATED = '" & PERIOD & "'" & vbCrLf _
    '            & " HAVING SUM(D.ORDR_QTY_SHIP) <> 0"
    '        Dim TOT_SHIP_TM As Int64 = Val(ASCDATA1.GetDataValue)

    '        'Calculate Sales Prior to this Month.
    '        ASCMAIN1.sql = "Select SUM(D.ORDR_QTY_SHIP) TRAN_QTY" & vbCrLf _
    '            & " FROM SOTINVH1 H, SOTINVH2 D" & vbCrLf _
    '            & " WHERE H.INV_TYPE = D.INV_TYPE" & vbCrLf _
    '            & " AND H.INV_NO = D.INV_NO" & vbCrLf _
    '            & " AND D.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
    '            & " AND D.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
    '            & " AND H.INV_TYPE = 'I'" & vbCrLf _
    '            & " AND D.ORDR_YYYYPP_UPDATED < '" & PERIOD & "'" & vbCrLf _
    '            & " AND D.ORDR_YYYYPP_UPDATED >= '" & YP_Baseline_Cost & "'" & vbCrLf _
    '            & " HAVING SUM(D.ORDR_QTY_SHIP) <> 0"
    '        Dim TOT_SHIP_TD As Int64 = Val(ASCDATA1.GetDataValue)

    '        'Find The Lot In Memory That Represents The Highest Sales.
    '        ASCMAIN1.sql = "Select" & vbCrLf _
    '            & " STYLE_CODE, COLOR_CODE, TRAN_DATE, OPS_YYYYPP, TRAN_TYPE, TRAN_REF, TRAN_QTY," & vbCrLf _
    '            & " TRAN_COST, 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
    '            & " FROM (" & vbCrLf _
    '            & " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
    '            & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
    '            & " FROM ICTCOST1" & vbCrLf _
    '            & " WHERE ICTCOST1.STYLE_CODE =  '" & STYLE_CODE & "'" & vbCrLf _
    '            & " AND ICTCOST1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
    '            & " AND ICTCOST1.OPS_YYYYPP >= '" & YP_Baseline_Cost & "'" & vbCrLf _
    '            & " UNION" & vbCrLf _
    '            & " SELECT POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE," & vbCrLf _
    '            & " ICTTRAN2.OPS_YYYYPP, 'R' TRAN_TYPE, POTSHIP1.PO_SHIPMENT_NO ||' - ' || POTSHIP1.PO_SHIP_VESSEL TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY," & vbCrLf _
    '            & " POTSHIP3.PO_COST_LANDED TRAN_COST" & vbCrLf _
    '            & " FROM POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, ICTTRAN2" & vbCrLf _
    '            & " WHERE POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
    '            & " AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
    '            & " AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
    '            & " AND POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
    '            & " AND POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
    '            & " AND POTORDR2.PO_ORDER_NO = ICTTRAN2.PO_ORDER_NO" & vbCrLf _
    '            & " AND POTORDR2.PO_ORDER_LNO = ICTTRAN2.PO_ORDER_LNO" & vbCrLf _
    '            & " AND POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
    '            & " AND POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
    '            & " AND ICTTRAN2.OPS_YYYYPP >= '" & YP_Baseline_Cost & "')"
    '        Dim dt As DataTable = ASCDATA1.GetDataTable
    '        If dt.Rows.Count = 0 Then
    '            SHIP_REMAINS = 0
    '            LOT_REMAIN = 0
    '            Calc_Cost_New_Add_ICTCOST1(frmASFBASE0, dt.NewRow, SHIP_REMAINS, LOT_REMAIN)
    '            GoTo No_Lots_Found
    '        Else
    '            SHIP_REMAINS = TOT_SHIP_TM
    '            For Each row As DataRow In dt.Select("", "STYLE_CODE, COLOR_CODE, TRAN_TYPE, TRAN_DATE")
    '                TOT_CONSUMED = TOT_CONSUMED + dynORA.Fields("TRAN_QTY").Value
    '                If TOT_CONSUMED > TOT_SHIP_TD Then
    '                    Exit For
    '                End If

    '            Next
    '        End If


    '            LOT_REMAIN = TOT_CONSUMED - TOT_SHIP_TD
    '            If LOT_REMAIN <= 0 Then
    '                dynORA.MovePrevious()
    '            End If
    '    GoSub CalcLot
    '            If Not dynORA.EOF Then
    '                dynORA.MoveNext()
    '                Do While Not dynORA.EOF
    '                    LOT_REMAIN = dynORA.Fields("TRAN_QTY").Value
    '                    If SHIP_REMAINS <= 0 Then
    '                        Exit Do
    '                    End If
    '            GoSub CalcLot
    '                    dynORA.MoveNext()
    '                Loop
    '            End If
    '            'Stop
    '            Dim ReturnCost As Double
    '            Dim CostQty As Double
    '            Dim CostTotal As Double
    '            Sql = "SELECT"
    '            Sql = Sql & " (TRAN_COST * QTY_USED) AS COST_TOTAL, QTY_USED"
    '            Sql = Sql & " FROM ICWCOST1"
    '            dynWK = AccD.OpenRecordset(Sql, dbOpenForwardOnly)
    '            Do While Not dynWK.EOF
    '                CostQty = CostQty + Val(dynWK.Fields("QTY_USED").Value)
    '                CostTotal = CostTotal + Val(dynWK.Fields("COST_TOTAL").Value & "")
    '                dynWK.MoveNext()
    '            Loop
    '            dynWK.Close()
    '            dynORA.Close()
    'No_Lots_Found:
    '            Sql = "SELECT D.STYLE_CODE, D.COLOR_CODE, H.INV_DATE TRAN_DATE," & vbCrLf
    '            Sql = Sql & "  D.ORDR_YYYYPP_UPDATED  OPS_YYYYPP, 'S' TRAN_TYPE, '' TRAN_REF," & vbCrLf
    '            Sql = Sql & "  SUM(D.ORDR_QTY_SHIP) TRAN_QTY, D.ORDR_UNIT_COST TRAN_COST, SUM(D.ORDR_QTY_SHIP) QTY_USED" & vbCrLf
    '            Sql = Sql & "  FROM SOTINVH1 H, SOTINVH2 D" & vbCrLf
    '            Sql = Sql & "  WHERE H.INV_TYPE = D.INV_TYPE" & vbCrLf
    '            Sql = Sql & "  AND H.INV_NO = D.INV_NO" & vbCrLf
    '            Sql = Sql & "  AND D.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf
    '            Sql = Sql & "  AND D.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf
    '            Sql = Sql & "  AND H.INV_TYPE = 'I'" & vbCrLf
    '            Sql = Sql & "  AND D.ORDR_YYYYPP_UPDATED = '" & PERIOD & "'" & vbCrLf
    '            Sql = Sql & "  HAVING SUM(D.ORDR_QTY_SHIP) <> 0" & vbCrLf
    '            Sql = Sql & "  GROUP BY D.STYLE_CODE, D.COLOR_CODE, H.INV_DATE," & vbCrLf
    '            Sql = Sql & "  D.ORDR_YYYYPP_UPDATED, 'S',  D.ORDR_UNIT_COST"
    '            Call Ora_to_Acc(Nothing, "ICWCOST1", 0, "N", Sql)

    '            If CostQty <= 0 Then
    '                ReturnCost = 0
    '            Else
    '                ReturnCost = CostTotal / CostQty
    '            End If
    '            A(0) = Val(ReturnCost)
    '            A(1) = Val(TOT_SHIP_TM)
    '            Calc_Cost_New = Join(A(), "|")


    '    End Function


    Public Shared Sub Calc_Cost_New_Add_ICTCOST1( _
                                                frmASFBASE0 As ASFBASE0, _
                                                row As DataRow, _
                                                ByRef SHIP_REMAINS As Int64, _
                                                ByRef LOT_REMAIN As Int64)
        Dim LOT_USED As Int64 = 0
        If LOT_REMAIN <= 0 Then
            LOT_USED = SHIP_REMAINS
        Else
            If LOT_REMAIN <= SHIP_REMAINS Then
                LOT_USED = LOT_REMAIN
            Else
                LOT_USED = SHIP_REMAINS
            End If
        End If
        SHIP_REMAINS = SHIP_REMAINS - LOT_USED

        Dim rowICTCOST1 As DataRow = frmASFBASE0.dst.Tables("ICTCOST1").NewRow
        With rowICTCOST1
            .Item("STYLE_CODE") = row.Item("STYLE_CODE")
            .Item("COLOR_CODE") = row.Item("COLOR_CODE")
            .Item("TRAN_DATE") = row.Item("TRAN_DATE")
            .Item("OPS_YYYYPP") = row.Item("OPS_YYYYPP")
            .Item("TRAN_TYPE") = row.Item("TRAN_TYPE")
            .Item("TRAN_REF") = row.Item("TRAN_REF")
            .Item("TRAN_QTY") = row.Item("TRAN_QTY")
            .Item("TRAN_COST") = row.Item("TRAN_COST")
            .Item("QTY_USED") = LOT_USED
            .Item("INIT_OPER") = row.Item("INIT_OPER")
            .Item("INIT_DATE") = row.Item("INIT_DATE")
        End With
        frmASFBASE0.dst.Tables("ICTCOST1").Rows.Add(rowICTCOST1)
    End Sub


    Public Shared Sub Update_ICTSTAT2( _
        STYLE_CODE As String, _
        COLOR_CODE As String, _
        WHSE_CODE As String, _
        COLUMN_NAME As String, _
        QTY As Int64)

        Dim WHSE_QTY_ON_HAND As Int64 = IIf(COLUMN_NAME = "WHSE_QTY_ON_HAND", QTY, 0)
        Dim WHSE_QTY_ON_ORDER As Int64 = IIf(COLUMN_NAME = "WHSE_QTY_ON_ORDER", QTY, 0)
        Dim WHSE_QTY_TRAN As Int64 = IIf(COLUMN_NAME = "WHSE_QTY_TRAN", QTY, 0)
        Dim WHSE_QTY_OPEN As Int64 = IIf(COLUMN_NAME = "WHSE_QTY_OPEN", QTY, 0)
        Dim WHSE_QTY_PICK As Int64 = IIf(COLUMN_NAME = "WHSE_QTY_PICK", QTY, 0)
        Dim WHSE_QTY_ALLO As Int64 = IIf(COLUMN_NAME = "WHSE_QTY_ALLO", QTY, 0)

        ASCDATA1.ExecuteSP("ICPSTAT2", "VVVNNNNNN", _
                           New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, _
                                         WHSE_QTY_ON_HAND, WHSE_QTY_ON_ORDER, WHSE_QTY_TRAN, _
                                         WHSE_QTY_OPEN, WHSE_QTY_PICK, 0}, _
                           New String() {"STYLE_CODE_IN", "COLOR_CODE_IN", "WHSE_CODE_IN", _
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in", _
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

#Region "ODG"

    Public Shared Function Update_Return(ByVal frm As ASFBASE1) As String

        Dim rowSOTRTRN1 As DataRow = frm.dst.Tables("SOTRTRN1").Rows(0)
        frm.dst.Tables("ARTOPEN1").Rows.Clear()
        frm.dst.Tables("SOTINVH1").Rows.Clear()
        frm.dst.Tables("SOTINVH2").Rows.Clear()
        frm.dst.Tables("SOTINVHM").Rows.Clear()

        'Dim rowSOTREAS1 As DataRow = frm.dst.Tables("SOTREAS1").Rows.Find(New String() _
        '    {rowSOTRTRN1.Item("REASON_CODE")})
        'Dim ACCT_CODE_RTN As String = rowSOTREAS1.Item("ACCT_CODE") & ""

        Dim rowARTCUST1 As DataRow = frm.LookUp("ARTCUST1", rowSOTRTRN1.Item("CUST_CODE"))
        If rowARTCUST1.Item("CUST_BILL_TO_CUST") & "" = "" Then
            rowSOTRTRN1.Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_CODE")
        Else
            rowSOTRTRN1.Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_BILL_TO_CUST")
        End If
        Dim rowARTCUST1_BT As DataRow = frm.LookUp("ARTCUST1", rowSOTRTRN1.Item("CUST_BILL_TO_CUST"))

        Dim rowARTPOST1 As DataRow = frm.LookUp("ARTPOST1", rowARTCUST1_BT.Item("POST_CODE"))
        Dim rowARTSTAX1 As DataRow = Nothing ' frm.LookUp("ARTSTAX1", rowARTCUST1_BT.Item("STAX_CODE") & "")
        Dim rowARTSTAX2 As DataRow = Nothing ' frm.LookUp("ARTSTAX2", frm.dst.Tables("SOTRTRN5").Rows(0).Item("CUST_ZIP_CODE") & "")
        'Dim rowSOTMISC1 As DataRow = frm.LookUp("SOTMISC1", frm.ROWs("SOTPARM1").Item("SO_PARM_MISC_CHG_RSF") & "")

        Dim INV_NO As String = ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            'INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
            INV_NO = rowSOTRTRN1.Item("RTRN_NO")
        Else
            INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        End If
        rowSOTRTRN1.Item("INV_NO") = INV_NO

        Dim RTRN_GNO As Integer = 0
        Dim rowSOTRTRN3 As DataRow
        Dim DIST_AMT As Decimal = 0
        Dim DTL_MISC_CHG_CODE As String = ""

        For Each rowSOTRTRN2 As DataRow In frm.dst.Tables("SOTRTRN2").Select("", "", DataViewRowState.CurrentRows) ' ISNULL(RTRN_QTY_REF,0) <> RTRN_QTY
            Dim RTRN_QTY As Int32 = Val(rowSOTRTRN2.Item("RTRN_QTY") & "")
            Dim RTRN_QTY_REF As Int32 = 0 ' Val(rowSOTRTRN2.Item("RTRN_QTY_REF") & "") - THIS WOULD BE FOR A QTY REFUSED FOR CREDIT - TO BE RETURNED TO THE CUSTOMER

            Dim DIST_AMT_SALES As Decimal = (RTRN_QTY - RTRN_QTY_REF) * Val(rowSOTRTRN2.Item("RTRN_PRICE") & "")
            Dim DIST_AMT_COSTS As Decimal = (RTRN_QTY - RTRN_QTY_REF) * Val(rowSOTRTRN2.Item("STYLE_COST") & "")

            Dim rowICTCLAS1 As DataRow = frm.dst.Tables("ICTCLAS1").Rows.Find(New String() {rowSOTRTRN2.Item("STYLE_CLASS_CODE")})

            For RTRN_GNO = 1 To 4
                If ((RTRN_GNO = 1 Or RTRN_GNO = 1) And DIST_AMT_SALES <> 0) _
                Or ((RTRN_GNO = 3 Or RTRN_GNO = 4) And DIST_AMT_COSTS <> 0) Then

                    ' DON'T DO AR HERE - DO IT ONCE FOR THE DOCUMENT

                    rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow

                    rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN2.Item("RTRN_NO")
                    rowSOTRTRN3.Item("RTRN_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
                    rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO

                    rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")

                    If RTRN_GNO = 1 Then ' Sales Returns
                        rowSOTRTRN3.Item("ACCT_CODE") = rowICTCLAS1.Item("ACCT_CODE_SLS_RTN")
                        rowSOTRTRN3.Item("DIST_TYPE") = "SLSRTN"
                        DIST_AMT = DIST_AMT_SALES
                    ElseIf RTRN_GNO = 2 Then ' Accts Receivable
                        rowSOTRTRN3.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
                        'rowSOTRTRN3.Item("SEG2_CODE") = rowARTPOST1.Item("SEG2_CODE")
                        'rowSOTRTRN3.Item("SEG3_CODE") = rowARTPOST1.Item("SEG3_CODE")
                        'rowSOTRTRN3.Item("SEG4_CODE") = rowARTPOST1.Item("SEG4_CODE")
                        rowSOTRTRN3.Item("DIST_TYPE") = "AR"
                        DIST_AMT = -1 * DIST_AMT_SALES
                    ElseIf RTRN_GNO = 3 Then ' Inventory
                        rowSOTRTRN3.Item("ACCT_CODE") = rowICTCLAS1.Item("ACCT_CODE_ONH")
                        rowSOTRTRN3.Item("DIST_TYPE") = "INVTY"
                        DIST_AMT = DIST_AMT_COSTS
                    ElseIf RTRN_GNO = 4 Then  ' Cost of Goods Returned
                        rowSOTRTRN3.Item("ACCT_CODE") = rowICTCLAS1.Item("ACCT_CODE_CGS_RTN")
                        rowSOTRTRN3.Item("DIST_TYPE") = "CGR"
                        DIST_AMT = -1 * DIST_AMT_COSTS
                    End If
                    rowSOTRTRN3.Item("DIST_AMT") = Round(DIST_AMT, 2)
                    frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)

                End If
            Next

            Dim rowSOTINVH2 As DataRow = frm.dst.Tables("SOTINVH2").NewRow
            rowSOTINVH2.Item("INV_TYPE") = "C"
            rowSOTINVH2.Item("INV_NO") = INV_NO
            rowSOTINVH2.Item("INV_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
            rowSOTINVH2.Item("STYLE_CODE") = rowSOTRTRN2.Item("STYLE_CODE")
            rowSOTINVH2.Item("COLOR_CODE") = rowSOTRTRN2.Item("COLOR_CODE")
            rowSOTINVH2.Item("ORDR_UNIT_PRICE") = rowSOTRTRN2.Item("RTRN_PRICE")
            rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") = rowSOTRTRN2.Item("RTRN_PRICE_CURR")
            rowSOTINVH2.Item("ORDR_QTY_SHIP") = -1 * (RTRN_QTY - RTRN_QTY_REF)
            rowSOTINVH2.Item("CUST_CODE") = rowSOTRTRN1.Item("CUST_CODE")
            rowSOTINVH2.Item("ORDR_YYYYPP_UPDATED") = rowSOTRTRN2.Item("OPS_YYYYPP")
            rowSOTINVH2.Item("ORDR_UNIT_COST") = rowSOTRTRN2.Item("STYLE_COST")
            frm.dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)

            If ASCMAIN1.CLIENT = "RGI" And Not IsDBNull(rowSOTRTRN2("MISC_CHG_CODE")) Then
                DTL_MISC_CHG_CODE = rowSOTRTRN2("MISC_CHG_CODE")
                Dim rowSOTINVHM As DataRow = frm.dst.Tables("SOTINVHM").NewRow
                With rowSOTINVHM
                    .Item("INV_TYPE") = "C"
                    .Item("INV_NO") = INV_NO
                    .Item("INV_MNO") = rowSOTRTRN2.Item("RTRN_LNO")
                    .Item("INV_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
                    .Item("MISC_CHG_CODE") = DTL_MISC_CHG_CODE
                    .Item("MISC_CHG_DESC") = DTL_MISC_CHG_CODE
                    .Item("MISC_CHG_NOTE") = "Returns Handling"
                    .Item("INV_MISC_CHG") = -1 * rowSOTRTRN2("LINE_TARIFF")
                    .Item("INV_MISC_CHG_CURR") = -1 * rowSOTRTRN2("LINE_TARIFF")
                    .Item("SURCHARGE_PERC") = rowSOTRTRN2("SURCHARGE_PERC")
                    .Item("MISC_CHARGE_TYPE") = "T"
                    .Item("COUNTRY_CODE") = rowSOTRTRN2("COUNTRY_CODE")
                End With
                frm.dst.Tables("SOTINVHM").Rows.Add(rowSOTINVHM)
            End If
        Next

        Dim rowSOTINVH1 As DataRow = frm.dst.Tables("SOTINVH1").NewRow
        rowSOTINVH1.Item("INV_TYPE") = "C"
        rowSOTINVH1.Item("INV_NO") = INV_NO
        rowSOTINVH1.Item("CUST_CODE") = rowSOTRTRN1.Item("CUST_CODE")
        rowSOTINVH1.Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO")
        rowSOTINVH1.Item("ORDR_CUST_PO") = rowSOTRTRN1.Item("CUST_CLAIM_NO")
        rowSOTINVH1.Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
        rowSOTINVH1.Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE")
        rowSOTINVH1.Item("TERM_CODE") = frm.ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0")
        rowSOTINVH1.Item("REASON_CODE") = rowSOTRTRN1.Item("REASON_CODE")
        rowSOTINVH1.Item("CUST_BILL_TO_CUST") = rowSOTRTRN1.Item("CUST_BILL_TO_CUST")
        rowSOTINVH1.Item("SREP_CODE") = rowSOTRTRN1.Item("SREP_CODE")
        rowSOTINVH1.Item("INV_SALES") = -1 * Val(rowSOTRTRN1.Item("RTRN_SALES") & "")
        rowSOTINVH1.Item("INV_COGS") = -1 * Val(rowSOTRTRN1.Item("RTRN_COSTS") & "")
        rowSOTINVH1.Item("INV_FREIGHT") = -1 * Val(rowSOTRTRN1.Item("RTRN_FREIGHT") & "")
        If Val(rowSOTRTRN1.Item("RTRN_HANDLING") & "") <> 0 Then
            rowSOTINVH1.Item("INV_MISC_CHG") = -1 * Val(rowSOTRTRN1.Item("RTRN_HANDLING") & "")
        End If
        rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = -1 * Val(rowSOTRTRN1.Item("RTRN_AMOUNT") & "")
        rowSOTINVH1.Item("INV_DATE") = rowSOTRTRN1.Item("RTRN_DATE")
        rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = rowSOTRTRN1.Item("OPS_YYYYPP")
        rowSOTINVH1.Item("INIT_DATE") = frm.DATETIME_STAMP
        rowSOTINVH1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSOTINVH1.Item("ORDR_TYPE_CODE") = "RTN"
        rowSOTINVH1.Item("INV_COMMENT") = rowSOTRTRN1.Item("RTRN_NOTE")

        rowSOTINVH1.Item("CURR_CODE") = rowSOTRTRN1.Item("CURR_CODE")
        rowSOTINVH1.Item("CURR_EXCH_RATE") = rowSOTRTRN1.Item("CURR_EXCH_RATE")
        rowSOTINVH1.Item("INV_SALES_CURR") = -1 * Val(rowSOTRTRN1.Item("RTRN_SALES_CURR") & "")
        rowSOTINVH1.Item("INV_STAX_CURR") = -1 * Val(rowSOTRTRN1.Item("RTRN_STAX_CURR") & "")
        rowSOTINVH1.Item("INV_FREIGHT_CURR") = -1 * Val(rowSOTRTRN1.Item("RTRN_FREIGHT_CURR") & "")
        rowSOTINVH1.Item("INV_MISC_CHG_CURR") = -1 * Val(rowSOTRTRN1.Item("RTRN_HANDLING_CURR") & "")
        rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR") = -1 * Val(rowSOTRTRN1.Item("RTRN_AMOUNT_CURR") & "")
        rowSOTINVH1.Item("INV_TOTAL_AMT_CURR") = -1 * Val(rowSOTRTRN1.Item("RTRN_AMOUNT_CURR") & "")
        rowSOTINVH1.Item("SALES_DIVISION_CODE") = ""

        frm.dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

        ' RGI has tariffs for individual items as MISC charges
        If ASCMAIN1.CLIENT <> "RGI" Then
            Dim INV_MISC_CHG As Decimal = Val(rowSOTINVH1.Item("INV_MISC_CHG") & "")
            If INV_MISC_CHG <> 0 Then
                Dim MISC_CHG_CODE = frm.ROWs("SOTPARM1").Item("SO_PARM_MISC_CHG_RTN")
                Dim rowSOTMISC1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTMISC1 WHERE MISC_CHG_CODE = '" & MISC_CHG_CODE & "'")
                Dim rowSOTINVHM As DataRow = frm.dst.Tables("SOTINVHM").NewRow
                With rowSOTINVHM
                    .Item("INV_TYPE") = "C"
                    .Item("INV_NO") = INV_NO
                    .Item("INV_MNO") = 1
                    .Item("MISC_CHG_CODE") = MISC_CHG_CODE
                    If rowSOTMISC1 IsNot Nothing Then
                        .Item("MISC_CHG_DESC") = rowSOTMISC1.Item("MISC_CHG_DESC")
                    End If
                    .Item("MISC_CHG_NOTE") = "Returns Handling"
                    .Item("INV_MISC_CHG") = INV_MISC_CHG
                End With
                frm.dst.Tables("SOTINVHM").Rows.Add(rowSOTINVHM)
            End If
        End If

        Dim rowARTOPEN1 As DataRow = frm.dst.Tables("ARTOPEN1").NewRow
        ' "STAX_CODE","INV_STAX",
        For Each C As String In New String() _
        {"CUST_CODE", "INV_TYPE", "INV_DATE", "CUST_STORE_NO", "POST_CODE", _
         "TERM_CODE", "SREP_CODE", _
         "ORDR_NO", "INV_SALES", "INV_FREIGHT", "INV_TOTAL_AMOUNT", _
         "REASON_CODE", "INIT_OPER", "INIT_DATE", "INV_MISC_CHG", "ORDR_TYPE_CODE", "SALES_DIVISION_CODE"}
            rowARTOPEN1.Item(C) = rowSOTINVH1.Item(C)
        Next

        rowARTOPEN1.Item("INV_TYPE") = "R"
        rowARTOPEN1.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")

        rowARTOPEN1.Item("INV_DUE_DATE") = rowSOTINVH1.Item("INV_DATE")
        rowARTOPEN1.Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
        rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("CUST_CODE_SO") = rowSOTINVH1.Item("CUST_CODE")
        rowARTOPEN1.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowARTOPEN1.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowARTOPEN1.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")

        rowARTOPEN1.Item("CURR_CODE") = rowSOTINVH1.Item("CURR_CODE")
        rowARTOPEN1.Item("CURR_EXCH_RATE") = rowSOTINVH1.Item("CURR_EXCH_RATE")
        rowARTOPEN1.Item("INV_SALES_CURR") = rowSOTINVH1.Item("INV_SALES_CURR")
        rowARTOPEN1.Item("INV_DISC_CURR") = rowARTOPEN1.Item("INV_DISC")
        rowARTOPEN1.Item("INV_FREIGHT_CURR") = rowSOTINVH1.Item("INV_FREIGHT_CURR")
        rowARTOPEN1.Item("INV_STAX_CURR") = rowSOTINVH1.Item("INV_STAX_CURR")
        rowARTOPEN1.Item("INV_MISC_CHG_CURR") = rowSOTINVH1.Item("INV_MISC_CHG_CURR")
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR")
        rowARTOPEN1.Item("INV_BALANCE_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR")

        rowARTOPEN1.Item("OPS_YYYYPP") = rowSOTRTRN1.Item("OPS_YYYYPP")
        rowARTOPEN1.Item("INV_NOTES") = rowSOTRTRN1.Item("RTRN_NOTE")
        frm.dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        frm.Update_Record_TDA("ARTOPEN1")
        frm.Update_Record_TDA("SOTINVH1")
        frm.Update_Record_TDA("SOTINVH2")
        frm.Update_Record_TDA("SOTINVHM")
        frm.Update_Record_TDA("SOTRTRN1")
        frm.Update_Record_TDA("SOTRTRN2")

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
        & " Select SOTRTRN1.OPS_YYYYPP, SOTRTRN1.WHSE_CODE" _
        & ", SOTRTRN2.STYLE_CODE, SOTRTRN2.COLOR_CODE, SOTRTRN2.RTRN_QTY, 0 RTRN_QTY_REF " _
        & " from SOTRTRN2,SOTRTRN1 where SOTRTRN2.RTRN_NO = :PARM1" _
        & " and SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " UPDATE ICTSTAT2 " _
        & " SET WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.RTRN_QTY,0) - NVL(R1.RTRN_QTY_REF,0)" _
        & " WHERE STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT2 (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_ON_HAND)" _
        & " VALUES (R1.STYLE_CODE, R1.COLOR_CODE, R1.WHSE_CODE, R1.RTRN_QTY-NVL(R1.RTRN_QTY_REF,0));" _
        & " END IF;" _
        & " UPDATE ICTSTAT1 " _
        & " SET WHSE_QTY_RTN = NVL(WHSE_QTY_RTN,0) + NVL(R1.RTRN_QTY,0) - NVL(R1.RTRN_QTY_REF,0)" _
        & " WHERE STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE AND WHSE_CODE = R1.WHSE_CODE" _
        & " AND OPS_YYYYPP = R1.OPS_YYYYPP;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT1 (OPS_YYYYPP, STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_RTN)" _
        & " VALUES (R1.OPS_YYYYPP, R1.STYLE_CODE, R1.COLOR_CODE, R1.WHSE_CODE, R1.RTRN_QTY - NVL(R1.RTRN_QTY_REF,0));" _
        & " END IF;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {rowSOTRTRN1.Item("RTRN_NO")})


        RTRN_GNO = 0

        DIST_AMT = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & "") ' AR
        If DIST_AMT <> 0 Then
            rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
            rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
            rowSOTRTRN3.Item("RTRN_LNO") = 0
            RTRN_GNO += 1
            rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO
            rowSOTRTRN3.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
            rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowSOTRTRN3.Item("DIST_TYPE") = "AR"
            rowSOTRTRN3.Item("DIST_AMT") = Round(DIST_AMT, 2)
            frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        End If

        'DIST_AMT = -1 * Val(rowSOTINVH1.Item("INV_STAX") & "") ' Sales Tax
        'If DIST_AMT <> 0 Then
        '    rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
        '    rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
        '    rowSOTRTRN3.Item("RTRN_LNO") = 0
        '    RTRN_GNO += 1
        '    rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO
        '    rowSOTRTRN3.Item("ACCT_CODE") = rowARTSTAX1.Item("ACCT_CODE")
        '    rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        '    rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        '    rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        '    rowSOTRTRN3.Item("DIST_TYPE") = "STAX"
        '    rowSOTRTRN3.Item("DIST_AMT") = Round(DIST_AMT, 2)
        '    frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        'End If

        DIST_AMT = -1 * Val(rowSOTINVH1.Item("INV_FREIGHT") & "") ' Freight
        If DIST_AMT <> 0 Then
            rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
            rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
            rowSOTRTRN3.Item("RTRN_LNO") = 0
            RTRN_GNO += 1
            rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO
            rowSOTRTRN3.Item("ACCT_CODE") = frm.ROWs.Item("SOTPARM1").Item("SO_PARM_ACCT_FRT_INC") & ""
            rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowSOTRTRN3.Item("DIST_TYPE") = "FRT"
            rowSOTRTRN3.Item("DIST_AMT") = Round(DIST_AMT, 2)
            frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        End If

        'SHOULD WE DO RSF AS AN ACCOUNT IN SOTPARM1 OR AS A MISC CHG CODE?
        ' CODE BELOW IS USING SOTPARM1, BUT ODG USES SOTMISC1 - ROWSOTMISC1 IS ALREADY SET UP ABOVE TO HANDLE THIS
        ' - LET'S WAIT UNTIL SOMEONE ASKS FOR IT
        DIST_AMT = -1 * Val(rowSOTINVH1.Item("INV_MISC_CHG") & "") ' Handling / RGI - Tariff
        If DIST_AMT <> 0 Then
            rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
            rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
            rowSOTRTRN3.Item("RTRN_LNO") = 0
            RTRN_GNO += 1
            rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO

            Dim MISC_CHG_CODE = frm.ROWs("SOTPARM1").Item("SO_PARM_MISC_CHG_RTN")
            If DTL_MISC_CHG_CODE <> "" Then
                MISC_CHG_CODE = DTL_MISC_CHG_CODE
            End If
            Dim rowSOTMISC1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTMISC1 WHERE MISC_CHG_CODE = '" & MISC_CHG_CODE & "'")

            rowSOTRTRN3.Item("ACCT_CODE") = rowSOTMISC1.Item("ACCT_CODE")
            rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowSOTRTRN3.Item("DIST_TYPE") = "HND"
            rowSOTRTRN3.Item("DIST_AMT") = Round(DIST_AMT, 2)
            frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        End If

        Call frm.Update_Record_TDA("SOTRTRN3")

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Shuttle_RTN_to_ICTTRAN1(frm)
        End If

        Return INV_NO
    End Function

    Public Shared Sub Shuttle_RTN_to_ICTTRAN1(ByVal frm As ASFBASE0)

        If frm.dst.Tables.Contains("ICTTRAN1") Then
            frm.dst.Tables("ICTTRAN1").Rows.Clear()
            frm.dst.Tables("ICTTRAN2").Rows.Clear()
        Else
            frm.Create_TDA(frm.dst.Tables.Add, "ICTTRAN1", "*")
            frm.Create_TDA(frm.dst.Tables.Add, "ICTTRAN2", "*")
        End If

        Dim rowSOTRTRN1 As DataRow = frm.dst.Tables("SOTRTRN1").Rows(0)
        Dim rowICTTRAN1 As DataRow = frm.dst.Tables("ICTTRAN1").NewRow
        With rowICTTRAN1
            .Item("OPS_YYYYPP") = rowSOTRTRN1.Item("OPS_YYYYPP")
            .Item("TRAN_TYPE") = "C"
            .Item("TRAN_NO") = Mid(rowSOTRTRN1.Item("RTRN_NO"), 5)
            .Item("TRAN_SOURCE_DOCUMENT") = rowSOTRTRN1.Item("RTRN_SOURCE_DOC_NO")
            .Item("TRAN_DATE") = rowSOTRTRN1.Item("RTRN_DATE")
            .Item("TRAN_WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
            .Item("TRAN_CUST_CODE") = rowSOTRTRN1.Item("CUST_CODE")
            .Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO")
            .Item("INIT_DATE") = rowSOTRTRN1.Item("INIT_DATE")
            .Item("INIT_OPER") = rowSOTRTRN1.Item("INIT_OPER")
            .Item("TRAN_STATUS_UPD") = "U"
            Dim rowARTCUST1 As DataRow = frm.LookUp("ARTCUST1", rowSOTRTRN1.Item("CUST_CODE"))
            .Item("TRAN_CCVRW_DESC") = rowSOTRTRN1.Item("CUST_NAME")
            .Item("TRAN_CCVRW_REF") = rowSOTRTRN1.Item("CUST_CLAIM_NO")
            .Item("TRAN_ORIGINATE") = "E"
            .Item("TRAN_COMMENT") = rowSOTRTRN1.Item("RTRN_NOTE")
            .Item("TRAN_FREIGHT") = rowSOTRTRN1.Item("RTRN_FREIGHT")
            .Item("TRAN_MISC_CHG") = rowSOTRTRN1.Item("RTRN_HANDLING")
            .Item("REASON_CODE") = rowSOTRTRN1.Item("REASON_CODE")
            .Item("SREP_CODE") = rowSOTRTRN1.Item("SREP_CODE")
            .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE")
        End With
        frm.dst.Tables("ICTTRAN1").Rows.Add(rowICTTRAN1)
        frm.Update_Record_TDA("ICTTRAN1")

        For Each rowSOTRTRN2 As DataRow In frm.dst.Tables("SOTRTRN2").Select("")
            Dim rowICTTRAN2 As DataRow = frm.dst.Tables("ICTTRAN2").NewRow
            With rowICTTRAN2
                .Item("OPS_YYYYPP") = rowSOTRTRN1.Item("OPS_YYYYPP")
                .Item("TRAN_TYPE") = "C"
                .Item("TRAN_NO") = Mid(rowSOTRTRN2.Item("RTRN_NO"), 5)
                .Item("TRAN_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
                .Item("STYLE_CODE") = rowSOTRTRN2.Item("STYLE_CODE")
                .Item("COLOR_CODE") = rowSOTRTRN2.Item("COLOR_CODE")
                Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", rowSOTRTRN2.Item("STYLE_CODE"))
                .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
                .Item("STYLE_COST") = rowSOTRTRN2.Item("STYLE_COST")
                .Item("TRAN_QTY") = rowSOTRTRN2.Item("RTRN_QTY")
                .Item("TRAN_QTY1") = 0 ' rowSOTRTRN2.Item("RTRN_QTY_1")
                .Item("TRAN_QTY2") = rowSOTRTRN2.Item("RTRN_QTY") ' rowSOTRTRN2.Item("RTRN_QTY_2")
                .Item("TRAN_QTY3") = 0 ' rowSOTRTRN2.Item("RTRN_QTY_3")
            End With
            frm.dst.Tables("ICTTRAN2").Rows.Add(rowICTTRAN2)
        Next
        frm.Update_Record_TDA("ICTTRAN2")

        frm.dst.Tables("ICTTRAN1").Rows.Clear()
        frm.dst.Tables("ICTTRAN2").Rows.Clear()
    End Sub

    Public Shared Sub Update_RTV(ByVal frm As ASFBASE1, ByVal rowICTIRTV1 As DataRow)

        Dim IC_PARM_ACCT_CODE_RTV_CLEARING As String = frm.ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_RTV_CLEARING")

        For Each rowICTIRTV2 As DataRow In frm.dst.Tables("ICTIRTV2").Select("", "", DataViewRowState.CurrentRows)
            Dim STANDARD_COST_TOTAL As Decimal = Val(rowICTIRTV2.Item("RTV_QTY") & "") * Val(rowICTIRTV2.Item("PRICE_CATGY_COST_TOTAL") & "")
            Dim PO_COST_TOTAL As Decimal = Val(rowICTIRTV2.Item("RTV_QTY") & "") * Val(rowICTIRTV2.Item("PO_COST") & "")
            Dim PURCHASE_PRICE_VARIANCE As Decimal = STANDARD_COST_TOTAL - PO_COST_TOTAL
            Dim rowICTCATG1 As DataRow = frm.dst.Tables("ICTCATG1").Rows.Find(New String() {rowICTIRTV2.Item("PROD_CATGY_CODE")})
            For RTV_GNO As Integer = 1 To 3
                Dim rowICTIRTV3 As DataRow = frm.dst.Tables("ICTIRTV3").NewRow
                rowICTIRTV3.Item("RTV_NO") = rowICTIRTV2.Item("RTV_NO")
                rowICTIRTV3.Item("RTV_LNO") = rowICTIRTV2.Item("RTV_LNO")
                rowICTIRTV3.Item("RTV_GNO") = RTV_GNO
                If RTV_GNO = 1 Then ' Inventory
                    rowICTIRTV3.Item("ACCT_CODE") = rowICTCATG1.Item("ACCT_CODE_INV")
                    rowICTIRTV3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowICTIRTV3.Item("DIST_TYPE") = "INVTY"
                    rowICTIRTV3.Item("DIST_AMT") = Round(STANDARD_COST_TOTAL * -1, 2)
                ElseIf RTV_GNO = 2 Then ' RTV Clearing
                    rowICTIRTV3.Item("ACCT_CODE") = IC_PARM_ACCT_CODE_RTV_CLEARING
                    rowICTIRTV3.Item("SEG2_CODE") = frm.dst.Tables("ICTWHSE1").Rows(0).Item("SEG2_CODE")
                    rowICTIRTV3.Item("DIST_TYPE") = "RTV"
                    rowICTIRTV3.Item("DIST_AMT") = Round(PO_COST_TOTAL, 2)
                Else 'Purchase Price Variance
                    rowICTIRTV3.Item("ACCT_CODE") = rowICTCATG1.Item("ACCT_CODE_PPV")
                    rowICTIRTV3.Item("SEG2_CODE") = frm.dst.Tables("ICTWHSE1").Rows(0).Item("SEG2_CODE")
                    rowICTIRTV3.Item("DIST_TYPE") = "PPV"
                    rowICTIRTV3.Item("DIST_AMT") = Round(PURCHASE_PRICE_VARIANCE, 2)
                End If
                'rowICTIRTV3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                rowICTIRTV3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                rowICTIRTV3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                frm.dst.Tables("ICTIRTV3").Rows.Add(rowICTIRTV3)
                rowICTIRTV2.Item("OPS_YYYYPP") = rowICTIRTV1.Item("OPS_YYYYPP")
                If PURCHASE_PRICE_VARIANCE = 0 And RTV_GNO = 2 Then
                    Exit For 'No entry for PPV if there is no variance
                End If
            Next
        Next

        Call frm.Update_Record_TDA("ICTIRTV1")
        Call frm.Update_Record_TDA("ICTIRTV2")
        Call frm.Update_Record_TDA("ICTIRTV3")



        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
        & " Select ICTIRTV1.OPS_YYYYPP, ICTIRTV1.WHSE_CODE" _
        & ", ICTIRTV2.ITEM_CODE, ICTIRTV2.RTV_QTY " _
        & " from ICTIRTV2,ICTIRTV1 where ICTIRTV2.RTV_NO = :PARM1" _
        & " and ICTIRTV1.RTV_NO = ICTIRTV2.RTV_NO;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " UPDATE ICTSTAT2 " _
        & " SET WHSE_QTY_HOLD = NVL(WHSE_QTY_HOLD,0) + NVL(R1.RTV_QTY,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_ON_HAND, WHSE_QTY_HOLD)" _
        & " VALUES (R1.ITEM_CODE, R1.WHSE_CODE, 0, R1.RTV_QTY);" _
        & " END IF;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {rowICTIRTV1.Item("RTV_NO")})

    End Sub

    Public Shared Function Finalize_RTV(ByVal RTV_NO As String, ByVal RANo As String, ByVal TrackingNo As String) As String

        Dim voucherNo As String = ASCDATA1.ExecuteSF("ICPIRTVF", New String() {"RTV_NO", "USER_ID", "RTV_RA_NO", "RTV_TRACKING_NO"}, New Object() {RTV_NO, ASCMAIN1.USER_ID, RANo, TrackingNo})

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
        & " Select ICTIRTV1.OPS_YYYYPP, ICTIRTV1.WHSE_CODE" _
        & ", ICTIRTV2.ITEM_CODE, ICTIRTV2.RTV_QTY " _
        & " from ICTIRTV2,ICTIRTV1 where ICTIRTV2.RTV_NO = :PARM1" _
        & " and ICTIRTV1.RTV_NO = ICTIRTV2.RTV_NO;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " UPDATE ICTSTAT2 " _
        & " SET WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) - NVL(R1.RTV_QTY,0)," _
        & " WHSE_QTY_HOLD = NVL(WHSE_QTY_HOLD,0) - NVL(R1.RTV_QTY,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_ON_HAND)" _
        & " VALUES (R1.ITEM_CODE, R1.WHSE_CODE, -1 * R1.RTV_QTY);" _
        & " END IF;" _
        & " UPDATE ICTSTAT1 " _
        & " SET WHSE_QTY_RTV = NVL(WHSE_QTY_RTV,0) + NVL(R1.RTV_QTY,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE" _
        & " AND OPS_YYYYPP = R1.OPS_YYYYPP;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT1 (OPS_YYYYPP, ITEM_CODE, WHSE_CODE, WHSE_QTY_RTV)" _
        & " VALUES (R1.OPS_YYYYPP, R1.ITEM_CODE, R1.WHSE_CODE, R1.RTV_QTY);" _
        & " END IF;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {RTV_NO})

        Return voucherNo
    End Function

    Public Shared Sub Update_Adjustment(ByVal frm As ASFBASE0)

        frm.Update_Record_TDA("ICTIADJ1")
        frm.Update_Record_TDA("ICTIADJ2")

        For Each rowICTIADJ1 As DataRow In frm.dst.Tables("ICTIADJ1").Select("")
            Dim ADJ_NO_in As String = rowICTIADJ1.Item("ADJ_NO")
            ASCDATA1.ExecuteSP("ICPIADJI", "VN", New Object() {ADJ_NO_in, 1}, New String() {"ADJ_NO_in", "S"})
            ASCDATA1.ExecuteSP("ICPIADJG", "V", New Object() {ADJ_NO_in}, New String() {"ADJ_NO_in"})
        Next

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Shuttle_ADJ_to_ICTTRAN1(frm)
        End If
    End Sub

    Public Shared Sub Shuttle_ADJ_to_ICTTRAN1(ByVal frm As ASFBASE0)

        If frm.dst.Tables.Contains("ICTTRAN1") Then
            frm.dst.Tables("ICTTRAN1").Rows.Clear()
            frm.dst.Tables("ICTTRAN2").Rows.Clear()
        Else
            frm.Create_TDA(frm.dst.Tables.Add, "ICTTRAN1", "*")
            frm.Create_TDA(frm.dst.Tables.Add, "ICTTRAN2", "*")
        End If

        For Each rowICTIADJ1 As DataRow In frm.dst.Tables("ICTIADJ1").Select("")
            Dim rowICTTRAN1 As DataRow = frm.dst.Tables("ICTTRAN1").NewRow
            With rowICTTRAN1
                .Item("OPS_YYYYPP") = rowICTIADJ1.Item("OPS_YYYYPP")
                .Item("TRAN_TYPE") = "A"
                .Item("TRAN_NO") = rowICTIADJ1.Item("ADJ_NO")
                .Item("TRAN_SOURCE_DOCUMENT") = ""
                .Item("TRAN_DATE") = rowICTIADJ1.Item("ADJ_DATE")
                .Item("TRAN_WHSE_CODE") = rowICTIADJ1.Item("WHSE_CODE")
                .Item("TRAN_ADJ_REASON_CODE") = rowICTIADJ1.Item("REASON_CODE")
                .Item("INIT_DATE") = rowICTIADJ1.Item("INIT_DATE")
                .Item("INIT_OPER") = rowICTIADJ1.Item("INIT_OPER")
                .Item("TRAN_STATUS_UPD") = "U"
                Dim rowICTWHSE1 As DataRow = frm.LookUp("ICTWHSE1", rowICTIADJ1.Item("WHSE_CODE"))
                .Item("TRAN_CCVRW_DESC") = rowICTWHSE1.Item("WHSE_DESC")
                .Item("TRAN_COMMENT") = rowICTIADJ1.Item("ADJ_NOTE")
                .Item("TRAN_ORIGINATE") = rowICTIADJ1.Item("ADJ_SOURCE")
            End With
            frm.dst.Tables("ICTTRAN1").Rows.Add(rowICTTRAN1)

            For Each rowICTIADJ2 As DataRow In frm.dst.Tables("ICTIADJ2").Select("ADJ_NO = '" & rowICTIADJ1.Item("ADJ_NO") & "'")
                Dim rowICTTRAN2 As DataRow = frm.dst.Tables("ICTTRAN2").NewRow
                With rowICTTRAN2
                    .Item("OPS_YYYYPP") = rowICTIADJ1.Item("OPS_YYYYPP")
                    .Item("TRAN_TYPE") = "A"
                    .Item("TRAN_NO") = rowICTIADJ2.Item("ADJ_NO")
                    .Item("TRAN_LNO") = rowICTIADJ2.Item("ADJ_LNO")
                    .Item("STYLE_CODE") = rowICTIADJ2.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = rowICTIADJ2.Item("COLOR_CODE")
                    Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", rowICTIADJ2.Item("STYLE_CODE"))
                    .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                    .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
                    .Item("STYLE_COST") = rowICTIADJ2.Item("STYLE_COST")
                    .Item("TRAN_QTY") = rowICTIADJ2.Item("ADJ_QTY")
                    .Item("TRAN_QTY2") = rowICTIADJ2.Item("ADJ_QTY")
                End With
                frm.dst.Tables("ICTTRAN2").Rows.Add(rowICTTRAN2)
            Next
        Next

        frm.Update_Record_TDA("ICTTRAN1")
        frm.Update_Record_TDA("ICTTRAN2")

        frm.dst.Tables("ICTTRAN1").Rows.Clear()
        frm.dst.Tables("ICTTRAN2").Rows.Clear()
    End Sub

    Public Shared Sub Shuttle_ADJ_to_ICTTRAN1_SQL(ADJ_NO As String)

        ASCMAIN1.sql = "INSERT INTO ICTTRAN1 (" & vbCrLf _
            & "OPS_YYYYPP,TRAN_TYPE,TRAN_NO,TRAN_DATE," & vbCrLf _
            & "TRAN_WHSE_CODE,TRAN_ADJ_REASON_CODE," & vbCrLf _
            & "INIT_DATE,INIT_OPER,TRAN_STATUS_UPD, " & vbCrLf _
            & "TRAN_CCVRW_DESC,TRAN_ORIGINATE,TRAN_COMMENT)" & vbCrLf _
            & "SELECT ICTIADJ1.OPS_YYYYPP, 'A', ICTIADJ1.ADJ_NO, ICTIADJ1.ADJ_DATE" & vbCrLf _
            & ", ICTIADJ1.WHSE_CODE, ICTIADJ1.REASON_CODE, ICTIADJ1.INIT_DATE, ICTIADJ1.INIT_OPER" & vbCrLf _
            & ", 'U', ICTWHSE1.WHSE_DESC, ICTIADJ1.ADJ_SOURCE, ICTIADJ1.ADJ_NOTE" & vbCrLf _
            & "FROM ICTIADJ1,ICTWHSE1 WHERE ICTWHSE1.WHSE_CODE = ICTIADJ1.WHSE_CODE" & vbCrLf _
            & "AND ICTIADJ1.ADJ_NO = '" & ADJ_NO & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "INSERT INTO ICTTRAN2 (" & vbCrLf _
            & "OPS_YYYYPP,TRAN_TYPE,TRAN_NO,TRAN_LNO," & vbCrLf _
            & "STYLE_CODE,COLOR_CODE,STYLE_DESC,STYLE_UOM,STYLE_COST,TRAN_QTY)" & vbCrLf _
            & "SELECT ICTIADJ2.OPS_YYYYPP, 'A', ICTIADJ2.ADJ_NO, ICTIADJ2.ADJ_LNO" & vbCrLf _
            & ", ICTIADJ2.STYLE_CODE, ICTIADJ2.COLOR_CODE" & vbCrLf _
            & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM" & vbCrLf _
            & ", ICTIADJ2.STYLE_COST, ICTIADJ2.ADJ_QTY" & vbCrLf _
            & "FROM ICTIADJ2,ICTSTYL1 WHERE ICTSTYL1.STYLE_CODE = ICTIADJ2.STYLE_CODE" & vbCrLf _
            & "AND ICTIADJ2.ADJ_NO = '" & ADJ_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Shared Sub Update_Transfer(ByVal frm As ASFBASE0)

        Dim rowICTIXFR1 As DataRow = frm.dst.Tables("ICTIXFR1").Rows(0)
        frm.Update_Record_TDA("ICTIXFR1")
        frm.Update_Record_TDA("ICTIXFR2")
        Dim XFR_NO_in As String = rowICTIXFR1.Item("XFR_NO")
        ASCDATA1.ExecuteSP("ICPIXFRI", "VN", New Object() {XFR_NO_in, 1}, New String() {"XFR_NO_in", "S"})
        ASCDATA1.ExecuteSP("ICPIXFRG", "V", New Object() {XFR_NO_in}, New String() {"XFR_NO_in"})

    End Sub

    Public Shared Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String, ByVal TT As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Double
        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))

        Dim YP As String = ""
        Dim TX As String = ""
        Dim TK As String = ""

        Select Case JOURNAL_TYPE

            Case "ICIA"
                TX = "ICTIADJ3"
                TK = "ADJ_NO"

            Case "ICIT"
                TX = "ICTIXFR3"
                TK = "XFR_NO"

            Case "ICIR"
                TX = "ICTIREC5"
                TK = "RECEIPT_NO"

            Case "ICIV"
                TX = "ICTIRTV3"
                TK = "RTV_NO"

        End Select

        ASCMAIN1.sql = "" _
        & " SELECT T1.OPS_YYYYPP, TX.ACCT_CODE" _
        & ", TX.SEG2_CODE, TX.SEG3_CODE, TX.SEG4_CODE" _
        & ", TX.DIST_TYPE, SUM (TX.DIST_AMT) DIST_AMT " _
        & " FROM " & TX & " TX," & TT & " T1" _
        & " where TX." & TK & " = T1." & TK _
        & " GROUP BY T1.OPS_YYYYPP, TX.ACCT_CODE, TX.SEG2_CODE, TX.SEG3_CODE, TX.SEG4_CODE, TX.DIST_TYPE " _
        & " ORDER BY T1.OPS_YYYYPP, TX.ACCT_CODE, TX.SEG2_CODE, TX.SEG3_CODE, TX.SEG4_CODE, TX.DIST_TYPE "

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            DETL_POSTING_AMT = Val(row.Item("DIST_AMT") & "")
            Dim rowGLTINTF1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").NewRow
            rowGLTINTF1("OPS_YYYYPP") = row("OPS_YYYYPP")
            rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
            JOURNAL_LNO += 1
            rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
            rowGLTINTF1("ACCT_CODE") = row("ACCT_CODE")
            rowGLTINTF1("SEG2_CODE") = row("SEG2_CODE")
            rowGLTINTF1("SEG3_CODE") = row("SEG3_CODE")
            rowGLTINTF1("SEG4_CODE") = row("SEG4_CODE")
            rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
            rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
            rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
            rowGLTINTF1("DETL_CTL_NO") = DBNull.Value
            rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
            rowGLTINTF1("DETL_CVX_NO") = DBNull.Value
            rowGLTINTF1("DETL_CVX_REF_DATE") = DBNull.Value
            rowGLTINTF1("DETL_CVX_REF_NO") = DBNull.Value
            rowGLTINTF1("DETL_DESC") = DBNull.Value
            rowGLTINTF1("DETL_CVX_TYPE") = DBNull.Value
            rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
            ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Next

        Return JOURNAL_NO

    End Function

    Public Shared Sub Update_WHTLOCBX(TRAN_TYPE As String, TRAN_NO As String)

        ASCDATA1.ExecuteSP("WHPLOCB2", _
                           "VVV", _
                           New String() {TRAN_TYPE, TRAN_NO, ASCMAIN1.SESSION_NO}, _
                           New String() {"WHSE_TRAN_TYPE_IN", "WHSE_TRAN_NO_IN", "SESSION_NO_IN"})
    End Sub

#End Region

    Public Shared Function Get_Image(F As ASFBASE0, rowICTSTYL1 As DataRow, ByRef imgba() As Byte) As Bitmap
        Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE") & ""
        Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""
        '   Dim imgba() As Byte = Nothing
        If IMAGE_NAME = "" Then IMAGE_NAME = STYLE_CODE
        Dim FOLDER_NAME As String = F.ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        If ASCMAIN1.Running_in_VS Then FOLDER_NAME = "C:\Users\wjz\Desktop\Data\Database\Images"
        '     rowICTSTYL1.Item("IMAGE") = imgba
        Return ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)

    End Function

    Public Shared Function Calculate_Suggested_SP( _
            LANDED_COST As Decimal, _
            ROYALTY_PCT As Decimal) As Decimal

        'Hi Walter
        'I thought about this a bit and want to adjust the formula slightly. 
        'I want to increase my initial markup a bit to cover some extra costs. I want my markup to be 35% of the sell price. 
        'So if landed cost is 2.00 then suggest sell is 3.07 (approximately). Formula is 2/.65=3.0769

        'Then you can do the rounding.

        '03/30/2015
        ' We pay royalty off the sell price so pls make sure it calculates correctly.  
        ' The net price after royalty should be the same as the price that we wanted originally.  
        ' You probably have your way of doing it buy it can be achieved by dividing the cost price by the converse of the royalty  100%-7%=93% 

        Dim MU As Decimal = 0.35

        'Dim SP As Decimal = System.Math.Round(LANDED_COST * 1.5 + 0.005, 2)
        'Dim SP As Decimal = System.Math.Round(LANDED_COST / 0.65 + 0.005, 2)
        Dim NP As Decimal = System.Math.Round(LANDED_COST / (1 - MU) + 0.005, 2)

        Dim SP As Decimal = NP / (1 - ROYALTY_PCT / 100)

        If SP < 0 Then SP = 0
        Dim SPX As Decimal = (SP * 100) Mod 25
        If SPX >= 10 Then
            SP += (25 - SPX) / 100
        Else
            SP -= SPX / 100
        End If
        Return SP
    End Function

    Public Shared Function Get_SIZEs_and_QTYs_and_COLORs(frmASFBASE0 As ASFBASE0, STYLE_CODE As String) As String

        Dim rowICTSTYLS As DataRow = frmASFBASE0.LookUp("ICTSTYLS", STYLE_CODE)

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

        Dim COLORs As String = ""
        ASCMAIN1.sql = "Select COLOR_CODE, STYLE_COLOR_DESC from ICTSTYC1 " & vbCrLf _
            & " where STYLE_CODE = '" & STYLE_CODE & "' and STYLE_COLOR_DESC is Not Null"
        For Each row As DataRow In ASCDATA1.GetDataTable().Select("", "COLOR_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim STYLE_COLOR_DESC As String = row.Item("STYLE_COLOR_DESC")
            ' COLORs &= vbCrLf & "#" & COLOR_CODE & " " & STYLE_COLOR_DESC ' PURPOSEFULLY KNOCKING OFF THE # TO KNOW THAT THE REPLACEMENT IS SHOWING
            COLORs &= vbCrLf & COLOR_CODE & " " & STYLE_COLOR_DESC
        Next

        '  Absx1.txtFor("SIZE_SCALE").Text = SIZEs_And_QTYs & COLORs
        Return SIZEs_And_QTYs & COLORs

    End Function
    Public Shared Function Calculate_Style_Royalty_Markup(frmASFBASE0 As ASFBASE0, STYLE_CODE As String, Optional STYLE_PRICE_FEFD As Decimal = 0) As Decimal
        Dim RetVal As Decimal = 0
        Dim S As New System.Text.StringBuilder With {.Length = 0}
        Dim NOW_DATE As Date = CDate((Now().ToShortDateString))

        Dim rowICTSTYL1 As DataRow = frmASFBASE0.LookUp("ICTSTYL1", STYLE_CODE)
        If Not IsNothing(rowICTSTYL1) Then
            Dim ROYALTY_CODE As String = rowICTSTYL1.Item("ROYALTY_CODE").ToString & String.Empty
            Dim STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Style_Price(frmASFBASE0, True, STYLE_CODE, rowICTSTYL1)
            If ROYALTY_CODE.Length > 0 And STYLE_PRICE > 0 Then
                S.Length = 0
                S.AppendLine("SELECT *")
                S.AppendLine("FROM ICTROYL2")
                S.AppendLine($"WHERE ROYALTY_CODE = '{ROYALTY_CODE}'")
                Dim PCT As Decimal = 0
                Dim tblICTROYL2 As DataTable = ASCDATA1.GetDataTable(S.ToString())
                For Each rowICTROYL2 As DataRow In tblICTROYL2.Select("", "ROYALTY_BEGIN")
                    Dim ROYALTY_BEGIN As String = rowICTROYL2.Item("ROYALTY_BEGIN").ToString & String.Empty
                    Dim ROYALTY_END As String = rowICTROYL2.Item("ROYALTY_END").ToString & String.Empty
                    Dim ROYALTY_PCT As Double = rowICTROYL2.Item("ROYALTY_PCT").ToString & String.Empty
                    If IsDate(ROYALTY_BEGIN) Then
                        If CDate(ROYALTY_BEGIN) <= NOW_DATE Then
                            If IsDate(ROYALTY_END) Then
                                If CDate(ROYALTY_END) >= NOW_DATE Then
                                    If IsNumeric(ROYALTY_PCT) Then
                                        PCT = Val(ROYALTY_PCT)
                                    End If
                                End If
                            Else
                                PCT = Val(ROYALTY_PCT)
                            End If
                        End If
                    End If
                Next
                If STYLE_PRICE_FEFD = 0 Then
                    If PCT > 0 And PCT < 100 Then
                        RetVal = STYLE_PRICE / ((100 - PCT) / 100) - STYLE_PRICE
                    End If
                Else
                    If PCT > 0 And PCT < 100 Then
                        RetVal = STYLE_PRICE_FEFD / ((100 - PCT) / 100)
                    Else
                        RetVal = STYLE_PRICE_FEFD
                    End If
                End If
            Else
                If STYLE_PRICE_FEFD > 0 Then
                    RetVal = STYLE_PRICE_FEFD
                End If
            End If
        End If

        Return RetVal
    End Function
    Public Shared Function Calculate_Style_Price(frmASFBASE0 As ASFBASE0, SILENT As Boolean, STYLE_CODE As String, Optional rowICTSTYL1 As DataRow = Nothing, Optional rowICTSTYV1 As DataRow = Nothing, Optional rowICTLSTC1 As DataRow = Nothing, Optional rowAPTVEND1 As DataRow = Nothing) As String

        'Public Shared Function Calculate_Style_Price(frmASFBASE0 As ASFBASE0, STYLE_CODE As String, Optional rowICTSTYL1 As DataRow = Nothing, Optional rowICTSTYV1 as DataRow = Nothing) As String


        Dim STYLE_PRICE As Decimal = 0
        Dim LIST_CALC_CODE As String
        Dim FRT_PER_CTN As Decimal = 0
        Dim FRT_PER_CUBE As Decimal = 0
        Dim MARGIN_FACTOR As Decimal = 0
        Dim CARTON_PACK_QTY As Integer
        Dim CASE_CUBE As Decimal = 0
        Dim PO_COST As Decimal
        Dim VEND_CODE As String
        Dim VEND_COUNTRY As String
        Dim DUTY_RATE_CODE As String
        Dim COUNTRY_CODE As String = ""
        If rowICTSTYL1 Is Nothing Then
            rowICTSTYL1 = frmASFBASE0.LookUp("ICTSTYL1", STYLE_CODE)
        End If
        If rowICTSTYL1 Is Nothing Then
            If Not SILENT Then
                MsgBox("Missing or Invalid Style Code", vbOKOnly, "Cannot Calculate Style Price")
            End If
            Return 0
        End If
        CARTON_PACK_QTY = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
        LIST_CALC_CODE = rowICTSTYL1.Item("LIST_CALC_CODE") & ""
        CASE_CUBE = Val(rowICTSTYL1.Item("CASE_CUBE") & "")
        VEND_CODE = rowICTSTYL1.Item("VEND_CODE") & ""
        DUTY_RATE_CODE = rowICTSTYL1.Item("DUTY_RATE_CODE") & ""
        COUNTRY_CODE = rowICTSTYL1.Item("COUNTRY_CODE") & ""

        If rowICTLSTC1 Is Nothing Then
            rowICTLSTC1 = frmASFBASE0.LookUp("ICTLSTC1", LIST_CALC_CODE)
        End If
        If rowICTLSTC1 Is Nothing Then
            If Not SILENT Then
                MsgBox("Missing or Invalid List Calc Code", vbOKOnly, "Cannot Calculate Style Price")
            End If
            Return 0
        End If

        If rowICTLSTC1 IsNot Nothing Then
            FRT_PER_CTN = Val(rowICTLSTC1.Item("FRT_PER_CTN") & "")
            FRT_PER_CUBE = Val(rowICTLSTC1.Item("FRT_PER_CUBE") & "")
            MARGIN_FACTOR = Val(rowICTLSTC1.Item("MARGIN_FACTOR") & "")
        End If

        If rowAPTVEND1 Is Nothing Then
            rowAPTVEND1 = frmASFBASE0.LookUp("APTVEND1", VEND_CODE)
        End If
        If rowAPTVEND1 Is Nothing Then
            If Not SILENT Then
                MsgBox("Missing or Invalid Vendor Code", vbOKOnly, "Cannot Calculate Style Price")
            End If
            Return 0
        End If

        VEND_COUNTRY = rowAPTVEND1.Item("VEND_COUNTRY") & ""
        If VEND_COUNTRY = "" Then
            If Not SILENT Then
                MsgBox("Vendor does not have a Country Defined", vbOKOnly, "Cannot Calculate Style Price")
            End If
            Return 0
        End If

        If rowICTSTYV1 Is Nothing AndAlso frmASFBASE0.dst.Tables.Contains("ICTSTYV1") Then
            rowICTSTYV1 = frmASFBASE0.dst.Tables("ICTSTYV1").Rows.Find(New String() {STYLE_CODE, VEND_CODE})
        End If
        If rowICTSTYV1 Is Nothing Then
            rowICTSTYV1 = frmASFBASE0.LookUp("ICTSTYV1", New String() {STYLE_CODE, VEND_CODE})
        End If
        If rowICTSTYV1 Is Nothing Then
            If Not SILENT Then
                MsgBox("Missing or Invalid Vendor Cost record", vbOKOnly, "Cannot Calculate Style Price")
            End If
            Return 0
        End If
        PO_COST = Val(rowICTSTYV1.Item("PO_COST").ToString & "")
        If IsDate(rowICTSTYV1.Item("NEW_PO_COST_DATE").ToString & "") Then
            If CDate(rowICTSTYV1.Item("NEW_PO_COST_DATE").ToString & "") < Now() Then
                PO_COST = Val(rowICTSTYV1.Item("NEW_PO_COST").ToString & "")
            End If
        End If


        Dim DUTY_RATE As Decimal = 0
        Dim DUTY_RATE_ADD As String = ""
        Dim DUTY_RATE_ADD_ON As Decimal = 0
        Dim COMPOUNDED_DUTY_RATE As Decimal = 0

        ASCMAIN1.sql = "Select ICTDUTY1.DUTY_RATE_CODE, ICTDUTY1.DUTY_RATE" & vbCrLf _
        & ", ICTDUTY4.DUTY_RATE DUTY_RATE_ADD_ON, ICTDUTY4.DUTY_RATE_ADD" & vbCrLf _
        & " from ICTDUTY1," & vbCrLf _
        & " (Select * from ICTDUTY4 " & vbCrLf _
        & "   where ICTDUTY4.DUTY_RATE_CODE  = '" & DUTY_RATE_CODE & "'" & vbCrLf _
        & "   and ICTDUTY4.COUNTRY_CODE (+) = '" & VEND_COUNTRY & "'" & vbCrLf _
        & "     and ICTDUTY4.DUTY_RATE_BEGIN <= '" & Now.ToString("dd-MMM-yyyy") & "'" & vbCrLf _
        & "     and (ICTDUTY4.DUTY_RATE_END is Null or ICTDUTY4.DUTY_RATE_END >= '" & Now.ToString("dd-MMM-yyyy") & "')" & vbCrLf _
        & " ) ICTDUTY4" & vbCrLf _
        & " where ICTDUTY4.DUTY_RATE_CODE (+) = ICTDUTY1.DUTY_RATE_CODE" & vbCrLf _
         & " AND ICTDUTY1.DUTY_RATE_CODE = '" & DUTY_RATE_CODE & "'"

        Dim rowICTDUTYX As DataRow = ASCDATA1.GetDataRow

        If rowICTDUTYX IsNot Nothing Then
            DUTY_RATE = Val(rowICTDUTYX.Item("DUTY_RATE") & "")
            DUTY_RATE_ADD = rowICTDUTYX.Item("DUTY_RATE_ADD") & ""
            DUTY_RATE_ADD_ON = Val(rowICTDUTYX.Item("DUTY_RATE_ADD_ON") & "")
            COMPOUNDED_DUTY_RATE = System.Math.Floor(DUTY_RATE + DUTY_RATE_ADD_ON + 10 + 0.99)
        End If
        If DUTY_RATE_ADD <> "1" Then
            DUTY_RATE_ADD_ON = 0
        End If


        If CARTON_PACK_QTY = 0 Then
            If Not SILENT Then
                MsgBox("Missing or Invalid Carton Pack Qty", vbOKOnly, "Cannot Calculate Style Price")
            End If
            Return 0
        Else

            STYLE_PRICE = (PO_COST * (100 + COMPOUNDED_DUTY_RATE) / 100 + FRT_PER_CTN / CARTON_PACK_QTY + FRT_PER_CUBE * CASE_CUBE / CARTON_PACK_QTY) * MARGIN_FACTOR

        End If

        If SILENT = False And LIST_CALC_CODE = "" Then
            MsgBox("List Calc Code is Needed to Calculate Correct Style Price", vbOKOnly, "Calculate Style Price")
        End If

        STYLE_PRICE = Math.Round(STYLE_PRICE, 1)

        ASCMAIN1.sql = "Select * from ICTPARM1 where IC_PARM_KEY = 'Z'"
        Dim rowICTPARM1 As DataRow = ASCDATA1.GetDataRow
        Dim IC_PARM_TARIFF_OFFSET_PCT As Decimal = Val(rowICTPARM1.Item("IC_PARM_TARIFF_OFFSET_PCT") & "")
        Dim trumpTax_amt As Decimal = STYLE_PRICE * (IC_PARM_TARIFF_OFFSET_PCT / 100)

        Dim tariffByCountry_amt As Decimal = Calculate_Tariff_By_Country(STYLE_PRICE, COUNTRY_CODE)

        Return Math.Round(STYLE_PRICE + trumpTax_amt + tariffByCountry_amt, 1)


    End Function
    Public Shared Function Calculate_Tariff_By_Country(STYLE_PRICE As Decimal, COUNTRY_CODE As String) As Decimal
        Dim tariffByCountry_amt As Decimal = 0
        If COUNTRY_CODE <> "" Then
            ASCMAIN1.sql = $"Select * from ICTTARF1 where COUNTRY_CODE = '{COUNTRY_CODE}' AND NVL(TARIFF_ACTIVE,'0') = '1'"
            Dim rowICTTARF1 As DataRow = ASCDATA1.GetDataRow
            If rowICTTARF1 IsNot Nothing Then
                Dim TARIFF_DATE As Date = Now
                ASCMAIN1.sql = $"Select * from ICTTARF2 where COUNTRY_CODE = '{COUNTRY_CODE}' 
                                    AND TARIFF_START <= '{TARIFF_DATE}' AND (TARIFF_END IS NULL OR TARIFF_END >= '{TARIFF_DATE}') "
                Dim rowICTTARF2 As DataRow = ASCDATA1.GetDataRow
                If rowICTTARF2 IsNot Nothing Then
                    Dim TARIFF_PCT As Decimal = Val(rowICTTARF2("TARIFF_PCT") & "")
                    tariffByCountry_amt = STYLE_PRICE * (TARIFF_PCT / 100)
                End If
            End If
        End If
        Return tariffByCountry_amt
    End Function
End Class