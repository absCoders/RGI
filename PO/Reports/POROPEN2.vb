Imports System.Text
Imports SpreadsheetGear
Public Class POROPEN2
    Dim XLS_NO As Integer = 0
    Dim exlExt As String = ".xlsx"
    Dim SOTORDRXSQL As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"
        Get_PARM("ICTPARM1")

        Range_Events(grpPO_DATE_ETA)
        Range_Events(grpPO_DATE_RECEIVED)
        Range_Events(grpPO_DATE_SHIP_BY)
        Range_Events(grpPO_INIT_DATE)

        If ASCMAIN1.USER_SECURITY_CODEs.Contains("X1") Then
            'Absx1.chkFor("CHKCOST").Checked = False
            'Absx1.chkFor("CHKCOST").Visible = False
        End If

        Absx1.chkFor("CHKUSEFIFOCOST").Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
        Absx1.chkFor("CHKCOLLAPSEDETAILS").Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
        chkExcel.Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
        chkStyleStats.Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
        txtPOREF.Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
        lblPOREF.Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")

        chkShowCustomerStyleInfo.Visible = (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = ""
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Dim STOCK_SUB As String = ""
            Select Case Absx1.optFor("OPTASN").Value
                Case "S"
                    STOCK_SUB = "Stock Only"
                Case "N"
                    STOCK_SUB = "Non-Stock Only"
                Case Else
                    STOCK_SUB = "All Stock"
            End Select
            If SUBT.Length = 0 Then
                SUBT = STOCK_SUB
            Else
                SUBT = SUBT & ", " & STOCK_SUB
            End If
        Else
            If Absx1.optFor("OPTASN").Value = "S" Then
                SUBT &= "Stock Styles Only"
            ElseIf Absx1.optFor("OPTASN").Value = "N" Then
                SUBT &= "Non-Stock Styles Only"
            End If
        End If


        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")

        For Each OS As String In New String() {"O", "S"}
            If Absx1.chkFor("CHKINCL_" & OS).Checked Then
                Dim sql_filter2 As String = Get_Dates(OS)

                If Absx1.optFor("OPTASN").Value = "S" Then
                    sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Null"
                ElseIf Absx1.optFor("OPTASN").Value = "N" Then
                    sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Not Null"
                End If

                If Absx1.optFor("OPTSHOW").Value = "O" Then
                    If OS = "O" Then
                        sql_filter2 &= "" _
                            & " and POTORDR1.PO_STATUS = 'O'" & vbCrLf _
                            & " and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
                            & " and NVL(POTORDR2.PO_QTY_OPN,0) <> 0" & vbCrLf
                    Else
                        sql_filter2 &= "" _
                            & " and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                            & " and NVL(POTSHIP3.PO_QTY_REC,0) = 0" & vbCrLf
                    End If
                End If

                ' ADD TO sql_filter2

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    If txtPOREF.Text <> "" Then
                        Dim i As Integer
                        Dim POREF As String = ""
                        Dim datarec() As String = Split(txtPOREF.Text, vbCrLf)
                        For i = 0 To UBound(datarec)
                            If datarec(i).Length <> 0 And POREF = "" Then
                                POREF = POREF & "("
                            End If
                            If datarec(i).Length <> 0 Then
                                POREF = POREF & "'" & datarec(i) & "',"
                            End If
                            '     MessageBox.Show(datarec(i))
                        Next i

                        If POREF <> "" Then
                            POREF = POREF.TrimEnd(CChar(","))
                            POREF = POREF & ")"
                            sql_filter2 &= "" _
                              & " and POTORDR1.PO_REFERENCE IN " & POREF & vbCrLf
                        End If
                    End If
                End If




                If OS = "S" Then
                    sql_TABLE_NAMEs &= ",POTSHIP2,POTSHIP3"
                    sql_JOIN &= "" _
                            & " and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                            & " and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                            & " and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                            & " and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                            & " and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf
                    If Not sql_TABLE_NAMEs.Contains("POTSHIP1") Then
                        sql_TABLE_NAMEs &= ",POTSHIP1"
                    End If
                End If

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

                    ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf _
                        & ", POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                        & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                        & ", POTORDR2.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_ORDERED" _
                        & IIf(OS = "O",
                              ", 'OPENPO' PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf,
                              ", POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf) _
                        & IIf(OS = "O",
                              ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_SHP, POTORDR2.PO_QTY_REC, POTORDR2.PO_QTY_OPN, 0 SHIP_QTY, 0 SHIP_OPN, 0 SHIP_REC" & vbCrLf,
                              ", 0 PO_QTY_ORD, 0 PO_QTY_SHP, 0 PO_QTY_REC, 0 PO_QTY_OPN, POTSHIP3.PO_QTY_SHP SHIP_QTY, DECODE (POTSHIP2.PO_SHIP_STATUS,'O',POTSHIP3.PO_QTY_SHP,0) SHIP_OPN, POTSHIP3.PO_QTY_REC SHIP_REC" & vbCrLf) _
                        & " from POTORDR2" & sql_TABLE_NAMEs & vbCrLf _
                        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf

                    ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                                        & " (" & G1thru9 _
                                        & ",PO_ORDER_NO,PO_ORDER_LNO,STYLE_CODE,COLOR_CODE,PO_DATE_SHIP_BY,PO_DATE_ORDERED,PO_SHIPMENT_NO,PO_SHIPMENT_LNO" _
                                        & ",PO_QTY_ORD,PO_QTY_SHP,PO_QTY_REC,PO_QTY_OPN,SHIP_QTY,SHIP_OPN,SHIP_REC" _
                                        & ") " _
                                        & " (" & ASCMAIN1.sql & ")")


                Else


                    ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf _
                        & ", POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                        & IIf(OS = "O",
                              ", 'OPENPO' PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf,
                              ", POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf) _
                        & IIf(OS = "O",
                              ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_SHP, POTORDR2.PO_QTY_REC, POTORDR2.PO_QTY_OPN, 0 SHIP_QTY, 0 SHIP_OPN, 0 SHIP_REC" & vbCrLf,
                              ", 0 PO_QTY_ORD, 0 PO_QTY_SHP, 0 PO_QTY_REC, 0 PO_QTY_OPN, POTSHIP3.PO_QTY_SHP SHIP_QTY, DECODE (POTSHIP2.PO_SHIP_STATUS,'O',POTSHIP3.PO_QTY_SHP,0) SHIP_OPN, POTSHIP3.PO_QTY_REC SHIP_REC" & vbCrLf) _
                        & " from POTORDR2" & sql_TABLE_NAMEs & vbCrLf _
                        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf

                    ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                                        & " (" & G1thru9 _
                                        & ",PO_ORDER_NO,PO_ORDER_LNO,PO_SHIPMENT_NO,PO_SHIPMENT_LNO" _
                                        & ",PO_QTY_ORD,PO_QTY_SHP,PO_QTY_REC,PO_QTY_OPN,SHIP_QTY,SHIP_OPN,SHIP_REC" _
                                        & ") " _
                                        & " (" & ASCMAIN1.sql & ")")


                End If


            End If
        Next


        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "Select * from (" & vbCrLf _
        & " Select SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE,'O' ORDR_TYPE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
        & ",  SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
        & ", MIN(SOTORDR1.SREP_CODE) SREP_CODE, MIN(SOTORDR1.WHSE_CODE) WHSE_CODE, SOTORDR0.ORDR_TYPE_CODE" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY) ORDR, SUM (SOTORDR2.ORDR_QTY_OPEN) OPEN" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY_PICK) PICK, SUM (SOTORDR2.ORDR_QTY_ALLO) ALLO" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY_SHIP) SHIP, SUM (SOTORDR2.ORDR_QTY_CANC) CANC, MAX (SOTORDR2.ORDR_UNIT_PRICE) PRICE" & vbCrLf _
        & ", COUNT (DISTINCT SOTORDR1.ORDR_NO) ORDERS" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY      * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
        & ", SUM (SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
        & ", ARTCUST1.CUST_NAME" & vbCrLf _
        & ", MIN (SOTORDR1.ORDR_DATE_RECD) ORDR_DATE_RECD, MIN (SOTORDR1.INIT_DATE) INIT_DATE" & vbCrLf _
        & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
        & " From SOTORDR2, SOTORDR1, SOTORDR0, ARTCUST1, ICTATOP1" & vbCrLf _
        & " Where (SOTORDR2.ORDR_STATUS = 'O' OR SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf _
        & " And SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
        & " And ICTATOP1.ORDR_TYPE(+) = 'O'" & vbCrLf _
        & " And ICTATOP1.ORDR_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
        & " And SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
        & " And ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE " & vbCrLf _
        & " And (SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & ASTSRPT1 & ")" & vbCrLf _
        & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
        & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, ARTCUST1.CUST_NAME, SOTORDR0.ORDR_TYPE_CODE" & vbCrLf _
        & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
        & " ) union (" & vbCrLf _
        & " Select SOTRSRV2.STYLE_CODE,SOTRSRV2.COLOR_CODE,'R' ORDR_TYPE, SOTRSRV2.RSRV_NO ORDR_GROUP_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO ORDR_CUST_PO" & vbCrLf _
        & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
        & ", MIN(SOTRSRV1.SREP_CODE) SREP_CODE, MIN(SOTRSRV1.WHSE_CODE) WHSE_CODE, NULL ORDR_TYPE_CODE" & vbCrLf _
        & ", SUM (SOTRSRV2.RSRV_QTY) ORDR, SUM (SOTRSRV2.RSRV_QTY_OPEN) OPEN" & vbCrLf _
        & ", SUM (0) PICK, SUM (SOTRSRV2.RSRV_QTY_ALLO) ALLO" & vbCrLf _
        & ", 0 SHIP, 0 CANC,MAX (SOTRSRV2.ORDR_UNIT_PRICE) PRICE" & vbCrLf _
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
        & " Where SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
        & " And SOTRSRV2.RSRV_QTY_OPEN <> 0" & vbCrLf _
        & " And SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
        & " And ICTATOP1.ORDR_TYPE (+) = 'R'" & vbCrLf _
        & " And ICTATOP1.ORDR_NO (+) = SOTRSRV2.RSRV_NO" & vbCrLf _
        & " And ARTCUST1.CUST_CODE = SOTRSRV1.CUST_CODE" & vbCrLf _
        & " And (SOTRSRV2.STYLE_CODE,SOTRSRV2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & ASTSRPT1 & ")" & vbCrLf _
        & " group by SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, SOTRSRV2.RSRV_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO" & vbCrLf _
        & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE, ARTCUST1.CUST_NAME, SOTRSRV1.INIT_DATE" & vbCrLf _
        & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
        & ")"

            SOTORDRXSQL = ASCMAIN1.sql
            Create_TDA(dst.Tables.Add, "SOTORDRX", "**", 0, False)

        End If




        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC " _
            & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        'ASCMAIN1.sql = "Select POTORDR2.*,ICTSTYL1.CASE_CUBE, ICTCOLR1.COLOR_DESC" & vbCrLf _
        '    & ", POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTORDR1.PO_REFERENCE, POTORDR1.WHSE_CODE" & vbCrLf _
        '    & ", POTORDR1.PO_SPEC_ORDR_NO, POTORDR1.PO_DATE_ORDERED, POTORDR1.PO_DATE_CANCEL,POTORDR1.PORT_CODE_ORIG,POTORDR1.PO_SHIP_VIA" & vbCrLf _
        '    & ", " & IIf(optSORT.Value & "" = "D", "TO_CHAR(POTORDR2.PO_DATE_SHIP_BY,'YYYYMMDD')", IIf(optSORT.Value & "" = "S", "POTORDR2.STYLE_CODE", "POTORDR2.PO_ORDER_NO")) & " SORT1" & vbCrLf _
        '    & ", " & IIf(optSORT.Value & "" = "D", "POTORDR2.PO_ORDER_NO", IIf(optSORT.Value & "" = "S", "POTORDR2.COLOR_CODE", "TO_CHAR(POTORDR2.PO_ORDER_LNO,'000000')")) & " SORT2" & vbCrLf _
        '    & " from POTORDR2,POTORDR1,ICTCOLR1, ICTSTYL1 " & vbCrLf _
        '    & " where (POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO) in (Select Distinct PO_ORDER_NO,PO_ORDER_LNO from " & ASTSRPT1 & ")" & vbCrLf _
        '    & "   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
        '    & "   and POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
        '    & "   and ICTCOLR1.COLOR_CODE = POTORDR2.COLOR_CODE"
        ASCMAIN1.sql = "Select POTORDR2.*,ICTSTYL1.CASE_CUBE, ICTCOLR1.COLOR_DESC" & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTORDR1.PO_REFERENCE, POTORDR1.WHSE_CODE" & vbCrLf _
            & ", POTORDR1.PO_SPEC_ORDR_NO, POTORDR1.PO_DATE_ORDERED, POTORDR1.PO_DATE_CANCEL,POTORDR1.PORT_CODE_ORIG,POTORDR1.PO_SHIP_VIA"
        Select Case optSORT.Value
            Case Is = "D" 'Ship By Date
                ASCMAIN1.sql += ", TO_CHAR(POTORDR2.PO_DATE_SHIP_BY,'YYYYMMDD') SORT1, POTORDR2.PO_ORDER_NO SORT2"
            Case Is = "S" 'Style / Color
                ASCMAIN1.sql += ", POTORDR2.STYLE_CODE SORT1, POTORDR2.COLOR_CODE SORT2"
            Case Is = "P" 'PO / Line
                ASCMAIN1.sql += ", POTORDR2.PO_ORDER_NO SORT1, PO_ORDER_LNO SORT2"
            Case Is = "O" 'Date Entered
                ASCMAIN1.sql += ", TO_CHAR(POTORDR1.PO_DATE_ORDERED,'YYYYMMDD') SORT1, POTORDR2.PO_ORDER_NO SORT2"
        End Select
        ASCMAIN1.sql += " from POTORDR2,POTORDR1,ICTCOLR1, ICTSTYL1 " & vbCrLf _
                    & " where (POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO) in (Select Distinct PO_ORDER_NO,PO_ORDER_LNO from " & ASTSRPT1 & ")" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                    & "   and ICTCOLR1.COLOR_CODE = POTORDR2.COLOR_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDRX", 2))

        ASCMAIN1.sql = "Select POTSHIP3.*" & vbCrLf _
            & ", POTSHIP2.CONTAINER_NO, POTSHIP2.PO_SHIP_STATUS SHIP_STATUS" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_REF_NO" & vbCrLf _
            & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
            & " from POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
            & " where (POTSHIP3.PO_SHIPMENT_NO,POTSHIP3.PO_SHIPMENT_LNO) in (Select Distinct PO_SHIPMENT_NO,PO_SHIPMENT_LNO from " & ASTSRPT1 & ")" & vbCrLf _
            & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTSHIPX", 4))
    End Sub

    Public Overrides Sub Print_Report()
        If ASCMAIN1.CLIENT = "VAN" Then
            RPT = "POROPEN3"
            Dim subtxt As String = txtDescription.Text & SUBT
            If Absx1.chkFor("CHKUSEFIFOCOST").Checked Then
                CR_params.Add("USEFIFOCOST", "1")
                subtxt = subtxt & " Using Initial Costs"
            Else
                CR_params.Add("USEFIFOCOST", "0")
            End If
            If Absx1.chkFor("CHKCOLLAPSEDETAILS").Checked Then
                CR_params.Add("COLLAPSEDETAILS", "1")
            Else
                CR_params.Add("COLLAPSEDETAILS", "0")
            End If
            If chkDZNCOST.Checked Then
                CR_params.Add("DZNCOST", "1")
                subtxt = subtxt & " (cost in dzn)"
            Else
                CR_params.Add("DZNCOST", "0")
            End If
            CR_params.Add("SUBT", subtxt)
            If chkStyleStats.Checked Then
                CR_params.Add("SIP", "")
            Else
                CR_params.Add("SIP", "1")
            End If


        End If
        If ASCMAIN1.CLIENT = "RGI" Then
            CR_params.Add("SUBT", txtDescription.Text & SUBT)
            RPT = "POROPENR"
        End If
        If ASCMAIN1.CLIENT = "NYA" Then
            '  CR_params.Add("SUBT", txtDescription.Text & SUBT)
            RPT = "POROPENN"
        End If
        Generate_Report(RPT, , SUBT)
        If ASCMAIN1.CLIENT = "VAN" Then
            If chkExcel.Checked Then
                Dim XLS_FILENAME1 As String = MakeExcelWorkbook()
                Dim XLS_FILENAME2 As String = ""
                Show_Document(XLS_FILENAME1)
                ASCMAIN1.Progress("", "")
            End If
            txtPOREF.Text = ""

        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
            '    EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            'End If
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length > 4 Then
                EMsg &= vbCr & "Maximum number of Sort Fields for this report is 4"
            End If
        End If
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        If parms.Length > 0 Then
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Fill_Records("SOTORDRX", "", True, SOTORDRXSQL)
        End If

        EnforceConstraints(False)
        'Fill_Records("ASTSRPT1")
        EnforceConstraints(True)
    End Sub

    Function Get_Dates(TYPE As String) As String
        Dim sql As String = ""

        Dim COLUMN_NAMEs() As String
        Dim CONTROL_NAMEs() As String
        If TYPE = "O" Then
            'COLUMN_NAMEs = New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA", "PO_DATE_RECEIVED"}
            'CONTROL_NAMEs = New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA", "PO_DATE_RECEIVED"}
            'COLUMN_NAMEs = New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA"}
            'CONTROL_NAMEs = New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA"}
            COLUMN_NAMEs = New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA", "INIT_DATE"}
            CONTROL_NAMEs = New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA", "PO_INIT_DATE"}
        Else
            COLUMN_NAMEs = New String() {"PO_DATE_SHIPPED", "PO_SHIP_ETA", "PO_DATE_RECEIVED", "INIT_DATE"}
            CONTROL_NAMEs = New String() {"PO_DATE_SHIP_BY", "PO_DATE_ETA", "PO_DATE_RECEIVED", "PO_INIT_DATE"}
        End If
        Dim ctlIndex As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Dim CONTROL_NAME As String = CONTROL_NAMEs(ctlIndex)
            If CONTROL_NAME = "PO_INIT_DATE" Then
                If Not Absx1.chkFor("CHKPO_INIT_DATE_F").Checked Then
                    sql = sql & " and POTORDR1.PO_DATE_ORDERED >= '" & Format(Absx1.dteFor("PO_INIT_DATE_F").Value, "dd-MMM-yyyy") & "'"
                End If
                If Not Absx1.chkFor("CHKPO_INIT_DATE_L").Checked Then
                    sql = sql & " and POTORDR1.PO_DATE_ORDERED <= '" & Format(Absx1.dteFor("PO_INIT_DATE_L").Value, "dd-MMM-yyyy") & "'"
                End If
            Else
                If Not Absx1.chkFor("CHK" & CONTROL_NAME & "_F").Checked Then
                    sql = sql & " and A." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(CONTROL_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
                End If
                If Not Absx1.chkFor("CHK" & CONTROL_NAME & "_L").Checked Then
                    sql = sql & " and A." & COLUMN_NAME & " <= '" & Format(Absx1.dteFor(CONTROL_NAME & "_L").Value, "dd-MMM-yyyy") & "'"
                End If
            End If
            ctlIndex += 1
        Next
        If TYPE = "S" Then
            sql = Replace(sql, "A.PO_DATE_RECEIVED", "POTSHIP2.PO_DATE_RECEIVED")
            sql = Replace(sql, "A.", "POTSHIP1.")
        Else
            sql = Replace(sql, "A.", "POTORDR2.")
        End If
        Return sql
    End Function

    Private Function MakeExcelWorkbook() As String
        Dim XLS_FILENAME As String = ""
        Dim REPORT_NAME As String = "POROPEN2"

        Dim StyleList As New List(Of String)

        ''  For Each rowPOTORDRQ As DataRow In dst.Tables("POTORDRX").Select("SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'", "FABRIC_CODE,SUB_BODY_CODE")
        Dim SORTEX As String = ""
        SORTEX = "STYLE_CODE,COLOR_CODE"

        Select Case optSORT.Value
            Case Is = "D" 'Ship By Date
                SORTEX = "PO_DATE_SHIP_BY,STYLE_CODE,COLOR_CODE"
            Case Is = "S" 'Style / Color
                SORTEX = "STYLE_CODE,COLOR_CODE"
            Case Is = "P" 'PO / Line
                SORTEX = "PO_ORDER_NO,PO_ORDER_LNO"
            Case Is = "O" 'Date Entered
                SORTEX = "PO_DATE_ORDERED,STYLE_CODE,COLOR_CODE"
                ''    ASCMAIN1.sql += ", TO_CHAR(POTORDR1.PO_DATE_ORDERED,'YYYYMMDD') SORT1, POTORDR2.PO_ORDER_NO SORT2"
        End Select




        'For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("", SORTEX)
        '    Dim STYLE_CODE As String = rowPOTORDRX.Item("STYLE_CODE").ToString & String.Empty
        '    If Not StyleList.Contains(STYLE_CODE) Then
        '        StyleList.Add(STYLE_CODE)
        '    End If
        'Next
        Dim NEWSORT As String = "G1,G2,G3,G4," & SORTEX

        Dim filter As String = ""
        ' Dim filter As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' AND PO_ORDER_LNO = " & PO_ORDER_LNO
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", NEWSORT)
            '  rowASTSRPT1.Item("SHIP_PO_EXT" & Group) = Val(rowASTSRPT1.Item("SHIP_QTY" & Group) & "") * Val(rowASTSRPT1.Item("PO_COST") & "")

            filter = "PO_ORDER_NO = '" & rowASTSRPT1.Item("PO_ORDER_NO") & "' AND PO_ORDER_LNO = " & rowASTSRPT1.Item("PO_ORDER_LNO")

            For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select(filter, SORTEX)
                Dim STYLE_CODE As String = rowPOTORDRX.Item("STYLE_CODE").ToString & String.Empty
                If Not StyleList.Contains(STYLE_CODE) Then
                    StyleList.Add(STYLE_CODE)
                End If
            Next

        Next




        If chk1Sheet.Checked Then
            Dim fileName As String = ""
            fileName = Create_Excel()
        Else
            Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
            Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
            worksheet.Name = "Style Info
"
            Create_Excel_WorkSheet(worksheet, StyleList)


            If ASCMAIN1.Folders("Temp").EndsWith("\") Then
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & String.Format("{0}.XLSX", REPORT_NAME)
            Else
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & "\" & String.Format("{0}.XLSX", REPORT_NAME)
            End If
            Dim success As Boolean = False

            ASCMAIN1.Progress("Now Saving Workbook")

            Do Until success
                Try
                    If System.IO.File.Exists(XLS_FILENAME) Then
                        System.IO.File.Delete(XLS_FILENAME)
                    End If
                    workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    ''If chkWebLinks.Checked Then
                    ''    SaveLinks(XLS_FILENAME)
                    ''End If
                    success = True
                Catch ex As Exception

                End Try
            Loop
            Return XLS_FILENAME


        End If

    End Function

    Private Function Create_Excel(Optional SALES_DIVISION_CODE As String = "") As String
        Dim RetVal As String = ""



        ''  RESEQ()

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim sqlWB As String = ""
        If SALES_DIVISION_CODE <> "" Then
            sqlWB = " and SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
            ASCMAIN1.Progress("Now Creating Workbook for Divison " & SALES_DIVISION_CODE, "")
        Else
            ASCMAIN1.Progress("Now Creating Workbook", "")
        End If
        Dim sql0 As String = ""
        ''  Dim sql0 As String = " and COUNT_COLOR > 0" ' & Val(numMinQty.Value & "")
        ''If chkShowSelectedOnly.Checked Then
        ''    sql0 &= " and SELECTED = '1'"
        ''End If


        ''  CUSTPOSs.Clear()

        Dim CUSTPOi As Integer = 0
        ''dst.Tables("SOTORDRC").Rows.Clear()

        ''For Each row As DataRow In dst.Tables("ICTSTYC1").Select("")
        ''    row.Item("OPEN_PICK_RSRV") = 0
        ''Next

        ''If chkShowPOs.Checked Then
        ''    For Each row As DataRow In dst.Tables("ICTQUOT2").Select("")
        ''        STYLE_CODE = row.Item("STYLE_CODE_PLM")
        ''        Fill_Records("SOTORDRC", New String() {txtQuoteCUST_CODE.Text, STYLE_CODE}, False)
        ''    Next
        ''    For Each row As DataRow In dst.Tables("SOTORDRC").Select("", "ORDR_CANCEL_DATE")
        ''        Dim OPO As String = row.Item("ORDR_TYPE") & vbTab & row.Item("ORDR_CUST_PO") & vbTab & Format(row.Item("ORDR_SHIP_DATE"), "MM/dd/yyyy") & vbTab & Format(row.Item("ORDR_CANCEL_DATE"), "MM/dd/yyyy")
        ''        If Not CUSTPOSs.ContainsKey(OPO) Then
        ''            CUSTPOi += 1
        ''            CUSTPOSs.Add(OPO, CUSTPOi)
        ''        End If
        ''        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
        ''        Dim COLOR_CODE As String = row.Item("COLOR_CODE")
        ''        Dim QTY As Int64 = Val(row.Item("QTY") & "")
        ''        Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
        ''        If rowICTSTYC1 IsNot Nothing Then
        ''            rowICTSTYC1.Item("OPEN_PICK_RSRV") = Val(rowICTSTYC1.Item("OPEN_PICK_RSRV") & "") + QTY
        ''        End If
        ''    Next
        ''End If

        Dim XLS_CREATED As Boolean = False

        If chk1Sheet.Checked Then
            Dim wsi As Integer = 0
            'Dim WJZ As Integer = dst.Tables("ICTQUOT2").Rows.Count

            Dim CODES As String = ""
            ''If opt1Sheet.Value = "S" Then
            ''    CODES = "SUB_BODY_CODE"
            ''ElseIf opt1Sheet.Value = "FS" Then
            ''    CODES = "FABRIC_CODE,SUB_BODY_CODE,STYLE_GROUP_CODE"
            ''ElseIf opt1Sheet.Value = "G" Then
            ''    CODES = "STYLE_GROUP_CODE,FABRIC_CODE,SUB_BODY_CODE"
            ''    ' CODES = "STYLE_GROUP_CODE"
            ''    ' DGJ
            ''ElseIf opt1Sheet.Value = "D" Then
            CODES = "SALES_DIVISION_CODE"

            '' End If

            For Each rowSB As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDRX").Select(Mid(sqlWB & sql0, 6)), Split(CODES, ",")).Select("")
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

                If CODES = "SALES_DIVISION_CODE" Then
                    Dim SALES_DIVISION_NAME As String = ""
                    SALES_DIVISION_CODE = Mid(SHEET_NAME, 2)
                    ASCMAIN1.sql = "Select SALES_DIVISION_NAME from SOTSDIV1 where SALES_DIVISION_CODE = :PARM1"
                    Dim rowSOTDIV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", Mid(SHEET_NAME, 2))
                    If rowSOTDIV1 IsNot Nothing Then
                        SALES_DIVISION_NAME = rowSOTDIV1.Item("SALES_DIVISION_NAME")
                    Else
                        SALES_DIVISION_NAME = ""
                    End If
                    SHEET_NAME = "Div-" & Mid(SHEET_NAME, 2) & "-" & SALES_DIVISION_NAME
                Else
                    SHEET_NAME = Mid(SHEET_NAME, 2)
                End If


                If dst.Tables("POTORDRX").Select(Mid(sqlWB & sqlSB & sql0, 6)).Length > 0 Then
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

                    Dim StyleList As New List(Of String)

                    '        For Each rowICTSTATD As DataRow In dst.Tables("ICTSTATD").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE, PO_DATE_SHIP_BY")

                    For Each rowPOTORDRQ As DataRow In dst.Tables("POTORDRX").Select("SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'", "FABRIC_CODE,SUB_BODY_CODE")
                        Dim STYLE_CODE As String = rowPOTORDRQ.Item("STYLE_CODE").ToString & String.Empty
                        If Not StyleList.Contains(STYLE_CODE) Then
                            StyleList.Add(STYLE_CODE)
                        End If
                    Next

                    Create_Excel_WorkSheet(worksheet, StyleList, sqlWB & sqlSB & sql0)
                    XLS_CREATED = True
                End If
            Next
        Else
            ''If dst.Tables("POTORDRX").Select(Mid(sqlWB & sql0, 6)).Length > 0 Then
            ''    Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
            ''    worksheet.Name = "Style Info"
            ''    Create_Excel_WorkSheet(worksheet, StyleList, sqlWB & sql0)
            ''    XLS_CREATED = True
            ''End If
        End If

        If XLS_CREATED Then
            Dim XLS_FILENAME As String = ""
            Dim success As Boolean = False

            ASCMAIN1.Progress("Now Saving Workbook")

            Do Until success
                Try
                    XLS_NO += 1
                    ' XLS_FILENAME = Absx1.txtFor("QUOTE_NO").Text
                    XLS_FILENAME = "OpenPOReport"

                    If SALES_DIVISION_CODE <> "" Then
                        XLS_FILENAME &= "-" & SALES_DIVISION_CODE
                    End If
                    XLS_FILENAME &= "-" & Format(XLS_NO, "000") & exlExt
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

    Sub Create_Excel_WorkSheet(worksheet As SpreadsheetGear.IWorksheet,
                               ByVal StyleList As List(Of String), Optional sqlWB As String = "")

        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        If (ASCMAIN1.Running_in_VS) Then
            If Not System.IO.Directory.Exists(IMAGE_FOLDER) Then
                Stop 'You Need to Set up Image Folder.
            End If
        End If

        Dim CX As Integer = 0
        Dim RX As Integer = 0

        Dim I As Integer = 0
        I += 4

        Dim COL0 As Integer = 12

        Dim COL As Integer = COL0

        Excel_DefaultColumns(worksheet, COL)

        With worksheet.Cells(I, 0, I, COL)
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        Dim I0 As Integer = 0
        Dim IA As Integer = 0
        Dim RT(16) As String
        Dim ROW0 As Integer = I
        Dim style_count As Integer = 0
        Dim pages As Integer = 0

        For Each STYLE_CODE As String In StyleList
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            ASCMAIN1.Progress("-", STYLE_CODE)
            I += 1
            I0 = I
            COL = COL0

            Excel_StyleHeader(worksheet, COL, I, COL0)

            I += 1

            Dim ImageRows = 0
            Dim ImageRowsBig = 0
            Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""
            Excel_ImageInsert(worksheet, IMAGE_NAME, IMAGE_FOLDER, ImageRows, ImageRowsBig, I)

            CX = 1

            Excel_StyleMasterfile(worksheet, I, CX, rowICTSTYL1, STYLE_CODE)

            Dim CI As Integer = 0
            Excel_ColorDetails(worksheet, STYLE_CODE, I, COL, COL0, CI)

            For iCOL As Integer = 1 To 16
                COL += 1
                Select Case iCOL
                    Case 5
                     '   worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        'Case 2
                        '    If chkShip2.Checked Then
                        '        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        '    Else
                        '        COL -= 1
                        '    End If
                    Case 3

                        ''If chkAveragePrice.Checked Then
                        ''    worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        ''Else
                        ''    COL -= 1
                        ''End If
                    Case 6
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"

                    Case 10, 15, 16
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"

                End Select

                RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
            Next

            COL += 1

            Dim colsLess As Int16 = 0
            ''If chkAveragePrice.Checked = False Then
            ''    colsLess += 1
            ''End If
            ''If chkShipDates.Checked = False Then
            ''    colsLess += 1
            ''End If
            ''If chkStyleStats.Checked Then
            ''    COL = COL - colsLess
            ''    For iCOL As Integer = 1 To 7
            ''        COL += 1
            ''        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
            ''        RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
            ''    Next
            ''    COL += 0
            ''End If

            ''If chkStyleStats.Checked Then
            ''    worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray
            ''Else
            ''    worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL - colsLess).Interior.Color = SpreadsheetGear.Colors.LightGray
            ''End If


            With worksheet.Cells(I, COL0 - 1, I + CI - 1, COL - colsLess)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
            End With

            I += ImageRowsBig

            Dim CJ As Integer = ImageRows

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

            If chkStyleStats.Checked And (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") Then

                Dim interior As SpreadsheetGear.IInterior
                Dim range As SpreadsheetGear.IRange
                '  I += 1
                COL = COL0
                Dim chkcnt As Int64 = 0
                Dim NEWSTYLE As Boolean = True


                For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE, ORDR_SHIP_DATE")
                    If NEWSTYLE = True Then
                        worksheet.Cells(I - 1, COL - 1).Value = "Ord/Res Details"
                        I += 1
                        ' Headinds and headingsFOrmat
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Col"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Ord Typ"
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Customer"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Customer PO"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Ord Shp Dt"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        range = worksheet.Cells(I - 1, COL - 1, I - 1, COL + 6)
                        interior = range.Interior
                        interior.Color = SpreadsheetGear.Colors.Aquamarine
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Ord Can Dt"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        range = worksheet.Cells(I - 1, COL - 1, I - 1, COL + 6)
                        interior = range.Interior
                        interior.Color = SpreadsheetGear.Colors.Aquamarine
                        chkcnt += 1


                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Qty Ord"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Price"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With

                        NEWSTYLE = False
                    End If



                    I += 1
                    chkcnt = 1
                    If sql = sql Then
                        ' avoid printing if no records in SOTORDRX
                        ' worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & "***"

                    End If



                    '  worksheet.Cells(I - 1, COL - 2 + chkcnt).Value = Val(rowSOTORDRXItem(1) & String.Empty)


                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = Format(Val(rowSOTORDRX.Item("COLOR_CODE") & String.Empty), "000")
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowSOTORDRX.Item("ORDR_TYPE") & String.Empty
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If
                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowSOTORDRX.Item("CUST_CODE") & String.Empty
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowSOTORDRX.Item("ORDR_CUST_PO") & String.Empty
                        .NumberFormat = ""
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = rowSOTORDRX.Item("ORDR_SHIP_DATE") & String.Empty
                        .NumberFormat = "MM/dd/yy"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = rowSOTORDRX.Item("ORDR_CANCEL_DATE") & String.Empty
                        .NumberFormat = "MM/dd/yy"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowSOTORDRX.Item("ORDR") & String.Empty)
                        .NumberFormat = "#,##0"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowSOTORDRX.Item("PRICE") & String.Empty)
                        .NumberFormat = "#,##0.00"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1



                Next
                'T = ""
                'COL += 1


            End If





            With worksheet.Cells(I0, 0, I + 1 - 1, COL)
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
        Next

        I += 2
        COL = COL0

        'Trying to get away without totals here :)
        'worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = 1 To 16
            COL += 1
            Select Case iCOL
                Case 5
                    ''worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    ''GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    '''Case 5
                    '    If chkShip2.Checked Then
                    '        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    '        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    '    Else
                    '        COL -= 1
                    '    End If
                Case 6
                    worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    '    If chkShip2.Checked Then
                    '        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    '        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    '    Else
                    '        COL -= 1
                    '    End If
                Case 10, 15, 16
                    worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
            End Select
        Next

        'worksheet.Cells(I - 1, COL0 - 1, I - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

        Excel_Header(worksheet)

        Excel_PageSetup(worksheet)
    End Sub

    Private Sub Excel_Header(worksheet As IWorksheet)
        Dim H0 As Integer = 8 + 9

        worksheet.Cells(0, H0).Value = "Prep"
        worksheet.Cells(1, H0).Value = "By"
        worksheet.Cells(2, H0).Value = "XNo"

        worksheet.Cells(0, H0, 2, H0).Interior.Color = SpreadsheetGear.Colors.LightGray


        worksheet.Cells(0, H0 + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Left
        worksheet.Cells(0, H0 + 1).Value = Now
        worksheet.Cells(0, H0 + 1).NumberFormat = "MM/dd/yy"

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

        Dim H1 As Integer = 16
        Dim HEAD1 As String = ""
        Dim HEAD2 As String = ""
        Dim HEAD3 As String = ""
        Dim HEAD4 As String = ""
        Dim HEAD5 As String = ""
        Dim HEAD6 As String = ""
        If optShow.Value = "O" Then
            HEAD1 = "Open"
        Else
            HEAD1 = "All"
        End If
        If optASN.Value = "S" Then
            HEAD2 = "Stock"
        ElseIf optASN.Value = "N" Then
            HEAD2 = "NonStock"
        Else
            HEAD2 = "All Styles"
        End If
        If Absx1.chkFor("CHKINCL_O").Checked Then
            HEAD3 = "Include Open Po's"
        End If
        If Absx1.chkFor("CHKINCL_S").Checked Then
            If HEAD3 = "" Then
                HEAD3 = "Include Shipped Po's"
            Else
                HEAD3 = HEAD3 & " & Shipped Po's"
            End If
        End If
        If optSORT.Value = "P" Then
            HEAD4 = "Sort Details By PO No, Line"
        ElseIf optSORT.Value = "S" Then
            HEAD4 = "Sort Details By Style, Color"
        ElseIf optSORT.Value = "D" Then
            HEAD4 = "Sort Details By Ship By Date"
        ElseIf optSORT.Value = "O" Then
            HEAD4 = "Sort Details By Date Ordered"
        End If
        HEAD5 = "By Factory,Customer,Sub-Body" ' Report Sort
        Dim G(4) As String
        Dim pos As Integer
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Rows

            For i As Integer = 1 To 4
                If rowASTSRPT1.Item(("G" & CStr(i)) & "") & "" = "x" Then
                    G(i) = ""
                Else
                    G(i) = rowASTSRPT1.Item(("G" & CStr(i)) & "")
                    pos = G(i).IndexOf(":")
                    G(i) = Mid(G(i), 1, pos)
                End If
            Next
            HEAD5 = G(1) & " " & G(2) & " " & G(3) & " " & G(4) & " "
            Exit For
        Next

        If Absx1.dteFor("PO_DATE_SHIP_BY_F").Value & "" = "" Then
        Else
            HEAD6 = "Ship By Dt: " & Absx1.dteFor("PO_DATE_SHIP_BY_F").Value & "" & " - " & Absx1.dteFor("PO_DATE_SHIP_BY_L").Value & "" & "  "
        End If

        If Absx1.dteFor("PO_DATE_ETA_F").Value & "" = "" Then
        Else
            HEAD6 = HEAD6 & "ETA Dt: " & Absx1.dteFor("PO_DATE_ETA_F").Value & "" & " - " & Absx1.dteFor("PO_DATE_ETA_L").Value & "" & "  "
        End If
        If Absx1.dteFor("PO_DATE_RECEIVED_F").Value & "" = "" Then
        Else
            HEAD6 = HEAD6 & "Rec Dt: " & Absx1.dteFor("PO_DATE_RECEIVED_F").Value & "" & " - " & Absx1.dteFor("PO_DATE_RECEIVED_L").Value & "" & "  "
        End If
        If Absx1.dteFor("PO_INIT_DATE_F").Value & "" = "" Then
        Else
            HEAD6 = HEAD6 & "Ord Dt: " & Absx1.dteFor("PO_INIT_DATE_F").Value & "" & " - " & Absx1.dteFor("PO_INIT_DATE_L").Value & ""
        End If
        worksheet.Cells(0, 2).Value = "Open PO Report with Details"
        worksheet.Cells(0, 2).Font.Bold = True
        worksheet.Cells(1, 2).Value = "PO's: " & HEAD1 & "   Styles: " & HEAD2 & "   Status: " & HEAD3
        worksheet.Cells(1, 2).Font.Bold = True
        worksheet.Cells(2, 2).Value = "Sort Details: " & HEAD4
        worksheet.Cells(2, 2).Font.Bold = True
        worksheet.Cells(3, 2).Value = "Report Sort: " & HEAD5
        worksheet.Cells(3, 2).Font.Bold = True
        If HEAD6 <> "" Then
            worksheet.Cells(4, 2).Value = "Dates: " & HEAD6
            worksheet.Cells(4, 2).Font.Bold = True
        End If


        worksheet.Cells(0, H1).Value = "Note"
        worksheet.Cells(1, H1).Value = "For"

        worksheet.Cells(0, H1, 2, H1).Interior.Color = SpreadsheetGear.Colors.LightGray

        worksheet.Cells(0, H1 + 1).NumberFormat = "MM/dd/yy"
        worksheet.Cells(0, H1 + 1).Value = "Notes"
        ''   worksheet.Cells(1, H1 + 1).Value = txtCUST_CODE.Text & String.Empty

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
        End With
    End Sub
    Private Sub Excel_PageSetup(ByRef worksheet As IWorksheet)
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
    Private Sub Excel_ColorDetails(ByRef worksheet As IWorksheet,
                                   ByVal STYLE_CODE As String,
                                   ByRef i As Integer,
                                   ByRef COL As Integer,
                                   ByRef COL0 As Integer,
                                   ByRef CI As Integer)
        Dim SZMAX As Integer = 0
        Dim SZTOT As Integer = 0
        Dim T As String = ""
        Dim styleTotal As Int64 = 0
        Dim LAST_COLOR As String = ""
        For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE")
            CI += 1
            COL = COL0 + 1
            'COL = COL0
            Dim chkcnt As Int64 = 0
            If LAST_COLOR <> rowPOTORDRX.Item("COLOR_CODE") & String.Empty Then
                worksheet.Cells(i + CI - 1, COL - 2).Value = "'" & rowPOTORDRX.Item("COLOR_CODE") & String.Empty
                'worksheet.Cells(i + CI - 1, COL - 1).Value = rowPOTORDRX.Item("COLOR_DESC") & String.Empty
                worksheet.Cells(i + CI - 1, COL - 1).Value = GetAltColorCode(STYLE_CODE, rowPOTORDRX.Item("COLOR_CODE") & String.Empty, rowPOTORDRX.Item("COLOR_DESC") & String.Empty)
                LAST_COLOR = rowPOTORDRX.Item("COLOR_CODE") & String.Empty
            End If

            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("PO_ORIG_DATE_SHIP_BY") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("PO_REFERENCE") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("VEND_CODE") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("FACTORY_CODE") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("PO_QTY_ORD") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("PO_QTY_OPN") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("PO_DATE_ORDERED") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("PO_DATE_SHIP_BY") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("PO_DATE_ETA") & String.Empty
            chkcnt += 1

            Dim VESSEL As String = ""
            Dim SHP_OPN As Integer = 0
            Dim PO_SHIP_ETA As String = ""
            Dim II As Integer = 1
            Dim PO_COST_CALC As Decimal = 0
            Dim PO_COST_TOT As Decimal = 0
            ' ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowPOTORDRQ.Item("COLOR_CODE") & String.Empty & "'"
            '       For Each rowICTSTAT2 As DataRow In ASCDATA1.GetDataTable.Select("")
            For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("SHIP_STATUS = 'O' AND PO_ORDER_NO = '" & rowPOTORDRX.Item("PO_ORDER_NO") & String.Empty & "' and PO_ORDER_LNO = '" & rowPOTORDRX.Item("PO_ORDER_LNO") & String.Empty & "'")
                ''If rowPOTORDRX.Item("PO_ORDER_NO") & String.Empty = "162223" Then
                ''    Stop
                ''End If
                ''If II > 1 Then
                ''    Stop
                ''End If
                SHP_OPN = SHP_OPN + Val(rowPOTSHIPX.Item("PO_QTY_SHP") & String.Empty)
                PO_SHIP_ETA = rowPOTSHIPX.Item("PO_SHIP_ETA") & String.Empty
                VESSEL = rowPOTSHIPX.Item("PO_SHIP_VESSEL") & String.Empty
                II = II + 1
            Next

            If Absx1.chkFor("CHKUSEFIFOCOST").Checked Then
                If chkDZNCOST.Checked = True Then
                    PO_COST_CALC = Val(rowPOTORDRX.Item("PO_COST_VCOST") & String.Empty) * 12
                Else
                    PO_COST_CALC = Val(rowPOTORDRX.Item("PO_COST_VCOST") & String.Empty)
                End If
                PO_COST_TOT = Val(rowPOTORDRX.Item("PO_COST_VCOST") & String.Empty)
            Else
                If chkDZNCOST.Checked = True Then
                    PO_COST_CALC = Val(rowPOTORDRX.Item("PO_COST") & String.Empty) * 12
                Else
                    PO_COST_CALC = Val(rowPOTORDRX.Item("PO_COST") & String.Empty)
                End If
                PO_COST_TOT = Val(rowPOTORDRX.Item("PO_COST") & String.Empty)
            End If


            ' 3 shipmrnt fielods
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = SHP_OPN
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = PO_SHIP_ETA
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = VESSEL
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("LAST_DATE_SHIP_BY") & String.Empty
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = PO_COST_CALC
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowPOTORDRX.Item("PO_QTY_OPN") & String.Empty) + SHP_OPN
            chkcnt += 1
            worksheet.Cells(i + CI - 1, COL + chkcnt).Value = (Val(rowPOTORDRX.Item("PO_QTY_OPN") & String.Empty) + SHP_OPN) * PO_COST_TOT




            'chkcnt += 1

            ''If chkAveragePrice.Checked Then
            ''    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("AVG_PRICE") & String.Empty
            ''    chkcnt += 1
            ''End If


            ''If chkAveragePrice.Checked Then
            ''    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowPOTORDRX.Item("VAL_SHP") & String.Empty
            ''    chkcnt += 1
            ''End If

            ''If chkShipDates.Checked Then
            ''    'worksheet.Cells(i + CI - 1, COL + chkcnt).Value = "1/1/2018 - 12/31/2018"
            ''    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = GetCustShipDates(rowPOTORDRX.Item("STYLE_CODE") & String.Empty, rowPOTORDRX.Item("COLOR_CODE") & String.Empty)
            ''    chkcnt += 1
            ''End If

            'If chkShip2.Checked Then
            '    worksheet.Cells(i + CI - 1, COL + 1).Value = rowPOTORDRX.Item("QTY_SHP_02") & String.Empty
            '    'chkcnt += 1
            'End If
            'If chkShip3.Checked Then
            '    worksheet.Cells(i + CI - 1, COL + 2).Value = rowPOTORDRX.Item("QTY_SHP_03") & String.Empty
            '    'chkcnt += 1
            'End If
            'T = ""
            'COL += 1

            ''If chkStyleStats.Checked Then

            ''    ' ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowPOTORDRQ.Item("COLOR_CODE") & String.Empty & "'"
            ''    '       For Each rowICTSTAT2 As DataRow In ASCDATA1.GetDataTable.Select("")
            ''    For Each rowICTSTAT2 As DataRow In dst.Tables("ICTSTAT2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowPOTORDRX.Item("COLOR_CODE") & String.Empty & "'")


            ''        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & String.Empty)
            ''        chkcnt += 1
            ''        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & String.Empty)
            ''        chkcnt += 1
            ''        Dim OTS As Integer = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "") - Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & "")
            ''        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = OTS
            ''        chkcnt += 1

            ''        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty)
            ''        chkcnt += 1

            ''        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty)
            ''        chkcnt += 1

            ''        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
            ''        chkcnt += 1

            ''        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = OTS + Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty) + Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty) - Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
            ''        chkcnt += 1

            ''    Next
            ''    T = ""
            ''    COL += 1
            ''End If


        Next

        CI += 2
        COL = COL0

        worksheet.Cells(i - 1, COL - 1, i + CI - 1, COL - 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & "***"
        worksheet.Cells(i + CI - 1, COL - 0).Value = "'" & "Total"
    End Sub

    Private Sub Excel_StyleMasterfile(ByRef worksheet As IWorksheet, ByRef i As Integer, ByRef cx As Integer, ByRef rowICTSTYL1 As DataRow, ByVal STYLE_CODE As String)
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        With worksheet.Cells(i - 1, 3)
            .Value = "'" & STYLE_CODE
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Size = 24
            .Font.Bold = True
        End With

        cx = 3

        worksheet.Cells(i + 2, cx).Value = "Case Qty"

        range = worksheet.Cells(i + 1, 3, i + 2, 4)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.LightGray

        range = worksheet.Cells(i + 1, 3 + 4, i + 2, 4 + 4)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.LightGray

        cx = 5
        worksheet.Cells(i, cx - 2).Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
        worksheet.Cells(i + 2, cx).Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
    End Sub

    Private Sub Excel_ImageInsert(ByRef worksheet As IWorksheet,
                                  ByVal iMAGE_NAME As String,
                                  ByVal IMAGE_FOLDER As String,
                                  ByRef ImageRows As Integer,
                                  ByRef ImageRowsBig As Integer,
                                  ByRef i As Integer)
        Dim imageFileStyle As String = IMAGE_FOLDER & "\" & iMAGE_NAME
        If Not System.IO.File.Exists(imageFileStyle) Then
            iMAGE_NAME = ""
        End If

        If iMAGE_NAME <> "" _
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
            Dim topStyle As Double = windowInfoStyle.RowToPoints(i - 1) + 0.1

            ImageRows = windowInfoStyle.PointsToRow(heightStyle)
            worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
        End If
    End Sub

    Private Sub Excel_StyleHeader(ByRef worksheet As IWorksheet, ByRef COL As Integer, ByRef i As Integer, ByVal COL0 As Integer)
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        worksheet.Cells(i, COL - 1).Value = "" & Chr(13) & Chr(10) & "Color"
        worksheet.Cells(i, COL - 1).Font.Size = 12
        worksheet.Cells(i, COL).Value = "" & Chr(13) & Chr(10) & "Description"
        worksheet.Cells(i, COL).Font.Size = 12

        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Orig Shp"
        End With

        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "RefNo"
        End With

        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Vendor"
        End With
        COL += 1
        With worksheet.Cells(i - 1, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Purchase Order"
        End With
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Factory"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "PO Qty"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Open"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Ordered"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "ShipBy"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "ETA"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "In Transit"
        End With
        COL += 1
        With worksheet.Cells(i - 1, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Shipment"
        End With
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "ETA"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Vessel"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Last Chg"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "PO Cost"
        End With
        COL += 1
        With worksheet.Cells(i - 1, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "WIP+"
        End With
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "In Transit"
        End With
        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Total $"
        End With



        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            ''Dim d1TEXT As String = Format(dteShip_Beg.DateTime, "MM/dd/yy")
            ''Dim d2TEXT As String = Format(dteShip_End.DateTime, "MM/dd/yy")
            ''If optSelectBy.Value = "D" Then
            ''    .Value = String.Format("{0} to {1}", d1TEXT, d2TEXT) & Chr(13) & Chr(10) & "Shp Units"
            ''Else
            ''    .Value = "Selected POs" & Chr(13) & Chr(10) & "Shp Units"
            ''End If
            .Font.Size = 12
        End With

        ''If chkAveragePrice.Checked Then
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "Price"
        ''        .Font.Size = 12
        ''    End With
        ''End If

        ''If chkAveragePrice.Checked Then
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        If optSelectBy.Value = "D" Then
        ''            .Value = "" & Chr(13) & Chr(10) & "Shp Amt"
        ''        Else
        ''            .Value = ""
        ''        End If

        ''        .Font.Size = 12
        ''    End With
        ''End If

        'If chkShip2.Checked Then
        '    COL += 1
        '    With worksheet.Cells(i, COL)
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        Dim d1TEXT As String = Format(dteShip_Beg_2.DateTime, "MM/dd/yy")
        '        Dim d2TEXT As String = Format(dteShip_End_2.DateTime, "MM/dd/yy")
        '        .Value = String.Format("{0} to {1}", d1TEXT, d2TEXT)
        '        .Font.Size = 12
        '    End With
        'End If

        ''If chkShipDates.Checked Then
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "1st & Last Ship Dates"
        ''        .Font.Size = 12
        ''    End With
        ''End If

        'If chkShip3.Checked Then
        '    COL += 1
        '    With worksheet.Cells(i, COL)
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        Dim d1TEXT As String = Format(dteShip_Beg_3.DateTime, "MM/dd/yy")
        '        Dim d2TEXT As String = Format(dteShip_End_3.DateTime, "MM/dd/yy")
        '        If chkLastOnFile.Checked Then
        '            d2TEXT = "Last"
        '        End If
        '        .Value = String.Format("{0} to {1}", d1TEXT, d2TEXT)
        '        .Font.Size = 12
        '    End With
        'End If

        range = worksheet.Cells(i, COL0 - 1, i, COL - 7)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.Gold

        range = worksheet.Cells(i, COL0 - 1 + 11, i, COL - 4)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.LightSkyBlue

        range = worksheet.Cells(i, COL0 - 1 + 14, i, COL)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.WhiteSmoke


        ''If chkStyleStats.Checked Then
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "On Hand"
        ''    End With
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "In Pick"
        ''    End With
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "OTS"
        ''    End With
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "In Transit"
        ''    End With
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "WIP"
        ''    End With
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "Open"
        ''    End With
        ''    COL += 1
        ''    With worksheet.Cells(i, COL)
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .Value = "" & Chr(13) & Chr(10) & "Net Pos"
        ''    End With

        ''    range = worksheet.Cells(i, COL - 6, i, COL)
        ''    interior = range.Interior
        ''    interior.Color = SpreadsheetGear.Colors.Aquamarine

        ''End If

    End Sub

    Private Sub Excel_DefaultColumns(ByRef worksheet As IWorksheet, ByRef COL As Int64)
        worksheet.Cells("A1:AC1").EntireColumn.Font.Size = 16

        Dim CWC() As String = Split("A,B, C,D,E,F,G,H,I,J,K,L, M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC", ",")
        Dim CWS() As String = Split("1,1,40,6,6,6,6,6,6,6,6,6,20,6,6,6,6,6,6,6,6,6,6,6,6,6, 6, 6, 6", ",")
        CWS(2) = 45
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
        Next

        worksheet.Cells(0, 0).EntireColumn.Hidden = True
        worksheet.Cells(0, 1).EntireColumn.Hidden = True

        Dim _COL As Int64 = 1


        'Orig Ship Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd/yy"
        End With

        'PO Ref Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        'vendor Column
        COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        'factory Column
        COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        'po qty Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.NumberFormat = "###,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With
        ' PO Open
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.NumberFormat = "###,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With


        'PO Ord Date
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd/yy"
        End With

        'PO Ship by
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd/yy"
        End With


        'PO ETA DATE
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd/yy"
        End With
        ' Shipment In transit
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.NumberFormat = "###,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With
        'Shipment ETA DATE
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd/yy"
        End With


        'Shipment Vessel Column
        COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        'PO Last Chg DATE
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd"
        End With
        COL += 1
        _COL += 1
        ' WIP PO COST
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.NumberFormat = "##0.00"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With
        ' WIP In Transit
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.NumberFormat = "###,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With
        ' WIP Total $
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 25
            .EntireColumn.NumberFormat = "###,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With




        ''If chkAveragePrice.Checked Then
        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 12
        ''        .EntireColumn.NumberFormat = "#,##0.00"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 12
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''End If

        ''If chkShipDates.Checked Then
        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 30
        ''        .EntireColumn.NumberFormat = "#,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With
        ''End If

        'Ship 2 Column
        'If chkShip2.Checked Then
        '    COL += 1
        '    _COL += 1
        '    With worksheet.Cells(_COL, COL)
        '        .ColumnWidth = 25
        '        .EntireColumn.NumberFormat = "#,##0"
        '        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    End With
        'End If

        'Ship 3 Column
        'If chkShip3.Checked Then
        '    COL += 1
        '    _COL += 1
        '    With worksheet.Cells(_COL, COL)
        '        .ColumnWidth = 25
        '        .EntireColumn.NumberFormat = "#,##0"
        '        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    End With

        'End If

        ''If chkStyleStats.Checked Then
        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 17
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 17
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 17
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 17
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 17
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 17
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With

        ''    COL += 1
        ''    _COL += 1
        ''    With worksheet.Cells(_COL, COL)
        ''        .ColumnWidth = 17
        ''        .EntireColumn.NumberFormat = "#,###,##0"
        ''        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        ''    End With
        ''End If


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
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
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

End Class