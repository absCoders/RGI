Public Class POROPEN2

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"

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

                ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf _
                    & ", POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                    & IIf(OS = "O", _
                          ", 'OPENPO' PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf, _
                          ", POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf) _
                    & IIf(OS = "O", _
                          ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_SHP, POTORDR2.PO_QTY_REC, POTORDR2.PO_QTY_OPN, 0 SHIP_QTY, 0 SHIP_OPN, 0 SHIP_REC" & vbCrLf, _
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
        Next
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

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
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
End Class