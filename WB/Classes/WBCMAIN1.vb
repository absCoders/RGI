Public Class WBCMAIN1
    Private Shared dicTASKS As Dictionary(Of String, Int16) = New Dictionary(Of String, Int16)
    Dim WithEvents aCheckBoxOnHeader_CreationFilter As New WBCCHKB1()

    Public Sub AddTaskDetail(ByVal TASK_NO As String, ByVal TASK_DETAIL As String)

        Dim TASK_LNO As Int16 = 0

        If dicTASKS.Keys.Contains(TASK_NO) Then
            TASK_LNO = dicTASKS(TASK_NO) + 1
            dicTASKS(TASK_NO) = TASK_LNO
        Else
            dicTASKS.Add(TASK_NO, 1)
            TASK_LNO = 1
        End If

        TASK_DETAIL = (TASK_DETAIL & String.Empty).Trim
        If TASK_DETAIL.Length > 250 Then
            TASK_DETAIL = TASK_DETAIL.Substring(0, 250).Trim
        End If
        TASK_DETAIL = TASK_DETAIL.Replace("'", "''")

        Try
            Dim sql As String = "INSERT INTO ASTTASK2"
            sql &= " VALUES "
            sql &= String.Format(" ('{0}', {1}, '{2}')", TASK_NO, TASK_LNO, TASK_DETAIL)
            ASCDATA1.ExecuteSQL(sql)

        Catch ex As Exception

        End Try

    End Sub

    Public Function UpdateTask(ByVal TASK_NO As String, ByVal FORM_NAME As String) As String

        Dim sql As String = String.Empty

        If TASK_NO.Length = 0 Then
            TASK_NO = ASCMAIN1.Next_Control_No("ASTTASK1.TASK_NO", 1)
            sql = "Insert Into ASTTASK1 (TASK_NO, FORM_NAME, INIT_OPER, START_TIME)"
            sql &= " VALUES "
            sql &= String.Format(" ('{0}', '{1}', '{2}', SYSDATE)", TASK_NO, FORM_NAME, ASCMAIN1.USER_ID)
            If Not dicTASKS.Keys.Contains(TASK_NO) Then
                dicTASKS.Add(TASK_NO, 1)
            End If
        Else
            sql = String.Format("Update ASTTASK1 SET END_TIME = SYSDATE WHERE TASK_NO = '{0}'", TASK_NO)
        End If

        ASCDATA1.ExecuteSQL(sql)

        Return TASK_NO

    End Function

    Public Sub ValidateOrderData(ByRef tblSOTORDR1 As DataTable, ByRef tblSOTORDR2 As DataTable _
                             , ByRef tblSOTORDR5 As DataTable, ByRef tblSOTORDRV As DataTable)

        Dim errorCodes As List(Of String) = New List(Of String)
        Dim telephoneNo As String = String.Empty
        Dim ORDR_NO As String = String.Empty

        Dim ORDR_QTY_OPEN As Integer = 0
        Dim ORDR_QTY As Integer = 0
        Dim ORDR_QTY_PICK As Integer = 0
        Dim ORDR_QTY_SHIP As Integer = 0
        Dim ORDR_QTY_CANC As Integer = 0

        ASCMAIN1.Progress("Validation Rules", "")

        ' Make sure the Ordr Qty Orig is set and fix Item Description
        For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("", "", DataViewRowState.CurrentRows)
            If Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty) >= 0 Then
                If Val(rowSOTORDR2.Item("ORDR_QTY_ORIG") & String.Empty) = 0 Then
                    rowSOTORDR2.Item("ORDR_QTY_ORIG") = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                End If
            End If

            ORDR_QTY = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
            ORDR_QTY_PICK = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
            ORDR_QTY_SHIP = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty)
            ORDR_QTY_CANC = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty)
            ORDR_QTY_OPEN = ORDR_QTY - (ORDR_QTY_PICK + ORDR_QTY_SHIP + ORDR_QTY_CANC)
            If ORDR_QTY_OPEN < 0 Then ORDR_QTY_OPEN = 0

            ' Place all Order Items into Open
            rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN

            ' Convert &amp; to &
            rowSOTORDR2.Item("ITEM_DESC") = (rowSOTORDR2.Item("ITEM_DESC") & String.Empty).ToString.Replace("&amp;", "&")
        Next

        ' Strip off non digits in the telephone number, convert Guam to GU, and Convert &amp; to &
        For Each rowSOTORDR5 As DataRow In tblSOTORDR5.Rows
            telephoneNo = String.Empty
            For Each ch As Char In rowSOTORDR5.Item("CUST_PHONE") & String.Empty
                If Char.IsDigit(ch) Then telephoneNo &= ch
            Next
            rowSOTORDR5.Item("CUST_PHONE") = telephoneNo

            telephoneNo = String.Empty
            For Each ch As Char In rowSOTORDR5.Item("CUST_FAX") & String.Empty
                If Char.IsDigit(ch) Then telephoneNo &= ch
            Next
            rowSOTORDR5.Item("CUST_FAX") = telephoneNo

            ' Update Full Name if it is missing
            If rowSOTORDR5.Item("CUST_FULL_NAME") & String.Empty = String.Empty Then
                rowSOTORDR5.Item("CUST_FULL_NAME") = (rowSOTORDR5.Item("CUST_FIRST_NAME") & " " & rowSOTORDR5.Item("CUST_LAST_NAME")).ToString.Trim
            End If

            'Convert GUAM to GU
            If (rowSOTORDR5.Item("CUST_STATE") & String.Empty).ToString.Trim.ToUpper = "GUAM" Then
                rowSOTORDR5.Item("CUST_STATE") = "GU"
            End If

            ' Convert &amp; to &
            rowSOTORDR5.Item("CUST_FIRST_NAME") = (rowSOTORDR5.Item("CUST_FIRST_NAME") & String.Empty).ToString.Replace("&amp;", "&")
            rowSOTORDR5.Item("CUST_LAST_NAME") = (rowSOTORDR5.Item("CUST_LAST_NAME") & String.Empty).ToString.Replace("&amp;", "&")
            rowSOTORDR5.Item("CUST_FULL_NAME") = (rowSOTORDR5.Item("CUST_FULL_NAME") & String.Empty).ToString.Replace("&amp;", "&")

            rowSOTORDR5.Item("CUST_ADDR1") = (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Replace("&amp;", "&")
            rowSOTORDR5.Item("CUST_ADDR2") = (rowSOTORDR5.Item("CUST_ADDR2") & String.Empty).ToString.Replace("&amp;", "&")
            rowSOTORDR5.Item("CUST_ADDR3") = (rowSOTORDR5.Item("CUST_ADDR3") & String.Empty).ToString.Replace("&amp;", "&")

            rowSOTORDR5.Item("CUST_COMPANY_NAME") = (rowSOTORDR5.Item("CUST_COMPANY_NAME") & String.Empty).ToString.Replace("&amp;", "&")

        Next

        ' Place the error codes into the Error Table, Recalc order totals
        Dim ordersalesAmt As Double = 0
        Dim isReturn As Boolean = False
        For Each rowSOTORDR1 As DataRow In tblSOTORDR1.Rows

            isReturn = (rowSOTORDR1.Item("ORDR_STATUS") = "R" OrElse rowSOTORDR1.Item("ORDR_TYPE_CODE") = "RET" OrElse rowSOTORDR1.Item("ORDR_TYPE_CODE") = "EXC")
            ORDR_NO = rowSOTORDR1.Item("ORDR_NO")
            ASCMAIN1.Progress("-", ORDR_NO)
            errorCodes.Clear()

            ' Recalculate sales order totals 
            ordersalesAmt = Val(tblSOTORDR2.Compute("SUM(ORDR_EXT_PRICE)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)
            rowSOTORDR1.Item("ORDR_SALES_AMT") = ordersalesAmt

            ordersalesAmt += Val(rowSOTORDR1.Item("ORDR_DISC_AMT") & String.Empty)
            ordersalesAmt += Math.Abs(Val(rowSOTORDR1.Item("ORDR_STAX_AMT") & String.Empty))
            ordersalesAmt += Math.Abs(Val(rowSOTORDR1.Item("ORDR_FRT_AMT") & String.Empty))
            ordersalesAmt += Val(rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") & String.Empty)

            rowSOTORDR1.Item("ORDR_TOT_AMT") = ordersalesAmt

            CalculateCOGS(tblSOTORDR1, tblSOTORDR2)

            If Not isReturn Then
                errorCodes = ValidateSalesOrder(rowSOTORDR1, tblSOTORDR2, tblSOTORDR5)
            End If

            For Each errCd As String In errorCodes.Distinct
                If tblSOTORDRV.Select("ORDR_NO = '" & ORDR_NO & "' AND ERROR_CODE = '" & errCd & "'", "").Length = 0 Then
                    Dim rowSOTORDRV As DataRow = tblSOTORDRV.NewRow
                    rowSOTORDRV.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDRV.Item("ERROR_CODE") = errCd
                    rowSOTORDRV.Item("ERROR_STATUS") = "0"
                    tblSOTORDRV.Rows.Add(rowSOTORDRV)
                End If
            Next

            ' This is also called by the Sales Order Entry Screen.
            ' It may be the case the user approved all errors.
            'If Not isReturn Then
            If tblSOTORDRV.Select("ISNULL(ERROR_STATUS, '0') = '0' AND ORDR_NO = '" & ORDR_NO & "'").Length > 0 Then
                rowSOTORDR1.Item("ORDR_STATUS") = "H"
            Else
                rowSOTORDR1.Item("ORDR_STATUS") = "O"
            End If
            'End If
        Next

    End Sub

    Private Sub CalculateCOGS(ByVal tblSOTORDR1 As DataTable, ByRef tblSOTORDR2 As DataTable)

        Dim ORDR_NO As String = String.Empty
        Dim ORDR_COGS_AMT As Decimal = 0

        For Each rowSOTORDR1 As DataRow In tblSOTORDR1.Rows
            ORDR_NO = rowSOTORDR1.Item("ORDR_NO") & String.Empty

            ORDR_COGS_AMT = 0
            For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("ORDR_NO = '" & ORDR_NO & "'")
                ORDR_COGS_AMT += Val(rowSOTORDR2.Item("ORDR_UNIT_COST") & String.Empty) * Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
            Next

            rowSOTORDR1.Item("ORDR_COGS_AMT") = ORDR_COGS_AMT
        Next

    End Sub

    Public Function ValidateSalesOrder(ByRef rowSOTORDR1 As DataRow, ByRef tblSOTORDR2 As DataTable, ByRef tblSOTORDR5 As DataTable) As List(Of String)

        Dim errorCodes As List(Of String) = New List(Of String)
        Dim sql As String = String.Empty
        Dim sqlFields As String = String.Empty

        Dim rowSOTZIPLK As DataRow = Nothing
        Dim rowICTITEM1 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing
        Dim rowSOTCUSTB As DataRow = Nothing
        Dim rowTATCNTRY As DataRow = Nothing

        Dim zipCode As String = String.Empty
        Dim stateCode As String = String.Empty
        Dim city As String = String.Empty
        Dim country As String = String.Empty
        Dim custTelephone As String = String.Empty

        Dim prefix As String = String.Empty
        Dim firstName As String = String.Empty
        Dim lastName As String = String.Empty
        Dim address As String = String.Empty
        Dim eMail As String = String.Empty
        Dim orderDate As String = String.Empty

        Dim STYLE_CODE As String = String.Empty
        Dim COLOR_CODE As String = String.Empty
        Dim SIZE_CODE As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim PYMT_METHOD_CODE As String = String.Empty
        Dim ORDR_SOURCE_CODE As String = String.Empty

        Dim SHIP_UPGRADE_AMOUNT As Double = Val(ASCDATA1.GetDataValue("SELECT SHIP_UPGRADE_AMOUNT FROM SOTPARM1 WHERE SO_PARM_KEY = 'Z'") & String.Empty)

        If rowSOTORDR1 Is Nothing Then
            Return Nothing
        End If

        PYMT_METHOD_CODE = (rowSOTORDR1.Item("PYMT_METHOD_CODE") & String.Empty).ToString.Trim
        ORDR_SOURCE_CODE = rowSOTORDR1.Item("ORDR_SOURCE_CODE") & String.Empty

        If PYMT_METHOD_CODE = "MP" Then
            AddErrorCode(errorCodes, "M")
        End If

        If SHIP_UPGRADE_AMOUNT > 0 AndAlso Val(rowSOTORDR1.Item("ORDR_SALES_AMT") & String.Empty) >= SHIP_UPGRADE_AMOUNT Then
            AddErrorCode(errorCodes, "UGA")
        End If

        ORDR_NO = rowSOTORDR1.Item("ORDR_NO") & String.Empty

        ' If the ship Via is already set then leave it alone
        If (rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty).ToString.Length = 0 Then
            If tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                Dim SHIP_COUNTRY As String = tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")(0).Item("CUST_COUNTRY") & String.Empty
                SHIP_COUNTRY = SHIP_COUNTRY.Trim.ToUpper

                rowTATCNTRY = Nothing
                rowTATCNTRY = ASCDATA1.GetDataRow("SELECT * FROM TATCNTRY WHERE COUNTRY_CODE = :PARM1", "V", SHIP_COUNTRY)
                If rowTATCNTRY Is Nothing Then
                    rowTATCNTRY = ASCDATA1.GetDataRow("SELECT * FROM TATCNTRY WHERE UPPER(COUNTRY_NAME) = :PARM1", "V", SHIP_COUNTRY)
                End If

                If rowTATCNTRY IsNot Nothing Then
                    SHIP_COUNTRY = rowTATCNTRY.Item("COUNTRY_CODE")
                    tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")(0).Item("CUST_COUNTRY") = SHIP_COUNTRY
                    SetShipViaCode(rowSOTORDR1, SHIP_COUNTRY)
                End If
            End If
        End If

        ' If Payment Type = None then set as error
        If (rowSOTORDR1.Item("PYMT_METHOD_CODE") & String.Empty).ToString <> "MP" Then
            If (rowSOTORDR1.Item("PYMT_TYPE_CODE") & String.Empty).ToString = "NONE" _
                OrElse (rowSOTORDR1.Item("PYMT_TYPE_CODE") & String.Empty).ToString.Trim.Length = 0 _
                OrElse (rowSOTORDR1.Item("PYMT_METHOD_CODE") & String.Empty).ToString = "NONE" Then
                AddErrorCode(errorCodes, "PTC")
            End If
        End If

        ' If the sales order has to have the credit card processed, set an error.
        ' the error is cleared after the card is processed
        'dst.Tables("SOTORDR1").Select("PYMT_METHOD_CODE = 'CC' AND ORDR_SOURCE_CODE = 'SHP'")
        If (rowSOTORDR1.Item("PYMT_METHOD_CODE") & String.Empty).ToString = "CC" _
            AndAlso (rowSOTORDR1.Item("ORDR_SOURCE_CODE") & String.Empty).ToString = "SHP" Then
            AddErrorCode(errorCodes, "CCP")
        End If

        ' If the Ship Via is empty then Set the error flag
        If (rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty).ToString.Length = 0 Then
            AddErrorCode(errorCodes, "S")
        End If

        ' Order Instructions ORDR_INSTR
        If (rowSOTORDR1.Item("ORDR_INSTR") & String.Empty).ToString.Trim.Length > 0 Then
            AddErrorCode(errorCodes, "OI")
        End If

        For Each addrType As String In New String() {"BT", "ST"}
            If tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = '" & addrType & "'", "").Length > 0 Then
                rowSOTORDR5 = tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = '" & addrType & "'", "")(0)

                If (PYMT_METHOD_CODE = "PP" OrElse PYMT_METHOD_CODE = "BUY") AndAlso addrType = "BT" Then
                    If (rowSOTORDR5.Item("CUST_FULL_NAME") & String.Empty).ToString.Trim.Length = 0 Then
                        AddErrorCode(errorCodes, addrType)
                    End If
                ElseIf ",US,CA,".Contains("," & rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty & ",") Then
                    If (rowSOTORDR5.Item("CUST_FULL_NAME") & String.Empty).ToString.Trim.Length = 0 _
                        OrElse (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Trim.Length = 0 _
                        OrElse (rowSOTORDR5.Item("CUST_CITY") & String.Empty).ToString.Trim.Length = 0 _
                        OrElse (rowSOTORDR5.Item("CUST_STATE") & String.Empty).ToString.Trim.Length = 0 _
                        OrElse (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim.Length = 0 Then
                        AddErrorCode(errorCodes, addrType)
                    End If
                Else
                    ' Non US and Canada; therefore, do not validate state code
                    If (rowSOTORDR5.Item("CUST_FULL_NAME") & String.Empty).ToString.Trim.Length = 0 _
                        OrElse (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Trim.Length = 0 _
                        OrElse (rowSOTORDR5.Item("CUST_CITY") & String.Empty).ToString.Trim.Length = 0 _
                        OrElse (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim.Length = 0 Then
                        AddErrorCode(errorCodes, addrType)
                    End If
                End If

                city = (rowSOTORDR5.Item("CUST_CITY") & String.Empty).ToString.Trim.ToUpper
                city = city.Replace("'", "").Replace(".", "").Trim.ToUpper
                stateCode = (rowSOTORDR5.Item("CUST_STATE") & String.Empty).ToString.Trim.ToUpper
                zipCode = (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim.ToUpper
                country = (rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.Trim.ToUpper

                ' Validate City / State / Zip code
                If Not ((PYMT_METHOD_CODE = "PP" OrElse PYMT_METHOD_CODE = "BUY") AndAlso addrType = "BT") AndAlso (country = "US" OrElse country = String.Empty) Then
                    sql = "SELECT * FROM SOTZIPLK"
                    sql &= " WHERE UPPER(ZIP_CODE) = :PARM1"
                    sql &= " AND UPPER(REPLACE(REPLACE(CITY, '''',''), '.','')) = :PARM2"
                    sql &= " AND UPPER(STATE_CODE) = :PARM3"

                    Dim ZIP As String = zipCode
                    If ZIP.Length > 5 Then
                        ZIP = ZIP.Substring(0, 5)
                    End If

                    rowSOTZIPLK = ASCDATA1.GetDataRow(sql, "VVV", New String() {ZIP, city, stateCode})
                    If rowSOTZIPLK Is Nothing Then
                        AddErrorCode(errorCodes, addrType & "Z")
                    End If
                End If

                ' Amazon and Buy.com perform the following checks on their site; therefore, no need to evaluate data
                If (ORDR_SOURCE_CODE = "AMZ" OrElse ORDR_SOURCE_CODE = "BUY") Then
                    Continue For
                End If

                '**********************************************************
                '***************** Look for bad customers *****************
                '**********************************************************
                firstName = (rowSOTORDR5.Item("CUST_FIRST_NAME") & String.Empty).ToString.ToUpper.Trim
                lastName = (rowSOTORDR5.Item("CUST_LAST_NAME") & String.Empty).ToString.ToUpper.Trim
                address = (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper.Trim
                eMail = (rowSOTORDR5.Item("CUST_EMAIL") & String.Empty).ToString.ToUpper.Trim

                custTelephone = rowSOTORDR5.Item("CUST_PHONE") & String.Empty

                If addrType = "BT" Then prefix = "BILL" Else prefix = "SHIP"

                sql = "SELECT UPPER(" & prefix & "_FIRST_NAME) FIRST_NAME, UPPER(" & prefix & "_LAST_NAME) LAST_NAME "
                sql &= ", UPPER(" & prefix & "_EMAIL) EMAIL"
                sql &= ", UPPER(" & prefix & "_ADDRESS1) ADDRESS"
                sql &= ", UPPER(" & prefix & "_PHONE) PHONE"
                sql &= " FROM SOTCUSTB"
                sql &= " WHERE ( UPPER(" & prefix & "_FIRST_NAME) = :PARM1 AND UPPER(" & prefix & "_LAST_NAME) = :PARM2)"
                sql &= " OR UPPER(" & prefix & "_EMAIL) = :PARM3"
                sql &= " OR UPPER(" & prefix & "_ADDRESS1) = :PARM4"
                sql &= " OR UPPER(" & prefix & "_PHONE) = :PARM5"

                ' Prevent Null Lookups
                If firstName.Length = 0 Then firstName = "@@@@"
                If lastName.Length = 0 Then lastName = "@@@@"
                If eMail.Length = 0 Then eMail = "@@@@"
                If address.Length = 0 Then address = "@@@@"
                If custTelephone.Length = 0 Then custTelephone = "@@@@"

                For Each rowSOTCUSTB In ASCDATA1.GetDataTable(sql, "", "VVVVV", New String() {firstName, lastName, eMail, address, custTelephone}).Rows

                    If rowSOTCUSTB IsNot Nothing Then

                        If rowSOTCUSTB.Item("FIRST_NAME") & String.Empty = firstName AndAlso rowSOTCUSTB.Item("LAST_NAME") & String.Empty = lastName Then
                            AddErrorCode(errorCodes, addrType & "N")
                        End If

                        If rowSOTCUSTB.Item("EMAIL") & String.Empty = eMail Then
                            AddErrorCode(errorCodes, addrType & "E")
                        End If

                        If rowSOTCUSTB.Item("ADDRESS") & String.Empty = address Then
                            AddErrorCode(errorCodes, addrType & "A")
                        End If

                        If rowSOTCUSTB.Item("PHONE") & String.Empty = custTelephone Then
                            AddErrorCode(errorCodes, addrType & "P")
                        End If

                    End If
                Next
            Else
                AddErrorCode(errorCodes, addrType)
            End If
        Next

        '**********************************************************
        '***************** Look for Recent Orders *****************
        '**********************************************************

        Dim rowSOTORDR5_BT As DataRow = Nothing
        Dim rowSOTORDR5_ST As DataRow = Nothing

        sql = String.Empty
        sqlFields = String.Empty

        If tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'BT'", "").Length > 0 Then
            rowSOTORDR5_BT = tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'BT'", "")(0)
        Else
            rowSOTORDR5_BT = tblSOTORDR5.NewRow
        End If

        If tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'", "").Length > 0 Then
            rowSOTORDR5_ST = tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'", "")(0)
        Else
            rowSOTORDR5_ST = tblSOTORDR5.NewRow
        End If

        Dim btData As String = String.Empty
        Dim stData As String = String.Empty

        btData = (rowSOTORDR5_BT.Item("CUST_EMAIL") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        stData = (rowSOTORDR5_ST.Item("CUST_EMAIL") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")

        If btData.Length > 0 AndAlso stData.Length > 0 Then
            sql &= " OR (UPPER(CUST_EMAIL) = '" & btData & "' OR UPPER(CUST_EMAIL) = '" & stData & "')"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_EMAIL) = '" & btData & "' OR UPPER(CUST_EMAIL) = '" & stData & "' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END EMAIL"

        ElseIf btData.Length > 0 Then
            sql &= " OR UPPER(CUST_EMAIL) = '" & btData & "'"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_EMAIL) = '" & btData & "' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END EMAIL"
        ElseIf stData.Length > 0 Then
            sql &= " OR UPPER(CUST_EMAIL) = '" & stData & "'"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_EMAIL) = '" & stData & "' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END EMAIL"
        Else
            sqlFields &= ", '0' EMAIL"
        End If

        btData = String.Empty
        stData = String.Empty
        btData = (rowSOTORDR5_BT.Item("CUST_LAST_NAME") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        stData = (rowSOTORDR5_BT.Item("CUST_FIRST_NAME") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        If btData.Length > 0 AndAlso stData.Length > 0 Then
            sql &= " OR (UPPER(CUST_LAST_NAME) = '" & btData & "' AND UPPER(CUST_FIRST_NAME) = '" & stData & "' AND CUST_ADDR_TYPE = 'BT')"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_LAST_NAME) = '" & btData & "' AND UPPER(CUST_FIRST_NAME) = '" & stData & "' AND CUST_ADDR_TYPE = 'BT' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END BT_NAME"
        Else
            sqlFields &= ", '0' BT_NAME"
        End If

        btData = String.Empty
        stData = String.Empty
        btData = (rowSOTORDR5_BT.Item("CUST_ADDR1") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        If btData.Length > 0 Then
            sql &= " OR (UPPER(CUST_ADDR1) = '" & btData & "' AND CUST_ADDR_TYPE = 'BT')"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_ADDR1) = '" & btData & "' AND CUST_ADDR_TYPE = 'BT' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END BT_ADDR"
        Else
            sqlFields &= ", '0' BT_ADDR"
        End If

        btData = String.Empty.Replace("'", "''")
        stData = String.Empty.Replace("'", "''")
        btData = (rowSOTORDR5_BT.Item("CUST_PHONE") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        If btData.Length > 0 Then
            sql &= " OR (UPPER(CUST_PHONE) = '" & btData & "' AND CUST_ADDR_TYPE = 'BT')"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_PHONE) = '" & btData & "' AND CUST_ADDR_TYPE = 'BT' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END BT_PHONE"
        Else
            sqlFields &= ", '0' BT_PHONE"
        End If

        btData = String.Empty
        stData = String.Empty
        btData = (rowSOTORDR5_BT.Item("CUST_LAST_NAME") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        stData = (rowSOTORDR5_ST.Item("CUST_FIRST_NAME") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        If btData.Length > 0 AndAlso stData.Length > 0 Then
            sql &= " OR (UPPER(CUST_LAST_NAME) = '" & btData & "' AND UPPER(CUST_FIRST_NAME) = '" & stData & "')"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_LAST_NAME) = '" & btData & "' AND UPPER(CUST_FIRST_NAME) = '" & stData & "' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END BT_ST_NAME"
        Else
            sqlFields &= ", '0' BT_ST_NAME"
        End If

        btData = String.Empty
        stData = String.Empty
        stData = (rowSOTORDR5_ST.Item("CUST_ADDR1") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        If stData.Length > 0 Then
            sql &= " OR (UPPER(CUST_ADDR1) = '" & stData & "' AND CUST_ADDR_TYPE = 'ST')"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_ADDR1) = '" & stData & "' AND CUST_ADDR_TYPE = 'ST' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END ST_ADDR"
        Else
            sqlFields &= ", '0' ST_ADDR"
        End If

        btData = String.Empty
        stData = String.Empty
        stData = (rowSOTORDR5_ST.Item("CUST_PHONE") & String.Empty).ToString.Trim.ToUpper.Replace("'", "''")
        If stData.Length > 0 Then
            sql &= " OR (UPPER(CUST_PHONE) = '" & stData & "' AND CUST_ADDR_TYPE = 'ST')"

            sqlFields &= ", CASE"
            sqlFields &= "     WHEN UPPER(CUST_PHONE) = '" & stData & "' AND CUST_ADDR_TYPE = 'ST' THEN '1'"
            sqlFields &= "     ELSE '0'"
            sqlFields &= "  END ST_PHONE"
        Else
            sqlFields &= ", '0' ST_PHONE"
        End If

        ' Amazon and Buy checks this data on their site. The guarantee payment, so no need to check
        If sql.Length > 0 AndAlso (ORDR_SOURCE_CODE <> "AMZ" AndAlso ORDR_SOURCE_CODE <> "BUY") Then
            sql = " AND (" & sql.Substring(3).Trim & ")"
            sqlFields = sqlFields.Substring(1).Trim

            orderDate = rowSOTORDR1.Item("ORDR_DATE") & String.Empty
            If IsDate(orderDate) Then
                orderDate = DateAdd(DateInterval.Day, -7, CDate(orderDate)).ToString("dd-MMM-yyyy")
            Else
                orderDate = DateAdd(DateInterval.Day, -7, DateTime.Now).ToString("dd-MMM-yyyy")
            End If

            sql = "SELECT " & sqlFields & " FROM SOTORDR5, SOTORDR1" _
            & " WHERE SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO" _
            & " AND SOTORDR1.ORDR_DATE >= '" & orderDate & "'" _
            & " AND SOTORDR1.ORDR_SOURCE_CODE <> 'RET'" _
            & " AND SOTORDR1.ORDR_NO <> '" & ORDR_NO & "'" _
            & sql

            For Each rowSOTORDR5 In ASCDATA1.GetDataTable(sql).Rows
                If rowSOTORDR5.Item("EMAIL") = "1" Then AddErrorCode(errorCodes, "BTWE") 'errorCodes.Add("BTWE")
                If rowSOTORDR5.Item("BT_NAME") = "1" Then AddErrorCode(errorCodes, "BTWN") 'errorCodes.Add("BTWN")
                If rowSOTORDR5.Item("BT_ADDR") = "1" Then AddErrorCode(errorCodes, "BTWA") 'errorCodes.Add("BTWA")
                If rowSOTORDR5.Item("BT_PHONE") = "1" Then AddErrorCode(errorCodes, "BTWP") 'errorCodes.Add("BTWP")
                If rowSOTORDR5.Item("BT_ST_NAME") = "1" Then AddErrorCode(errorCodes, "BTSTWN") 'errorCodes.Add("BTSTWN")
                If rowSOTORDR5.Item("ST_ADDR") = "1" Then AddErrorCode(errorCodes, "STWA") 'errorCodes.Add("STWA")
                If rowSOTORDR5.Item("ST_PHONE") = "1" Then AddErrorCode(errorCodes, "STWP") 'errorCodes.Add("STWP")
            Next
        End If

        For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("ORDR_NO = '" & ORDR_NO & "'", "")

            STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE") & String.Empty
            COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE") & String.Empty
            SIZE_CODE = rowSOTORDR2.Item("SIZE_CODE") & String.Empty

            sql = "SELECT * FROM ICTITEM1"
            sql &= " WHERE STYLE_CODE = :PARM1"
            sql &= " AND COLOR_CODE = :PARM2"
            sql &= " AND SIZE_CODE = :PARM3"

            rowICTITEM1 = ASCDATA1.GetDataRow(sql, "VVV", New String() {STYLE_CODE, COLOR_CODE, SIZE_CODE})
            If rowICTITEM1 Is Nothing Then
                AddErrorCode(errorCodes, "I", "I_" & rowSOTORDR2.Item("ORDR_LNO"))
            End If

            If Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty) < 0 Then
                AddErrorCode(errorCodes, "Q")
            End If

            If Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & String.Empty) <= 0 Then
                AddErrorCode(errorCodes, "P", "P_" & rowSOTORDR2.Item("ORDR_LNO"))
            End If

        Next

        Dim IP_ADDRESS As String = (rowSOTORDR1.Item("IP_ADDRESS") & String.Empty).ToString.Trim

        If IP_ADDRESS.Length > 0 Then

            Dim LongIp As String() = IP_ADDRESS.Split(".")
            Dim convertedIAddress As ULong = 0

            Try
                ' Prevents an Integer error
                For i As Integer = 1 To Convert.ToInt32(LongIp(0))
                    convertedIAddress += 16777216
                Next

                convertedIAddress += 65536 * Convert.ToInt32(LongIp(1)) + 256 * Convert.ToInt32(LongIp(2)) + Convert.ToInt32(LongIp(3))
            Catch ex As Exception
                convertedIAddress = 0
            End Try

            If LongIp.Length = 4 Then
                rowSOTORDR1.Item("IP_A") = LongIp(0)
                rowSOTORDR1.Item("IP_B") = LongIp(1)
                rowSOTORDR1.Item("IP_C") = LongIp(2)
                rowSOTORDR1.Item("IP_D") = LongIp(3)
            End If

            Dim rowSOTIPLKU As DataRow = Nothing
            sql = "SELECT * FROM SOTIPLKU WHERE IP_FROM <= :PARM1 AND IP_TO >= :PARM1"
            rowSOTIPLKU = ASCDATA1.GetDataRow(sql, "V", convertedIAddress)
            If rowSOTIPLKU IsNot Nothing Then
                rowSOTORDR1.Item("IP_COUNTRY") = rowSOTIPLKU.Item("COUNTRY_CODE") & String.Empty
                sql = "SELECT * FROM TATCNTRY WHERE UPPER(COUNTRY_NAME) = :PARM1 OR UPPER(COUNTRY_CODE) = :PARM1"
                Dim TATCNTRY As DataRow = ASCDATA1.GetDataRow(sql, "V", (rowSOTIPLKU.Item("COUNTRY_CODE") & String.Empty).ToString.Trim.ToUpper)

                If TATCNTRY Is Nothing Then
                    AddErrorCode(errorCodes, "IC")
                ElseIf TATCNTRY.Item("SELL_TO") & String.Empty <> "1" Then
                    AddErrorCode(errorCodes, "ICS")
                End If
            Else
                AddErrorCode(errorCodes, "IP")
            End If
        End If

        Return errorCodes
    End Function

    Private Sub AddErrorCode(ByRef errorCodes As List(Of String), ByVal ErrorCode As String, Optional ByVal DisplayErrorCode As String = "")

        Static tblSOTERRCD As DataTable = Nothing

        If tblSOTERRCD Is Nothing Then
            tblSOTERRCD = ASCDATA1.GetDataTable("SELECT * FROM SOTERRCD")
        End If

        If DisplayErrorCode = String.Empty Then
            DisplayErrorCode = ErrorCode
        End If

        ' No need to create a duplicate error code entry
        If errorCodes.Contains(DisplayErrorCode) Then
            Exit Sub
        End If

        ' If the error is in the active list then apply the error
        If tblSOTERRCD.Select("ERROR_CODE = '" & ErrorCode & "' AND ISNULL(ERROR_ACTIVE, '1') = '1'").Length > 0 Then
            errorCodes.Add(DisplayErrorCode)
        End If

    End Sub

    Public Sub UpdateInventoryOpenHoldPickStatus(ByVal ORDR_NO As String, ByVal S As Integer, _
                                             Optional ByVal UpdateOrderInPick As Boolean = False)

        Dim plusMinus As String = " + 1 * "
        Dim plusMinusPick As String = " - 1 * "

        If S = -1 Then
            plusMinus = " - 1 * "
            plusMinusPick = " + 1 * "
        End If

        ' August 12, 2010
        ' Doing an Item Code check since an Invalid item may be on an imported order.
        ' Although the order is on hold, Esmeralda wants the sales order items put into ORDR_QTY_OPEN

        ' August 18, 2010 - Esmeralda decided to add WHSE_QTY_HOLD so Avail Qty does not include Hold Quantities
        ' **** Uses ICTITEM1, since imported sales order may have Invalid Items
        ' **** This prevents entries in ICTSTAT2 for items that do not exist

        If UpdateOrderInPick Then
            ' If this is a Sales order in Order Entry and it is in Pick, then
            ' remove the qtys from Pick Only.
            ASCMAIN1.sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.SIZE_CODE, " _
            & " SOTORDR1.WHSE_CODE, SOTORDR2.ORDR_QTY_PICK, SOTORDR1.ORDR_STATUS" _
            & " FROM SOTORDR2, SOTORDR1, ICTITEM1 " _
            & " WHERE SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & " AND SOTORDR2.ORDR_NO = '" & ORDR_NO & "'" _
            & " AND NVL(SOTORDR2.ORDR_QTY_PICK,0) <> 0" _
            & " AND SOTORDR2.STYLE_CODE = ICTITEM1.STYLE_CODE" _
            & " AND SOTORDR2.COLOR_CODE = ICTITEM1.COLOR_CODE" _
            & " AND SOTORDR2.SIZE_CODE = ICTITEM1.SIZE_CODE" _
            & " AND SOTORDR1.ORDR_STATUS = 'P';" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & "     UPDATE ICTSTAT2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) " & plusMinus & " NVL(R1.ORDR_QTY_PICK,0), WEB_IND = '1'" _
            & "     WHERE STYLE_CODE = R1.STYLE_CODE " _
            & "     AND COLOR_CODE = R1.COLOR_CODE" _
            & "     AND SIZE_CODE = R1.SIZE_CODE" _
            & "     AND WHSE_CODE = R1.WHSE_CODE ;" _
            & "     IF SQL%NOTFOUND THEN" _
            & "     INSERT INTO ICTSTAT2 (STYLE_CODE, COLOR_CODE, SIZE_CODE, WHSE_CODE, WHSE_QTY_PICK, WEB_IND)" _
            & "     VALUES (R1.STYLE_CODE, R1.COLOR_CODE, R1.SIZE_CODE, R1.WHSE_CODE, " & plusMinus & " NVL(R1.ORDR_QTY_PICK,0), '1');" _
            & "     END IF; " _
            & " END LOOP; END; END;"
        Else
            ASCMAIN1.sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.SIZE_CODE, " _
            & " SOTORDR1.WHSE_CODE, SOTORDR2.ORDR_QTY_OPEN, SOTORDR1.ORDR_STATUS" _
            & " FROM SOTORDR2, SOTORDR1, ICTITEM1 " _
            & " WHERE SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & " AND SOTORDR2.ORDR_NO = '" & ORDR_NO & "'" _
            & " AND NVL(SOTORDR2.ORDR_QTY_OPEN,0) <> 0" _
            & " AND SOTORDR2.STYLE_CODE = ICTITEM1.STYLE_CODE" _
            & " AND SOTORDR2.COLOR_CODE = ICTITEM1.COLOR_CODE" _
            & " AND SOTORDR2.SIZE_CODE = ICTITEM1.SIZE_CODE;" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " IF R1.ORDR_STATUS = 'H' THEN" _
            & "     UPDATE ICTSTAT2 SET WHSE_QTY_HOLD = NVL(WHSE_QTY_HOLD,0) " & plusMinus & " NVL(R1.ORDR_QTY_OPEN,0), WEB_IND = '1'" _
            & "     WHERE STYLE_CODE = R1.STYLE_CODE " _
            & "     AND COLOR_CODE = R1.COLOR_CODE" _
            & "     AND SIZE_CODE = R1.SIZE_CODE" _
            & "     AND WHSE_CODE = R1.WHSE_CODE ;" _
            & "     IF SQL%NOTFOUND THEN" _
            & "     INSERT INTO ICTSTAT2 (STYLE_CODE, COLOR_CODE, SIZE_CODE, WHSE_CODE, WHSE_QTY_HOLD, WEB_IND)" _
            & "     VALUES (R1.STYLE_CODE, R1.COLOR_CODE, R1.SIZE_CODE, R1.WHSE_CODE, " & plusMinus & " NVL(R1.ORDR_QTY_OPEN,0), '1');" _
            & "     END IF; " _
            & " ELSIF R1.ORDR_STATUS = 'P' THEN" _
            & "     UPDATE ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) " & plusMinusPick & " NVL(R1.ORDR_QTY_OPEN,0)," _
            & "         WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) " & plusMinus & " NVL(R1.ORDR_QTY_OPEN,0), WEB_IND = '1'" _
            & "     WHERE STYLE_CODE = R1.STYLE_CODE " _
            & "     AND COLOR_CODE = R1.COLOR_CODE" _
            & "     AND SIZE_CODE = R1.SIZE_CODE" _
            & "     AND WHSE_CODE = R1.WHSE_CODE ;" _
            & "     IF SQL%NOTFOUND THEN" _
            & "     INSERT INTO ICTSTAT2 (STYLE_CODE, COLOR_CODE, SIZE_CODE, WHSE_CODE, WHSE_QTY_PICK, WEB_IND)" _
            & "     VALUES (R1.STYLE_CODE, R1.COLOR_CODE, R1.SIZE_CODE, R1.WHSE_CODE, " & plusMinus & " NVL(R1.ORDR_QTY_OPEN,0), '1');" _
            & "     END IF; " _
            & " ELSE" _
            & "     UPDATE ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) " & plusMinus & " NVL(R1.ORDR_QTY_OPEN,0), WEB_IND = '1'" _
            & "     WHERE STYLE_CODE = R1.STYLE_CODE " _
            & "     AND COLOR_CODE = R1.COLOR_CODE" _
            & "     AND SIZE_CODE = R1.SIZE_CODE" _
            & "     AND WHSE_CODE = R1.WHSE_CODE ;" _
            & "     IF SQL%NOTFOUND THEN" _
            & "     INSERT INTO ICTSTAT2 (STYLE_CODE, COLOR_CODE, SIZE_CODE, WHSE_CODE, WHSE_QTY_OPEN, WEB_IND)" _
            & "     VALUES (R1.STYLE_CODE, R1.COLOR_CODE, R1.SIZE_CODE, R1.WHSE_CODE, " & plusMinus & " NVL(R1.ORDR_QTY_OPEN,0), '1');" _
            & "     END IF; " _
            & " END IF; " _
            & " END LOOP; END; END;"
        End If
        ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub SetShipViaCode(ByVal rowSOTORDR1 As DataRow, ByVal SHIP_COUNTRY As String)

        Dim PARTNER_SHIP_METHOD As String = (rowSOTORDR1.Item("SHIP_VIA_ORIG") & String.Empty).ToString.ToUpper
        Dim ORDR_SOURCE_CODE As String = rowSOTORDR1.Item("ORDR_SOURCE_CODE") & String.Empty

        Dim sql As String = String.Empty

        sql = "SELECT * FROM SOTPART1 WHERE NVL(PARTNER_STATUS, 'A') = 'A' AND PARTNER_ORDR_SOURCE_CODE = :PARM1"
        Dim rowSOTPART1 As DataRow = ASCDATA1.GetDataRow(sql, "V", ORDR_SOURCE_CODE)
        If rowSOTPART1 Is Nothing Then
            Exit Sub
        End If

        Dim PARTNER_CODE As String = rowSOTPART1.Item("PARTNER_CODE") & String.Empty

        sql = "SELECT * FROM SOTSVIA2 WHERE PARTNER_CODE = :PARM1 AND PARTNER_SHIP_METHOD = :PARM2"
        Dim rowSOTSVIA2 As DataRow = ASCDATA1.GetDataRow(sql, "VV", New String() {PARTNER_CODE, PARTNER_SHIP_METHOD})

        If rowSOTSVIA2 IsNot Nothing Then
            If SHIP_COUNTRY = "US" Then
                rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSOTSVIA2.Item("SHIP_VIA_CODE_US") & String.Empty
            Else
                rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSOTSVIA2.Item("SHIP_VIA_CODE_INTL") & String.Empty
            End If
        End If

        If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty <> String.Empty Then
            sql = "SELECT * FROM SOTSVIA1 WHERE SHIP_VIA_CODE = :PARM1"
            Dim rowSOTSVIA1 As DataRow = ASCDATA1.GetDataRow(sql, "V", rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty)
            rowSOTORDR1.Item("CARRIER_CODE") = rowSOTSVIA1.Item("CARRIER_CODE")
        Else
            rowSOTORDR1.Item("SHIP_VIA_CODE") = String.Empty
            rowSOTORDR1.Item("CARRIER_CODE") = String.Empty
        End If

    End Sub

    Private Sub aCheckBoxOnHeader_CreationFilter_HeaderCheckBoxClicked(ByVal sender As Object, ByVal e As WBCCHKB1.HeaderCheckBoxEventArgs) Handles aCheckBoxOnHeader_CreationFilter.HeaderCheckBoxClicked
        ' Check to see if the column is of style checkbox.  If it is, set all the cells in that column to
        ' whatever value the header checkbox is.
        Dim aRow As UltraWinGrid.UltraGridRow
        Dim level As Integer = e.Header.Column.Level

        If e.Header.Column.Style = UltraWinGrid.ColumnStyle.CheckBox Then
            For Each aRow In e.Rows
                aRow.Cells(e.Header.Column.Index).Value = IIf((e.CheckState.Equals(CheckState.Checked)), "1", "0")
                aRow.Update()
            Next
        End If
    End Sub

    Public Sub DisplayHeaderCheckBox(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim columnList As List(Of String) = New List(Of String)
        DisplayHeaderCheckBox(grd, Nothing)
    End Sub

    Public Sub DisplayHeaderCheckBox(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal fieldColumnList() As String)

        Dim columnList As List(Of String) = New List(Of String)

        If fieldColumnList IsNot Nothing Then
            For Each key As String In fieldColumnList
                If key.Length > 0 Then
                    columnList.Add(key)
                End If
            Next
        End If

        ' Set the columns to receive the checkbox
        aCheckBoxOnHeader_CreationFilter.ColumnNames = columnList
        grd.CreationFilter = aCheckBoxOnHeader_CreationFilter
    End Sub

    Public Sub SendWebError(ByVal TO_SUBJECT As String, ByVal HTMLBody As String)
        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim TEMPLATE_NAME As String = "CREDIT"
        EMAIL_ADDRESSs.Add("whr@waynerichmond.net", "Wayne Richmond")
        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                        TO_SUBJECT, TEMPLATE_NAME, True, False, TEMPLATE_NAME, TEMPLATE_NAME, TO_SUBJECT, HTMLBody)
    End Sub
End Class
