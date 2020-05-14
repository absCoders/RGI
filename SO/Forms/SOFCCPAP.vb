Imports ABSolution

Public Class SOFCCPAP

    Private Enum CCProcessing
        NotAuthorized
        Locked
        Successful
        Declined
        ShippingMethod
        SalesOrderNotFound
        ProcessingError
        AlreadyProcessed
    End Enum

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Get_PARM("SOTPARM1")

            Create_TDA(.Tables.Add, "SOTORDC1", "*")

            ASCMAIN1.sql = " SELECT '1' SEL, ' ' AUTH_STATUS, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_DATE, SOTPICK1.PICK_RELEASED," _
             & " SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_STORE_NAME, SOTPICK1.PICK_NO, SOTPICK1.CCPA_NO_STATUS," _
             & " SUM(NVL(SOTORDR2.ORDR_UNIT_PRICE, 0) * SOTPICK2.PICK_QTY) PICK_AMT" _
             & " FROM SOTORDR1, SOTPICK1, SOTPICK2, SOTORDR2" _
             & " WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
             & " AND SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
             & " AND SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" _
             & " AND SOTPICK1.CCPA_NO_STATUS = '1'" _
             & " AND SOTPICK1.CCPA_NO_AUTH IS NULL" _
             & " AND SOTPICK1.PICK_STATUS = 'P'" _
             & " AND SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO" _
             & " AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO" _
             & " AND SOTPICK1.PICK_RELEASED >= '03-AUG-2015'" _
             & " GROUP BY SOTORDR1.ORDR_NO, SOTORDR1.ORDR_DATE, SOTPICK1.PICK_RELEASED, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME," _
             & " SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_STORE_NAME, SOTPICK1.PICK_NO, SOTPICK1.CCPA_NO_STATUS"

            Create_TDA(.Tables.Add, "ARTCCPAX", ASCMAIN1.sql, 0, False, String.Empty, 0)
            .Tables("ARTCCPAX").Columns.Add("CCPA_NO_AUTH", GetType(System.String))

        End With

        grdARTCCPAX.DataSource = dst.Tables("ARTCCPAX")
        Create_Summary(grdARTCCPAX, "SEL", "Count")
        Create_Summary(grdARTCCPAX, "PICK_AMT", "Sum")

        ASCMAIN1.Add_Value_List(grdARTCCPAX, "AUTH_STATUS", , New String() {":", Val(CCProcessing.AlreadyProcessed) & ":" & "Already Processed", _
                                                                       Val(CCProcessing.Declined) & ":" & "Declined", _
                                                                       Val(CCProcessing.Locked) & ":" & "Locked", _
                                                                       Val(CCProcessing.NotAuthorized) & ":" & "Not Authorized", _
                                                                       Val(CCProcessing.ProcessingError) & ":" & "Processing Error", _
                                                                       Val(CCProcessing.SalesOrderNotFound) & ":" & "Sales Order Not Found", _
                                                                       Val(CCProcessing.ShippingMethod) & ":" & "Shipping Method", _
                                                                       Val(CCProcessing.Successful) & ":" & "Successful"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty

        Select Case eItemKey

            Case "Load"

            Case "Process CC"
                If dst.Tables("ARTCCPAX").Select("SEL = '1' and CCPA_NO_STATUS = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select atleast one Order to Authorize."
                End If

            Case "Done"

        End Select

        If EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                MyBase.EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Process CC"
                ProcessCreditCardAuthorization()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Process CC").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        grdARTCCPAX.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
         dst.Tables("ARTCCPAX").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Orders that need CC processing")

        Me.Cursor = Cursors.WaitCursor
        EnforceConstraints(False)
        Fill_Records("ARTCCPAX")
        EnforceConstraints(True)

        Sort_grdColumns(grdARTCCPAX, "PICK_AMT")

        ASCMAIN1.Progress(String.Empty, String.Empty)
        Me.Cursor = Cursors.Default
    End Sub

#End Region

    Private Sub ProcessCreditCardAuthorization()

        For Each rowARTCCPAX As DataRow In dst.Tables("ARTCCPAX").Select("SEL = '1' and CCPA_NO_STATUS = '1'", "PICK_AMT")
            Dim ORDR_NO As String = rowARTCCPAX.Item("ORDR_NO")
            Dim PICK_NO As String = rowARTCCPAX.Item("PICK_NO")

            rowARTCCPAX.Item("SEL") = "0"

            Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
            If rowSOTORDR1 Is Nothing Then
                rowARTCCPAX.Item("AUTH_STATUS") = CCProcessing.SalesOrderNotFound
                Continue For
            End If

            ASCMAIN1.sql = "Select * from ARTCCPA1 where ORDR_NO = '" & ORDR_NO & "'"
            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
            For Each row As DataRow In tbl.Select("", "CCPA_NO")
                ' See if this was Web Authorized.
            Next

            Dim ORDR_GROUP_NO As String = rowSOTORDR1.Item("ORDR_GROUP_NO") & String.Empty
            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , , True, 4) Then
                rowARTCCPAX.Item("AUTH_STATUS") = CCProcessing.Locked
                Continue For
            End If

            If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO, , , True, 4) Then
                rowARTCCPAX.Item("AUTH_STATUS") = CCProcessing.Locked
                Continue For
            End If

            Try
                Dim CCPA_NO As String = String.Empty
                Dim chargeAmount As Decimal = rowARTCCPAX.Item("PICK_AMT")

                dst.Tables("SOTORDC1").Clear()
                Dim process As CCProcessing = AuthorizeCreditCard(rowSOTORDR1, chargeAmount, PICK_NO)

                Try
                    BeginTrans()
                    Update_Record_TDA("SOTORDC1")
                    rowARTCCPAX.Item("AUTH_STATUS") = Val(process)

                    If process = CCProcessing.Successful Then
                        ASCMAIN1.sql = "Update SOTPICK1 SET CCPA_NO_STATUS = '2', CCPA_NO_AUTH = '" & CCPA_NO & "' WHERE PICK_NO = '" & PICK_NO & "'"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    End If

                    CommitTrans()
                Catch ex As Exception
                    Rollback()
                    rowARTCCPAX.Item("AUTH_STATUS") = CCProcessing.ProcessingError
                End Try

            Catch ex As Exception
                rowARTCCPAX.Item("AUTH_STATUS") = CCProcessing.ProcessingError
                ASCMAIN1.MultiTask_Release(, , 4)
            End Try
        Next

    End Sub

    Private Function AuthorizeCreditCard(ByRef rowSOTORDR1 As DataRow, ByVal ChargeAmount As Decimal, ByVal PICK_NO As String) As CCProcessing

        Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
        Dim FRT_TERMS As String = rowSOTORDR1.Item("FRT_TERMS") & String.Empty
        Dim SHIP_VIA_CODE As String = rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty
        Dim ORDR_GROUP_NO As String = rowSOTORDR1.Item("ORDR_GROUP_NO") & String.Empty
        Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty

        Dim freightCost As Decimal = 0
        Dim rowSOTORDC1 As DataRow = Nothing

        AuthorizeCreditCard = CCProcessing.NotAuthorized

        Dim TRAN_TYPE As String = "A"

        If rowSOTORDR1.Item("CCPA_NO") & String.Empty <> String.Empty Then
            Dim dispMessage As Boolean = True

            ' Allow RGI to do a second, third, ... Authorization since they make multi shipments and only authorize what they need to ship
            ' the avaiable product
            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                Dim row As DataRow = ASCDATA1.GetDataRow("select * from artccpa1 where CCPA_NO_AUTH = '" & rowSOTORDR1.Item("CCPA_NO") & "'")
                If row IsNot Nothing Then
                    dispMessage = False
                End If
            End If
            If dispMessage Then
                MessageBox.Show("This sales order has an existing credit card authorization. You are not permitted to authorize additional funds.", _
                    "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return CCProcessing.ProcessingError
            End If
        End If

        'RGI Validates Credit Cards on the Web, no actual Credit Card Authorizations
        If rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty = "W" AndAlso ASCMAIN1.DBS_SERVER <> "RGI" Then
            MessageBox.Show("Web sales credit card authorization was processed on the website. You are not permitted to authorize additional funds.", _
                "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return CCProcessing.ProcessingError
        End If

        If Not ",O,P,".Contains(rowSOTORDR1.Item("ORDR_STATUS") & String.Empty) Then
            MessageBox.Show("Only Open and In-Pick statuses can perform a credit card Authorization. If the order has been shipped, you may charge the credit card in Customer Inquiry.", _
                "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return CCProcessing.ProcessingError
        End If

        ASCMAIN1.sql = "Select Count (*) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS in ('O','P','F')"
        If Val(ASCDATA1.GetDataValue) > 1 Then
            MessageBox.Show("You Cannot perform credit card processing on a Multiple Order Group", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return CCProcessing.ProcessingError
        End If

        EMsg = String.Empty
        If FRT_TERMS.Length > 0 Then
            If ASCDATA1.GetDataRow("select * from astcode1 where TABLE_NAME = 'SOTORDR1' AND COLUMN_NAME = 'FRT_TERMS' AND T_CODE = '" & FRT_TERMS & "'") Is Nothing Then
                EMsg &= vbCr & "Freight Terms are required to process a credit card."
            End If
        Else
            EMsg &= vbCr & "Freight Terms are required to process a credit card."
        End If

        If SHIP_VIA_CODE.Length > 0 Then
            If ASCDATA1.GetDataRow("SELECT * FROM SOTSVIA1 WHERE SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'") Is Nothing Then
                EMsg &= vbCr & "Ship Via Code is required for credit card processing."
            End If
        Else
            EMsg &= vbCr & "Ship Via Code is required for credit card processing."
        End If

        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow("select sotcarr1.carrier_type" _
                                                         & " from sotsvia1, sotcarr1" _
                                                         & " where sotsvia1.carrier_code = sotcarr1.carrier_code" _
                                                         & " and ship_via_code = :PARM1", "V", New Object() {SHIP_VIA_CODE})


        If rowSOTCARR1 Is Nothing Then
            EMsg &= vbCr & "Could not determine carrier for the Ship Via Code."
        End If

        If EMsg.Length > 0 Then

            EMsg &= "Errors processing Order Number: " & ORDR_NO & vbCr & EMsg
            MessageBox.Show(EMsg, "Authorize Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return CCProcessing.ProcessingError
        End If

        Dim ORDR_TOTAL_AMT As Decimal = ChargeAmount

        ' Fedex, UPS and similar pay for freight when freight terms of PPA 
        If rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty = "U" AndAlso FRT_TERMS.ToUpper = "PPA" Then
            ' New Rule 1/24/2013. 20% or $20 the greater of the two
            freightCost = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & String.Empty) * 0.2
            If freightCost < 20 Then
                freightCost = 20
            End If
        End If

        Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
        If Not (rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1") Then
            ORDR_TOTAL_AMT += freightCost
        End If

        If ORDR_TOTAL_AMT <= 0 Then
            MessageBox.Show("You cannot charge $0.00 for sales Order No: " & ORDR_NO, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return CCProcessing.ProcessingError
        End If

        If Not ASCMAIN1.Logical_Lock("ARTCUSTC", CUST_CODE) Then Exit Function
        If Not ASCMAIN1.Logical_Open("ARTCCPA1", "*") Then Exit Function

        Using frmCCProcessor As New TAC.TAFCARDF(Me)
            frmCCProcessor.test_mode = ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
            frmCCProcessor.CUST_CODE = CUST_CODE
            frmCCProcessor.CCPA_REASON = "O"
            frmCCProcessor.ORDR_NO = ORDR_NO
            frmCCProcessor.TRAN_TYPE = TRAN_TYPE

            With frmCCProcessor.rowARTCCPA1
                .Item("CUST_CODE") = CUST_CODE
                .Item("CCPA_AMT") = ORDR_TOTAL_AMT
                .Item("CCPA_NOTE") = "Credit Card Order"
            End With

            Try
                Fill_Records("SOTORDC1", String.Empty, True, "Select * from SOTORDC1 where ORDR_NO = '" & ORDR_NO & "'")
                frmCCProcessor.ShowDialog()
                Dim row As DataRow = ASCDATA1.GetDataRow("Select * from ARTCCPA1 where CCPA_NO = :PARM1", "V", New Object() {frmCCProcessor.CCPA_NO & String.Empty})
                If row IsNot Nothing AndAlso (row.Item("CCPA_STATUS") & String.Empty = "T" OrElse row.Item("CCPA_STATUS") & String.Empty = "S") Then
                    rowSOTORDR1.Item("CCPA_NO") = frmCCProcessor.CCPA_NO & String.Empty
                    'rowSOTORDR1.Item("CC_TRANS_ID") = row.Item("TRANS_ID")

                    If TRAN_TYPE = "A" AndAlso row.Item("CCPA_STATUS") & String.Empty = "T" Then
                        'ASCDATA1.ExecuteSQL("UPDATE SOTORDR1 SET CCPA_NO = '" & rowSOTORDR1.Item("CCPA_NO") & "', CC_TRANS_ID = '" & rowSOTORDR1.Item("CC_TRANS_ID") & "' WHERE ORDR_NO = '" & ORDR_NO & "'")
                        ASCDATA1.ExecuteSQL("UPDATE SOTPICK1 SET CCPA_NO_AUTH = '" & rowSOTORDR1.Item("CCPA_NO") & "' WHERE ORDR_NO = '" & PICK_NO & "'")
                    End If

                    dst.Tables("ARTCCPAX").Select("PICK_NO = '" & PICK_NO & "'")(0).Item("CCPA_NO_AUTH") = frmCCProcessor.CCPA_NO
                    dst.Tables("ARTCCPAX").Select("PICK_NO = '" & PICK_NO & "'")(0).Item("CCPA_NO_STATUS") = "2"
 
                    rowSOTORDC1 = dst.Tables("SOTORDC1").NewRow
                    rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                    rowSOTORDC1.Item("TRANS_TYPE") = IIf(TRAN_TYPE = "A", "C", "D")
                    rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                    rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowSOTORDC1.Item("CCPA_NO") = row.Item("CCPA_NO")
                    rowSOTORDC1.Item("CCPA_STATUS") = row.Item("CCPA_STATUS")
                    rowSOTORDC1.Item("AMOUNT") = row.Item("CCPA_AMT")
                    rowSOTORDC1.Item("BALANCE") = row.Item("CCPA_AMT")
                    rowSOTORDC1.Item("ACTIVE_IND") = "1"
                    dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)

                    AuthorizeCreditCard = CCProcessing.Successful

                ElseIf row IsNot Nothing Then
                    rowSOTORDC1 = dst.Tables("SOTORDC1").NewRow
                    rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                    rowSOTORDC1.Item("TRANS_TYPE") = IIf(TRAN_TYPE = "A", "C", "D")
                    rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                    rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowSOTORDC1.Item("CCPA_NO") = row.Item("CCPA_NO")
                    rowSOTORDC1.Item("CCPA_STATUS") = row.Item("CCPA_STATUS")
                    rowSOTORDC1.Item("AMOUNT") = row.Item("CCPA_AMT")
                    rowSOTORDC1.Item("BALANCE") = 0
                    rowSOTORDC1.Item("ACTIVE_IND") = "0"
                    dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)

                    AuthorizeCreditCard = CCProcessing.Declined
                End If

            Catch ex As Exception
                AuthorizeCreditCard = CCProcessing.ProcessingError
                MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Using

        ASCMAIN1.MultiTask_Release()
        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Function

End Class