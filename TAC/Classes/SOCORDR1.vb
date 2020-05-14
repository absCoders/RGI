Public Class SOCORDR1

    Private clsASFBASE1 As New ABSolution.ASFBASE1
    Private dst As New DataSet

    Public LastError As String = String.Empty

    Public Structure LineDetail

        ''' <summary>
        ''' Order Number
        ''' </summary>
        ''' <remarks></remarks>
        Dim OrderNo As String

        ''' <summary>
        ''' Sales Order Line Number
        ''' </summary>
        ''' <remarks></remarks>
        Dim OrderLineNo As Int16

        ''' <summary>
        ''' Style Code assigned to the Sales order Line Number
        ''' </summary>
        ''' <remarks></remarks>
        Dim StyleCode As String

        ''' <summary>
        ''' Color Code assigned to the Sales order Line Number
        ''' </summary>
        ''' <remarks></remarks>
        Dim ColorCode As String

        ''' <summary>
        ''' Set to -1 to Cancel the Entire Open Quantity; otherwise set to the quantity to cancel
        ''' </summary>
        ''' <remarks></remarks>
        Dim CancelQuantity As Int16

    End Structure

    Public Sub New(ByRef clsASFBASE1_in As ASFBASE1)
        clsASFBASE1 = clsASFBASE1_in
        dst = clsASFBASE1.dst

        With clsASFBASE1.dst
            If Not .Tables.Contains("SOTORDR1_CANC") Then
                clsASFBASE1.Create_TDA(.Tables.Add("SOTORDR1_CANC"), "SOTORDR1", "*")
                clsASFBASE1.Create_TDA(.Tables.Add("SOTORDR2_CANC"), "SOTORDR2", "*", 1)
                clsASFBASE1.Create_TDA(.Tables.Add("SOTORDR7_CANC"), "SOTORDR5", "*")

                clsASFBASE1.Create_TDA(.Tables.Add("SOTORDRG_CANC"), "SOTORDRG", "*")
                clsASFBASE1.Create_TDA(.Tables.Add("SOTORDXR_CANC"), "SOTORDXR", "*")
                clsASFBASE1.Create_TDA(.Tables.Add("TATEVNT1_CANC"), "TATEVNT1", "*")

                ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV2,SOTRSRV1" & vbCrLf _
                    & " where SOTRSRV1.CUST_CODE = :PARM1 " & vbCrLf _
                    & "   and SOTRSRV2.STYLE_CODE = :PARM2 " & vbCrLf _
                    & "   and SOTRSRV2.COLOR_CODE = :PARM3" & vbCrLf _
                    & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                    & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                    & "   and SOTRSRV2.RSRV_QTY_OPEN > 0" & vbCrLf
                clsASFBASE1.Create_TDA(.Tables.Add, "SOTRSRVX_CANC", "**", 0, False, "VVV", 0)
                clsASFBASE1.Create_TDA(.Tables.Add("SOTRSRV1_CANC"), "SOTRSRV1", "*")
                clsASFBASE1.Create_TDA(.Tables.Add("SOTRSRV2_CANC"), "SOTRSRV2", "*")
            End If
        End With

    End Sub

    ''' <summary>
    ''' Marks a Sales Order Line Detail as having the Quantity Cancelled increased.
    ''' </summary>
    ''' <param name="tblSOTORDR2"></param>
    ''' <returns>True if not Errors</returns>
    ''' <remarks>Parameter tblSOTORDR2 shold contain reocrds from SOTORDR2</remarks>
    Public Function EvaluateCancelledDetailLines(ByRef tblSOTORDR2 As DataTable) As Boolean

        Try
            LastError = String.Empty

            If Not tblSOTORDR2.Columns.Contains("ORDR_LINE_CANC") Then
                Return True
            End If

            For Each row As DataRow In tblSOTORDR2.Select("", "", DataViewRowState.CurrentRows)
                If row.RowState = DataRowState.Added Then
                    Continue For
                End If

                If Val(row.Item("ORDR_QTY_CANC", DataRowVersion.Original) & String.Empty) < Val(row.Item("ORDR_QTY_CANC") & String.Empty) Then
                    row.Item("ORDR_LINE_CANC") = "1"
                End If
            Next

            Return True
        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try
    End Function

    Public Function CancelItemsFormSalesOrder(ByVal ORDR_GROUP_NO As String, ByVal cLineDetail As List(Of LineDetail)) As Boolean

        Dim sql As String = String.Empty

        Try
            ClearTables()

            If cLineDetail Is Nothing OrElse cLineDetail.Count = 0 Then
                LastError = "No Line Details provided for Sales Order Number: " & ORDR_GROUP_NO
                Return False
            End If

            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , False, , 9) Then
                LastError = "Cannot lock Sales Order Group Number: " & ORDR_GROUP_NO
                Return False
            End If

            sql = "Select * from Sotordr1 where ordr_group_no = '" & ORDR_GROUP_NO & "'"
            clsASFBASE1.Fill_Records("SOTORDR1_CANC", String.Empty, True, sql)
            If dst.Tables("SOTORDR1_CANC").Rows.Count = 0 Then
                LastError = "Cannot locate Sales Order Group: " & ORDR_GROUP_NO
                Return False
            End If

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1_CANC").Rows(0)
            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
            ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO") & String.Empty

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                ' NO RESERVATIONS
            Else
                If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE, , False, , 9) Then
                    LastError = "Cannot lock Reservations for Customer: " & CUST_CODE
                    Return False
                End If
            End If

            If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE, , False, , 9) Then
                LastError = "Cannot lock SOFOREL1 for customer: " & CUST_CODE
                Return False
            End If


            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                If rowSOTORDR1.Item("ORDR_TYPE_CODE") & "" = "BTB" Then
                    If Not ASCMAIN1.Logical_Open("F", "POFCENT1", , , , 9) Then
                        LastError = "Cannot lock POFCENT1 for order: " & rowSOTORDR1.Item("ORDR_NO")
                        Return False
                    End If
                    ASCMAIN1.sql = "Select PO_ORDER_NO from POTORDR1 where ORDR_NO = '" & rowSOTORDR1.Item("ORDR_NO") & "'"
                    For Each rowPOTORDR1 As DataRow In ASCDATA1.GetDataTable.Rows
                        Dim PO_ORDER_NO As String = rowPOTORDR1.Item("PO_ORDER_NO")
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, , False, , 9) Then
                            LastError = "Cannot lock PO for order: " & rowSOTORDR1.Item("ORDR_NO")
                            Return False
                        End If
                    Next
                End If
            End If

            sql = " SELECT * FROM SOTORDR2 WHERE ORDR_NO IN (Select ORDR_NO from Sotordr1 where ordr_group_no = '" & ORDR_GROUP_NO & "')"
            clsASFBASE1.Fill_Records("SOTORDR2_CANC", String.Empty, True, sql)

            Dim ORDR_NO As String = String.Empty
            Dim ORDR_LNO As Int16 = 0
            Dim STYLE_CODE As String = String.Empty
            Dim COLOR_CODE As String = String.Empty
            Dim rowSOTORDR2 As DataRow = Nothing
            Dim CancelQuantity As Int16 = 0

            Dim ORDR_QTY_OPEN As Int16 = 0

            Dim modifiedOrderNumbers As New List(Of String)

            For Each lDetail As LineDetail In cLineDetail
                ORDR_NO = lDetail.OrderNo & String.Empty
                ORDR_LNO = Val(lDetail.OrderLineNo & String.Empty)
                STYLE_CODE = lDetail.StyleCode & String.Empty
                COLOR_CODE = lDetail.ColorCode & String.Empty
                CancelQuantity = Val(lDetail.CancelQuantity & String.Empty)
                If Not modifiedOrderNumbers.Contains(ORDR_NO) Then
                    modifiedOrderNumbers.Add(ORDR_NO)
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, , False, , 9) Then
                        LastError = "Cannot lock Sales Order No: " & ORDR_NO
                        Return False
                    End If
                End If

                rowSOTORDR2 = dst.Tables("SOTORDR2_CANC").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                If rowSOTORDR2 Is Nothing Then
                    LastError = "Cannot locate Sales Order Number: " & ORDR_NO & ", Line No: " & ORDR_LNO
                    Return False
                End If

                If rowSOTORDR2.Item("STYLE_CODE") & String.Empty <> STYLE_CODE Then
                    LastError = "Sales Order Number: " & ORDR_NO & ", Line No: " & ORDR_LNO & " is not Style: " & STYLE_CODE
                    Return False
                End If

                If rowSOTORDR2.Item("COLOR_CODE") & String.Empty <> COLOR_CODE Then
                    LastError = "Sales Order Number: " & ORDR_NO & ", Line No: " & ORDR_LNO & " is not Color: " & COLOR_CODE
                    Return False
                End If

                If CancelQuantity = 0 Then
                    LastError = "Sales Order Number: " & ORDR_NO & ", Line No: " & ORDR_LNO & ", Style/Color: " & STYLE_CODE & "/" & COLOR_CODE & " is assigned an invalid Cancel Quantity (0)."
                    Return False
                End If

                If CancelQuantity = -1 Then
                    CancelQuantity = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty)
                End If

                ORDR_QTY_OPEN = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty)
                If ORDR_QTY_OPEN <= 0 Then
                    Continue For
                End If

                If CancelQuantity > Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty) Then
                    CancelQuantity = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty)
                End If

                If CancelQuantity <= 0 Then
                    Continue For
                End If

                rowSOTORDR2.Item("ORDR_QTY_OPEN") = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty) - CancelQuantity
                rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + CancelQuantity
            Next

            If Not EvaluateCancelledDetailLines(dst.Tables("SOTORDR2_CANC")) Then
                Return False
            End If

            For Each ORDR_NO In modifiedOrderNumbers
                sql = "ORDR_NO = '" & ORDR_NO & "'"
                Dim ORDR_STATUS As String
                Dim OPEN As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", sql) & "")
                Dim PICK As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_PICK)", sql) & "")
                Dim SHIP As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_SHIP)", sql) & "")
                If OPEN <> 0 Then
                    ORDR_STATUS = "O"
                ElseIf PICK <> 0 Then
                    ORDR_STATUS = "P"
                ElseIf SHIP <> 0 Then
                    ORDR_STATUS = "F"
                Else
                    ORDR_STATUS = "C"
                End If
                rowSOTORDR1.Item("ORDR_STATUS") = ORDR_STATUS
                Record_Event(ORDR_NO, "UPDT", "Sales Order Updated")
                clsASFBASE1.INIT_LAST("SOTORDR1_CANC", False, , True)
            Next

            With clsASFBASE1
                Try
                    .BeginTrans()

                    For Each ORDR_NO In modifiedOrderNumbers
                        Dependent_Updates(-1, ORDR_NO)
                    Next
                    .Update_Record_TDA("SOTORDR1_CANC")
                    .Update_Record_TDA("SOTORDR2_CANC")
                    .Update_Record_TDA("TATEVNT1_CANC")

                    For Each ORDR_NO In modifiedOrderNumbers
                        Dependent_Updates(1, ORDR_NO)
                    Next
                    ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

                    If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                        ' IF WE EVER DO MULTIPLE ORDERS IN A GROUP - WE WILL NEED TO CALL THIS FOR EACH ORDER
                        ASCDATA1.ExecuteSP("SOPORDR1_COMM", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
                    End If

                    Dim rowSOTORDRG As DataRow = .Fill_Record("SOTORDRG_CANC", ORDR_GROUP_NO)
                    If rowSOTORDRG IsNot Nothing AndAlso rowSOTORDRG.Item("ORDR_REL_SHORT") & "" = "1" Then
                        rowSOTORDRG.Item("ORDR_REL_SHORT") = "0"
                        .Update_Record_TDA("SOTORDRG")
                    End If

                    .CommitTrans()
                Catch ex As Exception
                    .Rollback()
                    LastError = ex.Message
                    Return False
                End Try

            End With

            Return True

        Catch ex As Exception
            LastError = ex.Message
            Return False

        Finally
            ASCMAIN1.MultiTask_Release(, , 9)
        End Try
    End Function

    Private Sub ClearTables()

        For Each tableName As String In New String() {"SOTORDR1_CANC", "SOTORDR2_CANC", "SOTORDR7_CANC", _
                                                      "TATEVNT1_CANC", "SOTORDXR_CANC", "SOTRSRVX_CANC", _
                                                      "SOTRSRV1_CANC", "SOTRSRV2_CANC", "SOTORDRG_CANC"}
            clsASFBASE1.dst.Tables(tableName).Rows.Clear()
        Next

    End Sub

    Private Sub Record_Event(ByVal ORDR_NO As String, ByVal EVENT_TYPE As String, ByVal EVENT_DESC As String)
        Dim rowTATEVNT1 As DataRow = clsASFBASE1.dst.Tables("TATEVNT1_CANC").NewRow
        rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
        rowTATEVNT1.Item("TABLE_KEY") = ORDR_NO
        rowTATEVNT1.Item("INIT_DATE") = DateTime.Now
        rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowTATEVNT1.Item("EVENT_TYPE") = EVENT_TYPE
        rowTATEVNT1.Item("EVENT_DESC") = EVENT_DESC
        clsASFBASE1.dst.Tables("TATEVNT1_CANC").Rows.Add(rowTATEVNT1)
    End Sub

    Private Function Check_Changed_Fields(ByVal OrderNo As String, _
                                          ByRef rowSOTORDR1 As DataRow, _
                                          ByRef tblSOTORDR2 As DataTable) As Boolean


        Try

            Dim sql As String = String.Empty
            Dim EntryMode As String = "E"

            sql = "Select Max (REV_NO) from SOTORDXR where ORDR_NO = '" & OrderNo & "'"
            Dim REV_NO As Int16 = Val(ASCDATA1.GetDataValue & "")
            REV_NO += 1

            Dim LAST_DATE As Date = DateTime.Now + ASCMAIN1.NowTSD
            Dim REV_LNO As Integer = 0

            dst.Tables("SOTORDXR").Rows.Clear()

            For i As Integer = 0 To rowSOTORDR1.Table.Columns.Count - 1
                Dim COLUMN_NAME As String = rowSOTORDR1.Table.Columns(i).ColumnName

                If rowSOTORDR1.Item(COLUMN_NAME) & "" _
                <> rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                    Check_Changed_Fields = True
                    ASCMAIN1.Progress("-", COLUMN_NAME)
                    Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR_CANC").NewRow
                    With rowSOTORDXR
                        .Item("REV_NO") = REV_NO
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("ORDR_NO") = OrderNo
                        .Item("ORDR_LNO") = 0
                        .Item("INIT_DATE") = LAST_DATE
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("OLD_VALUE") = rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original)
                        .Item("NEW_VALUE") = rowSOTORDR1.Item(COLUMN_NAME)
                        .Item("EMODE") = EntryMode
                    End With
                    dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                    Check_Changed_Fields = True
                End If
            Next i

            sql = "Select * from SOTORDR2 where ORDR_NO = '" & OrderNo & "'"
            Dim dt As DataTable = ASCDATA1.GetDataTable(sql)

            For Each rowSOTORDR2_orig As DataRow In dt.Select("", "ORDR_LNO")
                Dim ORDR_LNO As Int64 = rowSOTORDR2_orig.Item("ORDR_LNO")
                Dim rowSOTORDR2 As DataRow = tblSOTORDR2.Rows.Find(New Object() {OrderNo, ORDR_LNO})
                If rowSOTORDR2 Is Nothing Then ' Line was Deleted
                    For i As Integer = 0 To dt.Columns.Count - 1
                        Dim COLUMN_NAME As String = rowSOTORDR2_orig.Table.Columns(i).ColumnName
                        Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR_CANC").NewRow
                        With rowSOTORDXR
                            .Item("REV_NO") = REV_NO
                            REV_LNO += 1
                            .Item("REV_LNO") = REV_LNO
                            .Item("ORDR_NO") = OrderNo
                            .Item("ORDR_LNO") = ORDR_LNO
                            .Item("INIT_DATE") = LAST_DATE
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("COLUMN_NAME") = COLUMN_NAME
                            .Item("OLD_VALUE") = rowSOTORDR2_orig.Item(COLUMN_NAME)
                            '.Item("NEW_VALUE") = ""
                            .Item("EMODE") = EntryMode
                            Dim CONTEXT As String
                            If rowSOTORDR2_orig.Item("RANGE_STYLE_CODE") & "" <> "" Then
                                CONTEXT = rowSOTORDR2_orig.Item("RANGE_STYLE_CODE")
                            Else
                                CONTEXT = rowSOTORDR2_orig.Item("STYLE_CODE") & "/" & rowSOTORDR2_orig.Item("COLOR_CODE")
                            End If
                            .Item("CONTEXT") = CONTEXT
                        End With
                        dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                    Next

                Else
                    For i As Integer = 0 To dt.Columns.Count - 1
                        Dim COLUMN_NAME As String = rowSOTORDR2_orig.Table.Columns(i).ColumnName
                        If rowSOTORDR2.Item(COLUMN_NAME) & "" <> rowSOTORDR2_orig.Item(COLUMN_NAME) & "" Then
                            ' Value in Column was Changed
                            Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                            With rowSOTORDXR
                                .Item("REV_NO") = REV_NO
                                REV_LNO += 1
                                .Item("REV_LNO") = REV_LNO
                                .Item("ORDR_NO") = OrderNo
                                .Item("ORDR_LNO") = ORDR_LNO
                                .Item("INIT_DATE") = LAST_DATE
                                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                .Item("COLUMN_NAME") = COLUMN_NAME
                                .Item("OLD_VALUE") = rowSOTORDR2_orig.Item(COLUMN_NAME)
                                .Item("NEW_VALUE") = rowSOTORDR2.Item(COLUMN_NAME)
                                .Item("EMODE") = EntryMode
                                Dim CONTEXT As String
                                If rowSOTORDR2_orig.Item("RANGE_STYLE_CODE") & "" <> "" Then
                                    CONTEXT = rowSOTORDR2_orig.Item("RANGE_STYLE_CODE")
                                Else
                                    CONTEXT = rowSOTORDR2_orig.Item("STYLE_CODE") & "/" & rowSOTORDR2_orig.Item("COLOR_CODE")
                                End If
                                '  .Item("CONTEXT") = CONTEXT
                            End With
                            dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                        End If
                    Next
                End If
            Next

            For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("", "", DataViewRowState.Added)
                Dim ORDR_LNO = rowSOTORDR2.Item("ORDR_LNO")
                ' For i As Integer = 0 To dt.Columns.Count - 1
                Dim COLUMN_NAME As String = "" ' dt.Columns(i).ColumnName
                Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                With rowSOTORDXR
                    .Item("REV_NO") = REV_NO
                    REV_LNO += 1
                    .Item("REV_LNO") = REV_LNO
                    .Item("ORDR_NO") = OrderNo
                    .Item("ORDR_LNO") = ORDR_LNO
                    .Item("INIT_DATE") = LAST_DATE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("COLUMN_NAME") = COLUMN_NAME
                    '.Item("OLD_VALUE") = ""
                    .Item("NEW_VALUE") = "PO Line Added" ' rowSOTORDR2.Item(COLUMN_NAME)
                    .Item("EMODE") = EntryMode
                    Dim CONTEXT As String
                    If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" Then
                        CONTEXT = rowSOTORDR2.Item("RANGE_STYLE_CODE")
                    Else
                        CONTEXT = rowSOTORDR2.Item("STYLE_CODE") & "/" & rowSOTORDR2.Item("COLOR_CODE")
                    End If
                    .Item("CONTEXT") = CONTEXT
                End With
                dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
            Next

            Return True

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    Private Sub Dependent_Updates(S As Integer, ORDR_NO As String)

        Dim QTY_TO_COMMIT As Int64 = 0
        Dim restore_reservation As Boolean = True

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow
        Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")
            Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")

            If S = -1 Then
                If rowSOTORDR2.Item("RSRV_NO") & "" <> "" And restore_reservation Then
                    'Only restore this reservation line if it hasn't been substitutioned.  Per Gabe 07/30/02 - WR.
                    Dim row As DataRow = dst.Tables("SOTORDR2_CANC").Rows.Find(New Object() {ORDR_NO, rowSOTORDR2.Item("ORDR_LNO")})
                    If row IsNot Nothing Then  'Added for Angela. 1/24/05.  She was adding styles to range that had pulled from reservation already.
                        If row.Item("STYLE_CODE_SUB") & "" = "" Then
                            Update_SOTRSRVx(rowSOTORDR2, S, rowSOTORDR1.Item("ORDR_GROUP_NO"))
                        End If
                    End If
                End If
            Else
                Dim rowSOTRSRVX As DataRow = clsASFBASE1.Fill_Record("SOTRSRVX_CANC", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                '& " order by SOTRSRV1.ORDR_CANCEL_DATE"

                Dim Ps() As Object

                If rowSOTRSRVX IsNot Nothing Then
                    rowSOTORDR2.Item("RSRV_NO") = rowSOTRSRVX.Item("RSRV_NO")
                    rowSOTORDR2.Item("RSRV_LNO") = rowSOTRSRVX.Item("RSRV_LNO")
                    Ps = {rowSOTRSRVX.Item("RSRV_NO"), rowSOTRSRVX.Item("RSRV_LNO")}
                    Update_SOTRSRVx(rowSOTORDR2, S, rowSOTORDR1.Item("ORDR_GROUP_NO"))
                Else
                    rowSOTORDR2.Item("RSRV_NO") = DBNull.Value
                    rowSOTORDR2.Item("RSRV_LNO") = DBNull.Value
                    Ps = {DBNull.Value, DBNull.Value}
                End If

                'Update_Record_TDA("SOTORDR2")
                ASCMAIN1.sql = "Update SOTORDR2 Set RSRV_NO = :PARM1, RSRV_LNO = :PARM2" _
                    & " where ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VN", Ps)
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", S * QTY_TO_COMMIT)
            End If
        Next

    End Sub

    Private Sub Update_SOTRSRVx(rowSOTORDR2 As DataRow, S As Integer, ByVal ORDR_GROUP_NO As String)

        Dim RSRV_NO As String = rowSOTORDR2.Item("RSRV_NO") & ""
        Dim RSRV_LNO As Int64 = Val(rowSOTORDR2.Item("RSRV_LNO") & "")

        Dim rowSOTRSRV1 As DataRow = clsASFBASE1.Fill_Record("SOTRSRV1_CANC", RSRV_NO)
        Dim WHSE_CODE As String = rowSOTRSRV1.Item("WHSE_CODE")

        Dim rowSOTRSRV2 As DataRow = clsASFBASE1.Fill_Record("SOTRSRV2_CANC", New String() {RSRV_NO, RSRV_LNO})
        With rowSOTRSRV2
            Dim RSRV_QTY As Int64 = .Item("RSRV_QTY")
            Dim RSRV_QTY_OPEN As Int64 = Val(.Item("RSRV_QTY_OPEN") & "")
            Dim RSRV_QTY_CANC As Int64 = Val(.Item("RSRV_QTY_CANC") & "")
            Dim RSRV_QTY_USED As Int64 = Val(.Item("RSRV_QTY_USED") & "") _
                          + S * Val(rowSOTORDR2.Item("ORDR_QTY") & "")

            '  + S * Val(rowSOTORDR2.Item("ORDR_QTY_ORIG") & "") - USING ORDR_QTY_ORIG WILL ALWAYS HAVE 0 IMPACT WHEN CHANGING THE ORDER
            Dim RSRV_QTY_OPEN_OLD As Int64 = RSRV_QTY_OPEN
            RSRV_QTY_OPEN = RSRV_QTY - RSRV_QTY_CANC - RSRV_QTY_USED
            If RSRV_QTY_OPEN < 0 Then
                RSRV_QTY_OPEN = 0
            End If
            Dim RSRV_QTY_OPEN_NEW As Int64 = RSRV_QTY_OPEN
            .Item("RSRV_QTY_USED") = RSRV_QTY_USED
            .Item("RSRV_QTY_OPEN") = RSRV_QTY_OPEN

            Dim QTY_TO_COMMIT As Int64 = RSRV_QTY_OPEN_NEW - RSRV_QTY_OPEN_OLD
            If QTY_TO_COMMIT <> 0 Then
                Dim STYLE_CODE As String = .Item("STYLE_CODE")
                Dim COLOR_CODE As String = .Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY_TO_COMMIT)
            End If

        End With

        If S = -1 Then
        Else

            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim rowSOTORDR7 As DataRow = clsASFBASE1.Fill_Record("SOTORDR7_CANC", New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})

            If rowSOTORDR7 Is Nothing Then
                rowSOTORDR7 = dst.Tables("SOTORDR7_CANC").NewRow
                rowSOTORDR7.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                rowSOTORDR7.Item("STYLE_CODE") = STYLE_CODE
                rowSOTORDR7.Item("COLOR_CODE") = COLOR_CODE
                dst.Tables("SOTORDR7_CANC").Rows.Add(rowSOTORDR7)
            End If
            If rowSOTRSRV2.Item("RSRV_PRIORITY_DATE") & "" = "" Then
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV1.Item("INIT_DATE")).Date ' DateValue(Format(rowSOTRSRV1.Item("INIT_DATE"), "MM/dd/yyyy"))
            Else
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV2.Item("RSRV_PRIORITY_DATE")).Date '  DateValue(Format$(rowSOTRSRV2.Item("RSRV_PRIORITY_DATE"), "MM/DD/YYYY"))
            End If
            rowSOTORDR7.Item("ORDR_PRIORITY") = rowSOTRSRV2.Item("RSRV_PRIORITY")
            clsASFBASE1.Update_Record_TDA("SOTORDR7_CANC", "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
        End If
        clsASFBASE1.Update_Record_TDA("SOTRSRV_CANC2")

        ASCMAIN1.sql = "Select Sum (RSRV_QTY_OPEN) from SOTRSRV2 where RSRV_NO = :PARM1"
        Dim RSRV_QTY_OPEN_total As Int64 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {RSRV_NO}))

        If RSRV_QTY_OPEN_total = 0 Then
            rowSOTRSRV1.Item("RSRV_STATUS") = "F"
        Else
            rowSOTRSRV1.Item("RSRV_STATUS") = "O"
        End If
        clsASFBASE1.Update_Record_TDA("SOTRSRV1")
    End Sub


End Class
