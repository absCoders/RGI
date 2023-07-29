Public Class SOCMAINL

    Public Shared Function IsValidTerms(ByVal CUST_CODE As String, ByVal SEL_TERM_CODE As String) As Boolean
        Dim RetVal As Boolean = False
        Dim sql As New Text.StringBuilder
        sql.Length = 0
        sql.AppendLine("SELECT NVL(TERM_CODE,'CRED') AS TERM_CODE")
        sql.AppendLine("FROM ARTCUST1")
        sql.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
        ASCMAIN1.sql = sql.ToString
        Dim CUST_TERM_CODE As String = ASCDATA1.GetDataValue
        If SEL_TERM_CODE = CUST_TERM_CODE Then
            RetVal = True
        Else
            Select Case CUST_TERM_CODE
                Case Is = "N30", "N30D", "N30ROG", "N45D", "N60", "N90", "N90D"
                    Select Case SEL_TERM_CODE
                        Case Is = "N30", "N30D", "N30ROG", "N45D", "N60", "N90", "N90D", "COD", "CBD", "CRED", "XMAS", "FALL"
                            RetVal = True
                        Case Else
                            MsgBox("Invalid Terms Code For This Customer", MsgBoxStyle.OkOnly, "Invalid Terms")
                            RetVal = False
                    End Select
                Case Is = "CRED", "COD", "CBD"
                    Select Case SEL_TERM_CODE
                        Case Is = "CRED", "COD", "CBD"
                            RetVal = True
                        Case Else
                            MsgBox(String.Format("Customers With {0} Terms Code Must Select CRED, COD or CBD Terms", CUST_TERM_CODE), MsgBoxStyle.OkOnly, "Invalid Terms")
                            RetVal = False
                    End Select
                Case Is = "AMEX"
                    MsgBox("AMEX Is No Longer Supported As A Terms Code" & CUST_TERM_CODE, MsgBoxStyle.OkOnly, "Invalid Terms")
                    RetVal = False
                Case Else
                    MsgBox("Customer Terms Code Is " & CUST_TERM_CODE, MsgBoxStyle.OkOnly, "Invalid Terms")
                    RetVal = False
            End Select
        End If
        Return RetVal
    End Function

    Private Function Pop_Control_No(ByVal TABLE_NAME As String, ByVal COLUMN_NAME As String) As String
        Dim RetVal As String = ""
        Dim VERBTYPE As String = ""
        Dim MSG As String = ""
        Dim RecsRemaining As Integer = 0
        Dim MAXREC As Integer = 0
        Dim MINREC As Integer = 0
        Dim CTLLENGTH As Integer = 0
        Select Case TABLE_NAME
            Case "SOTORDR1"
                VERBTYPE = "Order Numbers"
            Case "ARTCUST1"
                VERBTYPE = "Customer Numbers"
        End Select

        Dim sql As New Text.StringBuilder
        sql.Length = 0
        sql.AppendLine("SELECT MAX(CTL_NO_NEXT) AS MAXREC")
        sql.AppendLine(" FROM TATCTLNL")
        sql.AppendLine(String.Format(" WHERE TABLE_NAME = '{0}'", TABLE_NAME))
        sql.AppendLine(String.Format(" AND COLUMN_NAME = '{0}'", COLUMN_NAME))
        ASCMAIN1.sql = sql.ToString
        MAXREC = Val(ASCDATA1.GetDataValue)
        sql.Length = 0
        sql.AppendLine("SELECT MIN(CTL_NO_NEXT) AS MINREC")
        sql.AppendLine(" FROM TATCTLNL")
        sql.AppendLine(String.Format(" WHERE TABLE_NAME = '{0}'", TABLE_NAME))
        sql.AppendLine(String.Format(" AND COLUMN_NAME = '{0}'", COLUMN_NAME))
        ASCMAIN1.sql = sql.ToString
        MINREC = Val(ASCDATA1.GetDataValue)
        sql.Length = 0
        sql.AppendLine("SELECT MAX(CTL_NO_LENGTH) as CTLLENGTH")
        sql.AppendLine(" FROM TATCTLNL")
        sql.AppendLine(String.Format(" WHERE TABLE_NAME = '{0}'", TABLE_NAME))
        sql.AppendLine(String.Format(" AND COLUMN_NAME = '{0}'", COLUMN_NAME))
        ASCMAIN1.sql = sql.ToString
        CTLLENGTH = Val(ASCDATA1.GetDataValue)
        RecsRemaining = MAXREC - MINREC
        If RecsRemaining = 0 Then
            MSG = String.Format("You Have No More {0} Remaining.", VERBTYPE)
            MSG = MSG & vbCrLf & "Please Fetch More Using The Button Available In The Transfer Screen"
            RetVal = "NONE"
        Else
            If RecsRemaining < 10 And RecsRemaining > 0 Then
                MSG = String.Format("You Only Have {0} {1} Remaining.", RecsRemaining, VERBTYPE)
                MSG = MSG & vbCrLf & "Please Fetch More Using The Button Available In The Transfer Screen"
            End If
            RetVal = Format$(MINREC, "".PadLeft(CTLLENGTH, "0"))

        End If
        Return RetVal
    End Function

    Public Shared Function SalesReportCanRun(ByVal START_DATE As Date,
                                      ByVal END_DATE As Date,
                                      ByVal CK_UPDATES As Boolean,
                                      ByVal CK_TARIFFS As Boolean) As String
        Dim RetVal As New System.Text.StringBuilder With {.Length = 0}
        Dim S As New System.Text.StringBuilder With {.Length = 0}
        Dim START_DATE_ORA As String = Format(START_DATE, "dd-MMM-yyyy")
        Dim END_DATE_ORA As String = Format(END_DATE, "dd-MMM-yyyy")
        If CK_UPDATES Then
            S.Length = 0
            S.AppendLine("SELECT DISTINCT INV_DATE")
            S.AppendLine("FROM SOTINVH1")
            S.AppendLine("WHERE ORDR_YYYYPP_UPDATED is Null")
            S.AppendLine(String.Format("AND INV_DATE >= '{0}'", START_DATE_ORA))
            S.AppendLine(String.Format("AND INV_DATE <= '{0}'", END_DATE_ORA))
            ASCMAIN1.sql = S.ToString()
            Dim INV_DATE As String = ASCDATA1.GetDataValue
            If IsDate(INV_DATE) Then
                RetVal.AppendLine("There Are Billed Orders That Need To Be Run Through Sales Journal.")
            End If
        End If
        If CK_TARIFFS Then
            S.Length = 0
            S.AppendLine("SELECT MAX(INIT_DATE) AS INIT_DATE")
            S.AppendLine("FROM ICTCOSTP")
            ASCMAIN1.sql = S.ToString()
            Dim INIT_DATE As String = ASCDATA1.GetDataValue
            If IsDate(INIT_DATE) Then
                If END_DATE > CDate(CDate(INIT_DATE).ToShortDateString) Then
                    RetVal.AppendLine(String.Format("Tariffs Costs Have Only Been Run Through {0}", Format(CDate(INIT_DATE), "MM/dd/yyyy")))
                End If
            End If
        End If
        Return RetVal.ToString
    End Function

    Public Shared Function ImportDetailsFromExcel(ByVal frm As ASFBASE0) As Text.StringBuilder
        'This has been moved to SOFORDR1
        Dim RetVal As New Text.StringBuilder With {.Length = 0}
        Return RetVal
        Exit Function

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim ORDR_LNO As Int64 = 0
        Dim ORDR_NO As String = ""
        Dim CUST_CODE As String = ""

        If frm.dst.Tables.Contains("SOTORDR1") Then
            If frm.dst.Tables("SOTORDR1").Rows.Count = 1 Then
                rowSOTORDR1 = frm.dst.Tables("SOTORDR1").Rows(0)
                ORDR_NO = rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty
                CUST_CODE = rowSOTORDR1.Item("CUST_CODE").ToString & String.Empty
            Else
                MsgBox("Error in Form.  Please Let ABS Know", vbCritical, "Hmm")
                Return RetVal
                Exit Function
            End If
        Else
            MsgBox("Error in Form.  Please Let ABS Know", vbCritical, "Hmm")
            Return RetVal
            Exit Function
        End If
        If frm.dst.Tables.Contains("SOTORDR2") Then
            If frm.dst.Tables("SOTORDR2").Rows.Count > 0 Then
                Dim filter As String = ""
                ORDR_LNO = Val(frm.dst.Tables("SOTORDR2").Compute("max(ORDR_LNO)", filter)) + 1
            Else
                ORDR_LNO = 1
            End If
        Else
            MsgBox("Error in Form.  Please Let ABS Know", vbCritical, "Hmm")
            Return RetVal
            Exit Function
        End If
        Dim rowARTCUST1 As DataRow = frm.LookUp("ARTCUST1", CUST_CODE)

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            'Dim filter As String = "xlsb files (*.xlsb)|*.xlsx|All files (*.*)|*.*"
            Dim filter As String = "All files (*.*)|*.*"
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            ASCMAIN1.Progress("Now Building Order From Excel", "")
            frm.Cursor = Cursors.WaitCursor
            Dim COLUMNS As New Dictionary(Of String, Int64)
            Dim COLIST As New List(Of String)
            COLIST.Add(("Style Code").ToUpper)
            COLIST.Add(("Color Code").ToUpper)
            COLIST.Add(("Order Qty").ToUpper)
            COLIST.Add(("Price").ToUpper)
            COLIST.Add(("Cust SKU").ToUpper)
            COLIST.Add(("Cust Style").ToUpper)
            COLIST.Add(("Cust Color").ToUpper)

            'DATETIME_STAMP = Now + ASCMAIN1.NowTSD
            Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
            Dim xws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            xws = XWB.Worksheets(1)
            Dim BeginFound As Boolean = False
            Dim EndFound As Boolean = False
            Dim BlankRows As Int64 = 0
            For CurRow As Int64 = 1 To 2000
                If BeginFound And Not EndFound Then
                    If IsNothing(xws.Cells(CurRow, 1).value) Then
                        EndFound = True
                    Else
                        Dim STYLE_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "STYLE_CODE")
                        Dim COLOR_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "COLOR_CODE")
                        Dim CUST_SKU As String = GetValueFromExcel(xws, COLUMNS, CurRow, "CUST_SKU")
                        Dim CUST_STYLE_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "CUST_STYLE_CODE")
                        Dim CUST_COLOR_CODE As String = GetValueFromExcel(xws, COLUMNS, CurRow, "CUST_COLOR_CODE")
                        Dim QTY_STR As String = GetValueFromExcel(xws, COLUMNS, CurRow, "ORDR_QTY")
                        Dim PRICE_STR As String = GetValueFromExcel(xws, COLUMNS, CurRow, "ORDR_UNIT_PRICE")
                        Dim QTY As Int64 = 0
                        If IsNumeric(QTY_STR) Then
                            QTY = Val(QTY_STR)
                        End If
                        Dim PRICE As Decimal = 0.00
                        If IsNumeric(PRICE_STR) Then
                            PRICE = Val(PRICE_STR)
                        End If
                        Dim eMsg As New Text.StringBuilder With {.Length = 0}
                        'Stop
                        'Add Valiidation Here.
                        Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", STYLE_CODE)

                        If eMsg.Length > 0 Then
                            Dim err As New Text.StringBuilder With {.Length = 0}
                            err.AppendLine("Errors Found In File:")
                            err.AppendLine(eMsg.ToString)
                            err.AppendLine("")
                            err.AppendLine("Continue?")
                            Dim iResult As MsgBoxResult = MsgBox(err.ToString, vbYesNo, "Errors")
                            If Not iResult = MsgBoxResult.Yes Then
                                EndFound = True
                            End If
                        Else
                            Dim rowSOTORDR2 As DataRow = Nothing
                            rowSOTORDR2 = frm.dst.Tables("SOTORDR2").NewRow
                            Dim ORDR_UNIT_PRICE_CALC As Decimal = TAC.SOCMAIN1.Price_Line(frm, CUST_CODE, rowARTCUST1,
                                       STYLE_CODE, COLOR_CODE, QTY, "")
                            With rowSOTORDR2
                                .Item("ORDR_NO") = ORDR_NO
                                .Item("ORDR_LNO") = ORDR_LNO
                                .Item("STYLE_CODE") = STYLE_CODE
                                .Item("COLOR_CODE") = COLOR_CODE
                                .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC").ToString & String.Empty
                                .Item("ORDR_QTY") = QTY
                                .Item("ORDR_QTY_OPEN") = QTY
                                .Item("ORDR_QTY_ORIG") = QTY
                                .Item("ORDR_QTY_ALLO") = 0
                                .Item("INNER_PACK_QTY") = 0
                                .Item("ORDR_EXTD_COST") = 0
                                .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM").ToString & String.Empty
                                .Item("ORDR_QTY_PICK") = 0
                                .Item("ORDR_QTY_SHIP") = 0
                                .Item("ORDR_QTY_CANC") = 0
                                .Item("ORDR_STATUS") = "O"
                                .Item("ORDR_QTY_PRE_ALLO") = 0
                                .Item("QTY_PER_PP") = 0
                                .Item("CARTON_PACK_QTY") = 0
                                .Item("STYLE_PRICE") = 0
                                .Item("STYLE_RETAIL") = 0
                                .Item("PO_COST") = 0
                                .Item("COMM_RATE") = 0
                                .Item("ORDR_UNIT_PRICE_CALC") = ORDR_UNIT_PRICE_CALC
                                If PRICE <> 0 Then
                                    .Item("ORDR_UNIT_PRICE") = PRICE
                                    .Item("ORDR_UNIT_PRICE_CURR") = PRICE
                                    .Item("ORDR_UNIT_PRICE_MANUAL") = "1"
                                Else
                                    .Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_CALC
                                    .Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_CALC
                                End If
                                If CUST_SKU.Length > 0 Then
                                    .Item("CUST_SKU") = CUST_SKU
                                End If
                                If CUST_STYLE_CODE.Length > 0 Then
                                    .Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                                End If
                                If CUST_COLOR_CODE.Length > 0 Then
                                    .Item("CUST_COLOR_CODE") = CUST_COLOR_CODE
                                End If
                            End With
                            frm.dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
                            ORDR_LNO = ORDR_LNO + 1
                        End If
                    End If
                Else
                    BlankRows += 1
                    If BlankRows >= 10 Then
                        Exit For
                    End If
                    For R As Int64 = 1 To 7
                        If Not IsNothing(xws.Cells(CurRow, R).value) Then
                            Dim COL As String = xws.Cells(CurRow, R).value.ToString.ToUpper
                            If COLIST.Contains(COL) Then
                                COLUMNS.Add(COL, R)
                                BeginFound = True
                            End If
                        End If
                    Next
                    If BeginFound Then
                        If Not (COLUMNS.ContainsKey(("Style Code").ToUpper) And COLUMNS.ContainsKey(("Color Code").ToUpper)) Then
                            BeginFound = False
                        End If
                    End If
                End If
            Next
            xws = Nothing
            XWB = Nothing
            excel = Nothing
        End If
        Return RetVal
    End Function

    Private Shared Function GetValueFromExcel(xws As Microsoft.Office.Interop.Excel.Worksheet, ByVal COLUMNS As Dictionary(Of String, Int64), curRow As Long, ByVal CODE As String) As String
        Dim RetVal As String = ""
        Select Case CODE
            Case "STYLE_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Style Code").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Style Code").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "COLOR_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Color Code").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Color Code").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "ORDR_QTY"
                RetVal = "0"
                If COLUMNS.ContainsKey(("Order Qty").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Order Qty").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "ORDR_UNIT_PRICE"
                RetVal = "0"
                If COLUMNS.ContainsKey(("Price").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Price").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "CUST_SKU"
                RetVal = ""
                If COLUMNS.ContainsKey(("Cust SKU").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Cust SKU").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "CUST_STYLE_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Cust Style").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Cust Style").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
            Case "CUST_COLOR_CODE"
                RetVal = ""
                If COLUMNS.ContainsKey(("Cust Color").ToUpper) Then
                    Dim curCol As Int64 = COLUMNS.Item(("Cust Color").ToUpper)
                    If Not IsNothing(xws.Cells(curRow, curCol).value) Then
                        RetVal = xws.Cells(curRow, curCol).value.ToString.ToUpper
                    End If
                End If
        End Select

        Return RetVal
    End Function
End Class
