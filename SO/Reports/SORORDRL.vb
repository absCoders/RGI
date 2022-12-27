Imports System.Xml
Imports System.IO
Imports System.Text
Imports System
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports Infragistics.UltraChart.Resources.Appearance
Imports System.Net.Mail

Public Class SORORDRL

#Region "Declarations"

    Dim EDI_APPOINTMENT As String
    Dim sqlo As String = ""

    Private clsTACENCRY As TAC.ASCENCRY
    Private EncryptionType As TAC.ASCENCRY.EncrytpionTypes = TAC.ASCENCRY.EncrytpionTypes.AdvancedEncryptionStandard_AES
    Public EncryptionCode As String = String.Empty
    Private NewQuotes As Boolean = False
    Private QuoteAbandonHours As Int64 = 48
    Private AbandonLiveDate As Date = CDate("07/24/2021")
    Private CCPA_NOs As New List(Of String)
#End Region

#Region "ABS Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")

        Dim rowTATENCRY As DataRow = ASCDATA1.GetDataRow("Select * from TATENCRY Where ENCRYPT_CODE = '" & EncryptionCode & "'")
        Dim rowASTPARMP As DataRow = ASCDATA1.GetDataRow("Select * from ASTPARMP WHERE AS_PARM_KEY = 'Z'")

        If EncryptionCode.Length > 0 AndAlso rowTATENCRY IsNot Nothing Then
            EncryptionType = DirectCast(CInt(Val(rowTATENCRY.Item("ENCRYPT_TYPE") & String.Empty)), TAC.ASCENCRY.EncrytpionTypes)
            clsTACENCRY = New TAC.ASCENCRY(EncryptionType)
            clsTACENCRY.Key = rowTATENCRY.Item("ENCRYPT_KEY") & String.Empty
            clsTACENCRY.PaddingMode = rowTATENCRY.Item("ENCRYPT_PADDING") & String.Empty
            clsTACENCRY.CipherMode = rowTATENCRY.Item("ENCRYPT_CIPHER") & String.Empty
        Else
            clsTACENCRY = New TAC.ASCENCRY()
        End If

        If rowASTPARMP Is Nothing OrElse Not rowASTPARMP.Table.Columns.Contains("AS_PARM_USE_ENCRYPTION") OrElse rowASTPARMP.Item("AS_PARM_USE_ENCRYPTION") & String.Empty <> "1" Then
            clsTACENCRY.UseEncryption = False
        Else
            clsTACENCRY.UseEncryption = True
        End If

        OrdersMissingDetailsFound()

        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("UPDATE")
        sql.AppendLine("SOTORDR1_L")
        sql.AppendLine("SET ORDR_CANCEL_DATE = TO_DATE(TO_CHAR(ORDR_CANCEL_DATE,'DD-MON-YYYY'))")
        sql.AppendLine("WHERE ORDR_NO IN")
        sql.AppendLine("(")
        sql.AppendLine("SELECT")
        sql.AppendLine("ORDR_NO")
        sql.AppendLine("FROM SOTORDR1_L")
        sql.AppendLine("WHERE NVL(ORDR_STATUS,'') <> 'X'")
        sql.AppendLine("AND TO_CHAR(ORDR_CANCEL_DATE,'HH24:MI') <> '00:00'")
        sql.AppendLine(")")
        ASCMAIN1.sql = sql.ToString
        ASCDATA1.ExecuteSQL()

        sql.Length = 0
        sql.AppendLine("UPDATE")
        sql.AppendLine("SOTORDR1_L")
        sql.AppendLine("SET ORDR_SHIP_DATE = TO_DATE(TO_CHAR(ORDR_SHIP_DATE,'DD-MON-YYYY'))")
        sql.AppendLine("WHERE ORDR_NO IN")
        sql.AppendLine("(")
        sql.AppendLine("SELECT")
        sql.AppendLine("ORDR_NO")
        sql.AppendLine("FROM SOTORDR1_L")
        sql.AppendLine("WHERE NVL(ORDR_STATUS,'') <> 'X'")
        sql.AppendLine("AND TO_CHAR(ORDR_SHIP_DATE,'HH24:MI') <> '00:00'")
        sql.AppendLine(")")
        ASCMAIN1.sql = sql.ToString
        ASCDATA1.ExecuteSQL()

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        If optORDR_SOURCE.Value = "Q" Then
            ASCMAIN1.sql = "Select * from SOTQRDRP"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTQRDRP", 1))

            ASCMAIN1.sql = "Select * from SOTSREP1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))

            'ASCMAIN1.sql = "Select * from ARTCUST1 WHERE CUST_CODE IN (SELECT CUST_CODE FROM SOTQRDR1)"
            ASCMAIN1.sql = "Select * from ARTCUST1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1", 1))

            ASCMAIN1.sql = "Select * from SOTORDR1_L"
            Create_TDA(dst.Tables.Add, "SOTORDR1_L", "**", 0, , , 1)
            dst.Tables("SOTORDR1_L").Columns.Add("ERRORS")
            dst.Tables("SOTORDR1_L").Columns.Add("EXCEPTIONS")

            ASCMAIN1.sql = "Select * from SOTORDR2_L"
            Create_TDA(dst.Tables.Add, "SOTORDR2_L", "**", 0, , , 2)
            dst.Tables("SOTORDR2_L").Columns.Add("ERRORS")
            dst.Tables("SOTORDR2_L").Columns.Add("EXCEPTIONS")

            ASCMAIN1.sql = "Select * from SOTORDR5_L"
            Create_TDA(dst.Tables.Add, "SOTORDR5_L", "**", 0, , , 2)

            ASCMAIN1.sql = "Select * from SOTQRDR1"
            Create_TDA(dst.Tables.Add, "SOTQRDR1", "**", 0, , , 1)
            dst.Tables("SOTQRDR1").Columns.Add("ERRORS")
            dst.Tables("SOTQRDR1").Columns.Add("EXCEPTIONS")

            ASCMAIN1.sql = "Select * from SOTQRDR2"
            Create_TDA(dst.Tables.Add, "SOTQRDR2", "**", 0, , , 2)
            dst.Tables("SOTQRDR2").Columns.Add("ERRORS")
            dst.Tables("SOTQRDR2").Columns.Add("EXCEPTIONS")

            ASCMAIN1.sql = "Select * from SOTQRDR5"
            Create_TDA(dst.Tables.Add, "SOTQRDR5", "**", 0, , , 2)
            Fill_Records("SOTQRDR1")
            Fill_Records("SOTQRDR2")
            Fill_Records("SOTQRDR5")
            FetchWebQuotes()

            RWU = "N"
        Else
            Dim SOURCE As String = ""

            If optORDR_SOURCE.Value = "A" Then
                SOURCE = "('L','T')"
            ElseIf optORDR_SOURCE.Value = "L" Then
                SOURCE = "('L')"
            ElseIf optORDR_SOURCE.Value = "T" Then
                SOURCE = "('T')"
            ElseIf optORDR_SOURCE.Value = "W" Then
                SOURCE = "('W')"
            End If

            Create_TDA(dst.Tables.Add("WBTCUST1"), "WBTCUST1", "*")

            If optORDR_SOURCE.Value = "W" Then
                If ShopSiteFileExists() Then
                    Dim iTitle As String = "Shopsite File Exists"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("A File Has Been Found That")
                    iMSG.AppendLine("Has Not Been Processed Yet.")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("That File Will Be Used Rather")
                    iMSG.AppendLine("Than Requesting New Orders From")
                    iMSG.AppendLine("Shopsite.")
                    MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
                Else
                    If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
                    FetchShopSiteOrders()
                End If
                If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
                If ErrorsInShopSiteFile() Then
                    If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
                    RWU = "N"
                    xErrMsg = "Fix Customer Matching"
                    Exit Sub
                Else
                    If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
                    CCPA_NOs.Clear()
                    ProcessShopSiteXML()
                End If
            End If
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            ' Mark all orders which are in SOTORDR1_L with ORDR_STATUS = 'O', with ORDR_BATCH_NO

            EDI_APPOINTMENT = ASCMAIN1.Next_Control_No("SOTORDR1.EDI_APPOINTMENT")
            ASCMAIN1.sql = "Update SOTORDR1_L Set EDI_APPOINTMENT = '" & EDI_APPOINTMENT & "' where ORDR_STATUS = 'O' AND ORDR_SOURCE IN " & SOURCE & " and CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1)"
            'ASCMAIN1.sql &= " and ROWNUM < 10"
            ASCDATA1.ExecuteSQL()

            sqlo = " where ORDR_NO in (Select ORDR_NO from SOTORDR1_L where EDI_APPOINTMENT = '" & EDI_APPOINTMENT & "')"

            ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_PRIORITY = (Select CUST_PRIORITY_CODE" & vbCrLf _
                & " from ARTCUST1 where CUST_CODE = SOTORDR1_L.CUST_CODE)" & sqlo
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_PRIORITY = NVL(ORDR_PRIORITY,'9')" & sqlo
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR1_L Set POST_CODE = (Select POST_CODE" & vbCrLf _
                & " from ARTCUST1 where CUST_CODE = SOTORDR1_L.CUST_CODE)" & sqlo
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_TYPE_CODE = 'REG'" & sqlo
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR1_L Set REASON_CODE = '', EDI_VALUE_CHANGE_DATE = '', ORDR_DATE_CLOSED = '' " & sqlo
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_TYPE_CODE = 'BTB', ORDR_HOLD = '1', ORDR_HOLD_REASON = 'LAPBTB'" & sqlo & " and WHSE_CODE in ('FE','FD','SP','NY','NC')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_TYPE_CODE = 'SAM'" & sqlo & " and WHSE_CODE in ('ZZ')"
            ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_HOLD = '1', ORDR_HOLD_REASON = DECODE(ORDR_HOLD_REASON,NULL,'',',') || 'SREP'" & sqlo & " and SREP_CODE not in (Select SREP_CODE from SOTSREP1)"
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_HOLD = '1', ORDR_HOLD_REASON = DECODE(ORDR_HOLD_REASON,NULL,'',',') || 'TERM'" & sqlo & " and TERM_CODE not in (Select TERM_CODE from TATTERM1)"
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_HOLD = '1', ORDR_HOLD_REASON = DECODE(ORDR_HOLD_REASON,NULL,'',',') || 'FRT'" & sqlo & " and FRT_TERMS not in ('PPD','COL','PPA')"
            'ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select * from SOTSREP1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))
            ASCMAIN1.sql = "Select * from TATTERM1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "TATTERM1", 1))
            ASCMAIN1.sql = "Select * from ARTCUST1 WHERE CUST_CODE IN (SELECT CUST_CODE FROM SOTORDR1_L WHERE ORDR_STATUS = 'O')"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1", 1))
            ASCMAIN1.sql = "Select * from ICTWHSE1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTWHSE1", 1))
            ASCMAIN1.sql = "Select * from ICTUOMF1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTUOMF1", 1))

            ASCMAIN1.sql = "Select * from SOTORDR1_L where ORDR_STATUS = 'O' AND ORDR_SOURCE IN " & SOURCE & " and CUST_CODE NOT IN (SELECT CUST_CODE FROM ARTCUST1)"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR1_C", 1))

            If dst.Tables.Contains("SOTORDR1_L") Then
                dst.Tables.Remove("SOTORDR1_L")
            End If

            ASCMAIN1.sql = "Select * from SOTORDR1_L" & sqlo
            Create_TDA(dst.Tables.Add, "SOTORDR1_L", "**", 0, , , 1)
            Fill_Records("SOTORDR1_L")
            dst.Tables("SOTORDR1_L").Columns.Add("ERRORS")
            dst.Tables("SOTORDR1_L").Columns.Add("EXCEPTIONS")

            ASCMAIN1.sql = "Select * from SOTORDR2_L" & sqlo
            Create_TDA(dst.Tables.Add, "SOTORDR2_L", "**", 0, , , 2)
            Fill_Records("SOTORDR2_L")
            dst.Tables("SOTORDR2_L").Columns.Add("ERRORS")
            dst.Tables("SOTORDR2_L").Columns.Add("EXCEPTIONS")

            ASCMAIN1.sql = "Select * from SOTORDR5_L" & sqlo
            Create_TDA(dst.Tables.Add, "SOTORDR5_L", "**", 0, , , 2)
            Fill_Records("SOTORDR5_L")

            Prepare_Order_Data()

            Check_if_Empty("SOTORDR1_L")
        End If
    End Sub

    Sub Prepare_Order_Data()

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1_L").Select("")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            Dim ERRORS As String = ""
            Dim EXCEPTIONS As String = ""
            Dim ORDR_HOLD_REASON = ""
            Dim SREP_CODE As String = rowSOTORDR1.Item("SREP_CODE") & ""
            Dim FRT_TERMS As String = rowSOTORDR1.Item("FRT_TERMS") & ""
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE") & ""
            Dim ORDR_TYPE_CODE As String = rowSOTORDR1.Item("ORDR_TYPE_CODE") & ""

            Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(SREP_CODE)
            If rowSOTSREP1 Is Nothing Then
                ERRORS &= vbCrLf & "Invalid Sales Rep Code"
                ORDR_HOLD_REASON &= "," & "SRep"
            End If

            Dim TERM_CODE As String = rowSOTORDR1.Item("TERM_CODE") & ""
            Dim rowTATTERM1 As DataRow = dst.Tables("TATTERM1").Rows.Find(TERM_CODE)
            If rowTATTERM1 Is Nothing Then
                ERRORS &= vbCrLf & "Invalid Terms Code"
                ORDR_HOLD_REASON &= "," & " AR Term"
            End If

            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & ""
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            If rowARTCUST1 Is Nothing Then
                ERRORS &= vbCrLf & "Invalid Customer Code"
                ORDR_HOLD_REASON &= "," & "Invalid Customer"
            End If

            Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
            If rowICTWHSE1 Is Nothing Then
                ERRORS &= vbCrLf & "Invalid Warehouse Code"
                ORDR_HOLD_REASON &= "," & "Whse"
            End If

            Dim row As DataRow = LookUp("ASTCODE1", New String() {"SOTORDR1", "FRT_TERMS", FRT_TERMS})
            If row Is Nothing Then
                ERRORS &= vbCr & "Invalid Freight Terms"
            End If

            If ORDR_TYPE_CODE <> "REG" And ORDR_TYPE_CODE <> "BTB" Then
                EXCEPTIONS &= vbCrLf & " Invalid Order Type "
            End If

            If rowARTCUST1.Item("FRT_TERMS") & "" <> rowSOTORDR1.Item("FRT_TERMS") & "" Then
                EXCEPTIONS &= vbCrLf & " Freight Terms "
            End If

            If rowARTCUST1.Item("SREP_CODE") & "" <> SREP_CODE Then
                EXCEPTIONS &= vbCrLf & " Check Sales Rep "
            End If


            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                'Stop
            End If
            Dim ORDR_TOTAL As Double = 0
            Dim ORDR_SOURCE As String = rowSOTORDR1.Item("ORDR_SOURCE") & ""
            If ORDR_SOURCE = "W" Then
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2_L").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
                    ORDR_TOTAL = ORDR_TOTAL + Val(rowSOTORDR2.Item("ORDR_QTY").ToString & "") * Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & "")
                Next
                If ORDR_TOTAL < 500 Then
                    Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                    SQLS.AppendLine("SELECT COUNT(*)")
                    SQLS.AppendLine("FROM SOTORDR1")
                    SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim ORDR_COUNT As Int16 = Val(ASCDATA1.GetDataValue)
                    If ORDR_COUNT <= 1 Then
                        EXCEPTIONS &= vbCrLf & " 1st Order under 500 "
                    End If
                End If
            End If

            rowSOTORDR1.Item("ERRORS") = Mid(ERRORS, 3)
            rowSOTORDR1.Item("EXCEPTIONS") = Mid(EXCEPTIONS, 3)
            If ERRORS <> "" Then
                rowSOTORDR1.Item("ORDR_HOLD") = "1"
                rowSOTORDR1.Item("ORDR_HOLD_REASON") = Mid(Mid(ORDR_HOLD_REASON, 2), 1, 20)
            End If


            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2_L").Select("ORDR_NO = '" & ORDR_NO & "'")
                Dim ERRORS2 As String = ""
                Dim EXCEPTIONS2 As String = ""
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                Dim UOM As String = rowSOTORDR2.Item("STYLE_UOM") & String.Empty
                If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    If UOM.Length = 0 Then
                        Stop
                        UOM = "PC"
                    End If
                End If
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                Dim VEND_CODE As String = rowICTSTYL1.Item("VEND_CODE") & ""

                Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                If rowARTCUST1 Is Nothing Then
                    ERRORS2 &= vbCrLf & "Invalid Item Color " & STYLE_CODE
                End If

                If ORDR_TYPE_CODE = "BTB" Then
                    Dim PO_COST As Decimal = TAC.SOCMAIN1.Get_PO_Cost(Me, STYLE_CODE, rowICTSTYL1.Item("VEND_CODE"), rowSOTORDR1)
                    rowSOTORDR2.Item("PO_COST") = PO_COST
                    If PO_COST = 0 Then
                        ERRORS2 &= vbCrLf & "PO Cost missing for " & STYLE_CODE
                    End If
                End If

                ' Dim rowICTUOMF1 As DataRow = dst.Tables("rowICTUOMF1").Rows.Find(UOM)
                Dim rowICTUOMF1 As DataRow = LookUp("ICTUOMF1", New String() {UOM})
                If rowICTUOMF1 Is Nothing Then
                    ERRORS &= vbCrLf & "Invalid Style uom"
                    ORDR_HOLD_REASON &= "," & "UOM"
                End If

                If rowSOTORDR2.Item("STYLE_UOM") <> rowICTSTYL1("STYLE_UOM") & "" Then
                    EXCEPTIONS2 &= vbCrLf & "Item UOM Error " & STYLE_CODE
                End If

                rowSOTORDR2.Item("STYLE_PRICE") = Val(rowICTSTYL1("STYLE_PRICE") & "")

                Dim ORDR_PRICE_SOURCE As String = ""
                Dim ORDR_UNIT_PRICE_CALC As Decimal = 0
                Dim ORDR_UNIT_PRICE_STD As Decimal = 0

                If ORDR_TYPE_CODE = "BTB" Then
                    Dim BTB_Price As New FEFDPrice(Me, STYLE_CODE, 1)
                    If WHSE_CODE = "FE" Then
                        ORDR_UNIT_PRICE_CALC = BTB_Price.FEPrice
                        ORDR_PRICE_SOURCE = "FE"
                    Else
                        ORDR_UNIT_PRICE_CALC = BTB_Price.FDPrice
                        ORDR_PRICE_SOURCE = "FD"
                    End If
                    ORDR_UNIT_PRICE_STD = BTB_Price.FEPrice
                    rowSOTORDR2("ORDR_UNIT_PRICE_STD") = ORDR_UNIT_PRICE_STD
                    BTB_Price = Nothing
                Else
                    ORDR_UNIT_PRICE_CALC = (TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1,
                       STYLE_CODE, COLOR_CODE,
                       Val(rowSOTORDR2("ORDR_QTY")), ORDR_PRICE_SOURCE))
                End If

                rowSOTORDR2("ORDR_UNIT_PRICE_CALC") = ORDR_UNIT_PRICE_CALC
                rowSOTORDR2("ORDR_PRICE_SOURCE") = ORDR_PRICE_SOURCE


                If Val(rowSOTORDR2("ORDR_UNIT_PRICE") & "") <> ORDR_UNIT_PRICE_CALC Then
                    rowSOTORDR2("ORDR_UNIT_PRICE_MANUAL") = "1"
                Else
                    rowSOTORDR2("ORDR_UNIT_PRICE_MANUAL") = "0"
                End If

                If rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL") & "" = "1" Then
                    EXCEPTIONS2 &= vbCrLf & " Net Pricing " & STYLE_CODE
                End If


                'RWU = "N" - TO PREVENT UPDATE
                rowSOTORDR2.Item("ERRORS") = Mid(ERRORS2, 3)
                rowSOTORDR2.Item("EXCEPTIONS") = Mid(EXCEPTIONS2, 3)
            Next

        Next
        Update_Record_TDA("SOTORDR1_L")
        Update_Record_TDA("SOTORDR2_L")
    End Sub

    Public Overrides Sub Print_Report()
        If optORDR_SOURCE.Value = "Q" Then
            If NewQuotes Then
                'SUBT = "ThenOrder Batch No " & EDI_APPOINTMENT
                Generate_Report("SORORDRQ", "New Quotes Imported From Web", SUBT)
            End If
        Else
            SUBT = "ThenOrder Batch No " & EDI_APPOINTMENT
            Generate_Report(RPT, , SUBT)
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                If OrdersMissingDetailsFound() Then
                    EMsg &= "Inbound Order Out Of Balance. Can Not Proceed."
                End If
        End Select
    End Sub

    Overrides Sub Update_Record()


        ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_YYYYPP_BOOKED = '" & ASCMAIN1.CYP & "', ORDR_GROUP_NO = ORDR_NO" & sqlo
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR2_L Set ORDR_STATUS = 'O'" & sqlo
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select ORDR_NO from SOTORDR1_L" & sqlo & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   SOPORDR1_L(R1.ORDR_NO);" & vbCrLf _
            & "   SOPORDR1_COMM(R1.ORDR_NO);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR1_L Set ORDR_STATUS = 'X'" & sqlo
        ASCDATA1.ExecuteSQL()

    End Sub

    'Sub Move_Rows(TABLE_NAME As String)

    '    Dim sqlw As String = " where ORDR_NO in (Select ORDR_NO from SOTORDR1_L where XNO = '" & XNO & "')"

    '    ASCMAIN1.sql = "Insert into " & TABLE_NAME & "_A Select * from " & TABLE_NAME & "_L" & sqlw
    '    ASCDATA1.ExecuteSQL()

    '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_A" & sqlw
    '    ASCDATA1.ExecuteSQL()

    'End Sub

#End Region

#Region "Custom Methods"
    Private Function OrdersMissingDetailsFound() As Boolean
        Dim RetVal As Boolean = False
        Dim ORDR_NOs As New List(Of String)
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT ORDR_NO, ORDR_DATE, ORDR_SOURCE, INIT_OPER, ORDR_STATUS")
        sql.AppendLine("FROM SOTORDR1_L")
        sql.AppendLine("WHERE ORDR_NO NOT IN")
        sql.AppendLine("(SELECT ORDR_NO FROM SOTORDR2_L)")
        sql.AppendLine("AND ORDR_STATUS = 'O'")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        For Each rowSOTORDR1_L As DataRow In tbl.Rows
            ORDR_NOs.Add(rowSOTORDR1_L.Item("ORDR_NO").ToString & String.Empty)
        Next
        If ORDR_NOs.Count > 0 Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Missing Details"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("The Following Inbound Orders")
            iMSG.AppendLine("Were Found To Be Missing Details.")
            iMSG.AppendLine("Please Contact Wayne From ABS")
            iMSG.AppendLine("And Let Him Know.")
            iMSG.AppendLine("")
            For Each ORDR_NO As String In ORDR_NOs
                iMSG.AppendLine("* - " & ORDR_NO)
            Next
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            RetVal = True
        End If
        Return RetVal
    End Function

#End Region

#Region "Fetch ShopSite Orders"

    Private Sub AddOrder2Records(ByVal nodeMain As XmlNode, ByVal ORDR_NO As String)
        Dim ORDR_LNO As Integer = 0
        For Each Node1 As XmlNode In nodeMain.ChildNodes
            If Node1.Name = "Shipping" Then
                For Each Node2 As XmlNode In Node1.ChildNodes
                    If Node2.Name = "Products" Then
                        For Each Node3 As XmlNode In Node2.ChildNodes
                            If Node3.Name = "Product" Then
                                Dim rowSOTORDR2_W As DataRow = dst.Tables("SOTORDR2_W").NewRow
                                ORDR_LNO += 1
                                rowSOTORDR2_W.Item("ORDR_NO") = ORDR_NO
                                rowSOTORDR2_W.Item("ORDR_LNO") = ORDR_LNO
                                Dim STYLE_CODE As String = ""
                                Dim COLOR_CODE As String = ""
                                Dim ORDR_QTY As Integer = 0
                                Dim ORDR_UNIT_PRICE As Double = 0.0
                                For Each Node4 As XmlNode In Node3.ChildNodes
                                    If Node4.Name = "SKU" Then
                                        Dim posDash As Integer = Node4.InnerText.IndexOf("-")
                                        Dim fullLenght As Integer = Node4.InnerText.Length
                                        STYLE_CODE = Node4.InnerText.Substring(0, posDash)

                                        Dim SQLS As New System.Text.StringBuilder
                                        SQLS.Length = 0
                                        SQLS.AppendLine(String.Format("Select Count(*) from ICTSTYL1 where STYLE_CODE = '{0}'", STYLE_CODE))
                                        ASCMAIN1.sql = SQLS.ToString()
                                        Dim STYLE_COUNT As Int16 = Val(ASCDATA1.GetDataValue)
                                        If STYLE_COUNT = 0 Then
                                            STYLE_CODE = STYLE_CODE.Substring(0, STYLE_CODE.Length / 2)
                                        End If

                                        COLOR_CODE = Node4.InnerText.Substring(posDash + 1, fullLenght - posDash - 1)
                                    End If
                                    If Node4.Name = "Quantity" Then
                                        If IsNumeric(Node4.InnerText) Then
                                            ORDR_QTY = Val(Node4.InnerText)
                                        End If
                                    End If
                                    If Node4.Name = "ItemPrice" Then
                                        If IsNumeric(Node4.InnerText) Then
                                            ORDR_UNIT_PRICE = Val(Node4.InnerText.Replace(",", ""))
                                        End If
                                    End If
                                Next

                                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                                rowSOTORDR2_W.Item("STYLE_CODE") = STYLE_CODE
                                rowSOTORDR2_W.Item("COLOR_CODE") = COLOR_CODE
                                rowSOTORDR2_W.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC").ToString & ""
                                rowSOTORDR2_W.Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY").ToString & ""
                                rowSOTORDR2_W.Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM").ToString & ""
                                rowSOTORDR2_W.Item("ORDR_EXTD_COST") = 0
                                rowSOTORDR2_W.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE
                                rowSOTORDR2_W.Item("ORDR_QTY") = ORDR_QTY
                                rowSOTORDR2_W.Item("ORDR_QTY_OPEN") = ORDR_QTY
                                rowSOTORDR2_W.Item("ORDR_QTY_PICK") = 0
                                rowSOTORDR2_W.Item("ORDR_QTY_SHIP") = 0
                                rowSOTORDR2_W.Item("ORDR_QTY_CANC") = 0
                                rowSOTORDR2_W.Item("ORDR_STATUS") = "W"
                                rowSOTORDR2_W.Item("ORDR_QTY_ORIG") = ORDR_QTY
                                rowSOTORDR2_W.Item("QTY_PER_PP") = 1
                                rowSOTORDR2_W.Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE
                                rowSOTORDR2_W.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY").ToString & ""
                                rowSOTORDR2_W.Item("ITEM_CODE") = String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE)
                                rowSOTORDR2_W.Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString & ""
                                rowSOTORDR2_W.Item("ORDR_UNIT_PRICE_MANUAL") = 0
                                dst.Tables("SOTORDR2_W").Rows.Add(rowSOTORDR2_W)
                            End If
                        Next
                    End If
                Next
            End If
        Next


    End Sub

    Private Function ErrorsInShopSiteFile() As Boolean
        Dim RetVal As Boolean = False
        ASCMAIN1.Progress("Checking XML Files For Errors", String.Empty)
        Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")

        Dim WB_PARM_ORDERS_DIR As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString
        Dim FileList As New List(Of String)
        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            Stop
            WB_PARM_ORDERS_DIR = "C:\Shared\Test"
        End If

        Dim ErrList As New List(Of String)
        For Each FileName As String In IO.Directory.GetFiles(WB_PARM_ORDERS_DIR, "*.xml")
            Dim doc As XmlDocument = New XmlDocument()
            doc.Load(FileName)
            FileList.Add(FileName.Replace(WB_PARM_ORDERS_DIR & "\", ""))
            Dim nodeShopSiteOrder As XmlNode = doc.SelectNodes("ShopSiteOrders")(0)
            For Each nodeMain As XmlNode In nodeShopSiteOrder.ChildNodes
                Select Case nodeMain.Name
                    Case "Response"
                        If nodeMain.InnerText <> "1success" And nodeMain.InnerText <> "2success" Then
                            ' One means Successful
                            Stop 'Error out here
                        End If
                    Case "Order"
                        Dim ORDR_NO_WEB As String = GetXMLNodeData(nodeMain, "ORDR_NO_WEB")
                        Dim rowARTCUST1 As DataRow = GetCustomer(nodeMain, ErrList, ORDR_NO_WEB)
                End Select
            Next
        Next
        If ErrList.Count > 0 Then
            RetVal = True
            Dim ssErr As New StringBuilder With {.Length = 0}
            For Each ErrorS As String In ErrList
                ssErr.AppendLine("* - " & ErrorS)
            Next
            Using frmmsg As New ASFMSGBF
                frmmsg.Show_Formatted_txt("Please Fix The Following Issues With Shopsite", ssErr.ToString, Me)
            End Using
        End If
        Return RetVal
    End Function

    Private Sub FetchShopSiteOrders_old()
        Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")

        Dim WB_PARM_ORDERS_DIR As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString '"C:\VS\VDI\Archive\RGO\XML\ORDERS\"
        'Dim WB_PARM_SITE_ORDERS_POST_URL As String = rowWBTPARM1.Item("WB_PARM_SITE_ORDERS_POST_URL").ToString '"/regency-rib/bo/db_xml.cgi?"
        'Dim WB_PARM_SITE_ORDERS_POST_URL As String = "https://brown.secure-host.com/cgi-regency-rib/bo/db_xml.cgi?"
        Dim WB_PARM_SITE_ORDERS_POST_URL As String = "https://teal.secure-host.com/cgi-regency-rib/bo/db_xml.cgi?"
        'Dim WB_PARM_SITE_ORDERS_POST_URL As String = "https://www.regency-rib.com:443/cgi-regency-rib/bo/db_xml.cgi?"
        Dim WB_PARM_LAST_SALES_ORDER As Integer = Val(rowWBTPARM1.Item("WB_PARM_LAST_SALES_ORDER").ToString)
        Dim WB_PARM_SITE_NAME As String = "www.regency-rib.com"
        Dim WB_PARM_SITE_USER = rowWBTPARM1.Item("WB_PARM_SITE_USER").ToString
        Dim WB_PARM_SITE_PWD As String = rowWBTPARM1.Item("WB_PARM_SITE_PWD").ToString

        Dim salesFile As String = WB_PARM_ORDERS_DIR & "\SO_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
        Dim salesOrders As String = String.Empty

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim dataInStream As Boolean = True

        ASCMAIN1.Progress("Getting Sales Orders From Web", String.Empty)

        Dim script As String = WB_PARM_SITE_ORDERS_POST_URL
        script &= "startorder=" & Val(WB_PARM_LAST_SALES_ORDER) + 1
        script &= "&pay=yes"
        script &= "&secure=1"

        'Host: 69.94.109.131:443
        ' Port 443 is the secure port for ShopSite
        Dim sendText As String = String.Empty
        'sendText &= "GET " & script & " HTTP/1.1" & Chr(10)
        sendText &= "GET " & script & Chr(10)
        'sendText &= "Host: " & WB_PARM_SITE_NAME & ":443" & Chr(13) & Chr(10)
        sendText &= "Host: " & WB_PARM_SITE_NAME & Chr(13) & Chr(10)
        Dim pwd As [Byte]() = System.Text.Encoding.ASCII.GetBytes(WB_PARM_SITE_USER & ":" & WB_PARM_SITE_PWD)
        sendText &= "Authorization: Basic " & Convert.ToBase64String(pwd) & Chr(13) & Chr(10) & Chr(10)

        Dim shopSiteResponse As String = String.Empty

        Using tcpClient As New System.Net.Sockets.TcpClient()

            Try
                'tcpClient.Connect(WB_PARM_SITE_NAME, 80)
                tcpClient.Connect(WB_PARM_SITE_NAME, 443)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Shop Site Orders", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try


            Using networkStream As Net.Sockets.NetworkStream = tcpClient.GetStream()

                ' Post the request
                Dim sendBytes As [Byte]() = System.Text.Encoding.ASCII.GetBytes(sendText)
                networkStream.Write(sendBytes, 0, sendBytes.Length)

                ' Read the NetworkStream into a byte buffer.
                Dim bytes(tcpClient.ReceiveBufferSize) As Byte

                Dim myReadBuffer(1024) As Byte
                Dim myCompleteMessage As StringBuilder = New StringBuilder()
                Dim numberOfBytesRead As Integer = 0

                If networkStream.CanRead Then

                    ' Incoming message may be larger than the buffer size.
                    ' need to pause to allow buffer to fill up
                    Do
                        numberOfBytesRead = networkStream.Read(myReadBuffer, 0, myReadBuffer.Length)
                        'System.Threading.Thread.Sleep(500)
                        myCompleteMessage.AppendFormat("{0}", Encoding.ASCII.GetString(myReadBuffer, 0, numberOfBytesRead))
                        System.Threading.Thread.Sleep(500)
                    Loop While networkStream.DataAvailable
                End If

                shopSiteResponse = myCompleteMessage.ToString

            End Using

        End Using

        If Not shopSiteResponse.Contains("ShopSiteOrders") Then Exit Sub

        salesOrders = shopSiteResponse.Substring(shopSiteResponse.Split("<?xml")(0).Length).Trim
        ' trim the shit characters at the end of the file
        Dim FirstChar As Integer = Val(InStr(salesOrders, "<ShopSiteOrders>")) - 1
        Dim lastChar As Integer = InStr(salesOrders, "</ShopSiteOrders>") + "</ShopSiteOrders>".Length

        If lastChar > salesOrders.Length Then
            lastChar = salesOrders.Length - FirstChar
        End If

        salesOrders = salesOrders.Substring(FirstChar, lastChar)
        'salesOrders = salesOrders.Trim

        salesOrders = salesOrders.Replace(">" & Chr(10) & "<", ">" & Environment.NewLine & "<")

        Using objReader As New StreamWriter(salesFile)
            objReader.Write(salesOrders)
            objReader.Close()
        End Using
    End Sub

    Private Sub FetchShopSiteOrders()
        Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")

        Dim WB_PARM_ORDERS_DIR As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString '"C:\VS\VDI\Archive\RGO\XML\ORDERS\"
        'Dim WB_PARM_SITE_ORDERS_POST_URL As String = rowWBTPARM1.Item("WB_PARM_SITE_ORDERS_POST_URL").ToString '"/regency-rib/bo/db_xml.cgi?"
        'Dim WB_PARM_SITE_ORDERS_POST_URL As String = "https://brown.secure-host.com/cgi-regency-rib/bo/db_xml.cgi?"
        'Dim WB_PARM_SITE_ORDERS_POST_URL As String = "https://teal.secure-host.com/cgi-regency-rib/bo/db_xml.cgi?"
        Dim WB_PARM_SITE_ORDERS_POST_URL As String = "https://www.regency-rib.com:443/cgi-regency-rib/bo/db_xml.cgi?"
        Dim WB_PARM_LAST_SALES_ORDER As Integer = Val(rowWBTPARM1.Item("WB_PARM_LAST_SALES_ORDER").ToString)
        Dim WB_PARM_SITE_NAME As String = "www.regency-rib.com"
        Dim WB_PARM_SITE_USER = rowWBTPARM1.Item("WB_PARM_SITE_USER").ToString
        Dim WB_PARM_SITE_PWD As String = rowWBTPARM1.Item("WB_PARM_SITE_PWD").ToString

        Dim salesFile As String = WB_PARM_ORDERS_DIR & "\SO_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
        Dim salesOrders As String = String.Empty

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim dataInStream As Boolean = True

        ASCMAIN1.Progress("Getting Sales Orders From Web", String.Empty)

        Dim script As String = WB_PARM_SITE_ORDERS_POST_URL
        script &= "startorder=" & Val(WB_PARM_LAST_SALES_ORDER) + 1
        script &= "&pay=yes"
        script &= "&version=12.0"

        Dim pwd As [Byte]() = Encoding.ASCII.GetBytes(WB_PARM_SITE_USER & ":" & WB_PARM_SITE_PWD)
        Dim Authorization As String = "Basic " & Convert.ToBase64String(pwd)

        Dim shopSiteResponse As String = String.Empty
        Using http As New HttpClient()
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                Stop
                Dim uri As String = "C:\Shared\RGI\InOrders\ORDR00001.xml"
                'http.DefaultRequestHeaders.Accept.Clear()
                'http.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                'Dim response As HttpResponseMessage = http.GetAsync(uri).Result

                'If Not response.IsSuccessStatusCode Then
                '    MessageBox.Show("Invalid Response from ShopSite: " & response.ToString, "Shop Site Orders", MessageBoxButtons.OK, MessageBoxIcon.Error)
                '    Exit Sub
                'End If

                shopSiteResponse = New StreamReader(uri).ReadToEnd
                'shopSiteResponse = response.Content.ReadAsStringAsync().Result
            Else
                System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
                Dim uri As String = script
                http.DefaultRequestHeaders.Accept.Clear()
                http.DefaultRequestHeaders.Add("Host", WB_PARM_SITE_NAME & ":443")
                http.DefaultRequestHeaders.Add("Authorization", Authorization)
                http.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                Dim response As HttpResponseMessage = http.GetAsync(uri).Result

                If Not response.IsSuccessStatusCode Then
                    MessageBox.Show("Invalid Response from ShopSite: " & response.ToString, "Shop Site Orders", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                shopSiteResponse = response.Content.ReadAsStringAsync().Result
            End If

        End Using


        'Dim shopSiteResponse As String = String.Empty

        'Using tcpClient As New System.Net.Sockets.TcpClient()

        '    Try
        '        'tcpClient.Connect(WB_PARM_SITE_NAME, 80)
        '        tcpClient.Connect(WB_PARM_SITE_NAME, 443)
        '    Catch ex As Exception
        '        MessageBox.Show(ex.Message, "Shop Site Orders", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        Exit Sub
        '    End Try


        '    Using networkStream As Net.Sockets.NetworkStream = tcpClient.GetStream()

        '        ' Post the request
        '        Dim sendBytes As [Byte]() = System.Text.Encoding.ASCII.GetBytes(sendText)
        '        networkStream.Write(sendBytes, 0, sendBytes.Length)

        '        ' Read the NetworkStream into a byte buffer.
        '        Dim bytes(tcpClient.ReceiveBufferSize) As Byte

        '        Dim myReadBuffer(1024) As Byte
        '        Dim myCompleteMessage As StringBuilder = New StringBuilder()
        '        Dim numberOfBytesRead As Integer = 0

        '        If networkStream.CanRead Then

        '            ' Incoming message may be larger than the buffer size.
        '            ' need to pause to allow buffer to fill up
        '            Do
        '                numberOfBytesRead = networkStream.Read(myReadBuffer, 0, myReadBuffer.Length)
        '                'System.Threading.Thread.Sleep(500)
        '                myCompleteMessage.AppendFormat("{0}", Encoding.ASCII.GetString(myReadBuffer, 0, numberOfBytesRead))
        '                System.Threading.Thread.Sleep(500)
        '            Loop While networkStream.DataAvailable
        '        End If

        '        shopSiteResponse = myCompleteMessage.ToString

        '    End Using

        'End Using

        'If Not shopSiteResponse.Contains("ShopSiteOrders") Then Exit Sub

        'salesOrders = shopSiteResponse.Substring(shopSiteResponse.Split("<?xml")(0).Length).Trim
        '' trim the shit characters at the end of the file
        'Dim FirstChar As Integer = Val(InStr(salesOrders, "<ShopSiteOrders>")) - 1
        'Dim lastChar As Integer = InStr(salesOrders, "</ShopSiteOrders>") + "</ShopSiteOrders>".Length

        'If lastChar > salesOrders.Length Then
        '    lastChar = salesOrders.Length - FirstChar
        'End If

        salesOrders = shopSiteResponse.Substring(shopSiteResponse.Split("<?xml")(0).Length).Trim
        ' trim the shit characters at the end of the file
        Dim lastChar = InStr(salesOrders, "</ShopSiteOrders>") + "</ShopSiteOrders>".Length

        If lastChar > salesOrders.Length Then
            lastChar = salesOrders.Length
        End If

        salesOrders = salesOrders.Substring(0, lastChar)
        salesOrders = salesOrders.Trim

        salesOrders = salesOrders.Replace(">" & Chr(10) & "<", ">" & Environment.NewLine & "<")
        ' convert single low-9 quotation mark to apostrophe
        salesOrders = salesOrders.Replace("&sbquo;", "&#39;")
        salesOrders = salesOrders.Replace("&euro;", "")


        'salesOrders = salesOrders.Substring(FirstChar, lastChar)
        'salesOrders = salesOrders.Trim

        'salesOrders = salesOrders.Replace(">" & Chr(10) & "<", ">" & Environment.NewLine & "<")

        Using objReader As New StreamWriter(salesFile)
            objReader.Write(salesOrders)
            objReader.Close()
        End Using
    End Sub

    Private Function GetCustomer(nodeMain As XmlNode, ByRef ErrList As List(Of String), ByVal ORDR_NO_WEB As String) As DataRow
        Dim Email As String = ""
        Dim SHOPSITE_CUST_ID As String = ""
        For Each OrderNode As XmlNode In nodeMain.ChildNodes
            If OrderNode.Name = "Billing" Then
                For Each BillingNode As XmlNode In OrderNode.ChildNodes
                    If BillingNode.Name = "Email" Then
                        If Email.Length = 0 Then
                            Email = BillingNode.InnerText
                        End If
                    End If
                Next
            End If
            If OrderNode.Name = "Other" Then
                For Each Node2 As XmlNode In nodeMain.ChildNodes
                    If Node2.Name = "Other" Then
                        For Each Node3 As XmlNode In Node2.ChildNodes
                            If Node3.Name = "CustomerID" Then
                                If SHOPSITE_CUST_ID.Length = 0 Then
                                    SHOPSITE_CUST_ID = Node3.InnerText
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next
        Dim rowWBTCUST1 As DataRow = Nothing
        Dim rowARTCUST1 As DataRow = Nothing
        If Email.Length = 0 And SHOPSITE_CUST_ID.Length = 0 Then
            ErrList.Add(String.Format("Order {0} Has No E-mail Or Customer ID Provided!", ORDR_NO_WEB))
        Else
            rowWBTCUST1 = GetCustomerFallBack(Email, SHOPSITE_CUST_ID)
            If IsNothing(rowWBTCUST1) Then
                ErrList.Add(String.Format("Order {0} Can Not Be Linked To A Customer with CustID:{1} or EMail:{2}", ORDR_NO_WEB, SHOPSITE_CUST_ID, Email))
            End If
        End If
        If ErrList.Count = 0 Then
            rowARTCUST1 = LookUp("ARTCUST1", rowWBTCUST1.Item("CUST_CODE_ACTUAL").ToString)
            If IsNothing(rowARTCUST1) Then
                ErrList.Add(String.Format("Order {0} Can Not Be Linked To A Customer with CustID:{1} or EMail:{2}", ORDR_NO_WEB, SHOPSITE_CUST_ID, Email))
            End If
        End If

        'If IsNothing(rowWBTCUST1) Then
        '    Dim msg As New Text.StringBuilder With {.Length = 0}
        '    msg.AppendLine("No Customer Contact For Email " & Email.ToUpper)
        '    msg.AppendLine("Please Enter A Valid Customer And I Will")
        '    msg.AppendLine("Create One.")
        '    Dim NEW_CUST_CODE As String = InputBox(msg.ToString(), "Bad E-mail")
        '    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        '    SQLS.AppendLine("SELECT MIN(EMAIL) AS EXISTINGEMAIL")
        '    SQLS.AppendLine("FROM WBTCUST1")
        '    SQLS.AppendLine(String.Format("WHERE CUST_CODE_ACTUAL = '{0}'", NEW_CUST_CODE))
        '    SQLS.AppendLine("AND STATUS = 'A'")
        '    ASCMAIN1.sql = SQLS.ToString()
        '    Dim EXISTINGEMAIL As String = ASCDATA1.GetDataValue
        '    If EXISTINGEMAIL.Length = 0 Then
        '        MsgBox("Problem With Customer Provided: " & NEW_CUST_CODE, MsgBoxStyle.Critical, "Can Not Proceed")
        '        Stop 'Warn and Bail out.
        '    Else
        '        dst.Tables.Item("WBTCUST1").Clear()
        '        Dim oldWBTCUST1 As DataRow = LookUp("WBTCUST1", EXISTINGEMAIL.ToUpper)
        '        rowWBTCUST1 = dst.Tables.Item("WBTCUST1").NewRow()
        '        For Each col As DataColumn In rowWBTCUST1.Table.Columns
        '            Select Case col.ColumnName
        '                Case Is = "EMAIL"
        '                    rowWBTCUST1.Item(col.ColumnName) = Email.ToUpper
        '                Case Is = "STATUS"
        '                    rowWBTCUST1.Item(col.ColumnName) = "T"
        '                Case Is = "INIT_OPER"
        '                    rowWBTCUST1.Item(col.ColumnName) = ASCMAIN1.USER_ID
        '                Case Is = "LAST_OPER"
        '                    rowWBTCUST1.Item(col.ColumnName) = ASCMAIN1.USER_ID
        '                Case Is = "INIT_DATE"
        '                    rowWBTCUST1.Item(col.ColumnName) = Now + ASCMAIN1.NowTSD
        '                Case Is = "LAST_DATE"
        '                    rowWBTCUST1.Item(col.ColumnName) = Now + ASCMAIN1.NowTSD
        '                Case Else
        '                    rowWBTCUST1.Item(col.ColumnName) = oldWBTCUST1.Item(col.ColumnName)
        '            End Select
        '        Next
        '        dst.Tables.Item("WBTCUST1").Rows.Add(rowWBTCUST1)
        '        Update_Record_TDA("WBTCUST1")
        '    End If
        'End If
        'Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowWBTCUST1.Item("CUST_CODE_ACTUAL").ToString)
        'If IsNothing(rowARTCUST1) Then
        '    MsgBox("Problem With Customer: " & Email.ToUpper, MsgBoxStyle.Critical, "Can Not Proceed")
        '    Stop 'Warn and Bail out.
        'End If
        Return rowARTCUST1
    End Function

    Private Function GetCustomerFallBack(ByVal EMAIL As String, ByVal SHOPSITE_CUST_ID As String) As DataRow
        Dim RetVal As DataRow = Nothing
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM WBTCUST1")
        sql.AppendLine("WHERE SHOPSITE_CUST_ID = :PARM1")
        Dim tblWBTCUST1 As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", SHOPSITE_CUST_ID)
        If tblWBTCUST1.Rows.Count = 1 Then
            RetVal = tblWBTCUST1.Rows(0)
        Else
            RetVal = LookUp("WBTCUST1", EMAIL.ToUpper)
        End If
        Return RetVal
    End Function

    Private Function GetXMLNodeData(nodeMain As XmlNode, NodePath As String) As String
        Dim RetVal As String = ""
        Select Case NodePath
            Case "ORDR_NO_WEB"
                For Each Node1 As XmlNode In nodeMain.ChildNodes
                    If Node1.Name = "OrderNumber" Then
                        RetVal = Node1.InnerText
                        Exit For
                    End If
                Next
            Case "ORDR_SHIP_DATE"
                For Each Node1 As XmlNode In nodeMain.ChildNodes
                    If Node1.Name = "Other" Then
                        For Each Node2 As XmlNode In Node1.ChildNodes
                            If Node2.Name = "CustomCheckoutField" Then
                                Dim FoundNode As Boolean = False
                                For Each Node3 As XmlNode In Node2.ChildNodes
                                    If Node3.Name = "FieldName" Then
                                        If Node3.InnerText = "Ship Date" Then
                                            FoundNode = True
                                        End If
                                    End If
                                    If (Node3.Name = "FieldValue") And FoundNode Then
                                        If IsDate(Node3.InnerText) Then
                                            RetVal = CDate(Node3.InnerText)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            Case "ORDR_CANCEL_DATE"
                For Each Node1 As XmlNode In nodeMain.ChildNodes
                    If Node1.Name = "Other" Then
                        For Each Node2 As XmlNode In Node1.ChildNodes
                            If Node2.Name = "CustomCheckoutField" Then
                                Dim FoundNode As Boolean = False
                                For Each Node3 As XmlNode In Node2.ChildNodes
                                    If Node3.Name = "FieldName" Then
                                        If Node3.InnerText = "Cancel Date" Then
                                            FoundNode = True
                                        End If
                                    End If
                                    If (Node3.Name = "FieldValue") And FoundNode Then
                                        If IsDate(Node3.InnerText) Then
                                            RetVal = CDate(Node3.InnerText)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            Case "ORDR_MESSAGE"
                For Each Node1 As XmlNode In nodeMain.ChildNodes
                    If Node1.Name = "Other" Then
                        For Each Node2 As XmlNode In Node1.ChildNodes
                            If Node2.Name = "OrderInstructions" Then
                                RetVal += Node2.InnerText
                            End If
                            If Node2.Name = "Comments" Then
                                RetVal = RetVal & vbCrLf & Node2.InnerText
                            End If
                        Next
                    End If
                Next
            Case Else
                RetVal = ""
        End Select
        Return RetVal
    End Function

    Private Sub ProcessShopSiteXML()
        ASCMAIN1.Progress("Processing XML Files", String.Empty)
        Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")

        Dim WB_PARM_ORDERS_DIR As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString '"C:\VS\VDI\Archive\RGO\XML\ORDERS\"
        Dim WB_PARM_ORDERS_DIR_OLD As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString & "\old"
        Dim LAST_ORDR_NO_WEB As String = ""
        Dim FileList As New List(Of String)
        Dim FileListMove As New List(Of String)
        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            Stop
            WB_PARM_ORDERS_DIR = "C:\Shared\Test"
        End If

        For Each FileName As String In IO.Directory.GetFiles(WB_PARM_ORDERS_DIR, "*.xml_e")
            ASCMAIN1.TACMAIN1.ShopSiteEncrypt("D", FileName, WB_PARM_ORDERS_DIR, WB_PARM_ORDERS_DIR_OLD)
        Next

        For Each FileName As String In IO.Directory.GetFiles(WB_PARM_ORDERS_DIR, "*.xml")
            Dim doc As XmlDocument = New XmlDocument()
            doc.Load(FileName)
            FileList.Add(FileName.Replace(WB_PARM_ORDERS_DIR & "\", ""))
            Dim nodeShopSiteOrder As XmlNode = doc.SelectNodes("ShopSiteOrders")(0)
            For Each nodeMain As XmlNode In nodeShopSiteOrder.ChildNodes
                Select Case nodeMain.Name
                    Case "Response"
                        If nodeMain.InnerText <> "1success" And nodeMain.InnerText <> "2success" Then
                            ' One means Successful
                            Stop 'Error out here
                        End If
                    Case "Order"
                        If Not dst.Tables.Contains("SOTORDR1_W") Then
                            'ASCMAIN1.sql = "Select * from SOTORDR1_L"
                            'Create_TDA(dst.Tables.Add, "SOTORDR1_W", "**", , True, "V", 1)
                            Create_TDA(dst.Tables.Add("SOTORDR1_W"), "SOTORDR1_L", "*")
                        End If
                        If Not dst.Tables.Contains("SOTORDR2_W") Then
                            'ASCMAIN1.sql = "Select * from SOTORDR2_L"
                            'Create_TDA(dst.Tables.Add, "SOTORDR2_W", "**", , True, "VI", 2)
                            Create_TDA(dst.Tables.Add("SOTORDR2_W"), "SOTORDR2_L", "*")
                        End If
                        If Not dst.Tables.Contains("SOTORDR5_W") Then
                            'ASCMAIN1.sql = "Select * from SOTORDR5_L"
                            'Create_TDA(dst.Tables.Add, "SOTORDR5_W", "**", , True, "VI", 2)
                            Create_TDA(dst.Tables.Add("SOTORDR5_W"), "SOTORDR5_L", "*")
                        End If

                        If Not dst.Tables.Contains("ARTCCPA1") Then
                            Create_TDA(dst.Tables.Add("ARTCCPA1"), "ARTCCPA1", "*")
                        End If

                        If Not dst.Tables.Contains("SOTORDC1") Then
                            Create_TDA(dst.Tables.Add("SOTORDC1"), "SOTORDC1", "*")
                        End If

                        Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                        ASCMAIN1.Progress("-", ORDR_NO)
                        Dim ORDR_NO_WEB As String = GetXMLNodeData(nodeMain, "ORDR_NO_WEB")
                        Dim rowSOTORDR1_W As DataRow = dst.Tables("SOTORDR1_W").NewRow
                        Dim rowSOTORDR5_W As DataRow = dst.Tables("SOTORDR5_W").NewRow
                        Dim ErrList As New List(Of String)
                        Dim rowARTCUST1 As DataRow = GetCustomer(nodeMain, ErrList, ORDR_NO_WEB)
                        If ErrList.Count > 0 Then
                            Stop 'This Should Have Been Pre-Vetted Before Getting here!
                        End If
                        SetRowDefaults(rowSOTORDR1_W, rowSOTORDR5_W, rowARTCUST1, ORDR_NO, nodeMain)

                        'Dim ORDR_NO_WEB As String = GetXMLNodeData(nodeMain, "ORDR_NO_WEB")
                        rowSOTORDR1_W.Item("ORDR_NO_WEB") = ORDR_NO_WEB
                        LAST_ORDR_NO_WEB = ORDR_NO_WEB

                        Dim ORDR_SHIP_DATE As String = GetXMLNodeData(nodeMain, "ORDR_SHIP_DATE")
                        If IsDate(ORDR_SHIP_DATE) Then
                            rowSOTORDR1_W.Item("ORDR_SHIP_DATE") = CDate(CDate(ORDR_SHIP_DATE).ToShortDateString)
                        Else
                            rowSOTORDR1_W.Item("ORDR_SHIP_DATE") = CDate(Now().ToShortDateString)
                        End If

                        Dim ORDR_CANCEL_DATE As String = GetXMLNodeData(nodeMain, "ORDR_CANCEL_DATE")
                        If IsDate(ORDR_CANCEL_DATE) Then
                            rowSOTORDR1_W.Item("ORDR_CANCEL_DATE") = CDate(CDate(ORDR_CANCEL_DATE).ToShortDateString)
                        Else
                            rowSOTORDR1_W.Item("ORDR_CANCEL_DATE") = CDate(Now().ToShortDateString)
                        End If

                        If (rowSOTORDR1_W.Item("ORDR_MESSAGE").ToString & String.Empty).Length > 0 Then
                            rowSOTORDR1_W.Item("ORDR_MESSAGE") = rowSOTORDR1_W.Item("ORDR_MESSAGE").ToString & String.Empty & vbCrLf & GetXMLNodeData(nodeMain, "ORDR_MESSAGE")
                        Else
                            rowSOTORDR1_W.Item("ORDR_MESSAGE") = GetXMLNodeData(nodeMain, "ORDR_MESSAGE")
                        End If
                        If rowSOTORDR1_W.Item("ORDR_MESSAGE").length > 500 Then
                            rowSOTORDR1_W.Item("ORDR_MESSAGE") = rowSOTORDR1_W.Item("ORDR_MESSAGE").substring(0, 500)
                        End If

                        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop 'Skip This in Dev Mode
                        AddCCRecords(nodeMain, ORDR_NO, rowSOTORDR1_W.Item("CUST_CODE").ToString, rowSOTORDR1_W)

                        dst.Tables("SOTORDR1_W").Rows.Add(rowSOTORDR1_W)
                        dst.Tables("SOTORDR5_W").Rows.Add(rowSOTORDR5_W)

                        AddOrder2Records(nodeMain, ORDR_NO)
                End Select
            Next
            doc.Save(FileName)
            Dim fTemp As String = ASCMAIN1.TACMAIN1.ShopSiteEncrypt("E", FileName, WB_PARM_ORDERS_DIR, WB_PARM_ORDERS_DIR_OLD)
        Next
        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop 'Skip Down To Update_Records
        'For Each FileMove As String In FileList
        '    System.IO.File.Move(String.Format("{0}\{1}", WB_PARM_ORDERS_DIR, FileMove), String.Format("{0}\{1}", WB_PARM_ORDERS_DIR_OLD, FileMove))
        'Next

        If LAST_ORDR_NO_WEB <> "" Then
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine("UPDATE WBTPARM1")
            SQLS.AppendLine(String.Format("SET WB_PARM_LAST_SALES_ORDER = {0}", LAST_ORDR_NO_WEB))
            SQLS.AppendLine("WHERE WB_PARM_KEY = 'Z'")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
        End If

        Update_Record_TDA("SOTORDR1_W")
        Update_Record_TDA("SOTORDR2_W")
        Update_Record_TDA("SOTORDR5_W")
        'EncryptARTCCPA1()
        Update_Record_TDA("ARTCCPA1")
        Update_Record_TDA("SOTORDC1")
        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop 'Skip Down To Update_Records

        'SELECT * 
        'From ARTCCPA1
        'Where NVL(CCPA_STATUS,'NULL') = 'NULL'
        'ORDER BY CCPA_NO DESC;

        'CREATE TABLE WHR_ARTCCPA1_221122 AS
        'SELECT 
        'ORDR_NO AS ORDR_NO_SORT,
        'ARTCCPA1.*
        'FROM ARTCCPA1 
        'WHERE ORDR_NO IN
        '(
        '    SELECT ORDR_NO 
        '    FROM SOTORDR1 
        '    WHERE NVL(ORDR_NO_WEB,'NULL') = 'NULL'
        '    AND ORDR_DATE >= '20-NOV-2022'
        ');

        For Each CC As String In CCPA_NOs
            ASCDATA1.ExecuteSQL("Begin ARTCCPA1_ARTCUSTC('" & CC & "'); End;")
        Next
    End Sub

    Private Sub AddCCRecords(ByVal nodeMain As XmlNode, ByVal ORDR_NO As String, ByVal CUST_CODE As String, ByRef rowSOTORDR1_W As DataRow)

        Dim IsCC As Boolean = False
        For Each OrderNode As XmlNode In nodeMain.ChildNodes
            If OrderNode.Name = "Payment" Then
                For Each PaymentNode As XmlNode In OrderNode.ChildNodes
                    If PaymentNode.Name = "CreditCard" Then
                        IsCC = True
                        Exit For
                    End If
                Next
            End If
        Next
        If Not IsCC Then
            Exit Sub
        End If

        Dim CC_Issuer As String = ""
        Dim CC_Number As String = ""
        Dim CC_NumberLast4 As String = ""
        Dim CC_VerificationValue As String = ""
        Dim CC_FullName As String = ""
        Dim CC_ExpirationDate As String = ""
        Dim CC_Street1 As String = ""
        Dim CC_Street2 As String = ""
        Dim CC_City As String = ""
        Dim CC_State As String = ""
        Dim CC_ZipCode As String = ""
        Dim CC_Country As String = ""
        Dim TERM_CODE As String = ""

        For Each OrderNode As XmlNode In nodeMain.ChildNodes
            If OrderNode.Name = "Payment" Then
                For Each PaymentNode As XmlNode In OrderNode.ChildNodes
                    If PaymentNode.Name = "CreditCard" Then
                        For Each CCNode As XmlNode In PaymentNode.ChildNodes
                            Select Case CCNode.Name
                                Case Is = "Issuer"
                                    CC_Issuer = CCNode.InnerText
                                    Select Case CC_Issuer
                                        Case Is = "American Express"
                                            CC_Issuer = "AMEX"
                                            TERM_CODE = "CRED"
                                        Case Is = "Visa"
                                            CC_Issuer = "VISA"
                                            TERM_CODE = "CRED"
                                        Case Is = "MasterCard"
                                            CC_Issuer = "MC"
                                            TERM_CODE = "CRED"
                                        Case Is = "Discover"
                                            CC_Issuer = "DISC"
                                            TERM_CODE = "CRED"
                                    End Select
                                Case Is = "Number"
                                    CC_Number = CCNode.InnerText
                                    If CC_Number.Length > 4 Then
                                        CC_NumberLast4 = CC_Number.Substring(CC_Number.Length - 4)
                                        'CCNode.InnerText = "Data Expunged"
                                    End If
                                Case Is = "VerificationValue"
                                    CC_VerificationValue = CCNode.InnerText
                                    'CCNode.InnerText = "Data Expunged"
                                Case Is = "FullName"
                                    CC_FullName = CCNode.InnerText
                                Case Is = "ExpirationDate"
                                    CC_ExpirationDate = CCNode.InnerText
                                    If CC_ExpirationDate.IndexOf("/") = 2 Then
                                        CC_ExpirationDate = String.Format("{0}{1}", CC_ExpirationDate.Substring(0, 2), CC_ExpirationDate.Substring(5, 2))
                                    ElseIf CC_ExpirationDate.IndexOf("/") = 1 Then
                                        CC_ExpirationDate = String.Format("0{0}{1}", CC_ExpirationDate.Substring(0, 1), CC_ExpirationDate.Substring(4, 2))
                                    End If
                                    'CCNode.InnerText = "Data Expunged"
                            End Select
                        Next
                    End If
                Next
            End If
        Next

        For Each OrderNode As XmlNode In nodeMain.ChildNodes
            If OrderNode.Name = "Billing" Then
                For Each BillingNode As XmlNode In OrderNode.ChildNodes
                    If BillingNode.Name = "Address" Then
                        For Each AddressNode As XmlNode In BillingNode.ChildNodes
                            Select Case AddressNode.Name
                                Case Is = "Street1"
                                    CC_Street1 = AddressNode.InnerText
                                Case Is = "Street2"
                                    CC_Street2 = AddressNode.InnerText
                                Case Is = "City"
                                    CC_City = AddressNode.InnerText
                                Case Is = "State"
                                    CC_State = AddressNode.InnerText
                                Case Is = "Code"
                                    CC_ZipCode = AddressNode.InnerText
                                Case Is = "Country"
                                    CC_Country = getCountryCode(CC_Country)
                            End Select
                        Next
                    End If
                Next
            End If
        Next

        Dim rowARTCCPA1 As DataRow = dst.Tables("ARTCCPA1").NewRow
        rowARTCCPA1.Item("CUST_CODE") = CUST_CODE
        rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = CC_Issuer
        'rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") = CC_Number 'This gets encrypted now.
        rowARTCCPA1.Item("CUST_CREDIT_CARD_NO_E") = ASCMAIN1.EncryptAES(CC_Number)
        If CC_FullName.Length > 35 Then 'This Is limited on the web side now but I am leaving it just in case.
            CC_FullName = CC_FullName.Substring(0, 34)
        End If
        rowARTCCPA1.Item("CUST_CREDIT_CARD_NAME") = CC_FullName
        rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1") = CC_Street1
        rowARTCCPA1.Item("CUST_CREDIT_CARD_CITY") = CC_City
        If CC_State.Length <= 2 Then
            rowARTCCPA1.Item("CUST_CREDIT_CARD_STATE") = CC_State
        Else
            rowARTCCPA1.Item("CUST_CREDIT_CARD_STATE") = LookUpState(CC_State)
        End If
        rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE") = CC_ZipCode
        rowARTCCPA1.Item("CUST_CREDIT_CARD_COUNTRY") = CC_Country
        rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") = CC_ExpirationDate
        If CC_VerificationValue.Length >= 3 And CC_VerificationValue.Length <= 4 Then
            'rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE") = CC_VerificationValue 'This gets encrypted now.
            rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE_E") = ASCMAIN1.EncryptAES(CC_VerificationValue)
        End If
        rowARTCCPA1.Item("CCPA_AMT") = 1
        Dim newCCPA_NO As String = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
        CCPA_NOs.Add(newCCPA_NO)
        rowARTCCPA1.Item("CCPA_NO") = newCCPA_NO
        rowARTCCPA1.Item("TRANS_NUM") = ASCMAIN1.Next_Control_No("ARTCCPA1.TRANS_NUM")
        rowARTCCPA1.Item("WEB_PYMT_ID") = ASCMAIN1.Next_Control_No("ARTCCPA1.WEB_PYMT_ID")
        rowARTCCPA1.Item("CUST_CREDIT_CARD_LAST4") = CC_NumberLast4
        rowARTCCPA1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
        rowARTCCPA1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        rowARTCCPA1.Item("CCPA_DATE_VOID") = Now + ASCMAIN1.NowTSD
        rowARTCCPA1.Item("ORDR_NO") = ORDR_NO
        dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1)

        Dim rowSOTORDC1 As DataRow = dst.Tables("SOTORDC1").NewRow
        rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
        rowSOTORDC1.Item("TRANS_NO") = rowARTCCPA1.Item("TRANS_NUM")
        rowSOTORDC1.Item("TRANS_TYPE") = "C"
        rowSOTORDC1.Item("TRANS_DATE") = Now + ASCMAIN1.NowTSD
        rowSOTORDC1.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
        rowSOTORDC1.Item("CCPA_STATUS") = "T"
        rowSOTORDC1.Item("AMOUNT") = 1
        rowSOTORDC1.Item("BALANCE") = 1
        rowSOTORDC1.Item("ACTIVE_IND") = 1
        rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)

        rowSOTORDR1_W.Item("TERM_CODE") = TERM_CODE
        rowSOTORDR1_W.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
        'Removed Per Ed on 5/22/14
        'rowSOTORDR1_W.Item("CC_TRANS_ID") = rowARTCCPA1.Item("TRANS_NUM")

    End Sub

    Private Function getCountryCode(ByVal CC_Country As String) As String
        Dim RetVal As String = "US"
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT COUNTRY_CODE2")
        SQLS.AppendLine("FROM TATCNTRY")
        SQLS.AppendLine($"WHERE UPPER(COUNTRY_NAME) = UPPER('{CC_Country}')")
        ASCMAIN1.sql = SQLS.ToString()
        Dim COUNTRY_CODE2 As String = ASCDATA1.GetDataValue
        If COUNTRY_CODE2.Length = 2 Then
            RetVal = COUNTRY_CODE2
        End If
        Return RetVal
    End Function

    Private Function LookUpState(ByVal CC_State As String) As String
        Dim RetVal As String = CC_State
        Dim SQLS As New StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT STATE_CODE")
        SQLS.AppendLine("FROM TATSTATE")
        SQLS.AppendLine(String.Format("WHERE UPPER(STATE_NAME) = UPPER('{0}')", CC_State))
        ASCMAIN1.sql = SQLS.ToString()
        Dim STATE_CODE As String = ASCDATA1.GetDataValue
        If STATE_CODE.Length > 0 Then
            RetVal = STATE_CODE
        End If
        If RetVal > 2 Then
            RetVal = ""
        End If
        Return RetVal
    End Function

    Private Sub SetRowDefaults(ByRef rowSOTORDR1_W As DataRow,
                               ByRef rowSOTORDR5_L As DataRow,
                               ByVal rowARTCUST1 As DataRow,
                               ByVal ORDR_NO As String,
                               ByVal nodeMain As XmlNode)
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", "MS")
        If IsNothing(rowICTWHSE1) Then
            Stop 'Warn and bail out
        End If
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("SELECT MIN(CUST_ADDR_CODE) AS CUST_ADDR_CODE")
        SQLS.AppendLine("FROM ARTCUST2")
        SQLS.AppendLine("WHERE CUST_CODE = '" & rowARTCUST1.Item("CUST_CODE").ToString & "" & "'")
        ASCMAIN1.sql = SQLS.ToString()
        Dim CUST_ADDR_CODE As String = ASCDATA1.GetDataValue
        Dim CUST_ADDR_CODE_SEL As String = CUST_ADDR_CODE
        Dim CUST_ADDR_MANUAL As Boolean = True
        Dim CUST_ADDR_TEXT As String = ""
        If Not IsNothing(nodeMain) Then
            For Each node1 As XmlNode In nodeMain
                If node1.Name = "Other" Then
                    For Each node2 As XmlNode In node1
                        If node2.Name = "CustomCheckoutField" Then
                            'This is Also Where We Are Going To Capture The New Questions Once That Task Comes Off Hold
                            Dim FieldName As String = ""
                            Dim FieldValue As String = ""
                            For Each node3 As XmlNode In node2
                                Select Case node3.Name
                                    Case "FieldName"
                                        FieldName = node3.InnerText
                                    Case "FieldValue"
                                        FieldValue = node3.InnerText
                                End Select
                            Next
                            If FieldName = "Customer Address Code" And FieldValue.Length > 0 Then
                                CUST_ADDR_CODE_SEL = FieldValue
                                CUST_ADDR_MANUAL = False
                            End If
                        End If
                    Next
                End If

                If node1.Name = "Shipping" Then
                    For Each node2 As XmlNode In node1
                        If node2.Name = "Address" Then
                            Dim Street1 As String = ""
                            Dim Street2 As String = ""
                            Dim City As String = ""
                            Dim State As String = ""
                            Dim Code As String = ""
                            Dim Country As String = ""
                            Dim ADDR_FOUND As Boolean = False
                            For Each node3 As XmlNode In node2
                                Select Case node3.Name
                                    Case "Street1"
                                        Street1 = node3.InnerText & String.Empty
                                        If Street1.Length > 0 Then
                                            ADDR_FOUND = True
                                        End If
                                    Case "Street2"
                                        Street2 = node3.InnerText & String.Empty
                                        If Street2.Length > 0 Then
                                            ADDR_FOUND = True
                                        End If
                                    Case "City"
                                        City = node3.InnerText & String.Empty
                                        If City.Length > 0 Then
                                            ADDR_FOUND = True
                                        End If
                                    Case "State"
                                        State = node3.InnerText & String.Empty
                                        If State.Length > 0 Then
                                            ADDR_FOUND = True
                                        End If
                                    Case "Code"
                                        Code = node3.InnerText & String.Empty
                                        If Code.Length > 0 Then
                                            ADDR_FOUND = True
                                        End If
                                    Case "Country"
                                        Country = node3.InnerText & String.Empty
                                        If Country.Length > 0 Then
                                            ADDR_FOUND = True
                                        End If
                                End Select
                            Next
                            If ADDR_FOUND Then
                                CUST_ADDR_TEXT = "CUSTOMER PROVIDED NEW SHIPPING ADDRESS:"
                                If Street1.Length > 0 Then
                                    CUST_ADDR_TEXT = CUST_ADDR_TEXT & vbCrLf & Street1.ToString & String.Empty
                                End If
                                If Street2.Length > 0 Then
                                    CUST_ADDR_TEXT = CUST_ADDR_TEXT & vbCrLf & Street2.ToString & String.Empty
                                End If
                                If CUST_ADDR_TEXT.Length > 0 Or City.Length > 0 Or Code.Length > 0 Then
                                    CUST_ADDR_TEXT = CUST_ADDR_TEXT & vbCrLf & City.ToString & String.Empty & ", " & State.ToString & String.Empty & " " & Code.ToString & String.Empty
                                End If
                                If Country.Length > 0 Then
                                    CUST_ADDR_TEXT = CUST_ADDR_TEXT & vbCrLf & Country.ToString & String.Empty
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {rowARTCUST1.Item("CUST_CODE").ToString & "", "MK", CUST_ADDR_CODE_SEL})
        If IsNothing(rowARTCUST2) Then
            rowARTCUST2 = LookUp("ARTCUST2", New String() {rowARTCUST1.Item("CUST_CODE").ToString & "", "MK", CUST_ADDR_CODE})
        End If
        Dim INST_COLS As String() = New String() {"CUST_ADDR_CODE", "CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}
        rowSOTORDR5_L.Item("ORDR_NO") = ORDR_NO
        rowSOTORDR5_L.Item("CUST_ADDR_TYPE") = "ST"
        For Each COL As String In INST_COLS
            rowSOTORDR5_L.Item(COL) = rowARTCUST2.Item(COL)
        Next
        rowSOTORDR1_W.Item("ORDR_NO") = ORDR_NO
        rowSOTORDR1_W.Item("ORDR_DATE") = Now().Date
        rowSOTORDR1_W.Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE").ToString & ""
        rowSOTORDR1_W.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME").ToString & ""
        rowSOTORDR1_W.Item("CUST_STORE_NO") = CUST_ADDR_CODE
        rowSOTORDR1_W.Item("CUST_STORE_NAME") = rowSOTORDR5_L.Item("CUST_NAME").ToString & ""
        rowSOTORDR1_W.Item("ORDR_FOB") = rowICTWHSE1.Item("WHSE_CITY").ToString & "," & rowICTWHSE1.Item("WHSE_STATE").ToString
        'rowSOTORDR1_W.Item("ORDR_CUST_PO") = "Web Order"
        rowSOTORDR1_W.Item("ORDR_CUST_PO") = ""
        rowSOTORDR1_W.Item("POST_CODE") = rowARTCUST1.Item("POST_CODE").ToString & ""
        rowSOTORDR1_W.Item("SHIP_VIA_CODE") = "BST"
        rowSOTORDR1_W.Item("ORDR_SHIP_INSTR") = ""
        rowSOTORDR1_W.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE").ToString & ""
        rowSOTORDR1_W.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE").ToString & ""
        rowSOTORDR1_W.Item("WHSE_CODE") = "MS"
        rowSOTORDR1_W.Item("SALES_DIVISION_CODE") = "RIB"
        rowSOTORDR1_W.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSOTORDR1_W.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowSOTORDR1_W.Item("INIT_DATE") = DATETIME_STAMP
        rowSOTORDR1_W.Item("LAST_DATE") = DATETIME_STAMP
        rowSOTORDR1_W.Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date
        rowSOTORDR1_W.Item("ORDR_SOURCE") = "W"
        rowSOTORDR1_W.Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS").ToString & ""
        rowSOTORDR1_W.Item("ORDR_ADDR_TYPE_ST") = "MK"
        rowSOTORDR1_W.Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
        rowSOTORDR1_W.Item("ORDR_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE").ToString & ""
        rowSOTORDR1_W.Item("ORDR_STATUS") = "O"
        rowSOTORDR1_W.Item("ORDR_GROUP_NO") = Null
        rowSOTORDR1_W.Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_CODE").ToString & ""
        rowSOTORDR1_W.Item("CURR_CODE") = "USD"
        rowSOTORDR1_W.Item("CURR_EXCH_RATE") = "1"
        rowSOTORDR1_W.Item("ORDR_TYPE_CODE") = "REG"
        rowSOTORDR1_W.Item("ORDR_SHIP_COMPLETE") = rowARTCUST1.Item("CUST_SHIP_COMPLETE").ToString & ""
        If CUST_ADDR_MANUAL And CUST_ADDR_TEXT.Length > 0 Then
            rowSOTORDR1_W.Item("ORDR_MESSAGE") = CUST_ADDR_TEXT
        End If
    End Sub

    Private Function ShopSiteFileExists() As Boolean
        Dim RetVal As Boolean = False
        Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")
        Dim WB_PARM_ORDERS_DIR As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString
        If Directory.Exists(WB_PARM_ORDERS_DIR) Then
            Dim Files As String() = Directory.GetFiles(WB_PARM_ORDERS_DIR)
            If Files.Length > 0 Then
                RetVal = True
            End If
        End If
        Return RetVal
    End Function

    Private Sub EncryptARTCCPA1()
        'If clsTACENCRY.UseEncryption = False Then
        '    Exit Sub
        'End If
        For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"}
                rowARTCCPA1.Item(field & "_E") = ASCMAIN1.EncryptAES(rowARTCCPA1.Item(field) & String.Empty) 'clsTACENCRY.EncryptString(rowARTCCPA1.Item(field) & String.Empty)
                rowARTCCPA1.Item(field) = DBNull.Value
            Next
        Next
    End Sub

#End Region

#Region "Fetch Web Quotes"

    Private Sub FetchWebQuotes()
        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then Stop

        Me.Cursor = Cursors.WaitCursor

        Dim eMsg As New System.Text.StringBuilder With {.Length = 0}
        Dim eTitle As String = "Error(s) Fetching Web Quotes"

        If eMsg.Length = 0 Then
            RefreshQuoteData(eMsg)
        Else
            Me.Cursor = Cursors.Default
            MsgBox(eMsg.ToString, vbCritical, eTitle)
            Exit Sub
        End If

        If eMsg.Length = 0 Then
            FetchQuoteData(eMsg)
        Else
            Me.Cursor = Cursors.Default
            MsgBox(eMsg.ToString, vbCritical, eTitle)
            Exit Sub
        End If

        If eMsg.Length = 0 Then
            PostQuoteToOra(eMsg)
        Else
            Me.Cursor = Cursors.Default
            MsgBox(eMsg.ToString, vbCritical, eTitle)
            Exit Sub
        End If

        'If eMsg.Length = 0 Then
        '    PrintQuoteReport(eMsg)
        'Else
        '    Me.Cursor = Cursors.Default
        '    MsgBox(eMsg.ToString, vbCritical, eTitle)
        '    Exit Sub
        'End If

        If eMsg.Length = 0 Then
            SendQuoteEmails(eMsg)
            SendAbandonEmails(eMsg)
        Else
            Me.Cursor = Cursors.Default
            MsgBox(eMsg.ToString, vbCritical, eTitle)
            Exit Sub
        End If

        Me.Cursor = Cursors.Default
        MsgBox("Web Quote Imports Complete", vbOKOnly, "Web Quotes")

        'Else
        '    Me.Cursor = Cursors.Default
        '    MsgBox("This Feature Not Finished Yet", vbExclamation, "Call Wayne And Yell At Him")
        'End If
    End Sub
    Private Sub RefreshQuoteData(ByRef eMsg As System.Text.StringBuilder)
        Try
            Dim UpdateURL As String = "https://www.regency-rib.com/quote/view.php"
            ASCMAIN1.Progress("Refreshing Quote Data", "")
            Dim WB As New WebBrowser
            WB.Navigate("")
            WB.Navigate(UpdateURL)

            System.Threading.Thread.Sleep(2000)
        Catch ex As Exception
            eMsg.AppendLine(ex.InnerException.ToString)
        End Try
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub FetchQuoteData(ByRef eMsg As System.Text.StringBuilder)
        Try
            Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")

            Dim WB_PARM_ORDERS_DIR As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString
            Dim WB_PARM_SITE_NAME As String = "www.regency-rib.com"
            Dim WB_PARM_SITE_USER = rowWBTPARM1.Item("WB_PARM_SITE_USER").ToString
            Dim WB_PARM_SITE_PWD As String = rowWBTPARM1.Item("WB_PARM_SITE_PWD").ToString
            Dim WB_PARM_SITE_FILE As String = rowWBTPARM1.Item("WB_PARM_SITE_OUTPUT_DIR").ToString & "/quote/data.csv"

            Dim FolderQuote As String = WB_PARM_ORDERS_DIR & "\quotes" '\S:\Archive\xml\orders
            Dim FolderArchive As String = FolderQuote & "\archives"
            Dim FileQuote As String = "data.csv"
            Dim FileArchive As String = String.Format("data{0}{1}{2}{3}{4}{5}.csv", Now().Year, Now().Month(), Now().Day, Now().Hour, Now().Minute, Now().Second)

            'Dim dataInStream As Boolean = True
            ASCMAIN1.Progress("Getting Quotes From Web", String.Empty)
            If IO.File.Exists(FolderQuote & "\" & FileQuote) Then
                IO.File.Move(FolderQuote & "\" & FileQuote, FolderArchive & "\" & FileArchive)
            End If

            Dim Sftp1 As New nsoftware.IPWorks.Ftp
            Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

            Sftp1.RemoteHost = "69.39.227.201"
            Sftp1.User = WB_PARM_SITE_USER
            Sftp1.Password = WB_PARM_SITE_PWD
            Sftp1.LocalFile = FolderQuote & "\" & FileQuote
            Sftp1.RemoteFile = WB_PARM_SITE_FILE
            Sftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmDefault

            Sftp1.Timeout = (20 * 60)
            Sftp1.Logoff()
            Sftp1.Logon()

            If Not Sftp1.Connected Then
                eMsg.AppendLine("Could not connect to ShopSite FTP")
                Exit Sub
            End If

            Sftp1.Passive = True
            Sftp1.Download()
            Sftp1.Logoff()
        Catch ex As Exception
            eMsg.AppendLine(ex.InnerException.ToString)
        End Try
    End Sub

    Private Sub PostQuoteToOra(ByRef eMsg As System.Text.StringBuilder)
        'ProblemQuotes.Clear()
        'AddProblemQuotes("David Blackmore", "11/09/2020", "09:14")
        'AddProblemQuotes("Stephen Ferrante", "29/09/2020", "12:01")
        Try
            Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")

            Dim WB_PARM_ORDERS_DIR As String = rowWBTPARM1.Item("WB_PARM_ORDERS_DIR").ToString
            'Dim WB_PARM_SITE_NAME As String = "www.regency-rib.com"
            'Dim WB_PARM_SITE_USER = rowWBTPARM1.Item("WB_PARM_SITE_USER").ToString
            'Dim WB_PARM_SITE_PWD As String = rowWBTPARM1.Item("WB_PARM_SITE_PWD").ToString
            'Dim WB_PARM_SITE_FILE As String = rowWBTPARM1.Item("WB_PARM_SITE_OUTPUT_DIR").ToString & "/quote/data.csv"

            Dim FolderQuote As String = WB_PARM_ORDERS_DIR & "\quotes"
            Dim FileQuote As String = "data.csv"

            'Dim lines As String() = IO.File.ReadAllLines(FolderQuote & "\" & FileQuote)

            Dim afile As FileIO.TextFieldParser = New FileIO.TextFieldParser(FolderQuote & "\" & FileQuote)
            Dim CurrentRecord As String()
            afile.TextFieldType = FileIO.FieldType.Delimited
            afile.Delimiters = New String() {","}
            afile.HasFieldsEnclosedInQuotes = True

            Dim LineNo As Int64 = 0
            Do While Not afile.EndOfData
                CurrentRecord = afile.ReadFields
                If LineNo > 0 Then

                    'Dim LineCur As String() = line.Split(","c)
                    'Dim newRow = tblData.Rows.Add()
                    Dim Status As String = CurrentRecord(0)
                    If Status = "complete" Or Status = "" Then 'Go Live With Abandoned Quotes.
                        'If Status = "complete" Then

                        Dim DateString As String = CurrentRecord(1)
                        If DateString.Length <> 10 Then
                            eMsg.AppendLine(String.Format("Invalid Date In Import File: {0}", DateString))
                            Exit Sub
                        End If
                        Dim DAY As String = DateString.Substring(0, 2)
                        Dim MON As String = DateString.Substring(3, 2)
                        Dim YR As String = DateString.Substring(6, 4)
                        If Not (IsNumeric(DAY) And IsNumeric(MON) And IsNumeric(YR)) Then
                            eMsg.AppendLine(String.Format("Invalid Date In Import File: {0}", DateString))
                            Exit Sub
                        End If
                        Dim ORDR_DATE As Date = DateSerial(Val(YR), Val(MON), Val(DAY))
                        Dim TimeString As String = CurrentRecord(2)
                        If Not IsDate(TimeString) Then
                            eMsg.AppendLine(String.Format("Invalid Time In Import File: {0}", TimeString))
                            Exit Sub
                        End If
                        Dim INIT_DATE As Date = ORDR_DATE.AddHours(CDate(TimeString).Hour).AddMinutes(CDate(TimeString).Minute)
                        Dim CustomerName As String = CurrentRecord(3).Replace(Chr(34), "")

                        If IsProblemQuote(CustomerName, DateString, TimeString) Then
                            Continue Do
                        End If

                        Dim EmailAddress As String = CurrentRecord(5)

                        Dim CUST_CODE As String = ""
                        Dim rowWBTCUST1 As DataRow = LookUp("WBTCUST1", EmailAddress.ToUpper)
                        If IsNothing(rowWBTCUST1) Then
                            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                                Continue Do
                            Else
                                eMsg.AppendLine(String.Format("Can Not Find Web Customer For {0}", EmailAddress))
                                Exit Sub
                            End If
                        End If
                        CUST_CODE = rowWBTCUST1.Item("CUST_CODE_ACTUAL").ToString & String.Empty
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If IsNothing(rowARTCUST1) Then
                            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                                Continue Do
                            Else
                                eMsg.AppendLine(String.Format("Can Not Find Web Customer For {0}", EmailAddress))
                                Continue Do
                            End If
                        End If

                        If Not QuoteHasItems(CurrentRecord) Then
                            'eMsg.AppendLine(String.Format("Quote Found For {0} With No Items", EmailAddress))
                            Continue Do
                        End If

                        If Status = "" Then
                            'Go Live With Abandoned Quotes.
                            If Not QuoteAbandoned(INIT_DATE) Then
                                Continue Do
                            End If
                            'Continue Do
                        End If

                        'Dim CompanyName As String = LineCur(4).Replace(Chr(34), "")
                        Dim EmailQuoteTo As String = CurrentRecord(6).Replace(Chr(34), "")
                        Dim FILTER As String = String.Format("CUST_CODE = '{0}' AND ORDR_DATE = '{1}' AND INIT_DATE = '{2}'", CUST_CODE, ORDR_DATE, INIT_DATE)
                        Dim rowSOTQRDR1 As DataRow = dst.Tables.Item("SOTQRDR1").Select(FILTER).FirstOrDefault
                        If IsNothing(rowSOTQRDR1) Then
                            Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                            ASCMAIN1.Progress("-", ORDR_NO)
                            Dim newSOTQRDR1 As DataRow = dst.Tables("SOTQRDR1").NewRow
                            Dim newSOTQRDR5 As DataRow = dst.Tables("SOTQRDR5").NewRow
                            SetRowDefaults(newSOTQRDR1, newSOTQRDR5, rowARTCUST1, ORDR_NO, Nothing)
                            If Status = "complete" Then
                                newSOTQRDR1.Item("ERRORS") = "NEW"
                            End If
                            'Go Live With Abandoned Quotes.
                            If Status = "" Then
                                newSOTQRDR1.Item("ERRORS") = "ABANDON"
                            End If

                            newSOTQRDR1.Item("ORDR_DATE") = ORDR_DATE
                            newSOTQRDR1.Item("ORDR_SHIP_DATE") = ORDR_DATE
                            newSOTQRDR1.Item("ORDR_CANCEL_DATE") = ORDR_DATE
                            newSOTQRDR1.Item("INIT_DATE") = INIT_DATE
                            If EmailQuoteTo.Length > 0 Then
                                newSOTQRDR1.Item("ORDR_MESSAGE") = String.Format("Please E-Mail Quote To: {0}", EmailQuoteTo)
                            End If

                            dst.Tables("SOTQRDR1").Rows.Add(newSOTQRDR1)
                            dst.Tables("SOTQRDR5").Rows.Add(newSOTQRDR5)
                            NewQuotes = True

                            Dim ORDR_LNO As Int64 = 0
                            For Lno As Int64 = 7 To CurrentRecord.Length Step 4
                                If Lno < CurrentRecord.Length Then
                                    ORDR_LNO += 1
                                    Dim newSOTQRDR2 As DataRow = dst.Tables("SOTQRDR2").NewRow
                                    newSOTQRDR2.Item("ORDR_NO") = ORDR_NO
                                    newSOTQRDR2.Item("ORDR_LNO") = ORDR_LNO
                                    Dim STYLE_CODE As String = CurrentRecord(Lno + 1).Substring(0, CurrentRecord(Lno + 1).IndexOf("-"))
                                    Dim COLOR_CODE As String = CurrentRecord(Lno + 1).Substring(CurrentRecord(Lno + 1).IndexOf("-") + 1, CurrentRecord(Lno + 1).Length - CurrentRecord(Lno + 1).IndexOf("-") - 1)
                                    Dim ORDR_UNIT_PRICE As Double = Val(CurrentRecord(Lno + 2))
                                    Dim ORDR_QTY As Integer = Val(CurrentRecord(Lno + 3))

                                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                                    newSOTQRDR2.Item("STYLE_CODE") = STYLE_CODE
                                    newSOTQRDR2.Item("COLOR_CODE") = COLOR_CODE
                                    newSOTQRDR2.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC").ToString & ""
                                    newSOTQRDR2.Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY").ToString & ""
                                    newSOTQRDR2.Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM").ToString & ""
                                    newSOTQRDR2.Item("ORDR_EXTD_COST") = 0
                                    newSOTQRDR2.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE
                                    newSOTQRDR2.Item("ORDR_QTY") = ORDR_QTY
                                    newSOTQRDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY
                                    newSOTQRDR2.Item("ORDR_QTY_PICK") = 0
                                    newSOTQRDR2.Item("ORDR_QTY_SHIP") = 0
                                    newSOTQRDR2.Item("ORDR_QTY_CANC") = 0
                                    newSOTQRDR2.Item("ORDR_STATUS") = "W"
                                    newSOTQRDR2.Item("ORDR_QTY_ORIG") = ORDR_QTY
                                    newSOTQRDR2.Item("QTY_PER_PP") = 1
                                    newSOTQRDR2.Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE
                                    newSOTQRDR2.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY").ToString & ""
                                    newSOTQRDR2.Item("ITEM_CODE") = String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE)
                                    newSOTQRDR2.Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString & ""
                                    newSOTQRDR2.Item("ORDR_UNIT_PRICE_MANUAL") = 0
                                    dst.Tables("SOTQRDR2").Rows.Add(newSOTQRDR2)
                                End If

                            Next
                            'AddOrder2Records(nodeMain, ORDR_NO)
                        End If
                    End If
                End If
                LineNo += 1
            Loop

            If NewQuotes Then
                Update_Record_TDA("SOTQRDR1")
                Update_Record_TDA("SOTQRDR2")
                Update_Record_TDA("SOTQRDR5")
            Else
                eMsg.AppendLine("No new Quotes Found")
            End If

        Catch ex As Exception
            eMsg.AppendLine(ex.InnerException.ToString)
        End Try
    End Sub

    Private Function QuoteAbandoned(ByVal INIT_DATE As Date) As Boolean
        Dim RetVal As Boolean = False
        If INIT_DATE > AbandonLiveDate Then
            If DateDiff(DateInterval.Hour, INIT_DATE, Now()) > QuoteAbandonHours Then
                RetVal = True
            End If
        End If
        Return RetVal
    End Function

    Private Function QuoteHasItems(ByRef currentRecord() As String) As Boolean
        '07 - Product 1
        '08 - SKU 1
        '09 - Price 1
        '10 - Quantity 1
        Dim HasItem As Boolean = True
        If currentRecord.Length < 10 Then
            HasItem = False
        Else
            For I As Int64 = 7 To 10
                If currentRecord(I).ToString & String.Empty = "" Then
                    HasItem = False
                End If
            Next
            For I As Int64 = 9 To 10
                If Not IsNumeric(currentRecord(9).ToString & String.Empty) Then
                    HasItem = False
                End If
            Next
        End If
        Return HasItem
    End Function

    Private Sub SendQuoteEmails(ByRef eMsg As System.Text.StringBuilder)
        'Throw New NotImplementedException()
        Try
            Dim ORDR_FILTER As String = "ERRORS = 'NEW'"

            For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select(ORDR_FILTER)
                Dim ORDR_NO As String = rowSOTQRDR1.Item("ORDR_NO").ToString & String.Empty
                Dim CUST_CODE As String = rowSOTQRDR1.Item("CUST_CODE").ToString & String.Empty
                Dim CUST_NAME As String = rowSOTQRDR1.Item("CUST_NAME").ToString & String.Empty

                Dim SREP_CODE As String = rowSOTQRDR1.Item("SREP_CODE").ToString & String.Empty
                Dim SREP_FILTER As String = String.Format("SREP_CODE = '{0}'", SREP_CODE)
                Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Select(SREP_FILTER).FirstOrDefault

                If IsNothing(rowSOTSREP1) Then
                    eMsg.AppendLine(String.Format("Invalid Sales Rep Code: {0}", SREP_CODE))
                Else
                    Dim SREP_NAME As String = rowSOTSREP1.Item("SREP_NAME").ToString & String.Empty
                    Dim SREP_EMAIL As String = rowSOTSREP1.Item("SREP_EMAIL").ToString & String.Empty

                    'Per Danny, Send All Quotes To Rita Now
                    If SREP_CODE = "HO" Then
                        SREP_EMAIL = "rita@regency-rib.com"
                    End If

                    Dim SUBJECT As String = "Quote Request Recieved On Regency Website"
                    Dim EBODY As New System.Text.StringBuilder With {.Length = 0}
                    EBODY.AppendLine(String.Format("Dear {0}", SREP_NAME))
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("We have received the following quote request on Regency's Website:")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine(String.Format("Quote #: {0}", ORDR_NO))
                    EBODY.AppendLine("<br>")
                    EBODY.AppendLine(String.Format("Customer Code: {0}", CUST_CODE))
                    EBODY.AppendLine("<br>")
                    EBODY.AppendLine(String.Format("Customer Name: {0}", CUST_NAME))
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("You can review this quote in the Laptop Sales Order Entry Program by performing a master data transfer after waiting 15 minutes from the receipt of this e-mail.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("Once Completed, you will be able to activate the quote on your laptop in the data transfer screen.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("After reviewing the quote you must contact the customer by email or phone or both and send them an official quote from you with pictures for the customers review.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("Following up immediately is of the utmost importance.  We need all of you to inform customer service that you have taken care of the customer by sending them a quote ASAP.   We have created a screen which will be monitored and updated when the quote is taken care of.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("If you have any questions, please contact customer service.")
                    'Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email(ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, SUBJECT, "SORORDRL", True, False, "", "", "", EMAIL_BODY)
                    Dim mail As New MailMessage()
                    mail.IsBodyHtml = True

                    mail.From = New MailAddress("hq@regency-rib.com", "Regency International")
                    mail.To.Add(New MailAddress("whr@waynerichmond.net", "Wayne Richmond"))
                    mail.To.Add(New MailAddress("mariog@regency-rib.com", "Mario Arenas Jr."))
                    'If (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "mariog") Then
                    '    Don 't Sent Srep E-mails Yet?
                    '    Dim iResult As MsgBoxResult
                    '    Dim iTitle As String = "E-mails"
                    '    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    '    iMSG.AppendLine("Do You Want Email Sent To " & SREP_NAME)
                    '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    '    If iResult = MsgBoxResult.Yes Then
                    '        mail.To.Add(New MailAddress(SREP_EMAIL, SREP_NAME))
                    '    End If
                    'Else
                    If SREP_EMAIL.Length > 0 Then
                        If Not (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne") Then
                            mail.To.Add(New MailAddress(SREP_EMAIL, SREP_NAME))
                        End If
                    End If

                    'End If
                    'mail.CC.Add(New String("whr@waynerichmond.net", "Wayne Richmond"))
                    mail.Subject = SUBJECT
                    mail.Body = EBODY.ToString
                    Dim smtp As New SmtpClient("192.168.110.221", 25)
                    smtp.Credentials = New System.Net.NetworkCredential("", "")
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network

                    Dim retry As Boolean = True
                    Dim retrys As Integer = 0
                    While retry
                        Try
                            If ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne") Then
                                Stop
                            End If
                            smtp.Send(mail)
                            retrys = 0
                            retry = False
                        Catch ex As Exception
                            retrys += 1
                            If retrys > 10 Then
                                eMsg.AppendLine("Maximum Retrys (10) Exceeded")
                                retry = False
                            End If
                        End Try
                    End While


                End If
            Next
        Catch ex As Exception
            eMsg.AppendLine(ex.InnerException.ToString)
        End Try
    End Sub

    Private Sub SendAbandonEmails(ByRef eMsg As System.Text.StringBuilder)
        'Throw New NotImplementedException()
        Try
            Dim ORDR_FILTER As String = "ERRORS = 'ABANDON'"

            For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select(ORDR_FILTER)
                Dim ORDR_NO As String = rowSOTQRDR1.Item("ORDR_NO").ToString & String.Empty
                Dim OFILTER As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
                Dim CUST_CODE As String = rowSOTQRDR1.Item("CUST_CODE").ToString & String.Empty
                Dim CUST_NAME As String = rowSOTQRDR1.Item("CUST_NAME").ToString & String.Empty

                Dim SREP_CODE As String = rowSOTQRDR1.Item("SREP_CODE").ToString & String.Empty
                Dim SREP_FILTER As String = String.Format("SREP_CODE = '{0}'", SREP_CODE)
                Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Select(SREP_FILTER).FirstOrDefault

                If IsNothing(rowSOTSREP1) Then
                    eMsg.AppendLine(String.Format("Invalid Sales Rep Code: {0}", SREP_CODE))
                Else
                    Dim SREP_NAME As String = rowSOTSREP1.Item("SREP_NAME").ToString & String.Empty
                    Dim SREP_EMAIL As String = rowSOTSREP1.Item("SREP_EMAIL").ToString & String.Empty

                    'Per Danny, Send All Quotes To Rita Now
                    If SREP_CODE = "HO" Then
                        SREP_EMAIL = "rita@regency-rib.com"
                    End If

                    Dim SUBJECT As String = "Abandoned Quote Found On Regency Website"
                    Dim EBODY As New System.Text.StringBuilder With {.Length = 0}
                    EBODY.AppendLine(String.Format("Dear {0}", SREP_NAME))
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine(String.Format("We have identified the following quote that was not finalized on Regency's Website for {0} hours:", QuoteAbandonHours))
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine(String.Format("Quote #: {0}", ORDR_NO))
                    EBODY.AppendLine("<br>")
                    EBODY.AppendLine(String.Format("Customer Code: {0}", CUST_CODE))
                    EBODY.AppendLine("<br>")
                    EBODY.AppendLine(String.Format("Customer Name: {0}", CUST_NAME))
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("You can review this quote in the Laptop Sales Order Entry Program by performing a master data transfer after waiting 15 minutes from the receipt of this e-mail.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("Once Completed, you will be able to activate the quote on your laptop in the data transfer screen.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("After reviewing the quote you must contact the customer by email or phone or both and send them an official quote from you with pictures for the customers review.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("Following up immediately is of the utmost importance.  We need all of you to inform customer service that you have taken care of the customer by sending them a quote ASAP.   We have created a screen which will be monitored and updated when the quote is taken care of.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("If you have any questions, please contact customer service.")
                    EBODY.AppendLine("<br><br>")
                    EBODY.AppendLine("<table>")
                    EBODY.AppendLine("<tr>")
                    EBODY.AppendLine(" <th>Style</th>")
                    EBODY.AppendLine(" <th>Color</th>")
                    EBODY.AppendLine(" <th>Description</th>")
                    EBODY.AppendLine(" <th>Qty</th>")
                    EBODY.AppendLine(" <th>Price</th>")
                    EBODY.AppendLine("</tr>")
                    For Each rowSOTQRDR2 As DataRow In dst.Tables("SOTQRDR2").Select(OFILTER, "ORDR_LNO")
                        EBODY.AppendLine("<tr>")
                        EBODY.AppendLine(String.Format("  <td>{0}</td>", rowSOTQRDR2.Item("STYLE_CODE").ToString & String.Empty))
                        EBODY.AppendLine(String.Format("  <td>{0}</td>", rowSOTQRDR2.Item("COLOR_CODE").ToString & String.Empty))
                        EBODY.AppendLine(String.Format("  <td>{0}</td>", rowSOTQRDR2.Item("STYLE_DESC").ToString & String.Empty))
                        EBODY.AppendLine(String.Format("  <td>{0}</td>", rowSOTQRDR2.Item("ORDR_QTY").ToString & String.Empty))
                        EBODY.AppendLine(String.Format("  <td>{0}</td>", rowSOTQRDR2.Item("ORDR_UNIT_PRICE").ToString & String.Empty))
                        EBODY.AppendLine("</tr>")
                    Next
                    EBODY.AppendLine("</table>")

                    'Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email(ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, SUBJECT, "SORORDRL", True, False, "", "", "", EMAIL_BODY)
                    Dim mail As New MailMessage()
                    mail.IsBodyHtml = True

                    mail.From = New MailAddress("hq@regency-rib.com", "Regency International")
                    mail.To.Add(New MailAddress("whr@waynerichmond.net", "Wayne Richmond"))
                    mail.To.Add(New MailAddress("mariog@regency-rib.com", "Mario Arenas Jr."))
                    'If (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "mariog") Then
                    '    Don't Sent Srep E-mails Yet?
                    '    Dim iResult As MsgBoxResult
                    '    Dim iTitle As String = "E-mails"
                    '    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    '    iMSG.AppendLine("Do You Want Email Sent To " & SREP_NAME)
                    '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    '    If iResult = MsgBoxResult.Yes Then
                    '        mail.To.Add(New MailAddress(SREP_EMAIL, SREP_NAME))
                    '    End If
                    'Else
                    If SREP_EMAIL.Length > 0 Then
                        If Not (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne") Then
                            mail.To.Add(New MailAddress(SREP_EMAIL, SREP_NAME))
                        End If
                    End If

                    'End If
                    'mail.CC.Add(New String("whr@waynerichmond.net", "Wayne Richmond"))
                    mail.Subject = SUBJECT
                    mail.Body = EBODY.ToString
                    Dim smtp As New SmtpClient("192.168.110.221", 25)
                    smtp.Credentials = New System.Net.NetworkCredential("", "")
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network

                    Dim retry As Boolean = True
                    Dim retrys As Integer = 0
                    While retry
                        Try
                            If ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne") Then
                                Stop
                            End If
                            smtp.Send(mail)
                            retrys = 0
                            retry = False
                        Catch ex As Exception
                            retrys += 1
                            If retrys > 10 Then
                                eMsg.AppendLine("Maximum Retrys (10) Exceeded")
                                retry = False
                            End If
                        End Try
                    End While


                End If
            Next
        Catch ex As Exception
            eMsg.AppendLine(ex.InnerException.ToString)
        End Try
    End Sub

    Public Function IsProblemQuote(ByVal CustomerName As String, ByVal DateString As String, ByVal TimeString As String) As Boolean
        Dim RetVal As Boolean = False
        For Each rowSOTQRDRP As DataRow In dst.Tables("SOTQRDRP").Select()
            Dim DATE_STRING As String = rowSOTQRDRP.Item("DATE_STRING").ToString & String.Empty
            Dim TIME_STRING As String = rowSOTQRDRP.Item("TIME_STRING").ToString & String.Empty
            Dim CUST_STRING As String = rowSOTQRDRP.Item("CUST_STRING").ToString & String.Empty
            If CUST_STRING = CustomerName And DATE_STRING = DateString And TIME_STRING = TimeString Then
                RetVal = True
            End If
        Next
        Return RetVal
    End Function

    Private Sub btnEDcrypt_Click(sender As Object, e As EventArgs) Handles btnEncrypt.Click, btnDecrypt.Click
        Dim btn As Button = sender
        Dim WAY As String = ""
        Dim PWord As String = "ShopEncrypt22"
        Dim folder As String = "S:\Archive\xml\orders\"
        Dim ext As String = ""
        Dim title As String = ""
        Dim iMsg As New StringBuilder With {.Length = 0}
        If txtEncryptPass.Text <> PWord Then
            iMsg.Length = 0
            iMsg.AppendLine("Password Does Not Match.")
            iMsg.AppendLine("Thanks For Playing.")
            MsgBox(iMsg.ToString, vbOKOnly, "Invalid Password")
            Exit Sub
        End If
        If btn.Name = "btnEncrypt" Then
            WAY = "E"
            title = "Select Your File For Encryption"
            ext = "xml"
        End If
        If btn.Name = "btnDecrypt" Then
            WAY = "D"
            title = "Select Your File For Encryption"
            ext = "xml_e"
        End If
        If WAY = "E" Or WAY = "D" Then
            Dim ofd As OpenFileDialog = New OpenFileDialog
            ofd.DefaultExt = ext
            'ofd.FileName = "defaultname"
            ofd.InitialDirectory = folder
            ofd.Filter = $"Shopsite|*.{ext}"
            ofd.Title = title
            If ofd.ShowDialog() <> DialogResult.Cancel Then
                If ofd.FileName.EndsWith(ext) Then
                    Dim SelFolder = ofd.FileName.Replace(ofd.SafeFileName, "")
                    ASCMAIN1.TACMAIN1.ShopSiteEncrypt(WAY, ofd.FileName, SelFolder, SelFolder)
                    title = "Conversion Complete!"
                    iMsg.Length = 0
                    iMsg.AppendLine("Encryption/Decryption Complete!")
                    iMsg.AppendLine("Please Note that Removal Of The")
                    iMsg.AppendLine("Original File Is Your Responsibility!")
                    iMsg.AppendLine("")
                    iMsg.AppendLine("Make Sure You Do Not Create Security")
                    iMsg.AppendLine("Issues By Leaving Un-Encrypted Data")
                    iMsg.AppendLine("Exposed.")
                    MsgBox(iMsg.ToString, vbOKOnly, title)
                Else
                    MsgBox("Invalid File extension", vbExclamation, "No Good")
                End If
            End If
        End If

    End Sub
#End Region
End Class