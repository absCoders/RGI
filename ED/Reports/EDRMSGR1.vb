
Public Class EDRMSGR1

    Private EDI_DOC_SEQ_NO_812 As String = String.Empty
    Private EDI_DOC_SEQ_NO_816 As String = String.Empty
    Private EDI_DOC_SEQ_NO_824 As String = String.Empty
    Private EDI_DOC_SEQ_NO_860 As String = String.Empty
    Private EDI_DOC_SEQ_NO_864 As String = String.Empty
    Private EDI_DOC_SEQ_NO_947 As String = String.Empty
    Private COMPANY_CODE As String = String.Empty
    Private tblEDT860T1 As String = String.Empty

    Private Const RegencyWalmartCustCode As String = "231551"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            COMPANY_CODE = "VAN"
        ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            COMPANY_CODE = "NYA"
        ElseIf ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            COMPANY_CODE = "RGI"
        Else
            COMPANY_CODE = ASCMAIN1.CLIENT
        End If

        Dim tblEDTTRPM1 As DataTable = ASCDATA1.GetDataTable("SELECT DISTINCT EDI_DOC_NO FROM EDTTRPM1")

        chkEDT812.Enabled = False
        chkEDT812.Checked = False

        chkEDT816.Enabled = False
        chkEDT816.Checked = False

        chkEDT824.Enabled = False
        chkEDT824.Checked = False

        chkEDT860.Enabled = False
        chkEDT860.Checked = False

        chkEDT864.Enabled = False
        chkEDT864.Checked = False

        chkEDT947.Enabled = False
        chkEDT947.Checked = False

        For Each row As DataRow In tblEDTTRPM1.Select("", "EDI_DOC_NO")
            Select Case row.Item("EDI_DOC_NO") & String.Empty

                Case "812"
                    If COMPANY_CODE = "NYA" Then
                        chkEDT812.Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                        chkEDT812.Checked = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                    Else
                        chkEDT812.Enabled = True
                        chkEDT812.Checked = True
                    End If

                Case "816"
                    If COMPANY_CODE = "NYA" Then
                        chkEDT816.Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                        chkEDT816.Checked = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                    Else
                        chkEDT816.Enabled = True
                        chkEDT816.Checked = True
                    End If

                Case "824"
                    If COMPANY_CODE = "NYA" Then
                        chkEDT824.Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                        chkEDT824.Checked = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                    Else
                        chkEDT824.Enabled = True
                        chkEDT824.Checked = True
                    End If

                Case "860"
                    If COMPANY_CODE = "NYA" Then
                        chkEDT860.Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                        chkEDT860.Checked = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                    Else
                        chkEDT860.Enabled = True
                        chkEDT860.Checked = True
                    End If

                Case "864"
                    If COMPANY_CODE = "NYA" Then
                        chkEDT864.Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                        chkEDT864.Checked = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                    Else
                        chkEDT864.Enabled = True
                        chkEDT864.Checked = True
                    End If

            End Select
        Next

        If COMPANY_CODE = "NYA" Then
            chkEDT947.Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED") OrElse ASCMAIN1.USER_SECURITY_CODEs.Contains("IC")
            chkEDT947.Checked = ASCMAIN1.USER_SECURITY_CODEs.Contains("ED") OrElse ASCMAIN1.USER_SECURITY_CODEs.Contains("IC")
        End If

    End Sub

    Public Overrides Sub Proceed_PreReq(eItemKey As String)
        MyBase.Proceed_PreReq(eItemKey)

        Select Case eItemKey

            Case "Proceed"
                If chkEDT816.Checked OrElse chkEDT824.Checked OrElse chkEDT860.Checked OrElse chkEDT864.Checked OrElse chkEDT947.Checked OrElse chkEDT812.Checked Then
                    ' nothing at this time
                Else
                    EMsg &= vbCr & "You must select at least one EDI table to extract."
                End If

        End Select
    End Sub

    Protected Overrides Sub Build_Workfile()
        Prepare_dst(True)
        RWU = "R"
    End Sub

    Public Overrides Sub Print_Report()

        If chkEDT816.Checked Then
            RPT_TITLE = "816 - Store Address Changes"
            'Generate_Report("EDR816R1", RPT_TITLE, String.Empty)
            Generate_Report("EDR816R2", RPT_TITLE, String.Empty)
        End If

        If chkEDT824.Checked Then
            RPT_TITLE = "824 - Advice"
            Generate_Report("EDR824R1", RPT_TITLE, String.Empty)
        End If

        If chkEDT860.Checked Then
            RPT_TITLE = "860 - PO Changes"
            Generate_Report("EDR860R1", RPT_TITLE, String.Empty)
        End If

        If chkEDT864.Checked Then
            RPT_TITLE = "864 - Messages"
            Generate_Report("EDR864R1", RPT_TITLE, String.Empty)
        End If

        If chkEDT947.Checked AndAlso COMPANY_CODE = "NYA" Then
            RPT_TITLE = "947 - Inventory Adjustment"
            Generate_Report("EDR947R1", RPT_TITLE, String.Empty)
        End If

        If chkEDT812.Checked AndAlso COMPANY_CODE = "NYA" Then
            RPT_TITLE = "812 - Credit/Debit Adjustment"
            Generate_Report("EDR812R1", RPT_TITLE, String.Empty)
        End If

    End Sub

    Overrides Function Prepare_dst( _
          ByVal perform_fill As Boolean, _
          ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then
            Clear_dst()
        End If

        Dim sql As String = String.Empty

        With dst

            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "ARTCUST2", "*")

            Create_TDA(.Tables.Add, "EDTTRPM1", "*")
            Create_TDA(.Tables.Add, "ICTSTYC1", "*")

            If chkEDT812.Checked Then
                Create_TDA(.Tables.Add, "EDT812T1", "SELECT * FROM EDT812T1")
                Create_TDA(.Tables.Add, "EDT812T2", "SELECT * FROM EDT812T2")
                Create_TDA(.Tables.Add, "EDT812T4", "SELECT * FROM EDT812T4")
                .Tables("EDT812T1").Columns.Add("EDI_CMT_REF", GetType(System.String))
            End If

            If chkEDT816.Checked Then
                Create_TDA(.Tables.Add, "EDT816T1", "SELECT * FROM EDT816T1")
                Create_TDA(.Tables.Add, "EDT816T2", "SELECT * FROM EDT816T2")
                Create_TDA(.Tables.Add, "EDT816T3", "SELECT * FROM EDT816T3")

                sql = "Select CUST_CODE, NULL MAINT_TYPE ,CUST_ADDR_CODE CUST_STORE_NO, CUST_NAME CUST_STORE_NAME,"
                sql &= " CUST_ADDR1 CUST_STORE_ADDR1, CUST_ADDR2 CUST_STORE_ADDR2, CUST_CITY CUST_STORE_CITY,"
                sql &= " CUST_STATE CUST_STORE_STATE, CUST_ZIP_CODE CUST_STORE_ZIP_CODE, CUST_DC_NO,"
                sql &= " CUST_NAME CUST_STORE_NAME_OLD, CUST_ADDR1 CUST_STORE_ADDR1_OLD, "
                sql &= " CUST_ADDR2 CUST_STORE_ADDR2_OLD, CUST_CITY CUST_STORE_CITY_OLD, "
                sql &= " CUST_STATE CUST_STORE_STATE_OLD, CUST_ZIP_CODE CUST_STORE_ZIP_CODE_OLD, "
                sql &= " CUST_DC_NO CUST_DC_NO_OLD "
                sql &= " FROM ARTCUST2 WHERE ROWNUM < 1"
                Create_TDA(.Tables.Add, "EDT816TR", sql, 0, False, "", 0)
                .Tables("EDT816TR").Columns("MAINT_TYPE").MaxLength = 10
            End If

            If chkEDT824.Checked Then
                Create_TDA(.Tables.Add, "EDT824T1", "*")
                Create_TDA(.Tables.Add, "EDT824T2", "*")
                Create_TDA(.Tables.Add, "EDT824T3", "*")
                Create_TDA(.Tables.Add, "EDT824T4", "*")
            End If

            If chkEDT860.Checked Then
                Create_TDA(.Tables.Add, "EDT860T1", "SELECT * FROM EDT860T1")
                Create_TDA(.Tables.Add, "EDT860T2", "SELECT * FROM EDT860T2")
                .Tables("EDT860T2").Columns.Add("PO_CHANGE_TYPE", GetType(System.String))
                Create_TDA(.Tables.Add, "EDT860T3", "SELECT * FROM EDT860T3")
                'EDT860T4 – header comments
                'EDT860TC – detail line comments
                Create_TDA(.Tables.Add, "EDT860T4", "SELECT * FROM EDT860T4")
                Create_TDA(.Tables.Add, "EDT860TC", "SELECT * FROM EDT860TC")

                tblEDT860T1 = ASCMAIN1.Temp_Table("Select EDI_DOC_SEQ_NO from EDT860T1 WHERE ROWNUM < 1")
                Create_TDA(.Tables.Add, tblEDT860T1, "SELECT * FROM " & tblEDT860T1)

                Create_TDA(.Tables.Add, "EDT850T1", "*")
                Create_TDA(.Tables.Add, "EDT850T2", "*")

            End If

            If chkEDT864.Checked Then
                Create_TDA(.Tables.Add, "EDT864T1", "SELECT * FROM EDT864T1")
                Create_TDA(.Tables.Add, "EDT864T2", "*")
                Create_TDA(.Tables.Add, "EDT864T3", "*")
                Create_TDA(.Tables.Add, "EDT864T5", "*")
            End If

            If chkEDT947.Checked Then
                Create_TDA(.Tables.Add, "EDT947T1", "*")
                Create_TDA(.Tables.Add, "EDT947T2", "*")
                .Tables("EDT947T2").Columns.Add("STYLE_DESC", GetType(System.String))
                .Tables("EDT947T2").Columns.Add("EDI_ADDR_CODE_QUAL", GetType(System.String))
                Create_TDA(.Tables.Add, "ICTWHSE1", "*")
                Create_TDA(.Tables.Add, "EDTXREF1", "*")
            End If

            Create_Lookup("ICTSTYC1", "*", "UPC_CODE = :PARM1", "V", False)
            Create_Lookup("ARTCUST1")
            Create_Lookup("ARTCUST2")

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        EnforceConstraints(False)

        Dim tableNames As New List(Of String)


        sql = "Select * from EDTTRPM1"
        Fill_Records("EDTTRPM1", "", True, sql)

        If chkEDT812.Checked Then
            Load812()
            tableNames.Add("EDT812T1")
        End If

        If chkEDT816.Checked Then
            Load816()
            tableNames.Add("EDT816T1")
        End If

        If chkEDT824.Checked Then
            Load824()
            tableNames.Add("EDT824T1")
        End If

        If chkEDT860.Checked Then
            Load860()
            tableNames.Add("EDT860T1")
        End If

        If chkEDT864.Checked Then
            Load864()
            tableNames.Add("EDT864T1")
        End If

        If chkEDT947.Checked AndAlso COMPANY_CODE = "NYA" Then
            Load947()
            tableNames.Add("EDT947T1")
        End If

        EnforceConstraints(True)

        RWU = "N"
        For Each table As String In tableNames
            If dst.Tables(table).Rows.Count > 0 Then
                RWU = "R"
                Exit For
            End If
        Next

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    Public Overrides Sub Update_Record()
        MyBase.Update_Record()
        Dim SQL As String = String.Empty

        If chkEDT812.Checked AndAlso EDI_DOC_SEQ_NO_812.Length > 0 Then
            SQL = "UPDATE EDT812T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_812 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If


        If chkEDT816.Checked AndAlso EDI_DOC_SEQ_NO_816.Length > 0 Then
            SQL = "UPDATE EDT816T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_816 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

        If chkEDT824.Checked AndAlso EDI_DOC_SEQ_NO_824.Length > 0 Then
            SQL = "UPDATE EDT824T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

        If chkEDT860.Checked AndAlso EDI_DOC_SEQ_NO_860.Length > 0 Then
            SQL = "UPDATE EDT860T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

        If chkEDT864.Checked AndAlso EDI_DOC_SEQ_NO_864.Length > 0 Then
            SQL = "UPDATE EDT864T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_864 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

        If chkEDT947.Checked AndAlso EDI_DOC_SEQ_NO_947.Length > 0 AndAlso COMPANY_CODE = "NYA" Then
            SQL = "UPDATE EDT947T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_947 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

    End Sub

    Private Sub Load812()
        Dim sql As String = String.Empty
        EDI_DOC_SEQ_NO_812 = String.Empty
        Dim rowARTCUST2 As DataRow = Nothing

        ASCMAIN1.Progress("Processing 812's", String.Empty)

        Try
            ASCDATA1.ExecuteSQL("Update EDT812T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT812T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT812T1.EDI_OUR_ID and EDI_TP_ID = EDT812T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT812T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = EDT812T1.EDI_TP_QUAL and EDI_TP_ID = EDT812T1.EDI_TP_ID and EDI_DOC_NO = 812)" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & COMPANY_CODE & "'"
            ASCDATA1.ExecuteSQL()

            For Each TABLE_NAME As String In New String() {"EDT812T1", "EDT812T2", "EDT812T4"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "Select * from EDT812T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & COMPANY_CODE & "'"
            Fill_Records("EDT812T1", String.Empty, True, sql)

            For Each rowEDT812T1 As DataRow In dst.Tables("EDT812T1").Rows
                EDI_DOC_SEQ_NO_812 &= ", '" & rowEDT812T1.Item("EDI_DOC_SEQ_NO") & "'"

                Dim EDI_TP_ID As String = rowEDT812T1.Item("EDI_TP_ID") & String.Empty
                If dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '812' AND EDI_TP_ID = '" & EDI_TP_ID & "'").Length > 0 Then
                    rowEDT812T1.Item("CUST_CODE") = dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '812' AND EDI_TP_ID = '" & EDI_TP_ID & "'")(0).Item("CUST_CODE")
                End If
            Next

            If EDI_DOC_SEQ_NO_812.Length > 0 Then
                EDI_DOC_SEQ_NO_812 = EDI_DOC_SEQ_NO_812.Substring(2).Trim
                sql = "Select * from EDT812T2 where EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_812 & ")"
                Fill_Records("EDT812T2", String.Empty, True, sql)

                sql = "Select * from EDT812T4 where EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_812 & ")"
                Fill_Records("EDT812T4", String.Empty, True, sql)
            End If

            For Each rowEDT812T1 As DataRow In dst.Tables("EDT812T1").Select("")
                Dim EDI_DOC_SEQ_NO As String = rowEDT812T1.Item("EDI_DOC_SEQ_NO") & String.Empty
                Dim EDI_CMT_REF As String = String.Empty

                For Each rowEDT812T4 As DataRow In dst.Tables("EDT812T4").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'", "EDI_CMT_SEQ")
                    EDI_CMT_REF &= rowEDT812T4.Item("EDI_CMT_REF") & String.Empty
                Next

                EDI_CMT_REF = EDI_CMT_REF.Replace("  ", " ")
                rowEDT812T1.Item("EDI_CMT_REF") = StrConv(EDI_CMT_REF, VbStrConv.ProperCase)
            Next

        Catch ex As Exception
            MessageBox.Show("Error processing 812's: " & ex.Message, "EDI 812", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try
    End Sub

    Private Sub Load816()
        Dim sql As String = String.Empty
        EDI_DOC_SEQ_NO_816 = String.Empty
        Dim rowARTCUST2 As DataRow = Nothing

        ASCMAIN1.Progress("Processing 816's", String.Empty)

        Try

            ASCDATA1.ExecuteSQL("Update EDT816T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT816T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT816T1.EDI_OUR_ID and EDI_TP_ID = EDT816T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT816T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = EDT816T1.EDI_TP_QUAL and EDI_TP_ID = EDT816T1.EDI_TP_ID and EDI_DOC_NO = 816)" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & COMPANY_CODE & "'"
            ASCDATA1.ExecuteSQL()

            For Each TABLE_NAME As String In New String() {"EDT816T1", "EDT816T2", "EDT816T3", "EDT816TR"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "Select * from EDT816T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & COMPANY_CODE & "'"
            Fill_Records("EDT816T1", String.Empty, True, sql)

            For Each rowEDT816T1 As DataRow In dst.Tables("EDT816T1").Rows
                EDI_DOC_SEQ_NO_816 &= ", '" & rowEDT816T1.Item("EDI_DOC_SEQ_NO") & "'"

                Dim EDI_TP_ID As String = rowEDT816T1.Item("EDI_TP_ID") & String.Empty
                If dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'").Length > 0 Then
                    rowEDT816T1.Item("CUST_CODE") = dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'")(0).Item("CUST_CODE")
                End If
            Next

            If EDI_DOC_SEQ_NO_816.Length > 0 Then
                EDI_DOC_SEQ_NO_816 = EDI_DOC_SEQ_NO_816.Substring(2).Trim
                sql = "Select * from EDT816T2 where EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_816 & ")"
                Fill_Records("EDT816T2", String.Empty, True, sql)

                sql = "Select * from EDT816T3 where EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_816 & ")"
                Fill_Records("EDT816T3", String.Empty, True, sql)
            End If

            For Each rowEDT816T1 As DataRow In dst.Tables("EDT816T1").Rows
                Dim rowEDTTRPM1 As DataRow = Nothing
                Dim EDI_TP_ID As String = (rowEDT816T1.Item("EDI_TP_ID") & String.Empty).ToString.Trim

                If dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'").Length > 0 Then
                    rowEDTTRPM1 = dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'")(0)
                End If

                If rowEDTTRPM1 IsNot Nothing Then
                    Dim CUST_CODE As String = rowEDTTRPM1.Item("CUST_CODE") & String.Empty
                    For Each rowEDT816T2 As DataRow In dst.Tables("EDT816T2").Select("EDI_DOC_SEQ_NO = '" & rowEDT816T1.Item("EDI_DOC_SEQ_NO") & "'")
                        Dim CUST_STORE_NO As String = rowEDT816T2.Item("CUST_ADDR_CODE") & String.Empty
                        Dim rowEDT816TR As DataRow = dst.Tables("EDT816TR").NewRow
                        rowEDT816TR.Item("CUST_CODE") = rowEDTTRPM1.Item("CUST_CODE") & ""
                        rowEDT816TR.Item("MAINT_TYPE") = rowEDT816T2.Item("MAINT_TYPE") & ""
                        rowEDT816TR.Item("CUST_STORE_NO") = rowEDT816T2.Item("CUST_ADDR_CODE") & ""
                        rowEDT816TR.Item("CUST_STORE_NAME") = rowEDT816T2.Item("CUST_NAME") & ""
                        rowEDT816TR.Item("CUST_STORE_ADDR1") = rowEDT816T2.Item("CUST_ADDR1") & ""
                        rowEDT816TR.Item("CUST_STORE_ADDR2") = rowEDT816T2.Item("CUST_ADDR2") & ""
                        rowEDT816TR.Item("CUST_STORE_ZIP_CODE") = rowEDT816T2.Item("CUST_ZIP_CODE") & ""
                        rowEDT816TR.Item("CUST_STORE_CITY") = rowEDT816T2.Item("CUST_CITY") & ""
                        rowEDT816TR.Item("CUST_STORE_STATE") = rowEDT816T2.Item("CUST_STATE") & ""
                        rowEDT816TR.Item("CUST_DC_NO") = rowEDT816T2.Item("CUST_DC_CODE") & ""

                        rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                        If rowARTCUST2 IsNot Nothing Then
                            rowEDT816TR.Item("CUST_STORE_NAME_OLD") = rowARTCUST2.Item("CUST_STORE_NAME") & ""
                            rowEDT816TR.Item("CUST_STORE_ADDR1_OLD") = rowARTCUST2.Item("CUST_STORE_ADDR1") & ""
                            rowEDT816TR.Item("CUST_STORE_ADDR2_OLD") = rowARTCUST2.Item("CUST_STORE_ADDR2") & ""
                            rowEDT816TR.Item("CUST_STORE_ZIP_CODE_OLD") = rowARTCUST2.Item("CUST_STORE_ZIP_CODE") & ""
                            rowEDT816TR.Item("CUST_STORE_CITY_OLD") = rowARTCUST2.Item("CUST_STORE_CITY") & ""
                            rowEDT816TR.Item("CUST_STORE_STATE_OLD") = rowARTCUST2.Item("CUST_STORE_STATE") & ""
                            rowEDT816TR.Item("CUST_DC_NO_OLD") = rowARTCUST2.Item("CUST_DC_NO") & ""
                        End If

                        dst.Tables("EDT816TR").Rows.Add(rowEDT816TR)
                    Next
                End If

            Next
        Catch ex As Exception
            MessageBox.Show("Error processing 816's: " & ex.Message, "EDI 816", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try
    End Sub

    Private Sub Load824()

        EDI_DOC_SEQ_NO_824 = String.Empty

        ASCMAIN1.Progress("Processing 824's", String.Empty)

        Try
            For Each TABLE_NAME As String In New String() {"EDT824T1", "EDT824T2", "EDT824T3", "EDT824T4"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "SELECT EDT824T1.*, '' CUST_CODE FROM EDT824T1 where NVL(EDI_PROCESS_IND, '0') = '0' "
            Fill_Records("EDT824T1", String.Empty, True, sql)

            For Each rowEDT824T1 As DataRow In dst.Tables("EDT824T1").Select()
                Dim EDI_TP_ID As String = rowEDT824T1.Item("EDI_TP_ID") & String.Empty
                Dim EDI_TP_QUAL As String = rowEDT824T1.Item("EDI_TP_QUAL") & String.Empty

                sql = "EDI_DOC_NO = '824' AND EDI_TP_ID = '" & EDI_TP_ID & "' and EDI_TP_QUAL = '" & EDI_TP_QUAL & "'"
                If dst.Tables("EDTTRPM1").Select(sql).Length > 0 Then
                    rowEDT824T1.Item("CUST_CODE") = dst.Tables("EDTTRPM1").Select(sql)(0).Item("CUST_CODE")
                End If

                EDI_DOC_SEQ_NO_824 = EDI_DOC_SEQ_NO_824 & ", '" & rowEDT824T1.Item("EDI_DOC_SEQ_NO") & "'"
            Next

            If EDI_DOC_SEQ_NO_824.Length > 0 Then
                EDI_DOC_SEQ_NO_824 = EDI_DOC_SEQ_NO_824.Substring(2).Trim

                sql = "SELECT * from EDT824T2 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
                Fill_Records("EDT824T2", String.Empty, True, sql)

                sql = "SELECT * from EDT824T3 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
                Fill_Records("EDT824T3", String.Empty, True, sql)

                sql = "SELECT * from EDT824T4 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
                Fill_Records("EDT824T4", String.Empty, True, sql)
            End If

        Catch ex As Exception
            MessageBox.Show("Error processing 824's: " & ex.Message, "EDI 864", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

    Private Sub Load860()

        Dim sql As String = String.Empty
        Dim rowICTSTYC1 As DataRow = Nothing
        EDI_DOC_SEQ_NO_860 = String.Empty

        For Each TABLE_NAME As String In New String() {"EDT860T1", "EDT860T2", "EDT860T3", "EDT850T1", "EDT850T2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        ASCMAIN1.Progress("Processing 860's", String.Empty)
        Try

            ASCDATA1.ExecuteSQL("Update EDT860T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT860T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT860T1.EDI_OUR_ID and EDI_TP_ID = EDT860T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT860T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = EDT860T1.EDI_TP_QUAL and EDI_TP_ID = EDT860T1.EDI_TP_ID and EDI_DOC_NO = 860)" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & COMPANY_CODE & "'"
            ASCDATA1.ExecuteSQL()

            sql = "Select * from EDT860T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & COMPANY_CODE & "'"
            Fill_Records("EDT860T1", String.Empty, True, sql)

            dst.Tables(tblEDT860T1).Rows.Clear()
            For Each rowEDT860T1 As DataRow In dst.Tables("EDT860T1").Rows
                'EDI_DOC_SEQ_NO_860 &= ", '" & rowEDT860T1.Item("EDI_DOC_SEQ_NO") & "'"
                dst.Tables(tblEDT860T1).Rows.Add(New Object() {rowEDT860T1.Item("EDI_DOC_SEQ_NO")})
            Next
            ASCDATA1.ExecuteSQL("DELETE FROM " & tblEDT860T1)
            Update_Record_TDA(tblEDT860T1)

            ' SPECIAL PROCESSING FOR WALMART
            ASCMAIN1.sql = "SELECT * FROM EDT850T1 WHERE (EDI_TP_ID, EDI_PO_NO) IN (SELECT EDI_TP_ID, EDI_PO_NO FROM EDT860T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblEDT860T1 & "))"
            Fill_Records("EDT850T1", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM EDT850T2 WHERE EDI_DOC_SEQ_NO IN (" _
                & " SELECT EDI_DOC_SEQ_NO FROM EDT850T1 WHERE (EDI_TP_ID, EDI_PO_NO) IN (SELECT EDI_TP_ID, EDI_PO_NO FROM EDT860T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblEDT860T1 & "))" _
                & " )"
            Fill_Records("EDT850T2", String.Empty, True, ASCMAIN1.sql)

            If dst.Tables(tblEDT860T1).Rows.Count > 0 Then
                'EDI_DOC_SEQ_NO_860 = EDI_DOC_SEQ_NO_860.Substring(2).Trim
                EDI_DOC_SEQ_NO_860 = "Select EDI_DOC_SEQ_NO from " & tblEDT860T1
                sql = "Select * from EDT860T2 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
                Fill_Records("EDT860T2", String.Empty, True, sql)

                sql = "Select * from EDT860T3 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
                Fill_Records("EDT860T3", String.Empty, True, sql)

                sql = "Select * from EDT860T4 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
                Fill_Records("EDT860T4", String.Empty, True, sql)

                sql = "Select * from EDT860TC WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
                Fill_Records("EDT860TC", String.Empty, True, sql)
            End If

            For Each rowEDT860T2 As DataRow In dst.Tables("EDT860T2").Select("", "EDI_DOC_SEQ_NO, EDI_ORIG_DTL_SEQ")

                Dim EDI_DOC_SEQ_NO As String = rowEDT860T2.Item("EDI_DOC_SEQ_NO") & String.Empty

                Dim EDI_UPC As String = rowEDT860T2.Item("EDI_UPC") & String.Empty
                If EDI_UPC.Length = 0 Then
                    EDI_UPC = rowEDT860T2.Item("EDI_EAN") & String.Empty
                End If

                rowICTSTYC1 = Nothing
                If EDI_UPC.Length > 0 Then
                    rowICTSTYC1 = LookUp("ICTSTYC1", New String() {EDI_UPC})
                End If

                ' SPECIAL PROCESSING FOR REGENCY / WALMART 
                If rowICTSTYC1 Is Nothing AndAlso ASCMAIN1.CLIENT = "RGI" Then
                    Dim rowEDT860T1 As DataRow = dst.Tables("EDT860T1").Rows.Find(EDI_DOC_SEQ_NO)
                    If rowEDT860T1 IsNot Nothing AndAlso rowEDT860T1.Item("CUST_CODE") & String.Empty = RegencyWalmartCustCode Then
                        Dim EDI_TP_QUAL As String = rowEDT860T1.Item("EDI_TP_QUAL") & String.Empty
                        Dim EDI_TP_ID As String = rowEDT860T1.Item("EDI_TP_ID") & String.Empty
                        Dim EDI_PO_NO As String = rowEDT860T1.Item("EDI_PO_NO") & String.Empty
                        Dim CUST_CODE As String = rowEDT860T1.Item("CUST_CODE") & String.Empty

                        Dim EDI_PO_LNO As Int16 = Val(rowEDT860T2.Item("EDI_PO_LNO") & String.Empty)
                        ASCMAIN1.sql = "EDI_TP_QUAL = '" & EDI_TP_QUAL & "' AND EDI_TP_ID = '" & EDI_TP_ID & "' AND EDI_PO_NO = '" & EDI_PO_NO & "'"
                        If dst.Tables("EDT850T1").Select(ASCMAIN1.sql).Length > 0 Then
                            Dim EDI_DOC_SEQ_NO_T1 As String = dst.Tables("EDT850T1").Select(ASCMAIN1.sql)(0).Item("EDI_DOC_SEQ_NO") & String.Empty

                            ASCMAIN1.sql = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO_T1 & "' AND EDI_PO_LNO = " & EDI_PO_LNO
                            If dst.Tables("EDT850T2").Select(ASCMAIN1.sql).Length > 0 Then
                                Dim rowEDT850T2 As DataRow = dst.Tables("EDT850T2").Select(ASCMAIN1.sql)(0)

                                EDI_UPC = rowEDT850T2.Item("EDI_UPC") & String.Empty
                                rowEDT860T2.Item("EDI_UPC") = EDI_UPC
                                rowEDT860T2.Item("EDI_SKU") = rowEDT850T2.Item("EDI_SKU")
                                rowEDT860T2.Item("EDI_ITEM") = rowEDT850T2.Item("EDI_ITEM")

                                If EDI_UPC.Length > 0 Then
                                    rowICTSTYC1 = LookUp("ICTSTYC1", New String() {EDI_UPC})
                                End If

                            End If
                        End If
                    End If
                End If

                Select Case rowEDT860T2.Item("EDI_CHANGE_TYPE") & String.Empty
                    Case "QI", "QD"
                        If rowEDT860T2.Item("EDI_CHANGE_TYPE") = "QD" Then
                            rowEDT860T2.Item("EDI_QTY_CHANGE") = Val(rowEDT860T2.Item("EDI_QTY_CHANGE") & String.Empty) * -1
                        End If
                    Case Else
                        rowEDT860T2.Item("EDI_QTY_CHANGE") = rowEDT860T2.Item("EDI_QTY_OPEN")
                End Select

                If rowICTSTYC1 IsNot Nothing Then
                    rowEDT860T2.Item("EDI_ITEM") = rowICTSTYC1.Item("STYLE_CODE")
                End If
            Next

            ' Create dummy records for T2 where it does not exist for a T1
            For Each rowEDT860T1 As DataRow In dst.Tables("EDT860T1").Select
                Dim EDI_DOC_SEQ_NO As String = rowEDT860T1.Item("EDI_DOC_SEQ_NO") & String.Empty
                If dst.Tables("EDT860T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length = 0 Then
                    Dim rowEDT860T2 As DataRow = dst.Tables("EDT860T2").NewRow
                    rowEDT860T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    rowEDT860T2.Item("EDI_ORIG_DTL_SEQ") = -9
                    dst.Tables("EDT860T2").Rows.Add(rowEDT860T2)
                End If

                If dst.Tables("EDT860T3").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length = 0 Then
                    Dim rowEDT860T3 As DataRow = dst.Tables("EDT860T3").NewRow
                    rowEDT860T3.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    rowEDT860T3.Item("EDI_ORIG_DTL_SEQ") = -9
                    rowEDT860T3.Item("EDI_SDQ_SEQ") = 1
                    dst.Tables("EDT860T3").Rows.Add(rowEDT860T3)
                End If

            Next


        Catch ex As Exception
            MessageBox.Show("Error processing 860's: " & ex.Message, "EDI 860", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try
    End Sub

    Private Sub Load864()
        EDI_DOC_SEQ_NO_864 = String.Empty

        ASCMAIN1.Progress("Processing 864's", String.Empty)

        Try
            ASCDATA1.ExecuteSQL("Update EDT864T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")
 
            ASCMAIN1.sql = "Update EDT864T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT864T1.EDI_OUR_ID and EDI_TP_ID = EDT864T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT864T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = EDT864T1.EDI_TP_QUAL and EDI_TP_ID = EDT864T1.EDI_TP_ID and EDI_DOC_NO = 864)" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & COMPANY_CODE & "'"
            ASCDATA1.ExecuteSQL()

            For Each TABLE_NAME As String In New String() {"EDT864T1", "EDT864T2", "EDT864T3", "EDT864T5"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "SELECT EDT864T1.* FROM EDT864T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & COMPANY_CODE & "'"
            Fill_Records("EDT864T1", String.Empty, True, sql)

            For Each rowEDT864T1 As DataRow In dst.Tables("EDT864T1").Select()
                EDI_DOC_SEQ_NO_864 = EDI_DOC_SEQ_NO_864 & ", '" & rowEDT864T1.Item("EDI_DOC_SEQ_NO") & "'"
            Next

            If EDI_DOC_SEQ_NO_864.Length > 0 Then
                EDI_DOC_SEQ_NO_864 = EDI_DOC_SEQ_NO_864.Substring(2).Trim

                sql = "SELECT * from EDT864T2 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_864 & ")"
                Fill_Records("EDT864T2", String.Empty, True, sql)

                sql = "SELECT * from EDT864T3 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_864 & ")"
                Fill_Records("EDT864T3", String.Empty, True, sql)

                sql = "SELECT * from EDT864T5 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_864 & ")"
                Fill_Records("EDT864T5", String.Empty, True, sql)
            End If

        Catch ex As Exception
            MessageBox.Show("Error processing 864's: " & ex.Message, "EDI 864", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try
    End Sub

    Private Sub Load947()

        EDI_DOC_SEQ_NO_947 = String.Empty

        If COMPANY_CODE <> "NYA" Then
            Exit Sub
        End If

        ASCMAIN1.Progress("Processing 947's", String.Empty)

        Try

            ASCDATA1.ExecuteSQL("Update EDT947T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT947T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT947T1.EDI_OUR_ID and EDI_TP_ID = EDT947T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT947T1 Set WHSE_CODE = (Select Distinct WHSE_CODE from ICTWHSE1" _
               & " where WHSE_EDI_QUAL = EDT947T1.EDI_TP_QUAL and WHSE_EDI_ID = EDT947T1.EDI_TP_ID and ICTWHSE1.LP_WHSE_ID = EDT947T1.EDI_ADDR_CODE)" _
               & " where EDI_PROCESS_IND = '0' and WHSE_CODE IS NULL and COMPANY_CODE = '" & COMPANY_CODE & "'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT947T1 Set WHSE_CODE = '95'" _
                   & " where EDI_PROCESS_IND = '0' and WHSE_CODE IS NULL and COMPANY_CODE = '" & COMPANY_CODE & "'" _
                   & "   and EDI_ADDR_CODE in ('NYDG','NYWM','NYWB')"
            ASCDATA1.ExecuteSQL()

            For Each TABLE_NAME As String In New String() {"EDT947T1", "EDT947T2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "SELECT EDT947T1.* FROM EDT947T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & COMPANY_CODE & "'"
            Fill_Records("EDT947T1", String.Empty, True, sql)

            For Each rowEDT947T1 As DataRow In dst.Tables("EDT947T1").Select()
                EDI_DOC_SEQ_NO_947 = EDI_DOC_SEQ_NO_947 & ", '" & rowEDT947T1.Item("EDI_DOC_SEQ_NO") & "'"
                If rowEDT947T1.Item("WHSE_CODE") & String.Empty = String.Empty Then
                    rowEDT947T1.Item("WHSE_CODE") = "**"
                End If
            Next

            If EDI_DOC_SEQ_NO_947.Length > 0 Then
                EDI_DOC_SEQ_NO_947 = EDI_DOC_SEQ_NO_947.Substring(2).Trim

                sql = "SELECT EDT947T2.*, ICTSTYL1.STYLE_DESC from EDT947T2, ICTSTYL1 WHERE EDT947T2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+) AND EDT947T2.EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_947 & ")"
                Fill_Records("EDT947T2", String.Empty, True, sql)

                Fill_Records("ICTWHSE1", String.Empty, True, "SELECT * FROM ICTWHSE1")
                Fill_Records("EDTXREF1", String.Empty, True, "SELECT * FROM EDTXREF1")

                dst.Tables("ICTWHSE1").Rows.Add(New Object() {"**", "A", "Unknown"})
            End If

        Catch ex As Exception
            MessageBox.Show("Error processing 947's: " & ex.Message, "EDI 947", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try
    End Sub

End Class