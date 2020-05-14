Public Class ARRCBPR1

    Dim ARTCBPR1 As String
    Dim tblH As DataTable

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -1)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        ASCMAIN1.Progress("Work Tables")
        Create_WorkTable()

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        MyBase.Get_SQL("*", ARTCBPR1)

        If Not Absx1.chkFor("CHKDTL").Checked Then
            COLUMN_NAMEs_appended = ""
        End If

        sql_Data = "" _
            & ", SUM (GL_DIST_AMT) TOTAL" & vbCrLf

        sql_Cols = "" _
            & ",TOTAL"

        sql_filter = ""

        sql = "Select " & sql_SELECT_cols & COLUMN_NAMEs_appended & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from " & ARTCBPR1 & " ARTCBPR1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols & COLUMN_NAMEs_appended

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()

        ASCMAIN1.sql = "Select * from " & ARTCBPR1 & IIf(Absx1.chkFor("CHKDTL").Checked, "", " where ROWNUM < 1")
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCBPR1", 0))

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME || ' ' || CUST_CITY || ' ' || CUST_STATE || ' ' || CUST_CONTACT || Decode (CUST_COUNTRY, Null, ' ',' (' || CUST_COUNTRY || ')') CUST_NAME from ARTCUST1 where CUST_CODE in (Select CUST_CODE from " & ARTCBPR1 & ")"
        Dim tblARTCUST1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1", 1)
        For Each row As DataRow In dst.Tables("ASTGROUP").Select("")
            Dim CUST_CODE As String = row.Item("GROUP_CODE")
            Dim rowARTCUST1 As DataRow = tblARTCUST1.Rows.Find(CUST_CODE)
            If rowARTCUST1 IsNot Nothing Then
                row.Item("GROUP_DESC") = rowARTCUST1.Item("CUST_NAME")
            End If
        Next
    End Sub
    Public Overrides Sub Print_Report()
        If RYP0 = RYP1 Then
            SUBT = RYPLEGEND0
        Else
            SUBT = RYPLEGEND0 & " thru " & RYPLEGEND1
        End If

        Dim RPT_TITLE_modified As String = ""
        If Absx1.optFor("OPTDC").Value = "C" Then
            RPT_TITLE_modified = "Chargebacks Summary"
        ElseIf Absx1.optFor("OPTDC").Value = "D" Then
            RPT_TITLE_modified = "Deductions Summary"
        End If

        If Absx1.chkFor("CHKOPENONLY").Checked Then
            SUBT &= " - Open Chargebacks Only"
        End If

        CR_params.Add("SUBT", SUBT)
        CR_params.Add("CHKDTL", IIf(Absx1.chkFor("CHKDTL").Checked, "1", "0"))

        If Not Absx1.chkFor("CHKDTL").Checked Then
            If tblH Is Nothing Then
                tblH = tblASTDSQLH.Copy
            End If
            tblASTDSQLH.Rows.Clear()
        End If

        Generate_Report(RPT, RPT_TITLE_modified, SUBT)

        If Not Absx1.chkFor("CHKDTL").Checked Then
            tblASTDSQLH.Merge(tblH)
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            End If
        End If
    End Sub
    Sub Create_WorkTable()

        ASCMAIN1.sql = "Select " & vbCrLf _
            & " X5.PYMT_BATCH_NO, X5.PYMT_BATCH_LNO, 'D' PYMT_BATCH_LTYP, X5.PYMT_BATCH_DLNO PYMT_BATCH_XLNO, " & vbCrLf _
            & " X5.REASON_CODE, X1.PYMT_BATCH_DATE, " & vbCrLf _
            & " X2.CUST_CODE, X2.CUST_PYMT_REF_NO, X2.CUST_PYMT_REF_DATE, X2.CUST_PYMT_AMT, " & vbCrLf _
            & " X5.GL_DIST_AMT, X5.CHARGEBACK_NO, X5.CUST_REFERENCE" & vbCrLf _
            & " from ARTPYMT1 X1, ARTPYMT2 X2, ARTPYMT5 X5, ARTCUST1" _
            & IIf(Absx1.chkFor("CHKOPENONLY").Checked, ",ARTOPEN1", "") & vbCrLf _
            & " where X1.PYMT_BATCH_NO = X5.PYMT_BATCH_NO" & vbCrLf _
            & "   and X2.PYMT_BATCH_NO = X5.PYMT_BATCH_NO" & vbCrLf _
            & "   and X2.PYMT_BATCH_LNO = X5.PYMT_BATCH_LNO" & vbCrLf _
            & "   and X1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
            & "   and X1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = X2.CUST_CODE" & vbCrLf

        If Absx1.optFor("OPTDC").Value = "C" Then
            ASCMAIN1.sql &= "   and NVL(X5.CHARGEBACK_IND,'0') = '1'" & vbCrLf
            If Not Absx1.chkFor("CHKINCOA").Checked Then
                ASCMAIN1.sql &= "   and X5.INV_TYPE_CB <> 'O'" & vbCrLf
            End If
            If Absx1.chkFor("CHKOPENONLY").Checked Then
                ASCMAIN1.sql &= "   and ARTOPEN1.CUST_CODE = X2.CUST_CODE and ARTOPEN1.INV_TYPE = X5.INV_TYPE_CB and ARTOPEN1.INV_NUM = X5.CHARGEBACK_NO" & vbCrLf
                ASCMAIN1.sql &= "   and ARTOPEN1.INV_BALANCE <> 0" & vbCrLf
            End If

        Else
            ASCMAIN1.sql &= "   and NVL(X5.CHARGEBACK_IND,'0') <> '1'" & vbCrLf
        End If

        ASCMAIN1.sql &= SQL_in("CUST_CODE", "X2.CUST_CODE") & vbCrLf
        ASCMAIN1.sql &= SQL_in("REASON_CODE", "X5.REASON_CODE") & vbCrLf


        ARTCBPR1 = ASCMAIN1.Temp_Table


        If Absx1.optFor("OPTDC").Value = "D" Then
            Dim DEDz As String
            For DED As Integer = 1 To 2
                If DED = 1 Then
                    DEDz = "X5.INV_DISC_TAKEN"
                Else
                    DEDz = "X5.INV_WRITE_OFF"
                End If

                Dim REASON_CODE As String = ""
                Get_PARM("ARTPARM1")
                If DED = 1 Then REASON_CODE = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_DISC") & ""
                If DED = 2 Then REASON_CODE = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_WOFF") & ""

                If REASON_CODE <> "" Then
                    ASCMAIN1.sql = "Select " & vbCrLf _
                        & "X5.PYMT_BATCH_NO, X5.PYMT_BATCH_LNO, '" & CStr(DED) & "' PYMT_BATCH_LTYP, X5.PYMT_BATCH_ILNO PYMT_BATCH_XLNO, " & vbCrLf _
                        & " '" & REASON_CODE & "' REASON_CODE, X1.PYMT_BATCH_DATE, " & vbCrLf _
                        & " X2.CUST_CODE, X2.CUST_PYMT_REF_NO, X2.CUST_PYMT_REF_DATE, X2.CUST_PYMT_AMT, " & vbCrLf _
                        & " " & DEDz & " GL_DIST_AMT, NULL CHARGEBACK_NO, NULL CUST_REFERENCE" & vbCrLf _
                        & " from ARTPYMT1 X1, ARTPYMT2 X2, ARTPYMT3 X5, ARTCUST1" & vbCrLf _
                        & " where X1.PYMT_BATCH_NO = X5.PYMT_BATCH_NO" & vbCrLf _
                        & "   and X2.PYMT_BATCH_NO = X5.PYMT_BATCH_NO" & vbCrLf _
                        & "   and X2.PYMT_BATCH_LNO = X5.PYMT_BATCH_LNO" & vbCrLf _
                        & "   and " & DEDz & " <> 0 " & vbCrLf _
                        & "   and X1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
                        & "   and X1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf _
                        & "   and ARTCUST1.CUST_CODE = X2.CUST_CODE" & vbCrLf

                    ASCMAIN1.sql &= SQL_in("CUST_CODE", "X2.CUST_CODE") & vbCrLf
                    ASCMAIN1.sql &= Replace(SQL_in("REASON_CODE", "X5.REASON_CODE") & vbCrLf, "X5.REASON_CODE", "'" & REASON_CODE & "'")

                    ASCMAIN1.sql = "Insert into " & ARTCBPR1 & " " & ASCMAIN1.sql
                    ASCDATA1.ExecuteSQL()

                End If
            Next DED


            ASCMAIN1.sql = "Select " & vbCrLf _
                & "X5.PYMT_BATCH_NO, X5.PYMT_BATCH_LNO, 'G' PYMT_BATCH_LTYP, X5.PYMT_BATCH_GLNO PYMT_BATCH_XLNO, " & vbCrLf _
                & " 'GL' || X5.ACCT_CODE REASON_CODE, X1.PYMT_BATCH_DATE, " & vbCrLf _
                & " X2.CUST_CODE, X2.CUST_PYMT_REF_NO, X2.CUST_PYMT_REF_DATE, X2.CUST_PYMT_AMT, " & vbCrLf _
                & " X5.GL_DIST_AMT, NULL CHARGEBACK_NO, X5.GL_DIST_REF CUST_REFERENCE" & vbCrLf _
                & " from ARTPYMT1 X1, ARTPYMT2 X2, ARTPYMT4 X5, ARTCUST1" & vbCrLf _
                & " where X1.PYMT_BATCH_NO = X5.PYMT_BATCH_NO" & vbCrLf _
                & "   and X2.PYMT_BATCH_NO = X5.PYMT_BATCH_NO" & vbCrLf _
                & "   and X2.PYMT_BATCH_LNO = X5.PYMT_BATCH_LNO" & vbCrLf _
                & "   and X5.GL_DIST_AMT <> 0 " & vbCrLf _
                & "   and X1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
                & "   and X1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = X2.CUST_CODE" & vbCrLf

            ASCMAIN1.sql &= SQL_in("CUST_CODE", "X2.CUST_CODE") & vbCrLf
            ASCMAIN1.sql &= Replace(SQL_in("REASON_CODE", "X5.REASON_CODE") & vbCrLf, "X5.REASON_CODE", "'GL'")

            ASCMAIN1.sql = "Insert into " & ARTCBPR1 & " " & ASCMAIN1.sql
            'ASCDATA1.ExecuteSQL()

        End If

    End Sub

    Private Sub optDC_ValueChanged(sender As Object, e As EventArgs) Handles optDC.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Absx1.chkFor("CHKOPENONLY").Visible = optDC.Value = "C"
        Absx1.chkFor("CHKINCOA").Visible = optDC.Value = "C"
    End Sub
End Class