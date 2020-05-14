Public Class ARRSTMTR

    Dim ARTSTMTR As String
    Dim PRD_END_DATE As Date

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        ' NEED TO MOVE ROUTINE Get_Aging_Data_RGI IN ARCMAIN2 to ARCMAIN1 - Make regency Specific
        'Dim DRCCYP As String = "201403"
        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
        'Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", DRCCYP)
        PRD_END_DATE = rowGLTPARM2.Item("PRD_END_DATE")

        'TAC.ARCMAIN1.Get_Aging_Data(ROWs("ARTPARM1"), Now.Date)
        TAC.ARCMAIN2.Get_Aging_Data_RGI( _
        ROWs("ARTPARM1"), _
        PRD_END_DATE, True)

        ASCMAIN1.sql = "Select ARTOPEN1.* " & TAC.ARCMAIN2.DAYS_AND_BUCKETS _
        & ", DECODE (ARTOPEN1.INV_TYPE,'B',ARTOPEN1.INV_BALANCE,0) CHARGEBACKS " & vbCrLf _
        & ", CASE WHEN ARTOPEN1.INV_TYPE = 'C' OR ARTOPEN1.INV_TYPE = 'O' THEN ARTOPEN1.INV_BALANCE ELSE 0 END CREDITS" & vbCrLf _
        & " from ARTOPEN1, ARTCUST1 ARTCUSTX, TATTERM1 " & vbCrLf _
        & " where ARTOPEN1.INV_BALANCE <> 0" & vbCrLf _
        & " and ARTCUSTX.CUST_CODE = ARTOPEN1.CUST_CODE" & vbCrLf _
        & " and TATTERM1.TERM_CODE = ARTOPEN1.TERM_CODE"

        ASCMAIN1.sql &= SQL_in("CUST_CODE", "ARTCUSTX.CUST_CODE")
        ASCMAIN1.sql &= SQL_in("SREP_CODE", "ARTCUSTX.SREP_CODE")

        ARTSTMTR = ASCMAIN1.Temp_Table
        Create_TDA(dst.Tables.Add, "ARTSTMTR", "**", 0, False, "", 3)
        dst.Tables("ARTSTMTR").Columns.Add("AR_PARM_KEY")
        dst.Tables("ARTSTMTR").Columns("AR_PARM_KEY").DefaultValue = "Z"
        Fill_Records("ARTSTMTR")


        ASCMAIN1.sql = "Select ARTCUST1.* from ARTCUST1 where CUST_CODE in (Select CUST_CODE from " & ARTSTMTR & ")"
        Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)
        Fill_Records("ARTCUST1")

        ASCMAIN1.sql = "Select CUST_CODE" & vbCrLf _
            & ", SUM (DECODE(AGE_BUCKET,0,INV_BALANCE,0)) AGE_0" & vbCrLf _
            & ", SUM (DECODE(AGE_BUCKET,1,INV_BALANCE,0)) AGE_1" & vbCrLf _
            & ", SUM (DECODE(AGE_BUCKET,2,INV_BALANCE,0)) AGE_2" & vbCrLf _
            & ", SUM (DECODE(AGE_BUCKET,3,INV_BALANCE,0)) AGE_3" & vbCrLf _
            & ", SUM (DECODE(AGE_BUCKET,4,INV_BALANCE,0)) AGE_4" & vbCrLf _
            & ", SUM (DECODE(DUE_BUCKET,0,INV_BALANCE,0)) DUE_0" & vbCrLf _
            & ", SUM (DECODE(DUE_BUCKET,1,INV_BALANCE,0)) DUE_1" & vbCrLf _
            & ", SUM (DECODE(DUE_BUCKET,2,INV_BALANCE,0)) DUE_2" & vbCrLf _
            & ", SUM (DECODE(DUE_BUCKET,3,INV_BALANCE,0)) DUE_3" & vbCrLf _
            & ", SUM (DECODE(DUE_BUCKET,4,INV_BALANCE,0)) DUE_4" & vbCrLf _
            & " from " & ARTSTMTR & " ARTSTMTR group by CUST_CODE"
        Create_TDA(dst.Tables.Add, "ARTCUSTA", "**", 0, False, "", 1)
        Fill_Records("ARTCUSTA")


        With dst.Tables.Add("ARTSTMTZ")
            .Columns.Add("AR_PARM_KEY")
            .Columns.Add("REMIT0")
            .Columns.Add("REMIT1")
            .Columns.Add("REMIT2")
            .Columns.Add("REMIT3")
            .Columns.Add("AR_PARM_REMIT_MESSAGE")
            .Columns.Add("AR_PARM_DUNS_NO")
            .Columns.Add("ADDRESS_LINE")
            .Columns.Add("LOGO", GetType(System.Byte()))
            .Columns.Add("AR_PARM_FIN_CHG_RATE", GetType(System.Decimal))
            .Columns.Add("STMT_DATE", GetType(System.DateTime))
            .PrimaryKey = New DataColumn() {.Columns("AR_PARM_KEY")}
        End With

        Dim rowARTSTMTZ As DataRow = dst.Tables("ARTSTMTZ").NewRow
        With ROWs("ARTPARM1")
            rowARTSTMTZ.Item("AR_PARM_KEY") = "Z"
            rowARTSTMTZ.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowARTSTMTZ.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowARTSTMTZ.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            If .Item("AR_PARM_REMIT_PHONE") & "" <> "" And .Item("AR_PARM_REMIT_FAX") & "" <> "" Then
                rowARTSTMTZ.Item("REMIT3") = "" _
                    & " Tel " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_PHONE")) _
                    & ", Fax " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_FAX"))
            End If
            rowARTSTMTZ.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
            If 1 = 1 Then
                rowARTSTMTZ.Item("AR_PARM_REMIT_MESSAGE") = rowARTSTMTZ.Item("AR_PARM_REMIT_MESSAGE") & vbCrLf & .Item("AR_PARM_REMIT_MESSAGE_EXPORT")
            End If
            rowARTSTMTZ.Item("AR_PARM_DUNS_NO") = .Item("AR_PARM_DUNS_NO") & ""
            rowARTSTMTZ.Item("AR_PARM_FIN_CHG_RATE") = .Item("AR_PARM_FIN_CHG_RATE") & ""
            rowARTSTMTZ.Item("ADDRESS_LINE") = "" _
                & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ADDR1") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_CITY") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_STATE") _
                & " " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE") _
                & IIf(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "" <> "" _
                  And ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & "" <> "", "" _
                      & ", Tel " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "") _
                      & ", Fax " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & ""), "")
        End With
        rowARTSTMTZ.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        rowARTSTMTZ.Item("STMT_DATE") = PRD_END_DATE
        dst.Tables("ARTSTMTZ").Rows.Add(rowARTSTMTZ)

        ASCMAIN1.sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
        Create_TDA(dst.Tables.Add, "SOTSREP1", "**", 0, False, "", 1)
        Fill_Records("SOTSREP1")
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("SUBT", "")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class