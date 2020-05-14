Public Class SORCANC0

#Region "General Declarations"

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""
        'sqlw &= " and SOTINVH1.INV_NO_REV_BY is null" & vbCrLf

        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTORDR1")
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
        ASCMAIN1.sql = "Update SOTCANC0 " _
            & " Set REGISTER_IND = '1', REGISTER_DATE = SYSDATE, REGISTER_XNO = '" & XNO & "'" _
            & " where REGISTER_IND = '0'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        ASCMAIN1.sql = "Update SOTCANC0 Set REGISTER_IND = '0' where REGISTER_XNO is Null"
        ASCDATA1.ExecuteSQL()

        With dst

            ASCMAIN1.sql = "Select SOTORDR2.*, SOTCANC0.ORDR_QTY_CANC_NOW" & vbCrLf _
                & " from SOTCANC0,SOTORDR2" & vbCrLf _
                & " where SOTCANC0.REGISTER_IND = '0'" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTCANC0.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTCANC0.ORDR_LNO"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 0)
            
            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from SOTCANC0 where REGISTER_IND = '0')"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)
            .Tables("SOTORDR1").Columns.Add("SO_PARM_KEY")
            .Tables("SOTORDR1").Columns("SO_PARM_KEY").DefaultValue = "Z"

            ASCMAIN1.sql = "Select ARTCUST1.*" & vbCrLf _
                & " from ARTCUST1" & vbCrLf _
                & " where CUST_CODE in (Select Distinct CUST_CODE from SOTORDR1 where ORDR_NO in (Select Distinct ORDR_NO from SOTCANC0 where REGISTER_IND = '0'))"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTPARMR.*" & vbCrLf _
                & " from SOTPARMR" & vbCrLf _
                & " where SO_PARM_KEY = 'Z'"
            Create_TDA(.Tables.Add, "SOTPARMR", "**", 0, False, "", 1)

            With .Tables("SOTPARMR")
                '.Columns.Add("REMIT0")
                '.Columns.Add("REMIT1")
                '.Columns.Add("REMIT2")
                '.Columns.Add("REMIT3")
                .Columns.Add("ADDRESS_LINE1")
                .Columns.Add("ADDRESS_LINE2")
                .Columns.Add("LOGO", GetType(System.Byte()))
            End With

            Create_TDA(.Tables.Add, "SOTSREP1", "*", 0)
        End With

        Dim rowSOTPARMR As DataRow = Fill_Record("SOTPARMR")
        With ROWs("ARTPARM1")
            'rowSOTPARMR.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            'rowSOTPARMR.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            'rowSOTPARMR.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
            '        & .Item("AR_PARM_REMIT_STATE") & " " _
            '        & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
            '        & .Item("AR_PARM_REMIT_COUNTRY")
            'If .Item("AR_PARM_REMIT_PHONE") & "" <> "" And .Item("AR_PARM_REMIT_FAX") & "" <> "" Then
            '    rowSOTPARMR.Item("REMIT3") = "" _
            '        & "  Tel " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_PHONE")) _
            '        & ", Fax " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_FAX"))
            'End If

            rowSOTPARMR.Item("ADDRESS_LINE1") = "" _
                & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ADDR1") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_CITY") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_STATE") _
                & " " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE")

            rowSOTPARMR.Item("ADDRESS_LINE2") = "" _
                      & "Phone " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "") _
                      & ", Fax " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & "")
        End With
        rowSOTPARMR.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""

        If parms.Length > 0 Then
             
        End If
            
        EnforceConstraints(False)
        Fill_Records("SOTORDR2")
        Fill_Records("SOTORDR1")
        Fill_Records("ARTCUST1")
        Fill_Records("SOTSREP1")
        EnforceConstraints(True)
    End Sub
  
    Sub Email_Invoice(ByVal rowSOTINVH1 As DataRow, ByVal rowARTCUST1 As DataRow, ByVal attachment As String)
        Me.Cursor = Cursors.WaitCursor

        Using frmTAFSEND1 As New TAFSEND1(Me)

            With frmTAFSEND1
                .EMAIL_KEY = "INV"
                .SEND_TO = rowARTCUST1.Item("CUST_EMAIL_TO") & ""
                If ASCMAIN1.USER_EMAIL = "" Then
                    .SEND_FROM = "donotreply" & "@" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_DEFAULT_EMAIL_DOMAIN")
                Else
                    .SEND_FROM = ASCMAIN1.USER_EMAIL
                End If
                .SEND_FROM_NAME = ASCMAIN1.USER_NAME
                If rowARTCUST1.Item("CUST_EMAIL_CC") & "" <> "" Then
                    .SEND_CC = rowARTCUST1.Item("CUST_EMAIL_CC") & ""
                End If

                Dim customInfo As String = rowARTCUST1.Item("CUSTOM_SUBJECT") & " "
                Dim companyName As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " "
                Dim custPO As String = rowSOTINVH1.Item("ORDR_CUST_PO") & " "
                Dim invNo As String = "INV " & rowSOTINVH1.Item("ORDR_INV_NO") & " "
                Dim invDate As String = rowSOTINVH1.Item("ORDR_INV_DATE") & ""

                Dim subjectLine As String = customInfo & companyName & custPO & invNo & invDate

                .SEND_SUBJECT = subjectLine

                Dim sal As String = ""
                If rowARTCUST1.Item("CUST_SALUTATION") & "" <> "" Then
                    sal = rowARTCUST1.Item("CUST_SALUTATION") & "," & vbCrLf
                Else
                    sal = "To whom it may concern," & vbCrLf
                End If

                Dim body As String = ""
                If "" <> "" Then
                    body = "Please find your invoice attached."
                Else
                    body = rowARTCUST1.Item("CUST_BILLING_NOTE") & "" <> ""
                End If

                .SEND_BODY = sal & body
                .SEND_ATTACHMENT = attachment
                .SEND_METHOD = "E"
                .SEND_ENTITY_CAPTION = "Sold-To"
                .SEND_ENTITY_TABLE = "ARTCUST1"
                .SEND_ENTITY_KEY = rowSOTINVH1.Item("CUST_CODE")
                .SEND_ENTITY_NAME = rowARTCUST1.Item("CUST_NAME") & ""

                .Send_email_automatically(False)

                If .SEND_STATUS <> "S" Then
                    TAC.TACMAIN1.Record_Event("SOTORDR1", rowSOTINVH1.Item("SO_ORDER_NO"), DATETIME_STAMP, ASCMAIN1.USER_ID, "E", "Emailed Invoice to " & .SEND_TO, rowSOTINVH1.Item("CUST_CODE"))
                Else
                    MsgBox("Error Occured: Could Not Send Email for Invoice: " & rowSOTINVH1.Item("ORDR_INV_NO"), MsgBoxStyle.OkOnly, "Error")
                End If
            End With
        End Using
    End Sub
End Class