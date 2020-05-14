Public Class EDR855T1

    Dim EDT855T1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        ASCMAIN1.sql = "Update EDT855T1 Set EDI_PROCESS_IND = '0', COMPANY_CODE = 'NYA', EDI_TP_ID = TRIM(EDI_TP_ID)" & vbCrLf _
            & " where EDI_PROCESS_IND is Null AND (TRIM(EDI_TP_QUAL),TRIM(EDI_TP_ID),TRIM(EDI_OUR_ID),TRIM(EDI_OUR_QUAL)) IN " & vbCrLf _
            & " (SELECT TRIM(EDI_TP_QUAL),TRIM(EDI_TP_ID),TRIM(EDI_OUR_ID),TRIM(EDI_OUR_QUAL)" & vbCrLf _
            & " from EDTTRPM1 where EDI_DOC_NO = '855')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from EDT855T1 where EDI_PROCESS_IND = '0'"
        EDT855T1 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Select * from " & EDT855T1
        Create_TDA(dst.Tables.Add, "EDT855T1", "**", 0, False, "", 1)
        Fill_Records("EDT855T1")

        Check_if_Empty("EDT855T1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
       
    End Sub

    Overrides Sub Update_Record()
        ASCDATA1.ExecuteSQL("Update EDT855T1 Set EDI_PROCESS_IND = '1' where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from " & EDT855T1 & ")")
        TAC.SOCMAIN1.Update_Credit_Authorizations()
    End Sub

End Class