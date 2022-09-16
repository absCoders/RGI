Imports System.Text
Imports System.Math

Public Class PORACCR1
    Dim POTACCR1 As String
    Dim RYP_Legend As String = ""
    Dim JOURNAL_LNO As Integer = 0
    Dim NYP As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("POTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        RYP_Legend = Absx1.cmbFor("RYP0").Value
        RYP = Mid(RYP_Legend, 1, 4) & Mid(RYP_Legend, 6, 2)

        NYP = ASCMAIN1.Period_Calc(RYP, 1)

        If RYP = ASCMAIN1.CYP Then

            If ASCMAIN1.EOM <> "1" Then
                RWU = "N"
            End If

            ASCMAIN1.sql = $"
            SELECT '{ASCMAIN1.CYP}' OPS_YYYYPP, X.* FROM (
            SELECT 'S' STATUS, POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, MIN (POTORDR1.VEND_CODE) VEND_CODE
            , POTSHIP1.PO_DATE_SHIPPED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP2.TRAN_NO RECEIPT_NO
            , SUM (POTSHIP3.PO_QTY_SHP) QTY
            , SUM (POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST_LANDED) AMT
            , POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.CONTAINER_NO
            FROM POTSHIP2,POTSHIP3,POTORDR1,POTSHIP1
            WHERE POTSHIP2.ACCRUAL_STATUS = '0' AND POTSHIP2.TRAN_NO IS NULL
               AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO
               AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO
               AND POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO
               AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO
            GROUP BY POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.CONTAINER_NO
            , POTSHIP1.PO_DATE_SHIPPED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP2.TRAN_NO
            UNION
            SELECT 'R' STATUS, ICTIREC1.PO_SHIPMENT_NO, ICTIREC1.PO_SHIPMENT_LNO, ICTIREC1.VEND_CODE
            , POTSHIP1.PO_DATE_SHIPPED, POTSHIP2.PO_DATE_RECEIVED, ICTIREC1.RECEIPT_NO
            , ICTIREC1.QTY_REC QTY
            , ICTIREC1.AMT_REC AMT
            , POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO, POTSHIP2.CONTAINER_NO
            FROM ICTIREC1,POTSHIP2,POTSHIP1
            WHERE ICTIREC1.ACCRUAL_STATUS = '0'
               AND POTSHIP2.PO_SHIPMENT_NO = ICTIREC1.PO_SHIPMENT_NO
               AND POTSHIP2.PO_SHIPMENT_LNO = ICTIREC1.PO_SHIPMENT_LNO
               AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO
            ) X"

        Else
            RWU = "N"
            ASCMAIN1.sql = $"Select * from POTACCR1 where OPS_YYYYPP = '{RYP}'"
        End If

        POTACCR1 = ASCMAIN1.Temp_Table


        Dim sqlw As String = ""
        sqlw &= SQL_in("VEND_CODE", "POTACCR1.VEND_CODE")
        If sqlw <> "" Then RWU = "N"

        With dst
            ASCMAIN1.sql = $"Select POTACCR1.* from {POTACCR1} POTACCR1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
            Create_TDA(.Tables.Add, "POTACCR1", "**", , False)

            Create_TDA(.Tables.Add, "GLTINTF1", "*")
        End With

        Fill_Records("POTACCR1")

        GL_Interface()

        Check_if_Empty("POTACCR1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT)
        Print_GL()
    End Sub

    Overrides Sub Update_Record()

        Dim sql As String

        sql = $"Insert into POTACCR1 Select * from {POTACCR1}"
        ASCDATA1.ExecuteSQL(sql)

        GL_Update()

    End Sub

    Sub GL_Interface()

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_TYPE As String = "POAC"
        JOURNAL_LNO = 0

        Dim PO_ACCR_AMT_S As Decimal = Val(dst.Tables("POTACCR1").Compute("SUM(AMT)", "STATUS = 'S'") & "")
        Dim PO_ACCR_AMT_R As Decimal = Val(dst.Tables("POTACCR1").Compute("SUM(AMT)", "STATUS = 'R'") & "")

        Write_GLTINTF1(JOURNAL_NO, JOURNAL_TYPE, ROWs("POTPARM1").Item("PO_PARM_ACCT_ACCR_SHP"), PO_ACCR_AMT_S, "")
        Write_GLTINTF1(JOURNAL_NO, JOURNAL_TYPE, ROWs("POTPARM1").Item("PO_PARM_ACCT_ACCR_REC"), PO_ACCR_AMT_R, "")

        Write_GLTINTF1(JOURNAL_NO, JOURNAL_TYPE, ROWs("POTPARM1").Item("PO_PARM_ACCT_ACCR_LIA"), -1 * (PO_ACCR_AMT_R + PO_ACCR_AMT_S), "")

    End Sub

    Sub Write_GLTINTF1(JOURNAL_NO As String, JOURNAL_TYPE As String, ACCT_CODE As String, DETL_POSTING_AMT As Decimal, Optional VEND_CODE As String = "")

        Dim rowGLTINTF1 As DataRow
        JOURNAL_LNO += 1

        For I As Integer = 0 To 1
            rowGLTINTF1 = dst.Tables("GLTINTF1").NewRow
            With rowGLTINTF1
                If I = 0 Then
                    .Item("OPS_YYYYPP") = RYP
                Else
                    .Item("OPS_YYYYPP") = NYP
                End If

                .Item("JOURNAL_NO") = JOURNAL_NO
                .Item("JOURNAL_LNO") = JOURNAL_LNO
                .Item("ACCT_CODE") = ACCT_CODE
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                .Item("DETL_POSTING_AMT") = DETL_POSTING_AMT
                .Item("DETL_EXE_NO") = XNO
                .Item("DETL_CTL_NO") = DBNull.Value
                .Item("DETL_CTL_LNO") = DBNull.Value
                .Item("DETL_CVX_NO") = VEND_CODE
                .Item("DETL_CVX_REF_DATE") = DBNull.Value
                .Item("DETL_CVX_REF_NO") = DBNull.Value
                .Item("DETL_DESC") = DBNull.Value
                .Item("DETL_CVX_TYPE") = "V"
                .Item("JOURNAL_TYPE") = JOURNAL_TYPE
            End With
            dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)

            DETL_POSTING_AMT = -1 * DETL_POSTING_AMT
        Next

    End Sub
End Class