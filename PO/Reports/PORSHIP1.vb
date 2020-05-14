Imports System.Math

Public Class PORSHIP1

    Dim POTSHIP0 As String
    Dim sqlPOTSHIP0 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("ICTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()
 
        SUBT = ""
        Dim sqlw As String = ""

        If Absx1.chkFor("CHKOPEN_ONLY").Checked Then
            sqlw &= " and PO_SHIPMENT_NO in (Select Distinct PO_SHIPMENT_NO from POTSHIP2 where PO_SHIP_STATUS = 'O')"
        End If

        sqlw &= SQLA_filter("PO_SHIPMENT_NO", "POTSHIP1")
        Prepare_dst(True, sqlw)
        Check_if_Empty("POTSHIP1")
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("MODE", "S")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.chkFor("CHKOPEN_ONLY").Checked Then
            Else
                If tblASTDSQLA.Select("CODE_VALUES <> ''").Length = 0 Then
                    EMsg &= vbCr & "You must Specify some Filter Criteria"
                End If
            End If
        End If
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"

        sqlPOTSHIP0 = "Select POTSHIP1.PO_SHIPMENT_NO from POTSHIP1"
        ASCMAIN1.sql = sqlPOTSHIP0 & " where ROWNUM < 1"
        POTSHIP0 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTSHIP0 & " Add Primary Key (PO_SHIPMENT_NO)")

        ASCMAIN1.sql = "Select * from POTSHIP1 where PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from " & POTSHIP0 & ")"
        Create_TDA(dst.Tables.Add, "POTSHIP1", "**", 0, False, "", 1)

        ASCMAIN1.sql = "Select * from POTSHIP2 where PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from " & POTSHIP0 & ")"
        Create_TDA(dst.Tables.Add, "POTSHIP2", "**", 0, False, "", 2)
        'With .Tables("POTSHIP2")
        '    .Columns.Add("CLOSE")
        'End With

        ASCMAIN1.sql = "Select POTSHIP3.* " & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE " & vbCrLf _
            & ", POTORDR2.PO_QTY_OPN, POTORDR2.PO_QTY_UOM, POTORDR2.PO_COST ORDR2_COST" & vbCrLf _
            & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_BODY_CODE, POTORDR2.SUB_UNIT_PACK_QTY, POTORDR2.CARTON_PACK_QTY" & vbCrLf _
            & ", POTORDR1.PO_REFERENCE, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
            & ", POTSHIP3.PO_QTY_REC PO_QTY_REC_OLD" & vbCrLf _
            & " from POTSHIP3,POTORDR2,ICTSTYL1,POTORDR1 " & vbCrLf _
            & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_NO in (Select PO_SHIPMENT_NO from " & POTSHIP0 & ")"
        Create_TDA(dst.Tables.Add, "POTSHIP3", "**", 0, False, "", 4)

        If perform_fill Then
            Fill_Records_RPT(New Object() {sqlw})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        
        If parms IsNot Nothing Then
            sqlw = CStr(parms(0))

            ASCDATA1.ExecuteSQL("Truncate Table " & POTSHIP0)
            ASCDATA1.ExecuteSQL("Insert into " & POTSHIP0 & " " & sqlPOTSHIP0 & ASCMAIN1.SQL_Add_WHERE(sqlw))
 
        End If

        EnforceConstraints(False)
        Fill_Records("POTSHIP1")
        Fill_Records("POTSHIP2")
        Fill_Records("POTSHIP3")
        EnforceConstraints(True)
    End Sub
End Class