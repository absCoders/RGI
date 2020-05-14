Public Class SORORDR1

#Region "Declarations"
    Dim SOTORDR1 As String = ""
#End Region

#Region "ABS Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

    End Sub

    Public Overrides Sub Print_Report()

        SUBT = ""
        Generate_Report(RPT, , SUBT)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"

        End Select
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        If SOTORDR1 = "" Then
            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sqlw)
            SOTORDR1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")
        End If

        With dst

            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & " where ORDR_NO in (Select ORDR_NO from " & SOTORDR1 & ")"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)


            ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTSTYL1.CASE_CUBE, ICTSTYC1.UPC_CODE" & vbCrLf _
                & " from SOTORDR2,ICTCOLR1,ICTSTYL1,ICTSTYC1" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO in (Select ORDR_NO from " & SOTORDR1 & ")" & vbCrLf _
                    & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and ICTSTYC1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and ICTSTYC1.COLOR_CODE = SOTORDR2.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)
            .Tables("SOTORDR2").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(ORDR_QTY,0) / ISNULL(CARTON_PACK_QTY,0))")
            .Tables("SOTORDR2").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")

            With .Tables("SOTORDR2").Columns
                .Add("RANGE_STYLE_QTY_PER_PP", GetType(System.Int64))
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_ALLO", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            End With
             
        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""

        If parms.Length > 0 Then
            sqlw = parms(0)
            ASCDATA1.ExecuteSQL("Delete from " & SOTORDR1)
            ASCDATA1.ExecuteSQL("Insert into " & SOTORDR1 & " Select ORDR_NO from SOTORDR1" & ASCMAIN1.SQL_Add_WHERE(sqlw))
        End If

        EnforceConstraints(False)
        Fill_Records("SOTORDR1")
        Fill_Records("SOTORDR2")
        EnforceConstraints(True)
    End Sub

#End Region
End Class