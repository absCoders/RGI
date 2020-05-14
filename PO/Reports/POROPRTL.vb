Public Class POROPRTL

#Region "General Declarations"
    Private xDTE0 As Date
    Private xDTE1 As Date

    Dim SQLs As New Dictionary(Of String, String)

    Dim POTORDR1 As String
    Dim POTORDR2 As String

    Dim sqlPOTORDR1 As String
    Dim sqlPOTORDR2 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")
        Get_PARM("ICTPARM1")
        Absx1.optFor("RANGE").CheckedIndex = 1

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            Absx1.optFor("RANGE").CheckedIndex = 2
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "POs Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "POs Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = " and POTORDR1.PO_DATE_ORDERED between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'" & vbCrLf
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "S" Then
            SUBT = "Selected POs"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "U" Then
            SUBT = "All POs not Printed Yet"
            sqlw &= " and POTORDR1.PO_PRINTED_IND = '0'" & vbCrLf

        End If

        sqlw &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
        sqlw &= SQL_in("PO_ORDER_NO", "POTORDR1.PO_ORDER_NO")

        Prepare_dst(True, sqlw)

        Check_if_Empty("POTORDR1")
    End Sub

    Public Overrides Sub Print_Report()
        Dim PO_PARM_PO_RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
        If PO_PARM_PO_RPT <> "" Then RPT = PO_PARM_PO_RPT

        CR_params.Add("SUBT", "")
        CR_params.Add("FORM_TYPE", "P")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            Else

            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpDATE_RANGE.Enabled = (optRANGE.Value = "D")

        If optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()
        'Dim sql As String = ""
        'If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
        'Else
        '    sql = "Update POTORDR1 " _
        '        & " Set PO_PRINTED_IND = '1', PO_DATE_PRINTED = SYSDATE" _
        '        & " where (PO_ORDER_NO) in (Select PO_ORDER_NO from " & POTORDR1 & " )"
        '    ASCDATA1.ExecuteSQL(sql)
        'End If

        'sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
        '    & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'PO_PRT','PO Printed', PO_ORDER_NO" _
        '    & " from " & POTORDR1
        'ASCDATA1.ExecuteSQL(sql)
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("POTPARM1")
        Get_PARM("ICTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        sqlPOTORDR1 = "Select POTORDR1.* from POTORDR1 "
        ASCMAIN1.sql = sqlPOTORDR1 & ASCMAIN1.SQL_Add_WHERE(sqlw)
        POTORDR1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add Primary Key (PO_ORDER_NO)")

        sqlPOTORDR2 = "Select POTORDR2.* from POTORDR2, " & POTORDR1 _
            & " POTORDR1 where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
        ASCMAIN1.sql = sqlPOTORDR2
        POTORDR2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR2 & " Add Primary Key (PO_ORDER_NO, PO_ORDER_LNO)")

        SQLs.Clear()

        With dst
            ASCMAIN1.sql = "Select POTORDR1.*, 'Z' PO_PARM_KEY" & vbCrLf _
                & " from " & POTORDR1 & " POTORDR1"
            SQLs.Add("POTORDR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select POTPARM1.* from POTPARM1 where PO_PARM_KEY = 'Z'"
            Create_TDA(.Tables.Add, "POTPARM1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
                & " from " & POTORDR2 & " POTORDR2"
            SQLs.Add("POTORDR2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, False, "", 2)
            .Tables("POTORDR2").Columns.Add("DUPLICATE_IMAGE")

            Create_Relation("POTORDR1", "POTORDR2", "PO_ORDER_NO")

            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM" & vbCrLf _
                & ", ICTSTYL1.STYLE_COST, ICTSTYL1.STYLE_MATL_DESC, ICTSTYL1.COUNTRY_CODE" & vbCrLf _
                & ", ICTSTYL1.CASE_CUBE, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.LABEL_TYPE_CODE" & vbCrLf _
                & ", ICTSTYL1.DUTY_RATE_CODE, ICTDUTY1.DUTY_HTS_CODE, ICTSTYL1.IMAGE_NAME" & vbCrLf _
                & ", ICTSTYL1.SIZE_SCALE, ICTSTYL1.STYLE_CODE_PLM, ICTPLIN2.DESIGN_STYLE_NO" & vbCrLf _
                & " from ICTSTYL1,ICTDUTY1,ICTPLIN2" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE in (Select Distinct STYLE_CODE from " & POTORDR2 & ")" & vbCrLf _
                & "   and ICTDUTY1.DUTY_RATE_CODE (+) = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                & "   and ICTPLIN2.STYLE_CODE_PLM (+) = ICTSTYL1.STYLE_CODE_PLM"
            SQLs.Add("ICTSTYL1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 1)
            .Tables("ICTSTYL1").Columns.Add("IMAGE", GetType(System.Byte()))

            ASCMAIN1.sql = "Select ICTSTYC1.*" & vbCrLf _
                & " from ICTSTYC1" & vbCrLf _
                & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & POTORDR2 & ")"
            SQLs.Add("ICTSTYC1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False, "", 2)
                 
        End With

        Try
            Fill_Records("POTPARM1")
        Catch ex As Exception
            Stop
        End Try

        If perform_fill Then
            Fill_Records_RPT(New String() {sqlw})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        sqlw = parms(0)

        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR1)
        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR2)

        ASCDATA1.ExecuteSQL("Insert into " & POTORDR1 & " " & sqlPOTORDR1 & ASCMAIN1.SQL_Add_WHERE(sqlw))
        ASCDATA1.ExecuteSQL("Insert into " & POTORDR2 & " " & sqlPOTORDR2)

        EnforceConstraints(False)

        Fill_Records("POTORDR1")
        Fill_Records("POTORDR2")
        Fill_Records("ICTSTYL1")
        Fill_Records("ICTSTYC1")
 
        EnforceConstraints(True)
    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                sqlw = "APTVEND1.VEND_CODE in (Select Distinct VEND_CODE from POTORDR1)"
        End Select
        Return sqlw
    End Function
End Class