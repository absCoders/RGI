Imports ABSolution
Imports Infragistics.Win

Public Class POROPEN1

    Dim POTOPEN1 As String = ""
    Dim H(12) As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Range_Events(grpPO_DATE_RANGE)
        Get_PARM("POTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()
        ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        Dim DTX_DATE As String = IIf(Absx1.optFor("OPTDATE").Value = "ETD", "POTORDR2.PO_DATE_SHIP_BY", "POTORDR2.PO_DATE_ETA")

        Dim DTX As String = IIf(Absx1.chkFor("CHKADD_TERMS").Checked, _
                                DTX_DATE, _
                                DTX_DATE & " + NVL(TATTERM1.TERM_DAYS_DUE,0)")

        ' Extracts from Data Sources

        Dim P(12) As String
        If Absx1.optFor("OPTAGE").Value = "W" Then
            Dim dt As Date = Now.AddDays(-7)
            For i As Int32 = 0 To 12
                P(i) = Format(dt, "dd-MMM-yyyy")
                H(i) = Format(dt, "MM/dd/yy")
                dt = dt.AddDays(7)
            Next
        Else
            For i As Int32 = 0 To 12
                P(i) = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, i - 1)
                H(i) = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, i), False, True)
            Next
            DTX = "TO_CHAR(" & DTX & ", 'YYYYMM')"
        End If

        Dim VX As String = IIf(Absx1.optFor("OPTSHOW").Value = "U", _
                               "NVL(POTORDR2.PO_QTY_OPN,0)", _
                               "NVL(POTORDR2.PO_QTY_OPN,0) * NVL(POTORDR2.PO_COST,0)")

        ' NEED TO DEVELOP THE SQL FOR PLANS

        ' Create the Work File - as a flattened out result set containing all PB columns and data fields that you would like to put in the detail section of the report

        Dim sqlsum As String = ""
        Dim sql As String = ""

        If POTOPEN1 <> "" Then
            ASCMAIN1.sql = "Delete from " & POTOPEN1
            ASCDATA1.ExecuteSQL()
        End If

        If Absx1.chkFor("CHKPO").Checked Then

            sql = "Select POTORDR1.PO_ORDER_NO, POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTORDR1.WHSE_CODE, POTORDR2.STYLE_CODE" & vbCrLf
            sql &= ", SUM (CASE WHEN " & DTX & " <= '" & P(0) & "' THEN " & VX & " ELSE 0 END) PO_PD" & vbCrLf

            sqlsum = "PO_TOT = NVL(PO_PD,0)"
            For i As Int32 = 1 To 12
                sql &= ", SUM (CASE WHEN " & DTX & " > '" & P(i - 1) & "' and " & DTX & " <= '" & P(i) & "' THEN " & VX & " ELSE 0 END) PO_" & Format(i, "00") & vbCrLf
                sqlsum &= "+NVL(PO_" & Format(i, "00") & ",0)"
            Next
            sql &= ", 0 PO_TOT" & vbCrLf

            sql &= ", 0 PP_PD" & vbCrLf
            For i As Int32 = 1 To 12
                sql &= ", 0 PP_" & Format(i, "00")
            Next
            sql &= ", 0 PP_TOT" & vbCrLf

            sql &= "   from POTORDR1,POTORDR2,ICTSTYL1" _
                & IIf(Absx1.chkFor("CHKADD_TERMS").Checked, "", ",TATTERM1") & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR1.PO_STATUS = 'O' " & vbCrLf _
                & "   and POTORDR2.PO_STATUS = 'O' " & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & IIf(Absx1.chkFor("CHKADD_TERMS").Checked, "", " and TATTERM1.TERM_CODE (+) = POTORDR1.TERM_CODE") & vbCrLf

            If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
                Dim z As String = Format(Absx1.dteFor("PO_DATE_F").Value, "dd-MMM-yyyy")
                sql = sql & " and POTORDR2.PO_DATE_ETA >= '" & z & "'" & vbCrLf
                Page0.Add("PO ETA Date >= " & z)
            End If
            If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
                Dim z As String = Format(Absx1.dteFor("PO_DATE_L").Value, "dd-MMM-yyyy")
                sql = sql & " and POTORDR2.PO_DATE_ETA <= '" & z & "'" & vbCrLf
                Page0.Add("PO ETA Date <= " & z)
            End If

            ' one of these for each filter in Report Maintenance
            sql &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
            sql &= SQL_in("STYLE_CODE", "POTORDR2.STYLE_CODE")
            sql &= SQL_in("WHSE_CODE", "POTORDR1.WHSE_CODE")
            sql &= SQL_in("STYLE_CLASS_CODE", "ICTSTYL1.STYLE_CLASS_CODE")


            sql &= " group by POTORDR1.PO_ORDER_NO, POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTORDR1.WHSE_CODE, POTORDR2.STYLE_CODE" & vbCrLf

            POTOPEN1 = ASCMAIN1.Temp_Table(sql)

            ASCMAIN1.sql = "Update " & POTOPEN1 & " Set " & sqlsum
            ASCDATA1.ExecuteSQL()
        End If

        If Absx1.chkFor("CHKPP").Checked Then

            VX = IIf(Absx1.optFor("OPTSHOW").Value = "U", _
                       "NVL(POTSHIP3.PO_QTY_SHP,0)", _
                       "NVL(POTSHIP3.PO_QTY_SHP,0) * NVL(POTORDR2.PO_COST,0)")

            sql = "Select POTORDR1.PO_ORDER_NO, POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTSHIP1.WHSE_CODE, POTORDR2.STYLE_CODE" & vbCrLf

            sql &= ", 0 PO_PD" & vbCrLf
            For i As Int32 = 1 To 12
                sql &= ", 0 PO_" & Format(i, "00")
            Next
            sql &= ", 0 PO_TOT" & vbCrLf

            sql &= ", SUM (CASE WHEN " & DTX & " <= '" & P(0) & "' THEN " & VX & " ELSE 0 END) PP_PD" & vbCrLf
            sqlsum = "PP_TOT = NVL(PP_PD,0)"
            For i As Int32 = 1 To 12
                sql &= ", SUM (CASE WHEN " & DTX & " > '" & P(i - 1) & "' and " & DTX & " <= '" & P(i) & "' THEN " & VX & " ELSE 0 END) PP_" & Format(i, "00") & vbCrLf
                sqlsum &= "+NVL(PP_" & Format(i, "00") & ",0)"
            Next
            sql &= ", 0 PP_TOT" & vbCrLf

            sql &= "   from POTORDR1,POTORDR2,ICTSTYL1,POTSHIP1,POTSHIP2,POTSHIP3" _
                & IIf(Absx1.chkFor("CHKADD_TERMS").Checked, "", ",TATTERM1") & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIP_STATUS = 'O' " & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & IIf(Absx1.chkFor("CHKADD_TERMS").Checked, "", " and TATTERM1.TERM_CODE (+) = POTORDR1.TERM_CODE") & vbCrLf

            If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
                Dim z As String = Format(Absx1.dteFor("PO_DATE_F").Value, "dd-MMM-yyyy")
                sql = sql & " and POTSHIP1.PO_SHIP_ETA >= '" & z & "'" & vbCrLf
            End If
            If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
                Dim z As String = Format(Absx1.dteFor("PO_DATE_L").Value, "dd-MMM-yyyy")
                sql = sql & " and POTSHIP1.PO_SHIP_ETA <= '" & z & "'" & vbCrLf
            End If

            ' one of these for each filter in Report Maintenance
            sql &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
            sql &= SQL_in("STYLE_CODE", "POTORDR2.STYLE_CODE")
            sql &= SQL_in("WHSE_CODE", "POTSHIP1.WHSE_CODE")
            sql &= SQL_in("STYLE_CLASS_CODE", "ICTSTYL1.STYLE_CLASS_CODE")


            sql &= " group by POTORDR1.PO_ORDER_NO, POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTSHIP1.WHSE_CODE, POTORDR2.STYLE_CODE" & vbCrLf

            If POTOPEN1 = "" Then
                POTOPEN1 = ASCMAIN1.Temp_Table(sql)
            Else
                ASCDATA1.ExecuteSQL("Insert into " & POTOPEN1 & " " & sql)
            End If

            ASCMAIN1.sql = "Update " & POTOPEN1 & " Set " & sqlsum
            ASCDATA1.ExecuteSQL()
        End If

        'ASCMAIN1.sql = "Select * from " & POTOPEN1
        '.Tables.Add(ASCDATA1.GetDataTable("**", "POTOPEN1", 1))

        MyBase.Get_SQL("*", POTOPEN1)
        ASCMAIN1.Progress("Building Tiers")

        sql = "Select " & sql_SELECT_cols & ASTSRPT1_sum_columns & vbCrLf _
            & " from " & POTOPEN1 & " POTOPEN1 " & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        For i As Int32 = 1 To 12
            Dim rowASTDSQLS As DataRow
            rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PO_" & Format(i, "00") & "'")(0)
            rowASTDSQLS.Item("COLUMN_CAPTION") = "PO " & H(i)
            rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PP_" & Format(i, "00") & "'")(0)
            rowASTDSQLS.Item("COLUMN_CAPTION") = "PP " & H(i)
        Next

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        If Not Absx1.chkFor("CHKPO_DATE_F").Checked _
        Or Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
            SUBT = SUBT & "ETA Dates"
            If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
                SUBT = SUBT & " from " & Format(Absx1.dteFor("PO_DATE_F").Value, "MM/dd/yyyy")
            End If
            If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
                SUBT = SUBT & " thru " & Format(Absx1.dteFor("PO_DATE_L").Value, "MM/dd/yyyy")
            End If
        End If

        CR_params.Add("PO", IIf(Absx1.chkFor("CHKPO").Checked, "1", "0"))
        CR_params.Add("PP", IIf(Absx1.chkFor("CHKPP").Checked, "1", "0"))
        CR_params.Add("PO_TEXT", Absx1.chkFor("CHKPO").Text)
        CR_params.Add("PP_TEXT", Absx1.chkFor("CHKPP").Text)

        For i As Int32 = 1 To 12
            CR_params.Add("H" & Format(i, "00"), H(i))
        Next
        Generate_Report(RPT, , SUBT)
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        If eItemKey = "Proceed" Then
            If Not Absx1.chkFor("CHKPO").Checked And Not Absx1.chkFor("CHKPP").Checked Then
                EMsg &= vbCr & "You must select at least 1: Purchase Orders and/or Purchase Shipments"
            End If
        End If
    End Sub

End Class
