Public Class SORPCKR1

    Dim SOTPICKX As String
    Dim OPTSORT As String = ""
    Dim SOTSHIPC As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Range_Events(grpPICK_RELEASED)
        Range_Events(grpORDR_SHIP_DATE)
        Range_Events(grpORDR_CANCEL_DATE)

        If ASCMAIN1.CLIENT <> "VAN" Then
            Absx1.chkFor("CHKPOSHIPMENT").Visible = False
        End If

        Get_PARM("SOTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        If ASCMAIN1.DBS_COMPANY = "VAN" And ASCMAIN1.DBS_SERVER = "VAN" Then
            TAC.WHCMAIN1.Update_ADS_SOTSHIP1()
        End If

        ASCMAIN1.Progress("Building Work File")

        ASCMAIN1.sql = "SELECT SOTPICK1.*" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE_TO, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
            & " from SOTPICK1, SOTORDR0, SOTSHIP1, SOTORDR1" & vbCrLf _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "   and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
            & Get_Dates()
        ASCMAIN1.sql &= SQL_in("CUST_CODE", "SOTORDR0.CUST_CODE")
        ASCMAIN1.sql &= SQL_in("PICK_BATCH_NO", "SOTPICK1.PICK_BATCH_NO")
        ASCMAIN1.sql &= SQL_in("WHSE_CODE", "SOTSHIP1.WHSE_CODE")

        OPTSORT = Absx1.optFor("OPTSORT").Value

        SOTPICKX = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add PICK_QTY NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add PICK_AMT NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Update " & SOTPICKX & " SOTPICKX Set PICK_QTY = (Select Sum (PICK_QTY) from SOTPICK2 where PICK_NO = SOTPICKX.PICK_NO)")
        ASCDATA1.ExecuteSQL("Update " & SOTPICKX & " SOTPICKX Set PICK_AMT = (Select Sum (NVL(SOTPICK2.PICK_QTY,0)*NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) from SOTPICK2,SOTORDR2 where SOTPICK2.PICK_NO = SOTPICKX.PICK_NO and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO)")

        ASCMAIN1.sql = "Select * from " & SOTPICKX
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTPICKX", 1))

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources
        ASCMAIN1.sql = "Select SOTSHIP1.* from SOTSHIP1" _
            & " where SOTSHIP1.SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTPICKX & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSHIP1", 1))
        dst.Tables("SOTSHIP1").Columns.Add("SHIP_WINDOW_CHG")
        dst.Tables("SOTSHIP1").Columns.Add("SHIP_UNITS_CHG")

        ASCMAIN1.sql = "Select Distinct SOTSHIP3.SHIP_BOL_NO from SOTSHIP3" & vbCrLf _
            & " where SOTSHIP3.SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTPICKX & ")" & vbCrLf _
            & "   and (NVL(ORDR_SHIP_DATE_OLD,TRUNC(SYSDATE)) <> NVL(ORDR_SHIP_DATE_NEW,TRUNC(SYSDATE))" & vbCrLf _
            & "    or" & vbCrLf _
            & "        NVL(ORDR_CANCEL_DATE_OLD,TRUNC(SYSDATE)) <> NVL(ORDR_CANCEL_DATE_NEW,TRUNC(SYSDATE)))"

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & ""
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                rowSOTSHIP1.Item("SHIP_WINDOW_CHG") = "1"
            Next
        End If


        ASCMAIN1.sql = "Select Distinct SOTSHIP3.SHIP_BOL_NO from SOTSHIP3,SOTSHIP6" & vbCrLf _
            & " where SOTSHIP3.SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTPICKX & ")" & vbCrLf _
            & "   and SOTSHIP6.SHIP_CHGREQ_NO = SOTSHIP3.SHIP_CHGREQ_NO" & vbCrLf

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & ""
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                rowSOTSHIP1.Item("SHIP_UNITS_CHG") = "1"
            Next
        End If


        ASCMAIN1.sql = "Select SOTORDR0.*, ARTCUST1.CUST_NAME from SOTORDR0, ARTCUST1" _
            & " where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE and SOTORDR0.ORDR_GROUP_NO in " _
            & " (Select Distinct ORDR_GROUP_NO from " & SOTPICKX & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTORDR0", 1))
        dst.Tables("SOTORDR0").Columns.Add("SORTBY")

        For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select("")
            Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO")

            If OPTSORT = "G" Then
                rowSOTORDR0.Item("SORTBY") = rowSOTORDR0.Item("ORDR_GROUP_NO")
            ElseIf OPTSORT = "P" Then
                rowSOTORDR0.Item("SORTBY") = rowSOTORDR0.Item("ORDR_CUST_PO")
            ElseIf OPTSORT = "C" Then
                rowSOTORDR0.Item("SORTBY") = Format(rowSOTORDR0.Item("ORDR_CANCEL_DATE"), "yyyyMMdd")
            End If

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Else
                If rowSOTORDR0.Item("ORDR_ORIG_SHIP_DATE") & "" <> "" _
                    And rowSOTORDR0.Item("ORDR_ORIG_CANCEL_DATE") & "" <> "" Then
                    If Format(rowSOTORDR0.Item("ORDR_ORIG_SHIP_DATE") & "", "yyyyMMdd") <> Format(rowSOTORDR0.Item("ORDR_SHIP_DATE") & "", "yyyyMMdd") _
                    Or Format(rowSOTORDR0.Item("ORDR_ORIG_CANCEL_DATE") & "", "yyyyMMdd") <> Format(rowSOTORDR0.Item("ORDR_CANCEL_DATE") & "", "yyyyMMdd") Then

                        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                            rowSOTSHIP1.Item("SHIP_WINDOW_CHG") = "1"
                        Next
                    End If
                End If
            End If
        Next


        ASCMAIN1.sql = "Select ORDR_NO SHIP_BOL_NO, STYLE_CODE, COLOR_CODE from SOTORDR2 where ROWNUM < 1"
            SOTSHIPC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPC & " Add Primary Key (SHIP_BOL_NO, STYLE_CODE, COLOR_CODE)")
        If Absx1.chkFor("CHKPOSHIPMENT").Checked Then
            ASCMAIN1.sql = "Select Distinct SOTPICKX.SHIP_BOL_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                     & " from SOTORDR2,SOTPICK2,SOTPICK1, " & SOTPICKX & " SOTPICKX " & vbCrLf _
                     & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                     & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                     & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                     & "   and SOTPICK1.PICK_NO = SOTPICKX.PICK_NO"
            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPC & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select Distinct * from (" & vbCrLf _
                    & "Select POTSHIP1.WHSE_CODE, SOTSHIPC.SHIP_BOL_NO" & vbCrLf _
                    & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
                    & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
                    & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO,  Sum(POTSHIP3.PO_QTY_SHP) over (PARTITION BY POTSHIP2.PO_SHIPMENT_NO) PO_QTY_SHP" & vbCrLf _
                    & ", POTORDR1.VEND_CODE" & vbCrLf _
                    & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
                    & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, " & SOTSHIPC & " SOTSHIPC " & vbCrLf _
                    & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                    & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO " & vbCrLf _
                    & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                    & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & "   and POTORDR2.STYLE_CODE = SOTSHIPC.STYLE_CODE" & vbCrLf _
                    & "   and POTORDR2.COLOR_CODE = SOTSHIPC.COLOR_CODE" & vbCrLf _
                    & ")"
            dst.Tables.Add(ASCDATA1.GetDataTable("", "POTSHIPX", 4))

    End Sub

    Function Get_Dates() As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"PICK_RELEASED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            Dim TABLE_NAME As String = "SOTORDR0"
            If COLUMN_NAME = "PICK_RELEASED" Then TABLE_NAME = "SOTPICK1"
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                sql = sql & " and " & TABLE_NAME & "." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
            End If
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                sql = sql & " and " & TABLE_NAME & "." & COLUMN_NAME & " <= '" & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "dd-MMM-yyyy") & "'"
            End If
        Next
        Return sql
    End Function

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        'ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = 0")
        'ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = TRUNC(100 * ORDR_AMT / ORDR_QTY) / 100 where ORDR_QTY <> 0")
    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = ""
        Page0.Add("Report Detail: " & IIf(Absx1.chkFor("CHKDETAILED").Checked, "Yes", "No"))
        For Each COLUMN_NAME As String In New String() {"PICK_RELEASED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            Dim Z As String = Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text & ":"
            Dim real_date_selected As Boolean = False
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                Z &= " from First"
            Else
                Z &= " from " & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "MM/dd/yyyy")
                real_date_selected = True
            End If
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                Z &= " to Last"
            Else
                Z &= " to " & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "MM/dd/yyyy")
                real_date_selected = True
            End If
            If real_date_selected Then
                SUBT &= ", " & Z
            End If
            Page0.Add(Z)
        Next
        If optPrimary.Value = "P" Then
        Else
            Page0.Add("Sorted by " & Absx1.optFor("OPTSORT").Text)
            SUBT &= ", " & "Sorted by " & Absx1.optFor("OPTSORT").Text
            SUBT = Mid(SUBT, 3)
        End If


        CR_params.Add("DETAILED", IIf(Absx1.chkFor("CHKDETAILED").Checked, "1", "0"))

        If optPrimary.Value = "P" Then
            RPT = "SORPCKR2"
        Else
            CR_params.Add("CHKSHOW_AMT", IIf(Absx1.chkFor("CHKSHOW_AMT").Checked, "1", "0"))
        End If
        Generate_Report(RPT, RPT_TITLE, SUBT)

     End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then

                    Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("WHSE_CODE")

                    Dim CODE_VALUES As String = (rowASTDSQLA("CODE_VALUES") & String.Empty).ToString.Trim
                    Dim validSelections As Boolean = CODE_VALUES.Length > 0

                    For Each whse As String In CODE_VALUES.Split(",")
                        whse = whse.Trim
                        If Not TAC.TACMAIN1.NyaCanadaWhseList.Contains(whse) Then
                            validSelections = False
                            Exit For
                        End If
                    Next

                    If rowASTDSQLA("EXCLUDE") & "" = "0" AndAlso validSelections Then
                        ' OK
                    Else
                        EMsg &= "Warehouses " & TAC.TACMAIN1.NyaCanadaWhseCommaSeparatedString & " Only"
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Private Sub optPrimary_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPrimary.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_PrimarySort()
    End Sub

    Sub Setup_PrimarySort()
        grpCustomer.Visible = (optPrimary.Value = "C")
    End Sub
End Class