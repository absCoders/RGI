Public Class SORSDTL1

#Region "Declarations"

    Dim SOTINVH2 As String
    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        grpPERIOD_RANGE.Visible = True
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left

        'Absx1.optFor("RANGE").CheckedIndex = 2
        'Absx1.optFor("RANGE").CheckedIndex = 2 ' don't know why setting this twice is nec

        chkCURR.Visible = (ASCMAIN1.CLIENT = "NYA")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 60, "RYP0", 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Invoices Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Invoices Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sql_filter = " and SOTINVH1.INV_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"

        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Invoices Posted in " & xRYP0_legend
            Else
                SUBT = "Invoices Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sql_filter = " and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & xRYP0 & "' and '" & xRYP1 & "'"

        End If

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")

        Dim SOURCE_TABLE_NAME As String = "SOTINVH2"

        sql = "Select " & sql_SELECT_cols & Replace(Replace(COLUMN_NAMEs_appended, ",", ",SOTINVH2."), "SOTINVH2.PICK_NO", "SOTINVH1.PICK_NO") & vbCrLf _
        & " from " & SOURCE_TABLE_NAME & " SOTINVH2 " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & "" ' " group by " & sql_GROUP_BY_cols

        If chkCURR.Checked Then
            sql &= " and SOTINVH1.CURR_CODE = 'CAD'"
        End If

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 _
        & " (G1,G2,G3,G4,G5,G6,G7,G8,G9" _
        & COLUMN_NAMEs_appended _
        & ") (" & sql & ")"

        If chkCURR.Checked Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, ",SOTINVH2.ORDR_UNIT_PRICE", ",SOTINVH2.ORDR_UNIT_PRICE_CURR ORDR_UNIT_PRICE")
        End If
        ASCDATA1.ExecuteSQL()

        If chkOriginalStyle.Checked Then
            ASCDATA1.ExecuteSQL("Create Index I_" & ASTSRPT1 & "_INV_NO_INV_LNO on " & ASTSRPT1 & " (INV_NO,INV_LNO)")

            Dim sqlG_SUB As String = ""
            Dim sqlG_RANGE As String = ""
            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("STYLE_CODE")
            If Val(rowASTDSQLA.Item("SEQUENCE") & "") <> 0 Then
                sqlG_SUB = ", G" & CStr(Val(rowASTDSQLA.Item("SEQUENCE") & "")) & " = R1.STYLE_CODE_SUB" & vbCrLf
                sqlG_RANGE = ", G" & CStr(Val(rowASTDSQLA.Item("SEQUENCE") & "")) & " = R1.RANGE_STYLE_CODE" & vbCrLf
            End If

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select ASTSRPT1.INV_NO, ASTSRPT1.INV_LNO, SOTORDR2.STYLE_CODE_SUB, SOTORDR2.RANGE_STYLE_CODE" & vbCrLf _
                & " from " & ASTSRPT1 & " ASTSRPT1,SOTPICK2,SOTORDR2" & vbCrLf _
                & " where SOTPICK2.PICK_NO = ASTSRPT1.PICK_NO and SOTPICK2.PICK_LNO = ASTSRPT1.INV_LNO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and (SOTORDR2.STYLE_CODE_SUB is Not Null or SOTORDR2.RANGE_STYLE_CODE is Not Null);" & vbCrLf _
                & " Begin for R1 in C1 Loop" & vbCrLf _
                & "  If R1.RANGE_STYLE_CODE is Not Null Then" & vbCrLf _
                & "   Update " & ASTSRPT1 & vbCrLf _
                & "    Set STYLE_CODE = R1.RANGE_STYLE_CODE" & vbCrLf _
                & sqlG_RANGE _
                & "    where INV_NO = R1.INV_NO and INV_LNO = R1.INV_LNO;" & vbCrLf _
                & "  Else " & vbCrLf _
                & "   Update " & ASTSRPT1 & vbCrLf _
                & "    Set STYLE_CODE = R1.STYLE_CODE_SUB" & vbCrLf _
                & sqlG_SUB _
                & "    where INV_NO = R1.INV_NO and INV_LNO = R1.INV_LNO;" & vbCrLf _
                & "  End If;" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select G1,G2,G3,G4,G5,G6,G7,G8,G9," & vbCrLf _
                & "INV_TYPE,INV_NO,MIN (INV_LNO) INV_LNO,STYLE_CODE,COLOR_CODE,CUST_CODE,ORDR_UNIT_PRICE," & vbCrLf _
                & "SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP,PICK_NO" & vbCrLf _
                & " from " & ASTSRPT1 & vbCrLf _
                & " group by G1,G2,G3,G4,G5,G6,G7,G8,G9," & vbCrLf _
                & "INV_TYPE,INV_NO,STYLE_CODE,COLOR_CODE,CUST_CODE,ORDR_UNIT_PRICE,PICK_NO"

            Dim ASTSRPT1_SUM As String = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Delete from " & ASTSRPT1
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & ASTSRPT1 & " Select * from " & ASTSRPT1_SUM
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Select SOTINVH2.* from SOTINVH2," & ASTSRPT1 & " ASTSRPT1" & vbCrLf _
            & " where SOTINVH2.INV_TYPE = ASTSRPT1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = ASTSRPT1.INV_NO" & vbCrLf _
            & "   and SOTINVH2.INV_LNO = ASTSRPT1.INV_LNO"
        Create_TDA(dst.Tables.Add, "SOTINVH2", "**", 0, False, "", 3)
        Fill_Records("SOTINVH2")

        ASCMAIN1.sql = "Select Distinct SOTINVH1.*" & vbCrLf _
            & ",SOTSHIP1.ORDR_GROUP_NO,SOTSHIP1.SHIP_ADDR_TYPE,NVL(SOTSHIP1.SHIP_ADDR_CODE,SOTINVH1.CUST_STORE_NO)" & vbCrLf _
            & " from SOTINVH1," & ASTSRPT1 & " ASTSRPT1, SOTSHIP1" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = ASTSRPT1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = ASTSRPT1.INV_NO" & vbCrLf _
            & "   and SOTSHIP1.SHIP_BOL_NO (+) = SOTINVH1.SHIP_BOL_NO"
        Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, False, "", 2)
        Fill_Records("SOTINVH1")

        ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC from ICTSTYL1 where STYLE_CODE in " & vbCrLf _
            & " (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")" & vbCrLf _
            & " union " & vbCrLf _
            & "Select ASTSRPT1.STYLE_CODE, ICTSTYL1.STYLE_DESC from " & ASTSRPT1 & " ASTSRPT1, SOTINVH2, ICTSTYL1" & vbCrLf _
            & " where SOTINVH2.INV_NO = ASTSRPT1.INV_NO and SOTINVH2.INV_LNO = ASTSRPT1.INV_LNO" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
            & "   and ASTSRPT1.STYLE_CODE IN (SELECT DISTINCT STYLE_CODE FROM " & ASTSRPT1 & " MINUS SELECT STYLE_CODE FROM ICTSTYL1)"

        ASCMAIN1.sql = "Select STYLE_CODE, MIN (STYLE_DESC) STYLE_DESC from (" _
            & ASCMAIN1.sql & ") group by STYLE_CODE"

        Create_TDA(dst.Tables.Add, "ICTSTYL1", "**", 0, False, "", 1)
        Fill_Records("ICTSTYL1")

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1 where CUST_CODE in " _
            & " (Select Distinct CUST_CODE from " & ASTSRPT1 & ")"
        Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)
        Fill_Records("ARTCUST1")

    End Sub
     
    Public Overrides Sub Print_Report()
        ' CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        CR_params.Add("SORT_CUST", IIf(tblASTDSQLA.Rows.Find("CUST_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        CR_params.Add("SORT_STYLE", IIf(tblASTDSQLA.Rows.Find("STYLE_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        ' SUBT = ""
        'If RYP0 = RYP1 Then
        '    SUBT = "Sales Posted in " & RYPLEGEND0
        'Else
        '    SUBT = "Sales Posted in " & RYPLEGEND0 & " thru " & RYPLEGEND1
        'End If

        If chkSummary.Checked Then
            RPT = "SORSDTL2"
            SUBT &= " - Sales Summarized by Order Group"
        Else
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                RPT = "SORSDTL3"
            End If
        End If

        If chkCURR.Checked Then
            SUBT &= " - Sales in $CAD"
        End If

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If Absx1.cmbFor("RYP").Value & "" = "" Then
            '    EMsg &= vbCr & "You must Specify a Report Period"
            'End If

            If ASCMAIN1.CLIENT = "NYA" Then
                If chkCURR.Checked Then
                    Dim xx As String = SQLA("ROYALTY_CODE")
                    If Split(xx, ",").Contains("CA") And Split(xx, ",").Length > 1 Then
                        EMsg &= vbCr & "You cannot mix CA with other Royalty Codes on the same report when reporting in CAD"
                    ElseIf xx = "" Then
                        EMsg &= vbCr & "You must select  Royalty Code CA (only) when reporting in CAD"
                    End If
                End If
            End If
        End If
    End Sub

    Overrides Sub Verify_Special_Pre(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            End If

            If tblASTDSQLA.Rows.Find("CUST_CODE").Item("SEQUENCE") & "" = "" _
                And tblASTDSQLA.Rows.Find("STYLE_CODE").Item("SEQUENCE") & "" = "" Then
                EMsg &= vbCr & "You must Specify either Customer or Style somwhere in the Sort Sequence"
            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub
End Class