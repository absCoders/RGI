Imports System.Math

Public Class SORSLSJ1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Dim REGISTER_DATE As Date

    Dim SOTINVH1 As String
    Dim SOTINVHU As String
    Dim SOTINVHG As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("SOTPARM1")
        Call Get_PARM("ARTPARM1")

        Absx1.optFor("RANGE").CheckedIndex = 2
        Absx1.optFor("RANGE").CheckedIndex = 2
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        dteINV_DATE_CUTOFF.DateTime = CDate(Now + ASCMAIN1.NowTSD).Date.AddDays(-1)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left

        If MENU_ITEM_PP = "CJ" Then
            optRANGE.Value = "P"
            grpSelectBy.Visible = False
            grpInclude.Visible = False
            grpFactorOptions.Visible = False
        End If


        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
        Else
            grpFactorOptions.Visible = False
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        EnforceConstraints(False)
        SUBT = ""

        Dim sqlw As String = "NVL(SOTINVH1.REGISTER_IND,'0') = '0'"
        Dim segs As String = ""

        REGISTER_DATE = DATETIME_STAMP.Date '  Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy")
        If chkINV_DATE_CUTOFF.Checked Then
            REGISTER_DATE = dteINV_DATE_CUTOFF.Value
            SUBT = "Invoice Cut-Off Date " & Format(REGISTER_DATE, "MM/dd/yyyy")
            sqlw &= " and INV_DATE <= '" & Format(REGISTER_DATE, "dd-MMM-yyyy") & "'"
        End If

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Invoices Posted " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Invoices Posted between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "SOTINVH1.REGISTER_IND = '1' and SOTINVH1.REGISTER_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"
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
            sqlw = "SOTINVH1.REGISTER_IND = '1' and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        End If

        If MENU_ITEM_PP = "CJ" Then
            RWU = "N"
        End If

        sqlw &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE")
        sqlw &= SQL_in("ORDR_TYPE_CODE", "SOTINVH1.ORDR_TYPE_CODE")
        sqlw &= SQL_in("SALES_DIVISION_CODE", "SOTINVH1.SALES_DIVISION_CODE")
        sqlw &= SQL_in("INV_NO", "SOTINVH1.INV_NO")
        sqlw &= SQL_in("SREP_CODE", "SOTINVH1.SREP_CODE")

        If ASCMAIN1.CLIENT = "NYA" AndAlso Absx1.optFor("RANGE").Value = "P" Then
            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("SEG4_CODE")
            Dim CODE_VALUES As String = rowASTDSQLA.Item("CODE_VALUES") & ""
            Dim EXCLUDE As String = rowASTDSQLA.Item("EXCLUDE") & ""
            If CODE_VALUES = "001" Then
                Page0.Add("CAD")
                HFs.Add("NYAG-CAD", "NYAG Canada (CAD)")
                RWU = "N"
            End If
            sqlw &= SQL_in("SEG4_CODE", "NVL(ICTWHSE1.SEG4_CODE,NVL(SOTSDIV1.SEG4_CODE,'000'))")
        End If


        If optFACTORED.Value = "1" Then
            sqlw &= " and NVL(SOTINVH1.CUST_FACTOR_IND,'0') = '1'"
        ElseIf optFACTORED.Value = "0" Then
            sqlw &= " and NVL(SOTINVH1.CUST_FACTOR_IND,'0') = '0'"
        End If

        If Not Absx1.chkFor("INV_TYPE_I").Checked Or Not Absx1.chkFor("INV_TYPE_C").Checked Then
            If Absx1.chkFor("INV_TYPE_I").Checked Then
                sqlw &= " and SOTINVH1.INV_TYPE = 'I'"
            End If
            If Absx1.chkFor("INV_TYPE_C").Checked Then
                sqlw &= " and SOTINVH1.INV_TYPE = 'C'"
            End If
        End If

        ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1,ICTWHSE1,SOTSDIV1" & vbCrLf _
            & " where ICTWHSE1.WHSE_CODE (+) = SOTINVH1.WHSE_CODE" & vbCrLf _
            & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and " & sqlw
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_NO)")

        Dim SOTINVH2 As String = "SOTINVH2"
        If ASCMAIN1.CLIENT = "NYA" Then
            If Page0.Contains("CAD") Then
                ASCMAIN1.sql = "UPDATE " & SOTINVH1 & " SET INV_SALES = INV_SALES_CURR, INV_FREIGHT = INV_FREIGHT_CURR, INV_TOTAL_AMOUNT = INV_TOTAL_AMOUNT_CURR, INV_MISC_CHG = INV_MISC_CHG_CURR, GST_TAX = GST_TAX_CURR, INV_STAX = INV_STAX_CURR"
                ASCDATA1.ExecuteSQL()

                SOTINVH2 = ""

                ASCMAIN1.sql = "Select SOTINVH2.* from SOTINVH2," & SOTINVH1 & " SOTINVH1 where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO"
                SOTINVH2 = ASCMAIN1.Temp_Table
                ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add Primary Key (INV_NO, INV_LNO)")

                ASCMAIN1.sql = "UPDATE " & SOTINVH2 & " SET ORDR_UNIT_PRICE = ORDR_UNIT_PRICE_CURR"
                ASCDATA1.ExecuteSQL()

            End If
        End If

        ASCMAIN1.sql = "Select SOTINVH1.*,ARTCUST1.CUST_NAME " & vbCrLf _
        & " from " & SOTINVH1 & " SOTINVH1,ARTCUST1,ICTWHSE1,SOTSDIV1 " & vbCrLf _
        & " where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
        & "   and ICTWHSE1.WHSE_CODE (+) = SOTINVH1.WHSE_CODE" & vbCrLf _
        & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE"

        Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, False)
        With dst.Tables("SOTINVH1")
            .Columns.Add("INV_SALES_OTHER_DIV", GetType(System.Decimal))
            .Columns.Add("INV_COGS_OTHER_DIV", GetType(System.Decimal))
            .PrimaryKey = New DataColumn() {.Columns("INV_TYPE"), .Columns("INV_NO"), .Columns("SALES_DIVISION_CODE")}
        End With

        Fill_Records("SOTINVH1")


        ASCMAIN1.sql = "Select ICTSTYL1.SALES_DIVISION_CODE, SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
            & ", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
            & ", Sum (ORDR_QTY_SHIP * ORDR_UNIT_PRICE) as TOTAL_SALES" & vbCrLf _
            & ", Sum (ORDR_QTY_SHIP * ORDR_UNIT_COST) as TOTAL_COSTS" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1, " & SOTINVH2 & " SOTINVH2, ICTSTYL1" & vbCrLf _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
            & "   and NVL(SOTINVH1.ORDR_TYPE_CODE,'???') <> 'XFR'" & vbCrLf _
            & " group by ICTSTYL1.SALES_DIVISION_CODE, SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE"

        Create_TDA(dst.Tables.Add, "SOTINVHC", "**", 0, False)

        Fill_Records("SOTINVHC")

        ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE SALES_DIVISION_CODE_DTL" & vbCrLf _
            & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS" & vbCrLf _
            & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0)) CGS" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1, " & SOTINVH2 & " SOTINVH2,ICTSTYL1 " & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and NVL(SOTINVH1.ORDR_TYPE_CODE,'???') <> 'XFR'" & vbCrLf _
            & " group by SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE"
        ASCMAIN1.sql = "Select X.* from (" & ASCMAIN1.sql & ") X where SALES_DIVISION_CODE <> SALES_DIVISION_CODE_DTL"
        Create_TDA(dst.Tables.Add, "SOTINVHD", "**", 0, False)

        Fill_Records("SOTINVHD")

        For Each row As DataRow In dst.Tables("SOTINVHD").Select("SALES_DIVISION_CODE <> SALES_DIVISION_CODE_DTL")
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows.Find(New String() {row.Item("INV_TYPE"), row.Item("INV_NO"), row.Item("SALES_DIVISION_CODE")})
            rowSOTINVH1.Item("INV_SALES_OTHER_DIV") = Val(rowSOTINVH1.Item("INV_SALES_OTHER_DIV") & "") + Val(row.Item("SLS") & "")
            rowSOTINVH1.Item("INV_SALES") = Val(rowSOTINVH1.Item("INV_SALES") & "") - Val(row.Item("SLS") & "")
            rowSOTINVH1.Item("INV_COGS_OTHER_DIV") = Val(rowSOTINVH1.Item("INV_COGS_OTHER_DIV") & "") + Val(row.Item("CGS") & "")
            rowSOTINVH1.Item("INV_COGS") = Val(rowSOTINVH1.Item("INV_COGS") & "") - Val(row.Item("CGS") & "")
            Dim rowSOTINVH1_DTL As DataRow = dst.Tables("SOTINVH1").NewRow
            rowSOTINVH1_DTL.ItemArray = rowSOTINVH1.ItemArray
            rowSOTINVH1_DTL.Item("INV_SALES") = Val(row.Item("SLS") & "")
            rowSOTINVH1_DTL.Item("INV_SALES_OTHER_DIV") = -1 * Val(row.Item("SLS") & "")
            rowSOTINVH1_DTL.Item("INV_COGS") = Val(row.Item("CGS") & "")
            rowSOTINVH1_DTL.Item("INV_COGS_OTHER_DIV") = -1 * Val(row.Item("CGS") & "")
            rowSOTINVH1_DTL.Item("INV_FREIGHT") = 0
            rowSOTINVH1_DTL.Item("INV_MISC_CHG") = 0
            rowSOTINVH1_DTL.Item("GST_TAX") = 0
            rowSOTINVH1_DTL.Item("INV_STAX") = 0
            rowSOTINVH1_DTL.Item("INV_TOTAL_AMOUNT") = 0
            rowSOTINVH1_DTL.Item("SALES_DIVISION_CODE") = row.Item("SALES_DIVISION_CODE_DTL")
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1_DTL)
        Next



        ASCMAIN1.sql = "Select X.*" & vbCrLf _
            & ", ARTSTAX1.STAX_DESC, ARTSTAX1.ACCT_CODE" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select SOTINVH1.CUST_SHIP_TO_STATE, SOTINVH1.STAX_CODE" & vbCrLf _
            & ", SUM (NVL(SOTINVH1.INV_STAX,0)) INV_STAX" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1 " & vbCrLf _
            & " where NVL(SOTINVH1.INV_STAX,0) <> 0" & vbCrLf _
            & " group by SOTINVH1.CUST_SHIP_TO_STATE, SOTINVH1.STAX_CODE" & vbCrLf _
            & ") X, ARTSTAX1 where ARTSTAX1.STAX_CODE (+) = X.STAX_CODE"

        Create_TDA(dst.Tables.Add, "SOTINVHT", "**", 0, False, "", 2)
        Fill_Records("SOTINVHT")
        ' BE CAREFUL- THERE AS A POOR CHOICE OF TABLE NAMES 
        ' - ORACLE TEMP TABLE SOTINVHT HAS NOTHING TO DO WITH THIS DATA TABLE, WHICH IS USED FOR SALES TAX 


        ASCMAIN1.sql = "Select * from SOTINVHU where ROWNUM < 1"
        SOTINVHU = ASCMAIN1.Temp_Table

        ' Sales & CGS Summary
        ' MAYBE SHOULD RECORD SEG2 FROM EVENT IF IT IS DEFINED THERE
        ' MAYBE SHOULD ALSO SET UP SEG3 AND SEG4 IN THIS SUMMARY

        ' NOTE - THE EXCLUSION OF XFR IN THE SQL BELOW IS HOW WE KEEP CGS AND INVTY FROM BEING UPDATED FOR XFR ORDERS
        ' NOTE - WE ARE HARD CODING THE PLACEMENT OF DIVISION AT SEG3 FOR NYA - MAYBE NEED SOMETHING DIFFERENT

        If ASCMAIN1.CLIENT = "RGI" Then
            segs = "" _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' " & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' " & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' " & vbCrLf
        ElseIf ASCMAIN1.CLIENT = "NYA" Then
            segs = "" _
            & ", NVL(SOTSDIV1.SEG2_CODE,NVL(SOTTYPE1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')) " & vbCrLf _
            & ", NVL(SOTSDIV1.SEG3_CODE,NVL(SOTTYPE1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "')) " & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' " & vbCrLf
        Else
            segs = "" _
            & ", NVL(SOTSDIV1.SEG2_CODE,NVL(SOTTYPE1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')) " & vbCrLf _
            & ", NVL(SOTSDIV1.SEG3_CODE,NVL(SOTTYPE1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "')) " & vbCrLf _
            & ", NVL(SOTSDIV1.SEG4_CODE,NVL(SOTTYPE1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "')) " & vbCrLf
        End If

        Dim sqlGL_BY_INV As String = ""
        If chkGL_BY_INV.Checked Then
            sqlGL_BY_INV = "SOTINVH1.INV_NO"
        Else
            sqlGL_BY_INV = "NULL"
        End If

        ASCMAIN1.sql = "" _
            & "Select '" & XNO & "' REGISTER_XNO, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE, SOTINVH1.ORDR_TYPE_CODE" & vbCrLf & segs _
            & ", SOTINVH2.INV_TYPE, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
            & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS " & vbCrLf _
            & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0)) CGS " & vbCrLf _
            & ", " & sqlGL_BY_INV & " INV_NO " & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1, " & SOTINVH2 & " SOTINVH2,ICTSTYL1,SOTTYPE1,SOTSDIV1 " & vbCrLf _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE (+) = SOTINVH2.STYLE_CODE" & vbCrLf _
            & "   and NVL(SOTINVH1.ORDR_TYPE_CODE,'???') <> 'XFR'" & vbCrLf _
            & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, ICTSTYL1.SALES_DIVISION_CODE, SOTINVH1.ORDR_TYPE_CODE" & vbCrLf & segs _
            & ", SOTINVH2.INV_TYPE, ICTSTYL1.STYLE_CLASS_CODE, " & sqlGL_BY_INV
        ASCDATA1.ExecuteSQL("Insert into " & SOTINVHU & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from " & SOTINVHU
        Create_TDA(dst.Tables.Add, "SOTINVHU", "**", 0, False)

        ASCMAIN1.sql = "Select X.*" & vbCrLf _
            & ", SOTMISC1.MISC_CHG_DESC, SOTMISC1.ACCT_CODE, SOTMISC1.MISC_GP" & vbCrLf _
            & " from (Select MISC_CHG_CODE, Sum (INV_MISC_CHG) INV_MISC_CHG from (" & vbCrLf _
            & "Select '1', SOTINVHM.MISC_CHG_CODE, SUM (SOTINVHM.INV_MISC_CHG) INV_MISC_CHG" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1, SOTINVHM " & vbCrLf _
            & " where SOTINVHM.INV_MISC_CHG <> 0" & vbCrLf _
            & "   and SOTINVHM.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   and SOTINVHM.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & " group by SOTINVHM.MISC_CHG_CODE" & vbCrLf
        ASCMAIN1.sql &= "" _
            & ") group by MISC_CHG_CODE" & vbCrLf _
            & ") X, SOTMISC1 where SOTMISC1.MISC_CHG_CODE (+) = X.MISC_CHG_CODE"
        Create_TDA(dst.Tables.Add, "SOTINVHM", "**", 0, False, "", 1)
        Fill_Records("SOTINVHM")

        Create_TDA(dst.Tables.Add, "ICTCLAS1", "*", 0)
        Fill_Records("ICTCLAS1")

        Create_TDA(dst.Tables.Add, "SOTSDIV1", "*", 0)
        Fill_Records("SOTSDIV1")

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1 where CUST_CODE in (Select Distinct CUST_CODE from " & SOTINVH1 & " SOTINVH1)"
        Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0)
        Fill_Records("ARTCUST1")

        ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC, SALES_DIVISION_CODE, STYLE_COST from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTINVH1 & " SOTINVH1,SOTINVH2 where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO)"
        Create_TDA(dst.Tables.Add, "ICTSTYL1", "**", 0)
        Fill_Records("ICTSTYL1")

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        ' DO THIS ALWAYS - THAT IS WHY i ADDED THE "" = ""
        If Absx1.optFor("RANGE").Value = "N" Or "" = "" Then

            ASCMAIN1.sql = "Select * from SOTINVHG where ROWNUM < 1"
            SOTINVHG = ASCMAIN1.Temp_Table

            Dim ACCT_CODE_sql As String
            Dim sqlG As String = ""

            ' Sales / Customer Returns

            If chkGL_BY_INV.Checked Then
                sqlGL_BY_INV = "SOTINVHU.INV_NO"
            Else
                sqlGL_BY_INV = "NULL"
            End If

            ACCT_CODE_sql = "DECODE(SOTINVHU.INV_TYPE,'I',NVL(ICTCLAS1.ACCT_CODE_SLS_SHP,'" & ROWs("SOTPARM1").Item("SO_PARM_ACCT_SLS_SHP") & "'),NVL(ICTCLAS1.ACCT_CODE_SLS_RTN,'" & ROWs("SOTPARM1").Item("SO_PARM_ACCT_SLS_RTN") & "'))"
            sqlG = "" _
                & " Select '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVHU.OPS_YYYYPP" & vbCrLf _
                & ", " & ACCT_CODE_sql & " ACCT_CODE" & vbCrLf _
                & ", SOTINVHU.SEG2_CODE, SOTINVHU.SEG3_CODE, SOTINVHU.SEG4_CODE" & vbCrLf _
                & ", SOTINVHU.ORDR_TYPE_CODE DETL_CVX_REF_NO" & vbCrLf _
                & ", 'L' DETL_CVX_TYPE, SOTINVHU.STYLE_CLASS_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * SOTINVHU.SLS) DIST_AMT " & vbCrLf _
                & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                & " from " & SOTINVHU & " SOTINVHU,ICTCLAS1,SOTSDIV1" & vbCrLf _
                & " where ICTCLAS1.STYLE_CLASS_CODE (+) = SOTINVHU.STYLE_CLASS_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVHU.SALES_DIVISION_CODE" & vbCrLf _
                & " group by SOTINVHU.OPS_YYYYPP, " & ACCT_CODE_sql & ", SOTINVHU.INV_TYPE, SOTINVHU.SEG2_CODE, SOTINVHU.ORDR_TYPE_CODE, SOTINVHU.STYLE_CLASS_CODE " & vbCrLf _
                & ", SOTINVHU.SEG2_CODE, SOTINVHU.SEG3_CODE, SOTINVHU.SEG4_CODE" & vbCrLf _
                & ", " & sqlGL_BY_INV & vbCrLf _
                & " order by SOTINVHU.OPS_YYYYPP, " & ACCT_CODE_sql & ", SOTINVHU.INV_TYPE, SOTINVHU.SEG2_CODE, SOTINVHU.ORDR_TYPE_CODE, SOTINVHU.STYLE_CLASS_CODE " & vbCrLf _
                & ", SOTINVHU.SEG2_CODE, SOTINVHU.SEG3_CODE, SOTINVHU.SEG4_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            ' Cost of Goods Sold / Returned

            ACCT_CODE_sql = "DECODE(SOTINVHU.INV_TYPE,'I',NVL(ICTCLAS1.ACCT_CODE_CGS_SHP,'" & ROWs("SOTPARM1").Item("SO_PARM_ACCT_CGS_SHP") & "'),NVL(ICTCLAS1.ACCT_CODE_CGS_RTN,'" & ROWs("SOTPARM1").Item("SO_PARM_ACCT_CGS_RTN") & "'))"

            sqlG = "" _
                & " Select '" & XNO & "' REGISTER_XNO, 'OPCJ' JOURNAL_TYPE, SOTINVHU.OPS_YYYYPP" & vbCrLf _
                & ", " & ACCT_CODE_sql & " ACCT_CODE" & vbCrLf _
                & ", SOTINVHU.SEG2_CODE, SOTINVHU.SEG3_CODE, SOTINVHU.SEG4_CODE" & vbCrLf _
                & ", SOTINVHU.ORDR_TYPE_CODE DETL_CVX_REF_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, NULL DETL_CVX_NO" & vbCrLf _
                & ", SUM (SOTINVHU.CGS) DIST_AMT " & vbCrLf _
                & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                & " from " & SOTINVHU & " SOTINVHU,ICTCLAS1,SOTSDIV1" & vbCrLf _
                & " where ICTCLAS1.STYLE_CLASS_CODE (+) = SOTINVHU.STYLE_CLASS_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVHU.SALES_DIVISION_CODE" & vbCrLf _
                & " group by SOTINVHU.OPS_YYYYPP, " & ACCT_CODE_sql & ", SOTINVHU.INV_TYPE, SOTINVHU.ORDR_TYPE_CODE" & vbCrLf _
                & ", SOTINVHU.SEG2_CODE, SOTINVHU.SEG3_CODE, SOTINVHU.SEG4_CODE" & vbCrLf _
                & ", " & sqlGL_BY_INV & vbCrLf _
                & " order by SOTINVHU.OPS_YYYYPP, " & ACCT_CODE_sql & ", SOTINVHU.INV_TYPE, SOTINVHU.ORDR_TYPE_CODE" & vbCrLf _
                & ", SOTINVHU.SEG2_CODE, SOTINVHU.SEG3_CODE, SOTINVHU.SEG4_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            sqlG = "" _
                & " SELECT '" & XNO & "' REGISTER_XNO, 'OPCJ' JOURNAL_TYPE, SOTINVHU.OPS_YYYYPP" & vbCrLf _
                & ", ICTCLAS1.ACCT_CODE_ONH ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", NULL DETL_CVX_REF_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, NULL DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * SOTINVHU.CGS) DIST_AMT " & vbCrLf _
                & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                & " from " & SOTINVHU & " SOTINVHU,ICTCLAS1" & vbCrLf _
                & " where ICTCLAS1.STYLE_CLASS_CODE (+) = SOTINVHU.STYLE_CLASS_CODE" & vbCrLf _
                & " group by SOTINVHU.OPS_YYYYPP, ICTCLAS1.ACCT_CODE_ONH, SOTINVHU.INV_TYPE" & vbCrLf _
                & ", " & sqlGL_BY_INV
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)


            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                segs = "" _
                & ", NVL(SOTMISC1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') " & vbCrLf _
                & ", NVL(SOTMISC1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') " & vbCrLf _
                & ", NVL(SOTMISC1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') " & vbCrLf
            Else
                'If ASCMAIN1.CLIENT = "NYA" Then ' BILL TO ADVISE WHEN TO MOVE TO COMPANY 001
                '    'segs = "" _
                '    '    & ", NVL(SOTSDIV1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') " & vbCrLf _
                '    '    & ", NVL(SOTSDIV1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') " & vbCrLf _
                '    '    & ", '000' " & vbCrLf
                'Else
                '    segs = "" _
                '    & ", NVL(SOTSDIV1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') " & vbCrLf _
                '    & ", NVL(SOTSDIV1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') " & vbCrLf _
                '    & ", NVL(SOTSDIV1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') " & vbCrLf
                'End If

                segs = "" _
                    & ", NVL(SOTSDIV1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') " & vbCrLf _
                    & ", NVL(SOTSDIV1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') " & vbCrLf _
                    & ", NVL(SOTSDIV1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') " & vbCrLf
            End If


            If chkGL_BY_INV.Checked Then
                sqlGL_BY_INV = "SOTINVH1.INV_NO"
            Else
                sqlGL_BY_INV = "NULL"
            End If

            Dim COLUMN_NAME_MC As String = "INV_MISC_CHG"
            If ASCMAIN1.CLIENT = "NYA" Then
                If Page0.Contains("CAD") Then
                    COLUMN_NAME_MC = "INV_MISC_CHG_CURR"
                End If
            End If

            sqlG = "" _
                & " SELECT '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", SOTMISC1.ACCT_CODE" & vbCrLf & segs _
                & ", DECODE(SOTINVH1.INV_TYPE,'I','MISC-I','MISC-C') DETL_CVX_REF_NO" & vbCrLf _
                & ", 'M' DETL_CVX_TYPE, SOTINVHM.MISC_CHG_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * NVL(SOTINVHM." & COLUMN_NAME_MC & ",0)) DIST_AMT " & vbCrLf _
                & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1,SOTINVHM,SOTMISC1,SOTSDIV1,SOTTYPE1" & vbCrLf _
                & " where SOTMISC1.MISC_CHG_CODE = SOTINVHM.MISC_CHG_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTINVHM.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
                & "   and SOTINVHM.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
                & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
                & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, SOTMISC1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTINVHM.MISC_CHG_CODE" & vbCrLf _
                & ", SOTINVH1.ORDR_YYYYPP_UPDATED, SOTMISC1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTINVHM.MISC_CHG_CODE" & segs & vbCrLf _
                & ", " & sqlGL_BY_INV
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)


            ' Sales Tax Payable

            sqlG = "" _
                & " SELECT " & XNO & " REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", ARTSTAX1.ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", DECODE(SOTINVH1.INV_TYPE,'I','STAX-I','STAX-C') DETL_CVX_REF_NO" & vbCrLf _
                & ", 'T' DETL_CVX_TYPE, SOTINVH1.CUST_SHIP_TO_STATE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * (NVL(SOTINVH1.INV_STAX,0))) DIST_AMT " & vbCrLf _
                & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                & " FROM " & SOTINVH1 & " SOTINVH1,ARTSTAX1,SOTSDIV1" & vbCrLf _
                & " where ARTSTAX1.STAX_CODE (+) = SOTINVH1.STAX_CODE" & vbCrLf _
                & "   and (SOTINVH1.STAX_CODE is Not Null or (NVL(SOTINVH1.INV_STAX,0)) <> 0)" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & " group by " & XNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, ARTSTAX1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTINVH1.CUST_SHIP_TO_STATE" & vbCrLf _
                & ", " & sqlGL_BY_INV
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)



            ' Freight Income

            sqlG = "" _
                & " SELECT '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", '" & ROWs("SOTPARM1").Item("SO_PARM_ACCT_FRT_INC") & "' ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", 'FRT-' || SOTINVH1.INV_TYPE DETL_CVX_REF_NO" & vbCrLf _
                & ", NULL DETL_CVX_TYPE, NULL DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * NVL(SOTINVH1.INV_FREIGHT,0)) DIST_AMT " & vbCrLf _
                & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1,SOTSDIV1,SOTTYPE1" & vbCrLf _
                & " where SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
                & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.INV_TYPE" & vbCrLf _
                & ", " & sqlGL_BY_INV
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            ' Misc Charge / Handling Income
            ' DO NOT ENABLE THIS SECTION FOR VDI (NYA) - B/C MISC CHARGE DETAILS ARE RECORDED IN SOTINVHM AND THERE IS A GL INTERFACE SECTION FOR THIS ABOVE

            'sqlG = "" _
            '    & " SELECT '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
            '    & ", '" & ROWs("SOTPARM1").Item("SO_PARM_ACCT_HND_FEE") & "' ACCT_CODE" & vbCrLf _
            '    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            '    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            '    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            '    & ", 'MISC-' || SOTINVH1.INV_TYPE DETL_CVX_REF_NO" & vbCrLf _
            '    & ", NULL DETL_CVX_TYPE, NULL DETL_CVX_NO" & vbCrLf _
            '    & ", SUM (-1 * NVL(SOTINVH1.INV_MISC_CHG,0)) DIST_AMT " & vbCrLf _
            '    & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
            '    & " FROM " & SOTINVH1 & " SOTINVH1,SOTSDIV1,SOTTYPE1" & vbCrLf _
            '    & " where SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
            '    & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            '    & " GROUP BY SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.INV_TYPE" & vbCrLf _
            '    & " ORDER BY SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.INV_TYPE" & vbCrLf _
            '    & ", " & sqlGL_BY_INV
            'ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            ' Walmart discounts are now in SOTINVHM - March 2016 - Effective April 2016
            'If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            '    Dim rowARTREAS1 As DataRow = LookUp("ARTREAS1", "DI")
            '    Dim ACCT_CODE As String = rowARTREAS1.Item("ACCT_CODE")

            '    sqlG = "" _
            '      & " Select '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
            '      & ", '" & ACCT_CODE & "' ACCT_CODE" & vbCrLf _
            '      & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            '      & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            '      & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            '      & ", 'WALMARTDISC' DETL_CVX_REF_NO" & vbCrLf _
            '      & ", NULL DETL_CVX_TYPE, NULL DETL_CVX_NO" & vbCrLf _
            '      & ", SUM (1 * 0.025 * NVL(SOTINVH1.INV_SALES,0)) DIST_AMT " & vbCrLf _
            '      & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
            '      & " from " & SOTINVH1 & " SOTINVH1,SOTSDIV1" & vbCrLf _
            '      & " where SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
            '      & "   and SOTINVH1.CUST_CODE = 'WALMART'" & vbCrLf _
            '      & " group by SOTINVH1.ORDR_YYYYPP_UPDATED" & vbCrLf _
            '      & ", " & sqlGL_BY_INV
            '    ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            '    sqlG = Replace(sqlG, "1 * 0.025", "-1 * 0.025")
            '    ' sqlG = Replace(sqlG, ", '" & ACCT_CODE & "' ACCT_CODE", ", '" & rowGLTBANK1.Item("ACCT_CODE") & "' ACCT_CODE")
            '    sqlG = Replace(sqlG, ", '" & ACCT_CODE & "' ACCT_CODE", ", '1250' ACCT_CODE")
            '    ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            'End If


            ' AR - Normal AR

            sqlG = "" _
                & " SELECT '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", ARTPOST1.ACCT_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') SEG2_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') SEG3_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') SEG4_CODE" & vbCrLf _
                & ", DECODE(SOTINVH1.INV_TYPE,'I','AR-I','AR-C') DETL_CVX_REF_NO" & vbCrLf _
                & ", 'R' DETL_CVX_TYPE, SOTINVH1.POST_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (SOTINVH1.INV_TOTAL_AMOUNT) DIST_AMT " & vbCrLf _
                & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1,ARTPOST1,SOTSDIV1" & vbCrLf _
                & " where ARTPOST1.POST_CODE (+) = SOTINVH1.POST_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and NVL(SOTINVH1.CUST_FACTOR_IND,'0') = '0'" & vbCrLf _
                & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, ARTPOST1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTINVH1.POST_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')" & vbCrLf _
                & ", NVL(ARTPOST1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "')" & vbCrLf _
                & ", NVL(ARTPOST1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "')" & vbCrLf _
                & ", " & sqlGL_BY_INV
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            ' AR - Due from Factor

            ' THIS CODE USES A BANK CODE THAT REPRESENTS DUE FROM FACTOR
            'sqlG = "" _
            '    & " SELECT '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
            '    & ", GLTBANK1.ACCT_CODE" & vbCrLf _
            '    & ", NVL(GLTBANK1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') SEG2_CODE" & vbCrLf _
            '    & ", NVL(GLTBANK1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') SEG3_CODE" & vbCrLf _
            '    & ", NVL(GLTBANK1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') SEG4_CODE" & vbCrLf _
            '    & ", DECODE(SOTINVH1.INV_TYPE,'I','AR-I','AR-C') DETL_CVX_REF_NO" & vbCrLf _
            '    & ", 'R' DETL_CVX_TYPE, SOTINVH1.POST_CODE DETL_CVX_NO" & vbCrLf _
            '    & ", SUM (SOTINVH1.INV_TOTAL_AMOUNT) DIST_AMT " & vbCrLf _
            '    & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
            '    & " from " & SOTINVH1 & " SOTINVH1,GLTBANK1,SOTSDIV1" & vbCrLf _
            '    & " where GLTBANK1.BANK_CODE (+) = '" & ROWs("ARTPARM1").Item("AR_PARM_BANK_CODE_FACTOR") & "'" & vbCrLf _
            '    & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
            '    & "   and NVL(SOTINVH1.CUST_FACTOR_IND,'0') = '1'" & vbCrLf _
            '    & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, GLTBANK1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTINVH1.POST_CODE" & vbCrLf _
            '    & ", NVL(GLTBANK1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')" & vbCrLf _
            '    & ", NVL(GLTBANK1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "')" & vbCrLf _
            '    & ", NVL(GLTBANK1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "')" & vbCrLf _
            '    & ", " & sqlGL_BY_INV

            ' THIS CODE USES A GL ACCT - FOR USE WHEN WE DO NOT SET UP A BANK CODE THAT REPRESENTS DUE FROM FACTOR

            If ASCMAIN1.CLIENT = "NYA" Then
                sqlG = "" _
                    & " SELECT '" & XNO & "' REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                    & ", DECODE(SOTSDIV1.SEG4_CODE,'001','1251','1250') ACCT_CODE" & vbCrLf _
                    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                    & ", DECODE(SOTINVH1.INV_TYPE,'I','AR-I','AR-C') DETL_CVX_REF_NO" & vbCrLf _
                    & ", 'R' DETL_CVX_TYPE, SOTINVH1.POST_CODE DETL_CVX_NO" & vbCrLf _
                    & ", SUM (SOTINVH1.INV_TOTAL_AMOUNT) DIST_AMT " & vbCrLf _
                    & ", NULL DETL_CTL_DATE, " & sqlGL_BY_INV & " DETL_CTL_NO, NULL DETL_CTL_LNO " & vbCrLf _
                    & " from " & SOTINVH1 & " SOTINVH1,SOTSDIV1" & vbCrLf _
                    & " where SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                    & "   and NVL(SOTINVH1.CUST_FACTOR_IND,'0') = '1'" & vbCrLf _
                    & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.INV_TYPE, SOTINVH1.POST_CODE" & vbCrLf _
                    & ", DECODE(SOTSDIV1.SEG4_CODE,'001','1251','1250')" & vbCrLf _
                    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'" & vbCrLf _
                    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'" & vbCrLf _
                    & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'" & vbCrLf _
                    & ", " & sqlGL_BY_INV
                ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)
            End If

            ASCDATA1.ExecuteSQL("Delete from " & SOTINVHG & " where DIST_AMT = 0")

            If MENU_ITEM_PP = "CJ" Then
                'Prepare_GL_Interface("OPCJ")
            Else
                Prepare_GL_Interface("OPSJ")
            End If

        End If

        Check_if_Empty("SOTINVH1")
    End Sub

    Public Overrides Sub Print_Report()

        'Dim AS_PARM_INST_NAME As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & ""
        'If ASCMAIN1.CLIENT = "NYA" AndAlso Page0.Contains("CAD") Then
        '    ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") = "NYAG Canada (CAD)"
        'End If

        RPT_TITLE = "Sales Journal Invoice & Memo Register"
        CR_params.Add("SUMMARY", "0")
        CR_params.Add("CJ", IIf(MENU_ITEM_PP = "CJ", "1", "0"))
        Generate_Report(RPT, RPT_TITLE, SUBT)

        RPT_TITLE = "Sales Journal Summary by Style"
        CR_params.Add("CHKCOSTS", IIf(MENU_ITEM_PP = "CJ", "1", "0"))
        Generate_Report("SORMTDS3", RPT_TITLE, SUBT)

        RPT_TITLE = "Sales Journal Summary by Customer"
        CR_params.Add("CHKCOSTS", IIf(MENU_ITEM_PP = "CJ", "1", "0"))
        Generate_Report("SORMTDS4", RPT_TITLE, SUBT)

        If MENU_ITEM_PP = "CJ" Then
        Else
            Print_GL()
        End If

        If chkGL_BY_INV.Checked Then
            grdGLTINTF1.DataSource = dst.Tables("GLTINTF1")
            ASCMAIN1.grdInitializeLayout(grdGLTINTF1)
            Show_Filter(grdGLTINTF1, True)
            grdGLTINTF1.DisplayLayout.Bands(0).SortedColumns.Add("DETL_CTL_NO", False, True)
            If grdGLTINTF1.Tag = "" Then
                grdGLTINTF1.Tag = "X"
                Create_Summary(grdGLTINTF1, "DETL_POSTING_AMT")
                grdGLTINTF1.Visible = True
            End If
        End If

        'If ASCMAIN1.CLIENT = "NYA" AndAlso Page0.Contains("CAD") Then
        '    ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") = AS_PARM_INST_NAME
        'End If

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "N" Then
                Dim dte() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
                If Format(dteINV_DATE_CUTOFF.Value, "yyyyMMdd") < Format(dte(1), "yyyyMMdd") _
                Or Format(dteINV_DATE_CUTOFF.Value, "yyyyMMdd") > Format(dte(dte.Length - 1), "yyyyMMdd") Then
                    EMsg &= vbCr & "Cut-Off Date must be between " & Format(dte(1), "MM/dd/yyyy") & " and " & Format(dte(dte.Length - 1), "MM/dd/yyyy") & " - Current Period is " & ASCMAIN1.CYP
                End If
            End If


            If MENU_ITEM_PP = "CJ" Then
                'If ASCMAIN1.EOM = "1" Then
                '    If tblASTDSQLA.Select("CODE_VALUES <> ''").Length <> 0 Then
                '        EMsg &= vbCr & "Filters are not permitted on Period End Report"
                '    End If
                'End If
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

        chkINV_DATE_CUTOFF.Visible = (optRANGE.Value = "N")
        dteINV_DATE_CUTOFF.Visible = (optRANGE.Value = "N")

    End Sub

    Overrides Sub Update_Record()

        If MENU_ITEM_PP = "CJ" Then
            ' no update to mark the rows? they are retreived by period for the CJ report, 
            '  so that would have to be modified as well if we were to update a flag here
        Else
            Dim sql As String = "Update SOTINVH1 " _
                & " Set REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2, REGISTER_DATE = :PARM3" _
                & " where INV_NO in (Select INV_NO from " & SOTINVH1 & " )"
            ASCDATA1.ExecuteSQL(sql, "VVD", New Object() {"1", MyBase.XNO, REGISTER_DATE})
        End If

        ' maybe we should record these only at month end in the CJ?

        ASCDATA1.ExecuteSQL("Insert into SOTINVHU Select * from " & SOTINVHU)
        ASCDATA1.ExecuteSQL("Insert into SOTINVHG Select * from " & SOTINVHG)

        GL_Update()
    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Decimal
        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))
        If chkINV_DATE_CUTOFF.Checked Then
            DETL_CTL_DATE = dteINV_DATE_CUTOFF.Value
        End If

        ASCMAIN1.sql = "Select * from " & SOTINVHG & " where JOURNAL_TYPE = '" & JOURNAL_TYPE & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            DETL_POSTING_AMT = Val(row.Item("DIST_AMT") & "")
            Dim DETL_CVX_NO As String = row.Item("DETL_CVX_NO") & ""
            Dim DETL_CVX_REF_NO As String = row.Item("DETL_CVX_REF_NO") & ""
            Dim DETL_CVX_TYPE As String = row.Item("DETL_CVX_TYPE") & ""

            Dim rowGLTINTF1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").NewRow
            rowGLTINTF1("OPS_YYYYPP") = row("OPS_YYYYPP")
            rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
            JOURNAL_LNO += 1
            rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
            rowGLTINTF1("ACCT_CODE") = row("ACCT_CODE")
            rowGLTINTF1("SEG2_CODE") = row("SEG2_CODE")
            rowGLTINTF1("SEG3_CODE") = row("SEG3_CODE")
            rowGLTINTF1("SEG4_CODE") = row("SEG4_CODE")
            rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
            rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
            rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
            rowGLTINTF1("DETL_CTL_NO") = row("DETL_CTL_NO")
            rowGLTINTF1("DETL_CTL_LNO") = row("DETL_CTL_LNO")
            rowGLTINTF1("DETL_CVX_NO") = DETL_CVX_NO
            rowGLTINTF1("DETL_CVX_REF_DATE") = REGISTER_DATE
            rowGLTINTF1("DETL_CVX_REF_NO") = DETL_CVX_REF_NO
            rowGLTINTF1("DETL_DESC") = DBNull.Value
            rowGLTINTF1("DETL_CVX_TYPE") = DETL_CVX_TYPE
            rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
            ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Next

        Return JOURNAL_NO

    End Function

    Private Sub chkINV_DATE_CUTOFF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkINV_DATE_CUTOFF.CheckedChanged
        dteINV_DATE_CUTOFF.Visible = chkINV_DATE_CUTOFF.Checked
    End Sub

    Private Sub chkGL_BY_INV_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkGL_BY_INV.CheckedChanged

    End Sub
End Class