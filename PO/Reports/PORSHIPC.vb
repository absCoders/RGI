Imports System.Math

Public Class PORSHIPC

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Shadows SUBT As String = ""

    Dim ICTIREC1 As String
    Dim sqlICTIREC1 As String = ""
    Dim variances_only As Boolean = False

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")
        Absx1.optFor("RANGE").CheckedIndex = 2

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        SUBT = ""

        Dim sqlw As String = "ICTIREC1.REGISTER_IND = '0'"

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Receipts Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Receipts Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "ICTIREC1.RECEIPT_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"

        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Receipts Posted in " & xRYP0_legend
            Else
                SUBT = "Receipts Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw = "ICTIREC1.OPS_YYYYPP between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"

        ElseIf Absx1.optFor("RANGE").Value = "A" Then
            sqlw = "(ICTIREC1.REGISTER_IND = '0' or ICTIREC1.REGISTER_IND = '1')"
            SUBT = "Selected Receipts"
            RWU = "N"
        End If

        If Absx1.chkFor("CHKVAR_ONLY").Checked Then
            RWU = "N"
        End If
        'If ASCMAIN1.EOM <> "1" Then
        'RWU = "N"
        'End If

        sqlw &= SQLA_filter("PO_SHIPMENT_NO", "ICTIREC1")
        sqlw &= SQLA_filter("VEND_CODE", "ICTIREC1")

        Prepare_dst(True, sqlw)

        Check_if_Empty("ICTIREC1")
    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("CHKVAR_ONLY", IIf(variances_only, "1", "0"))
        Generate_Report(RPT, , SUBT)
        If Absx1.optFor("RANGE").Value = "N" Then
            Call Print_GL()
        End If
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
            ElseIf Absx1.optFor("RANGE").Value = "A" Then
                If tblASTDSQLA.Select("CODE_VALUES <> ''").Length = 0 Then
                    EMsg &= vbCr & "You must Specify some Filter Criteria"
                End If
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

    Overrides Sub Update_Record()

        Dim sql As String = "Update ICTIREC1 " _
            & " Set REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2" _
            & " where RECEIPT_NO in (Select RECEIPT_NO from " & ICTIREC1 & " )"
        ASCDATA1.ExecuteSQL(sql, "VV", New Object() {"1", MyBase.XNO})

        'sql = "BEGIN DECLARE CURSOR C1 IS" _
        '& " SELECT ICTIREC2.OPS_YYYYPP, ICTITEM1.PRICE_CATGY_CODE" _
        '& ", SUM (NVL(ICTIREC2.QTY_REC,0) * (NVL(ICTIREC2.PO_COST,0) - NVL(ICTCOSTA.PRICE_CATGY_COST_TOTAL,0))) VAR" _
        '& " FROM ICTIREC2,ICTITEM1,ICTCOSTA,ICTCATG1,ICTPCAT1," & ICTIREC1 & " ICTIREC1" _
        '& " WHERE ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" _
        '& " AND ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP" _
        '& " AND ICTCOSTA.PRICE_CATGY_CODE = ICTITEM1.PRICE_CATGY_CODE" _
        '& " AND ICTPCAT1.PRICE_CATGY_CODE = ICTITEM1.PRICE_CATGY_CODE" _
        '& " AND ICTCATG1.PROD_CATGY_CODE = ICTPCAT1.PROD_CATGY_CODE" _
        '& " AND ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" _
        '& " AND (NVL(ICTIREC2.PO_COST,0) - NVL(ICTCOSTA.PRICE_CATGY_COST_TOTAL,0)) <> 0" _
        '& " GROUP BY ICTIREC2.OPS_YYYYPP, ICTITEM1.PRICE_CATGY_CODE;" _
        '& " BEGIN FOR R1 IN C1 LOOP" _
        '& "   Update ICTIVAR1 Set PV_EXP = NVL(PV_EXP,0) + NVL(R1.VAR,0)" _
        '& "    where PRICE_CATGY_CODE = R1.PRICE_CATGY_CODE and OPS_YYYYPP = R1.OPS_YYYYPP;" _
        '& "   If SQL%NOTFOUND Then " _
        '& "    Insert into ICTIVAR1 (PRICE_CATGY_CODE, OPS_YYYYPP, PV_EXP) " _
        '& "     Values (R1.PRICE_CATGY_CODE, R1.OPS_YYYYPP, R1.VAR);" _
        '& "   End If;" _
        '& " END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(sql)

        Call GL_Update()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"

        sqlICTIREC1 = "Select ICTIREC1.*" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ETA, POTSHIP1.PO_SHIP_ADV_DATE, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
            & ", POTSHIP2.CONTAINER_NO, POTSHIP2.BOL_NO, POTSHIP2.PO_SHIP_CTNS" & vbCrLf _
            & " from ICTIREC1, POTSHIP1, POTSHIP2" & vbCrLf _
            & " where POTSHIP1.PO_SHIPMENT_NO = ICTIREC1.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = ICTIREC1.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = ICTIREC1.PO_SHIPMENT_LNO" & vbCrLf
        ASCMAIN1.sql = sqlICTIREC1 & " and ROWNUM < 1"
        ICTIREC1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTIREC1 & " Add Primary Key (RECEIPT_NO)")

        ASCMAIN1.sql = "Select ICTIREC1.*,ICTWHSE1.WHSE_DESC,APTVEND1.VEND_NAME " _
            & " from " & ICTIREC1 & " ICTIREC1,ICTWHSE1,APTVEND1" _
            & " where ICTWHSE1.WHSE_CODE (+) = ICTIREC1.WHSE_CODE " _
            & "   and APTVEND1.VEND_CODE (+) = ICTIREC1.VEND_CODE "
        Call Create_TDA(dst.Tables.Add, "ICTIREC1", "**", 0)

        ASCMAIN1.sql = "Select ICTIREC2.*, ICTSTYL1.STYLE_DESC" _
            & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" _
            & ", POTSHIP3.PO_QTY_SHP" _
            & ", ICTSTYL1.STYLE_CLASS_CODE, ICTCOSTA.STYLE_COST STYLE_COST_TOTAL" _
            & " from ICTIREC2," & ICTIREC1 & " ICTIREC1, ICTSTYL1, POTORDR1, POTSHIP3, ICTCOSTA " _
            & " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" _
            & " and ICTCOSTA.OPS_YYYYPP (+) = ICTIREC2.OPS_YYYYPP" _
            & " and ICTCOSTA.STYLE_CODE (+) = ICTIREC2.STYLE_CODE" _
            & " and ICTCOSTA.COLOR_CODE (+) = ICTIREC2.COLOR_CODE" _
            & " and POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" _
            & " and POTSHIP3.PO_SHIPMENT_NO = ICTIREC2.PO_SHIPMENT_NO" _
            & " and POTSHIP3.PO_SHIPMENT_LNO = ICTIREC2.PO_SHIPMENT_LNO" _
            & " and POTSHIP3.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" _
            & " and POTSHIP3.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" _
            & " and ICTSTYL1.STYLE_CODE = ICTIREC2.STYLE_CODE"
        Create_TDA(dst.Tables.Add, "ICTIREC2", "**", 0)

        ASCMAIN1.sql = "Select ICTIREC5.*, GLTACCT1.ACCT_DESC " _
            & " from ICTIREC5," & ICTIREC1 & " ICTIREC1, GLTACCT1 " _
            & " where ICTIREC5.RECEIPT_NO = ICTIREC1.RECEIPT_NO" _
            & " and GLTACCT1.ACCT_CODE = ICTIREC5.ACCT_CODE"
        Create_TDA(dst.Tables.Add, "ICTIREC5", "**", 0)

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        ASCMAIN1.sql = "Select * from ICTCLAS1"
        Create_TDA(dst.Tables.Add, "ICTCLAS1", "**", 0)
        Fill_Records("ICTCLAS1")

        If perform_fill Then
            Fill_Records_RPT(New Object() {sqlw, Absx1.chkFor("CHKVAR_ONLY").Checked})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        variances_only = False

        If parms IsNot Nothing Then
            sqlw = CStr(parms(0))
            variances_only = parms(1)
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTIREC1)
            ASCDATA1.ExecuteSQL("Insert into " & ICTIREC1 & " " & sqlICTIREC1 & "   and " & sqlw)
            If variances_only Then
                ASCMAIN1.sql = "Delete from " & ICTIREC1 _
                    & " where RECEIPT_NO in " _
                    & "(Select RECEIPT_NO from " & ICTIREC1 _
                    & " minus " _
                    & " Select Distinct RECEIPT_NO fron ICTIREC2," & ICTIREC1 & " ICTIREC1" _
                    & " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO and NVL(QTY_REC,0) <> NVL(QTY_REC,0))"
            End If
        End If

        EnforceConstraints(False)
        Fill_Records("ICTIREC1")
        Fill_Records("ICTIREC2")
        Fill_Records("ICTIREC5")

        If variances_only Then
            Dim dvw As New DataView(dst.Tables("ICTIREC2"))
            dvw.RowFilter = "ISNULL(QTY_REC,0) <> ISNULL(QTY_SHP,0)"
            Dim tbl As DataTable = dvw.ToTable
            dst.Tables("ICTIREC2").Rows.Clear()
            dst.Tables("ICTIREC2").Merge(tbl)
            dst.Tables("ICTIREC5").Clear()
        End If

        TAC.ICCMAIN1.Prepare_GL_Interface("ICIR", ICTIREC1)
        EnforceConstraints(True)
    End Sub
End Class