Imports System.Math

Public Class ICRICRR1

    Dim xRYP0_legend As String
    Dim xRYP0 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)

        grpPERIOD_RANGE.Visible = True
    End Sub

    Protected Overrides Sub Build_Workfile()

        SUBT = ""
        RWU = "N"

        Dim sqlw As String = ""

        xRYP0_legend = Absx1.cmbFor("RYP0").Value
        xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)

        SUBT = "For Period Ending " & xRYP0_legend

        Prepare_dst(True, sqlw)

        'Check_if_Empty("ICTIADJ1")

    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
        Generate_Report("ICRICRR2", "Reconciliation Units Detail Report", SUBT)
        Generate_Report("ICRICRR3", "Reconciliation Dollars Detail Report", SUBT)
    End Sub
    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE", True).Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1
        Dim SQLS As New System.Text.StringBuilder
        Dim SQLU As New System.Text.StringBuilder
        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"

        SQLS.AppendLine("Select 1 As Lorder,")
        SQLS.AppendLine(String.Format(" '{0}' As Description,", New String(" ", 50)))
        SQLS.AppendLine(" 10000 As Units,")
        SQLS.AppendLine(" 10000000000 As Dollars")
        SQLS.AppendLine(" From Dual")
        SQLS.AppendLine(" Where(Rownum < 1)")
        ASCMAIN1.sql = SQLS.ToString()
        Dim ICTICRR1 As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL(String.Format("INSERT INTO {0} VALUES ({1},'{2}',{3},{4})", ICTICRR1, 1, "Beginning Balance", 0, 0))
        ASCDATA1.ExecuteSQL(String.Format("INSERT INTO {0} VALUES ({1},'{2}',{3},{4})", ICTICRR1, 2, "Receipts", 0, 0))
        ASCDATA1.ExecuteSQL(String.Format("INSERT INTO {0} VALUES ({1},'{2}',{3},{4})", ICTICRR1, 3, "Returns", 0, 0))
        ASCDATA1.ExecuteSQL(String.Format("INSERT INTO {0} VALUES ({1},'{2}',{3},{4})", ICTICRR1, 4, "Adjustments", 0, 0))
        ASCDATA1.ExecuteSQL(String.Format("INSERT INTO {0} VALUES ({1},'{2}',{3},{4})", ICTICRR1, 5, "Sales", 0, 0))
        ASCDATA1.ExecuteSQL(String.Format("INSERT INTO {0} VALUES ({1},'{2}',{3},{4})", ICTICRR1, 6, "Ending Balance", 0, 0))
        ASCMAIN1.sql = "Select * from " & ICTICRR1
        Create_TDA(dst.Tables.Add, "ICTICRR1", "**", 0, False)

        SQLS.Clear()
        SQLS.AppendLine(" SELECT")
        SQLS.AppendLine(" S1.STYLE_CODE,")
        SQLS.AppendLine(" S1.COLOR_CODE,")
        SQLS.AppendLine(" I1.STYLE_DESC,")
        SQLS.AppendLine(" NVL(S1.WHSE_QTY_BEG,0) WHSE_QTY_BEG,")
        SQLS.AppendLine(" 1000000.99 AS WHSE_QTY_BEG_D,")
        SQLS.AppendLine(" NVL(S1.WHSE_QTY_REC,0) WHSE_QTY_REC,")
        SQLS.AppendLine(" 1000000.99 AS WHSE_QTY_REC_D,")
        SQLS.AppendLine(" NVL(S1.WHSE_QTY_RTN,0) WHSE_QTY_RTN,")
        SQLS.AppendLine(" 1000000.99 AS WHSE_QTY_RTN_D,")
        SQLS.AppendLine(" (NVL(S1.WHSE_QTY_ADJ,0) + NVL(S1.WHSE_QTY_PHY,0) + NVL(S1.WHSE_QTY_XFR,0)) WHSE_QTY_ADJ,")
        SQLS.AppendLine(" 1000000.99 AS WHSE_QTY_ADJ_D,")
        SQLS.AppendLine(" NVL(S1.WHSE_QTY_SHP,0) WHSE_QTY_SHP,")
        SQLS.AppendLine(" 1000000.99 AS WHSE_QTY_SHP_D,")
        SQLS.AppendLine(" 1000000 WHSE_QTY_END,")
        SQLS.AppendLine(" 1000000.99 AS WHSE_QTY_END_D")
        SQLS.AppendLine(" FROM ICTSTAT1 S1, ICTSTYL1 I1")
        SQLS.AppendLine(" WHERE S1.STYLE_CODE = I1.STYLE_CODE")
        SQLS.AppendLine(String.Format(" AND S1.OPS_YYYYPP = '{0}'", xRYP0))
        ASCMAIN1.sql = SQLS.ToString()
        Dim ICTICRRD As String = ASCMAIN1.Temp_Table
        SQLU.Clear()
        SQLU.AppendLine(String.Format(" UPDATE {0}", ICTICRRD))
        SQLU.AppendLine(" SET WHSE_QTY_BEG_D = 0,")
        SQLU.AppendLine(" WHSE_QTY_REC_D = 0,")
        SQLU.AppendLine(" WHSE_QTY_RTN_D = 0,")
        SQLU.AppendLine(" WHSE_QTY_ADJ_D = 0,")
        SQLU.AppendLine(" WHSE_QTY_SHP_D = 0,")
        SQLU.AppendLine(" WHSE_QTY_END = 0,")
        SQLU.AppendLine(" WHSE_QTY_END_D = 0")
        ASCDATA1.ExecuteSQL(SQLU.ToString())
        ASCMAIN1.sql = "Select * from " & ICTICRRD
        Create_TDA(dst.Tables.Add, "ICTICRRD", "**", 0, False)

        SQLS.Clear()
        SQLS.AppendLine("SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP,")
        SQLS.AppendLine(String.Format(" ICTCOST1.TRAN_TYPE, '{0}' TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST, 0.00 QTY_USED", New String(" ", 50)))
        SQLS.AppendLine(" FROM ICTCOST1")
        ASCMAIN1.sql = SQLS.ToString()
        Create_TDA(dst.Tables.Add, "ICTCOST1", "**", 0, False)

        SQLS.Clear()
        SQLU.AppendLine(" SELECT")
        SQLU.AppendLine(" STYLE_CODE, COLOR_CODE, TRAN_DATE, OPS_YYYYPP, TRAN_TYPE, TRAN_REF, TRAN_QTY,")
        SQLU.AppendLine(" TRAN_COST, 'N' CONSUMED, 0.00 CUM_QTY")
        SQLU.AppendLine(" FROM (")
        SQLU.AppendLine(" SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP,")
        SQLU.AppendLine(" ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST")
        SQLU.AppendLine(" FROM ICTCOST1)")
        SQLU.AppendLine(" WHERE ROWNUM < 0")
        Create_TDA(dst.Tables.Add, "ICTLOTSM", "**", 0, False)
        Create_TDA(dst.Tables.Add, "ICTLOTST", "**", 0, False)

        If perform_fill Then
            Fill_Records_RPT()
        End If

        For Each rowICTICRRD As DataRow In dst.Tables("ICTICRRD").Select("", "STYLE_CODE")
            ASCMAIN1.Progress(String.Format("Now Costing {0}", rowICTICRRD.Item("STYLE_CODE")))
            Dim Style_Cost As List(Of Double) = CalculateCosts(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), rowICTICRRD.Item("STYLE_CODE"), rowICTICRRD.Item("COLOR_CODE"))
            rowICTICRRD.Item("WHSE_QTY_BEG_D") = Val(rowICTICRRD.Item("WHSE_QTY_BEG")) * Val(Style_Cost(0))
            'TODO: Cost Recipts for Month by Style.
            'TODO: Cost Shipments for Month by Style.
            'TODO: Cost Ending Balance for Month by Style.
        Next

        'TODO: Sum up units and cost to fill in the summary report.

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)
        EnforceConstraints(False)
        Fill_Records("ICTICRR1")
        Fill_Records("ICTICRRD")
        EnforceConstraints(True)
    End Sub
    Function CalculateCosts(PERIOD As String, STYLE_CODE As String, COLOR_CODE As String) As List(Of Double)
        Dim ThisMonthCost As Double = 0
        Dim LastMonthCost As Double = 0
        Dim RetVal As List(Of Double)
        Dim OH_REMAINS As Long
        Dim TOT_OH As Long
        Dim LOT_REMAIN As Long
        Dim LIVE_PERIOD As String = "200607"
        Dim SQLS As New System.Text.StringBuilder

        SQLS.Clear()
        SQLS.AppendLine(" SELECT MAX(OPS_YYYYPP) FROM ICTCOST1")
        SQLS.AppendLine(" WHERE (TRAN_TYPE = 'B' OR TRAN_TYPE = 'M')")
        SQLS.AppendLine(String.Format(" AND (STYLE_CODE = '{0}' AND COLOR_CODE = '{1}')", STYLE_CODE, COLOR_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim OPS_YYYYPP As String = ASCDATA1.GetDataValue
        If OPS_YYYYPP.Length = 0 Then
            LIVE_PERIOD = OPS_YYYYPP
        End If

        If PERIOD = ASCMAIN1.CYP Then
            SQLS.Clear()
            SQLS.AppendLine(" SELECT NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND")
            SQLS.AppendLine(" FROM ICTSTAT2")
            SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", COLOR_CODE))
            SQLS.AppendLine(" AND WHSE_CODE = 'NJ'")
        Else
            SQLS.Clear()
            SQLS.AppendLine(" SELECT NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND")
            SQLS.AppendLine(" FROM ICTSTAT5")
            SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", COLOR_CODE))
            SQLS.AppendLine(" AND WHSE_CODE = 'NJ'")
            SQLS.AppendLine(String.Format(" AND OPS_YYYYPP = '{0}'", PERIOD))
        End If
        ASCMAIN1.sql = SQLS.ToString()
        TOT_OH = Val(ASCDATA1.GetDataValue)
        OH_REMAINS = TOT_OH

        If OH_REMAINS > 0 Then
            SQLS.Clear()
            SQLS.AppendLine(" SELECT")
            SQLS.AppendLine(" STYLE_CODE, COLOR_CODE, TRAN_DATE, OPS_YYYYPP, TRAN_TYPE, TRAN_REF, TRAN_QTY,")
            SQLS.AppendLine(" TRAN_COST, 'N' CONSUMED, 0.00 CUM_QTY")
            SQLS.AppendLine(" FROM (")
            SQLS.AppendLine(" SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP,")
            SQLS.AppendLine(" ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST")
            SQLS.AppendLine(" FROM ICTCOST1")
            SQLS.AppendLine(String.Format("  WHERE ICTCOST1.STYLE_CODE =  '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format(" AND ICTCOST1.COLOR_CODE = '{0}'", COLOR_CODE))
            SQLS.AppendLine(String.Format(" AND ICTCOST1.OPS_YYYYPP >= '{0}'", LIVE_PERIOD))
            SQLS.AppendLine(String.Format(" AND ICTCOST1.OPS_YYYYPP <= '{0}'", PERIOD))
            SQLS.AppendLine(" UNION")
            SQLS.AppendLine(" SELECT POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE,")
            SQLS.AppendLine(" ICTTRAN2.OPS_YYYYPP, 'R' TRAN_TYPE, POTSHIP1.PO_SHIPMENT_NO ||' - ' || POTSHIP1.PO_SHIP_VESSEL TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY,")
            SQLS.AppendLine(" POTSHIP3.PO_COST_LANDED TRAN_COST")
            SQLS.AppendLine(" FROM POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, ICTTRAN2")
            SQLS.AppendLine(" WHERE POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO")
            SQLS.AppendLine(" AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
            SQLS.AppendLine(" AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO")
            SQLS.AppendLine(" AND POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
            SQLS.AppendLine(" AND POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO")
            SQLS.AppendLine(" AND POTORDR2.PO_ORDER_NO = ICTTRAN2.PO_ORDER_NO")
            SQLS.AppendLine(" AND POTORDR2.PO_ORDER_LNO = ICTTRAN2.PO_ORDER_LNO")
            SQLS.AppendLine(String.Format(" AND POTORDR2.STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format(" AND POTORDR2.COLOR_CODE = '{0}'", COLOR_CODE))
            SQLS.AppendLine(String.Format(" AND ICTTRAN2.OPS_YYYYPP >= '{0}'", LIVE_PERIOD))
            SQLS.AppendLine(String.Format(" AND ICTTRAN2.OPS_YYYYPP <= '{0}')", PERIOD))
            SQLS.AppendLine(" ORDER BY TRAN_DATE DESC")
            ASCMAIN1.sql = SQLS.ToString()
            Fill_Records("ICTLOTSM", , True, ASCMAIN1.sql)
            dst.Tables("ICTLOTST").Rows.Clear()
            For Each rowICTLOTSM As DataRow In dst.Tables("ICTLOTSM").Select()
                OH_REMAINS = OH_REMAINS - rowICTLOTSM.Item("TRAN_QTY")
                If OH_REMAINS <= 0 Then
                    LOT_REMAIN = OH_REMAINS + rowICTLOTSM.Item("TRAN_QTY")
                    Calclot(rowICTLOTSM, LOT_REMAIN)
                    Exit For
                Else
                    LOT_REMAIN = rowICTLOTSM.Item("TRAN_QTY")
                    Calclot(rowICTLOTSM, LOT_REMAIN)
                End If
            Next
            Dim CostQty As Double
            Dim CostTotal As Double
            For Each rowICTLOTST As DataRow In dst.Tables("ICTLOTST").Select("", "TRAN_DATE DESC")
                CostQty = CostQty + Val(rowICTLOTST.Item("QTY_USED"))
                If IsDBNull(rowICTLOTST.Item("TRAN_COST")) Then
                    CostTotal = 0
                Else
                    CostTotal = CostTotal + Val(rowICTLOTST.Item("TRAN_COST") * rowICTLOTST.Item("QTY_USED"))
                End If
            Next
            If CostQty <= 0 Then
                ThisMonthCost = 0
            Else
                ThisMonthCost = CostTotal / CostQty
            End If

        End If

        RetVal = MakeCostCalc(ThisMonthCost, LastMonthCost)
        Return RetVal
    End Function

    Private Function MakeCostCalc(ByVal ThisMonthCost As Double, LastMonthCost As Double) As List(Of Double)
        Dim retVal As New List(Of Double)
        retVal.Add(ThisMonthCost)
        retVal.Add(LastMonthCost)
        Return retVal
    End Function

    Private Sub Calclot(ByVal RowToAdd As DataRow, ByVal LOT_REMAIN As Long)
        Dim rowADDNEW As DataRow = dst.Tables("ICTLOTST").NewRow
        rowADDNEW.Item("STYLE_CODE") = RowToAdd.Item("STYLE_CODE")
        rowADDNEW.Item("COLOR_CODE") = RowToAdd.Item("COLOR_CODE")
        rowADDNEW.Item("TRAN_DATE") = RowToAdd.Item("TRAN_DATE")
        rowADDNEW.Item("OPS_YYYYPP") = RowToAdd.Item("OPS_YYYYPP")
        rowADDNEW.Item("TRAN_TYPE") = RowToAdd.Item("TRAN_TYPE")
        rowADDNEW.Item("TRAN_REF") = RowToAdd.Item("TRAN_REF")
        rowADDNEW.Item("TRAN_QTY") = RowToAdd.Item("TRAN_QTY")
        rowADDNEW.Item("TRAN_COST") = RowToAdd.Item("TRAN_COST")
        rowADDNEW.Item("QTY_USED") = Val(LOT_REMAIN)
        dst.Tables("ICTLOTST").Rows.Add(rowADDNEW)
    End Sub
End Class