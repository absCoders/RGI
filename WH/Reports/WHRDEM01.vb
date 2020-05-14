Public Class WHRDEM01

#Region "Declarations"
    Dim DTE0 As Date

    Dim RptDateBasis As Boolean
    Dim SSDTEND_DATE As Date
    Dim WHSE_CODE As String

#End Region

#Region "ABS Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
        Absx1.dteFor("DTE0").Value = Today.Date

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

        Dim SS As String = SQLA("PO_SHIPMENT_NO", , True)

        If SS = "" Then
            RptDateBasis = True
            SSDTEND_DATE = Absx1.dteFor("DTE0").Value
        Else
            RptDateBasis = False
        End If


        ASCMAIN1.sql = " SELECT X1.PO_SHIPMENT_NO, X1.PO_SHIP_VESSEL,X1.CONTAINER_NO," & vbCrLf _
            & " X1.PO_SHIP_ETA, X1.STYLE_CODE, X1.COLOR_CODE," & vbCrLf _
            & " X1.STYLE_DESC, X1.PO_QTY_SHP," & vbCrLf _
            & " NVL(T2.WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND," & vbCrLf _
            & " NVL(T2.WHSE_QTY_OPEN,0) WHSE_QTY_OPEN," & vbCrLf _
            & " NVL(T2.WHSE_QTY_PICK,0) WHSE_QTY_PICK" & vbCrLf _
            & " FROM" & vbCrLf _
            & " (SELECT S1.PO_SHIPMENT_NO, NVL(S1.PO_SHIP_VESSEL,'NOT ASSIGNED') PO_SHIP_VESSEL," & vbCrLf _
            & " NVL(S2.CONTAINER_NO,'NOT ASSIGNED') CONTAINER_NO," & vbCrLf _
            & " S1.PO_SHIP_ETA, P2.STYLE_CODE, P2.COLOR_CODE, I1.STYLE_DESC," & vbCrLf _
            & " SUM(S3.PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & " FROM POTSHIP1 S1, POTSHIP2 S2, POTSHIP3 S3, POTORDR2 P2, ICTSTYL1 I1" & vbCrLf _
            & " WHERE S1.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO" & vbCrLf _
            & " AND S2.PO_SHIPMENT_NO = S3.PO_SHIPMENT_NO" & vbCrLf _
            & " AND S2.PO_SHIPMENT_LNO = S3.PO_SHIPMENT_LNO" & vbCrLf _
            & " AND S3.PO_ORDER_NO = P2.PO_ORDER_NO" & vbCrLf _
            & " AND S3.PO_ORDER_LNO = P2.PO_ORDER_LNO" & vbCrLf _
            & " AND P2.STYLE_CODE = I1.STYLE_CODE" & vbCrLf _
            & " AND S2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & " AND S1.WHSE_CODE  ='" & WHSE_CODE & "'" & vbCrLf

        If RptDateBasis Then
            ASCMAIN1.sql &= " AND S1.PO_SHIP_ETA <= '" & Format(SSDTEND_DATE, "dd-MMM-yyyy") & "'" & vbCrLf
        Else
            ASCMAIN1.sql &= " AND S1.PO_SHIPMENT_NO IN (" & SS & ")" & vbCrLf
        End If

        ASCMAIN1.sql &= "" _
            & " GROUP BY S1.PO_SHIPMENT_NO, NVL(S1.PO_SHIP_VESSEL,'NOT ASSIGNED')," & vbCrLf _
            & " NVL(S2.CONTAINER_NO,'NOT ASSIGNED')," & vbCrLf _
            & " S1.PO_SHIP_ETA, P2.STYLE_CODE, P2.COLOR_CODE, I1.STYLE_DESC) X1, ICTSTAT2 T2" & vbCrLf _
            & " WHERE X1.STYLE_CODE = T2.STYLE_CODE" & vbCrLf _
            & "  AND X1.COLOR_CODE = T2.COLOR_CODE" & vbCrLf _
            & "  AND T2.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf
        dst.Tables.Add(ASCDATA1.GetDataTable("", "WHTSTYL1", 6))

        ASCMAIN1.sql = "SELECT O2.STYLE_CODE, O2.COLOR_CODE, 'O' AS SOURCE, O1.CUST_CODE, O1.ORDR_CUST_PO, O1.ORDR_SHIP_DATE, O1.ORDR_CANCEL_DATE," & vbCrLf _
        & "  SUM(ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM(ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
        & "  FROM SOTORDR1 O1, SOTORDR2 O2" & vbCrLf _
        & "  WHERE O1.ORDR_NO = O2.ORDR_NO" & vbCrLf _
        & "  AND O1.ORDR_STATUS IN ('O','P') AND (O2.STYLE_CODE, O2.COLOR_CODE) IN" & vbCrLf _
        & "  (" & vbCrLf _
        & "  SELECT X1.STYLE_CODE, X1.COLOR_CODE" & vbCrLf _
        & "  FROM" & vbCrLf _
        & "  (SELECT P2.STYLE_CODE, P2.COLOR_CODE" & vbCrLf _
        & "  FROM POTSHIP1 S1, POTSHIP2 S2, POTSHIP3 S3, POTORDR2 P2, ICTSTYL1 I1" & vbCrLf _
        & "  WHERE S1.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO" & vbCrLf _
        & "  AND S2.PO_SHIPMENT_NO = S3.PO_SHIPMENT_NO" & vbCrLf _
        & "  AND S2.PO_SHIPMENT_LNO = S3.PO_SHIPMENT_LNO" & vbCrLf _
        & "  AND S3.PO_ORDER_NO = P2.PO_ORDER_NO" & vbCrLf _
        & "  AND S3.PO_ORDER_LNO = P2.PO_ORDER_LNO" & vbCrLf _
        & "  AND P2.STYLE_CODE = I1.STYLE_CODE" & vbCrLf _
        & "  AND S2.PO_SHIP_STATUS = 'O'" & vbCrLf _
        & "  AND S1.WHSE_CODE  ='" & WHSE_CODE & "'" & vbCrLf

        If RptDateBasis Then
            ASCMAIN1.sql &= " AND S1.PO_SHIP_ETA <= '" & Format(SSDTEND_DATE, "dd-MMM-yyyy") & "'" & vbCrLf
        Else
            ASCMAIN1.sql &= " AND S1.PO_SHIPMENT_NO IN (" & SS & ")" & vbCrLf
        End If

        ASCMAIN1.sql &= "  GROUP BY P2.STYLE_CODE, P2.COLOR_CODE) X1, ICTSTAT2 T2" & vbCrLf _
            & "  WHERE X1.STYLE_CODE = T2.STYLE_CODE" & vbCrLf _
            & "  AND X1.COLOR_CODE = T2.COLOR_CODE" & vbCrLf _
            & "  )" & vbCrLf _
            & "  GROUP BY O2.STYLE_CODE, O2.COLOR_CODE, O1.CUST_CODE, O1.ORDR_CUST_PO, O1.ORDR_SHIP_DATE, O1.ORDR_CANCEL_DATE" & vbCrLf _
            & " UNION" & vbCrLf _
            & "  SELECT O2.STYLE_CODE, O2.COLOR_CODE, 'R' AS SOURCE, O1.CUST_CODE, O1.ORDR_CUST_PO, O1.ORDR_SHIP_DATE, O1.ORDR_CANCEL_DATE," & vbCrLf _
            & "  SUM(RSRV_QTY_OPEN) ORDR_QTY_OPEN, SUM(0) ORDR_QTY_PICK" & vbCrLf _
            & "  FROM SOTRSRV1 O1, SOTRSRV2 O2" & vbCrLf _
            & "  WHERE O1.RSRV_NO = O2.RSRV_NO" & vbCrLf _
            & "  AND O1.RSRV_STATUS IN ('O','P') AND (O2.STYLE_CODE, O2.COLOR_CODE) IN" & vbCrLf _
            & "  (" & vbCrLf _
            & "  SELECT X1.STYLE_CODE, X1.COLOR_CODE" & vbCrLf _
            & "  FROM" & vbCrLf _
            & "  (SELECT P2.STYLE_CODE, P2.COLOR_CODE" & vbCrLf _
            & "  FROM POTSHIP1 S1, POTSHIP2 S2, POTSHIP3 S3, POTORDR2 P2, ICTSTYL1 I1" & vbCrLf _
            & "  WHERE S1.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO" & vbCrLf _
            & "  AND S2.PO_SHIPMENT_NO = S3.PO_SHIPMENT_NO" & vbCrLf _
            & "  AND S2.PO_SHIPMENT_LNO = S3.PO_SHIPMENT_LNO" & vbCrLf _
            & "  AND S3.PO_ORDER_NO = P2.PO_ORDER_NO" & vbCrLf _
            & "  AND S3.PO_ORDER_LNO = P2.PO_ORDER_LNO" & vbCrLf _
            & "  AND P2.STYLE_CODE = I1.STYLE_CODE" & vbCrLf _
            & "  AND S2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & "  AND S1.WHSE_CODE  ='" & WHSE_CODE & "'" & vbCrLf

        If RptDateBasis Then
            ASCMAIN1.sql &= " AND S1.PO_SHIP_ETA <= '" & Format(SSDTEND_DATE, "dd-MMM-yyyy") & "'" & vbCrLf
        Else
            ASCMAIN1.sql &= " AND S1.PO_SHIPMENT_NO IN (" & SS & ")" & vbCrLf
        End If
        ASCMAIN1.sql &= "" _
                & "  GROUP BY P2.STYLE_CODE, P2.COLOR_CODE) X1, ICTSTAT2 T2" & vbCrLf _
                & "  WHERE X1.STYLE_CODE = T2.STYLE_CODE" & vbCrLf _
                & "  AND X1.COLOR_CODE = T2.COLOR_CODE" & vbCrLf _
                & "  )" & vbCrLf _
                & "  GROUP BY O2.STYLE_CODE, O2.COLOR_CODE, O1.CUST_CODE, O1.ORDR_CUST_PO, O1.ORDR_SHIP_DATE, O1.ORDR_CANCEL_DATE" & vbCrLf

        dst.Tables.Add(ASCDATA1.GetDataTable("", "WHTORDRX", 6))


        ASCMAIN1.sql = "SELECT" & vbCrLf _
                & " S1.PO_SHIPMENT_NO," & vbCrLf _
                & " S1.PO_SHIP_VESSEL," & vbCrLf _
                & " S1.PO_SHIP_ETA," & vbCrLf _
                & " S1.PO_SHIP_REF_NO," & vbCrLf _
                & " COUNT(NVL(S2.CONTAINER_NO,'1234')) CONTAINERS," & vbCrLf _
                & " SUM(NVL(S2.PO_SHIP_CTNS,0)) CARTONS" & vbCrLf _
                & " FROM POTSHIP1 S1, POTSHIP2 S2" & vbCrLf _
                & " Where S1.PO_SHIPMENT_NO = S2.PO_SHIPMENT_NO" & vbCrLf _
                & " AND S2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                & " AND S1.WHSE_CODE  ='" & WHSE_CODE & "'" & vbCrLf

        If RptDateBasis Then
            ASCMAIN1.sql &= " AND S1.PO_SHIP_ETA <= '" & Format(SSDTEND_DATE, "dd-MMM-yyyy") & "'" & vbCrLf
        Else
            ASCMAIN1.sql &= " AND S1.PO_SHIPMENT_NO IN (" & SS & ")" & vbCrLf
        End If

        ASCMAIN1.sql &= "" _
            & " GROUP BY S1.PO_SHIPMENT_NO," & vbCrLf _
            & " S1.PO_SHIP_VESSEL," & vbCrLf _
            & " S1.PO_SHIP_ETA," & vbCrLf _
            & " S1.PO_SHIP_REF_NO" & vbCrLf _
            & " ORDER BY S1.PO_SHIP_ETA"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "WHTSHIPX", 1))

        'ASCMAIN1.sql = "Select ICTSTYL1.* from ICTSTYL1"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTYL1", 1))

        Check_if_Empty("WHTSTYL1")

    End Sub

    Public Overrides Sub Print_Report()
         
        If RptDateBasis Then
            SUBT = "Run For All Shipments Arriving On Or Before " & SSDTEND_DATE
        Else
            SUBT = "Run For Selected Shipments Shown In Summary"
        End If
        CR_params.Add("SUBT", SUBT)

        Generate_Report("WHRDEM01", "Receiving Demand Report", SUBT)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
              
        End Select
    End Sub

#End Region
 
  
End Class