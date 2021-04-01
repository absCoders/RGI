Imports Infragistics.Win.UltraWinGrid

Public Class POFBATC1

    'AUTOEXPANDING THE SPECIFIC WHSE IN DETAIL
    'EXPANDS DETAILS AUTOMATICALLY
    'CHANGE IN STATUS INSIDE PO BATCH GEN

    Dim PO_BATCH_NO As String
    Dim WHSE_CODE As String
    Dim rowPOTBATC1 As DataRow
    Dim rowICTWHSE1 As DataRow

    Dim SOTORDC1 As String
    Dim SOTSLSC1 As String
    Dim POTORDRX As String
    Dim POTBATC2 As String
    Dim sqlSOTORDC1 As String
    Dim sqlSOTSLSC1 As String
    Dim sqlPOTORDRX As String
    Dim TC As New Dictionary(Of String, Dictionary(Of String, String))
    Dim STYLE_CLASS_CODEs As New List(Of String)
    Dim generate_POs As Boolean = False

    Dim POTBATCS_expressions As New Dictionary(Of String, String)
    Dim POTBATC2_expressions As New Dictionary(Of String, String)
    Dim TTM As New UltraWinToolTip.UltraToolTipManager

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")
        Get_PARM("ICTPARM1")

        Create_Temp_Tables()

        With dst
            ASCMAIN1.sql = "Select POTBATC1.*" & vbCrLf _
            & " from POTBATC1 " & vbCrLf
            Create_TDA(.Tables.Add, "POTBATCX", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select POTBATC1.*" & vbCrLf _
            & " from POTBATC1 " & vbCrLf _
            & " where POTBATC1.PO_BATCH_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTBATC1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select :PARM1 PO_BATCH_NO" & vbCrLf _
                & ", ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYL1.STYLE_UOM" & vbCrLf _
                & ", CASE WHEN NEW_PO_COST_DATE IS NOT NULL AND :PARM2 >= NEW_PO_COST_DATE THEN NEW_PO_COST ELSE PO_COST END PO_COST" & vbCrLf _
                & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.SUB_UNIT_PACK_QTY" & vbCrLf _
                & ", ICTSTYL1.CASE_CUBE, ICTSTYL1.STYLE_PO_QTY_MIN" & vbCrLf _
                & ", ICTSTYL1.VEND_CODE, ICTSTYL1.STYLE_CLASS_CODE STYLE_CLASS_CODE, ICTSTYL1.INNER_PACK_QTY" & vbCrLf _
                & " from ICTSTYL1,ICTSTYV1" & vbCrLf _
                & " where ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE in (Select STYLE_CODE from " & POTBATC2 & ")"
            Create_TDA(.Tables.Add, "POTBATCS", "**", 0, False, "VD", 2)
            .Tables("POTBATCS").Columns("PO_BATCH_NO").MaxLength = 6

            ASCMAIN1.sql = "Select POTBATC2.*, NVL(QTY_OPEN,0)+NVL(QTY_PICK,0) QTY_OPEN_PICK" & vbCrLf _
            & " from " & POTBATC2 & " POTBATC2"
            Create_TDA(.Tables.Add, "POTBATC2", "**", 0, True)
            .Tables("POTBATC2").Columns("QTY_OPEN_PICK").DataType = GetType(System.Int64)

            Create_Relation("POTBATCS", "POTBATC2", "PO_BATCH_NO,STYLE_CODE")
            With .Tables("POTBATC2").Columns
                .Add("CUST_SOLD", GetType(System.Int32))
                .Add("QTY_SOLD", GetType(System.Int32))
                .Add("AMT_SOLD", GetType(System.Decimal))
                .Add("CASE_CUBE", GetType(System.Decimal), "PARENT(POTBATCS_POTBATC2).CASE_CUBE")
                .Add("NET_POS", GetType(System.Int32), "ISNULL(QTY_ONH,0)+ISNULL(QTY_PO,0)+ISNULL(QTY_PS,0)-ISNULL(QTY_OPEN,0)-ISNULL(QTY_PICK,0)")
                .Add("QTY_SHORT", GetType(System.Int32), "IIF(NET_POS<0,-1*NET_POS,NULL)")
                .Add("CASE_QTY", GetType(System.Int32), "IIF(ISNULL(PARENT(POTBATCS_POTBATC2).CARTON_PACK_QTY,0)>0,ISNULL(PARENT(POTBATCS_POTBATC2).CARTON_PACK_QTY,0),1)")
                .Add("CUBE", GetType(System.Decimal), "CASE_CUBE * PO_QTY_ROUNDED / CASE_QTY")
                .Add("PO_QTY_CALC", GetType(System.Int32), "PO_QTY + (CASE_QTY - PO_QTY%CASE_QTY)%CASE_QTY")
                .Add("NET_POS2", GetType(System.Int32), "NET_POS + PO_QTY_CALC")
                .Add("STYLE_CLASS_CODE", GetType(System.String), "PARENT(POTBATCS_POTBATC2).STYLE_CLASS_CODE")
                .Add("VEND_CODE", GetType(System.String), "PARENT(POTBATCS_POTBATC2).VEND_CODE")
                .Add("CARTON_PACK_QTY", GetType(System.Int32), "PARENT(POTBATCS_POTBATC2).CARTON_PACK_QTY")
                .Add("SUB_UNIT_PACK_QTY", GetType(System.Int32), "PARENT(POTBATCS_POTBATC2).SUB_UNIT_PACK_QTY")
                .Add("STYLE_PO_QTY_MIN", GetType(System.Int32), "PARENT(POTBATCS_POTBATC2).STYLE_PO_QTY_MIN")
                .Add("OK_COUNTER", GetType(System.Int32), "IIF(OK_IF_SHORT='1',1,0)")
                .Add("CHG_NET_POS", GetType(System.Int32), "ISNULL(NET_POS2,0) - ISNULL(NET_POS2_PRIOR,0)")

                .Add("QTY_SHORT_STATIC", GetType(System.Int32))
                .Add("CASE_QTY_STATIC", GetType(System.Int32))
            End With

            With .Tables("POTBATCS").Columns
                .Add("CUBE", GetType(System.Decimal), "SUM(CHILD.CUBE)")
                .Add("QTY_SHORT", GetType(System.Int32), "SUM(CHILD.QTY_SHORT)")
                .Add("PO_QTY_ROUNDED", GetType(System.Int32), "SUM(CHILD.PO_QTY_ROUNDED)")
                .Add("TOTAL_COST", GetType(System.Int32), "PO_COST * PO_QTY_ROUNDED")
                .Add("QTY_OPEN_PICK", GetType(System.Int32), "SUM(CHILD.QTY_OPEN_PICK)")
                .Add("QTY_SOLD", GetType(System.Int32), "SUM(CHILD.QTY_SOLD)")
                .Add("COLOR_COUNT", GetType(System.Int32), "COUNT(CHILD.COLOR_CODE)")
                .Add("OK_COUNT", GetType(System.Int32), "SUM(CHILD.OK_COUNTER)")
                .Add("OK_ALL", GetType(System.String), "IIF(OK_COUNT = COLOR_COUNT,'Y','')")
            End With

            ASCMAIN1.sql = "Select POTBATC3.*, APTVEND1.VEND_NAME" & vbCrLf _
                & " from POTBATC3,APTVEND1 where POTBATC3.PO_BATCH_NO = :PARM1" _
                & " and APTVEND1.VEND_CODE (+) = POTBATC3.VEND_CODE"
            Create_TDA(.Tables.Add, "POTBATC3", "**", 0, True, "V")

            Create_Relation("POTBATC3", "POTBATCS", "VEND_CODE")
            With .Tables("POTBATC3").Columns
                .Add("TOTAL_UNITS", GetType(System.Int32), "SUM(CHILD.PO_QTY_ROUNDED)")
                .Add("CUBE", GetType(System.Decimal), "SUM(CHILD.CUBE)")
                .Add("TOTAL_COST", GetType(System.Int32), "SUM(CHILD.TOTAL_COST)")
            End With

            ASCMAIN1.sql = "Select POTBATC4.*, ICTCLAS1.STYLE_CLASS_DESC STYLE_CLASS_DESC" & vbCrLf _
                & " from POTBATC4,ICTCLAS1" & vbCrLf _
                & " where POTBATC4.PO_BATCH_NO = :PARM1" & vbCrLf _
                & "   and ICTCLAS1.STYLE_CLASS_CODE = POTBATC4.STYLE_CLASS_CODE"
            Create_TDA(.Tables.Add, "POTBATC4", "**", 0, True, "V")

            ASCMAIN1.sql = "Select POTBATC5.*" & vbCrLf _
                & " from POTBATC5 where PO_BATCH_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTBATC5", "**", 0, True, "V")

            Create_Relation("POTBATC4", "POTBATC5", "PO_BATCH_NO,STYLE_CLASS_CODE")

            ASCMAIN1.sql = "Select SOTSLSC1.*, ARTCUST1.CUST_NAME" & vbCrLf _
                & " from " & SOTSLSC1 & " SOTSLSC1, ARTCUST1 " & vbCrLf _
                & " where SOTSLSC1.STYLE_CODE = :PARM1 and SOTSLSC1.COLOR_CODE = :PARM2" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE (+) = SOTSLSC1.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTSLSC1", "**", 0, False, "VV", 0)

            ASCMAIN1.sql = "Select SOTORDC1.*, ARTCUST1.CUST_NAME" & vbCrLf _
                & " from " & SOTORDC1 & " SOTORDC1, ARTCUST1 " & vbCrLf _
                & " where SOTORDC1.STYLE_CODE = :PARM1 and SOTORDC1.COLOR_CODE = :PARM2" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE (+) = SOTORDC1.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTORDC1", "**", 0, False, "VV", 0)

            ASCMAIN1.sql = "Select POTORDRX.*, APTVEND1.VEND_NAME" & vbCrLf _
                & " from " & POTORDRX & " POTORDRX, APTVEND1 " & vbCrLf _
                & " where POTORDRX.STYLE_CODE = :PARM1 and POTORDRX.COLOR_CODE = :PARM2" & vbCrLf _
                & "   and APTVEND1.VEND_CODE (+) = POTORDRX.VEND_CODE"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "VV", 0)

            ASCMAIN1.sql = "Select ICTCLAS1.STYLE_CLASS_CODE STYLE_CLASS_CODE, ICTCLAS1.STYLE_CLASS_DESC STYLE_CLASS_DESC" & vbCrLf _
                & " from ICTCLAS1" & vbCrLf
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)
            .Tables("ICTCLAS1").Columns.Add("SELECTED")

            Create_TDA(.Tables.Add, "POTORDR1", "*")
            Create_TDA(.Tables.Add, "POTORDR2", "*")

        End With

        grdPOTBATCX.DataSource = dst.Tables("POTBATCX")
        grdPOTBATC2.DataSource = dst.Tables("POTBATCS")

        grdPOTBATC2_changes.DataSource = dst.Tables("POTBATC2")
        Dim dvw As DataView = DirectCast(grdPOTBATC2_changes.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ISNULL(NET_POS2,0) <> ISNULL(NET_POS2_PRIOR,0)"

        grdPOTBATC3.DataSource = dst.Tables("POTBATC3")
        grdSOTSLSC1.DataSource = dst.Tables("SOTSLSC1")
        grdSOTORDC1.DataSource = dst.Tables("SOTORDC1")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdICTCLAS1.DataSource = dst.Tables("ICTCLAS1")
        grdPOTBATC4.DataSource = dst.Tables("POTBATC4")

        Fill_Records("ICTCLAS1")
        Sort_grdColumns(grdICTCLAS1, "STYLE_CLASS_DESC", False)


        grdPOTBATC2.DisplayLayout.UseFixedHeaders = True
        With grdPOTBATC2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "OK_ALL", "COLOR_COUNT", "STYLE_DESC", "STYLE_UOM", "STYLE_STATUS", "VEND_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdPOTBATC2.DisplayLayout.Bands(1)
            For Each COLUMN_NAME As String In New String() {"COLOR_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdPOTBATC2.DisplayLayout.Bands(0).Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False

        End With

        With grdPOTBATC2.DisplayLayout.Bands(1)
            .Columns("QTY_SHORT_STATIC").Hidden = True
            .Columns("CASE_QTY_STATIC").Hidden = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "PO_QTY" Or gcol.Key = "OK_IF_SHORT" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightYellow
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    '  gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                If New String() {"CUST_SOLD", "QTY_SOLD", "AMT_SOLD"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"PO_QTY", "PO_QTY_CALC", "PO_QTY_ROUNDED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    If gcol.Key <> "PO_QTY" Then gcol.CellAppearance.BackColor = Drawing.Color.Orange
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"CUST_OPEN", "QTY_OPEN", "QTY_PICK", "QTY_OPEN_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.CellAppearance.BackColor = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_COLOR_STATUS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.CellAppearance.BackColor = Drawing.Color.Pink
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"QTY_SHORT", "NET_POS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"QTY_ONH", "QTY_PO", "QTY_PS", "QTY_OPEN", "CUST_OPEN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With

        With grdPOTBATC3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_SEL", "PO_NOTES", "PO_MESSAGE", "PO_DATE_SHIP_BY", "PO_DATE_CANCEL"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Create_Summary(grdPOTBATCX, "PO_BATCH_NO", "Count")

        Create_Summary(grdPOTBATC2, "STYLE_CODE", "Count")
        Create_Summary(grdPOTBATC2, New String() {"CUBE", "QTY_SHORT", "PO_QTY_ROUNDED", "TOTAL_COST"}, , "POTBATCS")

        'Create_Summary(grdPOTBATC2, "COLOR_CODE", "Count", "POTBATCS_POTBATC2")
        'Create_Summary(grdPOTBATC2, New String() {"PO_QTY", "PO_QTY_CALC"}, , "POTBATCS_POTBATC2")

        Create_Summary(grdPOTBATC3, "VEND_CODE", "Count")
        Create_Summary(grdPOTBATC3, New String() {"TOTAL_UNITS", "CUBE", "TOTAL_COST"})

        Create_Summary(grdSOTORDC1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDC1, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK"})
        grdSOTORDC1.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.InGroupByRows

        Create_Summary(grdSOTSLSC1, "CUST_CODE", "Count")
        Create_Summary(grdSOTSLSC1, New String() {"QTY", "AMT"})
        grdSOTSLSC1.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.InGroupByRows

        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRX, New String() {"PO_QTY_ORD", "PO_QTY_OPN", "PO_QTY_SHP", "PO_QTY_REC"})
        grdPOTORDRX.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.InGroupByRows

        splPOTBATCA.Panel2Collapsed = True

        dteFrom.Value = Now.Date.AddMonths(-3)
        dteTo.Value = Now.Date

        grdICTCLAS1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        With grdICTCLAS1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Show_Filter(grdPOTBATC2_changes, True)
        'grdPOTBATC2_changes.Text = "Style/Colors"

        grdPOTBATC2.DisplayLayout.Override.SummaryFooterCaptionVisible = DefaultableBoolean.False
        'ASCMAIN1.Add_Value_List(grdPOTBATC2_changes, "STYLE_COLOR_STATUS", Nothing, New String() {":", "A:Active", "D:Discontinued", "N:Do Not Re-Order"}, 1)
        'ASCMAIN1.Add_Value_List(grdPOTBATC2, "STYLE_COLOR_STATUS", Nothing, New String() {":", "A:Active", "D:Discontinued", "N:Do Not Re-Order"}, 1)
        'ASCMAIN1.Add_Value_List(grdPOTBATC2, "STYLE_STATUS", Nothing, New String() {":", "A:Active", "D:Discontinued", "N:Do Not Re-Order"}, 0)
        ASCMAIN1.Add_Value_List(grdPOTBATC2_changes, "STYLE_COLOR_STATUS")
        ASCMAIN1.Add_Value_List(grdPOTBATC2, "STYLE_COLOR_STATUS", , , 1)
        ASCMAIN1.Add_Value_List(grdPOTBATC2, "STYLE_STATUS")
        ASCMAIN1.Add_Value_List(grdPOTBATCX, "BATCH_STATUS", Nothing, New String() {":", "O:Open", "P:Posted", "D:Deleted"}, 0)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 IsNot Nothing Then
                        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Whse " & Absx1.txtFor("WHSE_CODE").Text
                    End If
                End If

                STYLE_CLASS_CODEs.Clear()
                For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select("SELECTED = '1'")
                    Dim STYLE_CLASS_CODE As String = rowICTCLAS1.Item("STYLE_CLASS_CODE")
                    STYLE_CLASS_CODEs.Add(STYLE_CLASS_CODE)
                Next
                If STYLE_CLASS_CODEs.Count = 0 Then
                    EMsg &= vbCr & "No Item Classes Selected"
                Else
                    ASCMAIN1.sql = "Select Min(POTBATC1.PO_BATCH_NO) PO_BATCH_NO" _
                        & "  from POTBATC1,POTBATC4 " _
                        & " where POTBATC4.PO_BATCH_NO = POTBATC1.PO_BATCH_NO" _
                        & "   and POTBATC1.BATCH_STATUS = 'O'" _
                        & "   and POTBATC4.STYLE_CLASS_CODE in ('" & Join(STYLE_CLASS_CODEs.ToArray, "','") & "')"
                    Dim PO_BATCH_NO_in_use As String = ASCDATA1.GetDataValue
                    If PO_BATCH_NO_in_use <> "" Then
                        EMsg &= vbCr & "Batch is Already in progress with some of the Item Classes Selected (see " & PO_BATCH_NO_in_use & ")"""
                    End If
                End If

                If EMsg = "" Then
                    ' If Not ASCMAIN1.Logical_Lock("POTBATC1", "WHSE_CODE:" & WHSE_CODE) Then Exit Sub
                    For Each STYLE_CLASS_CODE As String In STYLE_CLASS_CODEs
                        'If Not ASCMAIN1.Logical_Lock("POTBATC1", "STYLE_CLASS_CODE:" & STYLE_CLASS_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("POTBATC1", "WHSE_CODE:" & WHSE_CODE & ":" & "STYLE_CLASS_CODE:" & STYLE_CLASS_CODE) Then Exit Sub
                    Next
                End If

            Case "Edit", "Load"

                WHSE_CODE = ""
                PO_BATCH_NO = ""

                If Absx1.txtFor("PO_BATCH_NO").Text = "" Then
                    EMsg &= vbCr & "No  Batch No Specified"
                Else
                    PO_BATCH_NO = Absx1.txtFor("PO_BATCH_NO").Text
                    rowPOTBATC1 = LookUp("POTBATC1", PO_BATCH_NO)
                    If rowPOTBATC1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Batch No " & PO_BATCH_NO
                    Else
                        WHSE_CODE = rowPOTBATC1.Item("WHSE_CODE")
                        If rowPOTBATC1.Item("BATCH_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowPOTBATC1.Item("BATCH_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " has been Cancelled"
                                Case "D"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " has been Deleted"
                                Case Else ' such as "F"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" Then
                    STYLE_CLASS_CODEs.Clear()
                    ASCMAIN1.sql = "Select STYLE_CLASS_CODE from POTBATC4 where PO_BATCH_NO = '" & PO_BATCH_NO & "'"
                    For Each rowPOTBATC4 As DataRow In ASCDATA1.GetDataTable.Rows
                        Dim STYLE_CLASS_CODE As String = rowPOTBATC4.Item("STYLE_CLASS_CODE")
                        STYLE_CLASS_CODEs.Add(STYLE_CLASS_CODE)
                    Next
                End If

                If EMsg = "" And EntryMode = "E" Then
                    If Not ASCMAIN1.Logical_Lock("POTBATC1", PO_BATCH_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("POTBATC1", "WHSE_CODE:" & WHSE_CODE) Then Exit Sub
                    For Each STYLE_CLASS_CODE As String In STYLE_CLASS_CODEs
                        If Not ASCMAIN1.Logical_Lock("POTBATC1", "STYLE_CLASS_CODE:" & STYLE_CLASS_CODE) Then Exit Sub
                    Next
                End If

            Case "Cancel"
                If MsgBox("Are you sure that you want to Cancel?", MsgBoxStyle.YesNo, "Verification to Cancel Changes Made to this Batch") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update", "Generate PO"
                'If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                '    Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                '    EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                'Else
                '    If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                '     > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                '        EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                '    End If
                'End If

                If grdPOTBATC2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Styles on Batch"
                Else
                    If Val(dst.Tables("POTBATC2").Compute("COUNT(STYLE_CODE)", "PO_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Styles on Batch with PO Qty >0"
                    End If
                End If


                If eItemKey = "Generate PO" Then

                    If dst.Tables("POTBATC3").Select("PO_SEL = '1'").Length = 0 Then
                        EMsg &= vbCr & "No Suppliers were selected for PO Generation"
                    End If

                    If dst.Tables("POTBATC3").Select("PO_SEL = '1' and ISNULL(TOTAL_UNITS,0) = 0").Length <> 0 Then
                        EMsg &= vbCr & "At least 1 Supplier Selected has 0 Units to Order - please De-Select"
                    End If

                    'If dst.Tables("POTBATCS").Select("ISNULL(PO_QTY_ROUNDED,0) <> 0 AND ISNULL(PO_COST,0) = 0").Length <> 0 Then
                    '    EMsg &= vbCr & "There are Styles with non-zero PO Qty which do not have a Cost"
                    'End If
                    'If dst.Tables("POTBATCS").Select("ISNULL(PO_QTY_ROUNDED,0) <> 0 AND STYLE_STATUS = 'N'").Length <> 0 Then
                    '    EMsg &= vbCr & "There are Styles marked as Do Not Re-Order with Non-Zero Order Qtys"
                    'End If
                    'If dst.Tables("POTBATC2").Select("STYLE_COLOR_STATUS = 'N' and ISNULL(PO_QTY_CALC,0) <> 0").Length <> 0 Then
                    '    EMsg &= vbCr & "There are Style/Colors marked as Do Not Re-Order with Non-Zero Order Qtys"
                    'End If

                    Dim MSG As String = "000"
                    For Each rowPOTBATC3 As DataRow In dst.Tables("POTBATC3").Select("PO_SEL = '1'")
                        If rowPOTBATC3.Item("PO_DATE_SHIP_BY") & "" = "" _
                        Or rowPOTBATC3.Item("PO_DATE_CANCEL") & "" = "" Then
                            EMsg &= vbCr & "Ship By and Cancel Dates are Mandatory"
                            Exit For
                        Else
                            If Format(rowPOTBATC3.Item("PO_DATE_SHIP_BY"), "yyyyMMdd") _
                             > Format(rowPOTBATC3.Item("PO_DATE_CANCEL"), "yyyyMMdd") Then
                                EMsg &= vbCr & "Ship By Date may NOT be later than Cancel Date"
                                Exit For
                            End If
                        End If


                        For Each rowPOTBATCS As DataRow In rowPOTBATC3.GetChildRows("POTBATC3_POTBATCS")
                            If Val(rowPOTBATCS.Item("PO_QTY_ROUNDED") & "") <> 0 Then
                                If Mid(MSG, 1, 1) = "0" And Val(rowPOTBATCS.Item("PO_COST") & "") = 0 Then
                                    EMsg &= vbCr & "There are Styles with non-zero PO Qty which do not have a Cost"
                                    Mid(MSG, 1, 1) = "1"
                                End If
                                If Mid(MSG, 2, 1) = "0" And rowPOTBATCS.Item("STYLE_STATUS") & "" = "N" Then
                                    EMsg &= vbCr & "There are Styles marked as Do Not Re-Order with Non-Zero Order Qtys"
                                    Mid(MSG, 2, 1) = "1"
                                End If
                                If Mid(MSG, 3, 1) = "0" Then
                                    For Each rowPOTBATC2 As DataRow In rowPOTBATCS.GetChildRows("POTBATCS_POTBATC2")
                                        If rowPOTBATC2.Item("STYLE_COLOR_STATUS") & "" = "N" And Val(rowPOTBATC2.Item("PO_QTY_CALC") & "") <> 0 Then
                                            EMsg &= vbCr & "There are Style/Colors marked as Do Not Re-Order with Non-Zero Order Qtys"
                                            Mid(MSG, 3, 1) = "1"
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    Next

                    If EMsg = "" Then
                        Dim TOTAL_COST As Decimal = Val(dst.Tables("POTBATC3").Compute("SUM(TOTAL_COST)", "PO_SEL = '1'") & "")
                        Dim PO_COUNT As Int64 = Val(dst.Tables("POTBATC3").Select("PO_SEL = '1'").Length)
                        If MsgBox("You have Selected the option to Generate POs" _
                                  & vbCrLf & vbCrLf & CStr(PO_COUNT) & " POs will be Generated, totaling " & Format(TOTAL_COST, "$#,##0.00") _
                                  & vbCrLf & vbCrLf & "Once you Generate these POs you may NOT change them using this screen; You may change them only by using PO Entry." _
                                  & vbCrLf & vbCrLf & "OK to continue to Generate POs?" _
                                  & "",
                                  MsgBoxStyle.YesNo,
                                  "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Delete"
                If EMsg = "" Then
                    If MsgBox("Do you want to Mark this Batch as Deleted",
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If



        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update", "Generate PO"
                generate_POs = (eItemKey = "Generate PO")
                Update_Record()
                Mode_Settings(False)

            Case "Save"
                generate_POs = False
                Update_Record(True)
            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Scan for New Styles"
                Scan_for_New_Styles()
        End Select

    End Sub


    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "L" And ScreenMode) Then
                        If rowPOTBATC1.Item("BATCH_STATUS") & "" = "O" Then
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                        Else
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                        End If
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Save").Settings.Enabled = iScreenMode
                    .Items("Generate PO").Visible = (EntryMode = "N" Or EntryMode = "E")

                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    .Items("Print").Visible = ScreenMode
                    If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                        .Items("Update").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                    Else
                        .Items("Update").Visible = False
                    End If

                    .Items("Save").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                    .Items("Delete").Visible = (EntryMode = "E")
                    .Items("Cancel").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                    .Items("Scan for New Styles").Visible = (EntryMode = "E")
                End With
                .Groups("Sales History").Visible = ScreenMode
                .Groups("Style Filters").Visible = ScreenMode
                .Groups("Pre-Load Options").Visible = Not ScreenMode
                .Groups("Style Class Selection").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        If ScreenMode Then
            Set_Read_Only_for_ctl(txtPO_BATCH_DESC, False)
            tabMain.Tabs("Changes in Excess").Visible = (EntryMode = "E" Or EntryMode = "N")
        End If

        lblStatusQtys.Visible = Not (EntryMode = "N" Or EntryMode = "E")

        optASN.Visible = Not ScreenMode

        lblStatus.Visible = ScreenMode
        grdPOTBATCX.Visible = Not ScreenMode
        splPOTBATCA.Visible = ScreenMode

        lblPO_BATCH_DESC.Visible = ScreenMode
        txtPO_BATCH_DESC.Visible = ScreenMode

        If ScreenMode Then

            With grdPOTBATC2.DisplayLayout.Bands(1)
                If (EntryMode = "E" Or EntryMode = "N") Then
                    .Columns("PO_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    'If EntryMode = "E" Then
                    '    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                    'Else
                    '    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                    'End If
                Else
                    .Columns("PO_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    '.Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End With

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTBATC2, grdPOTBATC3, grdPOTBATC4}
                With grd.DisplayLayout.Override
                    If EntryMode = "L" Then
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.True
                        .AllowDelete = DefaultableBoolean.False
                        If grd.Name = "grdPOTBATC3" Then
                            grdPOTBATC3.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                        End If
                    End If

                    If grd.Name = "grdPOTBATC4" Then
                        With grd.DisplayLayout.Bands(1).Override
                            If EntryMode = "L" Then
                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                .AllowDelete = DefaultableBoolean.False
                            Else
                                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                                .AllowDelete = DefaultableBoolean.True
                            End If
                        End With
                    End If
                End With
            Next
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTBATC1", "POTBATC2", "POTBATC3", "POTBATC4", "POTBATC5", "POTBATCS",
                 "SOTORDC1", "SOTSLSC1", "POTORDRX", "POTORDR1", "POTORDR2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select("")
            rowICTCLAS1.Item("SELECTED") = "0"
        Next
        Load_POTBATCX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        grdPOTBATC2.BeginUpdate()
        grdPOTBATC3.BeginUpdate()

        TC.Clear()
        For Each TABLE_NAME As String In New String() {"POTBATCS", "POTBATC1", "POTBATC3"}
            With dst.Tables(TABLE_NAME)
                Dim CE As New Dictionary(Of String, String)
                For c As Integer = .Columns.Count - 1 To 0 Step -1
                    If .Columns(c).Expression <> "" Then
                        CE.Add(.Columns(c).ColumnName, .Columns(c).Expression)
                        .Columns(c).Expression = ""
                    End If
                Next
                TC.Add(TABLE_NAME, CE)
            End With
        Next

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ASCDATA1.ExecuteSQL("Truncate Table " & POTBATC2)

        If EntryMode = "N" Then
            PO_BATCH_NO = ASCMAIN1.Next_Control_No("POTBATC1.PO_BATCH_NO")

            rowPOTBATC1 = dst.Tables("POTBATC1").NewRow
            With rowPOTBATC1
                .Item("PO_BATCH_NO") = PO_BATCH_NO
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("BATCH_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("POTBATC1").Rows.Add(rowPOTBATC1)

            'ASCMAIN1.sql = "Select '" & PO_BATCH_NO & "' PO_BATCH_NO" _
            '    & ", ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
            '    & " from ICTSTYL1,ICTSTYC1,ICTCOLR1" & vbCrLf _
            '    & " where ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            '    & "   and ICTSTYL1.STYLE_CLASS_CODE in ('" & Join(STYLE_CLASS_CODEs.ToArray, "','") & "')" & vbCrLf _
            '    & IIf(chkCustomerStylesOnly.Checked, _
            '          " and ICTSTYL1.CUST_CODE is Not Null", _
            '          " and ICTSTYL1.CUST_CODE is Null") & vbCrLf _
            '    & "   and ICTCOLR1.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE"

            ASCMAIN1.sql = "Select '" & PO_BATCH_NO & "' PO_BATCH_NO" _
                & ", ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
                & " from ICTSTYL1,ICTSTYC1,ICTCOLR1" & vbCrLf _
                & " where ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CLASS_CODE in ('" & Join(STYLE_CLASS_CODEs.ToArray, "','") & "')" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE"
            Select Case optASN.Value
                Case "S"
                    ASCMAIN1.sql += "   and ICTSTYL1.CUST_CODE is Null"
                Case "N"
                    ASCMAIN1.sql += "   and ICTSTYL1.CUST_CODE is Not Null"
            End Select
            If chkHideDN.Checked = True Then
                ASCMAIN1.sql += " AND NVL(ICTSTYC1.STYLE_COLOR_STATUS,'A') = 'A'"
            End If
            ASCDATA1.ExecuteSQL("Insert into " & POTBATC2 _
                                & " (PO_BATCH_NO,STYLE_CODE,COLOR_CODE,COLOR_DESC,STYLE_COLOR_STATUS) " _
                                & ASCMAIN1.sql)

            Dim PO_BATCH_DESC As String = ""

            For Each STYLE_CLASS_CODE As String In STYLE_CLASS_CODEs
                Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)
                PO_BATCH_DESC &= "," & rowICTCLAS1.Item("STYLE_CLASS_DESC")
                Dim rowPOTBATC4 As DataRow = dst.Tables("POTBATC4").NewRow
                rowPOTBATC4.Item("PO_BATCH_NO") = PO_BATCH_NO
                rowPOTBATC4.Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                rowPOTBATC4.Item("STYLE_CLASS_DESC") = rowICTCLAS1.Item("STYLE_CLASS_DESC")
                dst.Tables("POTBATC4").Rows.Add(rowPOTBATC4)

                ASCMAIN1.sql = "Select * from ICTCLAS2 where STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'"
                For Each rowICTCLAS2 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim rowPOTBATC5 As DataRow = dst.Tables("POTBATC5").NewRow
                    rowPOTBATC5.Item("PO_BATCH_NO") = PO_BATCH_NO
                    rowPOTBATC5.Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                    rowPOTBATC5.Item("CUST_COUNT") = rowICTCLAS2.Item("CUST_COUNT")
                    rowPOTBATC5.Item("PCT_INCREASE") = rowICTCLAS2.Item("PCT_INCREASE")
                    dst.Tables("POTBATC5").Rows.Add(rowPOTBATC5)
                Next
            Next

            rowPOTBATC1.Item("PO_BATCH_DESC") = Mid(PO_BATCH_DESC, 2)
        Else
            ASCMAIN1.sql = "Select POTBATC2.*, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_STATUS" _
                & " from POTBATC2,ICTCOLR1,ICTSTYC1" _
                & " where ICTCOLR1.COLOR_CODE (+) = POTBATC2.COLOR_CODE" _
                & "   and POTBATC2.PO_BATCH_NO = '" & PO_BATCH_NO & "'" _
                & "   and ICTSTYC1.STYLE_CODE = POTBATC2.STYLE_CODE" _
                & "   and ICTSTYC1.COLOR_CODE = POTBATC2.COLOR_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & POTBATC2 & " " & ASCMAIN1.sql)

            rowPOTBATC1 = Fill_Record("POTBATC1", PO_BATCH_NO)

            Fill_Records("POTBATC4", PO_BATCH_NO)
            Fill_Records("POTBATC5", PO_BATCH_NO)
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & SOTORDC1)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDC1 & " " & sqlSOTORDC1)

        ASCDATA1.ExecuteSQL("Truncate Table " & SOTSLSC1)
        ASCDATA1.ExecuteSQL("Insert into " & SOTSLSC1 & " " & sqlSOTSLSC1)

        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDRX)
        ASCDATA1.ExecuteSQL("Insert into " & POTORDRX & " " & sqlPOTORDRX)
        ASCDATA1.ExecuteSQL("Update " & POTORDRX & " Set PO_QTY_SHP = 0")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO, POTSHIP3.PO_QTY_SHP" & vbCrLf _
            & "   from POTSHIP3,POTSHIP2,POTORDR2," & POTBATC2 & " POTBATC2" & vbCrLf _
            & "   where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "     and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "     and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "     and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & "     and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & "     and POTORDR2.STYLE_CODE = POTBATC2.STYLE_CODE" & vbCrLf _
            & "     and POTORDR2.COLOR_CODE = POTBATC2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & POTORDRX & vbCrLf _
            & "    Set PO_QTY_SHP = R1.PO_QTY_SHP" & vbCrLf _
            & "     where PO_ORDER_NO = R1.PO_ORDER_NO" & vbCrLf _
            & "       and PO_ORDER_LNO = R1.PO_ORDER_LNO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"

        ASCDATA1.ExecuteSQL()

        If EntryMode = "N" Or EntryMode = "E" Then

            ASCMAIN1.sql = "Update " & POTBATC2 & " Set QTY_ONH = 0, QTY_OPEN = 0, QTY_PICK = 0, QTY_PO = 0"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select * from ICTSTAT2" & vbCrLf _
                & "   where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" & vbCrLf _
                & "     and (STYLE_CODE, COLOR_CODE) in" & vbCrLf _
                & "   (Select STYLE_CODE, COLOR_CODE from " & POTBATC2 & ");" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & POTBATC2 & " Set QTY_ONH = NVL(R1.WHSE_QTY_ON_HAND,0)" & vbCrLf _
                & "    , QTY_OPEN = NVL(R1.WHSE_QTY_OPEN,0)" & vbCrLf _
                & "    , QTY_PICK =  + NVL(R1.WHSE_QTY_PICK,0)" & vbCrLf _
                & "    , QTY_PO = NVL(R1.WHSE_QTY_ON_ORDER,0)" & vbCrLf _
                & "    , QTY_PS = NVL(R1.WHSE_QTY_TRAN,0)" & vbCrLf _
                & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update " & POTBATC2 & " Set CUST_OPEN = 0"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is" _
                & "  Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, Count (Distinct SOTORDR1.CUST_CODE) CUST_OPEN" _
                & "   from SOTORDR2,SOTORDR1 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                & "   and SOTORDR1.ORDR_STATUS >= 'O' and SOTORDR1.ORDR_STATUS <= 'P'" _
                & "   and SOTORDR2.ORDR_STATUS >= 'O' and SOTORDR2.ORDR_STATUS <= 'P'" _
                & "   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) <> 0 or NVL(SOTORDR2.ORDR_QTY_PICK,0) <> 0)" _
                & "   and SOTORDR1.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" _
                & "   and (SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE) in" _
                & "   (Select STYLE_CODE, COLOR_CODE from " & POTBATC2 & ")" _
                & "   group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Update " & POTBATC2 & " Set CUST_OPEN = R1.CUST_OPEN" _
                & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        Fill_Records("POTBATC3", PO_BATCH_NO)

        Fill_Records("POTBATCS", New Object() {PO_BATCH_NO, CDate(rowPOTBATC1.Item("INIT_DATE")).Date})
        For Each rowPOTBATCS As DataRow In dst.Tables("POTBATCS").Select("ISNULL(VEND_CODE,'') = ''")
            rowPOTBATCS.Item("VEND_CODE") = "."
        Next

        For Each rowVEND_CODE As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTBATCS"), New String() {"VEND_CODE"}).Rows
            Dim VEND_CODE As String = rowVEND_CODE.Item("VEND_CODE")
            Dim rowPOTBATC3 As DataRow = dst.Tables("POTBATC3").Rows.Find(New String() {PO_BATCH_NO, VEND_CODE})
            If rowPOTBATC3 Is Nothing Then
                rowPOTBATC3 = dst.Tables("POTBATC3").NewRow
                rowPOTBATC3.Item("PO_BATCH_NO") = PO_BATCH_NO
                rowPOTBATC3.Item("VEND_CODE") = VEND_CODE
                'rowPOTBATC3.Item("") = ""
                dst.Tables("POTBATC3").Rows.Add(rowPOTBATC3)
            End If
        Next
        Sort_grdColumns(grdPOTBATC3, "VEND_CODE")

        WHSE_CODE = rowPOTBATC1.Item("WHSE_CODE")
        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)

        Manage_Expressions("Remove")
        Fill_Records("POTBATC2")
        Manage_Expressions("Restore")

        Sort_grdColumns(grdPOTBATC2, "STYLE_CODE")
        Sort_grdColumns(grdPOTBATC2, "COLOR_CODE", , 1)

        Setup_SOTSLSC1()
        Setup_SOTORDC1()
        Setup_POTORDRX()

        If EntryMode = "N" Then
            lblStatus.Text = "New Batch"
        Else
            Select Case rowPOTBATC1.Item("BATCH_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Closed"
                Case "D"
                    lblStatus.Text = "Deleted"
            End Select
        End If

        Setup_POTBATC2()

        'Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        For Each TABLE_NAME As String In New String() {"POTBATCS", "POTBATC1", "POTBATC3"}
            With dst.Tables(TABLE_NAME)
                Dim CE As Dictionary(Of String, String) = TC(TABLE_NAME)
                For c As Integer = 0 To .Columns.Count - 1
                    If CE.ContainsKey(.Columns(c).ColumnName) Then
                        .Columns(c).Expression = CE(.Columns(c).ColumnName)
                    End If
                Next
            End With
        Next
        grdPOTBATC2.EndUpdate()
        grdPOTBATC3.EndUpdate()

        grdPOTBATC2.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

    End Sub

    Sub Delete_Record()
        BeginTrans()
        ASCDATA1.ExecuteSQL("Update POTBATC1 Set BATCH_STATUS = 'D' where PO_BATCH_NO = '" & PO_BATCH_NO & "'")
        'For Each TABLE_NAME As String In New String() {"POTBATC1", "POTBATC2", "POTBATC3", "POTBATC4", "POTBATC5"}
        '    ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & " where PO_BATCH_NO = '" & PO_BATCH_NO & "'")
        'Next
        CommitTrans("Delete Complete")
    End Sub

    Sub Update_Record(Optional ByVal SaveOnly As Boolean = False)

        BeginTrans()

        For Each rowPOTBATC2 As DataRow In dst.Tables("POTBATC2").Select("")
            rowPOTBATC2.Item("NET_POS2_PRIOR") = rowPOTBATC2.Item("NET_POS2")
            rowPOTBATC2.Item("VEND_CODE_PRIOR") = rowPOTBATC2.GetParentRow("POTBATCS_POTBATC2").Item("VEND_CODE")
        Next

        Dim PO_ORDER_NOs As New List(Of String)
        If generate_POs Then
            For Each rowPOTBATC3 As DataRow In dst.Tables("POTBATC3").Select("TOTAL_UNITS <> 0 and PO_SEL = '1'")
                Dim PO_ORDER_NO As String = ASCMAIN1.Next_Control_No("POTORDR1.PO_ORDER_NO")
                PO_ORDER_NOs.Add(PO_ORDER_NO)
                Dim VEND_CODE As String = rowPOTBATC3.Item("VEND_CODE")

                rowPOTBATC3.Item("PO_ORDER_NO") = PO_ORDER_NO
                rowPOTBATC3.Item("PO_DATE_ETA") = CDate(rowPOTBATC3.Item("PO_DATE_CANCEL")).AddDays(Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETD_TO_ETA") & ""))

                Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").NewRow
                With rowPOTORDR1
                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("VEND_CODE") = VEND_CODE
                    .Item("VEND_NAME") = rowPOTBATC3.Item("VEND_NAME")
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("PO_DATE_ORDERED") = DATETIME_STAMP.Date
                    .Item("PO_REFERENCE") = Absx1.txtFor("PO_BATCH_NO").Text
                    .Item("WHSE_CODE") = Absx1.txtFor("WHSE_CODE").Text
                    .Item("PO_STATUS") = "O"
                    .Item("PO_DATE_SHIP_BY") = rowPOTBATC3.Item("PO_DATE_SHIP_BY")
                    .Item("PO_DATE_CANCEL") = rowPOTBATC3.Item("PO_DATE_CANCEL")
                    .Item("PO_DATE_ETA") = rowPOTBATC3.Item("PO_DATE_ETA")
                    ' PROB NEED A NEW COLUMN
                    .Item("PO_NOTES") = rowPOTBATC3.Item("PO_NOTES")
                    .Item("PO_MESSAGE") = rowPOTBATC3.Item("PO_MESSAGE")
                    .Item("PO_BATCH_NO") = Absx1.txtFor("PO_BATCH_NO").Text
                    '.Item("FACTORY_CODE") = "?"
                    .Item("PO_XMIT_IND") = "0"
                    .Item("FOB_CMT") = "F"

                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
                    '  .Item("VEND_NAME") = rowAPTVEND1.Item("VEND_NAME")
                    .Item("TERM_CODE") = rowAPTVEND1.Item("TERM_CODE")
                    .Item("PORT_CODE_ORIG") = rowAPTVEND1.Item("PORT_CODE")
                    ' rowPOTORDR1.Item("PORT_CODE_DEST") = rowICTWHSE1.Item("PORT_CODE")

                    .Item("COST_CODE") = rowAPTVEND1.Item("COST_CODE")
                    .Item("PO_FOB_DESC") = rowAPTVEND1.Item("VEND_PURCH_FOB_DESC")
                    .Item("PO_SHIP_VIA") = rowAPTVEND1.Item("VEND_PURCH_SHIP_VIA")
                End With
                dst.Tables("POTORDR1").Rows.Add(rowPOTORDR1)

                Dim PO_ORDER_LNO As Int64 = 0
                For Each rowPOTBATCS As DataRow In dst.Tables("POTBATCS").Select("VEND_CODE = '" & VEND_CODE & "'", "STYLE_CODE")
                    For Each rowPOTBATC2 As DataRow In rowPOTBATCS.GetChildRows("POTBATCS_POTBATC2")
                        Dim PO_QTY_ORD As Int64 = Val(rowPOTBATC2.Item("PO_QTY_ROUNDED") & "")
                        If PO_QTY_ORD <> 0 Then
                            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").NewRow
                            With rowPOTORDR2
                                .Item("PO_ORDER_NO") = PO_ORDER_NO
                                PO_ORDER_LNO += 1
                                .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                                .Item("STYLE_CODE") = rowPOTBATC2.Item("STYLE_CODE")
                                .Item("COLOR_CODE") = rowPOTBATC2.Item("COLOR_CODE")
                                .Item("PO_QTY_ORD") = PO_QTY_ORD
                                .Item("PO_QTY_OPN") = PO_QTY_ORD
                                .Item("PO_COST") = rowPOTBATCS.Item("PO_COST")

                                .Item("PO_DATE_SHIP_BY") = rowPOTBATC3.Item("PO_DATE_SHIP_BY")
                                .Item("PO_DATE_ETA") = rowPOTBATC3.Item("PO_DATE_ETA")
                                .Item("PO_ORIG_DATE_SHIP_BY") = rowPOTBATC3.Item("PO_DATE_SHIP_BY")
                                .Item("PO_ORIG_DATE_ETA") = rowPOTBATC3.Item("PO_DATE_ETA")
                                .Item("PO_STATUS") = "O"
                                .Item("PO_QTY_UOM") = 1 ' rowPOTBATCS.Item("STYLE_UOM")
                                .Item("PO_COST_VCOST") = .Item("PO_COST")
                                .Item("STYLE_NOTES") = ""
                                .Item("SUB_UNIT_PACK_QTY") = rowPOTBATCS.Item("SUB_UNIT_PACK_QTY")
                                .Item("PO_COST_VCOST_DZ") = Val(.Item("PO_COST") & "") * 12
                                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                .Item("INIT_DATE") = DATETIME_STAMP
                                .Item("CARTON_PACK_QTY") = rowPOTBATCS.Item("CARTON_PACK_QTY")
                                .Item("INNER_PACK_QTY") = rowPOTBATCS.Item("INNER_PACK_QTY")
                            End With
                            dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)
                        End If
                    Next
                Next
            Next
            Update_Record_TDA("POTORDR1")
            Update_Record_TDA("POTORDR2")

            For Each PO_ORDER_NO As String In PO_ORDER_NOs
                ASCMAIN1.sql = "" _
                        & "Begin" & vbCrLf _
                        & " Declare Cursor C1 is Select * from POTORDR1 where PO_ORDER_NO = '" & PO_ORDER_NO & "' for Update;" & vbCrLf _
                        & " Begin " & vbCrLf _
                        & "  For R1 in C1 Loop" & vbCrLf _
                        & "   Begin" & vbCrLf _
                        & "    Declare " & vbCrLf _
                        & "     Cursor C2 is Select * from POTORDR2 where PO_ORDER_NO = R1.PO_ORDER_NO for Update;" & vbCrLf _
                        & "     QTY Number(8,0);" & vbCrLf _
                        & "    Begin" & vbCrLf _
                        & "     For R2 in C2 Loop" & vbCrLf _
                        & "      QTY := 1 * NVL(R2.PO_QTY_OPN,0);" & vbCrLf _
                        & "      Update ICTSTAT2 Set WHSE_QTY_ON_ORDER = NVL(WHSE_QTY_ON_ORDER,0) + QTY" & vbCrLf _
                        & "       where STYLE_CODE = R2.STYLE_CODE and COLOR_CODE = R2.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                        & "      If SQL%NOTFOUND then" & vbCrLf _
                        & "       Insert into ICTSTAT2 (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_ON_ORDER)" & vbCrLf _
                        & "        values (R2.STYLE_CODE, R2.COLOR_CODE, R1.WHSE_CODE, QTY);" & vbCrLf _
                        & "      End If;" _
                        & "     End Loop;" & vbCrLf _
                        & "    End;" & vbCrLf _
                        & "   End;" & vbCrLf _
                        & "  End Loop;" & vbCrLf _
                        & " End;" & vbCrLf _
                        & "End;"
                ASCDATA1.ExecuteSQL()
            Next

            rowPOTBATC1.Item("BATCH_STATUS") = "P"
        End If

        INIT_LAST("POTBATC1", False)

        For Each TABLE_NAME As String In New String() {"POTBATC1", "POTBATC2", "POTBATC3", "POTBATC4", "POTBATC5"}
            Update_Record_TDA(TABLE_NAME, "PO_BATCH_NO = '" & PO_BATCH_NO & "'")
        Next

        If generate_POs Then
            If optGeneratePOs.Value = "O" Then
                Dim PO_BATCH_NO_NEW As String = ASCMAIN1.Next_Control_No("POTBATC1.PO_BATCH_NO")
                For Each TABLE_NAME As String In New String() {"POTBATC1", "POTBATC2", "POTBATC3", "POTBATC4", "POTBATC5"}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " Set PO_BATCH_NO = '" & PO_BATCH_NO_NEW & "' where PO_BATCH_NO = '" & PO_BATCH_NO & "'")
                    Update_Record_TDA(TABLE_NAME, "PO_BATCH_NO = '" & PO_BATCH_NO & "'")
                Next
                ASCDATA1.ExecuteSQL("Update POTBATC1 Set BATCH_STATUS = 'O' where PO_BATCH_NO = '" & PO_BATCH_NO_NEW & "'")
                ASCDATA1.ExecuteSQL("Update POTBATC2 Set PO_QTY = 0, PO_QTY_ROUNDED = 0 where PO_BATCH_NO = '" & PO_BATCH_NO_NEW & "' and STYLE_CODE in (Select POTORDR2.STYLE_CODE from POTORDR2,POTORDR1 where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO and POTORDR1.PO_BATCH_NO = '" & PO_BATCH_NO & "')")
                ASCDATA1.ExecuteSQL("Update POTBATC3 Set PO_ORDER_NO = NULL where PO_BATCH_NO = '" & PO_BATCH_NO_NEW & "'")

                ASCDATA1.ExecuteSQL("Delete from POTBATC3 where PO_BATCH_NO = '" & PO_BATCH_NO & "' and PO_ORDER_NO is Null")
                ASCDATA1.ExecuteSQL("Delete from POTBATC2 where PO_BATCH_NO = '" & PO_BATCH_NO & "' and VEND_CODE_PRIOR Not in (Select VEND_CODE from POTBATC3 where PO_BATCH_NO = '" & PO_BATCH_NO & "')")
            End If
        End If

        If SaveOnly Then
            CommitTrans("")
        Else
            CommitTrans("Update Complete" & IIf(generate_POs,
                                    vbCrLf & vbCrLf & " and " & CStr(PO_ORDER_NOs.Count) & " POs were Generated",
                                    ""))
        End If
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("PORWREC2")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_BATCH_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTBATCX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdPOTBATC2, "SSSBBSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry", "Style Master File", "Show Details", "Expand All", "Collapse All", "Add by %", "Reset Add by %", "Set PO QTY = Rnd", "Discontinue Color", "DNR Color", "Discontinue Item", "DNR Item", "Danny Order")
        Load_Popup_Menu(grdPOTBATC3, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "Vendor Inquiry", "Change All")
        Load_Popup_Menu(grdPOTBATC2_changes, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Master File")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdPOTBATC3"
                    tlb_btn = DirectCast(tlb_pop.Tools("Change All"), UltraWinToolbars.ButtonTool)
                    If grd.ActiveCell Is Nothing OrElse (grd.ActiveCell.Value & "" = "" _
                    Or (EntryMode <> "N" And EntryMode <> "E") _
                    Or (grd.ActiveCell.Column.Key <> "PO_DATE_SHIP_BY" And grd.ActiveCell.Column.Key <> "PO_DATE_CANCEL")) Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.SharedProps.Caption = "Change All " & grd.ActiveCell.Column.Header.Caption & "s to " & Format(grd.ActiveCell.Value, "MM/dd/yyyy")
                        tlb_btn.Tag = grd.ActiveCell.Column.Key
                    End If
                Case "grdPOTBATC2"
                    tlb_btn = DirectCast(tlb_pop.Tools("Discontinue Item"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("DNR Item"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("Discontinue Color"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("DNR Color"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                        If grdPOTBATC2.Selected.Rows.Count = 1 Then
                            If grdPOTBATC2.Selected.Rows(0).Band.Index = 0 Then
                                Dim STYLE_STATUS As String = grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value.ToString
                                If STYLE_STATUS <> "D" Then
                                    e.Tool.ToolbarsManager.Tools("Discontinue Item").SharedProps.Visible = True
                                Else
                                    e.Tool.ToolbarsManager.Tools("Discontinue Item").SharedProps.Visible = False
                                End If
                                If STYLE_STATUS <> "N" Then
                                    e.Tool.ToolbarsManager.Tools("DNR Item").SharedProps.Visible = True
                                Else
                                    e.Tool.ToolbarsManager.Tools("DNR Item").SharedProps.Visible = False
                                End If
                            Else
                                Dim STYLE_COLOR_STATUS As String = grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Value.ToString
                                If STYLE_COLOR_STATUS <> "D" Then
                                    e.Tool.ToolbarsManager.Tools("Discontinue Color").SharedProps.Visible = True
                                Else
                                    e.Tool.ToolbarsManager.Tools("Discontinue Color").SharedProps.Visible = False
                                End If
                                If STYLE_COLOR_STATUS <> "N" Then
                                    e.Tool.ToolbarsManager.Tools("DNR Color").SharedProps.Visible = True
                                Else
                                    e.Tool.ToolbarsManager.Tools("DNR Color").SharedProps.Visible = False
                                End If
                            End If
                        Else
                            e.Tool.ToolbarsManager.Tools("Discontinue Color").SharedProps.Visible = False
                            e.Tool.ToolbarsManager.Tools("DNR Color").SharedProps.Visible = False
                        End If
                    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

                'Case "Style Master File"
                '    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                '    If rowICTSTYL1 IsNot Nothing Then
                '        Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                '    End If

            Case "Style Master File"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If

            Case "Show Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splPOTBATCA.Panel2Collapsed = Not tlb_sbt.Checked

            Case "Show Calculations"

            Case "Expand All"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Expanding all Nodes")
                ' grdPOTBATC2.Visible = False
                grdPOTBATC2.Rows.ExpandAll(True)
                ' grdPOTBATC2.Visible = True
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Collapse All"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Collapsing all Nodes")
                ' grdPOTBATC2.Visible = False
                grdPOTBATC2.Rows.CollapseAll(True)
                ' grdPOTBATC2.Visible = True
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Change All"
                tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)

                Dim COLUMN_NAME As String = tlb_btn.Tag
                Dim DATE_TO_USE As Date = grdPOTBATC3.ActiveRow.Cells(COLUMN_NAME).Value

                For Each rowPOTBATC3 As DataRow In dst.Tables("POTBATC3").Select("")
                    rowPOTBATC3.Item(COLUMN_NAME) = DATE_TO_USE
                Next

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value & ""
                If PO_ORDER_NO = "" Then
                    MsgBox("PO has not been Geneated Yet", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "Reset Add by %"
                For Each rowPOTBATC2 As DataRow In dst.Tables("POTBATC2").Select("QTY_SHORT > 0 and CUST_OPEN >= 0")
                    Dim CUST_OPEN As Int64 = Val(rowPOTBATC2.Item("CUST_OPEN") & "")
                    rowPOTBATC2.Item("PO_QTY") = Val(rowPOTBATC2.Item("QTY_SHORT") & "")
                    rowPOTBATC2.Item("PO_QTY_ROUNDED") = rowPOTBATC2.Item("PO_QTY_CALC")
                Next

            Case "Add by %"
                For Each rowPOTBATC4 As DataRow In dst.Tables("POTBATC4").Select("")
                    Dim STYLE_CLASS_CODE As String = rowPOTBATC4.Item("STYLE_CLASS_CODE")
                    Dim PCT_INCREASE_last As Int64 = -1
                    For Each rowPOTBATC5 As DataRow In dst.Tables("POTBATC5").Select _
                           ("STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'", "CUST_COUNT")
                        Dim PCT_INCREASE As Int64 = Val(rowPOTBATC5.Item("PCT_INCREASE") & "")
                        If PCT_INCREASE < 0 Then
                            MsgBox("% Increases must be >= 0", MsgBoxStyle.OkOnly, "Problem with Item Class " & STYLE_CLASS_CODE)
                            Exit Sub
                        End If
                        If PCT_INCREASE <= PCT_INCREASE_last Then
                            MsgBox("% Increases not Stepping in same Sequential Order as Customer Counts", MsgBoxStyle.OkOnly, "Problem with Item Class " & STYLE_CLASS_CODE)
                            Exit Sub
                        Else
                            PCT_INCREASE_last = PCT_INCREASE
                        End If
                    Next
                Next
                ASCMAIN1.Progress("Now Adding by %")
                For Each rowPOTBATC4 As DataRow In dst.Tables("POTBATC4").Select("")
                    Dim STYLE_CLASS_CODE As String = rowPOTBATC4.Item("STYLE_CLASS_CODE")
                    Dim CUST_COUNT_cnt As Int64 = rowPOTBATC4.GetChildRows("POTBATC4_POTBATC5").Count
                    If CUST_COUNT_cnt > 0 Then
                        Dim CUST_COUNTs(CUST_COUNT_cnt) As Int64
                        Dim PCT_INCREASEs(CUST_COUNT_cnt) As Int64
                        Dim I As Int64 = 0
                        For Each rowPOTBATC5 As DataRow In dst.Tables("POTBATC5").Select _
                            ("STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'", "CUST_COUNT")
                            I += 1
                            CUST_COUNTs(I) = Val(rowPOTBATC5.Item("CUST_COUNT") & "")
                            PCT_INCREASEs(I) = Val(rowPOTBATC5.Item("PCT_INCREASE") & "")
                        Next

                        Dim iCUST_COUNT As Int64 = 0
                        Dim records As Int64 = 0
                        Dim r2() As DataRow = dst.Tables("POTBATC2").Select _
                            ("STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "' and QTY_SHORT > 0 and CUST_OPEN >= " & CStr(CUST_COUNTs(1)), "CUST_OPEN")
                        Dim records_max As Int64 = r2.Length
                        For Each rowPOTBATC2 As DataRow In r2
                            rowPOTBATC2.Item("QTY_SHORT_STATIC") = rowPOTBATC2.Item("QTY_SHORT")
                            rowPOTBATC2.Item("CASE_QTY_STATIC") = rowPOTBATC2.Item("CASE_QTY")
                        Next
                        Manage_Expressions("Remove")
                        dst.Tables("POTBATC2").Columns("PO_QTY_CALC").Expression = "PO_QTY + (CASE_QTY_STATIC - PO_QTY%CASE_QTY_STATIC)%CASE_QTY_STATIC"

                        For Each rowPOTBATC2 As DataRow In r2 ' dst.Tables("POTBATC2").Select("STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "' and QTY_SHORT > 0 and CUST_OPEN >= " & CStr(CUST_COUNTs(1)), "CUST_OPEN")
                            Dim CUST_OPEN As Int64 = Val(rowPOTBATC2.Item("CUST_OPEN") & "")
                            Do While CUST_COUNT_cnt > iCUST_COUNT AndAlso CUST_OPEN >= CUST_COUNTs(iCUST_COUNT + 1)
                                iCUST_COUNT += 1
                            Loop
                            rowPOTBATC2.Item("PO_QTY") = Val(rowPOTBATC2.Item("QTY_SHORT_STATIC") & "") * (100 + PCT_INCREASEs(iCUST_COUNT)) / 100
                            rowPOTBATC2.Item("PO_QTY_ROUNDED") = rowPOTBATC2.Item("PO_QTY_CALC")
                            records += 1
                            ASCMAIN1.Progress("-", CStr(records) & "/" & CStr(records_max))
                        Next
                        Manage_Expressions("Restore")
                    End If
                Next
                ASCMAIN1.Progress("")
            Case "Set PO QTY = Rnd"
                Manage_Expressions("Remove")
                For Each rowPOTBATC2 As DataRow In dst.Tables("POTBATC2").Select("PO_QTY_ROUNDED > 0 and PO_QTY <> PO_QTY_ROUNDED")
                    rowPOTBATC2.Item("PO_QTY") = Val(rowPOTBATC2.Item("PO_QTY_ROUNDED"))
                Next
                Manage_Expressions("Restore")
            Case "Danny Order"
                With grdPOTBATC2.DisplayLayout.Bands("POTBATCS_POTBATC2")
                    .Columns.Item("NET_POS").Header.VisiblePosition = .Columns.Item("QTY_ONH").Header.VisiblePosition + 1
                    .Columns.Item("NET_POS2").Header.VisiblePosition = .Columns.Item("NET_POS").Header.VisiblePosition + 1

                    .Columns.Item("CUST_SOLD").Header.VisiblePosition = .Columns.Item("QTY_OPEN").Header.VisiblePosition + 1
                    .Columns.Item("QTY_SOLD").Header.VisiblePosition = .Columns.Item("CUST_SOLD").Header.VisiblePosition + 1
                    .Columns.Item("AMT_SOLD").Header.VisiblePosition = .Columns.Item("QTY_SOLD").Header.VisiblePosition + 1
                End With
            Case "Discontinue Item", "DNR Item"
                Dim NEW_STATUS As String = "D"
                Dim NEW_STATUS_DESC As String = "Discontinue"
                If e.Tool.Key = "DNR Item" Then
                    NEW_STATUS = "N"
                    NEW_STATUS_DESC = "DNR"
                End If
                Dim STYLE_CODE As String = grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
                Dim iResult As MsgBoxResult
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine(String.Format("This Will {0} The Following Item", NEW_STATUS_DESC))
                iMSG.AppendLine(String.Format("And All Of It's Non-{0}ed Colors:", NEW_STATUS_DESC))
                iMSG.AppendLine(STYLE_CODE)
                iMSG.AppendLine("")
                iMSG.AppendLine("Is This Really What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, NEW_STATUS_DESC & " Item")
                If iResult = MsgBoxResult.Yes Then
                    If ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE) Then
                        BeginTrans()
                        Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                        SQLS.AppendLine(String.Format("UPDATE ICTSTYL1 SET STYLE_STATUS = '{0}' WHERE STYLE_CODE = '{1}'", NEW_STATUS, STYLE_CODE))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                        Dim ORIG_STATUS As String = grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value
                        grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value = NEW_STATUS
                        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                        With rowASTAUDT1
                            .Item("TABLE_NAME") = "ICTSTYL1"
                            .Item("KEY_VALUE") = STYLE_CODE
                            .Item("COLUMN_NAME") = "STYLE_STATUS"
                            .Item("FM_MODE") = "E"
                            .Item("USER_ID") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                            .Item("OLD_VALUE") = ORIG_STATUS
                            .Item("NEW_VALUE") = NEW_STATUS
                            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                            .Item("SELECTION_NO") = Me.SELECTION_NO
                            .Item("XNO") = Me.XNO
                        End With
                        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                        Update_Record_TDA("ASTAUDT1")
                        For Each rowPOTBATC2 As DataRow In dst.Tables("POTBATC2").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                            If rowPOTBATC2.Item("STYLE_COLOR_STATUS") & "" <> NEW_STATUS Then
                                Dim ORIG_COLOR_STATUS As String = rowPOTBATC2.Item("STYLE_COLOR_STATUS") & ""
                                If NEW_STATUS = "N" And rowPOTBATC2.Item("STYLE_COLOR_STATUS") & "" = "D" Then
                                    'Don't DNR colors that are Discontinued
                                Else
                                    Dim COLOR_CODE As String = rowPOTBATC2.Item("COLOR_CODE").ToString
                                    SQLS.Length = 0
                                    SQLS.AppendLine(String.Format("UPDATE ICTSTYC1 SET STYLE_COLOR_STATUS = '{0}' WHERE STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", NEW_STATUS, STYLE_CODE, COLOR_CODE))
                                    ASCMAIN1.sql = SQLS.ToString
                                    ASCDATA1.ExecuteSQL()
                                    rowPOTBATC2.Item("STYLE_COLOR_STATUS") = NEW_STATUS
                                    Dim rowASTAUDTC As DataRow = dst.Tables("ASTAUDT1").NewRow
                                    With rowASTAUDTC
                                        .Item("TABLE_NAME") = "ICTSTYC1"
                                        .Item("KEY_VALUE") = STYLE_CODE & ":" & COLOR_CODE
                                        .Item("COLUMN_NAME") = "STYLE_COLOR_STATUS"
                                        .Item("FM_MODE") = "E"
                                        .Item("USER_ID") = ASCMAIN1.USER_ID
                                        .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                                        .Item("OLD_VALUE") = ORIG_COLOR_STATUS
                                        .Item("NEW_VALUE") = NEW_STATUS
                                        .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                                        .Item("SELECTION_NO") = Me.SELECTION_NO
                                        .Item("XNO") = Me.XNO
                                    End With
                                    dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDTC)
                                    Update_Record_TDA("ASTAUDT1")
                                End If
                            End If
                        Next
                        CommitTrans()
                        ASCMAIN1.MultiTask_Release(, , )
                    End If
                Else
                    MsgBox("Nothing Was Done.", vbOKOnly, NEW_STATUS_DESC & " Item")
                End If

            Case "Discontinue Color", "DNR Color"
                Dim NEW_STATUS As String = "D"
                Dim NEW_STATUS_DESC As String = "Discontinue"
                If e.Tool.Key = "DNR Color" Then
                    NEW_STATUS = "N"
                    NEW_STATUS_DESC = "DNR"
                End If
                Dim STYLE_CODE As String = grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
                Dim COLOR_CODE As String = grdPOTBATC2.Selected.Rows(0).Cells.Item("COLOR_CODE").Text
                Dim iResult As MsgBoxResult
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine(String.Format("This Will {0} The Following Color For Style {1}", NEW_STATUS_DESC, STYLE_CODE))
                iMSG.AppendLine("")
                iMSG.AppendLine(COLOR_CODE)
                iMSG.AppendLine("")
                iMSG.AppendLine("Is This Really What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, NEW_STATUS_DESC & " Item")
                If iResult = MsgBoxResult.Yes Then
                    If ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE) Then
                        BeginTrans()
                        Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                        Dim ORIG_COLOR_STATUS As String = grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Value
                        SQLS.AppendLine(String.Format("UPDATE ICTSTYC1 SET STYLE_COLOR_STATUS = '{0}' WHERE STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", NEW_STATUS, STYLE_CODE, COLOR_CODE))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                        grdPOTBATC2.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Value = NEW_STATUS

                        Dim rowASTAUDTC As DataRow = dst.Tables("ASTAUDT1").NewRow
                        With rowASTAUDTC
                            .Item("TABLE_NAME") = "ICTSTYC1"
                            .Item("KEY_VALUE") = STYLE_CODE & ":" & COLOR_CODE
                            .Item("COLUMN_NAME") = "STYLE_COLOR_STATUS"
                            .Item("FM_MODE") = "E"
                            .Item("USER_ID") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                            .Item("OLD_VALUE") = ORIG_COLOR_STATUS
                            .Item("NEW_VALUE") = NEW_STATUS
                            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                            .Item("SELECTION_NO") = Me.SELECTION_NO
                            .Item("XNO") = Me.XNO
                        End With
                        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDTC)
                        Update_Record_TDA("ASTAUDT1")
                        CommitTrans()
                        ASCMAIN1.MultiTask_Release(, , )
                    End If
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Private Sub grdPOTBATCX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTBATCX.AfterRowActivate

    End Sub

    Private Sub grdPOTBATCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTBATCX.DoubleClickRow
        If grdPOTBATCX.ActiveRow IsNot Nothing Then
            Absx1.txtFor("PO_BATCH_NO").Text = grdPOTBATCX.ActiveRow.Cells("PO_BATCH_NO").Text
            Click_Command("Load")
        End If
    End Sub

    Sub Setup_SOTSLSC1()
        Try
            If grdPOTBATC2.ActiveRow Is Nothing OrElse Not grdPOTBATC2.ActiveRow.IsDataRow Then
                grdSOTSLSC1.Visible = False
            Else
                Dim STYLE_CODE As String = grdPOTBATC2.ActiveRow.Cells("STYLE_CODE").Value
                If grdPOTBATC2.ActiveRow.Band.Key = "POTBATCS_POTBATC2" Then
                    Dim COLOR_CODE As String = grdPOTBATC2.ActiveRow.Cells("COLOR_CODE").Value
                    Fill_Records("SOTSLSC1", New Object() {STYLE_CODE, COLOR_CODE})
                    'Sort_grdColumns(grdSOTSLSC1, "WHSE_CODE,CUST_CODE")
                    grdSOTSLSC1.Text = "Style " & STYLE_CODE & " Color " & COLOR_CODE & "; Sales Summary"
                    grdSOTSLSC1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False
                Else
                    ASCMAIN1.sql = "Select SOTSLSC1.*, ARTCUST1.CUST_NAME" & vbCrLf _
                        & " from " & SOTSLSC1 & " SOTSLSC1, ARTCUST1 " & vbCrLf _
                        & " where SOTSLSC1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                        & "   and ARTCUST1.CUST_CODE (+) = SOTSLSC1.CUST_CODE"
                    Fill_Records("SOTSLSC1", "", True, ASCMAIN1.sql)
                    'Sort_grdColumns(grdSOTSLSC1, "WHSE_CODE,CUST_CODE")
                    grdSOTSLSC1.Text = "Style " & STYLE_CODE & " All Colors; Sales Summary"
                    grdSOTSLSC1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = True
                End If
                grdSOTSLSC1.Visible = True
                grdSOTSLSC1.DisplayLayout.Bands(0).SortedColumns.Clear()
                grdSOTSLSC1.DisplayLayout.Bands(0).SortedColumns.Add("WHSE_CODE", False, True)
                grdSOTSLSC1.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False)
                Expand_WHSE(grdSOTSLSC1)
                ' grdSOTSLSC1.Rows.ExpandAll(True)
            End If


        Catch ex As Exception

        End Try
    End Sub

    Sub Expand_WHSE(grd As UltraWinGrid.UltraGrid)
        For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
            If grow.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow, UltraWinGrid.UltraGridGroupByRow)
                If gbrow.Value = Absx1.txtFor("WHSE_CODE").Text Then
                    gbrow.Expanded = True
                End If
            End If
        Next
    End Sub

    Sub Setup_SOTORDC1()
        If grdPOTBATC2.ActiveRow Is Nothing OrElse Not grdPOTBATC2.ActiveRow.IsDataRow Then
            grdSOTORDC1.Visible = False
        Else
            Dim STYLE_CODE As String = grdPOTBATC2.ActiveRow.Cells("STYLE_CODE").Value

            If grdPOTBATC2.ActiveRow.Band.Key = "POTBATCS_POTBATC2" Then
                Dim COLOR_CODE As String = grdPOTBATC2.ActiveRow.Cells("COLOR_CODE").Value
                Fill_Records("SOTORDC1", New Object() {STYLE_CODE, COLOR_CODE})
                grdSOTORDC1.Text = "Style " & STYLE_CODE & " Color " & COLOR_CODE & "; Open Orders"
                grdSOTORDC1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = True
            Else
                ASCMAIN1.sql = "Select SOTORDC1.*, ARTCUST1.CUST_NAME" & vbCrLf _
                    & " from " & SOTORDC1 & " SOTORDC1, ARTCUST1 " & vbCrLf _
                    & " where SOTORDC1.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE (+) = SOTORDC1.CUST_CODE"
                Fill_Records("SOTORDC1", "", True, ASCMAIN1.sql)
                grdSOTORDC1.Text = "Style " & STYLE_CODE & " All Colors; Sales Summary"
                grdSOTORDC1.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False
            End If
            grdSOTORDC1.DisplayLayout.Bands(0).SortedColumns.Clear()
            grdSOTORDC1.DisplayLayout.Bands(0).SortedColumns.Add("WHSE_CODE", False, True)
            grdSOTORDC1.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False)
            Expand_WHSE(grdSOTORDC1)
            ' grdSOTORDC1.Rows.ExpandAll(True)
            grdSOTORDC1.Visible = True
        End If
    End Sub

    Sub Setup_POTORDRX()
        If grdPOTBATC2.ActiveRow Is Nothing OrElse Not grdPOTBATC2.ActiveRow.IsDataRow Then
            grdPOTORDRX.Visible = False
        Else
            Dim STYLE_CODE As String = grdPOTBATC2.ActiveRow.Cells("STYLE_CODE").Value

            If grdPOTBATC2.ActiveRow.Band.Key = "POTBATCS_POTBATC2" Then
                Dim COLOR_CODE As String = grdPOTBATC2.ActiveRow.Cells("COLOR_CODE").Value
                Fill_Records("POTORDRX", New Object() {STYLE_CODE, COLOR_CODE})
                grdPOTORDRX.Text = "Style " & STYLE_CODE & " Color " & COLOR_CODE & "; Open Orders"
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = True
            Else
                ASCMAIN1.sql = "Select POTORDRX.*, APTVEND1.VEND_NAME" & vbCrLf _
                & " from " & POTORDRX & " POTORDRX, APTVEND1 " & vbCrLf _
                & " where STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and APTVEND1.VEND_CODE (+) = POTORDRX.VEND_CODE"
                Fill_Records("POTORDRX", "", True, ASCMAIN1.sql)
                grdPOTORDRX.Text = "Style " & STYLE_CODE & " All Colors; Purcase Order History Summary"
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False
            End If

            ' Sort_grdColumns(grdPOTORDRX, "PO_ORDER_NO")
            grdPOTORDRX.DisplayLayout.Bands(0).SortedColumns.Clear()
            grdPOTORDRX.DisplayLayout.Bands(0).SortedColumns.Add("WHSE_CODE", False, True)
            grdPOTORDRX.DisplayLayout.Bands(0).SortedColumns.Add("PO_ORDER_NO", False)
            Expand_WHSE(grdPOTORDRX)
            ' grdPOTORDRX.Rows.ExpandAll(True)
            grdPOTORDRX.Visible = True
        End If
    End Sub

    Sub Load_POTBATCX()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Fill_Records("POTBATCX")
        Sort_grdColumns(grdPOTBATCX, "PO_BATCH_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdPOTBATCX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTBATCX.InitializeRow
        If e.Row.Cells("BATCH_STATUS").Value & "" <> "O" Then
            e.Row.CellAppearance.BackColor = Drawing.Color.LightGray
        End If
    End Sub

#Region "grdPOTBATC2"

    Private Sub grdPOTBATC2_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdPOTBATC2.AfterCellUpdate
        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If e.Cell.Column.Key = "PO_QTY" Then
                If chkAutoSave.Checked Then
                    generate_POs = False
                    Update_Record(True)
                End If
            End If
        End If
    End Sub
    Private Sub grdPOTBATC2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTBATC2.AfterRowActivate
        Setup_SOTSLSC1()
        Setup_SOTORDC1()
        Setup_POTORDRX()

        If tabDetails.SelectedTab.Key = "Style Information" Then
            FetchImage()
        End If

        EcomIndicator()
    End Sub

    Private Sub EcomIndicator()
        Try
            If Not (grdPOTBATC2.ActiveRow Is Nothing OrElse Not grdPOTBATC2.ActiveRow.IsDataRow) Then
                Dim STYLE_CODE As String = grdPOTBATC2.ActiveRow.Cells("STYLE_CODE").Value
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    'ASCMAIN1.sql = String.Format("SELECT COUNT(*) FROM ECTESTY1 WHERE STYLE_CODE = '{0}'", STYLE_CODE)
                    'Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                    'If REC_CNT > 0 Then
                    '    lblEcomStyle.Visible = True
                    'Else
                    '    lblEcomStyle.Visible = False
                    'End If
                    Dim ECOM_MSG As String = TAC.TACMAIN1.getEcomInfo(Me, STYLE_CODE)
                    If ECOM_MSG.Length > 0 Then
                        lblEcomStyle.Visible = True
                        Dim TTI As New UltraWinToolTip.UltraToolTipInfo
                        If Not IsNothing(TTM.GetUltraToolTip(lblEcomStyle)) Then
                            TTI.ToolTipTitle = "E-Commerce Information:"
                            TTM.AutoPopDelay = 20000
                            TTI.ToolTipTextFormatted = ECOM_MSG
                            TTM.SetUltraToolTip(lblEcomStyle, TTI)
                        Else
                            TTI.ToolTipTextFormatted = ECOM_MSG
                        End If
                    Else
                        lblEcomStyle.Visible = False
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdPOTBATC2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTBATC2.AfterRowUpdate
        If grdPOTBATC2.Tag & "" <> "X" Then
            grdPOTBATC2.Tag = "X"
            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                If Not IsNothing(e.Row.Cells("PO_QTY_ROUNDED").Value) Then
                    If e.Row.Band.Index = 1 Then
                        e.Row.Cells("PO_QTY_ROUNDED").Value = e.Row.Cells("PO_QTY_CALC").Value
                    End If
                End If
            Else
                e.Row.Cells("PO_QTY_ROUNDED").Value = e.Row.Cells("PO_QTY_CALC").Value
            End If
            e.Row.Update()
            grdPOTBATC2.Tag = ""
        End If
    End Sub

    Private Sub grdPOTBATC2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTBATC2.BeforeRowUpdate
    End Sub

    Private Sub grdPOTBATC2_ClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.ClickCellEventArgs) Handles grdPOTBATC2.ClickCell

    End Sub

    Private Sub grdPOTBATC2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTBATC2.InitializeRow
        If e.Row.Band.Key = "POTBATCS" Then
            e.Row.Cells("STYLE_STATUS").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("STYLE_CODE").ToolTipText = ""
            e.Row.Cells("CARTON_PACK_QTY").Appearance.ForeColor = Drawing.Color.Empty
        End If
        e.Row.Cells("PO_QTY_ROUNDED").Appearance.ForeColor = Drawing.Color.Empty
        e.Row.Cells("PO_QTY_ROUNDED").ToolTipText = ""

        If e.Row.Band.Key = "POTBATCS" Then
            If e.Row.Cells("STYLE_STATUS").Value & "" = "D" Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("STYLE_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("STYLE_CODE").ToolTipText = "Style Status is Discontinued"
            ElseIf e.Row.Cells("STYLE_STATUS").Value & "" = "N" Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.DarkOrange
                e.Row.Cells("STYLE_STATUS").Appearance.ForeColor = Drawing.Color.DarkOrange
                e.Row.Cells("STYLE_CODE").ToolTipText = "Style Status is Do Not Re-Order"
            Else
                If Val(e.Row.Cells("STYLE_PO_QTY_MIN").Value & "") > Val(e.Row.Cells("PO_QTY_ROUNDED").Value & "") Then
                    e.Row.Cells("PO_QTY_ROUNDED").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("PO_QTY_ROUNDED").ToolTipText = "Less than MOQ"
                    ' If e.Row.Cells("STYLE_CODE").Value = "MTH10925" Then Stop
                Else
                    e.Row.Cells("PO_QTY_ROUNDED").Appearance.ForeColor = Drawing.Color.Empty
                    e.Row.Cells("PO_QTY_ROUNDED").ToolTipText = ""
                End If
            End If
            e.Row.Cells("CARTON_PACK_QTY").Appearance.ForeColor = Drawing.Color.Blue
        Else
            If Val(e.Row.ParentRow.Cells("STYLE_PO_QTY_MIN").Value & "") > Val(e.Row.Cells("PO_QTY_ROUNDED").Value & "") Then
                e.Row.Cells("PO_QTY_ROUNDED").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("PO_QTY_ROUNDED").ToolTipText = "Less than MOQ"
            Else
                e.Row.Cells("PO_QTY_ROUNDED").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("PO_QTY_ROUNDED").ToolTipText = ""
            End If
        End If
    End Sub

#End Region

    Private Sub cmdFetchSales_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetchSales.Click
        Fetch_Sales()
    End Sub

    Sub Fetch_Sales()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("Now Fetching Sales")

        Dim SQLw As String = ""
        If optHistory.Value = "M" Then
            SQLw = " and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * (Val(numLastXXMonths.Value & "") - 1)) & "' and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & ASCMAIN1.CYP & "'"
        Else
            SQLw = " and SOTINVH1.INV_DATE >= '" & Format(dteFrom.Value, "dd-MMM-yyyy") & "'" _
                 & " and SOTINVH1.INV_DATE <= '" & Format(dteTo.Value, "dd-MMM-yyyy") & "'"
        End If
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTSLSC1)
        ASCMAIN1.sql = Replace(sqlSOTSLSC1, " and ROWNUM < 1", SQLw)
        ASCDATA1.ExecuteSQL("Insert into " & SOTSLSC1 & " " & ASCMAIN1.sql)


        Manage_Expressions("Remove")
        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, Count (Distinct CUST_CODE) CUST_SOLD" _
            & ", SUM (QTY) QTY, SUM (AMT) AMT" _
            & "   from " & SOTSLSC1 & " SOTSLSC1" _
            & " where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" _
            & "   group by STYLE_CODE, COLOR_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowPOTBATC2 As DataRow = dst.Tables("POTBATC2").Rows.Find(New String() {PO_BATCH_NO, row.Item("STYLE_CODE"), row.Item("COLOR_CODE")})
            rowPOTBATC2.Item("CUST_SOLD") = row.Item("CUST_SOLD")
            rowPOTBATC2.Item("QTY_SOLD") = row.Item("QTY")
            rowPOTBATC2.Item("AMT_SOLD") = row.Item("AMT")
        Next
        Manage_Expressions("Restore")


        Setup_SOTSLSC1()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub optHistory_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optHistory.ValueChanged
        lblFrom.Visible = (optHistory.Value = "D")
        lblTo.Visible = (optHistory.Value = "D")
        dteFrom.Visible = (optHistory.Value = "D")
        dteTo.Visible = (optHistory.Value = "D")
        numLastXXMonths.Visible = (optHistory.Value = "M")
    End Sub

    Sub Create_Temp_Tables()
        ASCMAIN1.sql = "Select POTBATC2.*, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_STATUS" _
             & " from POTBATC2,ICTCOLR1,ICTSTYC1" _
             & " where ICTCOLR1.COLOR_CODE = POTBATC2.COLOR_CODE" _
             & "   and ICTSTYC1.STYLE_CODE = POTBATC2.STYLE_CODE" _
             & "   and ICTSTYC1.COLOR_CODE = POTBATC2.COLOR_CODE" _
             & "   and ROWNUM < 1"
        POTBATC2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTBATC2 & " Add Primary Key (PO_BATCH_NO, STYLE_CODE, COLOR_CODE)")
        ASCDATA1.ExecuteSQL("Create Unique Index I_" & POTBATC2 & "_1 on " & POTBATC2 & " (STYLE_CODE, COLOR_CODE)")

        sqlSOTSLSC1 = "Select SOTINVH1.WHSE_CODE, SOTINVH1.CUST_CODE" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
            & ", SUM(SOTINVH2.ORDR_QTY_SHIP) QTY" & vbCrLf _
            & ", SUM(SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) AMT" & vbCrLf _
            & ", COUNT(*) CNT, MAX(SOTINVH1.INV_DATE) LAST_INV" & vbCrLf _
            & " from SOTINVH1,SOTINVH2," & POTBATC2 & " POTBATC2" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.STYLE_CODE = POTBATC2.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH2.COLOR_CODE = POTBATC2.COLOR_CODE" & vbCrLf _
            & "   and ROWNUM < 1" & vbCrLf _
            & " group by SOTINVH1.WHSE_CODE, SOTINVH1.CUST_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE"
        SOTSLSC1 = ASCMAIN1.Temp_Table(sqlSOTSLSC1)
        ' ASCDATA1.ExecuteSQL("Alter Table " & SOTSLSC1 & " Add Primary Key (WHSE_CODE, CUST_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTSLSC1 & "_1 on " & SOTSLSC1 & " (STYLE_CODE, COLOR_CODE)")

        sqlSOTORDC1 = "Select SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_OPEN" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_PICK" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_DATE_RECD" & vbCrLf _
            & " from SOTORDR1,SOTORDR2," & POTBATC2 & " POTBATC2" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.STYLE_CODE = POTBATC2.STYLE_CODE" & vbCrLf _
            & "   and SOTORDR2.COLOR_CODE = POTBATC2.COLOR_CODE" & vbCrLf _
            & "   and SOTORDR2.ORDR_STATUS in ('O','P')" & vbCrLf _
            & "   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) <> 0 or NVL(SOTORDR2.ORDR_QTY_PICK,0) <> 0)"
        SOTORDC1 = ASCMAIN1.Temp_Table(sqlSOTORDC1)
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTORDC1 & "_1 on " & SOTORDC1 & " (STYLE_CODE, COLOR_CODE)")

        sqlPOTORDRX = "" _
            & "Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO, POTORDR1.PO_REFERENCE" & vbCrLf _
            & ", POTORDR1.WHSE_CODE, POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE" & vbCrLf _
            & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ", POTORDR2.PO_COST, POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_QTY_SHP, POTORDR2.PO_QTY_REC" & vbCrLf _
            & ", POTORDR1.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_ETA" & vbCrLf _
            & " from POTORDR1,POTORDR2," & POTBATC2 & " POTBATC2" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.STYLE_CODE = POTBATC2.STYLE_CODE" & vbCrLf _
            & "   and POTORDR2.COLOR_CODE = POTBATC2.COLOR_CODE"
        POTORDRX = ASCMAIN1.Temp_Table(sqlPOTORDRX)
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDRX & " add Primary Key (PO_ORDER_NO, PO_ORDER_LNO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & POTORDRX & "_1 on " & POTORDRX & " (STYLE_CODE, COLOR_CODE)")
    End Sub

    Sub Scan_for_New_Styles()

    End Sub

    Private Sub optAS_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optAS.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_POTBATC2()
    End Sub

    Sub Setup_POTBATC2()
        Dim dvw As DataView = DirectCast(grdPOTBATC2.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        Select Case optAS.Value
            Case "A"
                sql = "STYLE_STATUS = 'A'"
            Case "S"
                sql = "QTY_SHORT > 0"
            Case "AS"
                sql = "STYLE_STATUS = 'A' OR QTY_SHORT > 0"
        End Select
        dvw.RowFilter = sql
    End Sub

    Private Sub grdSOTORDC1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDC1.InitializeRow
        If e.Row.Cells("WHSE_CODE").Value & "" <> Absx1.txtFor("WHSE_CODE").Text Then
            e.Row.Cells("WHSE_CODE").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("WHSE_CODE").ToolTipText = "Warehouse other than " & Absx1.txtFor("WHSE_CODE").Text
        End If
    End Sub

    Private Sub grdSOTSLSC1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSLSC1.InitializeRow
        If e.Row.Cells("WHSE_CODE").Value & "" <> Absx1.txtFor("WHSE_CODE").Text Then
            e.Row.Cells("WHSE_CODE").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("WHSE_CODE").ToolTipText = "Warehouse other than " & Absx1.txtFor("WHSE_CODE").Text
        End If
    End Sub

    Private Sub grdPOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRX.InitializeRow
        If e.Row.Cells("WHSE_CODE").Value & "" <> Absx1.txtFor("WHSE_CODE").Text Then
            e.Row.Cells("WHSE_CODE").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("WHSE_CODE").ToolTipText = "Warehouse other than " & Absx1.txtFor("WHSE_CODE").Text
        End If
    End Sub

    Private Sub grdPOTBATC2_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdPOTBATC2.MouseUp
        If grdPOTBATC2.ActiveCell IsNot Nothing AndAlso grdPOTBATC2.ActiveCell.Column.Key = "PO_SEL" Then
            grdPOTBATC2.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdPOTBATC3_ClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.ClickCellEventArgs) Handles grdPOTBATC3.ClickCell

    End Sub

    Private Sub grdPOTBATC3_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTBATC3.InitializeLayout

    End Sub

    Private Sub grdPOTBATC3_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdPOTBATC3.MouseUp
        If grdPOTBATC3.ActiveCell IsNot Nothing AndAlso grdPOTBATC3.ActiveCell.Column.Key = "PO_SEL" Then
            grdPOTBATC3.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdPOTBATC2_changes_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTBATC2_changes.DoubleClickRow
        tabMain.SelectedTab = tabMain.Tabs("Worksheet by Style/Color")
        For Each grow As UltraWinGrid.UltraGridRow In grdPOTBATC2.Rows
            If grow.Cells("STYLE_CODE").Value = e.Row.Cells("STYLE_CODE").Value Then
                grow.Expanded = True
                grdPOTBATC2.DisplayLayout.RowScrollRegions(0).FirstRow = grow
                For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                    If grow2.Cells("COLOR_CODE").Value = e.Row.Cells("COLOR_CODE").Value Then
                        grdPOTBATC2.ActiveRow = grow2
                        grow2.Selected = True
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End Sub

    Private Sub grdPOTBATC2_AfterRowExpanded(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTBATC2.AfterRowExpanded
        grdPOTBATC2.ActiveRow = e.Row
    End Sub

    Private Sub imgSTYLE_Click(sender As System.Object, e As System.EventArgs) Handles imgSTYLE.Click

    End Sub

    Function FetchImage() As Byte()

        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        Dim IMAGE_NAME As String = ""

        If grdPOTBATC2.ActiveRow IsNot Nothing AndAlso grdPOTBATC2.ActiveRow.IsDataRow Then
            If grdPOTBATC2.ActiveRow.Band.Key = "POTBATCS_POTBATC2" Then
                STYLE_CODE = grdPOTBATC2.ActiveRow.Cells("STYLE_CODE").Value
                COLOR_CODE = grdPOTBATC2.ActiveRow.Cells("COLOR_CODE").Value
                IMAGE_NAME = STYLE_CODE & "-" & COLOR_CODE
            End If
        End If

        Dim imgba() As Byte = Nothing

        If IMAGE_NAME <> "" Then
            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            imgSTYLE.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
            grpStylePicture.Text = "Style " & STYLE_CODE & "-" & COLOR_CODE
        Else
            imgSTYLE.Image = Nothing
            grpStylePicture.Text = ""
        End If

        Return imgba

    End Function

    Private Sub tabDetails_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        If tabDetails.SelectedTab.Key = "Style Information" Then
            FetchImage()
        End If
    End Sub


    Sub Manage_Expressions(action As String)

        If action = "Remove" Then

            ' Remove Expressions

            POTBATCS_expressions.Clear()
            For Each dcol As DataColumn In dst.Tables("POTBATCS").Columns
                If dcol.Expression <> "" Then
                    POTBATCS_expressions.Add(dcol.ColumnName, dcol.Expression)
                    dcol.Expression = ""
                End If
            Next

            POTBATC2_expressions.Clear()
            For Each dcol As DataColumn In dst.Tables("POTBATC2").Columns
                If dcol.Expression <> "" Then
                    POTBATC2_expressions.Add(dcol.ColumnName, dcol.Expression)
                    dcol.Expression = ""
                End If
            Next

        Else

            ' Restore Expressions

            For Each COLUMN_NAME As String In POTBATC2_expressions.Keys
                dst.Tables("POTBATC2").Columns(COLUMN_NAME).Expression = POTBATC2_expressions(COLUMN_NAME)
            Next

            For Each COLUMN_NAME As String In POTBATCS_expressions.Keys
                dst.Tables("POTBATCS").Columns(COLUMN_NAME).Expression = POTBATCS_expressions(COLUMN_NAME)
            Next
        End If

    End Sub
End Class