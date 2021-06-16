
Public Class SOFRUSSE
    Dim CUST_CODE As String
    Dim ORDR_NO As String
    Dim WHSE_CODE As String
    Dim InquiryOnly As Boolean = False
    Dim FormLoaded As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("Select")
            SQLs.AppendLine(" ORDR_NO, ORDR_LNO,")
            SQLs.AppendLine(" STYLE_CODE, COLOR_CODE,")
            SQLs.AppendLine(" EDI_DTL_SEQ, ORDR_QTY, ORDR_UNIT_PRICE,")
            SQLs.AppendLine(" EDI_DOC_SEQ_NO, CUST_UPC, CUST_SKU,")
            SQLs.AppendLine(" ORDR_QTY NEW_QTY, CUST_UPC NEW_UPC,")
            SQLs.AppendLine(" CUST_SKU NEW_SKU, COLOR_CODE NEW_COLOR_CODE, ORDR_UNIT_PRICE NEW_ORDR_UNIT_PRICE,")
            SQLs.AppendLine(" RANGE_STYLE_CODE, 0 AS RANGE_STYLE_QTY_PP")
            SQLs.AppendLine(" from SOTORDR2 where ORDR_NO = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "SOTRUSSE", "**", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            Dim SQLW As String = ""
            ASCMAIN1.sql = String.Format("SELECT SOTORDR1.* FROM SOTORDR1 WHERE CUST_CODE = '{0}' AND ORDR_STATUS = 'O'{1}", CUST_CODE, SQLW)
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False)
            Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR2 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR5 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 1, False)
            Fill_Records("ICTSTYL1", "", , ASCMAIN1.sql)
        End With

        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdSOTRUSSE.DataSource = dst.Tables("SOTRUSSE")

        Create_Summary(grdSOTRUSSE, "NEW_QTY", "Sum", "", "###,##0")
        Create_Summary(grdSOTRUSSE, "ORDR_QTY", "Sum", "", "###,##0")

        Sort_grdColumns(grdSOTORDRX, "ORDR_DATE, ORDR_GROUP_NO, ORDR_NO".ToLower(), False)

        grdSOTRUSSE.DisplayLayout.Bands(0).Columns("NEW_QTY").Format = "###,##0"
        grdSOTRUSSE.DisplayLayout.Bands(0).Columns("ORDR_QTY").Format = "###,##0"
        grdSOTRUSSE.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = "###,##0.00"
        grdSOTRUSSE.DisplayLayout.Bands(0).Columns("NEW_ORDR_UNIT_PRICE").Format = "###,##0.00"
        grdSOTRUSSE.DisplayLayout.Bands(0).Columns("RANGE_STYLE_QTY_PP").Format = "###,##0"

        grdSOTRUSSE.DisplayLayout.UseFixedHeaders = True
        With grdSOTRUSSE.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "ORDR_QTY", "ORDR_UNIT_PRICE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        tab.Visible = False
        grdSOTORDRX.Parent = tab.Parent

        FormLoaded = True
    End Sub

    Private Sub Proceed_WhseCheck()
        If WHSE_CODE.Length = 0 Then
            EMsg &= vbCr & "Problem Figuring Out Warehouse For This Order."
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit"
                If Absx1.txtFor("ORDR_NO").Text.Length = 0 Then
                    If grdSOTORDRX.Selected.Rows.Count <= 0 Then
                        EMsg &= vbCr & "You Must First Select an Order"
                    Else
                        If grdSOTORDRX.Selected.Rows.Count <> 1 Then
                            EMsg &= vbCr & "You May Only Select One Order At A Time To Edit"
                        Else
                            Absx1.txtFor("ORDR_NO").Text = grdSOTORDRX.Selected.Rows(0).Cells("ORDR_NO").Text
                        End If
                    End If
                End If
                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", Absx1.txtFor("ORDR_NO").Text) Then
                        Exit Sub
                    End If
                End If

                Select Case Absx1.txtFor("CUST_CODE").Text
                    Case "TARGET", "SEARSCAN", "CHARLOT", "COSTCOUS"
                        'These are allowed
                    Case Else
                        EMsg &= vbCr & "You May Only Select Target or Sears Canada or Charlot for this screen."
                End Select

            Case "Cancel"
                Dim iResult As MsgBoxResult = MsgBox("Cancelling Will Lose Any Changes You May Have Made." & vbCrLf & "Are You Sure You Want To Cancel?", MsgBoxStyle.YesNo, "Cancel Confirmation")
                If iResult = MsgBoxResult.No Then
                    EMsg &= vbCr & "Cancel Option Aborted"
                End If
            Case "Update"
                Proceed_IfOrderIsOpen()
                Proceed_WhseCheck()
                Proceed_UpDateRange()
                Proceed_LineErrors()
                If chkCreateRanges.Checked Then
                    Proceed_CreateRange()
                End If
            Case "Done"
                Mode_Settings(False)
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)
                Call Clear_Record()
            Case "Cancel", "Done"
                Call Mode_Settings(False)
                Call Clear_Record()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Edit").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode

                .Groups("Screen Control").Items("Update").Visible = ScreenMode
                .Groups("Screen Control").Items("Cancel").Visible = ScreenMode

            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSOTORDRX.Visible = Not tf

        With grdSOTRUSSE.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdSOTRUSSE.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTRUSSE.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        For Each COLNAME As String In New String() {"NEW_QTY", "NEW_UPC", "NEW_SKU", "NEW_COLOR_CODE", "NEW_ORDR_UNIT_PRICE", "RANGE_STYLE_CODE", "RANGE_STYLE_QTY_PP"}
            grdSOTRUSSE.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        Next
        For Each COLNAME As String In New String() {"NEW_QTY", "NEW_UPC", "NEW_SKU", "NEW_COLOR_CODE", "NEW_ORDR_UNIT_PRICE", "RANGE_STYLE_CODE", "RANGE_STYLE_QTY_PP"}
            grdSOTRUSSE.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Next

        If Not ScreenMode Then
            RefreshSOTORDRX()
        End If
    End Sub

    Sub Clear_Record()
        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR2").Rows.Clear()
        dst.Tables("SOTORDR5").Rows.Clear()
        dst.Tables("SOTRUSSE").Rows.Clear()

        Absx1.txtFor("ORDR_NO").Text = ""
        Absx1.txtFor("ORDR_CUST_PO").Text = ""
    End Sub

    Sub Load_Record()
        ORDR_NO = Absx1.txtFor("ORDR_NO").Text
        Dim HasRangeStyles As Boolean = False
        Call Save_Header_Fields(UltraGroupBox1)

        Dim ORDR_NOs(0) As String
        Dim ORDRcnt As Integer = 0
        Using tblSOTORDR1 As DataTable = ASCDATA1.GetDataTable(String.Format("SELECT ORDR_NO FROM SOTORDR1 WHERE ORDR_NO = '{0}'", Absx1.txtFor("ORDR_NO").Text))
            For Each rowSOTORDR1 As DataRow In tblSOTORDR1.Rows
                ReDim Preserve ORDR_NOs(ORDRcnt)
                ORDR_NOs(ORDRcnt) = rowSOTORDR1.Item("ORDR_NO").ToString
                ORDRcnt += 1
            Next
        End Using

        EnforceConstraints(False)
        If EntryMode = "E" Then
            For Each rowORDR_NO As String In ORDR_NOs
                Call Fill_Records("SOTORDR1", rowORDR_NO, False)
                Call Fill_Records("SOTORDR2", rowORDR_NO, False)
                Call Fill_Records("SOTORDR5", rowORDR_NO, False)
                Call Fill_Records("SOTRUSSE", rowORDR_NO, False)
                Dim SQLS As New System.Text.StringBuilder
                SQLS.Length = 0
                SQLS.AppendLine("SELECT COUNT(DISTINCT(RANGE_STYLE_CODE)) AS RANGE_STYLE_CNT")
                SQLS.AppendLine(String.Format("FROM SOTORDR2 WHERE ORDR_NO = '{0}'", rowORDR_NO))
                ASCMAIN1.sql = SQLS.ToString()
                Dim RangeCount As Int16 = Val(ASCDATA1.GetDataValue)
                If HasRangeStyles = False And RangeCount > 0 Then
                    HasRangeStyles = True
                End If
            Next
        End If
        'Call Fill_Records("ARTCUST1", Absx1.txtFor("CUST_CODE").Text, True)
        'Call Fill_Records("ARTCUST2", Absx1.txtFor("CUST_CODE").Text, True)

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        If HasRangeStyles Then
            chkIsRANGE.Checked = True
            chkUpdateRange.Visible = True
            txtNewRange.Text = ""
        Else
            chkIsRANGE.Checked = False
            chkUpdateRange.Visible = False
            lblNewRange.Visible = False
            txtNewRange.Visible = False
            txtNewRange.Text = ""
        End If
        SetCreateRangeState()
        SetRangeColumnState()

    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()
        'For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        'Next
        'Call CommitTrans("Order / Quote Deleted")
    End Sub

    Sub Update_Record()
        Call BeginTrans()

        Dim SQLQ As New System.Text.StringBuilder
        Dim EDI_DOC_SEQ_NO As String = ""
        Dim NewRanges As New List(Of String)
        For Each rowSOTRUSSE As DataRow In dst.Tables("SOTRUSSE").Select()
            If rowSOTRUSSE.Item("RANGE_STYLE_CODE").ToString & "" <> "" Then
                If Not NewRanges.Contains(rowSOTRUSSE.Item("RANGE_STYLE_CODE").ToString & "") Then
                    NewRanges.Add(rowSOTRUSSE.Item("RANGE_STYLE_CODE").ToString & "")
                End If
            End If
            EDI_DOC_SEQ_NO = rowSOTRUSSE.Item("EDI_DOC_SEQ_NO")
            If Val(rowSOTRUSSE.Item("ORDR_QTY")) <> Val(rowSOTRUSSE.Item("NEW_QTY")) Then
                SQLQ.Length = 0
                SQLQ.AppendLine(String.Format("UPDATE ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - {0} + {1}", Val(rowSOTRUSSE.Item("ORDR_QTY")), Val(rowSOTRUSSE.Item("NEW_QTY"))))
                SQLQ.AppendLine(String.Format(" WHERE WHSE_CODE = '{0}'", WHSE_CODE))
                SQLQ.AppendLine(String.Format(" AND STYLE_CODE = '{0}'", rowSOTRUSSE.Item("STYLE_CODE")))
                SQLQ.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", rowSOTRUSSE.Item("COLOR_CODE")))
                ASCMAIN1.sql = SQLQ.ToString
                ASCDATA1.ExecuteSQL()
                SQLQ.Length = 0
                SQLQ.AppendLine(String.Format("UPDATE SOTORDR2 SET ORDR_QTY = {0},", Val(rowSOTRUSSE.Item("NEW_QTY"))))
                SQLQ.AppendLine(String.Format(" ORDR_QTY_OPEN = {0}", Val(rowSOTRUSSE.Item("NEW_QTY"))))
                SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", rowSOTRUSSE.Item("ORDR_NO")))
                SQLQ.AppendLine(String.Format(" AND ORDR_LNO = {0}", rowSOTRUSSE.Item("ORDR_LNO")))
                ASCMAIN1.sql = SQLQ.ToString
                ASCDATA1.ExecuteSQL()
            End If

            If rowSOTRUSSE.Item("CUST_UPC") <> rowSOTRUSSE.Item("NEW_UPC") Then
                If rowSOTRUSSE.Item("NEW_UPC").ToString.Length > 0 Then
                    SQLQ.Length = 0
                    SQLQ.AppendLine(String.Format("UPDATE SOTORDR2 SET CUST_UPC = '{0}'", rowSOTRUSSE.Item("NEW_UPC")))
                    SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", rowSOTRUSSE.Item("ORDR_NO")))
                    SQLQ.AppendLine(String.Format(" AND ORDR_LNO = {0}", rowSOTRUSSE.Item("ORDR_LNO")))
                    ASCMAIN1.sql = SQLQ.ToString
                    ASCDATA1.ExecuteSQL()
                    If Not chkUpdateRange.Checked Then
                        SQLQ.Length = 0
                        SQLQ.AppendLine(String.Format("UPDATE EDT850T2 SET EDI_UPC = '{0}'", rowSOTRUSSE.Item("NEW_UPC")))
                        SQLQ.AppendLine(String.Format(" WHERE EDT850T2.EDI_DOC_SEQ_NO = '{0}'", rowSOTRUSSE.Item("EDI_DOC_SEQ_NO")))
                        SQLQ.AppendLine(String.Format(" AND EDT850T2.EDI_DTL_SEQ = '{0}'", rowSOTRUSSE.Item("EDI_DTL_SEQ")))
                        ASCMAIN1.sql = SQLQ.ToString
                        ASCDATA1.ExecuteSQL()
                    End If
                End If
            End If

            If rowSOTRUSSE.Item("CUST_SKU") <> rowSOTRUSSE.Item("NEW_SKU") Then
                If rowSOTRUSSE.Item("NEW_SKU").ToString.Length > 0 Then
                    SQLQ.Length = 0
                    SQLQ.AppendLine(String.Format("UPDATE SOTORDR2 SET CUST_SKU = '{0}'", Val(rowSOTRUSSE.Item("NEW_SKU"))))
                    SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", rowSOTRUSSE.Item("ORDR_NO")))
                    SQLQ.AppendLine(String.Format(" AND ORDR_LNO = {0}", rowSOTRUSSE.Item("ORDR_LNO")))
                    ASCMAIN1.sql = SQLQ.ToString
                    ASCDATA1.ExecuteSQL()
                    If Not chkUpdateRange.Checked Then
                        SQLQ.Length = 0
                        SQLQ.AppendLine(String.Format("UPDATE EDT850T2 SET EDI_SKU = '{0}'", rowSOTRUSSE.Item("NEW_SKU")))
                        SQLQ.AppendLine(String.Format(" WHERE EDT850T2.EDI_DOC_SEQ_NO = '{0}'", rowSOTRUSSE.Item("EDI_DOC_SEQ_NO")))
                        SQLQ.AppendLine(String.Format(" AND EDT850T2.EDI_DTL_SEQ = '{0}'", rowSOTRUSSE.Item("EDI_DTL_SEQ")))
                        ASCMAIN1.sql = SQLQ.ToString
                        ASCDATA1.ExecuteSQL()
                    End If
                End If
            End If
        Next
        If chkUpdateRange.Checked Then
            SQLQ.Length = 0
            SQLQ.AppendLine(String.Format("UPDATE EDT850T2 SET EDI_SKU = '{0}'", txtNewRange.Text))
            SQLQ.AppendLine(String.Format(" WHERE EDT850T2.EDI_DOC_SEQ_NO = '{0}'", EDI_DOC_SEQ_NO))
            ASCMAIN1.sql = SQLQ.ToString
            ASCDATA1.ExecuteSQL()
            SQLQ.Length = 0
            SQLQ.AppendLine(String.Format("UPDATE SOTORDR2 SET RANGE_STYLE_CODE = '{0}'", txtNewRange.Text))
            SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", Absx1.txtFor("ORDR_NO").Text))
            ASCMAIN1.sql = SQLQ.ToString
            ASCDATA1.ExecuteSQL()
            SQLQ.Length = 0
            SQLQ.AppendLine(String.Format("UPDATE SOTORDR9 SET RANGE_STYLE_CODE = '{0}'", txtNewRange.Text))
            SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", Absx1.txtFor("ORDR_NO").Text))
            ASCMAIN1.sql = SQLQ.ToString
            ASCDATA1.ExecuteSQL()
        End If

        Dim TotalOrderPriceOrig As Double = 0
        Dim TotalOrderPriceNew As Double = 0
        Dim TotalOrderQTYOrig As Double = 0
        Dim TotalOrderQTYNew As Double = 0
        For Each rowSOTRUSSE As DataRow In dst.Tables("SOTRUSSE").Select()
            TotalOrderPriceOrig += Val(rowSOTRUSSE.Item("ORDR_UNIT_PRICE")) * Val(rowSOTRUSSE.Item("ORDR_QTY"))
            TotalOrderPriceNew += Val(rowSOTRUSSE.Item("NEW_ORDR_UNIT_PRICE")) * Val(rowSOTRUSSE.Item("NEW_QTY"))
            TotalOrderQTYOrig += Val(rowSOTRUSSE.Item("ORDR_QTY"))
            TotalOrderQTYNew += Val(rowSOTRUSSE.Item("NEW_QTY"))
            If Val(rowSOTRUSSE.Item("ORDR_UNIT_PRICE")) <> Val(rowSOTRUSSE.Item("NEW_ORDR_UNIT_PRICE")) Then
                SQLQ.Length = 0
                SQLQ.AppendLine(String.Format("UPDATE SOTORDR2 SET ORDR_UNIT_PRICE = {0}", Val(rowSOTRUSSE.Item("NEW_ORDR_UNIT_PRICE"))))
                SQLQ.AppendLine(",ORDR_UNIT_PRICE_CURR = " & Val(rowSOTRUSSE.Item("NEW_ORDR_UNIT_PRICE")))
                SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", Absx1.txtFor("ORDR_NO").Text))
                SQLQ.AppendLine(String.Format(" AND ORDR_LNO = {0}", rowSOTRUSSE.Item("ORDR_LNO")))
                ASCMAIN1.sql = SQLQ.ToString
                ASCDATA1.ExecuteSQL()
            End If
            If Val(rowSOTRUSSE.Item("NEW_QTY")) = 0 Then
                SQLQ.Length = 0
                SQLQ.AppendLine(String.Format("UPDATE SOTORDR2 SET ORDR_QTY_CANC = {0}", Val(rowSOTRUSSE.Item("ORDR_QTY"))))
                SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", Absx1.txtFor("ORDR_NO").Text))
                SQLQ.AppendLine(String.Format(" AND ORDR_LNO = {0}", rowSOTRUSSE.Item("ORDR_LNO")))
                ASCMAIN1.sql = SQLQ.ToString
                ASCDATA1.ExecuteSQL()
            End If
        Next
        If (TotalOrderPriceOrig <> TotalOrderPriceNew) Or (TotalOrderQTYOrig <> TotalOrderQTYNew) Then
            SQLQ.Length = 0
            SQLQ.AppendLine("UPDATE SOTORDR0")
            SQLQ.AppendLine(String.Format(" SET ORDR_AMT = {0},", TotalOrderPriceNew))
            SQLQ.AppendLine(String.Format(" ORDR_AMT_OPEN = {0},", TotalOrderPriceNew))
            SQLQ.AppendLine(String.Format(" ORDR_QTY = {0},", TotalOrderQTYNew))
            SQLQ.AppendLine(String.Format(" ORDR_QTY_OPEN = {0}", TotalOrderQTYNew))
            SQLQ.AppendLine(" WHERE ORDR_GROUP_NO IN")
            SQLQ.AppendLine(" (SELECT ORDR_GROUP_NO")
            SQLQ.AppendLine(" FROM SOTORDR1")
            SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}')", Absx1.txtFor("ORDR_NO").Text))
            ASCMAIN1.sql = SQLQ.ToString
            ASCDATA1.ExecuteSQL()
            'TODO: Update Range Price if Need be.
            'Still To Be Tested.  Make Sure to remove Update Check.
            'SQLQ.Length = 0
            'SQLQ.AppendLine("")
            'ASCMAIN1.sql = SQLQ.ToString
            'ASCDATA1.ExecuteSQL()
        End If

        For Each rowSOTRUSSE As DataRow In dst.Tables("SOTRUSSE").Select()
            If Val(rowSOTRUSSE.Item("COLOR_CODE")) <> Val(rowSOTRUSSE.Item("NEW_COLOR_CODE")) Then

                SQLQ.Length = 0
                SQLQ.AppendLine(String.Format("UPDATE ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - {0}", Val(rowSOTRUSSE.Item("NEW_QTY"))))
                SQLQ.AppendLine(String.Format(" WHERE WHSE_CODE = '{0}'", WHSE_CODE))
                SQLQ.AppendLine(String.Format(" AND STYLE_CODE = '{0}'", rowSOTRUSSE.Item("STYLE_CODE")))
                SQLQ.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", rowSOTRUSSE.Item("COLOR_CODE")))
                ASCMAIN1.sql = SQLQ.ToString
                ASCDATA1.ExecuteSQL()

                Dim SQLS As New System.Text.StringBuilder
                SQLS.Length = 0
                SQLS.AppendLine("SELECT COUNT(*) AS REC_CNT")
                SQLS.AppendLine(" FROM ICTSTAT2")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", rowSOTRUSSE.Item("STYLE_CODE")))
                SQLS.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", rowSOTRUSSE.Item("COLOR_CODE")))
                SQLS.AppendLine(String.Format(" AND WHSE_CODE = '{0}'", WHSE_CODE))
                Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                If REC_CNT = 1 Then
                    SQLQ.Length = 0
                    SQLQ.AppendLine(String.Format("UPDATE ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) + {0}", Val(rowSOTRUSSE.Item("NEW_QTY"))))
                    SQLQ.AppendLine(String.Format(" WHERE WHSE_CODE = '{0}'", WHSE_CODE))
                    SQLQ.AppendLine(String.Format(" AND STYLE_CODE = '{0}'", rowSOTRUSSE.Item("STYLE_CODE")))
                    SQLQ.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", rowSOTRUSSE.Item("COLOR_CODE")))
                    ASCMAIN1.sql = SQLQ.ToString
                    ASCDATA1.ExecuteSQL()
                Else
                    SQLQ.Length = 0
                    SQLQ.AppendLine("INSERT INTO ICTSTAT2")
                    SQLQ.AppendLine(" (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_OPEN)")
                    SQLQ.AppendLine(" VALUES")
                    SQLQ.AppendLine(String.Format(" ('{0}', '{1}', '{2}', 1)", rowSOTRUSSE.Item("STYLE_CODE"), rowSOTRUSSE.Item("COLOR_CODE"), WHSE_CODE))
                    ASCMAIN1.sql = SQLQ.ToString
                    ASCDATA1.ExecuteSQL()
                End If

                SQLQ.Length = 0
                SQLQ.AppendLine(String.Format("UPDATE SOTORDR2 SET COLOR_CODE = '{0}'", Val(rowSOTRUSSE.Item("NEW_COLOR_CODE"))))
                SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", rowSOTRUSSE.Item("ORDR_NO")))
                SQLQ.AppendLine(String.Format(" AND ORDR_LNO = {0}", rowSOTRUSSE.Item("ORDR_LNO")))
                ASCMAIN1.sql = SQLQ.ToString
                ASCDATA1.ExecuteSQL()
            End If
        Next
        If chkCreateRanges.Checked = True Then
            Dim LNO As Integer = 0
            For Each NewRange As String In NewRanges
                LNO += 1
                Dim RANGE_STYLE_LNO As Integer = LNO
                Dim RANGE_INNER_PACK_QTY As Integer = 1
                Dim RANGE_STYLE_UOM As String = "EA"
                Dim RANGE_STYLE_PRICE As Double
                Dim EDI_DOC_SEQ_NO_NEW As String = ""
                Dim EDI_DTL_SEQ As String = ""
                Dim RANGE_STYLE_DESC As String = ""
                Dim ORDR_NO As String = Absx1.txtFor("ORDR_NO").Text
                Dim RANGE_STYLE_CODE As String = NewRange
                Dim RANGE_STYLE_QTY As Integer = 0
                Dim RANGE_STYLE_VALUE As Double = 0
                Dim RANGE_STYLE_PP_PRICE As Double = 0
                Dim RANGE_STYLE_QTY_PER_PP As Integer = 0
                Dim RANGE_STYLE_PP_QTY As Integer = 0
                For Each rowSOTRUSSE As DataRow In dst.Tables("SOTRUSSE").Select(String.Format("RANGE_STYLE_CODE = '{0}'", NewRange))
                    RANGE_STYLE_QTY += Val(rowSOTRUSSE.Item("NEW_QTY"))
                    RANGE_STYLE_VALUE += Val(rowSOTRUSSE.Item("NEW_QTY")) * Val(rowSOTRUSSE.Item("NEW_ORDR_UNIT_PRICE"))
                    RANGE_STYLE_QTY_PER_PP += Val(rowSOTRUSSE.Item("RANGE_STYLE_QTY_PP"))
                    SQLQ.Length = 0
                    SQLQ.AppendLine("UPDATE SOTORDR2")
                    SQLQ.AppendLine(String.Format(" SET RANGE_STYLE_CODE = '{0}',", NewRange))
                    SQLQ.AppendLine(String.Format(" RANGE_STYLE_LNO = {0}", LNO))
                    SQLQ.AppendLine(String.Format(" WHERE ORDR_NO = '{0}'", Absx1.txtFor("ORDR_NO").Text))
                    SQLQ.AppendLine(String.Format(" AND ORDR_LNO = {0}", rowSOTRUSSE.Item("ORDR_LNO")))
                    ASCMAIN1.sql = SQLQ.ToString
                    ASCDATA1.ExecuteSQL()
                Next
                If RANGE_STYLE_QTY_PER_PP = 0 Then
                    RANGE_STYLE_PP_QTY = 0
                Else
                    RANGE_STYLE_PP_QTY = RANGE_STYLE_QTY / RANGE_STYLE_QTY_PER_PP
                End If

                If RANGE_STYLE_QTY = 0 Then
                    RANGE_STYLE_PRICE = 0
                    RANGE_STYLE_PP_PRICE = 0
                Else
                    RANGE_STYLE_PRICE = RANGE_STYLE_VALUE / RANGE_STYLE_QTY
                    If RANGE_STYLE_PP_QTY = 0 Then
                        RANGE_STYLE_PP_PRICE = 0
                    Else
                        RANGE_STYLE_PP_PRICE = RANGE_STYLE_VALUE / RANGE_STYLE_PP_QTY
                    End If
                End If

                SQLQ.Length = 0
                SQLQ.AppendLine("INSERT INTO SOTORDR9 (")
                SQLQ.AppendLine(" ORDR_NO,")
                SQLQ.AppendLine(" RANGE_STYLE_LNO,")
                SQLQ.AppendLine(" RANGE_STYLE_CODE,")
                SQLQ.AppendLine(" RANGE_STYLE_QTY,")
                SQLQ.AppendLine(" RANGE_STYLE_PRICE,")
                SQLQ.AppendLine(" RANGE_INNER_PACK_QTY,")
                SQLQ.AppendLine(" RANGE_STYLE_DESC,")
                SQLQ.AppendLine(" RANGE_STYLE_UOM,")
                SQLQ.AppendLine(" RANGE_STYLE_PP_QTY,")
                SQLQ.AppendLine(" RANGE_STYLE_PP_PRICE,")
                SQLQ.AppendLine(" EDI_DOC_SEQ_NO,")
                SQLQ.AppendLine(" EDI_DTL_SEQ,")
                SQLQ.AppendLine(" RANGE_STYLE_QTY_PER_PP,")
                SQLQ.AppendLine(" RANGE_STYLE_PRICE_CURR,")
                SQLQ.AppendLine(" RANGE_STYLE_PP_PRICE_CURR)")
                SQLQ.AppendLine(" VALUES")
                SQLQ.AppendLine(" (")
                SQLQ.AppendLine(String.Format("'{0}',", ORDR_NO))
                SQLQ.AppendLine(String.Format("{0},", RANGE_STYLE_LNO))
                SQLQ.AppendLine(String.Format("'{0}',", RANGE_STYLE_CODE))
                SQLQ.AppendLine(String.Format("{0},", RANGE_STYLE_QTY))
                SQLQ.AppendLine(String.Format("{0},", RANGE_STYLE_PRICE))
                SQLQ.AppendLine(String.Format("{0},", RANGE_INNER_PACK_QTY))
                SQLQ.AppendLine(String.Format("'{0}',", RANGE_STYLE_DESC))
                SQLQ.AppendLine(String.Format("'{0}',", RANGE_STYLE_UOM))
                SQLQ.AppendLine(String.Format("{0},", RANGE_STYLE_PP_QTY))
                SQLQ.AppendLine(String.Format("{0},", RANGE_STYLE_PP_PRICE))
                SQLQ.AppendLine(String.Format("'{0}',", EDI_DOC_SEQ_NO_NEW))
                SQLQ.AppendLine(String.Format("{0},", Val(EDI_DTL_SEQ & "")))
                SQLQ.AppendLine(String.Format("{0},", RANGE_STYLE_QTY_PER_PP))
                SQLQ.AppendLine(String.Format("{0},", RANGE_STYLE_PRICE))
                SQLQ.AppendLine(String.Format("{0})", RANGE_STYLE_PP_PRICE))
                ASCMAIN1.sql = SQLQ.ToString
                ASCDATA1.ExecuteSQL()
            Next

        End If
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        'Print_Report_Begin()
        'Generate_Report("SORORDRO")
        'Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDRX, "SSB", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTRUSSE, "SSB", "Show Filter", "Show GroupBox", "Make Range From Selected")
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
            Case "grdSOTRUSSE"
                If chkIsRANGE.Checked Then
                    e.Tool.ToolbarsManager.Tools("Make Range From Selected").SharedProps.Visible = False
                Else
                    e.Tool.ToolbarsManager.Tools("Make Range From Selected").SharedProps.Visible = True
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        'Select Case e.SourceControl.Name
        '    'Case "grdSOTORDR1"
        '    '    If Not InquiryOnly Then
        '    '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
        '    '    End If
        'End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            'Case "Make Range From Selected"
            '    If Not chkIsRANGE.Checked Then
            '        Dim MultiRangeFound As Boolean = False
            '        Dim BlankRangeFound As Boolean = False
            '        If grdSOTRUSSE.Selected.Rows.Count < 2 Then
            '            MsgBox("You Must Select At Least Two Lines To Create A New Range From", MsgBoxStyle.Critical, "Row Selection")
            '            Exit Sub
            '        End If
            '        Dim LastRange As String = ""
            '        For Each grow As UltraWinGrid.UltraGridRow In grdSOTRUSSE.Selected.Rows()
            '            If Not IsNothing(grow) Then
            '                If LastRange = "" Then
            '                    LastRange = grow.Cells.Item("RANGE_STYLE_CODE").Text
            '                End If
            '                If LastRange <> grow.Cells.Item("RANGE_STYLE_CODE").Text Then
            '                    MultiRangeFound = True
            '                End If
            '                If grow.Cells.Item("RANGE_STYLE_CODE").Text.Length = 0 Then
            '                    BlankRangeFound = True
            '                End If
            '            End If
            '        Next
            '        For Each grow As UltraWinGrid.UltraGridRow In grdSOTRUSSE.Selected.Rows()
            '            If Not IsNothing(grow) Then
            '                grow.Cells.Item("RANGE_STYLE_CODE").Value = txtNewRangeStyle.Text
            '                AddNewRanges = True
            '            End If
            '        Next
            '    End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Edit Ship To"
            '    If Not InquiryOnly Then
            '        MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
            '    End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                'FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "grdSOTORDRX"
    Private Sub grdSOTORDRX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDRX.DoubleClickRow
        If Not IsDBNull(e.Row.Cells("ORDR_NO").Value) And Not IsDBNull(e.Row.Cells("ORDR_CUST_PO").Value) Then
            Absx1.txtFor("ORDR_NO").Text = e.Row.Cells("ORDR_NO").Value
            Absx1.txtFor("ORDR_CUST_PO").Text = e.Row.Cells("ORDR_CUST_PO").Value
            Click_Command("Edit")
        End If
    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow
        With e.Row
            grdSOTORDRX.UpdateData()
        End With
    End Sub

#End Region

#Region "grdSOTRUSSE"


    Private Sub grdSOTRUSSE_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTRUSSE.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case "NEW_COLOR_CODE"
                Dim SQLS As New System.Text.StringBuilder
                SQLS.Length = 0
                SQLS.AppendLine("SELECT COUNT(*) AS REC_CNT")
                SQLS.AppendLine(" FROM ICTSTYC1")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", e.Cell.Row.Cells.Item("STYLE_CODE").Text))
                SQLS.AppendLine(String.Format(" AND COLOR_CODE = '{0}'", e.Cell.Text))
                ASCMAIN1.sql = SQLS.ToString()
                Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                If REC_CNT = 0 Then
                    MsgBox("Invalid Color Entered", MsgBoxStyle.OkOnly, "Color")
                    e.Cancel = True
                End If
        End Select
    End Sub
#End Region

#Region "Misc Form Controls"


    Private Sub chkCreateRanges_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkCreateRanges.CheckedChanged
        SetRangeColumnState()
    End Sub

    Private Sub chkIsRANGE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkIsRANGE.CheckedChanged
        SetCreateRangeState()
    End Sub

    Private Sub chkUpdateRange_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkUpdateRange.CheckedChanged
        Dim RANGE_STYLE_CODE As String = ""
        Dim ManyRanges As Boolean = False
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTRUSSE.Rows()
            If Not IsNothing(grow) Then
                If RANGE_STYLE_CODE.Length = 0 Then
                    RANGE_STYLE_CODE = grow.Cells.Item("RANGE_STYLE_CODE").Text
                Else
                    If RANGE_STYLE_CODE <> grow.Cells.Item("RANGE_STYLE_CODE").Text Then
                        ManyRanges = True
                    End If
                End If

            End If
        Next
        If ManyRanges Then
            MsgBox("There Is More Than One Range On This Order", MsgBoxStyle.Critical, "No Range Updates Allowed")
        Else
            If chkUpdateRange.Checked Then
                lblNewRange.Visible = True
                txtNewRange.Visible = True
                txtNewRange.Text = ""
            Else
                lblNewRange.Visible = False
                txtNewRange.Visible = False
                txtNewRange.Text = ""
            End If
        End If
    End Sub
#End Region

#Region "Custom Methods"

    Private Function IsDivisible(Current_Range_Total As Integer, Current_Range_PP As Integer) As Boolean
        Dim RetVal As Boolean = False
        If Current_Range_PP <> 0 And Current_Range_Total <> 0 Then
            If (Current_Range_Total Mod Current_Range_PP) = 0 Then
                RetVal = True
            End If
        End If
        Return RetVal
    End Function

    Private Sub RefreshSOTORDRX()
        Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Dim SQLW As String = ""
        ASCMAIN1.sql = String.Format("SELECT SOTORDR1.* FROM SOTORDR1 WHERE CUST_CODE = '{0}' AND ORDR_STATUS = 'O'{1}", CUST_CODE, SQLW)
        Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)
        Cursor = Cursors.Default
    End Sub

    Private Sub SetCreateRangeState()
        If chkIsRANGE.Checked Then
            chkCreateRanges.Visible = False
        Else
            chkCreateRanges.Visible = True
        End If
    End Sub

    Private Sub SetRangeColumnState()
        If chkCreateRanges.Checked Then
            grdSOTRUSSE.DisplayLayout.Bands(0).Columns("RANGE_STYLE_CODE").Hidden = False
            grdSOTRUSSE.DisplayLayout.Bands(0).Columns("RANGE_STYLE_QTY_PP").Hidden = False
        Else
            grdSOTRUSSE.DisplayLayout.Bands(0).Columns("RANGE_STYLE_CODE").Hidden = True
            grdSOTRUSSE.DisplayLayout.Bands(0).Columns("RANGE_STYLE_QTY_PP").Hidden = True
        End If
    End Sub
#End Region

#Region "Proceed Tests"
    Private Sub Proceed_CreateRange()
        Dim MISSING_RANGE_STYLE_CODE As Boolean = False
        Dim MISSING_RANGE_STYLE_QTY_PP As Boolean = False
        Dim Current_RANGE_STYLE_CODE As String = ""
        Dim Current_Range_Total As Integer
        Dim Current_Range_PP As Integer
        Dim Current_Range_Carts As Integer
        For Each rowSOTRUSSE As DataRow In dst.Tables("SOTRUSSE").Select("", "RANGE_STYLE_CODE")
            If rowSOTRUSSE.Item("RANGE_STYLE_CODE").ToString.Length = 0 Then
                MISSING_RANGE_STYLE_CODE = True
            End If
            If rowSOTRUSSE.Item("RANGE_STYLE_QTY_PP").ToString.Length = 0 Then
                MISSING_RANGE_STYLE_QTY_PP = True
            End If
            If Not rowSOTRUSSE.Item("NEW_QTY") = 0 Then
                If Not IsDivisible(rowSOTRUSSE.Item("NEW_QTY"), rowSOTRUSSE.Item("RANGE_STYLE_QTY_PP")) Then
                    EMsg &= vbCr & "Line " & rowSOTRUSSE.Item("ORDR_LNO") & " Is Not Divisble By Pack."
                End If
            End If
            If Current_RANGE_STYLE_CODE = "" Then
                Current_RANGE_STYLE_CODE = rowSOTRUSSE.Item("RANGE_STYLE_CODE").ToString
                If Val(rowSOTRUSSE.Item("RANGE_STYLE_QTY_PP") & "") <> 0 Then
                    Current_Range_Carts = Val(rowSOTRUSSE.Item("NEW_QTY") & "") / Val(rowSOTRUSSE.Item("RANGE_STYLE_QTY_PP") & "")
                End If
            End If
            If Current_RANGE_STYLE_CODE = rowSOTRUSSE.Item("RANGE_STYLE_CODE").ToString Then
                Current_Range_PP += rowSOTRUSSE.Item("RANGE_STYLE_QTY_PP")
                Current_Range_Total += rowSOTRUSSE.Item("NEW_QTY")
                If Val(rowSOTRUSSE.Item("NEW_QTY") & "") / Val(rowSOTRUSSE.Item("RANGE_STYLE_QTY_PP") & "") <> Current_Range_Carts Then
                    EMsg &= vbCr & "Range Style " & Current_RANGE_STYLE_CODE & "Has Multiple Carton Counts Based On Pack Provided."
                End If
            Else
                If Not IsDivisible(Current_Range_Total, Current_Range_PP) Then
                    EMsg &= vbCr & "Range Style " & Current_RANGE_STYLE_CODE & "Is Not Divisble By Pack."
                End If
            End If
        Next
        If MISSING_RANGE_STYLE_CODE Then
            EMsg &= vbCr & "Range Style Codes Can Not Be Blank When Creating Ranges."
            EMsg &= vbCr & "Mixing Range Styles and Non-Range Styles On The Same Order Is Not Permitted."
        End If
        If MISSING_RANGE_STYLE_QTY_PP Then
            EMsg &= vbCr & "Pack/Range Can Not Be 0 When Creating Ranges."
        End If
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("SELECT COUNT(DISTINCT(RANGE_STYLE_CODE)) AS RANGE_STYLE_CNT")
        SQLS.AppendLine(String.Format("FROM SOTORDR2 WHERE ORDR_NO = '{0}'", Absx1.txtFor("ORDR_NO").Text))
        ASCMAIN1.sql = SQLS.ToString()
        Dim RangeCount As Int16 = Val(ASCDATA1.GetDataValue)
        If RangeCount > 0 Then
            EMsg &= vbCr & "Order Already Has Range Styles On It."
        End If
    End Sub

    Private Sub Proceed_IfOrderIsOpen()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            WHSE_CODE = rowSOTORDR1.Item("WHSE_CODE")
            If rowSOTORDR1.Item("ORDR_STATUS") <> "O" Then
                EMsg &= vbCr & "Order Needs To Be In An Open Status."
            End If
        Next
    End Sub

    Private Sub Proceed_LineErrors()
        Dim ChangesFound As Boolean = False
        Dim BadUPC As Boolean = False
        Dim BadSKU As Boolean = False
        Dim BadQTY As Boolean = False
        Dim BadPrice As Boolean = False
        Dim NoRangePriceYet As Boolean = False
        Dim TotalOldQty As Double = 0
        Dim TotalNewQty As Double = 0

        For Each rowSOTRUSSE As DataRow In dst.Tables("SOTRUSSE").Select()
            If chkCreateRanges.Checked Then
                ChangesFound = True
            End If
            TotalOldQty += Val(rowSOTRUSSE.Item("ORDR_QTY"))
            TotalNewQty += Val(rowSOTRUSSE.Item("NEW_QTY"))
            If Val(rowSOTRUSSE.Item("ORDR_QTY")) <> Val(rowSOTRUSSE.Item("NEW_QTY")) Then
                ChangesFound = True
                If Not BadQTY Then
                    If rowSOTRUSSE.Item("NEW_QTY").ToString.Length <= 0 Then
                        BadQTY = True
                    End If
                End If
            End If
            If rowSOTRUSSE.Item("ORDR_UNIT_PRICE") <> rowSOTRUSSE.Item("NEW_ORDR_UNIT_PRICE") Then
                ChangesFound = True
                If Not BadPrice Then
                    If rowSOTRUSSE.Item("NEW_ORDR_UNIT_PRICE").ToString.Length <= 0 Then
                        BadPrice = True
                    End If
                End If
                If chkUpdateRange.Checked Then
                    NoRangePriceYet = True
                End If
            End If
            If rowSOTRUSSE.Item("CUST_UPC") <> rowSOTRUSSE.Item("NEW_UPC") Then
                ChangesFound = True
                If Not BadUPC Then
                    If rowSOTRUSSE.Item("NEW_UPC").ToString.Length = 0 Then
                        BadUPC = True
                    End If
                End If
            End If
            If rowSOTRUSSE.Item("CUST_SKU") <> rowSOTRUSSE.Item("NEW_SKU") Then
                ChangesFound = True
                If Not BadSKU Then
                    If rowSOTRUSSE.Item("NEW_SKU").ToString.Length = 0 Then
                        BadSKU = True
                    End If
                End If
            End If
        Next
        If Not ChangesFound Then
            EMsg &= vbCr & "It Dosn't Look Like You Made Any Changes To Update?!?"
        End If
        If NoRangePriceYet Then
            EMsg &= vbCr & "Price Changes With Ranges Not Tested Yet!"
        End If
        If BadUPC Then
            EMsg &= vbCr & "Invalid UPC Entered for Update."
        End If
        If BadSKU Then
            EMsg &= vbCr & "Invalid SKU Entered for Update."
        End If
        If BadQTY Then
            EMsg &= vbCr & "Invalid QTY Entered for Update."
        End If
        If BadPrice Then
            EMsg &= vbCr & "Invalid Price Entered for Update."
        End If
        If TotalOldQty <> TotalNewQty Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Total Qty Change"
            Dim iMSG As New System.Text.StringBuilder
            iMSG.AppendLine("This Will Alter The Total Qty On The Order.")
            iMSG.AppendLine("Is That What You Want?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult <> MsgBoxResult.Yes Then
                EMsg &= vbCr & "The Total New Quantity Does Not Equal The Original."
            End If
        End If
    End Sub

    Private Sub Proceed_UpDateRange()
        If chkUpdateRange.Checked And txtNewRange.Text.Length = 0 Then
            EMsg &= vbCr & "You Checked the Update Range But Didn't Give Me A New Range."
        End If
        If chkUpdateRange.Checked And txtNewRange.Text.Length > 0 Then
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("SELECT COUNT(*) AS RNG_CNT FROM ICTRSTY1 WHERE RANGE_STYLE_CODE = '{0}'", txtNewRange.Text))
            ASCMAIN1.sql = SQLS.ToString()
            Dim RNG_CNT As Int16 = Val(ASCDATA1.GetDataValue)
            If RNG_CNT = 0 Then
                EMsg &= vbCr & "The New Range Style You Provided Is Not Set-up In The System."
            End If
            If RNG_CNT > 0 Then
                EMsg &= vbCr & "You Can Only Update Range When There Is Only One Range On The Order."
            End If
        End If

    End Sub
#End Region

    Private Sub txtCUST_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtCUST_CODE.ValueChanged
        If FormLoaded Then
            CUST_CODE = Absx1.txtFor("CUST_CODE").Text
            RefreshSOTORDRX()
        End If
    End Sub
End Class