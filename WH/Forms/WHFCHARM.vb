Imports ABSolution

Public Class WHFCHARM

    Dim selectedTable = "EDT850T2"

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        With dst

            ASCMAIN1.sql = "SELECT EDI_SKU, EDI_DTL_SEQ,EDI_TOTAL_QTY,'' CUST_COLOR_CODE, '' CUST_SIZE_CODE, '' QTY_PRINT, '' START_CARTON, '' END_CARTON FROM EDT850T2 WHERE EDI_DOC_SEQ_NO IN ( " & _
                    "SELECT EDI_DOC_SEQ_NO FROM SOTORDR1 WHERE ORDR_NO = :PARM1 AND ORDR_STATUS NOT IN ('C'))"
            Create_TDA(.Tables.Add, "EDT850T2", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT CUST_SKU EDI_SKU,ORDR_LNO EDI_DTL_SEQ,CUST_COLOR_CODE,CUST_SIZE_CODE,ORDR_QTY EDI_TOTAL_QTY,'' QTY_PRINT, '' START_CARTON, '' END_CARTON " & _
                            "FROM SOTORDR2 WHERE ORDR_NO = :PARM1 AND ORDR_STATUS NOT IN ('C')"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "V", 2)

        End With

        grdEDT850T2.DataSource = dst.Tables("EDT850T2")
        For Each tblName As String In {"EDT850T2", "SOTORDR2"}
            For Each colName As String In {"QTY_PRINT", "START_CARTON", "END_CARTON"}
                dst.Tables(tblName).Columns(colName).MaxLength = 5
            Next
        Next


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey
            Case "Load"

                If txtOrderNo.Text = "" Then
                    EMsg &= "Order No is required"
                Else
                    Dim orderStatus As String = ASCDATA1.GetDataValue("SELECT ORDR_STATUS FROM SOTORDR1 WHERE ORDR_NO=:PARM1", "V", New Object() {txtOrderNo.Text})
                    If orderStatus = "C" Then
                        EMsg &= "Selected order has been canceled"
                    End If
                End If

                If txtOrdrCustPO.Text = "" Then
                    EMsg &= "PO No is required"
                End If

                If Not {"CATHERINE", "LANEBRY", "DRESSBARN"}.Contains(txtCustCode.Text) Then
                    EMsg &= "Invalid customer (must be Catherine's, Lane Bryant, or Dress Barn)"
                End If

            Case "Done"

            Case "Print"

                grdEDT850T2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

                'Dim numCartons As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD2'")(0).Item("Value") & String.Empty)
                'Dim numSets As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD100'")(0).Item("Value") & String.Empty)

                'If numCartons <= 1 Then numCartons = 1
                'If numSets <= 1 Then numSets = 1

                'If MessageBox.Show("Do you want to print " & numSets & " set(s) for " & numCartons & " carton(s)?", "Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                'Exit Sub
                'End If
        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Load"
                If txtCustCode.Text = "DRESSBARN" Then
                    selectedTable = "SOTORDR2"
                Else
                    selectedTable = "EDT850T2"
                End If
                Load_Record()
                Mode_Settings(True)
            Case "Done"
                Clear_Record()
                Mode_Settings(False)
            Case "Print"
                ProcessLabel(LabelProcess.Print)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Email Option").Visible = ScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdEDT850T2.Visible = ScreenMode

        If ScreenMode Then
            grdEDT850T2.DisplayLayout.Bands(0).Columns("CUST_COLOR_CODE").Hidden = selectedTable <> "SOTORDR2"
            grdEDT850T2.DisplayLayout.Bands(0).Columns("CUST_SIZE_CODE").Hidden = selectedTable <> "SOTORDR2"
        End If

        If ScreenMode Then
            grdEDT850T2.DataSource = dst.Tables(selectedTable)
        Else
            Me.Clear_Record()
        End If


    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables(selectedTable).Rows.Clear()
        txtCustCode.Text = ""
        txtOrderNo.Text = ""
        txtOrdrCustPO.Text = ""
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("")

        EnforceConstraints(False)
        Fill_Records(selectedTable, txtOrderNo.Text)
        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()
            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

    Private Sub ProcessLabel(ByVal process As LabelProcess)
        ASCMAIN1.Progress("Processing...")
        Try
            Dim selectedLabel As String = If(txtCustCode.Text = "DRESSBARN", "DRESSB_CTN", "CHARM_CTN2")

            Dim customLabel As New CharmingLabel(selectedLabel)
            Dim tblShippingData As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR5 WHERE ORDR_NO=:PARM1 AND CUST_ADDR_TYPE='ST'", "tblShippingData", "V", New Object() {txtOrderNo.Text})
            customLabel.division = txtCustCode.Text
            customLabel.poNumber = txtOrdrCustPO.Text
            customLabel.tblShippingData = tblShippingData

            Dim savedFileNames As New Dictionary(Of String, String)

            Dim qtyColumn As String = "TOTAL_QTY"
            If txtCustCode.Text = "DRESSBARN" Then
                qtyColumn = "CARTON_PACK_QTY"
            End If


            For Each row As DataRow In dst.Tables(selectedTable).Select("QTY_PRINT > 0")

                Dim numCartons As Integer = Val(row.Item("QTY_PRINT"))

                Dim dtLabelData As DataTable

                If txtCustCode.Text = "DRESSBARN" Then
                    dtLabelData = ASCDATA1.GetDataTable("SELECT O1.ORDR_CUST_PO, O1.ORDR_DEPT,O2.ORDR_NO, ORDR_LNO, O2.STYLE_CODE, " & _
                                "  O2.INNER_PACK_QTY, ORDR_QTY, O2.CUST_STYLE_CODE, O2.STYLE_CLASS_CODE, O2.CUST_SKU, O2.CUST_COLOR_CODE, O2.CUST_SIZE_CODE, S1.CARTON_PACK_QTY, S1.CASE_WEIGHT_GRS, GREATEST(S1.CASE_WEIGHT_GRS-1,0) NET_WEIGHT, S1.COUNTRY_CODE, CY.COUNTRY_NAME " & _
                                "FROM SOTORDR1 O1 " & _
                                "JOIN SOTORDR2 O2      ON (O2.ORDR_NO = O1.ORDR_NO) " & _
                                "JOIN ICTSTYL1 S1      ON (O2.STYLE_CODE = S1.STYLE_CODE) " & _
                                "LEFT JOIN TATCNTRY CY ON (S1.COUNTRY_CODE = CY.COUNTRY_CODE) " & _
                                "WHERE O1.ORDR_NO = :PARM1 AND O2.ORDR_LNO=:PARM2 " & _
                                "ORDER BY O2.ORDR_NO, O2.ORDR_LNO", "dtCharming", "VN", New Object() {txtOrderNo.Text, row.Item("EDI_DTL_SEQ")})

                Else
                    dtLabelData = ASCDATA1.GetDataTable("SELECT EDI_DTL_SEQ,EDI_SLN_PO_LNO,EDI_SLN_DEPT,EDI_SLN_LINE_MODE,EDI_SLN_COLOR,SUM(EDI_SLN_QTY) TOTAL_QTY, LISTAGG(EDI_SLN_QTY,'/') WITHIN GROUP (ORDER BY EDI_SLN_SEQ) SIZE_STRING FROM EDT850T6 WHERE EDI_DOC_SEQ_NO IN ( " & _
                                                                    " SELECT EDI_DOC_SEQ_NO FROM SOTORDR1 WHERE ORDR_NO = :PARM1 AND ORDR_STATUS NOT IN ('C')) " & _
                                                                    " and EDI_DTL_SEQ = :PARM2 " & _
                                                                    " GROUP BY EDI_DTL_SEQ,EDI_SLN_PO_LNO,EDI_SLN_DEPT,EDI_SLN_LINE_MODE,EDI_SLN_COLOR", "dtCharming", "VN", New Object() {txtOrderNo.Text, row.Item("EDI_DTL_SEQ")})
                End If

                Dim savedFileName As String = ASCMAIN1.Folders("Work") & Now.ToString("yyyyMMddhhmmss") & "_" & row.Item("EDI_SKU") & ".txt"
                Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(savedFileName)
                savedFileNames.Add(FI.Name, savedFileName)

                Dim startCarton As Integer = If(Val(row.Item("START_CARTON") & "") <> 0, row.Item("START_CARTON"), 1)
                Dim endCarton As Integer = If(Val(row.Item("END_CARTON") & "") <> 0, row.Item("END_CARTON"),numCartons)

                For i As Integer = startCarton To endCarton
                    customLabel.tblLabelData = dtLabelData
                    customLabel.cartonString = String.Format("{0} of {1}", i, numCartons)

                    If txtCustCode.Text = "DRESSBARN" And i * dtLabelData.Rows(0).Item(qtyColumn) > row.Item("EDI_TOTAL_QTY") Then
                        dtLabelData.Columns(qtyColumn).ReadOnly = False
                        dtLabelData.Rows(0).Item(qtyColumn) = Math.Max(row.Item("EDI_TOTAL_QTY") - (dtLabelData.Rows(0).Item(qtyColumn) * (i - 1)), 0)
                    End If

                    If process = LabelProcess.Print Then
                        customLabel.PrintLabel()
                    Else
                        customLabel.SaveLabelToFile(savedFileName, True, False)

                    End If

                    If dtLabelData.Columns(qtyColumn).ReadOnly = False Then
                        Exit For
                    End If
                Next
            Next

            If process = LabelProcess.Email Then
                'email the saved label files
                Dim emailForm As New TAFSEND1(Me)
                emailForm.SEND_TOs.Add(txtEmailAddress.Text, txtEmailAddress.Text)
                emailForm.SEND_FROM = "labels@nyagroup.com"
                emailForm.SEND_BODY = ""
                emailForm.EMAIL_KEY = "LABEL"
                emailForm.SEND_SUBJECT = String.Format("Labels from NYAG - Cust: {0} PO: {1}", txtCustName.Text, _txtOrdrCustPO.Text)

                emailForm.SEND_ATTACHMENTs = savedFileNames
                emailForm.Send_email_automatically(False)
                emailForm.Dispose()
                MsgBox("Email has been sent")
            End If
            ASCMAIN1.Progress("")
        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message)
        End Try

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
    (ByVal ctl As Windows.Forms.Control, _
     ByVal COLUMN_NAME As String, _
     Optional ByRef sql_where As String = "", _
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "ORDR_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""


                If COLUMN_NAME = "ORDR_NO" Then
                    'If InquiryMode Then
                    'Else
                    '    sql_where &= " and SOTORDR1.ORDR_STATUS = 'O' "
                    'End If
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTORDR1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If

        End Select
    End Sub

#End Region

    Private Sub btnEmail_Click(sender As System.Object, e As System.EventArgs) Handles btnEmail.Click
        If String.IsNullOrEmpty(txtEmailAddress.Text) Then
            MsgBox("You must enter an email address")
            Exit Sub
        End If
        grdEDT850T2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

        ProcessLabel(LabelProcess.Email)
    End Sub


    Enum LabelProcess
        Print
        Email
    End Enum
End Class