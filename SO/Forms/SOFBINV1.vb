Public Class SOFBINV1
    Public ORDR_GROUP_NOs As New List(Of String)
    Public BATCH_NO As String

    Private Sub SOFBINV1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst

            Create_TDA(.Tables.Add, "SOTBINV1", "*")


            Create_TDA(.Tables.Add, "SOTBINV2", "*")
            .Tables("SOTBINV2").Columns.Add("ORDR_GROUP_NO")
            .Tables("SOTBINV2").Columns.Add("ORDR_CUST_PO")
        End With

        grdSOTBINV2.DataSource = dst.Tables("SOTBINV2")

        With grdSOTBINV2.DisplayLayout.Bands(0)
            .Override.AllowUpdate = DefaultableBoolean.True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "BILL_OF_LADING_NO" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                End If
            Next
        End With

        Create_Summary(grdSOTBINV2, "SHIP_BOL_NO", "Count")

        Create_New_Batch()

    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then

                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"

        End Select
    End Sub
#End Region

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        EMsg = ""

        If Absx1.txtFor("SHIP_REF").Text = "" Then
            EMsg &= vbCr & "Shipper's Pro No is Mandatory"
        End If

        If Not chkBOL.Checked Then
            If Absx1.txtFor("BILL_OF_LADING_NO").Text = "" Then
                EMsg &= vbCr & "Bill of Lading No is Mandatory"
            End If
        End If

        If Absx1.dteFor("SHIP_DATE_SHIPPED").Value & "" = "" Then
            EMsg &= vbCr & "Date Shipped is Mandatory"
        Else
            Dim dteSHIP_DATE_SHIPPED As Date = Absx1.dteFor("SHIP_DATE_SHIPPED").Value
            If Format(dteSHIP_DATE_SHIPPED, "yyyyMMdd") > Format(DATETIME_STAMP.AddDays(3), "yyyyMMdd") _
            Or Format(dteSHIP_DATE_SHIPPED, "yyyyMMdd") < Format(DATETIME_STAMP.AddDays(-3), "yyyyMMdd") Then
                EMsg &= vbCr & "Date Shipped may not be more than 3 days away from today"
            End If
        End If
        If Absx1.dteFor("INV_DATE").Value & "" = "" Then
            EMsg &= vbCr & "Invoice Date Mandatory"
        Else
            Dim dteINV_DATE As Date = Absx1.dteFor("INV_DATE").Value
            If Format(dteINV_DATE, "yyyyMMdd") > Format(DATETIME_STAMP.AddDays(3), "yyyyMMdd") _
            Or Format(dteINV_DATE, "yyyyMMdd") < Format(DATETIME_STAMP.AddDays(-3), "yyyyMMdd") Then
                EMsg &= vbCr & "Invoice Date may not be more than 3 days away from today"
            End If
        End If

        If Absx1.txtFor("SHIP_VIA_CODE").Text = "" Then
            EMsg &= vbCr & "Ship Via is Mandatory"
        Else
            If LookUp("SOTSVIA1", Absx1.txtFor("SHIP_VIA_CODE").Text) Is Nothing Then
                EMsg &= vbCr & "Invalid Ship Via Specified"
            End If
        End If


        For Each row As DataRow In dst.Tables("SOTBINV2").Select("")
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            Dim row2 As DataRow = ASCDATA1.GetDataRow("Select SOTBINV2.* from SOTBINV2,SOTBINV1 where SOTBINV1.BATCH_NO = SOTBINV2.BATCH_NO and STATUS = '0' and SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
            If row2 IsNot Nothing Then
                EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is already Queued in Batch " & row2.Item("BATCH_NO")
                Exit For
            End If
        Next

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If

        BeginTrans()
        Update_Record_TDA("SOTBINV1")
        Update_Record_TDA("SOTBINV2", "BATCH_NO = '" & BATCH_NO & "'")

        If chkBOL.Checked Then
            ASCMAIN1.sql = "Update SOTSHIP1 SET SHIP_REF = :PARM1, " _
                & "SHIP_DATE_SHIPPED = :PARM2, INV_DATE = :PARM3, SHIP_VIA_CODE = :PARM4 " _
                & "where SHIP_BOL_NO in (SELECT SHIP_BOL_NO from SOTBINV2 where BATCH_NO = '" & BATCH_NO & "')"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VDDV", New Object() { _
                                Absx1.txtFor("SHIP_REF").Text, _
                                Absx1.dteFor("SHIP_DATE_SHIPPED").Value, _
                                Absx1.dteFor("INV_DATE").Value, _
                                Absx1.txtFor("SHIP_VIA_CODE").Text})
        Else
            ASCMAIN1.sql = "Update SOTSHIP1 SET SHIP_REF = :PARM1, BILL_OF_LADING_NO = :PARM2, " _
                & "SHIP_DATE_SHIPPED = :PARM3, INV_DATE = :PARM4, SHIP_VIA_CODE = :PARM5 " _
                & "where SHIP_BOL_NO in (SELECT SHIP_BOL_NO from SOTBINV2 where BATCH_NO = '" & BATCH_NO & "')"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVDDV", New Object() { _
                                Absx1.txtFor("SHIP_REF").Text, _
                                Absx1.txtFor("BILL_OF_LADING_NO").Text, _
                                Absx1.dteFor("SHIP_DATE_SHIPPED").Value, _
                                Absx1.dteFor("INV_DATE").Value, _
                                Absx1.txtFor("SHIP_VIA_CODE").Text})
        End If

        CommitTrans()

        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        BATCH_NO = ""
        Me.Close()
    End Sub


    Sub Create_New_Batch()
        BATCH_NO = ASCMAIN1.Next_Control_No("SOTBINV1.BATCH_NO")
        Dim rowSOTBINV1 As DataRow = dst.Tables("SOTBINV1").NewRow

        With rowSOTBINV1
            .ITEM("BATCH_NO") = BATCH_NO

            .Item("STATUS") = "0"
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .ITEM("INIT_DATE") = DATETIME_STAMP
        End With
        dst.Tables("SOTBINV1").Rows.Add(rowSOTBINV1)

        dst.Tables("SOTBINV2").Rows.Clear()
        dst.Tables("SOTBINV2").AcceptChanges()

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCMAIN1.sql = "Select '" & BATCH_NO & "' BATCH_NO, SOTSHIP1.SHIP_BOL_NO, NULL BILL_OF_LADING_NO, SOTSHIP1.ORDR_GROUP_NO, SOTORDR0.ORDR_CUST_PO" _
                & " from SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO and SOTSHIP1.SHIP_STATUS = 'P' and SOTSHIP1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            Fill_Records("SOTBINV2", , False, ASCMAIN1.sql)
        Next

        chkBOL.Checked = True

        Dim SHIP_BOL_NO As String = dst.Tables("SOTBINV2").Rows(0).Item("SHIP_BOL_NO")
        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
        rowSOTBINV1.Item("SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF")
        rowSOTBINV1.Item("BILL_OF_LADING_NO") = rowSOTSHIP1.Item("BILL_OF_LADING_NO")
        rowSOTBINV1.Item("SHIP_DATE_SHIPPED") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
        rowSOTBINV1.Item("INV_DATE") = rowSOTSHIP1.Item("INV_DATE")
        rowSOTBINV1.Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE")


        cmdDelete.Visible = False
        For Each row As DataRow In dst.Tables("SOTBINV2").Select("")
            SHIP_BOL_NO = row.Item("SHIP_BOL_NO")
            Dim row2 As DataRow = ASCDATA1.GetDataRow("Select SOTBINV2.* from SOTBINV2,SOTBINV1 where SOTBINV1.BATCH_NO = SOTBINV2.BATCH_NO and STATUS = '0' and SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
            If row2 IsNot Nothing Then
                MsgBox("Shipment " & SHIP_BOL_NO & " is already Queued in Batch " & row2.Item("BATCH_NO"), MsgBoxStyle.OkOnly, "Cannot Create a New Batch")
                BATCH_NO = row2.Item("BATCH_NO")

                ASCMAIN1.sql = "Select * from SOTBINV1 where BATCH_NO = '" & BATCH_NO & "'"
                Fill_Records("SOTBINV1", "", True, ASCMAIN1.sql)

                ASCMAIN1.sql = "Select SOTBINV2.BATCH_NO, SOTBINV2.SHIP_BOL_NO, SOTBINV2.BILL_OF_LADING_NO, SOTSHIP1.ORDR_GROUP_NO, SOTORDR0.ORDR_CUST_PO" _
                    & " from SOTBINV2,SOTSHIP1,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                    & " and SOTSHIP1.SHIP_BOL_NO = SOTBINV2.SHIP_BOL_NO and SOTBINV2.BATCH_NO = '" & BATCH_NO & "'"
                Fill_Records("SOTBINV2", , True, ASCMAIN1.sql)

                cmdDelete.Visible = True
                cmdUpdate.Visible = False


                If dst.Tables("SOTBINV1").Rows(0).Item("BILL_OF_LADING_NO") & "" = "" Then
                    chkBOL.Checked = True
                Else
                    chkBOL.Checked = False
                End If

                Set_Read_Only(UltraGroupBox2, True)
                Exit For
            End If
        Next

    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If MsgBox("OK to Delete Billing Batch " & BATCH_NO, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
 
        Dim rowSOTBINV1 As DataRow = dst.Tables("SOTBINV1").Rows.Find(BATCH_NO)
        rowSOTBINV1.Item("STATUS") = "D"
        Update_Record_TDA("SOTBINV1")
        MsgBox("Batch " & BATCH_NO & " has been Deleted", MsgBoxStyle.OkOnly, "Verification")

        BATCH_NO = ""
        Me.Close()

    End Sub

    Private Sub chkBOL_CheckedChanged(sender As Object, e As EventArgs) Handles chkBOL.CheckedChanged
        Absx1.txtFor("BILL_OF_LADING_NO").ReadOnly = (chkBOL.Checked)
        If (chkBOL.Checked) Then Absx1.txtFor("BILL_OF_LADING_NO").Text = ""
    End Sub
End Class