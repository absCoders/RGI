Imports System.Text
Imports Infragistics.Win.UltraWinGrid

Public Class ECTECOM1
    Private SQL As New StringBuilder With {.Length = 0}
    Private EC_PARM_APICUST_IMAGES As String = ""
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowECTPARM1 As DataRow = LookUp("ECTPARM1", "Z")
        EC_PARM_APICUST_IMAGES = rowECTPARM1.Item("EC_PARM_APICUST_IMAGES").ToString & String.Empty
        If Not EC_PARM_APICUST_IMAGES.EndsWith("\") Then
            EC_PARM_APICUST_IMAGES = EC_PARM_APICUST_IMAGES & "\"
        End If

        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("EC.*,")
            SQL.AppendLine("C1.CUST_NAME")
            SQL.AppendLine("FROM ECTECOMC EC, ARTCUST1 C1")
            SQL.AppendLine("WHERE EC.CUST_CODE = C1.CUST_CODE")
            SQL.AppendLine("AND EC.ECOM_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "ECTECOMC", "**", 0, True, "V", 2)
            '.Tables("ARTCUST2").Columns.Add("LAST_VERIFIED", GetType(System.DateTime))
        End With

        grdECTECOMC.DataSource = dst.Tables("ECTECOMC")

        With grdECTECOMC.DisplayLayout
            .Override.AllowAddNew = AllowAddNew.Yes
            .Override.AllowDelete = DefaultableBoolean.True
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = Activation.NoEdit
            Next i
            For Each COLNAME As String In New String() {"API_NAME", "API_PASSWORD"}
                .Bands(0).Columns(COLNAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLNAME As String In New String() {"API_NAME", "API_PASSWORD"}
                .Bands(0).Columns(COLNAME).CellClickAction = CellClickAction.EditAndSelectText
            Next
        End With

        ASCMAIN1.Add_Value_List(grdECTECOMC, "API_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive"})
    End Sub
    Overrides Sub Show_Record_Special()
        Select Case EntryMode
            Case "New"
                SetNewDefaults()
        End Select
        If Absx1.txtFor("ECOM_CODE").Text.ToString = "APICUST" Then
            UltraTabControl1.Tabs.Item("API Customers").Visible = True
            Fill_Records("ECTECOMC", Absx1.txtFor("ECOM_CODE").Text.ToString)
        Else
            UltraTabControl1.Tabs.Item("API Customers").Visible = False
        End If
    End Sub

    Private Sub SetNewDefaults()
        Absx1.numFor("ECOM_MIN_QTY_DEFAULT").Value = 4
        Absx1.numFor("ECOM_ALLOC_PCT_DEFAULT").Value = 100
        Absx1.numFor("ECOM_SHIP_WINDOW").Value = 7
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Update"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Dim CUST_CODE_CNT As Int16 = 0
                If CUST_CODE.Length > 0 Then
                    If CUST_CODE <> "APICUST" Then
                        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                        SQLS.AppendLine(String.Format("SELECT COUNT(*) FROM ARTCUST1 WHERE CUST_CODE = '{0}'", CUST_CODE))
                        ASCMAIN1.sql = SQLS.ToString()
                        CUST_CODE_CNT = Val(ASCDATA1.GetDataValue)
                        If CUST_CODE_CNT <> 1 Then
                            EMsg &= "Invalid Value Specified for Cust Code"
                        End If
                    End If
                Else
                    EMsg &= "Invalid Value Specified for Cust Code"
                End If
                'Clear_Record_Special()
            Case "Cancel"
                Clear_Record_Special()
        End Select
    End Sub
    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ECTECOMC"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
            UltraTabControl1.Tabs.Item("API Customers").Visible = False
        End If
    End Sub

    Private Sub grdECTECOMC_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdECTECOMC.ClickCellButton
        With e.Cell.Row
            Dim sql_where As String = ""

            Select Case e.Cell.Column.Key
                Case "CUST_CODE"
                    grdClickCellButton(grdECTECOMC, sql_where)
            End Select
        End With
    End Sub

    Private Sub grdECTECOMC_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdECTECOMC.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                Dim row As DataRow = LookUp("ARTCUST1", e.Cell.Value)
                If row IsNot Nothing Then
                    grdECTECOMC.ActiveRow.Cells("CUST_NAME").Value = row.Item("CUST_NAME")
                End If
        End Select
    End Sub

    Private Sub grdECTECOMC_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdECTECOMC.BeforeRowUpdate
        If e.Row.IsAddRow Then
            Dim eMsg As New StringBuilder With {.Length = 0}
            If e.Row.Cells("CUST_CODE").Value = "" Then
                eMsg.AppendLine("Customer Required")
            End If
            If e.Row.Cells("API_STATUS").Value & String.Empty = "" Then
                eMsg.AppendLine("Status Required")
            End If
            If e.Row.Cells("API_NAME").Value & String.Empty = "" Then
                eMsg.AppendLine("API Name Required")
            End If
            If e.Row.Cells("API_PASSWORD").Value & String.Empty = "" Then
                eMsg.AppendLine("API Password Required")
            Else
                If (e.Row.Cells("API_PASSWORD").Value).ToString.Length < 10 Then
                    eMsg.AppendLine("API Password Must Be At Least 10 Char")
                End If
            End If
            If eMsg.Length > 0 Then
                MsgBox(eMsg.ToString, vbExclamation, "No Additions")
                e.Cancel = True
            Else
                e.Row.Cells("ECOM_CUST").Value = Absx1.txtFor("ECOM_CUST").Text
                e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            End If
        End If
        e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        grdECTECOMC.UpdateData()
        Dim sqlDelete = ""
        Update_Record_TDA("ECTECOMC")
    End Sub
    Private Sub grdECTECOMC_AfterRowActivate(sender As Object, e As EventArgs) Handles grdECTECOMC.AfterRowActivate
        picAPILogo.Image = Nothing
        picAPILogo.ImageLocation = ""
        picAPIMsg.Image = Nothing
        picAPIMsg.ImageLocation = ""

        'Dim ImagesFolder As String = "APICust\"

        If Not IsNothing(grdECTECOMC.ActiveRow) Then

            Dim CUST_CODE As String = grdECTECOMC.ActiveRow.Cells.Item("CUST_CODE").Text & String.Empty

            Dim ImageLogo As String = String.Format("{0}{1}_LOGO.jpg", EC_PARM_APICUST_IMAGES, CUST_CODE)
            If IO.File.Exists(ImageLogo) Then
                picAPILogo.ImageLocation = ImageLogo
                btnAPILogo_Remove.Enabled = True
                btnAPILogo_Add.Enabled = False
            Else
                btnAPILogo_Remove.Enabled = False
                btnAPILogo_Add.Enabled = True
            End If

            Dim ImageMsg As String = String.Format("{0}{1}_MSG.jpg", EC_PARM_APICUST_IMAGES, CUST_CODE)
            If IO.File.Exists(ImageMsg) Then
                picAPIMsg.ImageLocation = ImageMsg
                btnAPIMsg_Remove.Enabled = True
                btnAPIMsg_Add.Enabled = False
                lblAPIMsg.Enabled = False
                rtbAPIMsg.Text = ""
                rtbAPIMsg.Enabled = False
            Else
                btnAPIMsg_Remove.Enabled = False
                btnAPIMsg_Add.Enabled = True
                rtbAPIMsg.Text = grdECTECOMC.ActiveRow.Cells.Item("API_MSG").Text & String.Empty
                rtbAPIMsg.Enabled = True
            End If
        End If
    End Sub

    Private Sub rtbAPIMsg_TextChanged(sender As Object, e As EventArgs) Handles rtbAPIMsg.TextChanged
        If Not IsNothing(grdECTECOMC.ActiveRow) Then
            grdECTECOMC.ActiveRow.Cells.Item("API_MSG").Value = rtbAPIMsg.Text
        End If
    End Sub

    Private Sub btnAPILogo_Add_Click(sender As Object, e As EventArgs) Handles btnAPILogo_Add.Click
        If Not IsNothing(grdECTECOMC.ActiveRow) Then
            Dim CUST_CODE As String = grdECTECOMC.ActiveRow.Cells.Item("CUST_CODE").Text & String.Empty

            Dim fDialog As New OpenFileDialog
            fDialog.Filter = "Jpg Files (*.Jpg*)|*.Jpg"
            fDialog.CheckFileExists = False
            If fDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim ImageLogo As String = String.Format("{0}{1}_LOGO.jpg", EC_PARM_APICUST_IMAGES, CUST_CODE)
                If IO.File.Exists(ImageLogo) Then
                    IO.File.Delete(ImageLogo)
                End If
                IO.File.Copy(fDialog.FileName, ImageLogo)
                picAPILogo.ImageLocation = ImageLogo
                btnAPILogo_Remove.Enabled = True
                btnAPILogo_Add.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnAPILogo_Remove_Click(sender As Object, e As EventArgs) Handles btnAPILogo_Remove.Click
        If Not IsNothing(grdECTECOMC.ActiveRow) Then
            Dim CUST_CODE As String = grdECTECOMC.ActiveRow.Cells.Item("CUST_CODE").Text & String.Empty

            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Remove Logo?"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("Are You Sure?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                Dim ImageLogo As String = String.Format("{0}{1}_LOGO.jpg", EC_PARM_APICUST_IMAGES, CUST_CODE)
                If IO.File.Exists(ImageLogo) Then
                    IO.File.Delete(ImageLogo)
                End If
                picAPILogo.ImageLocation = Nothing
                btnAPILogo_Remove.Enabled = False
                btnAPILogo_Add.Enabled = True
            End If
        End If
    End Sub

    Private Sub btnAPIMsg_Add_Click(sender As Object, e As EventArgs) Handles btnAPIMsg_Add.Click
        If Not IsNothing(grdECTECOMC.ActiveRow) Then
            Dim CUST_CODE As String = grdECTECOMC.ActiveRow.Cells.Item("CUST_CODE").Text & String.Empty

            Dim fDialog As New OpenFileDialog
            fDialog.Filter = "Jpg Files (*.Jpg*)|*.Jpg"
            fDialog.CheckFileExists = False
            If fDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim ImageMsg As String = String.Format("{0}{1}_MSG.jpg", EC_PARM_APICUST_IMAGES, CUST_CODE)
                If IO.File.Exists(ImageMsg) Then
                    IO.File.Delete(ImageMsg)
                End If
                IO.File.Copy(fDialog.FileName, ImageMsg)
                picAPIMsg.ImageLocation = ImageMsg
                btnAPIMsg_Remove.Enabled = True
                btnAPIMsg_Add.Enabled = False
                lblAPIMsg.Enabled = False
                rtbAPIMsg.Enabled = False
                rtbAPIMsg.Text = ""
            End If
        End If
    End Sub

    Private Sub btnAPIMsg_Remove_Click(sender As Object, e As EventArgs) Handles btnAPIMsg_Remove.Click
        If Not IsNothing(grdECTECOMC.ActiveRow) Then
            Dim CUST_CODE As String = grdECTECOMC.ActiveRow.Cells.Item("CUST_CODE").Text & String.Empty

            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Remove Message Image?"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("Are You Sure?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                Dim ImageMSG As String = String.Format("{0}{1}_MSG.jpg", EC_PARM_APICUST_IMAGES, CUST_CODE)
                If IO.File.Exists(ImageMSG) Then
                    IO.File.Delete(ImageMSG)
                End If
                picAPIMsg.ImageLocation = Nothing
                btnAPIMsg_Remove.Enabled = False
                btnAPIMsg_Add.Enabled = True
                lblAPIMsg.Enabled = True
                rtbAPIMsg.Enabled = True
                If Not IsNothing(grdECTECOMC.ActiveRow) Then
                    rtbAPIMsg.Text = grdECTECOMC.ActiveRow.Cells.Item("API_MSG").Value
                End If
            End If
        End If
    End Sub
End Class