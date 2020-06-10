Imports System.Text
Imports Infragistics.Win.UltraWinGrid

Public Class ECTECOM1
    Private SQL As New StringBuilder With {.Length = 0}

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
                Clear_Record_Special()
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
End Class