Imports System.Math

Public Class SOFXFER1
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Dim SQLS As New Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine("SELECT")
            SQLS.AppendLine(String.Format(" '{0}' AS TIME_STAMP,", New String(" ", 12)))
            SQLS.AppendLine(String.Format(" '{0}' AS EVENT,", New String(" ", 20)))
            SQLS.AppendLine(String.Format(" '{0}' AS MESSAGE", New String(" ", 50)))
            SQLS.AppendLine(" FROM DUAL")
            ASCMAIN1.sql = SQLS.ToString
            Create_TDA(.Tables.Add, "SOTMSGS1", "**", 0, False, "", 0)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False)
            Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)
        End With
        grdSOTMSGS1.DataSource = dst.Tables("SOTMSGS1")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Update"
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Load"
                EntryMode = "E"
            Case "Clear All Options"
                CheckOptions(False)
            Case "Select All Options"
                CheckOptions(True)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Clear All Options").Settings.Enabled = DefaultableBoolean.True
                .Groups("Style Options").Expanded = False
                .Groups("Order Options").Expanded = False
                .Groups("Partner Options").Expanded = False
                .Groups("System Options").Expanded = False
                .Groups("Parameter Options").Expanded = False
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

    End Sub

    Sub Load_Record()

    End Sub

    Sub Update_Record()

    End Sub
#End Region

    Private Sub btnService_Click(sender As System.Object, e As System.EventArgs) Handles btnService.Click
        If chkStyles.Checked Then
            MESSAGEAddStatusRemark("Styles", "Loading")
            SYNCDATA.SyncTable(Me, "ICTSTYL1")
            SYNCDATA.SyncTable(Me, "ICTSTYL3")
            SYNCDATA.SyncTable(Me, "ICTSTYLP")
            SYNCDATA.SyncTable(Me, "ICTSTYV1")
            chkStyles.Checked = False
        End If
        If chkColors.Checked Then
            MESSAGEAddStatusRemark("Colors", "Loading")
            SYNCDATA.SyncTable(Me, "ICTCOLR1")
            SYNCDATA.SyncTable(Me, "ICTSTYC1")
            chkColors.Checked = False
        End If
        If chkStatus.Checked Then
            MESSAGEAddStatusRemark("Status", "Loading")
            SYNCDATA.SyncTable(Me, "ICTSTAT1")
            SYNCDATA.SyncTable(Me, "ICTSTAT2")
            SYNCDATA.SyncTable(Me, "ICTSTAT5")
            SYNCDATA.SyncTable(Me, "ICTSTDQ1")
            chkStatus.Checked = False
        End If
        If chkClasses.Checked Then
            MESSAGEAddStatusRemark("Classes", "Loading")
            SYNCDATA.SyncTable(Me, "ICTCLAS1")
            chkClasses.Checked = False
        End If
        If chkOrders.Checked Then
            MESSAGEAddStatusRemark("Orders", "Not Enabled")
            'MESSAGEAddStatusRemark("Orders", "Loading")
            'SYNCDATA.SyncTable(Me, "SOTORDR0")
            'SYNCDATA.SyncTable(Me, "SOTORDR1")
            'SYNCDATA.SyncTable(Me, "SOTORDR2")
            'SYNCDATA.SyncTable(Me, "SOTORDR5")
            chkOrders.Checked = False
        End If
        If chkCustomers.Checked Then
            MESSAGEAddStatusRemark("Customers", "Loading")
            SYNCDATA.SyncTable(Me, "ARTCUST1")
            SYNCDATA.SyncTable(Me, "ARTCUST2")
            chkCustomers.Checked = False
        End If
        If chkVendors.Checked Then
            MESSAGEAddStatusRemark("Vendors", "Loading")
            SYNCDATA.SyncTable(Me, "APTVEND1")
            chkVendors.Checked = False
        End If
        If chkSalesReps.Checked Then
            MESSAGEAddStatusRemark("Sales Reps", "Loading")
            SYNCDATA.SyncTable(Me, "SOTSREP1")
            chkSalesReps.Checked = False
        End If
        If chkWarehouses.Checked Then
            MESSAGEAddStatusRemark("Warehouses", "Loading")
            SYNCDATA.SyncTable(Me, "ICTWHSE1")
            chkWarehouses.Checked = False
        End If
        If chkShippers.Checked Then
            MESSAGEAddStatusRemark("Shippers", "Loading")
            SYNCDATA.SyncTable(Me, "SOTSVIA1")
            chkShippers.Checked = False
        End If
        If chkUsers.Checked Then
            MESSAGEAddStatusRemark("Users", "Not Enabled")
            'MESSAGEAddStatusRemark("Users", "Loading")
            'ASCDATA1.ExecuteSQL("delete from astuser1 where user_id <> 'wayne'")
            'ASCDATA1.ExecuteSQL("delete from astuser2 where user_id <> 'wayne'")
            'ASCDATA1.ExecuteSQL("update astuser1 set sync_batch = '0000000000'")
            'ASCDATA1.ExecuteSQL("update astuser2 set sync_batch = '0000000000'")
            'SYNCDATA.SyncTable(Me, "ASTUSER1")
            'SYNCDATA.SyncTable(Me, "ASTUSER2")
            'SYNCDATA.SyncTable(Me, "ASTUSERT")
            chkUsers.Checked = False
        End If
        If chkTerms.Checked Then
            MESSAGEAddStatusRemark("Terms", "Not Enabled")
            'MESSAGEAddStatusRemark("Terms", "Loading")
            'SYNCDATA.SyncTable(Me, "TATTERM1")
            chkTerms.Checked = False
        End If
        MsgBox("Sync Complete", MsgBoxStyle.Exclamation, "Sync")
    End Sub

    Private Sub MESSAGEAddStatusRemark(ByVal EVENT_S As String, ByVal MESSAGE As String)
        Dim rowSOTMSGS1 As DataRow = dst.Tables("SOTMSGS1").NewRow
        rowSOTMSGS1.Item("TIME_STAMP") = Format(Now, "hh:mm:ss")
        rowSOTMSGS1.Item("EVENT") = EVENT_S
        rowSOTMSGS1.Item("MESSAGE") = MESSAGE
        dst.Tables("SOTMSGS1").Rows.Add(rowSOTMSGS1)
        Sort_grdColumns(grdSOTMSGS1, "TIME_STAMP".ToLower, True)
        Application.DoEvents()
    End Sub

    Private Sub CheckOptions(ChkOptions As Boolean)
        chkStyles.Checked = ChkOptions
        chkColors.Checked = ChkOptions
        chkStatus.Checked = ChkOptions
        chkClasses.Checked = ChkOptions
        chkOrders.Checked = ChkOptions
        chkCustomers.Checked = ChkOptions
        chkVendors.Checked = ChkOptions
        chkSalesReps.Checked = ChkOptions
        chkWarehouses.Checked = ChkOptions
        chkShippers.Checked = ChkOptions
        chkUsers.Checked = ChkOptions
        chkMenu.Checked = ChkOptions
        chkOrderEntry.Checked = ChkOptions
        chkTerms.Checked = ChkOptions
    End Sub

End Class