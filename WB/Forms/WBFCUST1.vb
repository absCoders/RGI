Imports System.Net.Mail
Imports System.Globalization
Imports Infragistics.Win.UltraWinGrid
Imports System.IO
Imports System.Text
Imports WB

Public Class WBFCUST1
    Dim InquiryOnly As Boolean = False
    Dim IncludeFilter As String = "AND"
    Dim sqlARTCONTX As String = ""
    Dim EMAIL_NAME As String = ""
    Dim EMAIL_ADDRESS As String = ""
    'Dim WithEvents FtpShopSite As New nsoftware.IPWorks.Ftp
    Private Enum EncryptType
        Encrypt
        Decrypt
    End Enum

    Private Enum LockType
        Lock
        Release
    End Enum

    Private Enum LockDirection
        Import
        Export
    End Enum

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Dim inBoundFile As String = "customers.csv"
    Dim OutBoundFile As String = "accounts.csv"
    'Dim UserName As String = "regdemo"
    'Dim Password As String = "Sa78bt5R"
    'Dim RemoteHost As String = "regdemo.regency-rib.com"
    'Dim RemotePath As String = "www/customers"
    Dim UserName As String = "regency-rib"
    Dim Password As String = "joydHUJ3"
    Dim RemoteHost As String = "regency-rib.com" '69.39.227.201
    Dim RemotePath As String = "www/customers"
    Dim ServerFilePath As String = "S:\RGI\Archive\Shopsite\"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New StringBuilder() With {.Length = 0}
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM ARTCUST1")
            SQLs.AppendLine("WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, True, "V", 1)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM ARTCUST2")
            SQLs.AppendLine("WHERE CUST_CODE = :PARM1")
            SQLs.AppendLine("AND CUST_ADDR_TYPE = :PARM2")
            SQLs.AppendLine("AND CUST_ADDR_CODE = :PARM3")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, True, "VVV", 3)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM ARTCUSTD")
            SQLs.AppendLine("WHERE CUST_CODE = :PARM1")
            SQLs.AppendLine("AND CONTACT_NO = :PARM2")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "VI", 3)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTCUST1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTCUST1", "**", 0, True, , 1)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTCUST2")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTCUST2", "**", 0, True, , 2)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT *")
            SQLs.AppendLine("FROM WBTCUST2")
            SQLs.AppendLine("WHERE EMAIL IN")
            SQLs.AppendLine("(")
            SQLs.AppendLine("SELECT EMAIL")
            SQLs.AppendLine("FROM WBTCUST2")
            SQLs.AppendLine("GROUP BY EMAIL")
            SQLs.AppendLine("HAVING COUNT(WEB_CUST_BATCH) = 1")
            SQLs.AppendLine(")")
            SQLs.AppendLine("AND WEB_CUST_BATCH = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTCUSTC", "**", 0, False, "V")

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTCUST3")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTCUST3", "**", 0, True, , 2)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTCUST9")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTCUST9", "**", 0, True, , 1)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM (SELECT 'Customer' AS REC_TYPE, CUST_CODE, CUST_CONTACT,")
            SQLs.AppendLine("CUST_NAME, CUST_ADDR1, CUST_ADDR2, CUST_CITY, CUST_STATE,")
            SQLs.AppendLine("CUST_ZIP_CODE, CUST_PHONE, CUST_EMAIL, 0 AS CONTACT_NO")
            SQLs.AppendLine("FROM ARTCUST1")
            SQLs.AppendLine("UNION")
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("'Ship-To' AS REC_TYPE, CUST_CODE, CUST_CONTACT,")
            SQLs.AppendLine("CUST_NAME, CUST_ADDR1, CUST_ADDR2, CUST_CITY, CUST_STATE,")
            SQLs.AppendLine("CUST_ZIP_CODE, CUST_PHONE, CUST_EMAIL, 0 AS CONTACT_NO")
            SQLs.AppendLine("FROM ARTCUST2")
            SQLs.AppendLine("UNION")
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("'Contact' AS REC_TYPE, D.CUST_CODE,")
            SQLs.AppendLine("D.CONTACT_NAME AS CUST_CONTACT,")
            SQLs.AppendLine("C.CUST_NAME,")
            SQLs.AppendLine("C.CUST_ADDR1,")
            SQLs.AppendLine("C.CUST_ADDR2,")
            SQLs.AppendLine("C.CUST_CITY,")
            SQLs.AppendLine("C.CUST_STATE,")
            SQLs.AppendLine("C.CUST_ZIP_CODE,")
            SQLs.AppendLine("D.CONTACT_PHONE AS CUST_PHONE,")
            SQLs.AppendLine("D.CONTACT_EMAIL AS CUST_EMAIL,")
            SQLs.AppendLine("D.CONTACT_NO")
            SQLs.AppendLine("FROM ARTCUSTD D, ARTCUST1 C")
            SQLs.AppendLine("WHERE D.CUST_CODE = C.CUST_CODE)")
            ASCMAIN1.sql = SQLs.ToString()
            sqlARTCONTX = SQLs.ToString()
            Create_TDA(.Tables.Add, "ARTCONTX", "**", 0, False)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("AD1.KEY_VALUE AS CUST_CODE,")
            SQLs.AppendLine("AC1.CUST_NAME,")
            SQLs.AppendLine("WC1.GIVENNAME FIRST_NAME,")
            SQLs.AppendLine("WC1.FAMILYNAME LAST_NAME,")
            SQLs.AppendLine("WC1.FULLNAME,")
            SQLs.AppendLine("WC1.EMAIL,")
            SQLs.AppendLine("MAX(AD1.INIT_DATE) AS DATE_CHANGED")
            SQLs.AppendLine("FROM ASTAUDT1 AD1, ARTCUST1 AC1, WBTCUST1 WC1")
            SQLs.AppendLine("WHERE AD1.KEY_VALUE = AC1.CUST_CODE")
            SQLs.AppendLine("AND  AD1.KEY_VALUE = WC1.CUST_CODE_ACTUAL")
            SQLs.AppendLine("AND AD1.TABLE_NAME = 'ARTCUST1'")
            SQLs.AppendLine("AND NVL(AD1.OLD_VALUE,'NULL') <> 'NULL'")
            SQLs.AppendLine("AND AD1.COLUMN_NAME IN")
            SQLs.AppendLine("(")
            SQLs.AppendLine("'CUST_PRICE_TIER',")
            SQLs.AppendLine("'CUST_DISC_PCT_EXTRA',")
            SQLs.AppendLine("'CUST_DISC_PCT',")
            SQLs.AppendLine("'CUST_PRICE_TIER_PVC'")
            SQLs.AppendLine(")")
            SQLs.AppendLine("GROUP BY")
            SQLs.AppendLine("AD1.KEY_VALUE,")
            SQLs.AppendLine("AC1.CUST_NAME,")
            SQLs.AppendLine("WC1.GIVENNAME,")
            SQLs.AppendLine("WC1.FAMILYNAME,")
            SQLs.AppendLine("WC1.FULLNAME,")
            SQLs.AppendLine("WC1.EMAIL")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTCUSTP", "**", 0, False)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("WC1.EMAIL,")
            SQLs.AppendLine("WC1.STATUS,")
            SQLs.AppendLine("WC1.GIVENNAME AS FIRST_NAME,")
            SQLs.AppendLine("WC1.FAMILYNAME AS LAST_NAME,")
            SQLs.AppendLine("WC1.COMPANY,")
            SQLs.AppendLine("WC1.DATEREGISTERED,")
            SQLs.AppendLine("WC1.LAST_OPER,")
            SQLs.AppendLine("WC1.LAST_DATE")
            SQLs.AppendLine("FROM WBTCUST1 WC1")
            SQLs.AppendLine("WHERE STATUS IN ('D','R')")
            SQLs.AppendLine("AND NVL(WC1.LAST_DATE,'01-JAN-1900') <> '01-JAN-1900'")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTCUSTR", "**", 0, False)
        End With

        grdWBTCUST1.DataSource = dst.Tables("WBTCUST1")
        grdARTCUSTX.DataSource = dst.Tables("ARTCONTX")
        grdWBTCUSTP.DataSource = dst.Tables("WBTCUSTP")
        grdWBTCUSTR.DataSource = dst.Tables("WBTCUSTR")

        Create_Summary(grdWBTCUST1, "STATUS", "Count", "", "###,##0")
        Create_Summary(grdARTCUSTX, "REC_TYPE", "Count", "", "###,##0")
        Create_Summary(grdWBTCUSTP, "EMAIL", "Count", "", "###,##0")
        Create_Summary(grdWBTCUSTR, "EMAIL", "Count", "", "###,##0")

        ASCMAIN1.Add_Value_List(grdWBTCUST1, "STATUS", , New String() {":", "N:New", "M:Matched", "C:Awaiting Credit", "A:Accepted", "U:Awaiting Upload", "D:Disabled", "R:Rejected"})
        ASCMAIN1.Add_Value_List(grdWBTCUSTR, "STATUS", , New String() {":", "D:Disabled", "R:Rejected"})

        Fill_Records("WBTCUST1")
        Fill_Records("WBTCUST2")
        Fill_Records("WBTCUST3")
        Fill_Records("WBTCUST9")
        Fill_Records("ARTCONTX")
        Fill_Records("WBTCUSTP")
        Fill_Records("WBTCUSTR")

        FilterWBTCUST1()

        Sort_grdColumns(grdWBTCUSTP, "DATE_CHANGED".ToLower, False)
        Sort_grdColumns(grdWBTCUSTR, "LAST_DATE".ToLower, False)

        CheckForCustomers()

        If (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "site.admin" Or ASCMAIN1.USER_ID = "mariog") Then
            chkExportTesting.Visible = True
        Else
            chkExportTesting.Visible = False
        End If

        tab.Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        Dim iMSG As New StringBuilder With {.Length = 0}
        Dim iResult As MsgBoxResult
        Dim iTitle As String = ""
        EMsg = ""

        Select Case eItemKey

            Case "Done"
                Mode_Settings(False)

            Case "Update"

            Case "Cancel"

            Case "Refresh"

            Case "Load Records"

            Case "Import Customers"
                If LOCK_IMPORT_EXPORT(LockType.Lock, LockDirection.Import) Then
                    iTitle = "Import Customers"
                    iMSG.Length = 0
                    iMSG.AppendLine("This Will Attempt To Connect To")
                    iMSG.AppendLine("Shopsite To Download New Customers.")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Are You Ready?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg &= vbCr & "Import Cancelled"
                        LOCK_IMPORT_EXPORT(LockType.Release, LockDirection.Import)
                    End If
                End If
            Case "Send Customers"
                If LOCK_IMPORT_EXPORT(LockType.Lock, LockDirection.Export) Then
                    iTitle = "Send Customers"
                    iMSG.Length = 0
                    If chkExportTesting.Checked = False Then
                        iMSG.AppendLine("This Will Attempt To Connect To Shopsite")
                        iMSG.AppendLine("To Send Customers Awaiting Upload.")
                        iMSG.AppendLine("")
                        iMSG.AppendLine("Are You Ready?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult <> MsgBoxResult.Yes Then
                            EMsg &= vbCr & "Send Customers"
                            LOCK_IMPORT_EXPORT(LockType.Release, LockDirection.Export)
                        End If
                    Else
                        iMSG.AppendLine("Because You Have Selected Testing Only")
                        iMSG.AppendLine("This Feature Will Generate The Send Customer")
                        iMSG.AppendLine("File Only For You To Upload To ShopSite")
                        iMSG.AppendLine("Manually.")
                        iMSG.AppendLine("")
                        iMSG.AppendLine("Are You Ready?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult <> MsgBoxResult.Yes Then
                            EMsg &= vbCr & "Send Customers"
                            LOCK_IMPORT_EXPORT(LockType.Release, LockDirection.Export)
                        End If
                    End If
                End If
            Case "Refresh Shopsite"
                If ASCMAIN1.USER_ID = "wayne" Then
                    iTitle = "Refresh Shopsite"
                    iMSG.Length = 0
                    iMSG.AppendLine("This Will Attempt To Connect To")
                    iMSG.AppendLine("Shopsite To Refresh Discounts")
                    iMSG.AppendLine("And Terms On All Existing Customers.")
                    iMSG.AppendLine("Are You Ready?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.No Then
                        EMsg &= vbCr & "Refresh Shopsite"
                    End If
                Else
                    MsgBox("Sorry. You Are Not Authorized For This Feature", MsgBoxStyle.Critical, "Danger!")
                    Exit Sub
                End If
            Case "Encrypt"
                If (ASCMAIN1.USER_ID <> "wayne") Then
                    EMsg &= vbCr & "Option Only Available For Wayne."
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Done"
                'Call Update_Record()
                UpdateAndRefreshData(True)
                Call Mode_Settings(False)
                Close()
                'UpdateAndRefreshData()
            Case "Update"
                'Call Update_Record()
                UpdateAndRefreshData()
                Call Mode_Settings(False)
                'UpdateAndRefreshData()
            Case "Refresh"
                UpdateAndRefreshData()
            Case "Import Customers"
                'ShopsiteFTP("R", False)
                UpdateAndRefreshData(True)
                Dim Success As Boolean = ImportCustomersFromShopsite()
                If Success Then
                    Print_Record()
                    MsgBox("Import Customers Complete", vbOKOnly, "Import Customers")
                End If
            Case "Send Customers"
                If chkExportTesting.Checked = False Then
                    Dim Success As Boolean = ExportCustomersToShopsite(False)
                    If Success Then
                        MsgBox("Customer Upload Complete", vbOKOnly, "Send Customers")
                    End If
                Else
                    Dim TempFolder As String = ASCMAIN1.Folders("Temp").ToString
                    If Not TempFolder.EndsWith("\") Then
                        TempFolder = TempFolder & "\"
                    End If
                    Dim LocalFile As String = String.Format("{0}{1}", TempFolder, OutBoundFile)

                    Dim Success As Boolean = ExportCustomersToShopsiteTesting(False)
                    If Success Then
                        MsgBox(LocalFile, vbOKOnly, "Test File Created")
                    End If
                End If
            Case "Refresh Shopsite"
                Dim Success As Boolean = ExportCustomersToShopsite(True)
                If Success Then
                    MsgBox("Customer Upload Complete", vbOKOnly, "Refresh Shopsite")
                End If
            Case "View Folder"
                Dim cmd As String = String.Format("explorer {0}", ServerFilePath)
                Shell(cmd, AppWinStyle.NormalFocus)
            Case "Encrypt"
                EncryptPasswords()
            Case "Print Last Batch"
                Print_Record()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Update").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Refresh").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Import Customers").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Print Last Batch").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Send Customers").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Refresh Shopsite").Visible = Not ScreenMode
                .Groups("Screen Control").Items("View Folder").Visible = Not ScreenMode
                If (ASCMAIN1.USER_ID = "wayne") Then
                    .Groups("Screen Control").Items("Encrypt").Visible = True
                Else
                    .Groups("Screen Control").Items("Encrypt").Visible = False
                End If

            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        With grdWBTCUST1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdWBTCUST1.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBTCUST1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        Dim editColumns As String() = New String() {"CUST_CODE_PROVIDED", "DATEREGISTERED", "FAMILYNAME", "GIVENNAME", "GROUPNAME", "GROUPNOTE", "COMPANY", "STREET", "CITY", "STATE", "ZIP_CODE", "TELEPHONE", "TAX_ID", "SREP_CODE"}
        For Each COLNAME As String In editColumns
            grdWBTCUST1.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        Next
        For Each COLNAME As String In editColumns
            grdWBTCUST1.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Next
        With grdWBTCUST1.DisplayLayout.Bands(0)
            For Each COL_NAME As String In New String() {"EMAIL", "GIVENNAME", "FAMILYNAME", "CLAIM_BY_OPER"}
                .Columns(COL_NAME).Header.Fixed = True
            Next
        End With
    End Sub

    Sub Clear_Record()
        'dst.Tables("ICTSTYL1").Rows.Clear()
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)

        Dim SQLS As New StringBuilder() With {.Length = 0}
        SQLS.AppendLine("UPDATE WBTCUST1 Set EMAIL = UPPER(EMAIL)")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        EnforceConstraints(False)

        Call Fill_Records("WBTCUST1")
        Call Fill_Records("WBTCUST2")
        Call Fill_Records("WBTCUST3")
        Call Fill_Records("WBTCUST9")
        Call Fill_Records("ARTCONTX")

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
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

    Sub Update_Record(Optional ByVal Silent As Boolean = False)
        Dim Msg As String = ""
        If Not Silent Then
            Msg = "Update Complete"
        End If
        Call BeginTrans()

        Call Update_Record_TDA("WBTCUST1")
        Call Update_Record_TDA("WBTCUST2")
        Call Update_Record_TDA("WBTCUST3")
        Call Update_Record_TDA("WBTCUST9")

        Call CommitTrans(Msg)

        Dim SQLS As New StringBuilder() With {.Length = 0}
        SQLS.AppendLine("UPDATE WBTCUST1 SET EMAIL = UPPER(EMAIL)")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub UpdateAndRefreshData(Optional ByVal Silent As Boolean = False)
        Me.Cursor = Cursors.WaitCursor

        Update_Record(Silent)

        Dim SQLS As New StringBuilder() With {.Length = 0}
        SQLS.AppendLine("DELETE FROM WBTCUST9 WHERE EMAIL IN")
        SQLS.AppendLine("(")
        SQLS.AppendLine("  SELECT EMAIL FROM WBTCUST1")
        SQLS.AppendLine("  WHERE NVL(PASSWORD,'NULL') <> 'NULL'")
        SQLS.AppendLine("  AND STATUS NOT IN ('N','M','C')")
        SQLS.AppendLine(")")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        EnforceConstraints(False)

        Call Fill_Records("WBTCUST1")
        Call Fill_Records("WBTCUST2")
        Call Fill_Records("WBTCUST3")
        Call Fill_Records("WBTCUST9")
        Call Fill_Records("ARTCONTX")
        Call Fill_Records("WBTCUSTP")
        Call Fill_Records("WBTCUSTR")

        EnforceConstraints(True)

        FilterWBTCUST1()
        CheckForCustomers()
        Me.Cursor = Cursors.Default
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(Optional WEB_CUST_BATCH As String = "", Optional EMAIL As String = "")
        Dim SQLS As New StringBuilder With {.Length = 0}
        If EMAIL.Length > 0 Then
            SQLS.Length = 0
            SQLS.AppendLine("Select *")
            SQLS.AppendLine("FROM WBTCUST2")
            SQLS.AppendLine(String.Format("WHERE EMAIL = '{0}'", EMAIL))
            SQLS.AppendLine(String.Format("AND WEB_CUST_BATCH = (SELECT MAX(WEB_CUST_BATCH) FROM WBTCUST2 WHERE EMAIL = '{0}')", EMAIL))
            'SELECT MAX(WEB_CUST_BATCH) FROM WBTCUST2 WHERE EMAIL = 'BRANDYSBOUTIQUEANDGIFTS@GMAIL.COM'
            Fill_Records("WBTCUSTC", "", True, SQLS.ToString)
        Else
            If WEB_CUST_BATCH.Length = 0 Then
                SQLS.Length = 0
                SQLS.AppendLine("Select MAX(WEB_CUST_BATCH) from WBTCUST2")
                ASCMAIN1.sql = SQLS.ToString()
                WEB_CUST_BATCH = ASCDATA1.GetDataValue
            End If
            Fill_Records("WBTCUSTC", WEB_CUST_BATCH)
        End If
        Print_Report_Begin()
        Generate_Report("WBRCUSTC")
        Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdWBTCUST1, "SSB", "Show Filter", "Show GroupBox", "Match All E-Mails", "Add Contact To Customer", "Match To Selected Contact", "Send Credit E-Mail", "Accept Customer", "Disable User Access", "Reject User", "Copy E-Mail", "Mass Update Sales Rep")
        Load_Popup_Menu(grdWBTCUST1, "SSBBBBBBBBBBBB", "Show Filter", "Show GroupBox", "Match To Selected Contact", "Send Credit E-Mail", "Accept Customer", "Disable User Access", "Reject User", "Move To New", "Re-Upload Contact", "Copy E-Mail", "Print Web Info", "Add Contact To Customer", "Claim Contact", "Release Claim", "Kill Claim", "Move To Another Customer")
        Load_Popup_Menu(grdWBTCUSTP, "SSB", "Show Filter", "Show GroupBox", "Re-Upload Contact")
        Load_Popup_Menu(grdWBTCUSTR, "SS", "Show Filter", "Show GroupBox")
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
            Case "grdWBTCUST1"
                Dim MY_CLAIM As Boolean = False
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    Dim CLAIM_BY_OPER As String = grdWBTCUST1.Selected.Rows(0).Cells("CLAIM_BY_OPER").Text & String.Empty
                    If CLAIM_BY_OPER = ASCMAIN1.USER_ID Then
                        MY_CLAIM = True
                    End If
                End If

                tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

                tlb_btn = DirectCast(tlb_pop.Tools("Add Contact To Customer"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = (rdoShowNew.Checked Or rdoShowCredit.Checked) And MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Match To Selected Contact"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = (rdoShowNew.Checked Or rdoShowCredit.Checked) And MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Send Credit E-Mail"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = (rdoShowNew.Checked) And MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Accept Customer"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = rdoShowMatched.Checked And MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Disable User Access"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = False 'Go Live Day.  We discussed this With Kyle that we need this feature in the future.
                    'tlb_btn.SharedProps.Visible = rdoShowAccepted.Checked
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Reject User"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = (rdoShowCredit.Checked Or rdoShowMatched.Checked Or rdoShowNew.Checked) And MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Move To New"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = (rdoShowDisabled.Checked Or rdoShowRejected.Checked Or rdoShowAccepted.Checked Or rdoShowMatched.Checked) And MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Move To Another Customer"), UltraWinToolbars.ButtonTool)
                If (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "drouse" Or ASCMAIN1.USER_ID = "mariog" Or ASCMAIN1.USER_ID = "site.admin") Then
                    If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                        tlb_btn.SharedProps.Visible = (rdoShowAccepted.Checked Or rdoShowMatched.Checked) And MY_CLAIM
                    Else
                        tlb_btn.SharedProps.Visible = False
                    End If
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Re-Upload Contact"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grd.ActiveRow IsNot Nothing And grd.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = rdoShowAccepted.Checked And MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If
                tlb_btn = DirectCast(tlb_pop.Tools("Copy E-Mail"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = rdoShowNew.Checked Or rdoShowCredit.Checked

                tlb_btn = DirectCast(tlb_pop.Tools("Claim Contact"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = Not MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Release Claim"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = MY_CLAIM
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Kill Claim"), UltraWinToolbars.ButtonTool)
                If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "mariog" Or ASCMAIN1.USER_ID = "site.admin")
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

                'tlb_btn = DirectCast(tlb_pop.Tools("Mass Update Sales Rep"), UltraWinToolbars.ButtonTool)
                'If Not ScreenMode And (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count > 1) Then
                '    tlb_btn.SharedProps.Visible = True
                'Else
                '    tlb_btn.SharedProps.Visible = False
                'End If

                'tlb_btn = DirectCast(tlb_pop.Tools("Match All E-Mails"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (rdoShowNew.Checked Or rdoShowCredit.Checked)

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTORDR1"
            '    If Not InquiryOnly Then
            '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
            '    End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            'Case "Match All E-Mails"
            '    If Not InquiryOnly Then
            '        MatchAllEMails()
            '    End If
            Case "Match To Selected Contact"
                If (grdARTCUSTX.ActiveRow IsNot Nothing And grdARTCUSTX.Selected.Rows.Count = 1) Then
                    If Not InquiryOnly Then
                        MatchContacts(grdWBTCUST1.Selected.Rows(0).ListObject.row, grdARTCUSTX.Selected.Rows(0).ListObject.row)
                        UpdateAndRefreshData(True)
                    End If
                Else
                    MsgBox("You Must Select A Contact To Match To", MsgBoxStyle.Exclamation, "Contact Selection")
                End If
            Case "Send Credit E-Mail"
                If Not InquiryOnly Then
                    Dim SREP_CODE As String = grdWBTCUST1.Selected.Rows(0).Cells.Item("SREP_CODE").Text & ""
                    Dim SendEmail As Boolean = True
                    If SREP_CODE.Length = 0 Then
                        Dim iResult As MsgBoxResult
                        Dim iTitle As String = "Sales Rep"
                        Dim iMSG As New StringBuilder With {.Length = 0}
                        iMSG.AppendLine("There Is No Sales Rep Assigned.")
                        iMSG.AppendLine("Is That OK With You?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult <> MsgBoxResult.Yes Then
                            SendEmail = False
                        End If
                    End If
                    If SendEmail Then
                        CreditEMail(grdWBTCUST1.Selected.Rows(0).ListObject.row)
                        UpdateAndRefreshData(True)
                    End If
                End If
            'Case "Mass Update Sales Rep"
            '    Dim SREP_CODE As String = InputBox("Please Enter A Valid Sales Rep Code", "Sales Rep")
            '    SREP_CODE = SREP_CODE.ToUpper
            '    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            '    SQLS.AppendLine(String.Format("SELECT COUNT(*) AS RECCNT FROM SOTSREP1 WHERE SREP_CODE = '{0}'", SREP_CODE))
            '    ASCMAIN1.sql = SQLS.ToString()
            '    Dim RECCNT As Int16 = Val(ASCDATA1.GetDataValue)
            '    If RECCNT > 0 Then
            '        For Each rowWBTCUST1 As UltraGridRow In grdWBTCUST1.Selected.Rows
            '            rowWBTCUST1.Cells.Item("SREP_CODE").Value = SREP_CODE
            '        Next
            '    Else
            '        MsgBox("Invalid Sales Rep Code Entered", vbExclamation, "Problem")
            '    End If

            Case "Add Contact To Customer"
                If Not InquiryOnly Then
                    If grdWBTCUST1.Selected.Rows.Count = 1 And grdARTCUSTX.Selected.Rows.Count = 1 Then
                        If grdARTCUSTX.Selected.Rows(0).Cells.Item("REC_TYPE").Text = "Customer" Then
                            AddContactsToCustomer(grdWBTCUST1.Selected.Rows(0).ListObject.row, grdARTCUSTX.Selected.Rows(0).ListObject.row)
                            UpdateAndRefreshData(True)
                        Else
                            MsgBox("You Must Select A Record Of Type Customer To Add The Contact To", MsgBoxStyle.Critical, "Selection Error")
                        End If
                    Else
                        MsgBox("You Must Select One And Only One Customer From The Grid Below To Add Contact To", MsgBoxStyle.Critical, "Selection Error")
                    End If
                End If
            Case Is = "Accept Customer"
                If Not InquiryOnly Then
                    AcceptCustomer(grdWBTCUST1.Selected.Rows(0).ListObject.row)
                    UpdateAndRefreshData(True)
                End If
            'Case Is = "Create Accepted Upload File"
            '    If Not InquiryOnly Then
            '        CreateCustomerUpload()
            '    End If
            Case Is = "Reject User"
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    If Not InquiryOnly Then
                        RejectContacts(grdWBTCUST1.Selected.Rows)
                        UpdateAndRefreshData(True)
                    End If
                Else
                    MsgBox("You Must Select A Contact To Reject", MsgBoxStyle.Exclamation, "Contact Selection")
                End If
            Case Is = "Re-Upload Contact"
                If (grd.ActiveRow IsNot Nothing And grd.Selected.Rows.Count = 1) Then
                    If Not InquiryOnly Then
                        Dim iResult As MsgBoxResult
                        Dim iTitle As String = "Re-Upload Contact"
                        Dim iMSG As New StringBuilder With {.Length = 0}
                        iMSG.AppendLine("This Will Move The Selected Contact")
                        iMSG.AppendLine("To The Awaiting Upload Section.")
                        iMSG.AppendLine("")
                        iMSG.AppendLine("Is That What You Want.")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult = MsgBoxResult.Yes Then
                            UploadContacts(grd.Selected.Rows)
                            UpdateAndRefreshData(True)
                        End If
                    End If
                Else
                    MsgBox("You Must Select A Contact To Re-Upload", MsgBoxStyle.Exclamation, "Contact Selection")
                End If
            Case Is = "Move To New"
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    If Not InquiryOnly Then
                        Dim iResult As MsgBoxResult
                        Dim iTitle As String = "Move To New"
                        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                        iMSG.AppendLine("This Will Move The Selected Contact Back")
                        iMSG.AppendLine("To New And Remove The Association To The")
                        iMSG.AppendLine("Customer.")
                        iMSG.AppendLine("")
                        iMSG.AppendLine("Are You Ready?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult = MsgBoxResult.Yes Then
                            MoveToNew(grdWBTCUST1.Selected.Rows)
                            UpdateAndRefreshData(True)
                        End If
                    End If
                Else
                    MsgBox("You Must Select A Contact To Move", MsgBoxStyle.Exclamation, "Contact Selection")
                End If
            Case Is = "Move To Another Customer"
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    If Not InquiryOnly Then
                        Dim iResult As MsgBoxResult
                        Dim iTitle As String = "Move To Another Customer"
                        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                        iMSG.AppendLine("This Will Move The Selected Contact To")
                        iMSG.AppendLine("The Customer Code You Will Be Asked For.")
                        iMSG.AppendLine("I Will Confirm That The Code You Provide")
                        iMSG.AppendLine("Is Valid.")
                        iMSG.AppendLine("")
                        iMSG.AppendLine("ALL FUTURE SALES FROM THIS CONTACT WILL")
                        iMSG.AppendLine("BE ASSIGNED TO THE NEW CUSTOMER!")
                        iMSG.AppendLine("")
                        iMSG.AppendLine("Are You Ready?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult = MsgBoxResult.Yes Then
                            MoveToNewCustomer(grdWBTCUST1.Selected.Rows.GetItem(0))
                            UpdateAndRefreshData(True)
                        End If
                    End If
                Else
                    MsgBox("You Must Select A Contact To Move", MsgBoxStyle.Exclamation, "Contact Selection")
                End If
            Case Is = "Copy E-Mail"
                If Not InquiryOnly Then
                    If grdWBTCUST1.Selected.Rows.Count = 1 Then
                        If Not InquiryOnly Then
                            Dim EMAIL As String = grdWBTCUST1.Selected.Rows(0).Cells.Item("EMAIL").Text
                            Clipboard.SetText(EMAIL)
                            MsgBox(EMAIL, MsgBoxStyle.Information, "Copied To Clipboard")
                            Context_Launch("", "", "", "ARTCUST1")
                        End If
                    Else
                        MsgBox("You Must Select A Contact To Copy", MsgBoxStyle.Exclamation, "Contact Selection")
                    End If
                End If
            Case "Print Web Info"
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    If Not InquiryOnly Then
                        Dim EMAIL As String = grdWBTCUST1.Selected.Rows(0).Cells.Item("EMAIL").Text
                        Print_Record("", EMAIL)
                    End If
                Else
                    MsgBox("You Must Select A Contact To Print", MsgBoxStyle.Exclamation, "Contact Selection")
                End If
            Case "Claim Contact"
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    Dim EMAIL As String = grdWBTCUST1.Selected.Rows(0).Cells.Item("EMAIL").Text
                    SQLS.AppendLine("SELECT NVL(CLAIM_BY_OPER,'') AS CLAIM_BY_OPER")
                    SQLS.AppendLine("FROM WBTCUST1")
                    SQLS.AppendLine(String.Format("WHERE EMAIL = '{0}'", EMAIL))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim CLAIM_BY_OPER As String = ASCDATA1.GetDataValue
                    If CLAIM_BY_OPER.Length > 0 Then
                        MsgBox(String.Format("This Contact Is Already Claimed By {0}", CLAIM_BY_OPER), vbOKOnly, "Conflict")
                    Else
                        grdWBTCUST1.Selected.Rows(0).Cells.Item("CLAIM_BY_OPER").Value = ASCMAIN1.USER_ID
                        SQLS.Length = 0
                        SQLS.AppendLine("UPDATE WBTCUST1")
                        SQLS.AppendLine(String.Format("SET CLAIM_BY_OPER = '{0}'", ASCMAIN1.USER_ID))
                        SQLS.AppendLine(String.Format("WHERE EMAIL = '{0}'", EMAIL))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                    End If
                Else
                    MsgBox("You Must Select A Contact To Claim", MsgBoxStyle.Exclamation, "Contact Selection")
                End If
            Case "Release Claim"
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    Dim EMAIL As String = grdWBTCUST1.Selected.Rows(0).Cells.Item("EMAIL").Text
                    Dim CLAIM_BY_OPER As String = grdWBTCUST1.Selected.Rows(0).Cells.Item("CLAIM_BY_OPER").Text
                    If CLAIM_BY_OPER <> ASCMAIN1.USER_ID Then
                        MsgBox("This Contact Is Not Claimed By You", vbOKOnly, "Conflict")
                    Else
                        grdWBTCUST1.Selected.Rows(0).Cells.Item("CLAIM_BY_OPER").Value = String.Empty
                        SQLS.Length = 0
                        SQLS.AppendLine("UPDATE WBTCUST1")
                        SQLS.AppendLine("SET CLAIM_BY_OPER = ''")
                        SQLS.AppendLine(String.Format("WHERE EMAIL = '{0}'", EMAIL))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                    End If
                End If
            Case "Kill Claim"
                If (grdWBTCUST1.ActiveRow IsNot Nothing And grdWBTCUST1.Selected.Rows.Count = 1) Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Kill Claim"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("This Will Kill The Claim That Exists")
                    iMSG.AppendLine("For This User.  They Will NOT Be Notified!")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Are You Sure You Know What You Are Doing?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.Yes Then
                        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                        Dim EMAIL As String = grdWBTCUST1.Selected.Rows(0).Cells.Item("EMAIL").Text
                        grdWBTCUST1.Selected.Rows(0).Cells.Item("CLAIM_BY_OPER").Value = String.Empty
                        SQLS.AppendLine("UPDATE WBTCUST1")
                        SQLS.AppendLine("SET CLAIM_BY_OPER = ''")
                        SQLS.AppendLine(String.Format("WHERE EMAIL = '{0}'", EMAIL))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                    Else
                        MsgBox("Chicken!", vbExclamation, "Kill Claim")
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'Select Case
        '        'Case "Edit Ship To"
        '    '    If Not InquiryOnly Then
        '    '        MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
        '    '    End If
        'End Select
    End Sub

    Private Sub MoveToNewCustomer(ByRef rowWBTCUST1 As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Dim OLD_CUST_CODE As String = rowWBTCUST1.Cells.Item("CUST_CODE_ACTUAL").Text.ToString & String.Empty
        If OLD_CUST_CODE.Length > 0 Then
            Dim NEW_CUST_CODE As String = InputBox("Please Provide A New Customer Code", "New Code", OLD_CUST_CODE)
            If NEW_CUST_CODE.Length > 0 Then
                If OLD_CUST_CODE <> NEW_CUST_CODE Then
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine(String.Format("Select Count(*) from ARTCUST1 where CUST_CODE = '{0}'", NEW_CUST_CODE))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                    If REC_CNT = 1 Then
                        rowWBTCUST1.Cells.Item("CUST_CODE_ACTUAL").Value = NEW_CUST_CODE
                        rowWBTCUST1.Cells.Item("CONTACT_NO").Value = 0
                        rowWBTCUST1.Cells.Item("CONTACT_TYPE").Value = 1
                    Else
                        MsgBox(String.Format("No Match Found For Customer Code {0}", NEW_CUST_CODE), vbExclamation, "Try Again")
                    End If
                End If
            Else
                MsgBox("New Customer Code Not Provided", vbExclamation, "Try Again")
            End If

        Else
            MsgBox("Origianl Customer Code Is Not Set", vbExclamation, "Hmmm")
        End If

    End Sub


#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "XXXXXXXX"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "YYYYYYYY"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "XXXXXXXX"
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

#Region "Form Controls"

#Region "Buttons"

    Private Sub btnInclude_Click(sender As System.Object, e As System.EventArgs) Handles btnInclude.Click
        If btnInclude.Text = "Include" Then
            btnInclude.Text = "Exclude"
            IncludeFilter = "OR"
        Else
            btnInclude.Text = "Include"
            IncludeFilter = "AND"
        End If
        ShowSelectedMatches()
    End Sub

#End Region

#Region "Check Boxes"
    Private Sub chkMCity_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMCity.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMCompany_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMCompany.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMCustCode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMCustCode.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMEMail_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMEMail.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMFName_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMFName.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMLName_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMLName.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMPhone_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMPhone.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMState_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMState.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMStreet_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMStreet.CheckedChanged
        ShowSelectedMatches()
    End Sub

    Private Sub chkMZip_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMZip.CheckedChanged
        ShowSelectedMatches()
    End Sub
#End Region

#Region "Grids"
    Private Sub grdWBTCUST1_AfterSelectChange(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdWBTCUST1.AfterSelectChange
        ShowSelectedMatches()
    End Sub

    Private Sub grdWBTCUST1_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdWBTCUST1.ClickCellButton
        If grdWBTCUST1.ActiveRow Is Nothing Then Exit Sub
        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "SREP_CODE"
                Call grdClickCellButton(grdWBTCUST1, sql_where)
        End Select

    End Sub
#End Region

#Region "Radio Buttons"

    Private Sub rdoShowAccepted_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoShowAccepted.CheckedChanged
        FilterWBTCUST1()
    End Sub

    Private Sub rdoShowCredit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoShowCredit.CheckedChanged
        FilterWBTCUST1()
    End Sub

    Private Sub rdoShowDisabled_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoShowDisabled.CheckedChanged
        FilterWBTCUST1()
    End Sub

    Private Sub rdoShowMatched_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoShowMatched.CheckedChanged
        FilterWBTCUST1()
    End Sub

    Private Sub rdoShowNew_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoShowNew.CheckedChanged
        FilterWBTCUST1()
    End Sub
#End Region

#End Region

#Region "Custom Methods"

    Private Sub AcceptCustomer(rowWBTCUST1 As DataRow)
        If IsNothing(rowWBTCUST1) Then
            MsgBox("Error Selecting Row", MsgBoxStyle.Critical, "Selection Error")
        Else
            'SendEMail(rowWBTCUST1, "ACCEPT")
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Accept Customer"
            Dim iMSG As New StringBuilder
            iMSG.AppendLine("Would You Like To Move This User To Accepted?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                rowWBTCUST1.Item("STATUS") = "U"
                rowWBTCUST1.Item("CLAIM_BY_OPER") = Null
                rowWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowWBTCUST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            End If
        End If
    End Sub

    Private Shared Sub AddBlankVAR(ByVal VAR As List(Of REPORT_VARIABLES))
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_CODE]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_NAME]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_ADDR1]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_ADDR2]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_CITY]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_STATE]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_ZIP_CODE]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_PHONE]", .VAR_VALUE = ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_EMAIL]", .VAR_VALUE = ""})
    End Sub

    Private Sub AddContactsToCustomer(rowWBTCUST1 As DataRow, rowARTCONTX As DataRow)
        Dim iResult As MsgBoxResult
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        Dim iTitle As String = "Problem Adding Contact"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}

        Dim CUST_CODE As String = rowARTCONTX.Item("CUST_CODE") & ""
        If CUST_CODE.Length <> 6 Then
            iMSG.Length = 0
            iMSG.AppendLine("Invalid Customer Code!")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            Exit Sub
        End If

        SQLS.Length = 0
        SQLS.AppendLine(String.Format("Select Count(*) from ARTCUST1 where CUST_CODE = '{0}'", CUST_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim CUST_CNT As Int16 = Val(ASCDATA1.GetDataValue)
        If CUST_CNT <> 1 Then
            iMSG.Length = 0
            iMSG.AppendLine("Invalid Customer Code!")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            Exit Sub
        End If

        iMSG.Length = 0
        iMSG.AppendLine("This Will Add The Selected Contact To")
        iMSG.AppendLine(String.Format("Customer {0}", CUST_CODE))
        iMSG.AppendLine("")
        iMSG.AppendLine("Is That What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, "You Sure?")
        If iResult <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        dst.Tables("ARTCUSTD").Clear()
        SQLS.Length = 0
        SQLS.AppendLine("Select MAX(CONTACT_NO) As MAX_NO")
        SQLS.AppendLine("FROM ARTCUSTD")
        SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", rowARTCONTX.Item("CUST_CODE") & ""))
        ASCMAIN1.sql = SQLS.ToString()
        Dim CONTACT_NO As Int16 = Val(ASCDATA1.GetDataValue) + 1

        rowWBTCUST1.Item("STATUS") = "M"
        rowWBTCUST1.Item("CLAIM_BY_OPER") = Null
        rowWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowWBTCUST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        rowWBTCUST1.Item("CUST_CODE_ACTUAL") = CUST_CODE
        SQLS.Length = 0
        SQLS.AppendLine(String.Format("SELECT SREP_CODE FROM ARTCUST1 WHERE CUST_CODE = '{0}'", CUST_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim SREP_CODE As String = ASCDATA1.GetDataValue
        If SREP_CODE.Length > 0 Then
            rowWBTCUST1.Item("SREP_CODE") = SREP_CODE
        End If
        rowWBTCUST1.Item("CONTACT_TYPE") = "D"
        rowWBTCUST1.Item("CONTACT_NO") = CONTACT_NO

        Dim newARTCUSTD As DataRow = dst.Tables("ARTCUSTD").NewRow
        newARTCUSTD.Item("CUST_CODE") = CUST_CODE
        newARTCUSTD.Item("CONTACT_NO") = CONTACT_NO
        newARTCUSTD.Item("CONTACT_NAME") = String.Format("{0} {1}", rowWBTCUST1.Item("GIVENNAME").ToString.Trim & "", rowWBTCUST1.Item("FAMILYNAME").ToString.Trim & "")
        newARTCUSTD.Item("CONTACT_EMAIL") = rowWBTCUST1.Item("EMAIL").ToString & ""
        newARTCUSTD.Item("CONTACT_PHONE") = rowWBTCUST1.Item("TELEPHONE").ToString & ""
        newARTCUSTD.Item("CONTACT_TYPE") = "B"
        newARTCUSTD.Item("CONTACT_PRIMARY") = "0"
        newARTCUSTD.Item("INIT_OPER") = ASCMAIN1.USER_ID
        newARTCUSTD.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        newARTCUSTD.Item("LAST_OPER") = ASCMAIN1.USER_ID
        newARTCUSTD.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
        dst.Tables("ARTCUSTD").Rows.Add(newARTCUSTD)
        Call BeginTrans()
        Call Update_Record_TDA("ARTCUSTD")
        Call CommitTrans("")
        Dim Temp_Select As String = String.Format("{0} WHERE REC_TYPE = 'Contact' AND CUST_CODE = {1} AND CONTACT_NO = {2}", sqlARTCONTX, rowARTCONTX.Item("CUST_CODE") & "", CONTACT_NO)
        Fill_Records("ARTCONTX", , False, Temp_Select)
    End Sub

    Private Function AddCurrentRow(currentRow As String(), ByRef FilePositions As List(Of LexMap)) As Integer
        Dim CurrEmail As String = ""
        Dim FamilyNameIndex As Integer = -1
        Dim CurrEmailIndex As Integer = -1
        For Each L As LexMap In FilePositions
            If L.LEX_COL = "FamilyName" Then
                FamilyNameIndex = L.COL_INDEX
            End If
            If L.LEX_COL = "Email" Then
                CurrEmailIndex = L.COL_INDEX
            End If
        Next
        If FamilyNameIndex = -1 Or CurrEmailIndex = -1 Then
            Return 0
            Exit Function
        End If

        If currentRow(FamilyNameIndex).ToString = "FamilyName" Then
            Return 0
            Exit Function
        End If
        CurrEmail = currentRow(CurrEmailIndex).ToString.ToUpper()
        If CurrEmail.Length = 0 Then
            Return 0
            Exit Function
        End If
        dst.Tables("WBTCUST1").CaseSensitive = False
        If dst.Tables("WBTCUST1").Select(String.Format("EMAIL = '{0}'", CurrEmail)).Count > 0 Then
            Return 0
            Exit Function
        End If
        Dim newWBTCUST1 As DataRow = dst.Tables("WBTCUST1").NewRow
        For i As Integer = 0 To currentRow.Length - 1
            Dim FoundInFile As Boolean = False
            Dim FoundIndex As Integer = 0
            For Each L As LexMap In FilePositions
                If L.COL_INDEX = i Then
                    Select Case L.ABS_COL
                        Case "DATEREGISTERED"
                            If IsDate(currentRow(i).ToString) Then
                                newWBTCUST1.Item(L.ABS_COL) = Format(CDate(currentRow(i).ToString), "MM/dd/yyyy")
                            Else
                                newWBTCUST1.Item(L.ABS_COL) = Format(CDate(Now()), "MM/dd/yyyy")
                            End If
                        Case "TAX_ID"
                            If currentRow(i).ToString.Length < 30 Then
                                newWBTCUST1.Item(L.ABS_COL) = currentRow(i).ToString
                            End If
                        Case Else
                            newWBTCUST1.Item(L.ABS_COL) = currentRow(i).ToString
                    End Select
                End If
            Next
        Next
        newWBTCUST1.Item("STATUS") = "N"
        dst.Tables("WBTCUST1").Rows.Add(newWBTCUST1)
        Return 1
    End Function

    Private Function CalculatepriceGroup(ByRef rowARTCUST1 As DataRow) As String
        Dim RetVal As String = "0"
        If chkExportTesting.Checked = False Then
            Dim CUST_PRICE_TIER As String = rowARTCUST1.Item("CUST_PRICE_TIER").ToString & String.Empty
            Dim CUST_DISC_PCT_EXTRA As String = rowARTCUST1.Item("CUST_DISC_PCT_EXTRA").ToString & String.Empty
            If CUST_DISC_PCT_EXTRA = "" Then
                CUST_DISC_PCT_EXTRA = "0"
            End If
            Select Case CUST_PRICE_TIER
                Case "PC"
                    Select Case CUST_DISC_PCT_EXTRA
                        Case "1"
                            RetVal = "1"
                        Case "2"
                            RetVal = "2"
                    End Select
                Case "HC"
                    RetVal = "3"
                Case "FC"
                    RetVal = "4"
                Case "SP"
                    RetVal = "5"
            End Select
        Else
            Dim CUST_PRICE_TIER As String = rowARTCUST1.Item("CUST_PRICE_TIER").ToString & String.Empty
            Dim CUST_DISC_PCT_EXTRA As String = rowARTCUST1.Item("CUST_DISC_PCT_EXTRA").ToString & String.Empty
            Dim CUST_DISC_PCT As Int64 = Val(rowARTCUST1.Item("CUST_DISC_PCT").ToString & String.Empty)
            If CUST_DISC_PCT_EXTRA = "" Then
                CUST_DISC_PCT_EXTRA = "0"
            End If
            Select Case CUST_PRICE_TIER
                Case "PC"
                    Select Case CUST_DISC_PCT_EXTRA
                        Case "1"
                            RetVal = "1"
                        Case "2"
                            RetVal = "2"
                    End Select
                Case "HC"
                    RetVal = "3"
                Case "FC"
                    RetVal = "4"
                Case "SP"
                    Select Case CUST_DISC_PCT
                        Case 52
                            RetVal = "5"
                        Case 54
                            RetVal = "6"
                        Case 55
                            RetVal = "7"
                        Case 56
                            RetVal = "8"
                        Case 57
                            RetVal = "9"
                        Case 59
                            RetVal = "10"
                    End Select
            End Select
        End If
        Return RetVal
    End Function

    Private Function CalculateTerms(ByRef rowARTCUST1 As DataRow) As String
        Dim RetVal As String = "1" 'Credit Card Only.
        Dim TERM_CODE As String = rowARTCUST1.Item("TERM_CODE").ToString & String.Empty
        If TERM_CODE = "N30" Then
            RetVal = "2" 'Net 30 or Credit Card.
        End If
        Return RetVal
    End Function

    Private Sub CheckForCustomers()
        lblCustomerWaiting.Text = "Customers Waiting"
        lblCustomerWaiting.Visible = False
        'Try
        '    FtpShopSite.User = UserName
        '    FtpShopSite.Password = Password
        '    FtpShopSite.RemoteHost = RemoteHost
        '    FtpShopSite.RemotePath = RemotePath
        '    FtpShopSite.Logon()
        '    FtpShopSite.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
        '    FtpShopSite.LocalFile = ""
        '    FtpShopSite.RemoteFile = inBoundFile
        '    FtpShopSite.Overwrite = False
        '    If FtpShopSite.FileExists() Then
        '        lblCustomerWaiting.Visible = True
        '    End If
        '    FtpShopSite.Logoff()
        'Catch ex As Exception
        '    lblCustomerWaiting.Text = "Not Logged In"
        '    lblCustomerWaiting.Visible = True
        'End Try
    End Sub

    Private Sub CreditEMail(rowWBTCUST1 As DataRow)
        If IsNothing(rowWBTCUST1) Then
            MsgBox("Error Selecting Row", MsgBoxStyle.Critical, "Selection Error")
        Else
            SendEMail(rowWBTCUST1, "CREDIT")
            rowWBTCUST1.Item("STATUS") = "C"
            rowWBTCUST1.Item("CLAIM_BY_OPER") = Null
            rowWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowWBTCUST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        End If
    End Sub

    Private Sub FilterWBTCUST1()
        If rdoShowAccepted.Checked = False _
            And rdoShowCredit.Checked = False _
            And rdoShowDisabled.Checked = False _
            And rdoShowMatched.Checked = False _
            And rdoShowNew.Checked = False _
            And rdoShowRejected.Checked = False _
            And rdoShowUpload.Checked = False Then
            rdoShowAccepted.Checked = True
        End If

        If Not IsNothing(grdWBTCUST1.DataSource) Then
            Dim GrdFilter As String = ""
            If rdoShowNew.Checked Then
                GrdFilter = "STATUS = 'N'"
            End If
            If rdoShowMatched.Checked Then
                GrdFilter = "STATUS = 'M'"
                'If GrdFilter.Length > 0 Then
                '    GrdFilter += " OR STATUS = 'M'"
                'Else
                '    GrdFilter = "STATUS = 'M'"
                'End If
            End If
            If rdoShowCredit.Checked Then
                GrdFilter = "STATUS = 'C'"
                'If GrdFilter.Length > 0 Then
                '    GrdFilter += " OR STATUS = 'C'"
                'Else
                '    GrdFilter = "STATUS = 'C'"
                'End If
            End If
            If rdoShowAccepted.Checked Then
                GrdFilter = "STATUS = 'A'"
                'If GrdFilter.Length > 0 Then
                '    GrdFilter += " OR (STATUS = 'A' OR STATUS = 'U' OR STATUS = 'T')"
                'Else
                '    GrdFilter = "(STATUS = 'A' OR STATUS = 'U' OR STATUS = 'T')"
                'End If
            End If
            If rdoShowUpload.Checked Then
                GrdFilter = "STATUS = 'U'"
                'If GrdFilter.Length > 0 Then
                '    GrdFilter += " OR (STATUS = 'A' OR STATUS = 'U' OR STATUS = 'T')"
                'Else
                '    GrdFilter = "(STATUS = 'A' OR STATUS = 'U' OR STATUS = 'T')"
                'End If
            End If
            If rdoShowDisabled.Checked Then
                GrdFilter = "STATUS = 'D'"
                'If GrdFilter.Length > 0 Then
                '    GrdFilter += " OR STATUS = 'D'"
                'Else
                '    GrdFilter = "STATUS = 'D'"
                'End If
            End If
            If rdoShowRejected.Checked Then
                GrdFilter = "STATUS = 'R'"
                'If GrdFilter.Length > 0 Then
                '    GrdFilter += " OR STATUS = 'R'"
                'Else
                '    GrdFilter = "STATUS = 'R'"
                'End If
            End If
            Dim dvw As DataView = DirectCast(grdWBTCUST1.DataSource, DataTable).DefaultView
            dvw.RowFilter = GrdFilter
        End If
        ShowSelectedMatches()
    End Sub

    Private Sub FindAndReplace(ByVal doc As Microsoft.Office.Interop.Word.Document,
                               ByVal FindText As String, ByVal ReplaceText As String)
        Dim WordRange As Microsoft.Office.Interop.Word.Range

        Try
            For Each WordRange In doc.StoryRanges

                With WordRange.Find
                    .Text = FindText ' "<<FULLNAME>>"
                    .Replacement.Text = ReplaceText
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                End With


            Next WordRange
        Catch ex As Exception
            ' Do Nothing
        Finally
            Beep()
        End Try

    End Sub

    Private Function FormatBadChars(ByVal InString As String, Optional UpperCase As Boolean = False) As String
        Dim RetVal As String = InString
        Dim BadChars As String() = {"[", "]", "'"}
        For Each BadChar As String In BadChars
            RetVal = RetVal.Replace(BadChar, "")
        Next
        If UpperCase Then
            RetVal = RetVal.ToUpper
        End If
        Return RetVal
    End Function

    Private Function GetCustPassword(ByVal EMAIL As String) As String
        Dim RetVal As String = ""
        Dim filter As String = String.Format("EMAIL = '{0}'", EMAIL.ToUpper)
        Dim rowWBTCUST9 As DataRow = dst.Tables.Item("WBTCUST9").Select(filter).FirstOrDefault
        If Not IsNothing(rowWBTCUST9) Then
            Dim PSWDE As String = rowWBTCUST9.Item("PSWDE").ToString & String.Empty
            Dim PSWDD As String = psEncrypt(PSWDE, EncryptType.Decrypt)
            RetVal = PSWDD
        End If
        Return RetVal
    End Function

    Private Function GetCustGroupCode(ByVal CUST_CODE_ACTUAL As String) As String
        Dim RetVal As String = ""
        Dim SQLS As New StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("SELECT TERM_CODE")
        SQLS.AppendLine("FROM ARTCUST1")
        SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE_ACTUAL))
        ASCMAIN1.sql = SQLS.ToString()
        Dim TERM_CODE As String = ASCDATA1.GetDataValue
        If TERM_CODE.Length > 0 Then
            If TERM_CODE = "N30" Then
                RetVal = "wholesale1"
            Else
                RetVal = "wholesale"
            End If
        End If
        Return RetVal
    End Function

    Private Function GetFilePositions(currentRow As String()) As List(Of LexMap)
        Dim RetVal As New List(Of LexMap)
        RetVal.Add(New LexMap() With {.ABS_COL = "FAMILYNAME", .LEX_COL = "FamilyName", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "GIVENNAME", .LEX_COL = "GivenName", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "EMAIL", .LEX_COL = "Email", .COL_INDEX = vbNull})
        RetVal.Add(New LexMap() With {.ABS_COL = "DATEREGISTERED", .LEX_COL = "DateRegistered", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "GROUPNAME", .LEX_COL = "GroupName", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "CITY", .LEX_COL = "company_city", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "CUST_CODE_PROVIDED", .LEX_COL = "company_cust_code", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "COMPANY", .LEX_COL = "company_name", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "TELEPHONE", .LEX_COL = "company_phone", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "STATE", .LEX_COL = "company_state", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "STREET", .LEX_COL = "company_street", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "TAX_ID", .LEX_COL = "company_taxid", .COL_INDEX = -1})
        RetVal.Add(New LexMap() With {.ABS_COL = "ZIP_CODE", .LEX_COL = "company_zip", .COL_INDEX = -1})
        For i As Integer = 0 To currentRow.Length - 1
            For Each L As LexMap In RetVal
                If L.LEX_COL = currentRow(i).ToString Then
                    L.COL_INDEX = i
                End If
            Next L
        Next
        For Each L As LexMap In RetVal
            If L.COL_INDEX = -1 Then
                MsgBox("Un-Mapped Columns Found In Lexicon File.", MsgBoxStyle.Critical, "Can Not Complete")
                Stop
            End If
        Next
        Return RetVal
    End Function

    Private Function GetReportVariables(ByRef rowWBTCUST1 As DataRow, ByVal REPORT_NAME As String) As List(Of REPORT_VARIABLES)
        Dim VAR As New List(Of REPORT_VARIABLES)
        Dim Error_found As Boolean = False
        Dim rowARTCONTX As DataRow = dst.Tables.Item("ARTCONTX").NewRow
        Dim HasMatch As Boolean = False
        If grdWBTCUST1.Selected.Rows(0).Cells.Item("STATUS").Text & "" = "Matched" Then
            Dim CONTACT_TYPE As String = ""
            Select Case rowWBTCUST1.Item("CONTACT_TYPE").ToString & ""
                Case Is = "1"
                    CONTACT_TYPE = "Customer"
                Case Is = "2"
                    CONTACT_TYPE = "Ship-To"
                Case Is = "D"
                    CONTACT_TYPE = "Contact"
            End Select
            Dim Filter = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND REC_TYPE = '{2}'",
                                       grdWBTCUST1.Selected.Rows(0).Cells.Item("CUST_CODE_ACTUAL").Text,
                                       grdWBTCUST1.Selected.Rows(0).Cells.Item("CONTACT_NO").Text, CONTACT_TYPE)
            If dst.Tables.Item("ARTCONTX").Select(Filter).Count = 1 Then
                HasMatch = True
                rowARTCONTX = dst.Tables.Item("ARTCONTX").Select(Filter).FirstOrDefault
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_CODE]", .VAR_VALUE = rowARTCONTX.Item("CUST_CODE").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_NAME]", .VAR_VALUE = rowARTCONTX.Item("CUST_NAME").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_ADDR1]", .VAR_VALUE = rowARTCONTX.Item("CUST_ADDR1").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_ADDR2]", .VAR_VALUE = rowARTCONTX.Item("CUST_ADDR2").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_CITY]", .VAR_VALUE = rowARTCONTX.Item("CUST_CITY").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_STATE]", .VAR_VALUE = rowARTCONTX.Item("CUST_STATE").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_ZIP_CODE]", .VAR_VALUE = rowARTCONTX.Item("CUST_ZIP_CODE").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_PHONE]", .VAR_VALUE = rowARTCONTX.Item("CUST_PHONE").ToString & ""})
                VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[MATCHED_CUST_EMAIL]", .VAR_VALUE = rowARTCONTX.Item("CUST_EMAIL").ToString & ""})
            Else
                rowARTCONTX = dst.Tables.Item("ARTCONTX").NewRow
                AddBlankVAR(VAR)
            End If
        Else
            AddBlankVAR(VAR)
        End If

        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[REPORT_NAME]", .VAR_VALUE = REPORT_NAME})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_EMAIL]", .VAR_VALUE = rowWBTCUST1.Item("EMAIL").ToString & ""})
        EMAIL_ADDRESS = rowWBTCUST1.Item("EMAIL").ToString & ""
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_CUST_CODE_PROVIDED]", .VAR_VALUE = rowWBTCUST1.Item("CUST_CODE_PROVIDED").ToString & ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_CUST_CODE_ACTUAL]", .VAR_VALUE = rowWBTCUST1.Item("CUST_CODE_ACTUAL").ToString & ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_DATEREGISTERED]", .VAR_VALUE = rowWBTCUST1.Item("DATEREGISTERED").ToString & ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_DATEAPPROVED]", .VAR_VALUE = rowWBTCUST1.Item("DATEAPPROVED").ToString & ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_FAMILYNAME]", .VAR_VALUE = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("FAMILYNAME").ToString.ToLower & "")})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_GIVENNAME]", .VAR_VALUE = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("GIVENNAME").ToString.ToLower & "")})
        EMAIL_NAME = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("GIVENNAME").ToString.ToLower & "") & " " & CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("FAMILYNAME").ToString.ToLower & "")
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_COMPANY]", .VAR_VALUE = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("COMPANY").ToString & "")})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_TAX_ID]", .VAR_VALUE = rowWBTCUST1.Item("TAX_ID").ToString & ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_STATUS]", .VAR_VALUE = rowWBTCUST1.Item("STATUS").ToString & ""})
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_TELEPHONE]", .VAR_VALUE = rowWBTCUST1.Item("TELEPHONE").ToString & ""})

        Dim WEB_STREET As String = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("STREET").ToString.ToLower() & "")
        If WEB_STREET.Length = 0 And HasMatch Then
            WEB_STREET = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowARTCONTX.Item("CUST_ADDR1").ToString.ToLower() & "")
        End If
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_STREET]", .VAR_VALUE = WEB_STREET})

        Dim WEB_CITY As String = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("CITY").ToString.ToLower() & "")
        If WEB_CITY.Length = 0 And HasMatch Then
            WEB_CITY = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowARTCONTX.Item("CUST_CITY").ToString.ToLower() & "")
        End If
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_CITY]", .VAR_VALUE = WEB_CITY})

        Dim WEB_STATE As String = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("STATE").ToString.ToLower() & "")
        If WEB_STATE.Length = 0 And HasMatch Then
            WEB_STATE = rowARTCONTX.Item("CUST_STATE").ToString & ""
        End If
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_STATE]", .VAR_VALUE = WEB_STATE})

        Dim ZIP_CODE As String = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rowWBTCUST1.Item("ZIP_CODE").ToString.ToLower() & "")
        If ZIP_CODE.Length = 0 And HasMatch Then
            ZIP_CODE = rowARTCONTX.Item("CUST_ZIP_CODE").ToString & ""
        End If
        VAR.Add(New REPORT_VARIABLES() With {.VAR_NAME = "[WEB_ZIP_CODE]", .VAR_VALUE = ZIP_CODE})

        Return VAR
    End Function

    Private Function LOCK_IMPORT_EXPORT(ByVal lock_type As LockType, ByVal lock_direction As LockDirection) As Boolean
        Dim RetVal As Boolean = False
        Dim SQLS As New StringBuilder With {.Length = 0}

        Dim WB_XFR_OPER As String = ASCMAIN1.USER_ID
        Dim WB_XFR_DATE As String = Format(Now(), "dd-MMM-yyyy")

        Dim LOC_DIR As String = ""
        If lock_direction = LockDirection.Import Then
            LOC_DIR = "I"
        Else
            LOC_DIR = "E"
        End If

        Select Case lock_type
            Case LockType.Lock
                SQLS.Length = 0
                SQLS.AppendLine("SELECT NVL(WB_XFR_STATUS,'A') AS WB_XFR_STATUS")
                SQLS.AppendLine("FROM WBTCUSTP")
                SQLS.AppendLine("WHERE WB_PARM_KEY = 'Z'")
                ASCMAIN1.sql = SQLS.ToString()
                Dim WB_XFR_STATUS As String = ASCDATA1.GetDataValue
                If WB_XFR_STATUS = "A" Then
                    SQLS.Length = 0
                    SQLS.AppendLine("UPDATE WBTCUSTP")
                    SQLS.AppendLine(String.Format("SET WB_XFR_STATUS = '{0}',", LOC_DIR))
                    SQLS.AppendLine(String.Format("WB_XFR_OPER = '{0}',", WB_XFR_OPER))
                    SQLS.AppendLine(String.Format("WB_XFR_DATE = '{0}'", WB_XFR_DATE))
                    SQLS.AppendLine("WHERE WB_PARM_KEY = 'Z'")
                    ASCMAIN1.sql = SQLS.ToString
                    ASCDATA1.ExecuteSQL()
                    RetVal = True
                Else
                    SQLS.Length = 0
                    SQLS.AppendLine("SELECT NVL(WB_XFR_OPER,'') AS WB_XFR_OPER")
                    SQLS.AppendLine("FROM WBTCUSTP")
                    SQLS.AppendLine("WHERE WB_PARM_KEY = 'Z'")
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim WB_XFR_OPER_CUR As String = ASCDATA1.GetDataValue
                    MsgBox("Import/Export Locked By " & WB_XFR_OPER_CUR, vbOKOnly, "Lock In Place")
                End If
            Case LockType.Release
                SQLS.Length = 0
                SQLS.AppendLine("SELECT NVL(WB_XFR_OPER,'') AS WB_XFR_OPER")
                SQLS.AppendLine("FROM WBTCUSTP")
                SQLS.AppendLine("WHERE WB_PARM_KEY = 'Z'")
                ASCMAIN1.sql = SQLS.ToString()
                Dim WB_XFR_OPER_CUR As String = ASCDATA1.GetDataValue
                If WB_XFR_OPER = WB_XFR_OPER_CUR Then
                    SQLS.Length = 0
                    SQLS.AppendLine("UPDATE WBTCUSTP")
                    SQLS.AppendLine("SET WB_XFR_STATUS = NULL,")
                    SQLS.AppendLine("WB_XFR_OPER = NULL,")
                    SQLS.AppendLine("WB_XFR_DATE = NULL")
                    SQLS.AppendLine("WHERE WB_PARM_KEY = 'Z'")
                    ASCMAIN1.sql = SQLS.ToString
                    ASCDATA1.ExecuteSQL()
                    RetVal = True
                Else
                    MsgBox("Import/Export Locked By " & WB_XFR_OPER_CUR, vbOKOnly, "Lock In Place")
                End If
            Case Else
                Stop 'This Can Never Happen Unless we change the enum.
        End Select

        Return RetVal
    End Function

    Private Sub MatchContacts(rowWBTCUST1 As DataRow, rowARTCONTX As DataRow)
        rowWBTCUST1.Item("CUST_CODE_ACTUAL") = rowARTCONTX.Item("CUST_CODE").ToString
        Dim SQLS As New StringBuilder With {.Length = 0}
        SQLS.AppendLine(String.Format("SELECT SREP_CODE FROM ARTCUST1 WHERE CUST_CODE = '{0}'", rowARTCONTX.Item("CUST_CODE").ToString))
        ASCMAIN1.sql = SQLS.ToString()
        Dim SREP_CODE As String = ASCDATA1.GetDataValue
        If SREP_CODE.Length > 0 Then
            rowWBTCUST1.Item("SREP_CODE") = SREP_CODE
        End If
        rowWBTCUST1.Item("CONTACT_NO") = rowARTCONTX.Item("CONTACT_NO").ToString
        Select Case rowARTCONTX.Item("REC_TYPE").ToString
            Case Is = "Customer"
                rowWBTCUST1.Item("CONTACT_TYPE") = "1"
            Case Is = "Ship-To"
                rowWBTCUST1.Item("CONTACT_TYPE") = "2"
            Case Is = "Contact"
                rowWBTCUST1.Item("CONTACT_TYPE") = "D"
        End Select
        rowWBTCUST1.Item("STATUS") = "M"
        rowWBTCUST1.Item("CLAIM_BY_OPER") = Null
        rowWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowWBTCUST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        If rowWBTCUST1.Item("EMAIL").ToString.ToUpper <> rowARTCONTX.Item("CUST_EMAIL").ToString.ToUpper & "" Then
            rowARTCONTX.Item("CUST_EMAIL") = rowWBTCUST1.Item("EMAIL").ToString.ToUpper
            Select Case rowARTCONTX.Item("REC_TYPE").ToString
                Case Is = "Customer"
                    Fill_Records("ARTCUST1", rowARTCONTX.Item("CUST_CODE").ToString & "")
                    If dst.Tables("ARTCUST1").Rows.Count = 1 Then
                        dst.Tables("ARTCUST1").Rows(0).Item("CUST_EMAIL") = rowWBTCUST1.Item("EMAIL").ToString.ToUpper
                    End If
                    Call BeginTrans()
                    Call Update_Record_TDA("ARTCUST1")
                    Call CommitTrans("")
                Case Is = "Ship-To"
                    Fill_Records("ARTCUST2", New Object() {rowARTCONTX.Item("CUST_CODE").ToString & "", "MK", rowARTCONTX.Item("CUST_ADDR_CODE") & ""}, True)
                    If dst.Tables("ARTCUST2").Rows.Count = 1 Then
                        dst.Tables("ARTCUST2").Rows(0).Item("CUST_EMAIL") = rowWBTCUST1.Item("EMAIL").ToString.ToUpper
                    End If
                    Call BeginTrans()
                    Call Update_Record_TDA("ARTCUST2")
                    Call CommitTrans("")
                Case Is = "Contact"
                    Fill_Records("ARTCUSTD", New Object() {rowARTCONTX.Item("CUST_CODE").ToString & "", CInt(rowARTCONTX.Item("CONTACT_NO"))}, True)
                    If dst.Tables("ARTCUSTD").Rows.Count = 1 Then
                        dst.Tables("ARTCUSTD").Rows(0).Item("CONTACT_EMAIL") = rowWBTCUST1.Item("EMAIL").ToString.ToUpper
                    End If
                    Call BeginTrans()
                    Call Update_Record_TDA("ARTCUSTD")
                    Call CommitTrans("")
            End Select
        End If
    End Sub

    Private Sub MoveToNew(ByRef rowsCUST As Infragistics.Win.UltraWinGrid.SelectedRowsCollection)
        For Each rowCUST As Infragistics.Win.UltraWinGrid.UltraGridRow In rowsCUST
            rowCUST.Cells.Item("STATUS").Value = "N"
            rowCUST.Cells.Item("CUST_CODE_ACTUAL").Value = Null
            rowCUST.Cells.Item("CLAIM_BY_OPER").Value = Null
            rowCUST.Cells.Item("LAST_OPER").Value = ASCMAIN1.USER_ID
            rowCUST.Cells.Item("LAST_DATE").Value = Now + ASCMAIN1.NowTSD
        Next
    End Sub

    Private Sub ShowSelectedMatches()
        Dim GridFilter As String = ""
        If grdWBTCUST1.Selected.Rows.Count = 1 Then
            Select Case grdWBTCUST1.Selected.Rows(0).Cells.Item("STATUS").Value
                Case Is = "N", "C"
                    If chkMCompany.Checked Then
                        Dim COMPANY As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("COMPANY").Text, True)
                        If COMPANY.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_NAME LIKE '%{0}%'", COMPANY)
                            Else
                                GridFilter += String.Format(" {0} CUST_NAME LIKE '%{1}%'", IncludeFilter, COMPANY)
                            End If
                        End If
                    End If

                    If chkMEMail.Checked Then
                        Dim EMAIL As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("EMAIL").Text, True)
                        If EMAIL.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_EMAIL LIKE '%{0}%'", EMAIL)
                            Else
                                GridFilter += String.Format(" {0} CUST_EMAIL LIKE '%{1}%'", IncludeFilter, EMAIL)
                            End If
                        End If
                    End If

                    If chkMCustCode.Checked Then
                        Dim CUST_CODE_PROVIDED As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("CUST_CODE_PROVIDED").Text, True)
                        If CUST_CODE_PROVIDED.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_CODE = '{0}'", CUST_CODE_PROVIDED)
                            Else
                                GridFilter += String.Format(" {0} CUST_CODE = '{1}'", IncludeFilter, CUST_CODE_PROVIDED)
                            End If
                        End If
                    End If

                    If chkMStreet.Checked Then
                        Dim STREET As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("STREET").Text, True)
                        If STREET.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_ADDR1 LIKE '%{0}%'", STREET)
                            Else
                                GridFilter += String.Format(" {0} CUST_ADDR1 LIKE '%{1}%'", IncludeFilter, STREET)
                            End If
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_ADDR2 LIKE '%{0}%'", STREET)
                            Else
                                GridFilter += String.Format(" {0} CUST_ADDR2 LIKE '%{1}%'", IncludeFilter, STREET)
                            End If
                        End If
                    End If

                    If chkMCity.Checked Then
                        Dim CITY As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("CITY").Text, True)
                        If CITY.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_CITY LIKE '%{0}%'", CITY)
                            Else
                                GridFilter += String.Format(" {0} CUST_CITY LIKE '%{1}%'", IncludeFilter, CITY)
                            End If
                        End If
                    End If

                    If chkMState.Checked Then
                        Dim STATE As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("CITY").Text, True)
                        If STATE.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_STATE LIKE '%{0}%'", STATE)
                            Else
                                GridFilter += String.Format(" {0} CUST_STATE LIKE '%{1}%'", IncludeFilter, STATE)
                            End If
                        End If
                    End If

                    If chkMZip.Checked Then
                        Dim ZIP_CODE As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("CITY").Text, True)
                        If ZIP_CODE.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_ZIP_CODE LIKE '%{0}%'", ZIP_CODE)
                            Else
                                GridFilter += String.Format(" {0} CUST_ZIP_CODE LIKE '%{1}%'", IncludeFilter, ZIP_CODE)
                            End If
                        End If
                    End If

                    If chkMPhone.Checked Then
                        Dim TELEPHONE As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("TELEPHONE").Text, True)
                        If TELEPHONE.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_PHONE LIKE '%{0}%'", TELEPHONE)
                            Else
                                GridFilter += String.Format(" {0} CUST_PHONE LIKE '%{1}%'", IncludeFilter, TELEPHONE)
                            End If
                        End If
                    End If

                    If chkMLName.Checked Then
                        Dim FAMILYNAME As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("FAMILYNAME").Text, True)
                        If FAMILYNAME.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_CONTACT LIKE '%{0}%'", FAMILYNAME)
                            Else
                                GridFilter += String.Format(" {0} CUST_CONTACT LIKE '%{1}%'", IncludeFilter, FAMILYNAME)
                            End If
                        End If
                    End If

                    If chkMFName.Checked Then
                        Dim GIVENNAME As String = FormatBadChars(grdWBTCUST1.Selected.Rows(0).Cells.Item("GIVENNAME").Text, True)
                        If GIVENNAME.Length > 0 Then
                            If GridFilter.Length = 0 Then
                                GridFilter = String.Format("CUST_CONTACT LIKE '%{0}%'", GIVENNAME)
                            Else
                                GridFilter += String.Format(" {0} CUST_CONTACT LIKE '%{1}%'", IncludeFilter, GIVENNAME)
                            End If
                        End If
                    End If
                Case Is = "M"
                    Dim CONTACT_TYPE As String = ""
                    Select Case grdWBTCUST1.Selected.Rows(0).Cells.Item("CONTACT_TYPE").Text
                        Case Is = "1"
                            CONTACT_TYPE = "Customer"
                        Case Is = "2"
                            CONTACT_TYPE = "Ship-To"
                        Case Is = "D"
                            CONTACT_TYPE = "Contact"
                    End Select
                    GridFilter = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND REC_TYPE = '{2}'",
                                               grdWBTCUST1.Selected.Rows(0).Cells.Item("CUST_CODE_ACTUAL").Text,
                                               grdWBTCUST1.Selected.Rows(0).Cells.Item("CONTACT_NO").Text, CONTACT_TYPE)
                Case Is = "A"
                    Dim CONTACT_TYPE As String = ""
                    Select Case grdWBTCUST1.Selected.Rows(0).Cells.Item("CONTACT_TYPE").Text
                        Case Is = "1"
                            CONTACT_TYPE = "Customer"
                        Case Is = "2"
                            CONTACT_TYPE = "Ship-To"
                        Case Is = "D"
                            CONTACT_TYPE = "Contact"
                    End Select
                    GridFilter = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = {1} AND REC_TYPE = '{2}'",
                                               grdWBTCUST1.Selected.Rows(0).Cells.Item("CUST_CODE_ACTUAL").Text,
                                               grdWBTCUST1.Selected.Rows(0).Cells.Item("CONTACT_NO").Text, CONTACT_TYPE)
            End Select
        End If

        If Not IsNothing(grdARTCUSTX.DataSource) Then
            If GridFilter.Length = 0 Then
                GridFilter = "CUST_ZIP_CODE = '9999999999'"
            End If
            Dim dvw As DataView = DirectCast(grdARTCUSTX.DataSource, DataTable).DefaultView
            dst.Tables.Item("ARTCONTX").CaseSensitive = False
            dvw.RowFilter = GridFilter
            picFilter.Visible = (GridFilter.Length > 0)
        End If
    End Sub

    Private Function ProcessWordToHTML(ByRef rowWBTCUST1 As DataRow, ByVal TEMPLATE_NAME As String) As String
        Dim RPTV As List(Of REPORT_VARIABLES) = GetReportVariables(rowWBTCUST1, TEMPLATE_NAME)
        Dim REV As Integer = 0
        Dim TEMPLATE As String = String.Format("{0}EMAIL\Templates\{1}.doc", ASCMAIN1.Folders("Archive"), TEMPLATE_NAME)
        Dim W As New Microsoft.Office.Interop.Word.Application
        Dim WD As Microsoft.Office.Interop.Word.Document = W.Documents.Open(TEMPLATE)
        For Each V As REPORT_VARIABLES In RPTV
            FindAndReplace(WD, V.VAR_NAME, V.VAR_VALUE)
        Next
        WD.SaveAs(String.Format("{0}EMAIL\Templates\TEMP_HTML.doc", ASCMAIN1.Folders("Archive")), Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatHTML)
        WD.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges,
                      Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat)
        W.Quit()
        W = Nothing
        Dim HTMLString As String = New System.Net.WebClient().DownloadString(ASCMAIN1.Folders("Archive") & "EMAIL\Templates\TEMP_HTML.doc")
        Return HTMLString
    End Function

    Private Sub RejectContacts(ByRef rowsCUST As Infragistics.Win.UltraWinGrid.SelectedRowsCollection)
        For Each rowCUST As Infragistics.Win.UltraWinGrid.UltraGridRow In rowsCUST
            rowCUST.Cells.Item("STATUS").Value = "R"
            rowCUST.Cells.Item("CLAIM_BY_OPER").Value = Null
            rowCUST.Cells.Item("LAST_OPER").Value = ASCMAIN1.USER_ID
            rowCUST.Cells.Item("LAST_DATE").Value = Now + ASCMAIN1.NowTSD
        Next
    End Sub

    Private Sub UploadContacts(ByRef rowsCUST As Infragistics.Win.UltraWinGrid.SelectedRowsCollection)

        For Each rowCUST As Infragistics.Win.UltraWinGrid.UltraGridRow In rowsCUST
            Dim EMAIL As String = rowCUST.Cells.Item("EMAIL").Text.ToString & String.Empty
            Dim FILTER As String = String.Format("EMAIL = '{0}'", EMAIL)
            Dim rowDB As DataRow = dst.Tables.Item("WBTCUST1").Select(FILTER).FirstOrDefault
            If Not IsNothing(rowDB) Then
                rowDB.Item("STATUS") = "U"
                rowDB.Item("CLAIM_BY_OPER") = Null
                rowDB.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowDB.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            End If
            'rowCUST.Cells.Item("STATUS").Value = "U"
            'rowCUST.Cells.Item("CLAIM_BY_OPER").Value = Null
        Next
    End Sub

    Private Function SendEMail(ByRef rowWBTCUST1 As DataRow, ByVal TEMPLATE_NAME As String) As String
        'These all need to be parameterized before Wayne Retires.
        Const FROM_ADDRESS As String = "New.accounts@regency-rib.com"
        Const FROM_NAME As String = "New Accounts At Regency-rib.com"
        Const BCC_ADDRESS As String = "mariog@regency-rib.com"
        Const BCC_NAME As String = "Mario Arenas Jr."
        Const SERVER_IP As String = "192.168.110.221"
        Const SERVER_PORT As Integer = 25
        Const SERVER_ACCOUNT As String = "New.accounts@regency-rib.com"
        Const SERVER_PASSWORD As String = "0Ff1c3"
        Dim TO_SUBJECT As String = ""
        Select Case TEMPLATE_NAME
            Case Is = "ACCEPT"
                TO_SUBJECT = "Welcome To Regency"
            Case Is = "CREDIT"
                TO_SUBJECT = "Welcome To Regency"
            Case Else
                TO_SUBJECT = "Welcome To Regency"
        End Select

        Try
            If TEMPLATE_NAME.Length = 0 Then
                MsgBox("Could Not Determine Correct Report To Print", MsgBoxStyle.Critical, "Problem Printing Report")
                Return ""
                Exit Function
            End If

            Dim AddSalesRep As Boolean = False
            Dim ccSPREPName As String = ""
            Dim ccSREPEmail As String = ""
            Dim SREP_CODE As String = rowWBTCUST1.Item("SREP_CODE").ToString
            If SREP_CODE.Length > 0 Then
                Dim SQLC As StringBuilder = New StringBuilder() With {.Length = 0}
                SQLC.AppendLine("Select ")
                SQLC.AppendLine("NVL(SOTSREP1.SREP_NAME,'') SREP_NAME,")
                SQLC.AppendLine("NVL(SOTSREP1.SREP_EMAIL,'') SREP_EMAIL,")
                SQLC.AppendLine("NVL(SOTSREP1.SREP_CODE,'') SREP_CODE")
                SQLC.AppendLine("FROM SOTSREP1")
                SQLC.AppendLine(String.Format("WHERE SOTSREP1.SREP_CODE =  '{0}'", rowWBTCUST1.Item("SREP_CODE")))
                Dim tbl As DataTable = ASCDATA1.GetDataTable(SQLC.ToString())
                If tbl.Rows.Count > 0 Then
                    If tbl.Rows(0).Item("SREP_CODE").ToString.Length > 0 And tbl.Rows(0).Item("SREP_CODE").ToString <> "HO" Then
                        ccSPREPName = tbl.Rows(0).Item("SREP_NAME").ToString
                        ccSREPEmail = tbl.Rows(0).Item("SREP_EMAIL").ToString
                        AddSalesRep = True
                    End If
                End If
            End If

            EMAIL_ADDRESS = ""
            EMAIL_NAME = ""
            Dim HTMLBody As String = ProcessWordToHTML(rowWBTCUST1, TEMPLATE_NAME)
            If EMAIL_ADDRESS.Length = 0 Or EMAIL_NAME.Length = 0 Then
                MsgBox("Either The Email Address Or The Name Could Not Be Found", MsgBoxStyle.Critical, "E-Mail Not Sent")
            Else
                Dim mail As New MailMessage()
                mail.From = New MailAddress(FROM_ADDRESS, FROM_NAME)
                mail.To.Add(New MailAddress(EMAIL_ADDRESS, EMAIL_NAME))
                mail.Subject = TO_SUBJECT
                mail.IsBodyHtml = True
                mail.Body = HTMLBody
                mail.Bcc.Add(New MailAddress(BCC_ADDRESS, BCC_NAME))
                If AddSalesRep Then
                    mail.CC.Add(New MailAddress(ccSREPEmail, ccSPREPName))
                End If
                'mail.Attachments.Add(New Attachment(ss.Trim))

                Dim smtp As New SmtpClient(SERVER_IP, SERVER_PORT)
                If smtp IsNot Nothing Then
                    smtp.Credentials = New System.Net.NetworkCredential(SERVER_ACCOUNT, SERVER_PASSWORD)
                Else
                    Dim eMsg As String = "SMTP Client could not be created."
                    MsgBox(eMsg, MsgBoxStyle.OkOnly, "Error")
                    Return False
                End If

                If ASCMAIN1.Running_in_VS Then
                    Stop
                Else
                    smtp.Send(mail)
                End If

                'MsgBox("email has been sent", MsgBoxStyle.OkOnly, "Verification")
            End If

        Catch ex As Exception
            MsgBox("Error Trying to Generate Document" & vbCrLf & ex.Message)
        End Try
        Return ""
    End Function

#End Region

#Region "Inbound Customer Transfers"
    Private Function ImportCustomersFromShopsite() As Boolean
        Dim retVal As Boolean = False
        Me.Cursor = Cursors.WaitCursor
        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop

        Dim ErrMsg As New StringBuilder With {.Length = 0}
        Dim TempFolder As String = ASCMAIN1.Folders("Temp").ToString
        If Not TempFolder.EndsWith("\") Then
            TempFolder = TempFolder & "\"
        End If

        Dim LocalFile As String = String.Format("{0}{1}", TempFolder, inBoundFile)
        Dim FileFound As Boolean = False
        WebCustFTPInbound(ErrMsg, LocalFile, FileFound)
        If FileFound Then
            If ErrMsg.Length = 0 Then
                WebCustInboundCreate(ErrMsg, LocalFile)
            End If
            WebCustTMPDelete(ErrMsg, LocalFile)
            If ErrMsg.Length = 0 Then
                WebCustFTPDelete(ErrMsg, LocalFile)
                retVal = True
                UpdateAndRefreshData(False)
            Else
                retVal = False
                EnforceConstraints(False)
                Call Fill_Records("WBTCUST1")
                Call Fill_Records("WBTCUST2")
                Call Fill_Records("WBTCUST9")
                Call Fill_Records("ARTCONTX")
                EnforceConstraints(True)
                MsgBox(ErrMsg.ToString, vbExclamation, "Error Importing Customers")
            End If
        Else
            MsgBox("No InBound Customers Waiting On Shopsite.", vbOKOnly, "Nothing To Do Now")
        End If
        Me.Cursor = Cursors.Default
        Return retVal
    End Function

    Private Sub WebCustFTPInbound(ByRef ErrMsg As StringBuilder, ByVal LocalFile As String, ByRef FileFound As Boolean)
        'Download File From FTP to tmp
        FileFound = False
        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            If File.Exists(LocalFile) Then
                File.Delete(LocalFile)
            End If
            Try
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = UserName
                .Password = Password
                .RemoteHost = RemoteHost
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = LocalFile
                .RemoteFile = inBoundFile
                .Overwrite = False
                If Not .FileExists() Then
                    FileFound = False
                    .Logoff()
                Else
                    FileFound = True
                    .Download()
                    .Logoff()
                End If
            Catch ex As Exception
                FileFound = False
                ErrMsg.AppendLine(ex.Message.ToString)
                FtpShopSite.Logoff()
            End Try
        End With

    End Sub

    Private Sub WebCustInboundCreate(ByRef errMsg As StringBuilder, ByVal localFile As String)
        'Write to BATCH file, New Customer table and Encrypted Password File.
        Try
            Dim RecsUpdated As Int64 = 0
            Dim RecsAdded As Int64 = 0
            Dim RecWarnings As Int64 = 0
            Dim FoundStartRow As Boolean = False
            Dim FoundLastRow As Boolean = False
            Dim curRow As Int64 = 0
            Dim WEB_CUST_BATCH As String = ASCMAIN1.Next_Control_No("WBTCUST2.WEB_CUST_BATCH")
            Dim BATCHDATE As Date = Now()
            Dim BATCH_LNO As Int64 = 0
            Using MyReader As New FileIO.TextFieldParser(localFile)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim emptyCount As Int64 = 0
                While Not MyReader.EndOfData
                    curRow += 1
                    ASCMAIN1.Progress("Processing Row", curRow)
                    Dim currentRow As String()
                    currentRow = MyReader.ReadFields()
                    If curRow = 1 Then
                        If currentRow(0).ToString & String.Empty <> "Business Name" Then
                            errMsg.AppendLine("Inbound File Not In Correct Format")
                            Exit Sub
                        End If
                    Else
                        If currentRow(4).ToString & String.Empty = "" Then
                            emptyCount += 1
                        Else
                            BATCH_LNO += 1
                            Dim FILTER As String = String.Format("EMAIL = '{0}'", currentRow(4).ToString.ToUpper & String.Empty)
                            Dim rowWBTCUST1 As DataRow = dst.Tables.Item("WBTCUST1").Select(FILTER).FirstOrDefault
                            Dim newWBTCUST2 As DataRow = dst.Tables.Item("WBTCUST2").NewRow
                            newWBTCUST2.Item("WEB_CUST_BATCH") = WEB_CUST_BATCH
                            newWBTCUST2.Item("BATCHDATE") = BATCHDATE
                            newWBTCUST2.Item("BATCH_LNO") = BATCH_LNO
                            newWBTCUST2.Item("EMAIL") = currentRow(4).ToString.ToUpper & String.Empty

                            If Not IsNothing(rowWBTCUST1) Then
                                If currentRow(0).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("COMPANY") = currentRow(0).ToString & String.Empty
                                End If
                                If currentRow(1).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("GIVENNAME") = currentRow(1).ToString & String.Empty
                                End If
                                If currentRow(2).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("FAMILYNAME") = currentRow(2).ToString & String.Empty
                                End If
                                If (currentRow(1).ToString & String.Empty & " " & currentRow(2).ToString & String.Empty) <> "" Then
                                    newWBTCUST2.Item("FULLNAME") = currentRow(1).ToString & String.Empty & " " & currentRow(2).ToString & String.Empty
                                End If
                                If currentRow(3).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("TELEPHONE") = currentRow(3).ToString & String.Empty
                                End If
                                'If currentRow(5).ToString & String.Empty <> "" Then
                                '    rowWBTCUST9.Item("EMAIL") = currentRow(4).ToString.ToUpper & String.Empty 'Started Storing Passwords Per Mario 2/15/19
                                '    rowWBTCUST9.Item("PASSWORD") = currentRow(5).ToString & String.Empty 'Started Storing Passwords Per Mario 2/15/19
                                'End If
                                If currentRow(6).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("STREET") = currentRow(6).ToString & String.Empty
                                End If
                                If currentRow(7).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("STREET2") = currentRow(7).ToString & String.Empty
                                End If
                                If currentRow(8).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("STREET3") = currentRow(8).ToString & String.Empty
                                End If
                                If currentRow(9).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("CITY") = currentRow(9).ToString & String.Empty
                                End If
                                If currentRow(10).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("STATE") = currentRow(10).ToString & String.Empty
                                End If
                                If currentRow(11).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("ZIP_CODE") = currentRow(11).ToString & String.Empty
                                End If
                                If currentRow(12).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("COUNTRY") = currentRow(12).ToString & String.Empty
                                End If
                                If currentRow(14).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("TAX_ID") = currentRow(14).ToString & String.Empty
                                End If
                                If currentRow(16).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("WEBSITE") = currentRow(15).ToString & String.Empty
                                End If
                                If currentRow(17).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("BUSINESS_YEARS") = currentRow(16).ToString & String.Empty
                                End If
                                If currentRow(18).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("INTERESTS") = currentRow(17).ToString & String.Empty
                                End If
                                If currentRow(19).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("REFERRED") = currentRow(18).ToString & String.Empty
                                End If
                                If currentRow(20).ToString & String.Empty <> "" Then
                                    newWBTCUST2.Item("COMMENTS") = currentRow(19).ToString & String.Empty
                                End If
                                newWBTCUST2.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                                newWBTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                dst.Tables.Item("WBTCUST2").Rows.Add(newWBTCUST2)
                            Else
                                Dim newWBTCUST1 As DataRow = dst.Tables.Item("WBTCUST1").NewRow
                                newWBTCUST1.Item("COMPANY") = currentRow(0).ToString & String.Empty
                                newWBTCUST2.Item("COMPANY") = currentRow(0).ToString & String.Empty

                                newWBTCUST1.Item("GIVENNAME") = currentRow(1).ToString & String.Empty
                                newWBTCUST2.Item("GIVENNAME") = currentRow(1).ToString & String.Empty

                                newWBTCUST1.Item("FAMILYNAME") = currentRow(2).ToString & String.Empty
                                newWBTCUST2.Item("FAMILYNAME") = currentRow(2).ToString & String.Empty

                                newWBTCUST1.Item("FULLNAME") = currentRow(1).ToString & String.Empty & " " & currentRow(2).ToString & String.Empty
                                newWBTCUST2.Item("FULLNAME") = currentRow(1).ToString & String.Empty & " " & currentRow(2).ToString & String.Empty

                                newWBTCUST1.Item("TELEPHONE") = currentRow(3).ToString & String.Empty
                                newWBTCUST2.Item("TELEPHONE") = currentRow(3).ToString & String.Empty

                                newWBTCUST1.Item("EMAIL") = currentRow(4).ToString.ToUpper & String.Empty
                                newWBTCUST2.Item("EMAIL") = currentRow(4).ToString.ToUpper & String.Empty

                                If currentRow(5).ToString & String.Empty <> "" Then
                                    Dim newWBTCUST9 As DataRow = dst.Tables.Item("WBTCUST9").NewRow
                                    newWBTCUST9.Item("EMAIL") = currentRow(4).ToString.ToUpper & String.Empty 'Started Storing Passwords Per Mario 2/15/19
                                    'newWBTCUST9.Item("PASSWORD") = currentRow(5).ToString & String.Empty 'Started Storing Passwords Per Mario 2/15/19
                                    newWBTCUST9.Item("PSWDE") = psEncrypt(currentRow(5).ToString & String.Empty, EncryptType.Encrypt)
                                    dst.Tables.Item("WBTCUST9").Rows.Add(newWBTCUST9)
                                End If

                                newWBTCUST1.Item("STREET") = currentRow(6).ToString & String.Empty
                                newWBTCUST2.Item("STREET") = currentRow(6).ToString & String.Empty

                                newWBTCUST1.Item("STREET2") = currentRow(7).ToString & String.Empty
                                newWBTCUST2.Item("STREET2") = currentRow(7).ToString & String.Empty

                                newWBTCUST1.Item("STREET3") = currentRow(8).ToString & String.Empty
                                newWBTCUST2.Item("STREET3") = currentRow(8).ToString & String.Empty

                                newWBTCUST1.Item("CITY") = currentRow(9).ToString & String.Empty
                                newWBTCUST2.Item("CITY") = currentRow(9).ToString & String.Empty

                                newWBTCUST1.Item("STATE") = currentRow(10).ToString & String.Empty
                                newWBTCUST2.Item("STATE") = currentRow(10).ToString & String.Empty

                                newWBTCUST1.Item("ZIP_CODE") = currentRow(11).ToString & String.Empty
                                newWBTCUST2.Item("ZIP_CODE") = currentRow(11).ToString & String.Empty

                                newWBTCUST1.Item("COUNTRY") = currentRow(12).ToString & String.Empty
                                newWBTCUST2.Item("COUNTRY") = currentRow(12).ToString & String.Empty

                                newWBTCUST1.Item("CUST_CODE_PROVIDED") = currentRow(13).ToString & String.Empty
                                newWBTCUST2.Item("CUST_CODE_PROVIDED") = currentRow(13).ToString & String.Empty

                                newWBTCUST1.Item("TAX_ID") = currentRow(14).ToString & String.Empty
                                newWBTCUST2.Item("TAX_ID") = currentRow(14).ToString & String.Empty

                                newWBTCUST1.Item("WEBSITE") = currentRow(15).ToString & String.Empty
                                newWBTCUST2.Item("WEBSITE") = currentRow(15).ToString & String.Empty

                                newWBTCUST1.Item("BUSINESS_YEARS") = currentRow(16).ToString & String.Empty
                                newWBTCUST2.Item("BUSINESS_YEARS") = currentRow(16).ToString & String.Empty

                                newWBTCUST1.Item("INTERESTS") = currentRow(17).ToString & String.Empty
                                newWBTCUST2.Item("INTERESTS") = currentRow(17).ToString & String.Empty

                                newWBTCUST1.Item("REFERRED") = currentRow(18).ToString & String.Empty
                                newWBTCUST2.Item("REFERRED") = currentRow(18).ToString & String.Empty

                                newWBTCUST1.Item("COMMENTS") = currentRow(19).ToString & String.Empty
                                newWBTCUST2.Item("COMMENTS") = currentRow(19).ToString & String.Empty

                                newWBTCUST1.Item("DATEREGISTERED") = Now + ASCMAIN1.NowTSD
                                newWBTCUST2.Item("DATEREGISTERED") = Now + ASCMAIN1.NowTSD

                                newWBTCUST1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                                newWBTCUST2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD

                                newWBTCUST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                newWBTCUST2.Item("INIT_OPER") = ASCMAIN1.USER_ID

                                newWBTCUST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                                newWBTCUST2.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD

                                newWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                newWBTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID

                                newWBTCUST1.Item("STATUS") = "N"
                                newWBTCUST2.Item("STATUS") = "N"

                                dst.Tables.Item("WBTCUST1").Rows.Add(newWBTCUST1)
                                dst.Tables.Item("WBTCUST2").Rows.Add(newWBTCUST2)
                            End If
                        End If
                    End If
                End While
            End Using
        Catch ex As Exception
            errMsg.AppendLine(ex.Message.ToString)
        End Try
    End Sub

    Private Sub WebCustFTPDelete(ByRef errMsg As StringBuilder, ByVal localFile As String)
        'Delete FTP file.
        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = UserName
                .Password = Password
                .RemoteHost = RemoteHost
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = localFile
                .RemoteFile = inBoundFile
                .Overwrite = False
                If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    Stop
                End If
                If Not .FileExists() Then
                    errMsg.AppendLine("No Customer File To Delete On ShopSite.")
                    .Logoff()
                Else
                    .DeleteFile(inBoundFile)
                    .Logoff()
                End If
            Catch ex As Exception
                errMsg.AppendLine(ex.Message.ToString)
            End Try
        End With
    End Sub

#End Region

#Region "Outbound Customer Transfers"
    Private Function ExportCustomersToShopsite(Optional ByVal RefreshAll As Boolean = False) As Boolean
        Dim RetVal As Boolean = False
        Dim ErrMsg As New StringBuilder With {.Length = 0}
        Dim TempFolder As String = ASCMAIN1.Folders("Temp").ToString
        If Not TempFolder.EndsWith("\") Then
            TempFolder = TempFolder & "\"
        End If
        Dim LocalFile As String = String.Format("{0}{1}", TempFolder, OutBoundFile)
        If File.Exists(LocalFile) Then
            File.Delete(LocalFile)
        End If

        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop

        If ErrMsg.Length = 0 Then
            WebCustOutboundCheck(ErrMsg, LocalFile)
        End If

        If ErrMsg.Length = 0 Then
            WebCustOutboundCreate(ErrMsg, LocalFile, RefreshAll)
        End If

        If ErrMsg.Length = 0 Then
            WebCustOutboundSend(ErrMsg, LocalFile)
        End If
        WebCustTMPDelete(ErrMsg, LocalFile)
        If ErrMsg.Length = 0 Then
            RetVal = True
            UpdateAndRefreshData(False)
        Else
            RetVal = False
            EnforceConstraints(False)
            Call Fill_Records("WBTCUST1")
            Call Fill_Records("WBTCUST2")
            Call Fill_Records("WBTCUST9")
            Call Fill_Records("ARTCONTX")
            EnforceConstraints(True)
            MsgBox(ErrMsg.ToString, vbExclamation, "Error Sending Customers")
        End If
        Return RetVal
    End Function

    Private Function ExportCustomersToShopsiteTesting(Optional ByVal RefreshAll As Boolean = False) As Boolean
        Dim RetVal As Boolean = False
        Dim ErrMsg As New StringBuilder With {.Length = 0}
        Dim TempFolder As String = ASCMAIN1.Folders("Temp").ToString
        If Not TempFolder.EndsWith("\") Then
            TempFolder = TempFolder & "\"
        End If
        Dim LocalFile As String = String.Format("{0}{1}", TempFolder, OutBoundFile)
        If File.Exists(LocalFile) Then
            File.Delete(LocalFile)
        End If

        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop

        If ErrMsg.Length = 0 Then
            WebCustOutboundCheck(ErrMsg, LocalFile)
        End If

        If ErrMsg.Length = 0 Then
            WebCustOutboundCreate(ErrMsg, LocalFile, RefreshAll)
        End If

        If ErrMsg.Length = 0 Then
            WebCustOutboundSend(ErrMsg, LocalFile)
        End If
        WebCustTMPDelete(ErrMsg, LocalFile)
        If ErrMsg.Length = 0 Then
            RetVal = True
            UpdateAndRefreshData(False)
        Else
            RetVal = False
            EnforceConstraints(False)
            Call Fill_Records("WBTCUST1")
            Call Fill_Records("WBTCUST2")
            Call Fill_Records("WBTCUST9")
            Call Fill_Records("ARTCONTX")
            EnforceConstraints(True)
            MsgBox(ErrMsg.ToString, vbExclamation, "Error Sending Customers")
        End If
        Return RetVal
    End Function

    Private Sub WebCustOutboundCreate(errMsg As StringBuilder, localFile As String, refreshAll As Boolean)
        Try
            Dim Retval As Boolean = False
            Dim str As New StringBuilder
            Dim sql As New StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT")
            sql.AppendLine(String.Format("A1.CUST_CODE {0}Regency Account #{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_NAME {0}Business Name{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.GIVENNAME {0}First Name{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.FAMILYNAME {0}Last Name{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.FULLNAME {0}Contact Name{0},", Chr(34)))
            sql.AppendLine(String.Format("'' {0}Contact Number{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.EMAIL {0}Email Address{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.PASSWORD {0}Password{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_ADDR1 {0}Business Address Line 1{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_ADDR2 {0}Business Address Line 2{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_CITY {0}City{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_STATE {0}State{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_ZIP_CODE {0}Zip Code{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_COUNTRY {0}Country{0},", Chr(34)))
            sql.AppendLine(String.Format("'00' {0}Price Group{0},", Chr(34)))
            sql.AppendLine(String.Format("'1' {0}Welcome Type{0},", Chr(34)))
            sql.AppendLine(String.Format("'1' {0}Terms{0}", Chr(34)))
            sql.AppendLine("FROM WBTCUST1 W1, ARTCUST1 A1")
            sql.AppendLine("WHERE W1.CUST_CODE_ACTUAL = A1.CUST_CODE")
            If refreshAll Then
                sql.AppendLine("AND (STATUS = 'A' OR STATUS = 'U')")
            Else
                sql.AppendLine("AND STATUS = 'U'")
            End If
            sql.AppendLine("ORDER BY W1.FULLNAME")
            'str.Append(Chr(34))
            Dim tblAccounts As DataTable = ASCDATA1.GetDataTable(sql.ToString())
            For Each dc As DataColumn In tblAccounts.Columns
                str.Append(Chr(34) & dc.ColumnName.ToString & Chr(34) & ",")
            Next
            str.Replace(",", vbNewLine, str.Length - 1, 1)
            tblAccounts.Columns.Item("Price Group").ReadOnly = False
            tblAccounts.Columns.Item("Terms").ReadOnly = False
            For Each rowACCOUNTS As DataRow In tblAccounts.Rows
                Dim EMAIL As String = rowACCOUNTS.Item("Email Address").ToString.ToUpper & String.Empty
                Dim CUST_CODE As String = rowACCOUNTS.Item("Regency Account #").ToString.ToUpper & String.Empty
                rowACCOUNTS.Item("PASSWORD") = GetCustPassword(EMAIL)
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If Not IsNothing(rowARTCUST1) Then
                    rowACCOUNTS.Item("Price Group") = CalculatepriceGroup(rowARTCUST1)
                    rowACCOUNTS.Item("Terms") = CalculateTerms(rowARTCUST1)
                End If
                If refreshAll Then
                    Dim colIndex As Integer = 0
                    For Each field As Object In rowACCOUNTS.ItemArray
                        If colIndex = 6 Or colIndex = 14 Or colIndex = 16 Then
                            str.Append(Chr(34) & field.ToString & Chr(34) & ",")
                        Else
                            str.Append(Chr(34) & "" & Chr(34) & ",")
                        End If
                        colIndex += 1
                    Next
                    str.Replace(",", vbNewLine, str.Length - 1, 1)
                Else
                    Dim Filter As String = String.Format("EMAIL = '{0}'", EMAIL)
                    Dim rowWBTCUST1 As DataRow = dst.Tables.Item("WBTCUST1").Select(Filter).FirstOrDefault
                    If Not IsNothing(rowWBTCUST1) Then
                        rowWBTCUST1.Item("STATUS") = "A"
                    End If
                    For Each field As Object In rowACCOUNTS.ItemArray
                        str.Append(Chr(34) & field.ToString & Chr(34) & ",")
                    Next
                    str.Replace(",", vbNewLine, str.Length - 1, 1)
                End If
            Next

            Dim WEB_CUST_BATCH As String = ASCMAIN1.Next_Control_No("WBTCUST3.WEB_CUST_BATCH")
            Dim BATCHDATE As Date = Now()
            Dim BATCH_LNO As Int64 = 0
            For Each rowACCOUNTS As DataRow In tblAccounts.Rows
                BATCH_LNO += 1
                Dim newWBTCUST3 As DataRow = dst.Tables.Item("WBTCUST3").NewRow
                newWBTCUST3.Item("WEB_CUST_BATCH") = WEB_CUST_BATCH
                newWBTCUST3.Item("BATCH_LNO") = BATCH_LNO
                newWBTCUST3.Item("BATCHDATE") = BATCHDATE
                newWBTCUST3.Item("CUST_CODE") = rowACCOUNTS.Item("Regency Account #").ToString & String.Empty
                newWBTCUST3.Item("CUST_NAME") = rowACCOUNTS.Item("Business Name").ToString & String.Empty
                newWBTCUST3.Item("GIVENNAME") = rowACCOUNTS.Item("First Name").ToString & String.Empty
                newWBTCUST3.Item("FAMILYNAME") = rowACCOUNTS.Item("Last Name").ToString & String.Empty
                newWBTCUST3.Item("FULLNAME") = rowACCOUNTS.Item("Contact Name").ToString & String.Empty
                newWBTCUST3.Item("CONTACT_NUMBER") = rowACCOUNTS.Item("Contact Number").ToString & String.Empty
                newWBTCUST3.Item("EMAIL") = rowACCOUNTS.Item("Email Address").ToString & String.Empty
                newWBTCUST3.Item("PASSWORD") = "" 'rowACCOUNTS.Item("Password").ToString & String.Empty
                newWBTCUST3.Item("CUST_ADDR1") = rowACCOUNTS.Item("Business Address Line 1").ToString & String.Empty
                newWBTCUST3.Item("CUST_ADDR2") = rowACCOUNTS.Item("Business Address Line 2").ToString & String.Empty
                newWBTCUST3.Item("CUST_CITY") = rowACCOUNTS.Item("City").ToString & String.Empty
                newWBTCUST3.Item("CUST_STATE") = rowACCOUNTS.Item("State").ToString & String.Empty
                newWBTCUST3.Item("CUST_ZIP_CODE") = rowACCOUNTS.Item("Zip Code").ToString & String.Empty
                newWBTCUST3.Item("CUST_COUNTRY") = rowACCOUNTS.Item("Country").ToString & String.Empty
                newWBTCUST3.Item("PRICE_GROUP") = rowACCOUNTS.Item("Price Group").ToString & String.Empty
                newWBTCUST3.Item("WELCOME_TYPE") = rowACCOUNTS.Item("Welcome Type").ToString & String.Empty
                newWBTCUST3.Item("TERMS") = rowACCOUNTS.Item("Terms").ToString & String.Empty
                dst.Tables.Item("WBTCUST3").Rows.Add(newWBTCUST3)
            Next
            Try
                My.Computer.FileSystem.WriteAllText(localFile, str.ToString, False)
                Retval = True
            Catch ex As Exception
                MsgBox("Error Creating Output File", vbExclamation, "Error")
                Retval = False
            End Try
        Catch ex As Exception
            errMsg.AppendLine(ex.Message.ToString)
        End Try
    End Sub

    Private Sub WebCustOutboundSend(errMsg As StringBuilder, localFile As String)
        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                If .Connected = True Then
                    .Logoff()
                End If
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = UserName
                .Password = Password
                .RemoteHost = RemoteHost
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = localFile
                .RemoteFile = OutBoundFile
                .Overwrite = False
                If Not .FileExists() Then
                    .Upload()
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                End If
            Catch ex As Exception
                errMsg.AppendLine(ex.Message.ToString)
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
            End Try
        End With
    End Sub

    Private Sub WebCustOutboundCheck(errMsg As StringBuilder, ByVal localFile As String)
        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = UserName
                .Password = Password
                .RemoteHost = RemoteHost
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = localFile
                .RemoteFile = OutBoundFile
                .Overwrite = False
                If .FileExists() Then
                    errMsg.AppendLine("New Customer File Still Waiting On ShopSite.")
                    .DoEvents()
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                End If
            Catch ex As Exception
                errMsg.AppendLine(ex.Message.ToString)
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
            End Try
        End With

    End Sub

    Private Sub WebCustTMPDelete(ByRef errMsg As StringBuilder, ByVal localFile As String)
        'Delete tmp file.
        Try
            If File.Exists(localFile) Then
                File.Delete(localFile)
            Else
                errMsg.AppendLine("No Local File Found To Delete.")
            End If
        Catch ex As Exception
            errMsg.AppendLine(ex.Message.ToString)
        End Try
    End Sub
#End Region

#Region "Encryption"
    Private Sub EncryptPasswords()
        For Each rowWBTCUST9 As DataRow In dst.Tables("WBTCUST9").Select()
            Dim EMAIL As String = rowWBTCUST9.Item("EMAIL").ToString & String.Empty
            Dim PASSWORD As String = rowWBTCUST9.Item("PASSWORD").ToString & String.Empty
            Dim PSWDE As String = psEncrypt(PASSWORD, EncryptType.Encrypt)
            Dim PSWDD As String = psEncrypt(PSWDE, EncryptType.Decrypt)
            If PSWDD = PASSWORD Then
                rowWBTCUST9.Item("PSWDE") = PSWDE
            End If
        Next
    End Sub

    Private Function psEncrypt(ByVal PASSWORD As String, ByVal EType As EncryptType) As String
        Dim RetVal As String = ""
        Dim h As New TAC.ASCSCRTY
        Select Case EType
            Case EncryptType.Encrypt
                RetVal = h.Encrypt_AES(PASSWORD)
            Case EncryptType.Decrypt
                RetVal = h.Decrypt_AES(PASSWORD)
        End Select
        Return RetVal
    End Function

    Private Sub rdoAutoMatched_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAutoMatched.CheckedChanged

    End Sub

#End Region

#Region "Old Space Code"
    'Private Sub CreateCustomerUpload()
    '    Dim FileContent As String = ""
    '    FileContent += "#E-mail Address" & vbTab
    '    FileContent += "Group Name" & vbCrLf

    '    For Each rowWBTCUST1 As DataRow In dst.Tables("WBTCUST1").Select("STATUS = 'U'")
    '        Dim GroupCode As String = ""
    '        GroupCode = GetCustGroupCode(rowWBTCUST1.Item("CUST_CODE_ACTUAL").ToString & "")
    '        FileContent += rowWBTCUST1.Item("EMAIL").ToString.ToLower & "" & vbTab
    '        FileContent += GroupCode & vbCrLf
    '        rowWBTCUST1.Item("STATUS") = "A"
    '    Next

    '    Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
    '    SQLS.AppendLine("SELECT WB_PARM_UPLOAD_DIR FROM WBTPARM1 WHERE WB_PARM_KEY = 'Z'")
    '    ASCMAIN1.sql = SQLS.ToString()
    '    Dim WB_PARM_UPLOAD_DIR As String = ASCDATA1.GetDataValue
    '    Dim FileCount As Integer = 1
    '    Dim DATEFORMAT As String = Format(Now, "yymmdd")
    '    Dim FileUpload As String = String.Format("{0}\CustomerUpload_{1}_{2}.txt", WB_PARM_UPLOAD_DIR, DATEFORMAT, FileCount)
    '    Do While IO.File.Exists(FileUpload)
    '        FileCount += 1
    '        FileUpload = String.Format("{0}\CustomerUpload_{1}_{2}.txt", WB_PARM_UPLOAD_DIR, DATEFORMAT, FileCount)
    '    Loop
    '    If WB_PARM_UPLOAD_DIR.Length = 0 Then
    '        MsgBox("Missing Upload Folder In Parameters", MsgBoxStyle.Critical, "Parameter Error")
    '        Exit Sub
    '    End If
    '    If Not IO.Directory.Exists(WB_PARM_UPLOAD_DIR) Then
    '        MsgBox("Missing Upload Folder In Parameters", MsgBoxStyle.Critical, "Parameter Error")
    '        Exit Sub
    '    End If
    '    Dim FileWrite As IO.FileStream = IO.File.Create(FileUpload)
    '    FileWrite.Close()
    '    IO.File.WriteAllText(FileUpload, FileContent)
    '    MsgBox("The File " & FileUpload & " Was Created For Upload", MsgBoxStyle.Information, "Upload This To Shopsite Before You Forget!")
    'End Sub

    'Private Sub MatchAllEMails()
    '    Dim RecCnt As Integer = 0
    '    For Each rowWBTCUST1 As DataRow In dst.Tables("WBTCUST1").Select()
    '        If rowWBTCUST1.Item("STATUS") & "" = "N" Then
    '            For Each rowARTCONTX As DataRow In dst.Tables("ARTCONTX").Select()
    '                If rowWBTCUST1.Item("EMAIL").ToString.ToUpper = rowARTCONTX.Item("CUST_EMAIL").ToString.ToUpper & "" Then
    '                    MatchContacts(rowWBTCUST1, rowARTCONTX)
    '                    RecCnt += 1
    '                End If
    '            Next
    '        End If
    '    Next
    '    If RecCnt = 0 Then
    '        MsgBox("No Matching Records Found", MsgBoxStyle.Information, "Match All EMails")
    '    Else
    '        MsgBox(String.Format("{0} Matching Record(s) Found", RecCnt), MsgBoxStyle.Information, "Match All EMails")
    '    End If
    'End Sub

    'Private Sub Import_Shopsite_IDs()
    '    Dim fDialog As New OpenFileDialog
    '    fDialog.Filter = "Comma Delimited Files|*.csv"
    '    fDialog.Title = "Select an Excel File To Import"
    '    If fDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
    '        Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
    '        Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Add
    '        Dim XWS As Microsoft.Office.Interop.Excel.Worksheet = XWB.Sheets(1)
    '        Dim FullFileName As String = fDialog.FileNames(0)
    '        Try
    '            XWB = excel.Workbooks.Open(FullFileName)
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
    '            XWB.Close()
    '            XWB = Nothing
    '            excel = Nothing
    '            Exit Sub
    '        End Try
    '        Me.Cursor = Cursors.WaitCursor
    '        XWS = XWB.Worksheets(1)
    '        Dim RecsUpdated As Int64 = 0
    '        Dim RecsNotUpdated As Int64 = 0
    '        Dim RecWarnings As Int64 = 0
    '        For i As Integer = 1 To 20000
    '            If (i Mod 100) = 0 Then
    '                ASCMAIN1.Progress("Updating", i)
    '            End If
    '            Dim SHOPSITE_ID As String = XWS.Cells(i, 2).text.ToString & ""
    '            Dim E_MAIL As String = XWS.Cells(i, 11).text.ToString & ""
    '            If SHOPSITE_ID = "" Then
    '                Exit For
    '            Else
    '                Dim filter As String = String.Format("EMAIL = '{0}'", E_MAIL)
    '                Dim rowWBTCUST1 As DataRow = dst.Tables.Item("WBTCUST1").Select(filter).FirstOrDefault
    '                If IsNothing(rowWBTCUST1) Then
    '                    RecsNotUpdated += 1
    '                Else
    '                    If rowWBTCUST1.Item("SHOPSITE_CUST_ID").ToString & String.Empty = "" Then
    '                        rowWBTCUST1.Item("SHOPSITE_CUST_ID") = SHOPSITE_ID
    '                        RecsUpdated += 1
    '                    Else
    '                        If rowWBTCUST1.Item("SHOPSITE_CUST_ID").ToString & String.Empty <> SHOPSITE_ID Then
    '                            RecWarnings += 1
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next
    '        If RecsUpdated > 0 Then
    '            Update_Record_TDA("WBTCUST1")
    '        End If
    '        XWB.Close()
    '        XWB = Nothing
    '        excel = Nothing
    '        Me.Cursor = Cursors.Default
    '        Dim iMsg As New System.Text.StringBuilder With {.Length = 0}
    '        iMsg.AppendLine(String.Format("Records Updated: {0}", RecsUpdated))
    '        iMsg.AppendLine(String.Format("Records Not Updated: {0}", RecsNotUpdated))
    '        iMsg.AppendLine(String.Format("Warnings: {0}", RecWarnings))
    '        MsgBox(iMsg.ToString, vbOKOnly, "Import Complete!")
    '    End If
    'End Sub

    'Private Sub ImportFile(ByVal FileName As String)
    '    Dim NewRecsFound As Integer = 0
    '    Dim ErrMsg As String = ""
    '    If Not IO.File.Exists(FileName) Then
    '        ErrMsg = ErrMsg & vbCrLf & "Invalid File Specified"
    '    End If
    '    If Not FileName.EndsWith(".xls") Then
    '        ErrMsg = ErrMsg & vbCrLf & "File Must Be An XLS File Type"
    '    End If
    '    If ErrMsg.ToString.Length > 0 Then
    '        MsgBox(ErrMsg, MsgBoxStyle.Critical, "File Error")
    '    Else
    '        Using MyReader As New Microsoft.VisualBasic.FileIO.
    '            TextFieldParser(FileName)

    '            MyReader.TextFieldType =
    '                Microsoft.VisualBasic.FileIO.FieldType.Delimited
    '            MyReader.Delimiters = New String() {vbTab}
    '            Dim currentRow As String()
    '            'Loop through all of the fields in the file.  
    '            'If any lines are corrupt, report an error and continue parsing.
    '            currentRow = MyReader.ReadFields()
    '            Dim FilePositions As New List(Of LexMap)
    '            FilePositions = GetFilePositions(currentRow)
    '            While Not MyReader.EndOfData
    '                Try
    '                    currentRow = MyReader.ReadFields()
    '                    NewRecsFound += AddCurrentRow(currentRow, FilePositions)
    '                    ' Include code here to handle the row. 
    '                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
    '                    MsgBox("Line " & ex.Message &
    '                    " is invalid.  Skipping")
    '                End Try
    '            End While
    '            MsgBox(NewRecsFound & " New Contacts Added", MsgBoxStyle.Information, "New Records Found")
    '        End Using
    '    End If

    'End Sub

    'Private Function ReadFTP(ByVal localFile As String) As Boolean
    '    Dim RetVal As Boolean = False
    '    Dim RecsUpdated As Int64 = 0
    '    Dim RecsAdded As Int64 = 0
    '    Dim RecWarnings As Int64 = 0
    '    Dim FoundStartRow As Boolean = False
    '    Dim FoundLastRow As Boolean = False
    '    Dim curRow As Int64 = 0
    '    Using MyReader As New FileIO.TextFieldParser(localFile)
    '        MyReader.TextFieldType = FileIO.FieldType.Delimited
    '        MyReader.SetDelimiters(",")
    '        Dim emptyCount As Int64 = 0
    '        While Not MyReader.EndOfData
    '            curRow += 1
    '            ASCMAIN1.Progress("Processing Row", curRow)
    '            Dim currentRow As String()
    '            currentRow = MyReader.ReadFields()
    '            If curRow = 1 Then
    '                If currentRow(0).ToString & String.Empty <> "Business Name" Then
    '                    MsgBox("Inbound File Not In Correct Format", vbExclamation, "Problem With File")
    '                    RetVal = False
    '                    Return RetVal
    '                    Exit Function
    '                End If
    '            Else
    '                If currentRow(4).ToString & String.Empty = "" Then
    '                    emptyCount += 1
    '                Else
    '                    Dim FILTER As String = String.Format("EMAIL = '{0}'", currentRow(4).ToString.ToUpper & String.Empty)
    '                    Dim rowWBTCUST1 As DataRow = dst.Tables.Item("WBTCUST1").Select(FILTER).FirstOrDefault
    '                    If Not IsNothing(rowWBTCUST1) Then
    '                        If currentRow(0).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("COMPANY") = currentRow(0).ToString & String.Empty
    '                        End If
    '                        If currentRow(1).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("GIVENNAME") = currentRow(1).ToString & String.Empty
    '                        End If
    '                        If currentRow(2).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("FAMILYNAME") = currentRow(2).ToString & String.Empty
    '                        End If
    '                        If (currentRow(1).ToString & String.Empty & " " & currentRow(2).ToString & String.Empty) <> "" Then
    '                            rowWBTCUST1.Item("FULLNAME") = currentRow(1).ToString & String.Empty & " " & currentRow(2).ToString & String.Empty
    '                        End If
    '                        If currentRow(3).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("TELEPHONE") = currentRow(3).ToString & String.Empty
    '                        End If
    '                        If currentRow(5).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("PASSWORD") = currentRow(5).ToString & String.Empty 'Started Storing Passwords Per Mario 2/15/19
    '                        End If
    '                        If currentRow(6).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("STREET") = currentRow(6).ToString & String.Empty
    '                        End If
    '                        If currentRow(7).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("STREET2") = currentRow(7).ToString & String.Empty
    '                        End If
    '                        If currentRow(8).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("STREET3") = currentRow(8).ToString & String.Empty
    '                        End If
    '                        If currentRow(9).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("CITY") = currentRow(9).ToString & String.Empty
    '                        End If
    '                        If currentRow(10).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("STATE") = currentRow(10).ToString & String.Empty
    '                        End If
    '                        If currentRow(11).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("ZIP_CODE") = currentRow(11).ToString & String.Empty
    '                        End If
    '                        If currentRow(12).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("COUNTRY") = currentRow(12).ToString & String.Empty
    '                        End If
    '                        If currentRow(14).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("TAX_ID") = currentRow(14).ToString & String.Empty
    '                        End If
    '                        If currentRow(15).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("WEBSITE") = currentRow(15).ToString & String.Empty
    '                        End If
    '                        If currentRow(16).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("BUSINESS_YEARS") = currentRow(16).ToString & String.Empty
    '                        End If
    '                        If currentRow(17).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("INTERESTS") = currentRow(17).ToString & String.Empty
    '                        End If
    '                        If currentRow(18).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("REFERRED") = currentRow(18).ToString & String.Empty
    '                        End If
    '                        If currentRow(19).ToString & String.Empty <> "" Then
    '                            rowWBTCUST1.Item("COMMENTS") = currentRow(19).ToString & String.Empty
    '                        End If
    '                        rowWBTCUST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
    '                        rowWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
    '                    Else
    '                        Dim newWBTCUST1 As DataRow = dst.Tables.Item("WBTCUST1").NewRow
    '                        newWBTCUST1.Item("COMPANY") = currentRow(0).ToString & String.Empty
    '                        newWBTCUST1.Item("GIVENNAME") = currentRow(1).ToString & String.Empty
    '                        newWBTCUST1.Item("FAMILYNAME") = currentRow(2).ToString & String.Empty
    '                        newWBTCUST1.Item("FULLNAME") = currentRow(1).ToString & String.Empty & " " & currentRow(2).ToString & String.Empty
    '                        newWBTCUST1.Item("TELEPHONE") = currentRow(3).ToString & String.Empty
    '                        newWBTCUST1.Item("EMAIL") = currentRow(4).ToString.ToUpper & String.Empty
    '                        newWBTCUST1.Item("PASSWORD") = currentRow(5).ToString & String.Empty 'Started Storing Passwords Per Mario 2/15/19
    '                        newWBTCUST1.Item("STREET") = currentRow(6).ToString & String.Empty
    '                        newWBTCUST1.Item("STREET2") = currentRow(7).ToString & String.Empty
    '                        newWBTCUST1.Item("STREET3") = currentRow(8).ToString & String.Empty
    '                        newWBTCUST1.Item("CITY") = currentRow(9).ToString & String.Empty
    '                        newWBTCUST1.Item("STATE") = currentRow(10).ToString & String.Empty
    '                        newWBTCUST1.Item("ZIP_CODE") = currentRow(11).ToString & String.Empty
    '                        newWBTCUST1.Item("COUNTRY") = currentRow(12).ToString & String.Empty
    '                        newWBTCUST1.Item("CUST_CODE_PROVIDED") = currentRow(13).ToString & String.Empty
    '                        newWBTCUST1.Item("TAX_ID") = currentRow(14).ToString & String.Empty
    '                        newWBTCUST1.Item("WEBSITE") = currentRow(15).ToString & String.Empty
    '                        newWBTCUST1.Item("BUSINESS_YEARS") = currentRow(16).ToString & String.Empty
    '                        newWBTCUST1.Item("INTERESTS") = currentRow(17).ToString & String.Empty
    '                        newWBTCUST1.Item("REFERRED") = currentRow(18).ToString & String.Empty
    '                        newWBTCUST1.Item("COMMENTS") = currentRow(19).ToString & String.Empty
    '                        newWBTCUST1.Item("DATEREGISTERED") = Now + ASCMAIN1.NowTSD
    '                        newWBTCUST1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
    '                        newWBTCUST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
    '                        newWBTCUST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
    '                        newWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
    '                        newWBTCUST1.Item("STATUS") = "N"
    '                        dst.Tables.Item("WBTCUST1").Rows.Add(newWBTCUST1)
    '                    End If
    '                End If
    '            End If
    '        End While
    '        If emptyCount > 0 Then
    '            MsgBox("Import File Contained Blank Emails" & vbCrLf & "These Lines Were Skipped", vbOKOnly, "Warning!")
    '        End If
    '    End Using
    '    Return RetVal
    'End Function
#End Region
End Class

Public Class REPORT_VARIABLES
    Public VAR_NAME As String
    Public VAR_VALUE As String
End Class

Public Class LexMap
    Public ABS_COL As String
    Public LEX_COL As String
    Public COL_INDEX As Integer
End Class