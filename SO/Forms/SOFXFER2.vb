Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Http.Formatting
'Imports Newtonsoft.Json
Imports System.Reflection
Imports nsoftware.IPWorksSSH
Imports Newtonsoft.Json

Public Class SOFXFER2
    Private Event OnDirList As nsoftware.IPWorks.Ftp.OnDirListHandler
    'Private Event OnDirListS As nsoftware.IPWorksSSH.Sftp.OnDirListHandler
    ' Dim sqlSOTWORK1 As String
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.FTP
    Dim WithEvents FtpS As New nsoftware.IPWorksSSH.SFTPClient ' Sftp

    Dim PrintSelected As Boolean = False
    Dim RefreshRequired As Boolean = False
    Dim FileList As New Dictionary(Of String, String)
    Dim FileListS As New Dictionary(Of String, String)
    Dim ImageList As New Dictionary(Of String, String)
    Dim FTPImages As Boolean = False
    Dim ImageListLocal As New Dictionary(Of String, Date)
    Dim ImageListFTP As New Dictionary(Of String, Date)
    Dim ImageListDownload As New List(Of String)
    Dim ImageListDelete As New List(Of String)
    Dim BackupComplete As Boolean = False
    Private VersionInfo As New Text.StringBuilder With {.Length = 0}
    Dim ServerLive As String = "https://api2.regency-rib.com:8086/"
    Dim ServerTest As String = "http://localhost:1977/"
    'Dim ServerLive As String = "http://api.regency-rib.com:8181/"
    'Dim ServerTest As String = "http://kreativekode.ngrok.io/"
    'Dim API_CONTROLLER_ORDERS As String = "api/SalesOrder/CreateSalesOrder"
    'Dim API_CONTROLLER_CUSTOMERS As String = "api/Customer/CreateCustomer"
    'Dim API_CONTROLLER_SHIPTO As String = "api/Customer/UpdateShipTo"
    'Dim API_CONTROLLER_TATCTLN1 As String = "api/ABS/GetTATCTLN1"

    Dim ctlPfx As String = "api/RGI/LT"
    Dim API_CONTROLLER_ORDERS As String = $"{ctlPfx}/CreateSalesOrder"
    Dim API_CONTROLLER_CUSTOMERS As String = $"{ctlPfx}/CreateCustomer"
    Dim API_CONTROLLER_SHIPTO As String = $"{ctlPfx}/UpdateShipTo" 'api/RGI/LT/UpdateShipTo
    Dim API_CONTROLLER_TATCTLN1 As String = $"{ctlPfx}/GetTATCTLN1"

    Dim SB As New System.Text.StringBuilder With {.Length = 0}
    Dim ImageGetProgress As Boolean = True

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARMR")

        setVersionNo()

        With dst

            With .Tables.Add("TATCTLN0")
                .Columns.Add("CTL_NO_TYPE")
                .Columns.Add("CTL_NO_GET")
                .Columns.Add("CTL_NO_COUNT", GetType(System.Int64))
                .PrimaryKey = New DataColumn() { .Columns("CTL_NO_TYPE")}
            End With

            ASCMAIN1.sql = "Select * from TATCTLN3"
            Create_TDA(.Tables.Add, "TATCTLN3", "**", 0)

            Create_Relation("TATCTLN0", "TATCTLN3", "CTL_NO_TYPE")
            .Tables("TATCTLN0").Columns("CTL_NO_COUNT").Expression = "COUNT(CHILD.CTL_NO)"

            Dim ORDRWHERE As String = " WHERE ORDR_NO IN (Select ORDR_NO from SOTORDR1_L WHERE ORDR_STATUS = 'L')"

            ASCMAIN1.sql = "Select SOTORDR1_L.* from SOTORDR1_L" & ORDRWHERE
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False)
            .Tables("SOTORDR1").Columns.Add("ORDR_AMT", GetType(System.Decimal))
            .Tables("SOTORDR1").Columns.Add("SEL")
            .Tables("SOTORDR1").Columns("SEL").DefaultValue = "0"
            Create_TDA(.Tables.Add, "SOTORDRR", "**", 0, False)
            .Tables("SOTORDRR").Columns.Add("ORDR_AMT", GetType(System.Decimal))
            .Tables("SOTORDRR").Columns.Add("SEL")
            .Tables("SOTORDRR").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select SOTORDR2_L.* from SOTORDR2_L" & ORDRWHERE
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False)
            .Tables("SOTORDR2").Columns.Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")

            Create_Relation("SOTORDR1", "SOTORDR2", "ORDR_NO")
            .Tables("SOTORDR1").Columns("ORDR_AMT").Expression = "SUM(CHILD.ORDR_AMT)"


            ASCMAIN1.sql = "Select SOTORDR5_L.* from SOTORDR5_L" & ORDRWHERE
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, False)

            ASCMAIN1.sql = "Select ARTCUST1_L.* from ARTCUST1_L"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False)
            ASCMAIN1.sql = "Select ARTCUST2_L.* from ARTCUST2_L"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False)
            ASCMAIN1.sql = "Select ARTCUSTD_L.* from ARTCUSTD_L"
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, False)
            ASCMAIN1.sql = "Select * from ARTCUSTQ_L"
            Create_TDA(.Tables.Add, "ARTCUSTQ_L", "**", 0, False)

            ASCMAIN1.sql = "Select * from SOTORDR1"
            Create_TDA(.Tables.Add("SOTORDR1_Q"), "SOTORDR1", "**", 0, True)
            ASCMAIN1.sql = "Select * from SOTORDR2"
            Create_TDA(.Tables.Add("SOTORDR2_Q"), "SOTORDR2", "**", 0, True)
            ASCMAIN1.sql = "Select * from SOTORDR5"
            Create_TDA(.Tables.Add("SOTORDR5_Q"), "SOTORDR5", "**", 0, True)
            ASCMAIN1.sql = "Select * from SOTQRDR1 WHERE ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTQRDR1", "**", 0, True, "V", 1)
            ASCMAIN1.sql = "Select * from SOTQRDR2 WHERE ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTQRDR2", "**", 0, True, "V", 2)
            ASCMAIN1.sql = "Select * from SOTQRDR5 WHERE ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTQRDR5", "**", 0, True, "V", 2)
        End With

        grdTATCTLN0.DataSource = dst.Tables("TATCTLN0")
        Create_Summary(grdTATCTLN0, "CTL_NO_TYPE", "Count")

        grdARTCUST1.DataSource = dst.Tables("ARTCUST1")
        Create_Summary(grdARTCUST1, "CUST_CODE", "Count")

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdTATCTLN0, grdARTCUST1, grdSOTORDR1}
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        Next

        grdSOTORDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For b As Integer = 0 To 1
            For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDR1.DisplayLayout.Bands(b).Columns
                If b = 0 And gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        Next

        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        'FtpS.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        'FtpS.RuntimeLicense = "31484E3941413153554252415331544531414D4831343236000000000000000000000000000000003335384A30543346000046365241325A505A504E36300000"
        FtpS.RuntimeLicense = "31484E46414431535542323032333033313352415331544531414D483134323600000000000000003335384A30543346000059554A4336594E46335047530000"
        ASCMAIN1.Add_Value_List(grdTATCTLN0, "CTL_NO_TYPE", Nothing, New String() {":", "SOTORDR1.ORDR_NO:Sales Order", "ARTCUST1.CUST_CODE:Customer"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"
                If Not ASCMAIN1.Logical_Lock("F", "SOFORDRO", , False) Then
                    EMsg &= vbCr & "You Should Close The Order Entry Screen Before Refreshing"
                End If
                If Not ASCMAIN1.Logical_Lock("T", "SOTCUST1", , False) Then
                    EMsg &= vbCr & "You Should Close The Customer Maintenance Screen Before Refreshing"
                End If
            Case "Update Masterfile", "Full Transmission", "Transmit Orders", "Print transfer sheet", "E-mail transfer sheet"
                'If dst.Tables("ARTCUST1").Select("").Length = 0 And dst.Tables("SOTORDR1").Select("SEL='1'").Length = 0 Then
                '    EMsg &= vbCr & "Nothing selected to be transmitted"
                'End If
                If RefreshRequired Then
                    EMsg &= vbCr & "You Should Refresh Your Data After Closing Other Forms"
                End If
                If Not ASCMAIN1.Logical_Lock("F", "SOFORDRO", , False) Then
                    EMsg &= vbCr & "You May Not Transfer/Update While Order Entry Screen Is Open"
                    RefreshRequired = True
                End If
                If Not ASCMAIN1.Logical_Lock("T", "SOTCUST1", , False) Then
                    EMsg &= vbCr & "You May Not Transfer/Update While Customer Maintenance Screen Is Open"
                    RefreshRequired = True
                End If
                For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select()
                    Dim CUST_CODE_SEL As String = rowARTCUST1.Item("CUST_CODE").ToString & ""
                    If dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_SEL & "'").Count = 0 Then
                        EMsg &= vbCr & "Customer Code " & CUST_CODE_SEL & " Has No Valid Ship Tos"
                    End If
                Next
                Dim ORDR_NOs As New List(Of String)
                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("SEL='1'")
                    ORDR_NOs.Add(rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
                Next
                Dim DUP_STYLES As Boolean = CHECK_STYLE_COLOR_DUPS(ORDR_NOs, "Transmitting The Order")
                If DUP_STYLES Then
                    EMsg &= vbCr & "Fix Duplicate Style/Colors."
                End If
            Case "Clear Orders Pending"
                Dim InPassWord As String = ASCMAIN1.Get_txt_from_User _
               ("Enter Password To Clear ALL Pending Orders!", "Hope You Know What Your Doing.", True, 12, "")
                If InPassWord.ToUpper <> "DELPENDING" Then
                    EMsg &= vbCr & "Invalid Password Entered"
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

            Case "Refresh"
                Refresh_Data()
                RefreshRequired = False
            Case "Full Transmission"
                Me.Cursor = Cursors.WaitCursor
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Update Masterfile"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.Length = 0
                iMSG.AppendLine("Would You Also Like to Backup Your Data?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    SOCMAIN2.BackUpLaptop(Me)
                    MsgBox("Backup Complete", MsgBoxStyle.OkOnly, iTitle)
                    BackupComplete = True
                End If

                Transmit_Orders()

                iTitle = "Update Masterfile"
                iMSG.Length = 0
                iMSG.AppendLine("Would You Also Like to Update The MasterFiles?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    UpdateMasterfiles()
                    MsgBox("Update Complete", MsgBoxStyle.OkOnly, iTitle)
                End If
                Refresh_Data()
                RefreshRequired = False
                Me.Cursor = Cursors.Default
            Case "Backup Orders"
                Me.Cursor = Cursors.WaitCursor
                SOCMAIN2.BackUpLaptop(Me)
                MsgBox("Backup Complete", MsgBoxStyle.OkOnly, "Backup")
                BackupComplete = True
                Refresh_Data()
                Me.Cursor = Cursors.Default
            Case "Transmit Orders"
                Me.Cursor = Cursors.WaitCursor
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Transmit Orders"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.Length = 0
                iMSG.AppendLine("Would You Also Like to Backup Your Data?")
                If BackupComplete Then
                    iResult = MsgBoxResult.No
                Else
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                End If
                If iResult = MsgBoxResult.Yes Then
                    SOCMAIN2.BackUpLaptop(Me)
                    MsgBox("Backup Complete", MsgBoxStyle.OkOnly, iTitle)
                    BackupComplete = True
                End If
                Transmit_Orders()
                Refresh_Data()
                Me.Cursor = Cursors.Default
                MsgBox("Transmit Complete", MsgBoxStyle.OkOnly, "Transmit")
            Case "Update Masterfile"
                Dim BackupTime As String = ""
                Dim FetchTime As String = ""
                Dim RefreshTime As String = ""
                Dim STime As DateTime
                Dim ETime As DateTime
                Me.Cursor = Cursors.WaitCursor
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Update Masterfile"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.Length = 0
                iMSG.AppendLine("Would You Also Like to Backup Your Data?")
                If BackupComplete Then
                    iResult = MsgBoxResult.Yes
                Else
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                End If
                If iResult = MsgBoxResult.Yes Then
                    STime = Now()
                    SOCMAIN2.BackUpLaptop(Me)
                    ETime = Now()
                    BackupTime = $"Backup Time: {Format(DateDiff(DateInterval.Second, STime, ETime) / 60, "###.#0")}"
                    MsgBox("Backup Complete", MsgBoxStyle.OkOnly, iTitle)
                    BackupComplete = True
                End If
                UpdateMasterfiles(FetchTime, RefreshTime)
                Refresh_Data()
                Dim msg As New Text.StringBuilder With {.Length = 0}
                msg.AppendLine("Update Complete")
                If BackupTime.Length > 0 Then
                    msg.AppendLine(BackupTime)
                End If
                If FetchTime.Length > 0 Then
                    msg.AppendLine(FetchTime)
                End If
                If RefreshTime.Length > 0 Then
                    msg.AppendLine(RefreshTime)
                End If
                MsgBox(msg.ToString, MsgBoxStyle.OkOnly, "Update Masterfile")
                Me.Cursor = Cursors.Default
            Case "Print transfer sheet"
                Print_Record()
            Case "E-mail transfer sheet"
                Print_Record(True)
            Case "Update Software"
                Dim cpwd As String = TodaysPwd(CDate(Now().ToShortDateString))
                Dim upwd As String = InputBox("Password", "Update Software")
                If cpwd = upwd Then
                    UpdateSoftware()
                Else
                    MsgBox("Incorrect Password.", vbOKOnly, "Thanks For Playing")
                End If
            Case "Clear Historical Orders"
                ClearHistory()
            Case "Clear Control Numbers"
                ClearControls()
            Case "Clear Customers Pending"
                ClearPendingCustomers()
            Case "Clear Orders Pending"
                ClearOrdersPending()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("View").Settings.Enabled = not_iScreenMode
                    '.Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode

                    '.Items("View").Visible = (EntryMode = "L" Or Not ScreenMode)
                    '.Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    '.Items("Print").Visible = ScreenMode
                    'Clear Orders Pending
                    If ASCMAIN1.USER_ID = "mariog" Or ASCMAIN1.USER_ID = "wayne" Then
                        .Items("Clear Orders Pending").Visible = True
                        '.Items("Update Software").Visible = True
                        'chkNEWAPI.Visible = True
                    Else
                        .Items("Clear Orders Pending").Visible = False
                        '.Items("Update Software").Visible = False
                        'chkNEWAPI.Visible = False
                    End If
                End With
                .Groups("Import Quotes").Expanded = False
                .Groups("Images").Expanded = False
                .Groups("Version").Expanded = False
            End With
        End If

        'chkNEWAPI.Checked = True

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdTATCTLN0.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"TATCTLN0", "TATCTLN3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Data()
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record(Optional ByVal viaEmail As Boolean = False)
        PrintSelected = True



        dst.Tables("SOTORDRR").Clear()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("SEL='1'")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            Dim rowSOTORDRR As DataRow = dst.Tables("SOTORDRR").NewRow
            rowSOTORDRR.Item("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO")
            rowSOTORDRR.Item("ORDR_DATE") = rowSOTORDR1.Item("ORDR_DATE")
            rowSOTORDRR.Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
            rowSOTORDRR.Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
            rowSOTORDRR.Item("CUST_NAME") = rowSOTORDR1.Item("CUST_NAME")
            rowSOTORDRR.Item("INIT_OPER") = rowSOTORDR1.Item("INIT_OPER")
            rowSOTORDRR.Item("ORDR_AMT") = rowSOTORDR1.Item("ORDR_AMT")
            dst.Tables("SOTORDRR").Rows.Add(rowSOTORDRR)
            'ASCDATA1.ExecuteSQL(String.Format("Insert Into SOTORDRR SELECT * from SOTORDR1 where ORDR_NO = '{0}'", ORDR_NO))
            'ASCDATA1.ExecuteSQL()
        Next
        Print_Report_Begin()



        Generate_Report("SOFXFER2")
        Print_Report_End()

        If viaEmail Then
            Dim tempFileName As String = "TransferSheet_" & Format(Now(), "yyMMddhhss")
            Dim attachments As String()
            ReDim attachments(0)
            attachments(0) = ASCMAIN1.Folders("Temp") & tempFileName & ".pdf"
            Dim REPORT_NO As String = Generate_Report("SOFXFER2", "Quote Sheet", "", "", "PDF", tempFileName, True)
            Print_Report_End(, True)
            Dim MAIL_BODY As New Text.StringBuilder With {.Length = 0}
            MAIL_BODY.AppendLine("Please Find Attached My Transfer Sheet Referring to A Transfer I Made Today.")
            MAIL_BODY.AppendLine("If You Do Not See The Orders Come Through When You Import The Transfers, Please")
            MAIL_BODY.AppendLine("Contact Me immediately.")
            'email_Quote(tempFileName)
            Create_Outlook_mailitem("Kathy <kathy@regency-rib.com>", "rita@regency-rib.com", "Transfer Sheet", MAIL_BODY.ToString(), attachments)
        Else
            Generate_Report("SOFXFER2")
            Print_Report_End()
        End If


    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "USER_ID"
                sql_where = "USER_ID in (Select Distinct WO_ASSIGNED_TO from SOTWORK1 union Select Distinct WO_ASSIGNED_TO from SOTWORK2)"
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDR1, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Select Selected", "Select All CUSTOMER")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdSOTORDR1"
                tlb_btn = DirectCast(tlb_pop.Tools("Select All CUSTOMER"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse (Not grd.ActiveRow.IsDataRow Or grd.ActiveRow.Band.Key <> "SOTORDR1") Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Visible = True
                    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                    tlb_btn.Tag = CUST_CODE
                    tlb_btn.SharedProps.Caption = "Select All " & CUST_CODE
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SEL").Value = "1"
                    grow.Update()
                Next

            Case "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SEL").Value = "0"
                    grow.Update()
                Next

            Case "Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = "1"
                    grow.Update()
                Next

            Case "Select All CUSTOMER"
                tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Dim CUST_CODE As String = tlb_btn.Tag
                For Each row As DataRow In dst.Tables("SOTORDR1").Select("CUST_CODE = '" & CUST_CODE & "'")
                    row.Item("SEL") = "1"
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "PO Inquiry"
            '    Dim WO_REF_TYPE As String = grd.ActiveRow.Cells("WO_REF_TYPE").Value
            '    Dim WO_REF_NO As String = grd.ActiveRow.Cells("WO_REF_NO").Value
            '    If WO_REF_TYPE = "P" Then Context_Launch("View", WO_REF_NO, e.Tool.Key, "POFORDRI", "F", "POE")

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

#Region "grdTATCTLNO"
    Private Sub grdTATCTLN0_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdTATCTLN0.ClickCellButton

        Dim CTL_NO_TYPE As String = e.Cell.Row.Cells("CTL_NO_TYPE").Value & ""
        Dim CTL_NO_TYPE_DESC As String = e.Cell.Row.Cells("CTL_NO_TYPE").Text & ""
        Dim rowTATCTN0 As DataRow = dst.Tables("TATCTLN0").Rows.Find(CTL_NO_TYPE)

        Dim CTL_NO_COUNT As Int64 = Val(rowTATCTN0.Item("CTL_NO_COUNT") & "")
        Dim CTL_NO_GET As Int64 = Val(rowTATCTN0.Item("CTL_NO_GET") & "")
        Dim CTL_NO_GET_LARGER As Int64 = CTL_NO_GET
        Select Case CTL_NO_TYPE
            Case "SOTORDR1.ORDR_NO"
                CTL_NO_GET_LARGER = 500
            Case "ARTCUST1.CUST_CODE"
                CTL_NO_GET_LARGER = 200
        End Select


        If (ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "mariog") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "You Are Special"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("Because You Are Mario, You Are Special")
            iMSG.AppendLine("And Can Request the Larger Option Of")
            iMSG.AppendLine(CTL_NO_GET_LARGER & " Control Numbers.")
            iMSG.AppendLine("Is That What You Want?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                CTL_NO_GET = CTL_NO_GET_LARGER
            End If
        Else
            If CTL_NO_COUNT > CInt(CTL_NO_GET / 2) Then
                MsgBox("You may not ask for new " & CTL_NO_TYPE_DESC & " Numbers unless your current available count is less than " & CStr(CInt((CTL_NO_GET / 2))), MsgBoxStyle.OkOnly, "Cannot Perform Reqeusted Action")
                Exit Sub
            End If
        End If


        Dim TABLE_NAME As String = Split(CTL_NO_TYPE, ".")(0)
        Dim COLUMN_NAME As String = Split(CTL_NO_TYPE, ".")(1)

        Try
            BeginTrans()
            'Dim CTL_NO_START As String = TAC.SOCMAIN1.Get_CTL_NOs(ASCMAIN1.USER_ID, ASCMAIN1.COMPUTER_NAME, TABLE_NAME, COLUMN_NAME, CTL_NO_GET)

            Dim ORDR_NO_QS As String = "?USER_ID=" & ASCMAIN1.USER_ID _
                                   & "&COMPUTER_NAME=laptop" _
                                   & "&TABLE_NAME=" & TABLE_NAME _
                                   & "&COLUMN_NAME=" & COLUMN_NAME _
                                   & "&HOW_MANY=" & CTL_NO_GET

            If (ASCMAIN1.Running_in_VS) Then Stop
            Dim errMsg As String = ""
            Dim CTL_NO_START As String = Generate_ORDR_NO(ORDR_NO_QS, errMsg)
            If errMsg.Length > 0 Then
                MsgBox(errMsg, vbCritical, "Error Fetching Numbers")
                Rollback()
                Exit Sub
            End If
            For I As Integer = 0 To CTL_NO_GET - 1
                Dim CTL_NO_X As String = Format(CTL_NO_START + I, "".PadLeft(Len(CTL_NO_START), "0"))
                Dim rowTATCTLN3 As DataRow = dst.Tables("TATCTLN3").NewRow
                rowTATCTLN3.Item("CTL_NO_TYPE") = CTL_NO_TYPE
                rowTATCTLN3.Item("CTL_NO") = CTL_NO_X
                rowTATCTLN3.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                rowTATCTLN3.Item("INIT_OPER") = ASCMAIN1.USER_ID
                dst.Tables("TATCTLN3").Rows.Add(rowTATCTLN3)
            Next
            Update_Record_TDA("TATCTLN3")
            CommitTrans()
            MsgBox(CStr(CTL_NO_GET) & " " & CTL_NO_TYPE_DESC & " Numbers have been added", MsgBoxStyle.OkOnly, "Success")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Occurred - please call ABS")
            Rollback()
        End Try

    End Sub
#End Region

#Region "Custom Methods"
    Sub Refresh_Data()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"TATCTLN0", "TATCTLN3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        With dst.Tables("TATCTLN0")
            .Rows.Add(New Object() {"SOTORDR1.ORDR_NO", 100})
            .Rows.Add(New Object() {"ARTCUST1.CUST_CODE", 20})
        End With

        Fill_Records("TATCTLN3")

        Fill_Records("SOTORDR1")
        Fill_Records("SOTORDR2")
        Fill_Records("SOTORDR5")

        Fill_Records("ARTCUST1")
        Fill_Records("ARTCUST2")
        Fill_Records("ARTCUSTD")

        Fill_Records("ARTCUSTQ_L")

        EnforceConstraints(True)

    End Sub

    Sub Transmit_Orders()
        If (ASCMAIN1.Running_in_VS) Then Stop

        Dim NoDetails As Boolean = False
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("SEL='1'")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty
            Dim filter As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            Dim rowCtn As Int64 = dst.Tables.Item("SOTORDR2").Select(filter).Count
            If rowCtn = 0 Then
                MsgBox(String.Format("Order {0} Has No Details!!", ORDR_NO), vbExclamation, "No Transmission Will Occur")
                NoDetails = True
                Exit Sub
            End If
        Next
        If Not PrintSelected Then
            Dim Msg As String = "You Have Not Printed A Transmission Report" & vbCrLf & "Do You Still Want To Transfer?"
            Dim iresult As MsgBoxResult = MsgBox(Msg, vbYesNo, "No Report Yet!")
            If iresult <> vbYes Then
                Exit Sub
            End If
        End If
        Try
            Dim XSD = ASCMAIN1.Next_Control_No("XSD")
            Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.USER_ID & "_" & XSD & ".xsd"

            dst.AcceptChanges()
            dst.WriteXml(FILENAME)

            BeginTrans()

            For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("")
                Dim CUST_CODE As String = rowARTCUST1.Item("CUST_CODE")
                If (ASCMAIN1.Running_in_VS) Then Stop
                Dim UPLOAD_ERR As String = CUST_UPLOAD(CUST_CODE)
                If UPLOAD_ERR.Length = 0 Then
                    For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUSTD"}
                        ASCDATA1.ExecuteSQL(String.Format("Delete from {0}_L where CUST_CODE = '{1}'", TABLE_NAME, CUST_CODE))
                    Next
                Else
                    MsgBox(UPLOAD_ERR, vbCritical, "Error Uploading Customers")
                    Rollback()
                    Exit Sub
                End If
            Next

            'If ((ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then

            Dim ORDRs As New List(Of String)
            ORDRs.Add("9999999999")
            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("SEL='1'")
                ORDRs.Add(rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty)
            Next
            For Each rowARTCUSTQ_L As DataRow In dst.Tables("ARTCUSTQ_L").Select()
                If ORDRs.Contains(rowARTCUSTQ_L.Item("LAST_ORDR_NO").ToString & String.Empty) Then
                    Dim CUST_CODE As String = rowARTCUSTQ_L.Item("CUST_CODE")
                    Dim CUST_ADDR_CODE As String = rowARTCUSTQ_L.Item("CUST_ADDR_CODE")
                    If (ASCMAIN1.Running_in_VS) Then Stop
                    Dim UPLOAD_ERR As String = SHIPTO_UPLOAD(CUST_CODE, CUST_ADDR_CODE)
                    If UPLOAD_ERR.Length = 0 Then
                        SB.Length = 0
                        SB.AppendLine("Delete from ARTCUSTQ_L")
                        SB.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                        SB.AppendLine(String.Format("AND CUST_ADDR_CODE = '{0}'", CUST_ADDR_CODE))
                        ASCDATA1.ExecuteSQL(SB.ToString)
                    Else
                        MsgBox(UPLOAD_ERR, vbCritical, "Error Uploading Ship Tos")
                        Rollback()
                        Exit Sub
                    End If
                End If
            Next
            'End If

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("SEL='1'")
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                Dim ORDR_BATCH_NO As String = rowSOTORDR1.Item("ORDR_BATCH_NO")
                If IsOrderTransferPartialGroup(ORDR_NO, ORDR_BATCH_NO) Then
                    rowSOTORDR1.Item("ORDR_BATCH_NO") = ASCMAIN1.Next_Control_No("ORDR_BATCH_NO")
                    ASCDATA1.ExecuteSQL(String.Format("UPDATE SOTORDR1 SET ORDR_BATCH_NO = '{0}' WHERE ORDR_NO = '{1}'", rowSOTORDR1.Item("ORDR_BATCH_NO").ToString(), ORDR_NO))
                End If

                If (ASCMAIN1.Running_in_VS) Then Stop
                Dim UPLOAD_ERR As String = ORDR_UPLOAD(ORDR_NO)
                If UPLOAD_ERR.Length = 0 Then
                    For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
                        ASCDATA1.ExecuteSQL(String.Format("Delete from {0}_L where ORDR_NO = '{1}'", TABLE_NAME, ORDR_NO))
                    Next
                    For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR2"}
                        ASCDATA1.ExecuteSQL(String.Format("UPDATE {0} SET ORDR_STATUS = 'O' WHERE ORDR_NO = '{1}'", TABLE_NAME, ORDR_NO))
                        If TABLE_NAME = "SOTORDR1" Then
                            ASCDATA1.ExecuteSQL(String.Format("UPDATE {0} SET ORDR_DATE_RECD = TRUNC(SYSDATE), INIT_DATE = SYSDATE WHERE ORDR_NO = '{1}'", TABLE_NAME, ORDR_NO))
                        End If
                    Next
                    For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
                        Dim TABLE_NAME_H As String = String.Format("{0}_H", TABLE_NAME)
                        ASCMAIN1.sql = String.Format("Insert into {0} Select * from {1} where ordr_no = '{2}'", TABLE_NAME_H, TABLE_NAME, ORDR_NO)
                        ASCDATA1.ExecuteSQL()
                    Next
                Else
                    MsgBox(UPLOAD_ERR, vbCritical, "Error Uploading Orders")
                    Rollback()
                    Exit Sub
                End If
            Next

            CommitTrans()

            MsgBox("Data has been successfully Transmitted", MsgBoxStyle.OkOnly, "Success")
            PrintSelected = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Please Call ABS")
            Try
                Rollback()
            Catch ex2 As Exception

            End Try
        End Try
    End Sub

    Sub ftp_File(ByVal RemoteFile As String, ByVal LocalFile As String, ByVal Direction As String, Optional ByVal OnlyNewest As Boolean = False)
        'Direction:
        'U = Up
        'D = Down
        Dim LocalFileOrig As String = ""
        Dim RemoteFileOrig As String = ""

        If Direction <> "D" And Direction <> "U" Then
            Exit Sub
        End If

        Ftp1.User = "ABS"
        Ftp1.Password = "0ff1c3"
        Ftp1.RemoteHost = "ftp.regency-rib.com" '"192.168.110.224" '"50.75.200.254"
        Ftp1.Logon()
        'Ftp1.TransferMode = nsoftware.IPWorks.FTPTransferModes.tmBinary
        Ftp1.ChangeTransferMode(nsoftware.IPWorks.FTPTransferModes.tmBinary)
        Ftp1.LocalFile = LocalFile
        LocalFileOrig = LocalFile
        Ftp1.RemoteFile = RemoteFile
        RemoteFileOrig = RemoteFile
        'Ftp1.Timeout = 0 'Don't Timeout
        Ftp1.Overwrite = True
        If OnlyNewest Then
            FileList.Clear()
            Ftp1.ListDirectoryLong()
            If FileList.Count > 0 Then
                For i As Integer = 0 To FileList.Count - 1
                    Ftp1.LocalFile = FileList.Keys(i)
                    Ftp1.RemoteFile = FileList.Values(i)
                    ASCMAIN1.Progress("Fetching " & Ftp1.RemoteFile)
                    Ftp1.Download()
                Next
            End If
        Else
            ASCMAIN1.Progress("Fetching " & RemoteFile)
            If Direction = "U" Then
                Ftp1.Upload()
            Else
                Ftp1.Download()
            End If
        End If
        Ftp1.Logoff()
        ASCMAIN1.Progress("")
    End Sub

    Function ftpS_File(ByVal RemoteFile As String, ByVal RemotePath As String, ByVal LocalFile As String, ByVal LocalPath As String) As Boolean
        Dim RetVal As Boolean = True
        Try
            AddHandler FtpS.OnSSHServerAuthentication, AddressOf SSHServerAuthentication
            AddHandler FtpS.OnSSHStatus, AddressOf SSHStatus

            If Not RemotePath.EndsWith("\") Then
                RemotePath = RemotePath & "\"
            End If
            If Not LocalPath.EndsWith("\") Then
                LocalPath = LocalPath & "\"
            End If

            FileListS.Clear()
            FtpS.SSHUser = "salesReps"
            FtpS.SSHPassword = "0ff1c3ABS"
            FtpS.SSHHost = "sftp.regency-rib.com"
            FtpS.SSHAuthMode = SCPSSHAuthModes.amPassword ' SftpSSHAuthModes.amPassword ' nsoftware.IPWorksSSH.SftpSSHAuthModes.amPassword
            FtpS.SSHEncryptionAlgorithms = "aes256-ctr"
            FtpS.LocalFile = String.Format("{0}{1}", LocalPath, LocalFile)
            FtpS.RemoteFile = RemoteFile
            'FtpS.RemotePath = RemotePath '"/DB"
            FtpS.ChangeRemotePath(RemotePath)
            FtpS.Overwrite = True
            FtpS.Config("PreserveFileTime=True")

            FtpS.SSHLogon("sftp.regency-rib.com", "22")
            FtpS.ListDirectory()
            If FtpS.DirList.Count > 0 Then
                For Each FL As nsoftware.IPWorksSSH.DirEntry In FtpS.DirList
                    If Not FL.IsDir Then
                        FtpS.LocalFile = String.Format("{0}{1}", LocalPath, FL.FileName)
                        FtpS.RemoteFile = FL.FileName

                        Dim DFile As Boolean = True
                        If IO.File.Exists(FtpS.LocalFile) Then
                            If IO.File.GetCreationTime(FtpS.LocalFile) >= FL.FileTime Then
                                DFile = False
                            End If
                        End If
                        If FtpS.RemoteFile = "nsoftware.IPWorksSSH.dll" Or FtpS.RemoteFile = "nsoftware.IPWorksSSH.System.dll" Then 'Force Security Update Always.
                            DFile = True
                        End If
                        If DFile Then
                            ASCMAIN1.Progress("Secure Fetch: " & FL.FileName)
                            FtpS.Download()
                        End If
                    End If
                Next
            End If
            FtpS.SSHLogoff()

            ASCMAIN1.Progress("")
        Catch ex As Exception
            FtpS.SSHLogoff()
            ASCMAIN1.Progress("")
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Secure FTP Error")
            RetVal = False
        End Try
        Return RetVal
    End Function

    Public Shared Sub SSHServerAuthentication(sender As Object, e As SFTPClientSSHServerAuthenticationEventArgs) ' nsoftware.IPWorksSSH.SftpSSHServerAuthenticationEventArgs)

        e.Accept = True
    End Sub

    Public Shared Sub SSHStatus(sender As Object, e As SFTPClientSSHStatusEventArgs) ' SftpSSHStatusEventArgs)nsoftware.IPWorksSSH.SftpSSHStatusEventArgs)

        ' MsgBox(e.Message, MsgBoxStyle.OkOnly, "SSHStatus Messages")
        'theLog &= e.Message & vbCrLf
    End Sub

    Public Shared Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.UTF8Encoding()
        Return encoding.GetBytes(str)
    End Function

    'Private Sub GetFileSInfo(sender As Object, e As nsoftware.IPWorksSSH.SftpDirListEventArgs) Handles FtpS.OnDirList
    '    If Not e.IsDir Then
    '        Dim localfile As String = Ftp1.LocalFile
    '        Dim RemoteFile As String = Ftp1.RemoteFile
    '        If RemoteFile.Substring(RemoteFile.Length - 1, 1) = "*" Then
    '            RemoteFile = RemoteFile.Substring(0, RemoteFile.Length - 1) + e.FileName
    '        Else
    '            Exit Sub
    '        End If
    '        If localfile.Length > 0 Then
    '            If localfile.Substring(localfile.Length - 1, 1) = "*" Then
    '                localfile = localfile.Substring(0, localfile.Length - 1) + e.FileName
    '                If IO.File.Exists(localfile) Then
    '                    Dim localDT As Date = IO.File.GetLastWriteTime(localfile)
    '                    If e.FileTime > localDT Then
    '                        FileListS.Add(localfile, RemoteFile)
    '                    End If
    '                Else
    '                    FileListS.Add(localfile, RemoteFile)
    '                End If
    '            End If
    '        Else
    '            Exit Sub
    '        End If
    '    End If
    'End Sub
    Private Sub GetFileInfo(sender As Object, e As nsoftware.IPWorks.FtpDirListEventArgs) Handles Ftp1.OnDirList
        If FTPImages Then
            If Not e.IsDir Then
                Dim localfile As String = Ftp1.LocalFile
                Dim RemoteFile As String = Ftp1.RemoteFile
                If RemoteFile.Substring(RemoteFile.Length - 1, 1) = "*" Then
                    RemoteFile = RemoteFile.Substring(0, RemoteFile.Length - 1) + e.FileName
                Else
                    Exit Sub
                End If
                Dim etm As String = e.FileTime
                If etm.Length > 5 Then
                    If etm.Substring(etm.Length - 5, 5).Contains(":") Then
                        If IsDate(etm.Replace(etm.Substring(etm.Length - 5, 5), Now.Year.ToString)) Then
                            If CDate(etm.Replace(etm.Substring(etm.Length - 5, 5), Now.Year.ToString)) > Now() Then
                                etm = etm.Replace(etm.Substring(etm.Length - 5, 5), Now.AddYears(-1).Year.ToString)
                            Else
                                etm = etm.Replace(etm.Substring(etm.Length - 5, 5), Now.Year.ToString)
                            End If
                        End If
                    End If
                End If
                If IsDate(etm) Then
                    ImageListFTP.Add(e.FileName, etm)
                Else
                    ImageListFTP.Add(e.FileName, Now())
                End If


            End If
        Else
            If Not e.IsDir Then
                Dim localfile As String = Ftp1.LocalFile
                Dim RemoteFile As String = Ftp1.RemoteFile
                If RemoteFile.Substring(RemoteFile.Length - 1, 1) = "*" Then
                    RemoteFile = RemoteFile.Substring(0, RemoteFile.Length - 1) + e.FileName
                Else
                    Exit Sub
                End If
                If localfile.Length > 0 Then
                    If localfile.Substring(localfile.Length - 1, 1) = "*" Then
                        localfile = localfile.Substring(0, localfile.Length - 1) + e.FileName
                        If IO.File.Exists(localfile) Then
                            Dim localDT As Date = IO.File.GetLastWriteTime(localfile)
                            If e.FileTime > localDT Then
                                FileList.Add(localfile, RemoteFile)
                            End If
                        Else
                            FileList.Add(localfile, RemoteFile)
                        End If
                    End If
                Else
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Function ORDR_UPLOAD(ByVal ORDR_NO As String) As String
        Dim RetVal As String = ""

        Dim url As New System.Uri(ServerLive)
        Dim API_BASE As String = ServerLive
        If (ASCMAIN1.Running_in_VS) Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Testing?"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine(String.Format("Test With: {0}", ServerTest))
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                url = New System.Uri(ServerTest)
                API_BASE = ServerTest
            Else
                url = New System.Uri(ServerLive)
                API_BASE = ServerLive
            End If
        Else
            url = New System.Uri(ServerLive)
            API_BASE = ServerLive
        End If

        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create($"{url}api/rgi/lt/serverstatus")
        Dim resptest As System.Net.WebResponse
        Try
            resptest = req.GetResponse()
            resptest.Close()
            req = Nothing
        Catch ex As Exception
            If (ASCMAIN1.Running_in_VS) Then Stop
            req = Nothing
            'API_BASE = ServerLive
            RetVal = vbCrLf & "Server Request Not Responding."
        End Try

        If (ASCMAIN1.Running_in_VS) Then Stop 'Check your URL shit and make dam sure you are not looking at live!

        Dim order As New SOTORDR1_L
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            rowSOTORDR1.Item("ORDR_NO_WEB") = ""
            'rowSOTORDR1.Item("CC_TRANS_ID") = ""
            Dim myType As Type = order.GetType
            Dim myPropertyInfo As PropertyInfo() = myType.GetProperties((BindingFlags.Public Or BindingFlags.Instance))
            For i As Integer = 0 To myPropertyInfo.Length - 1
                Dim myPropInfo As PropertyInfo = CType(myPropertyInfo(i), PropertyInfo)
                If myPropInfo.Name <> "OrderDetails" And myPropInfo.Name <> "OrderAddressTypes" Then
                    If Not IsDBNull(rowSOTORDR1.Item(myPropInfo.Name)) Then
                        Select Case myPropInfo.Name
                            Case "ORDR_STATUS"
                                myPropInfo.SetValue(order, "O", Nothing)
                            'Order Date Stays the Original Order Date Now - 9/5/23 W.R.
                            'Date Recd & Init Date Gets Set to Date of Transfer - 9/5/23 W.R.
                            'Case "ORDR_DATE"
                            '    myPropInfo.SetValue(order, Now(), Nothing)
                            Case "ORDR_DATE_RECD"
                                myPropInfo.SetValue(order, Now(), Nothing)
                            Case "INIT_DATE"
                                myPropInfo.SetValue(order, Now(), Nothing)
                            Case "ORDR_PICK_SEQ"
                                'myPropInfo.SetValue(order, CInt(0), Nothing)
                            Case Else
                                myPropInfo.SetValue(order, rowSOTORDR1.Item(myPropInfo.Name), Nothing)
                        End Select
                    End If
                End If
            Next i
        Next

        Dim ORDR_DETAILS As New List(Of SOTORDR2_L)()
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            Dim SOTORDR2_L_NEW As New SOTORDR2_L
            Dim myType As Type = SOTORDR2_L_NEW.GetType
            Dim myPropertyInfo As PropertyInfo() = myType.GetProperties((BindingFlags.Public Or BindingFlags.Instance))
            For i As Integer = 0 To myPropertyInfo.Length - 1
                Dim myPropInfo As PropertyInfo = CType(myPropertyInfo(i), PropertyInfo)
                If Not IsDBNull(rowSOTORDR2.Item(myPropInfo.Name)) Then
                    Select Case myPropInfo.Name
                        Case "ORDR_STATUS"
                            myPropInfo.SetValue(SOTORDR2_L_NEW, "O", Nothing)
                        Case Else
                            myPropInfo.SetValue(SOTORDR2_L_NEW, rowSOTORDR2.Item(myPropInfo.Name), Nothing)
                    End Select
                End If
            Next i
            ORDR_DETAILS.Add(SOTORDR2_L_NEW)
        Next

        Dim ORDR_ADDR_TYPES As New List(Of SOTORDR5_L)()
        For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            Dim SOTORDR5_L_NEW As New SOTORDR5_L
            Dim myType As Type = SOTORDR5_L_NEW.GetType
            Dim myPropertyInfo As PropertyInfo() = myType.GetProperties((BindingFlags.Public Or BindingFlags.Instance))
            For i As Integer = 0 To myPropertyInfo.Length - 1
                Dim myPropInfo As PropertyInfo = CType(myPropertyInfo(i), PropertyInfo)
                If Not IsDBNull(rowSOTORDR5.Item(myPropInfo.Name)) Then
                    myPropInfo.SetValue(SOTORDR5_L_NEW, rowSOTORDR5.Item(myPropInfo.Name), Nothing)

                End If
            Next i
            ORDR_ADDR_TYPES.Add(SOTORDR5_L_NEW)
        Next

        order.OrderDetails = ORDR_DETAILS
        order.OrderAddressTypes = ORDR_ADDR_TYPES

        Dim client As New HttpClient()
        client.BaseAddress = New Uri(API_BASE)

        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        'client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))
        client.Timeout = New TimeSpan(0, 5, 0)

        Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()

        'Dim content As HttpContent = New ObjectContent(Of SOTORDR1_L)(order, frmtr)
        Dim content As String = JsonConvert.SerializeObject(order)
        Dim Buffer = System.Text.Encoding.UTF8.GetBytes(content)
        Dim byteContent = New ByteArrayContent(Buffer)
        byteContent.Headers.ContentType = New MediaTypeHeaderValue("application/json")

        'Dim resp As HttpResponseMessage = client.PostAsync(API_CONTROLLER_ORDERS, content).Result

        If (ASCMAIN1.Running_in_VS) Then Stop
        Dim resp As HttpResponseMessage = client.PostAsync(API_CONTROLLER_ORDERS, byteContent).Result()

        Dim apiResponseString As String = Newtonsoft.Json.JsonConvert.SerializeObject(resp)

        'Dim fullErrorMsg As String = ""
        If Not resp.IsSuccessStatusCode Then
            RetVal = vbCrLf & RetVal & String.Format("{0} ({1})", CInt(resp.StatusCode), resp.ReasonPhrase)
            RetVal = vbCrLf & RetVal & $"Tried Post: {API_CONTROLLER_ORDERS}"
            RetVal = vbCrLf & RetVal & $"Base: {API_BASE}"
            '
        End If

        Return RetVal
    End Function

    Private Function CUST_UPLOAD(ByVal CUST_CODE As String) As String
        Dim RetVal As String = ""

        Dim url As New System.Uri(ServerLive)
        Dim API_BASE As String = ServerLive
        If (ASCMAIN1.Running_in_VS) Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Testing?"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine(String.Format("Test With: {0}", ServerTest))
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                url = New System.Uri(ServerTest)
                API_BASE = ServerTest
            Else
                url = New System.Uri(ServerLive)
                API_BASE = ServerLive
            End If
        Else
            url = New System.Uri(ServerLive)
            API_BASE = ServerLive
        End If

        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create($"{url}api/rgi/lt/serverstatus")
        Dim resptest As System.Net.WebResponse
        Try
            resptest = req.GetResponse()
            resptest.Close()
            req = Nothing
        Catch ex As Exception
            If (ASCMAIN1.Running_in_VS) Then Stop
            req = Nothing
            'API_BASE = ServerLive
            RetVal = vbCrLf & "Server Request Not Responding."
        End Try

        Dim newARTCUST1_L As New ARTCUST1_L
        For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select(String.Format("CUST_CODE = '{0}'", CUST_CODE))
            Dim myType As Type = newARTCUST1_L.GetType
            Dim myPropertyInfo As PropertyInfo() = myType.GetProperties((BindingFlags.Public Or BindingFlags.Instance))
            For i As Integer = 0 To myPropertyInfo.Length - 1
                Dim myPropInfo As PropertyInfo = CType(myPropertyInfo(i), PropertyInfo)
                If myPropInfo.Name <> "CustAddressTypes" And myPropInfo.Name <> "CustContacts" Then
                    If Not IsDBNull(rowARTCUST1.Item(myPropInfo.Name)) Then
                        Select Case myPropInfo.Name
                            Case "XXXXXX"
                                myPropInfo.SetValue(newARTCUST1_L, "X", Nothing)
                            Case Else
                                'MsgBox(myPropInfo.Name)
                                myPropInfo.SetValue(newARTCUST1_L, rowARTCUST1.Item(myPropInfo.Name), Nothing)
                        End Select
                    End If
                End If
            Next i
        Next

        Dim newARTCUST2_L As New List(Of ARTCUST2_L)
        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select(String.Format("CUST_CODE = '{0}'", CUST_CODE))
            Dim ARTCUST2_L_NEW As New ARTCUST2_L
            Dim myType As Type = ARTCUST2_L_NEW.GetType
            Dim myPropertyInfo As PropertyInfo() = myType.GetProperties((BindingFlags.Public Or BindingFlags.Instance))
            For i As Integer = 0 To myPropertyInfo.Length - 1
                Dim myPropInfo As PropertyInfo = CType(myPropertyInfo(i), PropertyInfo)
                If Not IsDBNull(rowARTCUST2.Item(myPropInfo.Name)) Then
                    Select Case myPropInfo.Name
                        Case "XXXXXX"
                            myPropInfo.SetValue(ARTCUST2_L_NEW, "X", Nothing)
                        Case Else
                            'MsgBox(myPropInfo.Name)
                            myPropInfo.SetValue(ARTCUST2_L_NEW, rowARTCUST2.Item(myPropInfo.Name), Nothing)
                    End Select
                End If
            Next i
            newARTCUST2_L.Add(ARTCUST2_L_NEW)
        Next

        Dim newARTCUSTD_L As New List(Of ARTCUSTD_L)
        For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(String.Format("CUST_CODE = '{0}'", CUST_CODE))
            Dim ARTCUSTD_L_NEW As New ARTCUSTD_L
            Dim myType As Type = ARTCUSTD_L_NEW.GetType
            Dim myPropertyInfo As PropertyInfo() = myType.GetProperties((BindingFlags.Public Or BindingFlags.Instance))
            For i As Integer = 0 To myPropertyInfo.Length - 1
                Dim myPropInfo As PropertyInfo = CType(myPropertyInfo(i), PropertyInfo)
                If Not IsDBNull(rowARTCUSTD.Item(myPropInfo.Name)) Then
                    Select Case myPropInfo.Name
                        Case "XXXXXX"
                            myPropInfo.SetValue(ARTCUSTD_L_NEW, "X", Nothing)
                        Case Else
                            'MsgBox(myPropInfo.Name)
                            myPropInfo.SetValue(ARTCUSTD_L_NEW, rowARTCUSTD.Item(myPropInfo.Name), Nothing)
                    End Select
                End If
            Next i
            newARTCUSTD_L.Add(ARTCUSTD_L_NEW)
        Next

        newARTCUST1_L.CustAddressTypes = newARTCUST2_L
        newARTCUST1_L.CustContacts = newARTCUSTD_L

        Dim client As New HttpClient()
        client.BaseAddress = New Uri(API_BASE)

        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        'client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))

        Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()

        'Dim content As HttpContent = New ObjectContent(Of ARTCUST1_L)(newARTCUST1_L, frmtr)
        Dim content As String = JsonConvert.SerializeObject(newARTCUST1_L)
        Dim Buffer = System.Text.Encoding.UTF8.GetBytes(content)
        Dim byteContent = New ByteArrayContent(Buffer)
        byteContent.Headers.ContentType = New MediaTypeHeaderValue("application/json")

        If (ASCMAIN1.Running_in_VS) Then Stop
        'Dim resp As HttpResponseMessage = client.PostAsync(API_CONTROLLER_CUSTOMERS, content).Result
        Dim resp As HttpResponseMessage = client.PostAsync(API_CONTROLLER_CUSTOMERS, byteContent).Result()

        If (ASCMAIN1.Running_in_VS) Then Stop
        Dim apiResponseString As String = Newtonsoft.Json.JsonConvert.SerializeObject(resp)

        'Dim fullErrorMsg As String = ""
        If Not resp.IsSuccessStatusCode Then
            RetVal = vbCrLf & String.Format("{0} ({1})", CInt(resp.StatusCode), resp.ReasonPhrase)
            'fullErrorMsg &= vbCrLf & vbCrLf & String.Format("{0} ({1})", CInt(resp.StatusCode), resp.ReasonPhrase)
        End If

        'RetVal = vbCrLf & fullErrorMsg
        Return RetVal
    End Function

    Private Function SHIPTO_UPLOAD(ByVal CUST_CODE As String, ByVal CUST_ADDR_CODE As String) As String
        Dim RetVal As String = ""
        'Dim fullErrorMsg As String = ""
        If (ASCMAIN1.Running_in_VS) Then Stop

        Dim url As New System.Uri(ServerLive)
        Dim API_BASE As String = ServerLive
        If (ASCMAIN1.Running_in_VS) Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Testing?"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine(String.Format("Test With: {0}", ServerTest))
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                url = New System.Uri(ServerTest)
                API_BASE = ServerTest
            Else
                url = New System.Uri(ServerLive)
                API_BASE = ServerLive
            End If
        Else
            url = New System.Uri(ServerLive)
            API_BASE = ServerLive
        End If

        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create($"{url}api/rgi/lt/serverstatus")
        Dim resptest As System.Net.WebResponse
        Try
            resptest = req.GetResponse()
            resptest.Close()
            req = Nothing
        Catch ex As Exception
            If (ASCMAIN1.Running_in_VS) Then Stop
            req = Nothing
            'API_BASE = ServerLive
            RetVal = vbCrLf & "Server Request Not Responding."
        End Try

        Dim newARTCUSTQ_L As New ARTCUSTQ_L
        Dim rowARTCUSTQ_L As DataRow = LookUp("ARTCUSTQ_L", New String() {CUST_CODE, CUST_ADDR_CODE})
        If Not IsNothing(rowARTCUSTQ_L) Then
            newARTCUSTQ_L.CUST_CODE = rowARTCUSTQ_L.Item("CUST_CODE").ToString & String.Empty
            newARTCUSTQ_L.CUST_ADDR_CODE = rowARTCUSTQ_L.Item("CUST_ADDR_CODE").ToString & String.Empty
            newARTCUSTQ_L.LAST_DATE = rowARTCUSTQ_L.Item("LAST_DATE")
            newARTCUSTQ_L.LAST_OPER = rowARTCUSTQ_L.Item("LAST_OPER").ToString & String.Empty
            newARTCUSTQ_L.LAST_ORDR_NO = rowARTCUSTQ_L.Item("LAST_ORDR_NO").ToString & String.Empty
            newARTCUSTQ_L.RESIDENTIAL_ORDR = rowARTCUSTQ_L.Item("RESIDENTIAL_ORDR").ToString & String.Empty
            newARTCUSTQ_L.GATE_LIFT_REQ = rowARTCUSTQ_L.Item("GATE_LIFT_REQ").ToString & String.Empty
            newARTCUSTQ_L.LIMITED_ACCESS = rowARTCUSTQ_L.Item("LIMITED_ACCESS").ToString & String.Empty
            newARTCUSTQ_L.LIMITED_ACCESS_NOTE = rowARTCUSTQ_L.Item("LIMITED_ACCESS_NOTE").ToString & String.Empty
            newARTCUSTQ_L.IRREGULAR_HOURS = rowARTCUSTQ_L.Item("IRREGULAR_HOURS").ToString & String.Empty
            newARTCUSTQ_L.IRREGULAR_HOURS_NOTE = rowARTCUSTQ_L.Item("IRREGULAR_HOURS_NOTE").ToString & String.Empty
            newARTCUSTQ_L.APPOINTMENT_REQUIRED = rowARTCUSTQ_L.Item("APPOINTMENT_REQUIRED").ToString & String.Empty
            newARTCUSTQ_L.APPOINTMENT_REQUIRED_NOTE = rowARTCUSTQ_L.Item("APPOINTMENT_REQUIRED_NOTE").ToString & String.Empty
            newARTCUSTQ_L.BROKER = rowARTCUSTQ_L.Item("BROKER").ToString & String.Empty
            newARTCUSTQ_L.BROKER_NOTE = rowARTCUSTQ_L.Item("BROKER_NOTE").ToString & String.Empty

            Dim client As New HttpClient()
            client.BaseAddress = New Uri(API_BASE)

            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            'client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))

            Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()

            'Dim content As HttpContent = New ObjectContent(Of ARTCUSTQ_L)(newARTCUSTQ_L, frmtr)
            Dim content As String = JsonConvert.SerializeObject(newARTCUSTQ_L)

            Dim Buffer = System.Text.Encoding.UTF8.GetBytes(content)
            Dim byteContent = New ByteArrayContent(Buffer)
            byteContent.Headers.ContentType = New MediaTypeHeaderValue("application/json")

            Dim resp As HttpResponseMessage = client.PostAsync(API_CONTROLLER_SHIPTO, byteContent).Result()

            'Dim resp As HttpResponseMessage = client.PostAsync(API_CONTROLLER_SHIPTO, content).Result

            Dim apiResponseString As String = Newtonsoft.Json.JsonConvert.SerializeObject(resp)
            'Dim content As HttpContent = New ObjectContent(Of uploadStylesRequest)(US_REQ, frmtr)

            If Not resp.IsSuccessStatusCode Then
                RetVal = vbCrLf & String.Format("{0} ({1})", CInt(resp.StatusCode), resp.ReasonPhrase)
            End If
        Else
            RetVal = vbCrLf & "No ShipTo Record Found!"
        End If

        'RetVal = fullErrorMsg
        Return RetVal
    End Function

    Function Generate_ORDR_NO(qs As String, ByRef errMsg As String) As String
        Dim ctl_no As String = ""

        Dim url As New System.Uri(ServerLive)
        Dim API_BASE As String = ServerLive
        If (ASCMAIN1.Running_in_VS) Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Testing?"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine(String.Format("Test With: {0}", ServerTest))
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                url = New System.Uri(ServerTest)
                API_BASE = ServerTest
            Else
                url = New System.Uri(ServerLive)
                API_BASE = ServerLive
            End If
        Else
            url = New System.Uri(ServerLive)
            API_BASE = ServerLive
        End If

        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create($"{url}api/rgi/lt/serverstatus")
        Dim resptest As System.Net.WebResponse
        Try
            resptest = req.GetResponse()
            resptest.Close()
            req = Nothing
        Catch ex As Exception
            If (ASCMAIN1.Running_in_VS) Then Stop
            req = Nothing
            'API_BASE = ServerLive
            errMsg = vbCrLf & "Server Request Not Responding."
        End Try

        If errMsg.Length = 0 Then
            Dim API_QUERY_STRING As String = qs

            Dim client As New HttpClient()
            client.BaseAddress = New Uri(API_BASE)

            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            If (ASCMAIN1.Running_in_VS) Then Stop
            Dim response As HttpResponseMessage = client.GetAsync(API_CONTROLLER_TATCTLN1 & API_QUERY_STRING).Result

            If (ASCMAIN1.Running_in_VS) Then Stop
            If response.IsSuccessStatusCode Then
                Try
                    Dim TATCTLN1_RESPONSE As New TATCTLN1
                    Dim apiResponseString As String = ""
                    Dim responseObject As Object = response.Content.ReadAsAsync(Of IEnumerable(Of TATCTLN1))().Result

                    TATCTLN1_RESPONSE = responseObject(0)
                    ctl_no = TATCTLN1_RESPONSE.CTL_NO_LAST.ToString.PadLeft(Val(TATCTLN1_RESPONSE.CTL_NO_LENGTH & ""), "0")
                Catch ex As Exception
                    errMsg = vbCrLf & ex.Message
                    'Dim fullErrorMsg As String = "Could not Generate Ordr No fetch Next Order Numbers"
                    'fullErrorMsg &= vbCrLf & vbCrLf & ex.Message
                End Try

            Else
                errMsg = vbCrLf & String.Format("{0} ({1})", CInt(response.StatusCode), response.ReasonPhrase)
                'Dim fullErrorMsg As String = ""
                'fullErrorMsg &= vbCrLf & vbCrLf & String.Format("{0} ({1})", CInt(response.StatusCode), response.ReasonPhrase)
            End If
        End If

        Return ctl_no

    End Function
#End Region

    Private Function IsOrderTransferPartialGroup(ByVal ORDR_NO As String, ByVal ORDR_BATCH_NO As String) As Boolean
        Dim Retval As Boolean = False
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(String.Format("ORDR_BATCH_NO = '{0}'", ORDR_BATCH_NO))
            If rowSOTORDR1.Item("ORDR_NO") <> ORDR_NO Then
                If rowSOTORDR1.Item("SEL") <> "1" Then
                    Retval = True
                End If
            End If
        Next
        Return Retval
    End Function

    Private Sub UpdateMasterfiles(Optional ByRef FetchTime As String = "", Optional ByRef RefreshTime As String = "")
        Dim STime As DateTime = Now()
        Dim ETime As DateTime
        Dim FTPFILES As String() = New String() {"DBUPDATES_TST.BAT", "DBUPDATES.BAT", "RGO_DB.ZIP", "RGO_DB.SQL", "ARTCUST.SQL", "SOTORDR.SQL", "RGO_DB2.SQL", "RGO_DB3.SQL"}
        For Each FTPFILE As String In FTPFILES
            ftpS_File(FTPFILE, "/DB/", FTPFILE, "C:\Shared\RGO\")
        Next
        'If chkNEWAPI.Checked Then
        '    For Each FTPFILE As String In FTPFILES
        '        ftpS_File(FTPFILE, "/DB/", FTPFILE, "C:\Shared\RGO\")
        '    Next
        'Else
        '    For Each FTPFILE As String In FTPFILES
        '        ftp_File(String.Format("\DB\{0}", FTPFILE), String.Format("C:\Shared\RGO\{0}", FTPFILE), "D")
        '    Next
        'End If

        Dim Zip1 As New nsoftware.IPWorksZip.Zip
        Dim p As New System.Diagnostics.ProcessStartInfo()
        Zip1.RuntimeLicense = nSoftwareKeys("nSoftwareZipkey")

        Zip1.ArchiveFile = "C:\Shared\RGO\RGO_DB.ZIP"
        Zip1.ExtractToPath = "C:\Shared\RGO"
        Zip1.OverwriteFiles = True
        Zip1.Extract("RGO_DB.DMP")
        Zip1.Dispose()
        SaveFavorites()

        ETime = Now()
        FetchTime = $"Fetch Time: {Format(DateDiff(DateInterval.Second, STime, ETime) / 60, "###.#0")}"

        STime = Now()
        With p
            .WindowStyle = ProcessWindowStyle.Minimized
            .WorkingDirectory = "C:\Shared\RGO\"
            .FileName = "C:\Shared\RGO\DBUPDATES.bat"
            .UseShellExecute = True
        End With
        ASCMAIN1.Progress("Now Refreshing Database")
        Dim Proc As System.Diagnostics.Process = System.Diagnostics.Process.Start(p)
        Do While Not Proc.HasExited
        Loop
        RestoreFavorites()
        ETime = Now()
        RefreshTime = $"Refresh Time: {Format(DateDiff(DateInterval.Second, STime, ETime) / 60, "###.#0")}"
        ASCMAIN1.Progress("")
    End Sub

    Private Sub ClearHistory()
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Clear History"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("This will Clear All Orders Not Transferred To The Home Office.")
        iMSG.AppendLine("The Next Transfer You Do Will Restore Them.")
        iMSG.AppendLine("Are You Ready?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine("DELETE FROM SOTORDR1")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
            SQLS.Length = 0
            SQLS.AppendLine("DELETE FROM SOTORDR2")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
            SQLS.Length = 0
            SQLS.AppendLine("DELETE FROM SOTORDR5")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
            SQLS.Length = 0
            SQLS.AppendLine("INSERT INTO SOTORDR1 SELECT * FROM SOTORDR1_L")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
            SQLS.Length = 0
            SQLS.AppendLine("INSERT INTO SOTORDR2 SELECT * FROM SOTORDR2_L")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
            SQLS.Length = 0
            SQLS.AppendLine("INSERT INTO SOTORDR5 SELECT * FROM SOTORDR5_L")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
            iMSG.Length = 0
            iMSG.AppendLine("Orders Cleared.")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        End If
    End Sub

    Private Sub UpdateSoftware()
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Software Update"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("This Will Update Your Software. When It Is Finished")
        iMSG.AppendLine("You Will Need To Restart ABSolution.")
        iMSG.AppendLine("Are You Ready?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = vbYes Then
            Dim RemoteRoot As String = "updates\"
            'If ASCMAIN1.USER_ID = "mariog" Or ASCMAIN1.USER_ID = "danny" Or ASCMAIN1.USER_ID = "wayne" Then
            '    iMSG.Length = 0
            '    iMSG.AppendLine("Are You Testing Software?")
            '    Dim iResult2 As MsgBoxResult
            '    iResult2 = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, "Testing Or Live")
            '    If iResult2 = MsgBoxResult.Yes Then
            '        RemoteRoot = "testing\"
            '    End If
            'End If
            'If iResult = MsgBoxResult.Yes Then
            Dim LocalRoot As String = "C:\Shared\RGO\"
            Dim FolderList As New Dictionary(Of String, String)
            FolderList.Add("bin", "bin")
            'FolderList.Add("Reports", "Reports")
            'FolderList.Add("Images\16", "Images\16")
            'FolderList.Add("Images\32", "Images\32")
            For i As Integer = 0 To FolderList.Count - 1
                ftpS_File("*", String.Format("\{0}{1}\", RemoteRoot, FolderList.Keys(i)), "*", String.Format("{0}{1}\", LocalRoot, FolderList.Keys(i)))
                'If chkNEWAPI.Checked Then
                '    ftpS_File("*", String.Format("\{0}{1}\", RemoteRoot, FolderList.Keys(i)), "*", String.Format("{0}{1}\", LocalRoot, FolderList.Keys(i)))
                'Else
                '    ftp_File(String.Format("{0}{1}\*", RemoteRoot, FolderList.Keys(i)), String.Format("{0}{1}\*", LocalRoot, FolderList.Values(i)), "D", True)
                'End If
            Next
            iMSG.Length = 0
            iMSG.AppendLine("Software Update Is Complete.")
            iMSG.AppendLine("You Need To Get Out Of ABSolution")
            iMSG.AppendLine("And Get Back In To Finish The Update.")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        End If

        'End If
        ASCMAIN1.Progress("")
    End Sub

    Private Sub ClearControls()
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Clear Control Numbers"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("This will Clear All Order And Customer Numbers Received.")
        iMSG.AppendLine("Once Done, You Can Request New Ones By Clicking On")
        iMSG.AppendLine("The Get Button In The Control Grid.")
        iMSG.AppendLine("Are You Ready?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine("DELETE FROM TATCTLN3")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
            iMSG.Length = 0
            Refresh_Data()
            iMSG.AppendLine("Numbers Cleared.")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        End If
    End Sub

    Private Sub ClearPendingCustomers()
        Dim iResult As MsgBoxResult
        Dim iTitle As String = ""
        Dim iMSG As New System.Text.StringBuilder
        Dim SQLS As New System.Text.StringBuilder

        SQLS.Length = 0
        SQLS.AppendLine("Select Count(*) as RECCNT from SOTORDR1_L")
        ASCMAIN1.sql = SQLS.ToString()
        Dim RECCNT As Int16 = Val(ASCDATA1.GetDataValue)
        If RECCNT > 0 Then
            iTitle = "Clear Pending Customers"
            iMSG.AppendLine("You May Not Clear Pending Customers While")
            iMSG.AppendLine("There Are Pending Orders To Be Transferred.")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        Else
            iTitle = "Clear Pending Customers"
            iMSG.AppendLine("This will Clear All Pending New and Edited Customers.")
            iMSG.AppendLine("Are You Ready?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                SQLS.Length = 0
                SQLS.AppendLine("DELETE FROM ARTCUST1_L")
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
                SQLS.Length = 0
                SQLS.AppendLine("DELETE FROM ARTCUST2_L")
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
                SQLS.Length = 0
                SQLS.AppendLine("DELETE FROM ARTCUSTD_L")
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
            End If
            Refresh_Data()
        End If
    End Sub

    Private Sub ClearOrdersPending()
        Dim SelFound As Boolean = False
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            If rowSOTORDR1.Item("SEL") = "1" Then
                SelFound = True
            End If
        Next
        If Not SelFound Then
            MsgBox("You Must Check Off A Pending Order To Clear!", MsgBoxStyle.Critical, "Please Make a Selection")
            Exit Sub
        End If
        BeginTrans()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            If rowSOTORDR1.Item("SEL") = "1" Then
                Dim ORDR_NO_DEL As String = rowSOTORDR1.Item("ORDR_NO").ToString & ""
                Dim SQLS As New System.Text.StringBuilder

                Dim T_NAME1 As String() = New String() {"SOTORDR1_L", "SOTORDR2_L", "SOTORDR5_L"}
                For Each TABLE_NAME As String In T_NAME1
                    SQLS.Length = 0
                    SQLS.AppendLine(String.Format("INSERT INTO {0}_H SELECT * FROM {1} WHERE ORDR_NO = '{2}'", TABLE_NAME.Substring(0, 8), TABLE_NAME, ORDR_NO_DEL))
                    ASCMAIN1.sql = SQLS.ToString
                    ASCDATA1.ExecuteSQL()
                Next

                Dim T_NAME2 As String() = New String() {"SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTORDR1_L", "SOTORDR2_L", "SOTORDR5_L"}
                For Each TABLE_NAME As String In T_NAME2
                    SQLS.Length = 0
                    SQLS.AppendLine(String.Format("DELETE FROM {0} WHERE ORDR_NO = '{1}'", TABLE_NAME, ORDR_NO_DEL))
                    ASCMAIN1.sql = SQLS.ToString
                    ASCDATA1.ExecuteSQL()
                Next
            End If
        Next
        CommitTrans("Selected Orders Cleared")
        Refresh_Data()
    End Sub

    Private Sub FillImageStats()
        'FTPImages = True
        FTPImages = True
        ImageListDownload.Clear()
        ImageListDelete.Clear()
        ImageListLocal.Clear()
        ImageListFTP.Clear()
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim IMAGES_FOLDER As String = "C:\"
        Dim ImageFilter As String = "*.jpg"
        Dim RemoteFolder As String = "/www/media/product/"
        If Not IsNothing(rowSOTPARM3) Then
            If rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString.EndsWith("\") Then
                IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            Else
                IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString & "\"
            End If
        End If
        If System.IO.Directory.Exists(IMAGES_FOLDER) Then

            For Each file As String In System.IO.Directory.GetFiles(IMAGES_FOLDER, ImageFilter)
                If file.Length >= 3 Then
                    If file.EndsWith(".JPG") Or file.EndsWith(".jpg") Then
                        Dim LastWriteTime As Date = CDate(System.IO.File.GetLastWriteTime(file))
                        Dim FileName As String = System.IO.Path.GetFileName(file)
                        ImageListLocal.Add(FileName, LastWriteTime)
                    End If
                End If
            Next

            Ftp1.User = "regency-rib"
            Ftp1.Password = "joydHUJ3"
            Ftp1.RemoteHost = "regency-rib.com"
            Ftp1.Logon()
            'Ftp1.TransferMode = nsoftware.IPWorks.FTPTransferModes.tmBinary
            Ftp1.ChangeTransferMode(nsoftware.IPWorks.FTPTransferModes.tmBinary)
            Ftp1.RemoteFile = RemoteFolder & "*"
            Ftp1.LocalFile = IMAGES_FOLDER & "*"
            Ftp1.Overwrite = True
            FileList.Clear()
            Ftp1.ListDirectoryLong()
            Ftp1.Logoff()

            ImageListDownload.Clear()
            ImageListDelete.Clear()

            Dim ImageListFTPtmp As New Dictionary(Of String, Date)
            For Each image As KeyValuePair(Of String, Date) In ImageListFTP
                If Not (ImageListFTPtmp.ContainsKey(image.Key) Or ImageListFTPtmp.ContainsKey(image.Key.Replace(".jpg", ".JPG")) Or ImageListFTPtmp.ContainsKey(image.Key.Replace(".JPG", ".jpg"))) Then
                    ImageListFTPtmp.Add(image.Key.Replace(RemoteFolder, ""), image.Value)
                End If
            Next
            ImageListFTP = ImageListFTPtmp

            Dim ImageListLocaltmp As New Dictionary(Of String, Date)
            For Each image As KeyValuePair(Of String, Date) In ImageListLocal
                ImageListLocaltmp.Add(image.Key.Replace(RemoteFolder, ""), image.Value)
            Next
            ImageListLocal = ImageListLocaltmp

            Dim sql As New Text.StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT")
            sql.AppendLine("(STYLE_CODE || '-' || COLOR_CODE || '.JPG') AS SC")
            sql.AppendLine("FROM")
            sql.AppendLine("(")
            sql.AppendLine("SELECT")
            sql.AppendLine("S1.STYLE_CODE,")
            sql.AppendLine("C1.COLOR_CODE,")
            sql.AppendLine("S1.STYLE_STATUS,")
            sql.AppendLine("C1.STYLE_COLOR_STATUS,")
            sql.AppendLine("S1.STYLE_DESC,")
            sql.AppendLine("SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))) AS AVAIL")
            sql.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1, ICTSTAT2 S2")
            sql.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE")
            sql.AppendLine("AND C1.STYLE_CODE = S2.STYLE_CODE (+)")
            sql.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
            sql.AppendLine("AND S2.WHSE_CODE = 'MS'")
            sql.AppendLine("GROUP BY")
            sql.AppendLine("S1.STYLE_CODE,")
            sql.AppendLine("C1.COLOR_CODE,")
            sql.AppendLine("S1.STYLE_STATUS,")
            sql.AppendLine("C1.STYLE_COLOR_STATUS,")
            sql.AppendLine("S1.STYLE_DESC")
            sql.AppendLine(")")
            If chkNoDiscInvPics.Checked Then
                sql.AppendLine("WHERE (STYLE_COLOR_STATUS = 'A' OR AVAIL > 0)")
            End If
            Dim tblMFLIST As DataTable = ASCDATA1.GetDataTable(sql.ToString())

            For Each image As KeyValuePair(Of String, Date) In ImageListFTPtmp
                'If image.Key = "MT17903-CAIR.jpg" Then Stop
                If tblMFLIST.Select($"SC = '{image.Key.Replace("'", "").ToUpper}'").Length > 0 Then
                    If ImageListLocaltmp.Keys.Contains(image.Key) Then
                        If image.Value > ImageListLocaltmp.Item(image.Key) Then
                            ImageListDownload.Add(image.Key)
                        End If
                    Else
                        ImageListDownload.Add(image.Key)
                    End If
                End If
            Next
            'For Each localFile As KeyValuePair(Of String, Date) In ImageListLocal
            '    If Not ImageListFTP.Keys.Contains(localFile.Key) Then
            '        ImageListDelete.Add(localFile.Key)
            '    End If
            'Next
        Else
            MsgBox("Error with Image Parameters", MsgBoxStyle.Critical, "Parameters")
        End If
        ASCMAIN1.Progress("")
        txtLocalImages.Value = ImageListLocal.Count
        txtUpdateImages.Value = ImageListDownload.Count
        txtDeleteImages.Value = ImageListDelete.Count
        FTPImages = False
        btnGetImages.Visible = True
    End Sub

    Private Sub btnUpdateImages_Click(sender As Object, e As EventArgs) Handles btnUpdateImages.Click
        FillImageStats()
    End Sub

    Private Sub btnGetImages_Click(sender As Object, e As EventArgs) Handles btnGetImages.Click
        If btnGetImages.Text = "Get Images" Then
            btnGetImages.Text = "Stop Get"
            ImageGetProgress = True
            getFTPImages()
            ImageGetProgress = False
            ASCMAIN1.Progress("")
            btnGetImages.Text = "Get Images"
        Else
            btnGetImages.Text = "Get Images"
            ImageGetProgress = False
            ASCMAIN1.Progress("")
        End If
        Application.DoEvents()
    End Sub

    Private Sub getFTPImages()
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim IMAGES_FOLDER As String = "C:\"
        Dim RemoteFolder As String = "/www/media/product/"
        If Not IsNothing(rowSOTPARM3) Then
            If rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString.EndsWith("\") Then
                IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            Else
                IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString & "\"
            End If
        End If
        Ftp1.User = "regency-rib"
        Ftp1.Password = "joydHUJ3"
        Ftp1.RemoteHost = "regency-rib.com"
        Ftp1.Logon()
        'Ftp1.TransferMode = nsoftware.IPWorks.FTPTransferModes.tmBinary
        Ftp1.ChangeTransferMode(nsoftware.IPWorks.FTPTransferModes.tmBinary)
        Ftp1.RemoteFile = RemoteFolder & "*"
        Ftp1.LocalFile = IMAGES_FOLDER & "*"
        Ftp1.Overwrite = True
        For Each DLFile As String In ImageListDownload
            Application.DoEvents()
            If ImageGetProgress Then
                Ftp1.LocalFile = IMAGES_FOLDER & DLFile
                Ftp1.RemoteFile = RemoteFolder & DLFile
                ASCMAIN1.Progress("Fetching " & DLFile)
                Ftp1.Download()
            Else
                Exit For
            End If
        Next
        Ftp1.Logoff()
        'For Each DELFile As String In ImageListDelete
        '    System.IO.File.Delete(IMAGES_FOLDER & DELFile)
        '    ASCMAIN1.Progress("Deleting " & DELFile)
        'Next
        FillImageStats()
        MsgBox("File Sync Complete!", vbOKOnly, "Sync")
        btnGetImages.Visible = False
    End Sub

    Private Sub SaveFavorites()
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT COUNT(TABLE_NAME) AS TBL_CNT FROM USER_TABLES WHERE TABLE_NAME = 'ASTMENU2_TMP'")
        ASCMAIN1.sql = SQLS.ToString()
        Dim TBL_CNT As Int16 = Val(ASCDATA1.GetDataValue)

        If TBL_CNT = 1 Then
            SQLS.Length = 0
            SQLS.AppendLine("DROP TABLE ASTMENU2_TMP")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
        End If

        For Each SQLString As String In New String() {"CREATE TABLE ASTMENU2_TMP AS SELECT * FROM ASTMENU2"}
            ASCMAIN1.sql = SQLString
            ASCDATA1.ExecuteSQL()
        Next
    End Sub

    Private Sub RestoreFavorites()
        For Each SQLString As String In New String() {"DELETE FROM ASTMENU2",
                                                              "INSERT INTO ASTMENU2 SELECT * FROM ASTMENU2_TMP"}
            ASCMAIN1.sql = SQLString
            ASCDATA1.ExecuteSQL()
        Next
    End Sub

    Private Sub grdTATCTLN0_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATCTLN0.InitializeLayout

    End Sub

    Private Sub lblVersionNo_Click(sender As Object, e As EventArgs) Handles lblVersionNo.Click
        MsgBox(VersionInfo.ToString, vbOKOnly, "This Version")
    End Sub

    Private Sub btnGetQuote_Click(sender As Object, e As EventArgs) Handles btnGetQuote.Click
        Dim eMsg As New System.Text.StringBuilder With {.Length = 0}
        Dim QUOTE_NO As String = ""
        Dim ORDR_NO As String = ""
        If Not IsNumeric(txtQuoteNumber.Text) Then
            eMsg.AppendLine("Invalid Quote Number.  Must Be Numberic.")
        Else
            If txtQuoteNumber.Text.Length <> 10 Then
                txtQuoteNumber.Text = txtQuoteNumber.Text.PadLeft(10, "0")
            Else
                QUOTE_NO = txtQuoteNumber.Text
            End If
        End If
        If eMsg.Length = 0 Then
            Fill_Records("SOTQRDR1", QUOTE_NO)
            Fill_Records("SOTQRDR2", QUOTE_NO)
            Fill_Records("SOTQRDR5", QUOTE_NO)
            If dst.Tables.Item("SOTQRDR1").Rows.Count = 1 Then
                'ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                'Stop
                'You need to follow the way Order Entry gets a new number.
                Dim TATCTLN3 As New TATCTLN3("SOTORDR1.ORDR_NO", Me)
                If Not IsNothing(TATCTLN3.ErrMsg) Then
                    MsgBox(TATCTLN3.ErrMsg, MsgBoxStyle.OkOnly, "Problem Getting Next Order Number")
                    Exit Sub
                End If
                If TATCTLN3.NumbersRemaining < 10 Then
                    Dim msg As String = String.Format("You Only Have {0} Order Numbers Left", TATCTLN3.NumbersRemaining)
                    msg = msg & vbCrLf & "You Should Fetch Some More From The Transfer Screen Soon."
                    MsgBox(msg, MsgBoxStyle.Critical, "Running Low On Order Numbers")
                End If
                ORDR_NO = TATCTLN3.Next_ctl_no
                Dim ORDR_BATCH_NO As String = ASCMAIN1.Next_Control_No("ORDR_BATCH_NO")
                Dim API_MSG As String = GetQuoteFromAPI(QUOTE_NO, ORDR_NO)
                If API_MSG = ORDR_NO Then
                    For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select()
                        Dim newSOTORDR1_Q As DataRow = dst.Tables.Item("SOTORDR1_Q").NewRow
                        For Each dc As DataColumn In dst.Tables.Item("SOTORDR1_Q").Columns
                            Dim name As String = dc.ColumnName
                            Select Case name
                                Case "ORDR_NO"
                                    newSOTORDR1_Q.Item(name) = ORDR_NO
                                Case "ORDR_GROUP_NO"
                                    newSOTORDR1_Q.Item(name) = ORDR_BATCH_NO
                                Case "ORDR_STATUS"
                                    newSOTORDR1_Q.Item(name) = "Q"
                                Case Else
                                    newSOTORDR1_Q.Item(name) = rowSOTQRDR1.Item(name)
                            End Select
                        Next
                        dst.Tables.Item("SOTORDR1_Q").Rows.Add(newSOTORDR1_Q)
                    Next
                    For Each rowSOTQRDR2 As DataRow In dst.Tables("SOTQRDR2").Select()
                        Dim newSOTORDR2_Q As DataRow = dst.Tables.Item("SOTORDR2_Q").NewRow
                        For Each dc As DataColumn In dst.Tables.Item("SOTORDR2_Q").Columns
                            Dim name As String = dc.ColumnName
                            Select Case name
                                Case "ORDR_NO"
                                    newSOTORDR2_Q.Item(name) = ORDR_NO
                                Case Else
                                    newSOTORDR2_Q.Item(name) = rowSOTQRDR2.Item(name)
                            End Select
                        Next
                        dst.Tables.Item("SOTORDR2_Q").Rows.Add(newSOTORDR2_Q)
                    Next
                    For Each rowSOTQRDR5 As DataRow In dst.Tables("SOTQRDR5").Select()
                        Dim newSOTORDR5_Q As DataRow = dst.Tables.Item("SOTORDR5_Q").NewRow
                        For Each dc As DataColumn In dst.Tables.Item("SOTORDR5_Q").Columns
                            Dim name As String = dc.ColumnName
                            Select Case name
                                Case "ORDR_NO"
                                    newSOTORDR5_Q.Item(name) = ORDR_NO
                                Case Else
                                    newSOTORDR5_Q.Item(name) = rowSOTQRDR5.Item(name)
                            End Select
                        Next
                        dst.Tables.Item("SOTORDR5_Q").Rows.Add(newSOTORDR5_Q)
                    Next

                    Update_Record_TDA("SOTORDR1_Q")
                    Update_Record_TDA("SOTORDR2_Q")
                    Update_Record_TDA("SOTORDR5_Q")
                Else
                    eMsg.AppendLine(API_MSG)
                End If
            Else
                eMsg.AppendLine("Quote Specified Is Not Yet Downloaded.")
                eMsg.AppendLine("Please Do A Data Transfer And Try Again.")
            End If

        End If
        If eMsg.Length > 0 Then
            MsgBox(eMsg.ToString, vbCritical, "Error Getting Quote")
        Else
            eMsg.AppendLine("Your Quote Has Been Imported From")
            eMsg.AppendLine("Regency's Main System.")
            eMsg.AppendLine("")
            eMsg.AppendLine("It Has Been Re-Assigned To A New")
            eMsg.AppendLine(String.Format("Quote Number On Your Laptop: {0}", ORDR_NO))
            eMsg.AppendLine("")
            eMsg.AppendLine("You Can Find The New Quote In The")
            eMsg.AppendLine("Order Entry Screen.")
            MsgBox(eMsg.ToString, vbCritical, "Your Quote Awaits")
        End If
    End Sub

    Private Function GetQuoteFromAPI(ByVal QUOTE_NO As String, ByVal ORDR_NO As String) As String
        Dim RetVal As String = ""
        'Dim URL As String = "https://localhost:44344/api/WebQuotes/" & QUOTE_NO
        Dim URL As String = "https://com.regency-rib.com:8090/api/WebQuotes/" & QUOTE_NO
        Dim client As New HttpClient()
        client.BaseAddress = New Uri(URL)

        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))
        client.DefaultRequestHeaders.Add("NEW_ORDR_NO", ORDR_NO)
        client.DefaultRequestHeaders.Add("WEBQUOTEKEY", "9jngTdghY@<hf9hg632gfe7![jjy865gf542gveKjhhghwuy8#jnshyu0)nH")

        client.Timeout = New TimeSpan(0, 5, 0)

        Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()
        'Dim frmtr As MediaTypeFormatter = New String()

        Dim content As HttpContent = New ObjectContent(Of String)(ORDR_NO, frmtr)

        Dim response As HttpResponseMessage = client.GetAsync(URL).Result
        If response.IsSuccessStatusCode Then
            Try
                Dim TATCTLN1_RESPONSE As New TATCTLN1
                Dim apiResponseString As String = ""
                Dim responseObject As Object = response.Content.ReadAsAsync(Of String)().Result

                RetVal = responseObject
            Catch ex As Exception
                RetVal = ex.InnerException.ToString
            End Try

        Else
            RetVal = "Can Not Cponnect To Serve"
        End If

        'RetVal = Newtonsoft.Json.JsonConvert.SerializeObject(resp)

        Return RetVal
    End Function

    Private Sub btnPWD_Click(sender As Object, e As EventArgs) Handles btnPWD.Click
        If ASCMAIN1.USER_ID = "mariog" Or ASCMAIN1.USER_ID = "wayne" Then
            Dim eMsg As New Text.StringBuilder With {.Length = 0}
            For i As Int64 = 0 To 6
                Dim thisDt As Date = CDate(Now().AddDays(i).ToShortDateString)
                eMsg.AppendLine($"{thisDt.ToShortDateString} : {TodaysPwd(thisDt)}")
            Next
            MsgBox(eMsg.ToString, vbOKOnly, "Update Passwords")
        End If
    End Sub

    Private Function TodaysPwd(ByVal DateForPW As Date) As String
        DateForPW = CDate(DateForPW.ToShortDateString)
        Dim retval As String = ""
        Dim D As Int64 = DateForPW.Day
        Dim M As Int64 = DateForPW.Month
        Dim Y As Int64 = Val(DateForPW.Year.ToString.Substring(2, 2))
        Dim E As Int64 = D Mod 2
        Dim C1 As String = Chr(M + 64)
        Dim C2 As String = Chr(Y - M + 64)
        Dim C3 As String = Chr(D + 64)
        If D >= 10 Then
            If E = 0 Then
                retval = $"{C1}{C2}{C3}"
            Else
                retval = $"{C3}{C1}{C2}"
            End If
        Else
            If E = 0 Then
                retval = $"{C3}{C2}{C1}"
            Else
                retval = $"{C1}{C3}{C2}"
            End If
        End If
        Return retval
    End Function

    Private Function CHECK_STYLE_COLOR_DUPS(ByVal ORDR_NOs As List(Of String), ByVal NextStep As String) As Boolean
        '--This function is found in several places, Laptop order entry, Laptop Transfers & order imports.
        '--If you make changes here you should make those changes there as well.  Someday we will move this
        '--to a shared class or evne beter find the problem that allows duplicate style/colors.
        Dim RETVAL As Boolean = False
        Dim EMSG As String = ""
        For Each ORDR_NO As String In ORDR_NOs
            Dim filter As String = $"ORDR_NO = '{ORDR_NO}'"
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select(filter).FirstOrDefault
            If Not IsNothing(rowSOTORDR1) Then
                Dim STYLE_COLORS_CHK As New List(Of String)
                Dim STYLE_COLORS_DUP As New List(Of String)
                Dim CUST_NAME As String = rowSOTORDR1.Item("CUST_NAME").ToString & String.Empty
                Dim rowFilter As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(rowFilter, "STYLE_CODE, COLOR_CODE")
                    Dim STYLE_COLOR As String = rowSOTORDR2.Item("STYLE_CODE").ToString & String.Empty & "-" & rowSOTORDR2.Item("COLOR_CODE").ToString
                    If STYLE_COLORS_CHK.Contains(STYLE_COLOR) Then
                        STYLE_COLORS_DUP.Add(STYLE_COLOR)
                    Else
                        STYLE_COLORS_CHK.Add(STYLE_COLOR)
                    End If
                Next
                If STYLE_COLORS_DUP.Count > 0 Then
                    If RETVAL = False Then
                        RETVAL = True
                        EMSG &= "The Following Duplicate Style / Colors <br>"
                        EMSG &= "Were Found On These Orders.  You Should <br>"
                        EMSG &= $"Fix Them Before {NextStep}. <br>"
                        EMSG &= "<br><hr>"
                        EMSG &= $"Order: {ORDR_NO} - {CUST_NAME} <br>"
                        For Each STYLE_COLOR As String In STYLE_COLORS_DUP
                            EMSG &= $"     - {STYLE_COLOR} <br>"
                        Next
                        EMSG &= "<br>"
                    Else
                        EMSG &= $"Order: {ORDR_NO} - {CUST_NAME} <br>"
                        For Each STYLE_COLOR As String In STYLE_COLORS_DUP
                            EMSG &= $"     - {STYLE_COLOR} <br>"
                        Next
                        EMSG &= "<br>"
                    End If
                End If
            End If
        Next
        If RETVAL Then
            Using frmmsg As New ASFMSGBF
                frmmsg.Show_Formatted_txt("You Must Fix This Before Transmitting The Order(s).", EMSG, Me)
            End Using
        End If

        Return RETVAL
    End Function
    Private Sub setVersionNo()
        Dim VersionNo As String = ""

        'VersionNo = "17.12.19.2"
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Fixes Issue With Find Attibute When Adding By Order Missing Styles.")
        'VersionInfo.AppendLine("* Buyer Information No Longer Manditory.")

        'VersionNo = "17.12.19.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Seperates The Transmission Functions Into Seperate Steps.")

        'VersionNo = "17.12.01.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Feature To Capture Buyer Information When Entering Orders.")

        'VersionNo = "19.03.08.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Feature Allow Bentlys to Work with Suzie.")

        'VersionNo = "19.05.31.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Capture And Send Contact Cell Phone Info.")
        'VersionInfo.AppendLine("* Fix Bug For Mario When Adding Contacts On Fly In Order Entry.")

        'VersionNo = "19.06.13.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Capture And Buyer Information Live and Manditory.")

        'VersionNo = "19.07.01.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Customer Masterfile to Capture Invoice E-mail Info.")

        'VersionNo = "19.07.02.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes to Order Entry and SKU Inquiry To Expose Cartons / Unit with Colors.")

        'VersionNo = "19.07.17.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes to box qty rules to lock down for zero availability items.")

        'VersionNo = "19.09.10.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Matrix For Sales Reps With Combined Accounts.") 'James and Dimple.

        'VersionNo = "19.10.10.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Sales Reps With Combined Accounts.") 'TN, CB, JB

        'VersionNo = "19.11.07.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Allow Item Pricing In Tablet Management.")

        'VersionNo = "19.12.03.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Add Lighting To Find By Attribute.")
        'VersionInfo.AppendLine("* Changes To Tablet Mangement For New iPads.")

        'VersionNo = "19.12.03.2"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Chanages to Order Entry Box Check to exclude RIBB.")

        'VersionNo = "19.12.08.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To FD Pricing During Show.")

        'VersionNo = "19.12.17.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* New FEFD Re-pricing For Rich.")

        'VersionNo = "20.01.02.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* ShipTo Capture Feature Out Live.")

        'VersionNo = "20.01.05.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Adding Check For Duplicate style/colors.")

        'VersionNo = "20.01.09.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Tablet Management Adds Missing Order Details Attributes.")

        'VersionNo = "20.05.02.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Update for Changes To Order Re-pricing.")

        'VersionNo = "20.06.16.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Item Pack Join.")

        'VersionNo = "20.11.20.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Web Quote System Going Live.")

        'VersionNo = "20.12.31.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Update to Sync With ABSolution Security Changes.")

        'VersionNo = "21.02.27.2"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change To Stop Disc & DNR Ordering Past Qty Avail.")

        'VersionNo = "21.03.17.1"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change To Fix Time Stamp Issue in FEFD Pricing.")

        'VersionNo = "21.04.28.2"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Addition of tariff code info to find style.")

        'If (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "mariog") Then
        'VersionNo = "21.05.05.2"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Secure FTP.")

        'VersionNo = "21.05.05.17"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Upgrade Of Secure FTP.")

        'VersionNo = "21.06.03.13"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes To Allocation Grids To Mimic Big ABS.")

        'VersionNo = "21.06.26.15"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes Allow Quotes Accept DNR with Qty > OH.")

        'VersionNo = "21.07.08.16"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Multi-Hang Tag Option.")

        'VersionNo = "21.07.11.17"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Default Term Code on New Customers Set to CRED.")
        'VersionInfo.AppendLine("* Add Duty Rate To Search By Attribute.")
        'VersionInfo.AppendLine("* Default Locations in Image Management.")

        'VersionNo = "21.12.09.17"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes to Image Mapper For Michale.")
        'VersionInfo.AppendLine("* Changes to Order Entry To Allow FE Disc Ordering With Warning.")

        'VersionNo = "22.01.03.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes to Add Extended PVC System To Laptops.")

        'VersionNo = "22.01.21.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes to Search By Attribute For Importing Spreadsheets.")

        'VersionNo = "22.03.21.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes to Search By Attribute For PVC Items.")

        'VersionNo = "22.11.18.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Upgrade to Label Printing.")

        'VersionNo = "22.11.18.02"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Upgrade to Encryption.")

        'VersionNo = "22.11.18.03"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Remove Software Update Option Unless Mario.")

        'VersionNo = "22.12.23.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Software Updates Are Locked With Passwords.")

        'VersionNo = "23.01.26.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change SO Excel To Show Order List Price.")
        'VersionInfo.AppendLine("* Proviid Multiple Days Of Passwords.")

        'VersionNo = "23.02.22.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change SO Extended PVC for Importing.")

        'VersionNo = "23.06.30.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change SO Excel To Show List From MF when Order Is Missing.")

        'VersionNo = "23.08.03.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change Search By Attribute For Stock/Non-Stock.")

        'VersionNo = "23.09.4.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change Order Transfer set Date Received & Init to Today.")

        'VersionNo = "23.11.30.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change Order Printing To Show Sub-UPCs.")

        'VersionNo = "24.02.15.03"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* New Attribute Excel Feature Added to SSBA Screen.")

        'VersionNo = "24.03.06.02"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* New Order Entry Excel Features.")

        'VersionNo = "24.05.03.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* New Feature to Pull Images From Web.")

        'VersionNo = "24.05.30.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Changes to Web Images to Optionaly Use Only Active/Disc > 0.")

        'VersionNo = "24.06.27.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change To Search By Attribute For Sorting And Discount Pricing When Printing.")

        'VersionNo = "24.07.25.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Change To Customer Matrix to show all orders.")

        'VersionNo = "24.09.06.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Fix Issue With Search By Attribute To Ignore ECommerce When Running On Laptops.")

        'VersionNo = "24.10.24.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* New version to solve newtonsoft issues and have new full distribution.")

        'VersionNo = "24.11.26.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Hang Tag Changes and Tariff Warnings.")

        'VersionNo = "24.12.08.01"
        'VersionInfo.AppendLine("")
        'VersionInfo.AppendLine(VersionNo)
        'VersionInfo.AppendLine("* Option To Show / Hide Discontinued Colors In Order Entry.")

        VersionNo = "25.01.22.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* API Calls Moved To Use New API Server.")

        VersionNo = "25.04.17.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Changes to Tariff Notices On Various Output.")

        VersionNo = "25.06.11.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Changes to Tariff Notices On Various Output.")

        VersionNo = "25.07.21.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Changes to Rollup MS and US Warehouse Data.")

        VersionNo = "25.09.04.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Prevent Transferring Orders that Have Dups.")

        VersionNo = "25.10.16.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Update To Transfer Process To Sync Pricing Tiers.")
        VersionInfo.AppendLine("* Warn When Starting Order If Customer Is On Credit Hold.")

        VersionNo = "25.12.03.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Stop When Starting Order If Customer Is On Sales Hold.")

        VersionNo = "25.12.07.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Expose 2nd Description In Order Entry.")

        VersionNo = "25.12.11.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Remove Factory From SKU Inq and Order Entry.")

        VersionNo = "25.12.18.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Addition Of Sales And Credit Hold To Cust Matrix.")

        VersionNo = "26.01.06.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Changes To Order Excel.")

        VersionNo = "26.03.26.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Add 3 Month Option To Orders Filter.")

        VersionNo = "26.04.22.01"
        VersionInfo.AppendLine("")
        VersionInfo.AppendLine(VersionNo)
        VersionInfo.AppendLine("* Changes to Contacts to Create Lead Contacts.")

        lblVersionNo.Text = VersionNo
    End Sub
End Class