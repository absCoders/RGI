Public Class POFCENT1

    Dim WithEvents FTP1 As nsoftware.IPWorks.Ftp
    Dim REMOTEDIRECTORYFILELIST As List(Of String) = New List(Of String)
    Dim displaycontrol As Control = Nothing

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")
        Get_PARM("POTPARMC")
        With dst

            ASCMAIN1.sql = "Select POTCENT2.PO_SHIPMENT_NO, POTCENT2.LM_VESSEL_NM PO_SHIP_VESSEL" & vbCrLf _
                & ", Decode(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO) PO_SHIP_REF_NO" & vbCrLf _
                & ", Min(POTCENT2.EST_ARRIVE_DT) PO_SHIP_ETA, Min(POTCENT2.LM_CLOSE_DT) PO_SHIP_ADV_DATE" & vbCrLf _
                & ", Min(POTCENT2.LM_EST_DEPART_DT) PO_DATE_SHIPPED" & vbCrLf _
                & ", Min(POTCENT2.COM_INVOICE_NO) PO_NOTES" & vbCrLf _
                & ", Min(POTCENT2.LOAD_PORT_CD) PORT_CODE_ORIG, Min(POTCENT2.DISC_PORT_CD) PORT_CODE_DEST" & vbCrLf _
                & ", Min(POTCENT2.LM_VOYAGE_NO) VOYAGE_NO " & vbCrLf _
                & " from POTCENT2 where PO_CENT_STATUS = '0'" & vbCrLf _
                & " group by POTCENT2.PO_SHIPMENT_NO, POTCENT2.LM_VESSEL_NM" & vbCrLf _
                & ", Decode(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO)"
            Create_TDA(.Tables.Add, "POTSHIPS", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from POTCENT2 where PO_SHIPMENT_NO =  :PARM1 AND PO_CENT_STATUS = '0' "
            Create_TDA(.Tables.Add, "POTCENT2", "**", 0, True, "V", 0)

            ASCMAIN1.sql = "Select * from POTCENT1 where PO_CENT_STATUS = '0' "
            Create_TDA(.Tables.Add, "POTCENT1", "**", 0, True, "", 1)


            Create_TDA(.Tables.Add, "POTSHIP1", "*", 0, True, "V", 0)
            Create_TDA(.Tables.Add, "POTSHIP2", "*", 0, True, "V", 0)
            Create_TDA(.Tables.Add, "POTSHIP3", "*", 0, True, "V", 0)

            Create_TDA(.Tables.Add, "POTSHIP7", "*", 0, True, "V", 0)
            Create_TDA(.Tables.Add, "POTSHIP8", "*", 0, True, "V", 0)
            Create_TDA(dst.Tables.Add, "POTSHIP4", "*")
        End With

        grdPOTSHIPS.DataSource = dst.Tables("POTSHIPS")
        grdPOTCENT2.DataSource = dst.Tables("POTCENT2")

        Create_Summary(grdPOTSHIPS, "PO_SHIPMENT_NO", "Count")

        'Create_Summary(grdPOTSHIPS, "PO_CENT_NO", "Count")
        'Create_Summary(grdPOTCENT2, New String() {"EDI_ON_HAND_QTY", "EDI_ALLOCATED_QTY", "WHSE_QTY_ON_HAND", "VARIANCE"})

        spl.Panel1Collapsed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Retrieve File"

            Case "Build Shipments"

                If grdPOTSHIPS.ActiveRow Is Nothing Then
                    EMsg &= vbCr & "No Shipments to be Built"
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

            Case "Retrieve File"

                retrieve_file()

                'PROCESS_file("C:\VS\VDI\TEMP\ASN\1.TXT")
                'PROCESS_file("C:\VS\VDI\TEMP\ASN\2.TXT")
                'PROCESS_file("C:\VS\VDI\TEMP\ASN\3.TXT")
                'PROCESS_file("C:\VS\VDI\TEMP\ASN\4.TXT")
                'PROCESS_file("C:\VS\VDI\TEMP\ASN\5.TXT")
                'Process_File("C:\VS\VDI\TEMP\ASN\6.TXT")

                Dim foldername As String = "S:\century"
                'foldername = "C:\Users\wjz\Desktop\RGI\century"
                foldername = "C:\century\IN"

                'foldername = "e:\century\IN"

                ' BELOW IS WHERE I NEED TO PUT IN ARCHIEVE

                For Each filename As String In My.Computer.FileSystem.GetFiles(foldername)
                    Process_File(filename)
                Next
                Dim ARCHIEVE_FOLDER As String = "S:\centuryhist"
                Dim archieve_file As String = ""
                'Dim ARCHIVE_FOLDER As String = ROWs("POTPARMC").Item("PO_PARM_ARCH_FOLDER")
                For Each filename As String In My.Computer.FileSystem.GetFiles(foldername)
                    archieve_file = Mid(filename, Len(foldername) + 1)
                    'My.Computer.FileSystem.CopyFile(filename, ARCHIVE_FOLDER, True)
                Next

                ASCMAIN1.sql = "" _
                    & " UPDATE POTPPRM1 SET CENT_IMP_EXECUTE_OPER = '" & ASCMAIN1.USER_ID & "',  CENT_IMP_EXECUTE_DATE = SYSDATE WHERE POTPPRM1_CODE = 'Z' "
                ASCDATA1.ExecuteSQL()

                Dim DRCSTRING As String = ""

                Load_POTSHIPS()

            Case "Build Shipments"

                build_shipments()

            Case "Load History"
                'ASCDATA1.ExecuteSQL("Truncate Table POTCENT1")
                'ASCDATA1.ExecuteSQL("Truncate Table POTCENT2")

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Processing File", "")

                ' "\\rbssco\century\histABS"
                Dim folder As String = "\\192.168.110.100\century\histABS"
                If ASCMAIN1.Running_in_VS Then folder = "C:\VS\VDI\Work\century"
                folder = "S:\RGI\Work\century"
                If ASCMAIN1.useUNCPath Then
                    folder = $"{ASCMAIN1.Folders("SharedRoot")}\RGI\Work\century"
                End If
                For Each FILENAME As String In My.Computer.FileSystem.GetFiles(folder)
                    ' Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                    ASCMAIN1.Progress("-", FILENAME)
                    BeginTrans()
                    Process_File(FILENAME)
                    CommitTrans()
                Next

                ASCMAIN1.Progress("", "")
                Me.Cursor = Cursors.Default

                Load_POTSHIPS()

                MsgBox("All Files Loaded")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Retrieve File").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Build Shipments").Settings.Enabled = not_iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"POTCENT1", "POTCENT2", "POTSHIPS"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Load_POTSHIPS()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then

        Else

        End If


        If EntryMode = "N" Then
        Else

        End If

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "SEASON_CODE"
            '    If Absx1.optFor("STMT_TYPE").CheckedIndex <> -1 Then
            '        sql_where = "STMT_TYPE = '" & Absx1.optFor("STMT_TYPE") & "'"
            '    End If
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()

        ' Load_Popup_Menu(grdEDT855, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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

        Select Case e.Tool.Key
            'Case "Style Multi-Color"
            '    Using F As New TAC.ICFSTYCX
            '        F.STYLE_CODE = ""
            '        F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
            '        F.ShowDialog()
            '        If F.STYLE_CODE <> "" Then
            '            Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
            '        End If
            '    End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "PO Inquiry"
            '    Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
            '    Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            'Case "Style Status Inquiry"
            '    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
            '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            '    If rowICTSTYL1 IsNot Nothing Then
            '        Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

            'Case "PO Shipment Inquiry"
            '    Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
            '    Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "LP_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Me.UltraGroupBox1.Select() ' to force txt_Leave event to fire, for formatting
            '        Load_EDT846T1()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "LP_CODE"
            '    Load_EDT846T1()
        End Select
    End Sub


    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        With Absx1.txtFor(COLUMN_NAME)
            Select Case COLUMN_NAME

                'Case "LP_CODE"
                '    Load_EDT846T1()

            End Select

        End With
    End Sub

#End Region

    Sub Load_POTSHIPS()

        Fill_Records("POTSHIPS")
        Sort_grdColumns(grdPOTSHIPS, "PO_SHIPMENT_NO")
        setup_potcent2()
    End Sub

    Private Sub grdPOTSHIPS_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIPS.AfterRowActivate
        setup_potcent2()
    End Sub

    Sub setup_potcent2()
        If grdPOTSHIPS.ActiveRow Is Nothing Then
            grdPOTCENT2.Visible = False
        Else
            Dim PO_SHIPMENT_NO As String = grdPOTSHIPS.ActiveRow.Cells("PO_SHIPMENT_NO").Value & ""
            Fill_Records("POTCENT2", PO_SHIPMENT_NO)
        End If
    End Sub

    Sub retrieve_file()

        Dim ifilesdownloaded As Long
        Dim filename As String = String.Empty
        Dim validdownloadfile As Boolean = False

        Try
            FTP1 = New nsoftware.IPWorks.Ftp
            FTP1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
            FTP1.RemoteHost = ROWs("POTPARMC").Item("PO_PARM_FTP_SITE")
            FTP1.RemotePath = ROWs("POTPARMC").Item("PO_PARM_INBOUND_FOLDER")
            FTP1.User = ROWs("POTPARMC").Item("PO_PARM_FTP_LOGIN")
            FTP1.Password = ROWs("POTPARMC").Item("PO_PARM_FTP_PWD")

            FTP1.RemotePath = "/Outbound/856/"

            FTP1.Logon()

            FTP1.RemotePath = "/Outbound/856/"

            If Not FTP1.Connected Then
                EMsg = "Connection to FTP Site Failed "
                MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
                Exit Sub
            End If

            Dim downloadedfiles As List(Of String) = New List(Of String)

            FTP1.ListDirectory()

            FTP1.Overwrite = True

            ifilesdownloaded = 0

            For Each remotediritem As String In REMOTEDIRECTORYFILELIST
                filename = remotediritem.Trim.ToUpper
                validdownloadfile = True

                If validdownloadfile Then
                    If displaycontrol IsNot Nothing Then

                        displaycontrol.Text = "downloading " & filename

                    End If
                End If

                FTP1.RemoteFile = filename
                FTP1.LocalFile = ROWs("POTPARMC").Item("PO_PARM_DOWNLOAD_FOLDER") & filename
                FTP1.Download()
                FTP1.DeleteFile(FTP1.RemotePath & "/" & filename)
                FTP1.DoEvents()

                My.Computer.FileSystem.CopyFile(ROWs("POTPARMC").Item("PO_PARM_DOWNLOAD_FOLDER") & filename, ROWs("POTPARMC").Item("PO_PARM_ARCH_FOLDER") & "\" & filename, True)

                BeginTrans()

                Process_File(FTP1.LocalFile)

                CommitTrans()

                downloadedfiles.Add(filename)
                ifilesdownloaded += 1
            Next

            FTP1.Logoff()

        Catch ex As Exception
            EMsg = "there were errors "
            MsgBox(EMsg & vbCr & ex.Message, MsgBoxStyle.OkOnly)
            Exit Sub


        End Try

        Try
            If FTP1.Connected Then FTP1.Logoff()
        Catch ex1 As Exception
            MsgBox(EMsg & vbCr & ex1.Message, MsgBoxStyle.OkOnly)
            Exit Sub
        End Try

        ASCMAIN1.sql = "" _
            & " UPDATE POTPPRM1 SET CENT_IMP_EXECUTE_OPER = '" & ASCMAIN1.USER_ID & "',  CENT_IMP_EXECUTE_DATE = SYSDATE WHERE POTPPRM1_CODE = 'Z' "
        ASCDATA1.ExecuteSQL()


    End Sub

    Sub Build_Shipments()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Shipments", "")

        Dim sqlPO_SHIPMENT_NO As String = ""
        Dim PO_SHIPMENT_NOs As New List(Of String)

        BeginTrans()

        Dim do_not_commit As Boolean = False

        Dim CONTAINER_TYPE_CODEs As New Dictionary(Of String, String)
        Dim CONTAINER_SEAL_NOs As New Dictionary(Of String, String)

        ' Dim PO_ORDER_NOs As New List(Of String)

        For Each row As DataRow In dst.Tables("POTSHIPS").Select("")

            Dim PO_SHIPMENT_NO As String = row.Item("PO_SHIPMENT_NO")
            sqlPO_SHIPMENT_NO &= ",'" & PO_SHIPMENT_NO & "'"
            PO_SHIPMENT_NOs.Add(PO_SHIPMENT_NO)

            ASCMAIN1.Progress("-", PO_SHIPMENT_NO)

            Dim rowPOTSHIP1 As DataRow = dst.Tables("POTSHIP1").NewRow
            With rowPOTSHIP1
                .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIP_VESSEL") = Trim(Mid(row.Item("PO_SHIP_VESSEL"), 1, 20) & "")
                .Item("PO_SHIP_ETA") = row.Item("PO_SHIP_ETA")

                .Item("PO_SHIP_LANDING_LEAD_DAYS") = ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR")
                .Item("PO_SHIP_REF_NO") = Trim(row.Item("PO_SHIP_REF_NO") & "")
                .Item("PO_SHIP_ADV_DATE") = row.Item("PO_SHIP_ADV_DATE")
                .Item("PO_DATE_SHIPPED") = row.Item("PO_DATE_SHIPPED")
                .Item("PO_NOTES") = Trim(row.Item("PO_NOTES") & "")
                .Item("PORT_CODE_ORIG") = row.Item("PORT_CODE_ORIG")
                .Item("PORT_CODE_DEST") = row.Item("PORT_CODE_DEST")
                .Item("VOYAGE_NO") = Trim(row.Item("VOYAGE_NO") & "")
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("COST_CODE") = "FOB"

                .Item("COST_NO_DUTY") = "0"
                .Item("REVIEW") = "0"
                .Item("AIR_SHIP") = "0"
                .Item("COST_COMPLETE") = "0"

                .Item("FREIGHT_ENTERED_BY") = ROWs("POTPARM1").Item("PO_PARM_FREIGHT_ENTERED_BY")
                .Item("COST_FRT_METHOD") = ROWs("POTPARM1").Item("PO_PARM_COST_FRT_METHOD")

                Dim Missing_POs As New List(Of String)
                Dim BTB_issue_POs As New List(Of String)

                Fill_Records("POTCENT2", PO_SHIPMENT_NO)
                Dim PO_SHIPMENT_LNO As Int32 = 0
                Dim CARTON_NO As Int32 = 0

                Dim rowPOTSHIP2 As DataRow
                Dim ORDR_NO As String = ""
                Dim first As Boolean = False

                CONTAINER_TYPE_CODEs.Clear()
                CONTAINER_SEAL_NOs.Clear()

                For Each row2 As DataRow In dst.Tables("POTCENT2").Select("UNITS <> 0", "PO_SHIPMENT_LNO,PO_NO,ITEM_SIZE")

                    If PO_SHIPMENT_LNO <> Val(row2.Item("PO_SHIPMENT_LNO") & "") Then
                        ORDR_NO = ""
                        first = True
                        rowPOTSHIP2 = dst.Tables("POTSHIP2").NewRow
                        With rowPOTSHIP2
                            PO_SHIPMENT_LNO = row2.Item("PO_SHIPMENT_LNO")
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO

                            Dim CONTAINER_NO As String = Trim(row2.Item("CONTAINER_NO") & "")
                            .Item("CONTAINER_NO") = CONTAINER_NO

                            If Not CONTAINER_TYPE_CODEs.ContainsKey(CONTAINER_NO) Then
                                CONTAINER_TYPE_CODEs.Add(CONTAINER_NO, Trim(row2.Item("CONTAINER_SIZE") & ""))
                                CONTAINER_SEAL_NOs.Add(CONTAINER_NO, Trim(row2.Item("SEAL_NO") & ""))
                            End If

                            .Item("BOL_NO") = Trim(row2.Item("FCR_NO") & "")

                            .Item("PO_SHIP_STATUS") = "O"
                            .Item("ACCRUAL_STATUS") = "0"
                            .Item("CONTAINER_SIZE") = Trim(row2.Item("CONTAINER_SIZE") & "")

                            Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
                            Dim CARTONS As Integer = 0
                            Dim VOLUME As Decimal = 0 ' Val(dst.Tables("POTCENT2").Compute("SUM(VOLUME)", sqlw) & "")
                            Dim WEIGHT As Decimal = 0 ' Val(dst.Tables("POTCENT2").Compute("SUM(WEIGHT)", sqlw) & "")
                            For Each rowS As DataRow In dst.Tables("POTCENT2").Select(sqlw)
                                VOLUME += Val(rowS.Item("VOLUME") & "")
                                WEIGHT += Val(rowS.Item("WEIGHT") & "")
                                CARTONS += Val(rowS.Item("CARTONS") & "")
                            Next
                            .Item("PO_SHIP_CTNS") = CARTONS
                            .Item("CBM") = VOLUME
                            .Item("TOTAL_WEIGHT") = WEIGHT
                            .Item("COMM_INV_NO") = Trim(row2.Item("COM_INVOICE_NO") & "")
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("LAST_DATE") = DATETIME_STAMP
                            dst.Tables("POTSHIP2").Rows.Add(rowPOTSHIP2)
                        End With

                    End If

                    Dim PO_ORDER_NO As String = Trim(row2.Item("PO_NO"))
                    Dim PO_ORDER_LNO As Integer = Trim(row2.Item("ITEM_SIZE"))

                    Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                    'If Not PO_ORDER_NOs.Contains(PO_ORDER_NO) Then PO_ORDER_NOs.Add(PO_ORDER_NO)

                    If rowPOTORDR1 Is Nothing Then

                        If Not Missing_POs.Contains(PO_ORDER_NO) Then
                            If Not Missing_POs.Contains(PO_ORDER_NO) Then
                                MsgBox("PO " & PO_ORDER_NO & " is referenced in the Century Data, but is not on file", _
                                       MsgBoxStyle.OkOnly, "Cannot Load part of this Shipment")
                                Missing_POs.Add(PO_ORDER_NO)
                                do_not_commit = True
                            End If
                        End If
                    Else
                        If first Then
                            ORDR_NO = rowPOTORDR1.Item("ORDR_NO") & ""
                            first = False
                        Else
                            If ORDR_NO <> rowPOTORDR1.Item("ORDR_NO") & "" Then
                                If Not BTB_issue_POs.Contains(PO_ORDER_NO) Then
                                    BTB_issue_POs.Add(PO_ORDER_NO)
                                    do_not_commit = True
                                    MsgBox("PO " & PO_ORDER_NO & " is combined with other, non-compatible POs in the Century Data", _
                                           MsgBoxStyle.OkOnly, "Cannot mix BTB POs with other or non-BTB POs in a single Shipment Line")
                                End If
                            End If
                        End If

                        Dim rowPOTORDR2 As DataRow

                        Dim rowPOTSHIP3 As DataRow = dst.Tables("POTSHIP3").NewRow
                        With rowPOTSHIP3

                            rowPOTORDR2 = LookUp("POTORDR2", New String() {PO_ORDER_NO, PO_ORDER_LNO})
                            Dim STYLE_CODE As String = rowPOTORDR2.Item("STYLE_CODE")
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)

                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO

                            .Item("PO_QTY_SHP") = Trim(row2.Item("UNITS"))
                            .Item("PO_COST") = Val(rowPOTORDR2.Item("PO_COST") & "")
                            .Item("PO_COST_VCOST") = Val(rowPOTORDR2.Item("PO_COST") & "")
                            .Item("PO_COST_VCOST_UM") = Val(rowPOTORDR2.Item("PO_COST") & "")
                            .Item("PO_COST_VCOST_DZ") = Val(rowPOTORDR2.Item("PO_COST") & "") * 12

                            .Item("DUTY_RATE_CODE") = rowICTSTYL1.Item("DUTY_RATE_CODE")
                            .Item("DUTY_RATE") = 0
                            '.Item("WEIGHT_CODE") = rowICTSTYL1.Item("WEIGHT_CODE")

                            If Val(row2.Item("UNITS") & "") <> 0 Then
                                .Item("WEIGHT_FACTOR") = Val(row2.Item("WEIGHT") & "") / Val(row2.Item("UNITS") & "")
                            Else
                                .Item("WEIGHT_FACTOR") = 0
                            End If

                            .Item("FOB_CMT") = "B"

                            .Item("PO_COST_OTHER") = 0
                            .Item("PO_COST_COMM") = 0
                            .Item("PO_COST_OTHER_DZ") = 0

                            .Item("PO_COST_LANDED") = Val(rowPOTORDR2.Item("PO_COST") & "")

                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("LAST_DATE") = DATETIME_STAMP
                            dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)
                        End With

                        CARTON_NO += 1

                        Dim rowPOTSHIP7 As DataRow = dst.Tables("POTSHIP7").NewRow
                        With rowPOTSHIP7
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            .Item("CARTON_NO") = CARTON_NO
                            .Item("CARTONS") = Val(row2.Item("CARTONS") & "")
                            '.Item("CARTON_COMMENTS") = DBNull.Value
                            '.Item("CUSTOM_PPK") = DBNull.Value
                            '.Item("PPK_CODE") = DBNull.Value
                            If Val(row2.Item("CARTONS") & "") <> 0 Then
                                .Item("PO_QTY_PER_CTN") = Val(row2.Item("UNITS") & "") / Val(row2.Item("CARTONS") & "")
                                .Item("CARTON_DIMS") = CStr(CInt(100 * 100 * Val(row2.Item("VOLUME") & "") / Val(row2.Item("CARTONS") & "")) / 100) & "x100x100"
                                .Item("CARTON_VOLUME") = 1000000 * Val(row2.Item("VOLUME") & "") / Val(row2.Item("CARTONS") & "")
                                .Item("CARTON_WEIGHT") = Val(row2.Item("WEIGHT") & "") / Val(row2.Item("CARTONS") & "")
                            Else
                                .Item("PO_QTY_PER_CTN") = 1
                                .Item("CARTON_DIMS") = ""
                                .Item("CARTON_VOLUME") = 0
                                .Item("CARTON_WEIGHT") = 0
                            End If
                            .Item("STYLE_CODE") = Replace(Trim(row2.Item("ITEM_NO") & ""), " ", "")
                            .Item("COLOR_CODE") = Trim(row2.Item("ITEM_COLOR") & "")
                            '.Item("PPK_INNER_QTY") =  DBNull.Value
                            dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
                        End With

                        Dim rowPOTSHIP8 As DataRow = dst.Tables("POTSHIP8").NewRow
                        With rowPOTSHIP8
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                            .Item("CARTON_NO") = CARTON_NO
                            .Item("STYLE_CODE") = Replace(Trim(row2.Item("ITEM_NO") & ""), " ", "")
                            .Item("COLOR_CODE") = Trim(row2.Item("ITEM_COLOR") & "")
                            If Val(row2.Item("CARTONS") & "") <> 0 Then
                                .Item("QTY") = Val(row2.Item("UNITS") & "") / Val(row2.Item("CARTONS") & "")
                            Else
                                .Item("QTY") = Val(row2.Item("UNITS") & "")
                            End If

                            .Item("DOZENS") = DBNull.Value
                            .Item("PPK_INNER_QTY") = rowPOTORDR2.Item("INNER_PACK_QTY")
                            dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)
                        End With
                    End If
                Next

                If dst.Tables("POTSHIP2").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'").Length <> 0 Then

                    ' IF YOU RECEIVED A MESSAGE THAT A PO WAS NOT ON FILE, 
                    ' THEN YOU WILL BLOW UP ON THIS NEXT LINE 
                    ' BECAUSE THERE ARE NO POTSHIP2 RECORDS 
                    ' AS A RESULT OF THE CENTURY DATA THAT COULD NOT BE LINKED TO POS
                    ' WHICH IS WHY WE HAVE THE IF JUST ABOVE THIS COMMENT BLOCK

                    Dim CONTAINER_LNO As Integer = 0

                    For Each rowCONTAINER_NO As DataRow In ASCDATA1.SelectDistinct _
                      (dst.Tables("POTSHIP2").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"), New String() {"CONTAINER_NO"}) _
                      .Select("", "CONTAINER_NO")
                        Dim CONTAINER_NO As String = rowCONTAINER_NO.Item("CONTAINER_NO")

                        Dim rowPOTSHIP4 As DataRow = dst.Tables("POTSHIP4").NewRow
                        With rowPOTSHIP4
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            ' NOTE: POTSHIP4.PO_SHIPMENT_LNO IS NOT THE SAME THINK AS POTSHIP2.PO_SHIPMENT_LNO, AND PROBABLY SHOULD HAVE BEEN NAMED POTSHIP4.CONTAINER_LNO
                            CONTAINER_LNO += 1
                            .Item("PO_SHIPMENT_LNO") = CONTAINER_LNO
                            .Item("CONTAINER_NO") = CONTAINER_NO
                            Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and CONTAINER_NO = '" & CONTAINER_NO & "'"
                            Dim PO_SHIP_CTNS As Int32 = Val(dst.Tables("POTSHIP2").Compute("SUM(PO_SHIP_CTNS)", sqlw) & "")
                            .Item("PO_SHIP_CTNS") = PO_SHIP_CTNS
                            .Item("CONTAINER_TYPE_CODE") = CONTAINER_TYPE_CODEs(CONTAINER_NO)
                            .Item("CONTAINER_SEAL_NO") = CONTAINER_SEAL_NOs(CONTAINER_NO)
                            Dim TOTAL_WEIGHT As Int32 = Val(dst.Tables("POTSHIP2").Compute("SUM(TOTAL_WEIGHT)", sqlw) & "")
                            .Item("TOTAL_WEIGHT") = TOTAL_WEIGHT
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("LAST_DATE") = DATETIME_STAMP
                            dst.Tables("POTSHIP4").Rows.Add(rowPOTSHIP4)
                        End With
                    Next


                    ASCMAIN1.sql = "Update POTSHIP2" & vbCrLf _
                        & " Set ORDR_NO = (Select DISTINCT ORDR_NO from POTORDR1 where PO_ORDER_NO in " & vbCrLf _
                        & "(Select DISTINCT PO_ORDER_NO from POTSHIP3" & vbCrLf _
                        & " where PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                        & "   and PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO))" & vbCrLf _
                        & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
                    ASCDATA1.ExecuteSQL()
                End If

                'Dim P As String = Join(BTB_issue_POs.ToArray, ",")
                'Stop
            End With
            dst.Tables("POTSHIP1").Rows.Add(rowPOTSHIP1)

            ASCMAIN1.sql = "Update POTCENT2 Set PO_CENT_STATUS = '1' " _
            & " where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next

        Update_Record_TDA("POTSHIP1")
        Update_Record_TDA("POTSHIP2")
        Update_Record_TDA("POTSHIP3")
        Update_Record_TDA("POTSHIP4")
        Update_Record_TDA("POTSHIP7")
        Update_Record_TDA("POTSHIP8")

        ASCMAIN1.sql = "" _
           & "Begin" & vbCrLf _
           & " Declare Cursor C1 is " & vbCrLf _
           & "  Select POTSHIP3.PO_SHIPMENT_NO, MIN(POTORDR1.WHSE_CODE) WHSE_CODE" & vbCrLf _
           & "   from POTSHIP3,POTSHIP1,POTORDR1" & vbCrLf _
           & "   where POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
           & "     and POTSHIP3.PO_SHIPMENT_NO IN (" & Mid(sqlPO_SHIPMENT_NO, 2) & ")" & vbCrLf _
           & "   group by POTSHIP3.PO_SHIPMENT_NO;" & vbCrLf _
           & " Begin" & vbCrLf _
           & "  For R1 in C1 Loop" & vbCrLf _
           & "   Update POTSHIP1 Set " & vbCrLf _
           & "    WHSE_CODE = R1.WHSE_CODE" & vbCrLf _
           & "    where PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO;" & vbCrLf _
           & "  End Loop;" & vbCrLf _
           & " End;" & vbCrLf _
           & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & ", MIN(POTORDR1.ORDR_NO) ORDR_NO, MAX(POTORDR1.ORDR_NO) ORDR_NO2" & vbCrLf _
            & "   from POTSHIP3,POTSHIP1,POTORDR1" & vbCrLf _
            & "   where POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "     and POTSHIP3.PO_SHIPMENT_NO IN (" & Mid(sqlPO_SHIPMENT_NO, 2) & ")" & vbCrLf _
            & "     and POTORDR1.ORDR_NO is Not Null" & vbCrLf _
            & "   group by POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update POTSHIP2 Set " & vbCrLf _
            & "    ORDR_NO = R1.ORDR_NO" & vbCrLf _
            & "    where PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO" & vbCrLf _
            & "      and PO_SHIPMENT_LNO = R1.PO_SHIPMENT_LNO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE, POTSHIP3.*" & vbCrLf _
            & "   from POTSHIP3,POTORDR2,POTSHIP1" & vbCrLf _
            & "   where POTSHIP3.PO_SHIPMENT_NO IN (" & Mid(sqlPO_SHIPMENT_NO, 2) & ")" & vbCrLf _
            & "     and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "     and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "     and POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Begin" & vbCrLf _
            & "    Declare Cursor C2 is " & vbCrLf _
            & "     Select POTORDR2.* from POTORDR2" & vbCrLf _
            & "      where POTORDR2.PO_ORDER_NO = R1.PO_ORDER_NO" & vbCrLf _
            & "        and POTORDR2.PO_ORDER_LNO = R1.PO_ORDER_LNO" & vbCrLf _
            & "        for Update;" & vbCrLf _
            & "     PO_QTY_OPN_before NUMBER (8,0);" & vbCrLf _
            & "     PO_QTY_OPN_after NUMBER (8,0);" & vbCrLf _
            & "    Begin" & vbCrLf _
            & "     For R2 in C2 Loop" & vbCrLf _
            & "      PO_QTY_OPN_before := R2.PO_QTY_OPN;" & vbCrLf _
            & "      PO_QTY_OPN_after := GREATEST(0,NVL(R2.PO_QTY_OPN,0) - NVL(R1.PO_QTY_SHP,0));" & vbCrLf _
            & "      Update POTORDR2 Set" & vbCrLf _
            & "        PO_QTY_SHP = NVL(PO_QTY_SHP,0) + R1.PO_QTY_SHP" & vbCrLf _
            & "      , PO_QTY_OPN = PO_QTY_OPN_after" & vbCrLf _
            & "      , PO_STATUS = (CASE WHEN PO_QTY_OPN_after > 0 THEN 'O' ELSE 'C' END)" & vbCrLf _
            & "       where Current of C2;" & vbCrLf _
            & "      Update ICTSTAT2 Set " & vbCrLf _
            & "        WHSE_QTY_TRAN = NVL(WHSE_QTY_TRAN,0) + NVL(R1.PO_QTY_SHP,0)" & vbCrLf _
            & "      , WHSE_QTY_ON_ORDER = NVL(WHSE_QTY_ON_ORDER,0) - NVL(PO_QTY_OPN_before,0) + NVL(PO_QTY_OPN_after,0)" & vbCrLf _
            & "       where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & "      If SQL%NOTFOUND Then" & vbCrLf _
            & "       Insert into ICTSTAT2 (STYLE_CODE,COLOR_CODE,WHSE_CODE,WHSE_QTY_TRAN,WHSE_QTY_ON_ORDER)" & vbCrLf _
            & "        Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,R1.PO_QTY_SHP, -1 * NVL(PO_QTY_OPN_before,0) + NVL(PO_QTY_OPN_after,0));" & vbCrLf _
            & "      End If;" & vbCrLf _
            & "     End Loop;" & vbCrLf _
            & "    End;" & vbCrLf _
            & "   End;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select PO_ORDER_NO, Sum (Decode(PO_STATUS,'O',1,0)) O, Sum (Decode(PO_STATUS,'O',0,1)) C" & vbCrLf _
            & "   from POTORDR2 where PO_ORDER_NO in (" & vbCrLf _
            & "  Select Distinct PO_ORDER_NO from POTSHIP3" & vbCrLf _
            & "   where POTSHIP3.PO_SHIPMENT_NO IN (" & Mid(sqlPO_SHIPMENT_NO, 2) & ")" & vbCrLf _
            & "  ) group by PO_ORDER_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update POTORDR1 Set PO_STATUS = (Case When NVL(R1.O,0) > 0 Then 'O' Else 'C' End)" & vbCrLf _
            & "    where PO_ORDER_NO = R1.PO_ORDER_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
            & " UPDATE POTPPRM1 SET CENT_IMP_UPDATE_OPER = '" & ASCMAIN1.USER_ID & "',  CENT_IMP_UPDATE_DATE = SYSDATE WHERE POTPPRM1_CODE = 'Z' "
        ASCDATA1.ExecuteSQL()

        For Each PO_SHIPMENT_NO_AT_ONCE As String In PO_SHIPMENT_NOs
            TAC.POCMAIN1.Create_At_Once_Shipment(PO_SHIPMENT_NO_AT_ONCE)
        Next

        If do_not_commit Then '
            Rollback("There were errors reported during this Process")
        Else
            CommitTrans("Process Complete")
        End If

        Load_POTSHIPS()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub Ftp1_OnDirList(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.FtpDirListEventArgs) Handles FTP1.OnDirList
        If Not e.IsDir Then
            REMOTEDIRECTORYFILELIST.Add(e.FileName)
        End If
    End Sub

    Sub Process_File(FILENAME As String)

        ' BELOW LINE WAS FOR TESTING 
        ' My.Computer.FileSystem.CopyFile(ROWs("POTPARMC").Item("PO_PARM_DOWNLOAD_FOLDER") & Mid(FILENAME, 20), ROWs("POTPARMC").Item("PO_PARM_ARCH_FOLDER") & Mid(FILENAME, 20), True)

        Dim PO_CENT_NO As String = ASCMAIN1.Next_Control_No("POTCENT1.PO_CENT_NO")

        ' LOAD POTSHIPS
        Dim rowPOTCENT1 As DataRow = dst.Tables("POTCENT1").NewRow
        With rowPOTCENT1
            .Item("PO_CENT_NO") = PO_CENT_NO
            .Item("PO_CENT_STATUS") = "0"
            .Item("PO_CENT_FTP_DATE") = Now
            .Item("PO_CENT_FILE_NAME") = FILENAME
            .Item("INIT_DATE") = Now
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
        End With
        dst.Tables("POTCENT1").Rows.Add(rowPOTCENT1)
        Update_Record_TDA("POTCENT1")

        Dim data As String = ""
        Using sr As New System.IO.StreamReader(FILENAME)
            data = sr.ReadToEnd
            Dim datarec() As String = Split(data, vbCrLf)

            For i As Integer = 0 To UBound(datarec)
                If datarec(i) <> "" Then
                    Dim rowPOTCENT2 As DataRow = dst.Tables("POTCENT2").NewRow
                    With rowPOTCENT2
                        .Item("PO_CENT_NO") = PO_CENT_NO
                        .Item("PO_CENT_STATUS") = "0"

                        .Item("SHIP_KEY") = Trim(Mid(datarec(i), 1, 20) & "")
                        .Item("PO_NO") = Trim(Mid(datarec(i), 21, 25) & "")
                        .Item("ITEM_NO") = Trim(Mid(datarec(i), 46, 25) & "")
                        .Item("ITEM_COLOR") = Trim(Mid(datarec(i), 71, 25) & "")
                        .Item("ITEM_SIZE") = Trim(Mid(datarec(i), 96, 25) & "")
                        .Item("RANGE_KEY") = Trim(Mid(datarec(i), 121, 7) & "")
                        .Item("LF_VESSEL_NM") = Trim(Mid(datarec(i), 128, 50) & "")
                        .Item("LF_VESSEL_CD") = Trim(Mid(datarec(i), 178, 8) & "")
                        .Item("LF_VOYAGE_NO") = Trim(Mid(datarec(i), 186, 10) & "")
                        .Item("CLOSE_DT") = FormatDATE(Mid(datarec(i), 196, 10) & "")
                        .Item("LOAD_PORT_CD") = Trim(Mid(datarec(i), 206, 6) & "")
                        .Item("LOAD_PORT_NAME") = Trim(Mid(datarec(i), 212, 30) & "")
                        .Item("LF_EST_DEPART_DT") = FormatDATE(Mid(datarec(i), 242, 10) & "")
                        .Item("ORIGIN_FLAG") = Trim(Mid(datarec(i), 252, 2) & "")
                        .Item("CARRIER_CD") = Trim(Mid(datarec(i), 254, 6) & "")
                        .Item("LM_VESSEL_NM") = Trim(Mid(datarec(i), 260, 20) & "") ' SHORTENED TO 20 TO MATCH POTSHIP1
                        .Item("LM_VESSEL_CD") = Trim(Mid(datarec(i), 310, 8) & "")
                        .Item("LM_VOYAGE_NO") = Trim(Mid(datarec(i), 318, 10) & "")
                        .Item("LM_CLOSE_DT") = FormatDATE(Mid(datarec(i), 328, 10) & "")
                        .Item("LM_LOAD_PORT_CD") = Trim(Mid(datarec(i), 338, 6) & "")
                        .Item("LM_LOAD_PORT_NAME") = Trim(Mid(datarec(i), 344, 6) & "")
                        .Item("LM_EST_DEPART_DT") = FormatDATE(Mid(datarec(i), 374, 10) & "")
                        .Item("DISC_PORT_CD") = Trim(Mid(datarec(i), 384, 6) & "")
                        .Item("DISC_PORT_NAME") = Trim(Mid(datarec(i), 390, 30) & "")
                        .Item("EST_DISC_DT") = FormatDATE(Mid(datarec(i), 420, 10) & "")
                        .Item("ARRIVE_PORT_CD") = Trim(Mid(datarec(i), 430, 6) & "")
                        .Item("ARRIVE_LOCATION") = Trim(Mid(datarec(i), 436, 30) & "")
                        .Item("EST_ARRIVE_DT") = FormatDATE(Mid(datarec(i), 466, 10) & "")
                        .Item("MPCREF_NO") = Trim(Mid(datarec(i), 476, 7) & "")
                        .Item("CONTAINER_NO") = Trim(Mid(datarec(i), 483, 15) & "")
                        .Item("CONTAINER_SIZE") = Trim(Mid(datarec(i), 498, 3) & "")
                        .Item("FREIGHT_TYPE") = Trim(Mid(datarec(i), 501, 7) & "")
                        .Item("CONTAINER_RATED") = Trim(Mid(datarec(i), 508, 10) & "")
                        .Item("SEAL_NO") = Trim(Mid(datarec(i), 518, 15) & "")
                        .Item("TOTAL_FREIGHT_AT") = Trim(Mid(datarec(i), 533, 10) & "")
                        .Item("MASTERBL_NO") = Trim(Mid(datarec(i), 548, 15) & "")
                        .Item("HOUSEBL_NO") = Trim(Mid(datarec(i), 563, 20) & "")
                        .Item("FCR_NO") = Trim(Mid(datarec(i), 583, 15) & "")
                        .Item("COM_INVOICE_NO") = Trim(Mid(datarec(i), 598, 20) & "")
                        .Item("CUST_REF") = Trim(Mid(datarec(i), 618, 25) & "")
                        .Item("CONTRACT_NO") = Trim(Mid(datarec(i), 643, 20) & "")
                        .Item("AIRWAY_BILL") = Trim(Mid(datarec(i), 668, 20) & "")
                        .Item("DOCS_SENT_DT") = FormatDATE(Mid(datarec(i), 688, 10) & "")
                        .Item("CUST_AIRWAY_BILL") = Trim(Mid(datarec(i), 698, 20) & "")
                        .Item("CUST_DOCS_SENT_DT") = FormatDATE(Mid(datarec(i), 718, 10) & "")
                        .Item("GOODS_ORG") = Trim(Mid(datarec(i), 728, 2) & "")
                        .Item("GOODS_RCVD") = FormatDATE(Mid(datarec(i), 730, 10) & "")
                        .Item("DOCS_RCVD") = FormatDATE(Mid(datarec(i), 740, 10) & "")
                        .Item("ITEM_DESC") = Trim(Mid(datarec(i), 750, 50) & "")
                        .Item("VENDOR_CD") = Trim(Mid(datarec(i), 800, 20) & "")
                        .Item("VENDOR_NAME") = Trim(Mid(datarec(i), 820, 35) & "")
                        .Item("CARTONS") = Trim(Mid(datarec(i), 855, 7) & "")
                        .Item("CARTONS_UM") = Trim(Mid(datarec(i), 862, 2) & "")
                        .Item("UNITS") = Trim(Mid(datarec(i), 864, 7) & "")
                        .Item("UNITS_UM") = Trim(Mid(datarec(i), 871, 2) & "")
                        .Item("VOLUME") = Trim(Mid(datarec(i), 873, 11) & "")
                        .Item("VOLUME_UM") = Trim(Mid(datarec(i), 884, 2) & "")
                        .Item("WEIGHT") = Trim(Mid(datarec(i), 886, 12) & "")
                        .Item("WEIGHT_UM") = Trim(Mid(datarec(i), 898, 2) & "")
                        .Item("FOB_AT") = Trim(Mid(datarec(i), 900, 14) & "")
                        .Item("COMMODITY_CD") = Trim(Mid(datarec(i), 914, 5) & "")
                        .Item("STOW_POSITION") = Trim(Mid(datarec(i), 919, 10) & "")
                        .Item("INIT_DATE") = Now
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = Now
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    End With
                    dst.Tables("POTCENT2").Rows.Add(rowPOTCENT2)
                End If
            Next
            Update_Record_TDA("POTCENT2")

            ASCMAIN1.sql = "Delete from POTCENT2 where PO_CENT_NO = '" & PO_CENT_NO & "' and PO_NO = 'SAMPLE'"
            ASCDATA1.ExecuteSQL()
        End Using

        Assign_PO_SHIPMENT_NO()

    End Sub

    Sub Assign_PO_SHIPMENT_NO()

        'ASCMAIN1.sql = "" _
        '    & "Begin" & vbCrLf _
        '    & " Declare" & vbCrLf _
        '    & "  PO_SHIPMENT_NO_X VARCHAR2(6);" & vbCrLf _
        '    & "  Cursor C1 is  " & vbCrLf _
        '    & "   Select Distinct LM_VESSEL_NM, DECODE(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO) BOL" & vbCrLf _
        '    & "    from POTCENT2 where PO_SHIPMENT_NO is Null " & vbCrLf _
        '    & "     order by LM_VESSEL_NM, DECODE(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO); " & vbCrLf _
        '    & " Begin" & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   PO_SHIPMENT_NO_X :=  TAPCTLN1('POTSHIP1.PO_SHIPMENT_NO',1); " & vbCrLf _
        '    & "   Update POTCENT2 Set PO_SHIPMENT_NO = PO_SHIPMENT_NO_X " & vbCrLf _
        '    & "    where LM_VESSEL_NM = R1.LM_VESSEL_NM and DECODE(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO) = R1.BOL; " & vbCrLf _
        '    & "   Begin" & vbCrLf _
        '    & "    Declare" & vbCrLf _
        '    & "     PO_SHIPMENT_LNO_X NUMBER (6,0);" & vbCrLf _
        '    & "     Cursor C2 is" & vbCrLf _
        '    & "      Select Distinct POTCENT2.FCR_NO, POTCENT2.CONTAINER_NO" & vbCrLf _
        '    & "      , DECODE(POTORDR1.ORDR_NO,NULL,NULL,POTCENT2.PO_NO) PO_NO from POTCENT2,POTORDR1" & vbCrLf _
        '    & "       where POTCENT2.PO_SHIPMENT_NO = PO_SHIPMENT_NO_X" & vbCrLf _
        '    & "         and POTORDR1.PO_ORDER_NO = TRIM(POTCENT2.PO_NO)" & vbCrLf _
        '    & "       order by POTCENT2.FCR_NO, POTCENT2.CONTAINER_NO" & vbCrLf _
        '    & "      , DECODE(POTORDR1.ORDR_NO,NULL,NULL,POTCENT2.PO_NO); " & vbCrLf _
        '    & "    Begin" & vbCrLf _
        '    & "     PO_SHIPMENT_LNO_X := 0; " & vbCrLf _
        '    & "     For R2 in C2 Loop" & vbCrLf _
        '    & "      PO_SHIPMENT_LNO_X := PO_SHIPMENT_LNO_X + 1; " & vbCrLf _
        '    & "      Update POTCENT2 Set PO_SHIPMENT_LNO = PO_SHIPMENT_LNO_X" & vbCrLf _
        '    & "       where PO_SHIPMENT_NO = PO_SHIPMENT_NO_X" & vbCrLf _
        '    & "         and FCR_NO = R2.FCR_NO" & vbCrLf _
        '    & "         and CONTAINER_NO = R2.CONTAINER_NO" & vbCrLf _
        '    & "         and (R2.PO_NO is Null or PO_NO = R2.PO_NO);" & vbCrLf _
        '    & "      Update POTCENT2 Set PO_SHIPMENT_PLNO = ROWNUM" & vbCrLf _
        '    & "       where PO_SHIPMENT_NO = PO_SHIPMENT_NO_X and PO_SHIPMENT_LNO = PO_SHIPMENT_LNO_X; " & vbCrLf _
        '    & "     End Loop;" & vbCrLf _
        '    & "    End;" & vbCrLf _
        '    & "   End;" & vbCrLf _
        '    & "  End Loop;" & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End; " & vbCrLf


        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare" & vbCrLf _
            & "  PO_SHIPMENT_NO_X VARCHAR2(6);" & vbCrLf _
            & "  Cursor C1 is  " & vbCrLf _
            & "   Select Distinct LM_VESSEL_NM, DECODE(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO) BOL, POTORDR1.WHSE_CODE" & vbCrLf _
            & "    from POTCENT2,POTORDR1 where PO_SHIPMENT_NO is Null and POTORDR1.PO_ORDER_NO = TRIM(POTCENT2.PO_NO)" & vbCrLf _
            & "     order by LM_VESSEL_NM, DECODE(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO); " & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   PO_SHIPMENT_NO_X :=  TAPCTLN1('POTSHIP1.PO_SHIPMENT_NO',1); " & vbCrLf _
            & "   Update POTCENT2 Set PO_SHIPMENT_NO = PO_SHIPMENT_NO_X " & vbCrLf _
            & "    where LM_VESSEL_NM = R1.LM_VESSEL_NM and DECODE(TRIM(HOUSEBL_NO),NULL,MASTERBL_NO,HOUSEBL_NO) = R1.BOL" & vbCrLf _
            & "      and TRIM(PO_NO) = (Select PO_ORDER_NO from POTORDR1 where PO_ORDER_NO = TRIM(POTCENT2.PO_NO) and WHSE_CODE = R1.WHSE_CODE); " & vbCrLf _
            & "   Begin" & vbCrLf _
            & "    Declare" & vbCrLf _
            & "     PO_SHIPMENT_LNO_X NUMBER (6,0);" & vbCrLf _
            & "     Cursor C2 is" & vbCrLf _
            & "      Select Distinct POTCENT2.FCR_NO, POTCENT2.CONTAINER_NO" & vbCrLf _
            & "      , DECODE(POTORDR1.ORDR_NO,NULL,NULL,POTCENT2.PO_NO) PO_NO from POTCENT2,POTORDR1" & vbCrLf _
            & "       where POTCENT2.PO_SHIPMENT_NO = PO_SHIPMENT_NO_X" & vbCrLf _
            & "         and POTORDR1.PO_ORDER_NO = TRIM(POTCENT2.PO_NO)" & vbCrLf _
            & "       order by POTCENT2.FCR_NO, POTCENT2.CONTAINER_NO" & vbCrLf _
            & "      , DECODE(POTORDR1.ORDR_NO,NULL,NULL,POTCENT2.PO_NO); " & vbCrLf _
            & "    Begin" & vbCrLf _
            & "     PO_SHIPMENT_LNO_X := 0; " & vbCrLf _
            & "     For R2 in C2 Loop" & vbCrLf _
            & "      PO_SHIPMENT_LNO_X := PO_SHIPMENT_LNO_X + 1; " & vbCrLf _
            & "      Update POTCENT2 Set PO_SHIPMENT_LNO = PO_SHIPMENT_LNO_X" & vbCrLf _
            & "       where PO_SHIPMENT_NO = PO_SHIPMENT_NO_X" & vbCrLf _
            & "         and FCR_NO = R2.FCR_NO" & vbCrLf _
            & "         and CONTAINER_NO = R2.CONTAINER_NO" & vbCrLf _
            & "         and (R2.PO_NO is Null or PO_NO = R2.PO_NO) and NVL(PO_SHIPMENT_LNO,0) = 0;" & vbCrLf _
            & "      Update POTCENT2 Set PO_SHIPMENT_PLNO = ROWNUM" & vbCrLf _
            & "       where PO_SHIPMENT_NO = PO_SHIPMENT_NO_X and PO_SHIPMENT_LNO = PO_SHIPMENT_LNO_X; " & vbCrLf _
            & "     End Loop;" & vbCrLf _
            & "    End;" & vbCrLf _
            & "   End;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End; " & vbCrLf

        ASCDATA1.ExecuteSQL()

    End Sub


    Public Shared Function FormatDATE(ByVal PDATE As String) As Date
        If Trim(PDATE) <> "" Then
            FormatDATE = CDate((Mid(PDATE, 6, 2) & "/" & Mid(PDATE, 9, 2) & "/" & Mid(PDATE, 1, 4)))
        Else
            FormatDATE = Nothing
        End If


    End Function

End Class