Public Class EDFASNO1
    Dim CUST_CODE As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CUST_CODE = "WALMART"
        Absx1.txtFor("CUST_CODE").Value = CUST_CODE
        Get_PARM("EDTPARM1")
        With dst
            ASCMAIN1.sql = "Select SHIP_BOL_NO, MASTER_BILL_OF_LADING_NO, SOTSHIP1.BILL_OF_LADING_NO, EDI_LOAD_ID,
            SHIP_REF, SHIP_856_BATCH_NO, SHIP_DATE_SHIPPED, SHIP_STATUS, ORDR_CUST_PO from 
            SOTSHIP1, SOTORDR0
            where SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO
            And CUST_CODE = 'WALMART' 
            And SHIP_DATE_SHIPPED between :PARM1 And : PARM2"
            Create_TDA(.Tables.Add, "EDTASNO1", "**", 0, False, "DD")

            ASCMAIN1.sql = "SELECT (NEW_TIME((TO_DATE('01/01/1970', 'MM-DD-YYYY') + ""TimeCreated"" / 86400), 'GMT', 'EST')) Created, DOC.""Direction"" Direction, DOC.""PartnerKEY"", DOC.""DocumentName"",
                DOC.""TransactionSetID"", DOC.""FunctionalGroupID"", DOC.""ControlNumber"", DOC.""ComplianceStatus"" ComplianceStatus, DOC.""DocumentBlobKEY"", DOC.""DocumentKEY"" FROM ""Document_tb"" DOC, ""Partner_tb"" PARTNER " &
               "WHERE ""TransactionSetID"" = '856' " &
               "AND TO_CHAR(NEW_TIME(TO_DATE('01/01/1970', 'MM-DD-YYYY') + " &
               """TimeCreated""/ 86400, 'GMT', 'EST'), 'yyyymmdd') BETWEEN :PARM1 AND :PARM2 " &
               "AND PARTNER.""PartnerKEY"" = DOC.""PartnerKEY"" AND PARTNER.""PartnerKEY"" = 'WMartTest'"
            Create_TDA(.Tables.Add, "EDTASNO2", "**", 0, False, "DD")

        End With

        grdEDTASNO1.DataSource = dst.Tables("EDTASNO1")
        grdEDTASNO2.DataSource = dst.Tables("EDTASNO2")
        With grdEDTASNO2
            If .DisplayLayout.Bands(0).Columns.Exists("ComplianceStatus") Then
                .DisplayLayout.Bands(0).Columns("ComplianceStatus").Header.Appearance.ForeColor = Drawing.Color.White
                .DisplayLayout.Bands(0).Columns("ComplianceStatus").Header.Appearance.BackColor = Drawing.Color.Blue
            End If
        End With


        'dteSearchE.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        'dteSearchE.MinDate = DateAdd(DateInterval.Year, -5, DateTime.Now)
        dteSearchE.DateTime = DateTime.Now

        'dteSearchS.MaxDate = dteSearchE.MaxDate
        'dteSearchS.MinDate = dteSearchE.MinDate
        dteSearchS.DateTime = DateAdd(DateInterval.Month, -1, DateTime.Now)

        ASCMAIN1.Add_Value_List(grdEDTASNO2, "Direction", , New String() {":", "0:Inbound", "1:Outbound"})
        ASCMAIN1.Add_Value_List(grdEDTASNO2, "ComplianceStatus", , New String() {":", "0:Incomplete", "7:Ack'd", "8:Waiting", "9:OverDue", "10:NetWarning"})

        Show_Filter(grdEDTASNO1)
        Show_Filter(grdEDTASNO2)
        'ASCMAIN1.Add_Value_List(grdEDTASNO2, "CYCLE_TYPE", , New String() {":", "C:COUNT", "V:VERIFY"})


        '0:Incomplete
        '1:NonCompliant
        '2:OK
        '3:DocQueued
        '4:Sent
        '5:NetReceived
        '6:NetDelivered
        '7:Ack'd
        '8:Waiting
        '9:OverDue
        '10:NetWarning
        '11:NetError
        '12:AckErr
        '13:FAPartial
        '14:FAReject
        '15:NetPickedUp
        '16:Duplicate
        '17:ReadyToSend

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Cancel"

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
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdEDTASNO1.Visible = ScreenMode
        grdEDTASNO2.Visible = ScreenMode
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"EDTASNO1", "EDTASNO2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Try
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"EDTASNO1", "EDTASNO2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            'need to parameterize
            ASCMAIN1.sql = $"Select SHIP_BOL_NO, MASTER_BILL_OF_LADING_NO, SOTSHIP1.BILL_OF_LADING_NO, EDI_LOAD_ID,
            SHIP_REF, SHIP_856_BATCH_NO, SHIP_DATE_SHIPPED, SHIP_STATUS, ORDR_CUST_PO from 
            SOTSHIP1, SOTORDR0
            where SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO
            And CUST_CODE = 'WALMART' 
            And SHIP_DATE_SHIPPED between '{dteSearchS.DateTime.ToString("dd-MMM-yyyy")}' And '{dteSearchE.DateTime.ToString("dd-MMM-yyyy")}'"

            Fill_Records("EDTASNO1", "", True, ASCMAIN1.sql)
            'Trunc(NEW_TIME((TO_DATE('01/01/1970', 'MM-DD-YYYY') + "TimeCreated" / 86400), 'GMT', 'EST')),
            Sort_grdColumns(grdEDTASNO1, "ORDR_CUST_PO")
            ASCMAIN1.sql = $"SELECT (NEW_TIME((TO_DATE('01/01/1970', 'MM-DD-YYYY') + ""TimeCreated"" / 86400), 'GMT', 'EST')) Created, DOC.""Direction"" Direction, DOC.""PartnerKEY"", DOC.""DocumentName"",
                DOC.""TransactionSetID"", DOC.""FunctionalGroupID"", DOC.""ControlNumber"", DOC.""ComplianceStatus"" ComplianceStatus, DOC.""DocumentBlobKEY"", DOC.""DocumentKEY"" FROM ""Document_tb"" DOC, ""Partner_tb"" PARTNER " &
               "WHERE ""TransactionSetID"" = '856' " &
               "AND TO_CHAR(NEW_TIME(TO_DATE('01/01/1970', 'MM-DD-YYYY') + " &
               $"""TimeCreated""/ 86400, 'GMT', 'EST'), 'yyyymmdd') BETWEEN 
               '{dteSearchS.DateTime.ToString("yyyyMMdd")}' and '{dteSearchE.DateTime.ToString("yyyyMMdd")}' " &
               "And PARTNER.""PartnerKEY"" = DOC.""PartnerKEY"" And PARTNER.""PartnerKEY"" = 'WMartTest'"

            Fill_Records("EDTASNO2", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdEDTASNO2, "compliancestatus, CREATED")
            'Catch ex As Exception
            '    MessageBox.Show($"Error Generating EDI 846 {ex.Message }", "Generate EDI 846", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Clear_Record()
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
    ByVal ctl As Control,
    ByVal COLUMN_NAME As String,
    Optional ByRef sql_where As String = "",
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "SEASON_CODE"
            '    If Absx1.optFor("STMT_TYPE").CheckedIndex <> -1 Then
            '        sql_where = "STMT_TYPE = '" & Absx1.optFor("STMT_TYPE").Value & "'"
            '    End If
        End Select
    End Sub


#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDTASNO1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show EDI ASN")
        Load_Popup_Menu(grdEDTASNO2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show EDI ASN")
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
                'SHIP_856_BATCH_NO
                Case "grdEDTASNO1"
                    tlb_pop.Tools("Show EDI ASN").SharedProps.Visible = grd.ActiveRow.Cells("SHIP_856_BATCH_NO").Value & "" <> ""
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

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("EDI_STYLE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            'Case "Show Raw EDI"

            '    'If grdSOTORDR1.ActiveRow IsNot Nothing Then
            '    '    Dim EDI_DOC_SEQ_NO As String = grdSOTORDR1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
            '    'End If
            '    '  Display_Raw(grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value & "")

            '    'grdEDTASNO1
            '    If grd.Name = "grdEDTASNO2" And grdEDTASNO2.ActiveRow IsNot Nothing Then
            '        Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
            '        Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"))
            '        Using frm As New ASFTEXT1
            '            frm.t = RAW_EDI
            '            frm.Text = "Raw EDI for " & CUST_CODE & " PO No " & grdEDTASNO1.ActiveRow.Cells("ORDR_CUST_PO").Value
            '            frm.ShowDialog()
            '        End Using
            '    End If

            Case "Show EDI ASN"
                If grd.ActiveRow IsNot Nothing Then
                    If grd.Name = "grdEDTASNO2" And grdEDTASNO2.ActiveRow IsNot Nothing Then
                        Dim EDI_DOCUMENT_NAME As String = grd.ActiveRow.Cells("DocumentName").Value & ""
                        Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI("", ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), "856", EDI_DOCUMENT_NAME)
                        Using frm As New ASFTEXT1
                            frm.t = RAW_EDI
                            frm.Text = "Raw EDI for " & CUST_CODE & " BOL No " & EDI_DOCUMENT_NAME
                            frm.ShowDialog()
                        End Using
                    ElseIf grd.Name = "grdEDTASNO1" And grdEDTASNO1.ActiveRow IsNot Nothing Then
                        Dim EDI_DOCUMENT_NAME As String = grd.ActiveRow.Cells("BILL_OF_LADING_NO").Value & ""
                        Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI("", ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), "856", EDI_DOCUMENT_NAME)
                        Using frm As New ASFTEXT1
                            frm.t = RAW_EDI
                            frm.Text = "Raw EDI for " & CUST_CODE & " BOL No " & EDI_DOCUMENT_NAME
                            frm.ShowDialog()
                        End Using
                    End If

                End If

        End Select
    End Sub

#End Region




End Class