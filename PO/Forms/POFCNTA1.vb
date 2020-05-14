Public Class POFCNTA1

    Dim POTCNTAX As String
    Dim CONTs As New List(Of String)
    Dim YPs As New List(Of String)
    Dim DT1 As String
    Dim DT2 As String
    Dim sqlPOTSHIPX As String
    Dim POTSHIPX As String
    Dim ORIG As String
    Dim DEST As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        With dst

            ASCMAIN1.sql = "Select * from POTCNTT1"
            Create_TDA(.Tables.Add, "POTCNTT1", "**", 0, False, "", 1)
            Fill_Records("POTCNTT1")

            Create_POTCNTAX()
            ASCMAIN1.sql = "Select * from " & POTCNTAX
            Create_TDA(.Tables.Add, "POTCNTAX", "**", 0, False, "", 0)


            ASCMAIN1.sql = "Select POTSHIP4.CONTAINER_TYPE_CODE" & vbCrLf _
                & ", COUNT (*) USED" & vbCrLf _
                & " from POTSHIP4,POTSHIP1" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP4.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP1.PORT_CODE_ORIG = ''" & vbCrLf _
                & "   and POTSHIP1.PORT_CODE_DEST = '' " & vbCrLf _
                & " group by POTSHIP4.CONTAINER_TYPE_CODE"
            Create_TDA(.Tables.Add, "POTCNTAY", "**", 0, False, "", 0)
            With .Tables("POTCNTAY")
                .Columns("USED").DataType = GetType(System.Int64)
                For i As Integer = 1 To 36
                    .Columns.Add("P" & Format(i, "00"), GetType(System.Int64))
                Next
            End With


            sqlPOTSHIPX = "SELECT POTSHIP1.*" & vbCrLf _
                & ", ICTWHSE1.LP_CODE, WHTLPXN1.INIT_OPER LP_XNO_INIT_OPER, WHTLPXN1.INIT_DATE LP_XNO_INIT_DATE" & vbCrLf _
                & " from POTSHIP1,WHTLPXN1,ICTWHSE1" & vbCrLf _
                & " where WHTLPXN1.LP_XNO (+) = POTSHIP1.LP_XNO" & vbCrLf _
                & "   and ICTWHSE1.WHSE_CODE (+) = POTSHIP1.WHSE_CODE"
            ASCMAIN1.sql = sqlPOTSHIPX & " and ROWNUM <1"
            POTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & POTSHIPX & " Add Primary Key (PO_SHIPMENT_NO)")

            ASCMAIN1.sql = "Select POTSHIPX.*" & vbCrLf _
                & ", X.CONTAINER_NO, X.BOL_NO, X.COMM_INV_NO, X.PO_SHIP_CTNS, X.PO_DATE_RECEIVED_MIN, X.PO_DATE_RECEIVED_MAX, X.ORDR_NO" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from " & POTSHIPX & " POTSHIPX, SOTORDR1" & vbCrLf _
                & ", (Select PO_SHIPMENT_NO" & vbCrLf _
                & ", Min (CONTAINER_NO) CONTAINER_NO" & vbCrLf _
                & ", Min (BOL_NO) BOL_NO" & vbCrLf _
                & ", Min (COMM_INV_NO) COMM_INV_NO" & vbCrLf _
                & ", Min (ORDR_NO) ORDR_NO" & vbCrLf _
                & ", Sum (PO_SHIP_CTNS) PO_SHIP_CTNS" & vbCrLf _
                & ", Min (PO_DATE_RECEIVED) PO_DATE_RECEIVED_MIN" & vbCrLf _
                & ", Max (PO_DATE_RECEIVED) PO_DATE_RECEIVED_MAX" & vbCrLf _
                & " from POTSHIP2 where PO_SHIPMENT_NO " & vbCrLf _
                & " in (Select PO_SHIPMENT_NO from " & POTSHIPX & ") group by PO_SHIPMENT_NO) X" & vbCrLf _
                & " where X.PO_SHIPMENT_NO (+) = POTSHIPX.PO_SHIPMENT_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO (+) = X.ORDR_NO"
            Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "", 1)
            .Tables("POTSHIPX").Columns.Add("LINES", GetType(System.Int64))
            .Tables("POTSHIPX").Columns.Add("LINES_REC", GetType(System.Int64))
            .Tables("POTSHIPX").Columns("PO_SHIP_CTNS").DataType = GetType(System.Int64)


        End With

        grdPOTCNTAX.DataSource = dst.Tables("POTCNTAX")
        For Each CONTAINER_TYPE_CODE As String In CONTs
            With grdPOTCNTAX.DisplayLayout.Bands(0).Columns("TYPE_" & CONTAINER_TYPE_CODE)
                If CONTAINER_TYPE_CODE = "" Then
                    .Header.Caption = "Unknown"
                Else
                    .Header.Caption = CONTAINER_TYPE_CODE
                End If
                .Width = 70
                .Format = "#,##0"
                Create_Summary(grdPOTCNTAX, "TYPE_" & CONTAINER_TYPE_CODE)
            End With
        Next
        With grdPOTCNTAX.DisplayLayout.Bands(0).Columns("USED")
            .Header.Caption = "Total"
            .Width = 70
            .Format = "#,##0"
            Create_Summary(grdPOTCNTAX, "USED")
        End With
        For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTCNTAX.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Drawing.Color.White
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            If gcol.Key = "USED" Then
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            ElseIf gcol.Key = "ORIG" Or gcol.Key = "DEST" Then
                If gcol.Key = "ORIG" Then gcol.Header.Caption = "Origin Port"
                If gcol.Key = "DEST" Then gcol.Header.Caption = "Destination Port"
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Width = 70
            Else
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            End If
        Next


        grdPOTSHIPX.DataSource = dst.Tables("POTSHIPX")
        Create_Summary(grdPOTSHIPX, "PO_SHIPMENT_NO", "Count")
        Create_Summary(grdPOTSHIPX, New String() {"LINES", "LINES_REC"})



        grdPOTSHIPX.DisplayLayout.UseFixedHeaders = True
        With grdPOTSHIPX.DisplayLayout.Bands(0)
            .Columns("PO_SHIPMENT_NO").Header.Fixed = True
            .Columns("PO_SHIP_VESSEL").Header.Fixed = True
            .Columns("PO_SHIP_ETA").Header.Fixed = True
        End With
        'grdPOTSHIPX.DisplayLayout.GroupByBox.Hidden = False


        With grdPOTSHIPX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_SHIP_VESSEL", "PO_SHIP_ETA", "PO_SHIP_REF_NO", "COST_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"PO_DATE_SHIPPED", "PORT_CODE", "WHSE_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Fuchsia
                ElseIf New String() {"LINES", "LINES_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"LP_STATUS", "LP_XNO", "LP_XNO_INIT_DATE", "LP_XNO_INIT_OPER", "LP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                    If gcol.Key <> "LP_CODE" Then gcol.CellAppearance.ForeColor = Drawing.Color.Green
                ElseIf New String() {"INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE", "PO_SHIPMENT_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With
        ASCMAIN1.Add_Value_List(grdPOTSHIPX, "LP_STATUS", Nothing, New String() {":", "1:Transmitted", "0:Not Transmitted"})


        grdPOTCNTAY.DataSource = dst.Tables("POTCNTAY")

        Create_Summary(grdPOTCNTAY, "USED")
        For i As Integer = 1 To 36
            Create_Summary(grdPOTCNTAY, "P" & Format(i, "00"))
        Next
        For Each gcol As UltraWinGrid.UltraGridColumn In grdPOTCNTAY.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Drawing.Color.White
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            If gcol.Key = "USED" Then
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            ElseIf gcol.Key = "CONTAINER_TYPE_CODE" Then
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Width = 70
            Else
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            End If
        Next

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
                Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

                Dim N As Integer = ASCMAIN1.Period_Diff(RYP0, RYP1)

                If N < 0 Or N > 35 Then
                    EMsg &= vbCr & "Invalid Period Range (must be from 1 to 36 months)"
                Else
                    YPs.Clear()
                    For i As Integer = 0 To N
                        Dim YP As String = ASCMAIN1.Period_Calc(RYP0, i)
                        YPs.Add(YP)
                        With grdPOTCNTAY.DisplayLayout.Bands(0).Columns("P" & Format(i + 1, "00"))
                            .Header.Caption = ASCMAIN1.Get_Legend(YP, False, True)
                            .Width = 60
                            .Format = "#,##0"
                            .Hidden = False
                        End With
                    Next
                    For i As Integer = N + 1 To 35
                        With grdPOTCNTAY.DisplayLayout.Bands(0).Columns("P" & Format(i + 1, "00"))
                            .Hidden = True
                        End With
                    Next
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

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

            Case "Print"
                Call Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                .Groups("Period Range").Enabled = Not ScreenMode
            End With
        End If

        spl.Panel1Collapsed = True

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        splPOTCNTAX.Visible = ScreenMode
        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POTCNTAX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Reading from Shipments Data")

        Save_Header_Fields(UltraGroupBox1)


        Create_POTCNTAX()

        EnforceConstraints(False)

        Fill_Records("POTCNTAX")
        Sort_grdColumns(grdPOTCNTAX, "ORIG,DEST")

        EnforceConstraints(True)


        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTCNTAX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdPOTSHIPX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Shipment Inquiry")
        '  Load_Popup_Menu(grdPOTCNTAX, "SSB", "Show Filter", "Show GroupBox", "PO Shipments Inquiry")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name
            Case "grdSATSLSC1 "
            Case Else
        End Select
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "PO Sipments Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Sub Create_POTCNTAX()

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value


        Dim DTE1 As Date = Now.Date
        Dim DTE2 As Date = Now.Date

        If RYP0 <> "" Then
            Dim RYP0_dates() As Date = ASCMAIN1.Get_Dates(RYP0)
            Dim RYP1_dates() As Date = ASCMAIN1.Get_Dates(RYP1)

            DTE1 = RYP0_dates(1)
            DTE2 = RYP1_dates(RYP1_dates.Length - 1)
        End If

        DT1 = Format(DTE1, "dd-MMM-yyyy")
        DT2 = Format(DTE2, "dd-MMM-yyyy")

        grdPOTCNTAX.Text = "Lane Analysis for " & optDate.Text & " range from " & Format(DTE1, "MM/dd/yy") & " to " & Format(DTE2, "MM/dd/yy")

        Dim sql As String = ", Sum (Decode (TYPE,NULL,USED, 0)) TYPE_" & vbCrLf
        CONTs.Clear()
        CONTs.Add("")
        For Each row As DataRow In dst.Tables("POTCNTT1").Select("", "CONTAINER_TYPE_CODE")
            Dim CONTAINER_TYPE_CODE As String = row.Item("CONTAINER_TYPE_CODE")
            CONTs.Add(CONTAINER_TYPE_CODE)
            sql &= ", Sum (Decode(TYPE, '" & CONTAINER_TYPE_CODE & "', USED, 0)) TYPE_" & CONTAINER_TYPE_CODE & vbCrLf
        Next

        ASCMAIN1.sql = "Select ORIG, DEST, Sum (USED) USED" & sql _
            & " from (" & vbCrLf _
            & "Select POTSHIP1.PORT_CODE_ORIG ORIG, POTSHIP1.PORT_CODE_DEST DEST, POTSHIP4.CONTAINER_TYPE_CODE TYPE, COUNT (*) USED" & vbCrLf _
            & " from POTSHIP4,POTSHIP1" & vbCrLf _
            & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP4.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP1." & optDate.Value & " between '" & DT1 & "' and " & "'" & DT2 & "'" & vbCrLf _
            & " group by POTSHIP1.PORT_CODE_ORIG, POTSHIP1.PORT_CODE_DEST, POTSHIP4.CONTAINER_TYPE_CODE" & vbCrLf _
            & ") group by ORIG, DEST"

        If POTCNTAX = "" Then
            POTCNTAX = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & POTCNTAX)
            ASCDATA1.ExecuteSQL("Insert into " & POTCNTAX & " " & ASCMAIN1.sql)
        End If
    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()
        Print_Report_End()
    End Sub

    Private Sub grdPOTCNTAX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTCNTAX.AfterRowActivate
        Setup_grdPOTCNTAX()
    End Sub

    Sub Setup_grdPOTCNTAX()
        If grdPOTCNTAX.ActiveRow Is Nothing OrElse Not grdPOTCNTAX.ActiveRow.IsDataRow Then
            tabDetails.Visible = False
        Else
            ASCMAIN1.sql = "Select POTSHIP4.CONTAINER_TYPE_CODE" & vbCrLf _
                & ", COUNT (*) USED" & vbCrLf

            Dim i As Integer = 0
            For Each YP As String In YPs
                i += 1
                ASCMAIN1.sql &= ", SUM (DECODE(TO_CHAR(PO_DATE_SHIPPED,'YYYYMM'),'" & YP & "',1,0)) P" & Format(i, "00") & vbCrLf
            Next

            ORIG = grdPOTCNTAX.ActiveRow.Cells("ORIG").Value & ""
            DEST = grdPOTCNTAX.ActiveRow.Cells("DEST").Value & ""

            ASCMAIN1.sql &= " from POTSHIP4,POTSHIP1" & vbCrLf _
            & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP4.PO_SHIPMENT_NO" & vbCrLf _
            & "   and NVL(POTSHIP1.PORT_CODE_ORIG,'') = '" & ORIG & "'" & vbCrLf _
            & "   and NVL(POTSHIP1.PORT_CODE_DEST,'') = '" & DEST & "' " & vbCrLf _
            & "   and POTSHIP1." & optDate.Value & " between '" & DT1 & "' and " & "'" & DT2 & "'" & vbCrLf _
            & " group by POTSHIP4.CONTAINER_TYPE_CODE"

            Fill_Records("POTCNTAY", "", True, ASCMAIN1.sql)

            Sort_grdColumns(grdPOTCNTAY, "CONTAINER_TYPE_CODE")

            Load_POTSHIPX()

            tabDetails.Visible = True

        End If
    End Sub

    Sub Load_POTSHIPX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ASCMAIN1.sql = sqlPOTSHIPX
        ASCMAIN1.sql &= " and NVL(POTSHIP1.PORT_CODE_ORIG,'') = '" & ORIG & "' and NVL(POTSHIP1.PORT_CODE_DEST,'') = '" & DEST & "'"

        grdPOTSHIPX.Text = "Shipments in Lane " & ORIG & "-" & DEST

        ASCDATA1.ExecuteSQL("Delete from " & POTSHIPX)

        ASCMAIN1.sql = "Insert into " & POTSHIPX & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        Fill_Records("POTSHIPX")
        Sort_grdColumns(grdPOTSHIPX, "PO_SHIP_ETA".ToLower)
 
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
 
End Class