Imports Infragistics.Win.UltraWinGrid

Public Class SOFPICKU
    Public SOTORDQ1 As String
    Public GROUP_KEY As String
    Public TOTEs_in_SLOTs As List(Of String)
    Public update_flag As Boolean = False
    Public TRUCK_NO As String
    Public PICK_DESCRIPTION As String
    Public tbl As DataTable
    Public rowSOTORDQ0 As DataRow
    Dim verifyToteSlot As Boolean = False
    'Dim NPIX_NO As String = ""
    Public TRUCK_TYPE As String = ""
    Dim SLOT_NO_to_Verify As Integer = 0
    Public SALES_DIVISION_CODE_DC As String = ""
    Public DC_CODE As String = String.Empty

    Dim Appearance_Magenta As New Infragistics.Win.Appearance
    Dim Appearance_Yellow As New Infragistics.Win.Appearance
    Dim Appearance_Empty As New Infragistics.Win.Appearance


    Public Sub New(ByVal FF As ASFBASE1, ByVal inDC_CODE As String)
        frmASFBASE1 = FF
        InitializeComponent()
        DC_CODE = inDC_CODE & String.Empty

        DC_CODE = DC_CODE.Trim
        If DC_CODE.Length = 0 Then
            DC_CODE = "???"
        End If
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Appearance_Magenta.BackColor = Drawing.Color.Magenta
        Appearance_Yellow.BackColor = Drawing.Color.Yellow

        With dst
            ASCMAIN1.sql = $"Select SOTORDQ1.* from {SOTORDQ1} SOTORDQ1"
            Create_TDA(.Tables.Add, "SOTORDQ1", "**", 0, False)
            If Not dst.Tables("SOTORDQ1").Columns.Contains("ORDR_QTY_ALLO") Then
                .Tables("SOTORDQ1").Columns.Add("ORDR_QTY_ALLO", GetType(System.Int32))
            End If

            With .Tables("SOTORDQ1")
                .Columns("ORDR_CNT").DataType = GetType(System.Int32)
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                .Columns.Add("TOTE_NO")
                .Columns.Add("SLOT_NO", GetType(System.Int32))
                .Columns.Add("TOTE_CLASS_CODE")
                .Columns.Add("TOTE_CLASS_MIN_QTY", GetType(System.Int32))
                .Columns.Add("TOTE_CLASS_MAX_QTY", GetType(System.Int32))
                .Columns.Add("VERIFIED")
                '.Columns("VERIFIED").DefaultValue = "0"
            End With

            With .Tables.Add("SOTSCAN1")
                .Columns.Add("SCAN_NO", GetType(System.Int32))
                .Columns.Add("SCAN")
                .Columns.Add("RESULT")
                .Columns.Add("ERR")
            End With

            ASCMAIN1.sql = "Select * from SOTTOTE1 where TRUCK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTTOTET", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from SOTTOTE0"
            Create_TDA(.Tables.Add, "SOTTOTE0", "**", 0, False)
            Fill_Records("SOTTOTE0")

            ASCMAIN1.sql = "Select * from SOTTRCK1 where DC_CODE = :PARM1 AND TRUCK_TYPE = 'P' AND PICK_BATCH_NO IS NULL"
            Create_TDA(.Tables.Add, "SOTTRCKT", "**", 0, False, "V", 1)
            Fill_Records("SOTTRCKT", DC_CODE)
            Dim lstTRUCK_NOs As New List(Of String)
            For Each rowSOTTRCKT As DataRow In dst.Tables("SOTTRCKT").Select("", "TRUCK_NO")
                lstTRUCK_NOs.Add(rowSOTTRCKT.Item("TRUCK_NO"))
            Next
            lblInstruction1.Text = String.Join(", ", lstTRUCK_NOs.ToArray)
        End With

        Dim rowQ1s() As DataRow = frmASFBASE1.dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' and SEL = '1'")
        For Each row As DataRow In rowQ1s
            dst.Tables("SOTORDQ1").Rows.Add(row.ItemArray)
        Next
        grdSOTORDQ1.DataSource = dst.Tables("SOTORDQ1")

        grdSOTSCAN1.DataSource = dst.Tables("SOTSCAN1")


        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTORDQ1, grdSOTSCAN1}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns

                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next
            End With
        Next


        Create_Summary(grdSOTORDQ1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDQ1, New String() {"ORDR_QTY_OPEN", "ORDR_QTY_BACK", "ORDR_QTY_ALLO"})

        grdSOTORDQ1.DisplayLayout.Bands(0).Columns("ORDR_QTY_ALLO").Hidden = (SALES_DIVISION_CODE_DC <> "SKIN")
        grdSOTORDQ1.DisplayLayout.Bands(0).Columns("ORDR_QTY_ALLO").Hidden = (SALES_DIVISION_CODE_DC <> "SKIN")

        cmdUpdate.Enabled = False
        chkCustomTruck.Visible = True
        'Set_Read_Only_for_ctl(chkCustomTruck, True)

        Dim rowSOTORDQ0 As DataRow = frmASFBASE1.dst.Tables("SOTORDQ0").Rows.Find(GROUP_KEY)

        Absx1.txtFor("TRUCK_NO").ReadOnly = True

        If SALES_DIVISION_CODE_DC = "SKIN" Then
            chkCustomTruck.Checked = False
            chkCustomTruck.Visible = True
        End If

        PICK_DESCRIPTION = rowSOTORDQ0.Item("PICK_DESCRIPTION") & ""

        Me.Text &= ": " & PICK_DESCRIPTION
        Me.Width = cmdCancel.Left + cmdCancel.Width + 25
        CenterToParent()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        If dst.Tables("SOTORDQ1").Select("TOTE_NO IS NULL").Length <> 0 Then
            MsgBox("Some Orders have Not been Assigned to Totes", MsgBoxStyle.OkOnly, "Cannot Close this Batch")
            Exit Sub
        End If

        tbl = dst.Tables("SOTORDQ1")
        update_flag = True

        Me.Close()
    End Sub

    Private Sub btnSimulateScan_Click(sender As Object, e As EventArgs) Handles btnSimulateScan.Click
        'Dim scans() As String = {"TBAD", "T202", "T101", "T102", "T104", "T103", "000001", "000001", "000004", "000006", "000004X", "000005", "000006", "000007", "000008", "000009"}

        If txtSCAN.Text <> "" Then
            Process_Scan()
        Else

            'Dim S As Integer = Val(txtSCAN.Tag & "")
            'txtSCAN.Text = scans(S)

            'Process_Scan()

            'S += 1
            'txtSCAN.Tag = CStr(S)
        End If

    End Sub

    Private Sub txtSCAN_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSCAN.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Process_Scan()
        End If
    End Sub

    Sub Process_Scan()
        Dim SCAN As String = txtSCAN.Text.ToUpper
        Dim RESULT As String = ""
        Dim ERR As String = "1"

        If txtTRUCK.Text = "" Then
            Dim rowSOTTRCK1 As DataRow = LookUp("SOTTRCK1", SCAN)
            If rowSOTTRCK1 Is Nothing Then
                RESULT = "Invalid Value for Truck"
            ElseIf rowSOTTRCK1.Item("TRUCK_TYPE") & "" = "X" And Not chkCustomTruck.checked Then
                RESULT = $"Truck {SCAN} is a Custom Truck only for NPI"
            Else
                TRUCK_NO = SCAN
                Dim DC_CODE_truck As String = rowSOTTRCK1.Item("DC_CODE") & ""
                If DC_CODE_truck <> DC_CODE Then
                    RESULT = $"Truck {TRUCK_NO} is in DC {DC_CODE_truck}"
                Else
                    Dim PICK_BATCH_NO As String = rowSOTTRCK1.Item("PICK_BATCH_NO") & ""
                    If PICK_BATCH_NO <> "" Then
                        RESULT = $"Truck {TRUCK_NO} is in use in Pick Batch {PICK_BATCH_NO}"
                    Else
                        TRUCK_TYPE = rowSOTTRCK1.Item("TRUCK_TYPE") & ""

                        If TRUCK_TYPE = "P" Then
                            Fill_Records("SOTTOTET", TRUCK_NO)
                            Dim TOTE_COUNT As Integer = dst.Tables("SOTTOTET").Rows.Count
                            If TOTE_COUNT = 0 Then
                                RESULT = $"Truck {TRUCK_NO} is Pre-Configured but Tote Configuration shows no Totes"
                            ElseIf TOTE_COUNT < dst.Tables("SOTORDQ1").Rows.Count Then
                                RESULT = $"Truck {TRUCK_NO} has {CStr(TOTE_COUNT)} Totes - not enough for this Pick Batch"
                                'ElseIf NPIX_NO <> "" Then
                                '    RESULT = $"Truck {TRUCK_NO} is Pre-Configured - not valid for NPI orders"
                            ElseIf chkCustomTruck.Checked Then
                                RESULT = $"Truck {TRUCK_NO} is Pre-Configured - not valid for Custom Truck"
                            Else
                                Dim SLOT_NO_ctr As Integer = 0
                                For Each row As DataRow In dst.Tables("SOTTOTET").Select("", "SLOT_NO")
                                    Dim PICK_NO As String = row.Item("PICK_NO") & ""
                                    Dim TOTE_NO As String = row.Item("TOTE_NO") & ""
                                    If PICK_NO <> "" Then
                                        RESULT = $"Truck {TRUCK_NO} is Pre-Configured but Tote {TOTE_NO} is associated with a Pick Ticket"
                                        Exit For
                                    Else
                                        SLOT_NO_ctr += 1
                                        If Val(row.Item("SLOT_NO") & "") <> SLOT_NO_ctr Then
                                            RESULT = $"Truck {TRUCK_NO} is Pre-Configured but Slots are not contiguous"
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If

                        If RESULT = "" Then
                            RESULT = $"Truck {TRUCK_NO} is Pre-Configured but Tote Configuration shows no Totes"

                            If Not ASCMAIN1.Logical_Lock("SOTTRCK1", TRUCK_NO,, False, False, 1) Then
                                RESULT = $"Cannot Lock Truck {TRUCK_NO}"
                            Else
                                RESULT = $"Truck {TRUCK_NO} Accepted"
                                txtTRUCK.Text = TRUCK_NO
                                TOTEs_in_SLOTs = New List(Of String)
                                lblInstruction.Text = "Scan the Tote in Slot 1"

                                lblPreConfigured.Visible = (TRUCK_TYPE = "P")
                                Dim TOTE_COUNT As Integer = dst.Tables("SOTTOTET").Rows.Count
                                lblPreConfigured.Text = $"Pre-Configured with {CStr(TOTE_COUNT)} Totes"

                                '  chkSaveTruck.Visible = (TRUCK_PRECONFIG <> "1" And NPIX_NO = "")

                                ERR = "0"

                                If TRUCK_TYPE = "P" Then
                                    chkCustomTruck.Checked = False
                                End If

                                If chkCustomTruck.Checked Then
                                    TRUCK_TYPE = "X"

                                    dst.Tables("SOTTOTET").Rows.Clear()

                                    For SLOT_NO As Integer = 1 To dst.Tables("SOTORDQ1").Rows.Count
                                        Dim TOTE_NO As String = ASCMAIN1.Next_Control_No("TOTE_NO_CUSTOM")
                                        TOTE_NO = "X" & Mid(TOTE_NO, TOTE_NO.Length - 4, 5)
                                        Dim rowSOTTOTET As DataRow = dst.Tables("SOTTOTET").NewRow
                                        With rowSOTTOTET
                                            .Item("TOTE_NO") = TOTE_NO
                                            .Item("TOTE_CLASS_CODE") = "X"
                                            .Item("DC_CODE") = rowSOTTRCK1.Item("DC_CODE")
                                            .Item("TRUCK_NO") = TRUCK_NO
                                            .Item("SLOT_NO") = SLOT_NO
                                            .Item("TOTE_TYPE") = "X"
                                        End With
                                        dst.Tables("SOTTOTET").Rows.Add(rowSOTTOTET)
                                    Next

                                    'BeginTrans()
                                    'Update_Record_TDA("SOTTOTE1")
                                    'ASCMAIN1.sql = "Update SOTTRCK1 Set TRUCK_CUSTOM = '1' where TRUCK_NO = :PARM1"
                                    'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", TRUCK_NO)
                                    'CommitTrans()

                                    Set_Read_Only(chkCustomTruck, True)
                                Else
                                    chkCustomTruck.Visible = False
                                End If


                                If TRUCK_TYPE = "P" Or TRUCK_TYPE = "X" Then
                                    Record_Scan(SCAN, RESULT, ERR)
                                    For SLOT_NO As Integer = 1 To dst.Tables("SOTORDQ1").Rows.Count
                                        Dim rowSOTTOTET As DataRow = dst.Tables("SOTTOTET").Select($"SLOT_NO = {CStr(SLOT_NO)}")(0)
                                        Dim TOTE_NO As String = rowSOTTOTET.Item("TOTE_NO")
                                        txtSCAN.Text = TOTE_NO
                                        Process_Scan()
                                    Next
                                    RESULT = ""
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Else

            If verifyToteSlot Then


                Dim row As DataRow = dst.Tables("SOTORDQ1").Select($"SLOT_NO = {CStr(SLOT_NO_to_Verify)}")(0)
                Dim TOTE_NO_expected As String = row.Item("TOTE_NO") & ""
                If SCAN <> TOTE_NO_expected Then
                    RESULT = $"Invalid Scan {SCAN} - expected Tote {TOTE_NO_expected} in slot {CStr(SLOT_NO_to_Verify)}"
                Else
                    RESULT = $"Verified Tote {SCAN} in slot {CStr(SLOT_NO_to_Verify)}"
                    ERR = "0"
                    row.Item("VERIFIED") = "1"
                    Dim rowsLeft() As DataRow = dst.Tables("SOTORDQ1").Select("VERIFIED = '0'", "SLOT_NO")
                    If rowsLeft.Length = 0 Then
                        btnSimulateScan.Enabled = False
                        cmdUpdate.Enabled = True
                        lblTruckIsReady.Visible = True
                        '  chkSaveTruck.Appearance.ForeColor = System.Drawing.Color.Red
                        '  chkSaveTruck.Appearance.FontData.Bold = DefaultableBoolean.True
                        lblInstruction2.Visible = False
                        verifyToteSlot = False
                    Else
                        SLOT_NO_to_Verify = Val(rowsLeft(0).Item("SLOT_NO"))
                        lblInstruction2.Text = $"Scan Tote in Slot {CStr(SLOT_NO_to_Verify)}"
                    End If
                End If

            Else

                Dim rowSOTTOTE1 As DataRow = Nothing
                If TRUCK_TYPE = "P" Or TRUCK_TYPE = "X" Then
                    rowSOTTOTE1 = dst.Tables("SOTTOTET").Rows.Find(SCAN)
                Else
                    rowSOTTOTE1 = LookUp("SOTTOTE1", SCAN)
                End If

                If rowSOTTOTE1 Is Nothing OrElse rowSOTTOTE1.Item("TOTE_NO") & "" <> SCAN Then
                    RESULT = "Invalid Value for Tote"
                Else
                    Dim TOTE_NO As String = SCAN
                    Dim DC_CODE_tote As String = rowSOTTOTE1.Item("DC_CODE") & ""

                    If DC_CODE_tote <> DC_CODE Then
                        RESULT = $"Tote {TOTE_NO} is in DC {DC_CODE_tote}"
                    Else
                        If TOTEs_in_SLOTs.Contains(TOTE_NO) Then
                            Dim SLOT As Integer = TOTEs_in_SLOTs.IndexOf(TOTE_NO) + 1
                            RESULT = $"Tote {TOTE_NO} has already been scanned into this Truck into Slot {CStr(SLOT)}"
                        Else
                            Dim PICK_NO As String = rowSOTTOTE1.Item("PICK_NO") & ""
                            If PICK_NO <> "" Then
                                RESULT = $"Tote {TOTE_NO} is in use in Pick No {PICK_NO}"
                            Else
                                Dim rowSOTORDRQ1s() As DataRow = dst.Tables("SOTORDQ1").Select("TOTE_NO IS NULL")
                                If rowSOTORDRQ1s.Length = 0 Then
                                    RESULT = $"No Orders Open to assign to Tote {TOTE_NO}"
                                Else
                                    If Not ASCMAIN1.Logical_Lock("SOTTOTE1", TOTE_NO,, False, False, 1) Then
                                        RESULT = $"Cannot Lock Tote {TOTE_NO}"
                                    Else
                                        RESULT = $"Tote {TOTE_NO} Accepted"
                                        TOTEs_in_SLOTs.Add(TOTE_NO)
                                        rowSOTORDRQ1s(0).Item("TOTE_NO") = TOTE_NO
                                        rowSOTORDRQ1s(0).Item("SLOT_NO") = TOTEs_in_SLOTs.Count

                                        Dim TOTE_CLASS_CODE As String = rowSOTTOTE1.Item("TOTE_CLASS_CODE") & ""
                                        Dim rowSOTTOTE0 As DataRow = dst.Tables("SOTTOTE0").Rows.Find(TOTE_CLASS_CODE)

                                        rowSOTORDRQ1s(0).Item("TOTE_CLASS_CODE") = TOTE_CLASS_CODE
                                        rowSOTORDRQ1s(0).Item("TOTE_CLASS_MIN_QTY") = rowSOTTOTE0.Item("TOTE_CLASS_MIN_QTY")
                                        rowSOTORDRQ1s(0).Item("TOTE_CLASS_MAX_QTY") = rowSOTTOTE0.Item("TOTE_CLASS_MAX_QTY")

                                        If dst.Tables("SOTORDQ1").Select("TOTE_NO IS NULL").Length = 0 Then
                                            lblInstruction.Text = "Tote Assignments are Complete"

                                            For Each row As DataRow In dst.Tables("SOTORDQ1").Select("")
                                                row.Item("VERIFIED") = DBNull.Value
                                            Next

                                            Dim OrderCount As Integer = dst.Tables("SOTORDQ1").Rows.Count

                                            verifyToteSlot = (dst.Tables("SOTORDQ1").Select("VERIFIED = '0'").Length > 0)
                                            verifyToteSlot = False ' TURN OFF ALL VERIFICATION

                                            If verifyToteSlot Then
                                                lblInstruction.Text = "Tote Assignments are Complete - Verify Totes/Slots"
                                                lblInstruction2.Visible = True
                                                lblInstruction2.Text = $"Scan Tote in Slot {CStr(SLOT_NO_to_Verify)}"
                                            Else
                                                btnSimulateScan.Enabled = False
                                                cmdUpdate.Enabled = True
                                                lblTruckIsReady.Visible = True
                                            End If

                                        Else
                                            lblInstruction.Text = $"Scan the Tote in Slot {TOTEs_in_SLOTs.Count + 1}"
                                        End If

                                        ERR = "0"
                                    End If

                                End If
                            End If
                        End If

                    End If

                End If
            End If
        End If

        If RESULT <> "" Then
            Record_Scan(SCAN, RESULT, ERR)
        End If

        txtSCAN.Text = ""

    End Sub

    Sub Record_Scan(SCAN As String, RESULT As String, ERR As String)
        Dim SCAN_NO As Integer = dst.Tables("SOTSCAN1").Rows.Count + 1
        Dim rowSOTSCAN1 As DataRow = dst.Tables("SOTSCAN1").NewRow
        With rowSOTSCAN1
            .Item("SCAN_NO") = SCAN_NO
            .Item("SCAN") = SCAN
            .Item("RESULT") = RESULT
            .Item("ERR") = ERR
        End With
        dst.Tables("SOTSCAN1").Rows.Add(rowSOTSCAN1)
        Sort_grdColumns(grdSOTSCAN1, "SCAN_NO".ToLower)
    End Sub
    Private Sub grdSOTSCAN1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTSCAN1.InitializeRow

    End Sub

    Private Sub grdSOTORDQ1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTORDQ1.InitializeRow

    End Sub

    Private Sub chkCustomTruck_CheckedChanged(sender As Object, e As EventArgs) Handles chkCustomTruck.CheckedChanged
        If chkCustomTruck.Checked Then
            chkCustomTruck.Appearance.ForeColor = System.Drawing.Color.Red
        Else
            chkCustomTruck.Appearance.ForeColor = System.Drawing.Color.Empty

        End If
    End Sub
End Class