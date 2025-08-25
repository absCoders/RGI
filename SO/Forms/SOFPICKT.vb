Imports Infragistics.Win.UltraWinGrid

Public Class SOFPICKT
    Public SOTORDQ1 As String
    Public SOTORDR0 As String
    Public TOTEs_in_SLOTs As List(Of String)
    Public update_flag As Boolean = False
    Public TRUCK_NO As String
    Public PICK_DESCRIPTION As String
    Public tbl As DataTable
    Public rowSOTORDQ0 As DataRow
    Public WHSE_CODE As String = ""
    Dim SLOT_NO_to_Verify As Integer = 0
    Public SALES_DIVISION_CODE_DC As String = ""

    Dim Appearance_Magenta As New Infragistics.Win.Appearance
    Dim Appearance_Yellow As New Infragistics.Win.Appearance
    Dim Appearance_Empty As New Infragistics.Win.Appearance


    Public Sub New(ByVal FF As ASFBASE1)
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Appearance_Magenta.BackColor = Drawing.Color.Magenta
        Appearance_Yellow.BackColor = Drawing.Color.Yellow

        With dst
            ASCMAIN1.sql = $"Select SOTORDQ1.* from {SOTORDQ1} SOTORDQ1, {SOTORDR0} SOTORDR0 where SOTORDQ1.ORDR_NO = SOTORDR0.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTORDQ1", "**", 0, False)
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
                .Columns.Add("TRUCK_NO")
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


            ASCMAIN1.sql = "Select * from SOTTRCK1 where WHSE_CODE = :PARM1 AND TRUCK_TYPE = 'P' AND PICK_BATCH_NO IS NULL"
            Create_TDA(.Tables.Add, "SOTTRCKT", "**", 0, False, "V", 1)
            Fill_Records("SOTTRCKT", WHSE_CODE)
            Dim TRUCK_NOs As String = ""
            For Each rowSOTTRCKT As DataRow In dst.Tables("SOTTRCKT").Select("", "TRUCK_NO")
                Dim TRUCK_NO As String = rowSOTTRCKT.Item("TRUCK_NO")
                TRUCK_NOs &= "," & TRUCK_NO
                If TRUCK_NOs.Length > 60 Then
                    TRUCK_NOs &= "," & "..."
                    Exit For
                End If
            Next
            lblInstruction1.Text = Mid(TRUCK_NOs, 2)
        End With

        Fill_Records("SOTORDQ1")
        'Dim rowQ1s() As DataRow = frmASFBASE1.dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' and SEL = '1'")
        'For Each row As DataRow In rowQ1s
        '    dst.Tables("SOTORDQ1").Rows.Add(row.ItemArray)
        'Next
        grdSOTORDQ1.DataSource = dst.Tables("SOTORDQ1")
        Sort_grdColumns(grdSOTORDQ1, "ORDR_QTY_ALLO,".ToLower & "ORDR_NO")
        grdSOTSCAN1.DataSource = dst.Tables("SOTSCAN1")
        Calculate_Totals()

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
        'Create_Summary(grdSOTORDQ1, New String() {"ORDR_QTY_OPEN", "ORDR_QTY_BACK", "ORDR_QTY_ALLO"})
        Create_Summary(grdSOTORDQ1, New String() {"ORDR_QTY_OPEN", "ORDR_QTY_ALLO"})

        grdSOTORDQ1.DisplayLayout.Bands(0).Columns("ORDR_QTY_ALLO").Hidden = (SALES_DIVISION_CODE_DC <> "SKIN")
        grdSOTORDQ1.DisplayLayout.Bands(0).Columns("ORDR_QTY_ALLO").Hidden = (SALES_DIVISION_CODE_DC <> "SKIN")

        'cmdUpdate.Enabled = False

        'Dim rowSOTORDQ0 As DataRow = frmASFBASE1.dst.Tables("SOTORDQ0").Rows.Find(GROUP_KEY)

        PICK_DESCRIPTION = "Release Batch"

        Me.Text &= ": " & PICK_DESCRIPTION
        Me.Width = cmdCancel.Left + cmdCancel.Width + 25
        CenterToParent()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        If dst.Tables("SOTORDQ1").Select("TOTE_NO IS NOT NULL AND BO <> '1'").Length = 0 And dst.Tables("SOTORDQ1").Select("BO <> '1'").Length > 0 Then
            MsgBox("No Orders have been Assigned to Trucks", MsgBoxStyle.OkOnly, "Update is NOT Permitted")
            Exit Sub
        End If

        Dim OrderCountNotInTrucks As Integer = dst.Tables("SOTORDQ1").Select("TOTE_NO IS NULL AND BO <> '1'").Length
        If OrderCountNotInTrucks <> 0 Then
            If MsgBox($"Some {CStr(OrderCountNotInTrucks)} Orders have NOT been Assigned to Trucks/Totes." & vbCrLf & vbCrLf & "OK to Continue?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                Exit Sub
            End If
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
        End If
    End Sub

    Private Sub txtSCAN_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSCAN.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Process_Scan()
        End If
    End Sub

    Sub Process_Scan()

        If lblDone.Visible Then
            Exit Sub
        End If

        Dim SCAN As String = txtSCAN.Text.ToUpper
        Dim RESULT As String = ""
        Dim ERR As String = "1"

        Dim rowSOTORDRQ1s() As DataRow

        Dim rowSOTTRCK1 As DataRow = LookUp("SOTTRCK1", SCAN)
        If rowSOTTRCK1 Is Nothing Then
            RESULT = "Invalid Value for Truck"
        Else
            TRUCK_NO = SCAN

            Dim WHSE_CODE_truck As String = rowSOTTRCK1.Item("WHSE_CODE") & ""
            If WHSE_CODE_truck <> WHSE_CODE Then
                RESULT = $"Truck {TRUCK_NO} is in DC {WHSE_CODE_truck}"
                Record_Scan(SCAN, RESULT, ERR)
            Else
                Dim PICK_BATCH_NO As String = rowSOTTRCK1.Item("PICK_BATCH_NO") & ""
                If PICK_BATCH_NO <> "" Then
                    RESULT = $"Truck {TRUCK_NO} is in use in Pick Batch {PICK_BATCH_NO}"
                    Record_Scan(SCAN, RESULT, ERR)
                Else
                    Dim TRUCK_TYPE As String = rowSOTTRCK1.Item("TRUCK_TYPE") & ""

                    If TRUCK_TYPE <> "P" Then
                        RESULT = $"Truck {TRUCK_NO} is NOT a Pre-Configured Truck"
                        Record_Scan(SCAN, RESULT, ERR)
                    Else
                        Fill_Records("SOTTOTET", TRUCK_NO)
                        Dim TOTE_COUNT As Integer = dst.Tables("SOTTOTET").Rows.Count
                        If TOTE_COUNT = 0 Then
                            RESULT = $"Truck {TRUCK_NO} is Pre-Configured but Tote Configuration shows no Totes"
                        Else

                            If Not ASCMAIN1.Logical_Lock("SOTTRCK1", TRUCK_NO,, False, False) Then
                                RESULT = $"Cannot Lock Truck {TRUCK_NO}"
                            Else
                                RESULT = $"Truck {TRUCK_NO} Accepted"
                                ERR = "0"
                                Record_Scan(SCAN, RESULT, ERR)

                                Dim SLOT_NO_ctr As Integer = 0
                                For Each rowSOTTOTE1 As DataRow In dst.Tables("SOTTOTET").Select("", "TOTE_NO")
                                    RESULT = ""
                                    rowSOTORDRQ1s = dst.Tables("SOTORDQ1").Select("TOTE_NO IS NULL AND BO <> '1'", "ORDR_QTY_ALLO DESC, ORDR_NO")
                                    If rowSOTORDRQ1s.Length = 0 Then
                                        Exit For
                                    End If

                                    Dim TOTE_NO As String = rowSOTTOTE1.Item("TOTE_NO")
                                    Dim WHSE_CODE_tote As String = rowSOTTOTE1.Item("WHSE_CODE") & ""

                                    If WHSE_CODE_tote <> WHSE_CODE Then
                                        RESULT = $"Tote {TOTE_NO} is in DC {WHSE_CODE_tote}"
                                    Else

                                        Dim PICK_NO As String = rowSOTTOTE1.Item("PICK_NO") & ""
                                        If PICK_NO <> "" Then
                                            RESULT = $"Tote {TOTE_NO} is in use in Pick No {PICK_NO}"
                                        Else
                                            If Not ASCMAIN1.Logical_Lock("SOTTOTE1", TOTE_NO,, False, False) Then
                                                RESULT = $"Cannot Lock Tote {TOTE_NO}"
                                            Else
                                                ' RESULT = $"Tote {TOTE_NO} Accepted"
                                                'Record_Scan(SCAN, RESULT, ERR)

                                                rowSOTORDRQ1s(0).Item("TRUCK_NO") = TRUCK_NO
                                                rowSOTORDRQ1s(0).Item("TOTE_NO") = TOTE_NO
                                                SLOT_NO_ctr += 1
                                                rowSOTORDRQ1s(0).Item("SLOT_NO") = SLOT_NO_ctr

                                                Dim TOTE_CLASS_CODE As String = rowSOTTOTE1.Item("TOTE_CLASS_CODE") & ""
                                                Dim rowSOTTOTE0 As DataRow = dst.Tables("SOTTOTE0").Rows.Find(TOTE_CLASS_CODE)

                                                rowSOTORDRQ1s(0).Item("TOTE_CLASS_CODE") = TOTE_CLASS_CODE
                                                rowSOTORDRQ1s(0).Item("TOTE_CLASS_MIN_QTY") = rowSOTTOTE0.Item("TOTE_CLASS_MIN_QTY")
                                                rowSOTORDRQ1s(0).Item("TOTE_CLASS_MAX_QTY") = rowSOTTOTE0.Item("TOTE_CLASS_MAX_QTY")
                                            End If
                                        End If
                                    End If
                                    If RESULT <> "" Then
                                        ERR = "1"
                                        Record_Scan(SCAN, RESULT, ERR)
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If

        txtSCAN.Text = ""
        Calculate_Totals()

        rowSOTORDRQ1s = dst.Tables("SOTORDQ1").Select("TOTE_NO IS NULL AND BO <> '1'")
        If rowSOTORDRQ1s.Length = 0 Then
            lblDone.Visible = True
            'cmdUpdate.Enabled = True
        End If

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
        Dim ORDR_QTY_OPEN As Integer = Val(e.Row.Cells("ORDR_QTY_OPEN").Value & "")
        'Dim ORDR_QTY_BACK As Integer = Val(e.Row.Cells("ORDR_QTY_BACK").Value & "")
        Dim ORDR_QTY_ALLO As Integer = Val(e.Row.Cells("ORDR_QTY_ALLO").Value & "")

        Dim BO As String = e.Row.Cells("BO").Value & ""
        If BO = "1" Then
            'e.Row.Appearance = Appearance_Magenta
            e.Row.Appearance.BackColor = System.Drawing.Color.LightPink
        Else
            If ORDR_QTY_ALLO <> ORDR_QTY_OPEN Then ' If ORDR_QTY_ALLO <> ORDR_QTY_OPEN + ORDR_QTY_BACK Then
                e.Row.Cells("ORDR_QTY_ALLO").Appearance = Appearance_Yellow
            Else
                e.Row.Cells("ORDR_QTY_ALLO").Appearance = Nothing
            End If
        End If


    End Sub

    Sub Calculate_Totals()
        Dim ASSIGNED As Integer = dst.Tables("SOTORDQ1").Select("TOTE_NO IS NOT NULL").Length
        Dim UNASSIGNED As Integer = dst.Tables("SOTORDQ1").Select("TOTE_NO IS NULL AND BO <> '1'").Length
        Dim BACKORDERED As Integer = dst.Tables("SOTORDQ1").Select("BO = '1'").Length
        lblAssigned.Text = Format(ASSIGNED, "#,##0")
        lblUnassigned.Text = Format(UNASSIGNED, "#,##0")
        lblBackOrdered.Text = Format(BACKORDERED, "#,##0")
    End Sub

End Class